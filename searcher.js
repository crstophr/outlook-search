/**
 * Outlook Searcher
 * 
 * Orchestrates multi-phase search through Outlook using the bridge.
 * Implements broad scan → LLM filter → deep read strategy.
 */

const { expandSearchIntent } = require('./query-expander');
const { parseDate } = require('./date-parser');
const { analyzeEmailRelevance, extractRelevantSnippets } = require('./llm');
const { searchEmails, getEmailById, healthCheck, extractLinks: bridgeExtractLinks } = require('./bridge-client');

/**
 * Create an Outlook deep link from email EntryID
 * This allows direct opening in Outlook desktop app
 */
function createOutlookLink(entryId) {
    const cleanId = encodeURIComponent(entryId);
    return `outlook://#email/${cleanId}`;
}

/**
 * Verify bridge is available
 */
async function verifyBridge() {
    try {
        const health = await healthCheck();
        if (!health.outlook_connected) {
            throw new Error('Outlook bridge is not connected to Outlook');
        }
        return true;
    } catch (error) {
        console.error('Bridge health check failed:', error.message);
        throw error;
    }
}

/**
 * PHASE 1: Broad search - get subjects and metadata
 */
async function broadSearch(query, options = {}) {
    const { dateRange = null, limit = 100 } = options;
    
    console.log(`\n📬 Phase 1: Broad search for "${query}"`);
    
    try {
        // Calculate days parameter from date range
        let daysParam = null;
        if (dateRange && dateRange.range) {
            const now = new Date();
            const startDate = dateRange.range.startDate || dateRange.range.start;
            const daysAgo = Math.ceil((now - startDate) / (1000 * 60 * 60 * 24));
            daysParam = daysAgo;
        }
        
        const result = await searchEmails(query, {
            limit,
            unreadOnly: false,
            days: daysParam,
            includeFullBody: false
        });
        
        const emails = result.emails || [];
        console.log(`  Found ${emails.length} emails`);
        return emails;
    } catch (error) {
        console.error('Error in broad search:', error.message);
        return [];
    }
}

/**
 * Analyze a batch of email summaries with LLM to find most relevant ones
 */
async function analyzeEmailSummaries(emails, searchContext) {
    if (emails.length === 0) {
        return { scored: [], filtered: [] };
    }
    
    console.log(`\n🧠 Analyzing ${emails.length} emails with LLM...`);
    
    // Use the LLM module for intelligent analysis
    const result = await analyzeEmailRelevance(emails, searchContext);
    
    console.log(`  Found ${result.filtered.length} relevant emails (score >= 4)`);
    
    return result;
}

/**
 * PHASE 2: Deep read - fetch full email bodies for top candidates
 */
async function deepRead(topEmails, maxRead = 10) {
    console.log(`\n📖 Phase 2: Deep reading top ${Math.min(topEmails.length, maxRead)} emails...`);
    
    const detailedResults = [];
    
    for (let i = 0; i < Math.min(topEmails.length, maxRead); i++) {
        const email = topEmails[i];
        
        try {
            // Fetch full email content with body
            const result = await getEmailById(email.id);
            const fullEmail = result.email || {};
            
            // Extract links from body (try html_body first, then body)
            const bodyText = fullEmail.html_body || fullEmail.body || '';
            const links = bridgeExtractLinks(bodyText);
            
            detailedResults.push({
                ...fullEmail,
                links,
                relevanceScore: email.relevanceScore,
                reasoning: email.reasoning
            });
            
        } catch (error) {
            console.error(`  Error reading email ${email.id}:`, error.message);
        }
    }
    
    return detailedResults;
}

/**
 * Format a single email for display
 */
function formatEmailForDisplay(email) {
    const subject = email.subject || '(no subject)';
    const sender = `${email.sender || ''} <${email.sender_email || ''}>`.trim();
    const date = email.received_datetime 
        ? new Date(email.received_datetime).toLocaleString()
        : 'unknown';
    const readStatus = email.is_read ? '' : '[unread]';
    const importance = email.importance === 'High' ? ' [HIGH]' : '';
    const attachments = email.has_attachments 
        ? ` (has ${email.attachments_count || 0} attachment(s))`
        : '';
    const links = email.links?.length 
        ? `\n  Links: ${email.links.slice(0, 3).join(', ')}${email.links.length > 3 ? '...' : ''}`
        : '';
    
    // Truncate body preview
    const bodyPreview = email.body 
        ? `\n  Preview: ${email.body.substring(0, 200).replace(/\n/g, ' ').trim()}...`
        : '';
    
    return `${readStatus}${importance} FROM: ${sender}\n   DATE: ${date}\nSUBJECT: ${subject}${attachments}\n${bodyPreview}\n${links}`;
}

/**
 * Format email result with Outlook deep link for easy opening
 */
function formatEmailResult(email) {
    const subject = email.subject || '(no subject)';
    const sender = `${email.sender || ''} <${email.sender_email || ''}>`.trim();
    const date = email.received_datetime 
        ? new Date(email.received_datetime).toLocaleString()
        : 'unknown';
    const attachments = email.has_attachments 
        ? ` (${email.attachments_count} attachment(s))`
        : '';
    
    // Create Outlook deep link
    const outlookLink = email.id ? createOutlookLink(email.id) : null;
    
    return {
        subject,
        sender,
        date,
        attachments,
        outlookLink,
        entryId: email.id,
        relevanceScore: email.relevanceScore,
        reasoning: email.reasoning
    };
}

/**
 * Main search orchestration function
 */
async function searchOutlook(searchContext, options = {}) {
    const { query, context = '', dateHint = null } = searchContext;
    const { topN = 10, deepReadLimit = 5 } = options;
    
    console.log('\n' + '='.repeat(60));
    console.log('OUTLOOK SEARCH');
    console.log('='.repeat(60));
    console.log(`Query: ${query}`);
    console.log(`Context: ${context || 'None'}\n`);
    
    // Verify bridge is available
    try {
        await verifyBridge();
        console.log('✓ Bridge connected to Outlook\n');
    } catch (error) {
        return { results: [], summary: 'Cannot connect to Outlook. Please start the bridge server.', error: error.message };
    }
    
    // Parse date hint if provided
    let dateRange = null;
    if (dateHint) {
        dateRange = parseDate(dateHint);
        console.log(`Date range: ${dateRange.description}`);
    }
    
    // Expand search intent into queries
    const expandedQueries = expandSearchIntent(query, { includeDocumentFilters: true });
    console.log(`\nExpanded queries:`);
    console.log(`  Primary: ${expandedQueries.primary.substring(0, 80)}...`);
    console.log(`  Variants: ${expandedQueries.variants.length}\n`);
    
    // PHASE 1: Broad search with multiple queries
    let allResults = [];
    for (const queryStr of expandedQueries.variants) {
        const results = await broadSearch(queryStr, { dateRange });
        allResults = [...allResults, ...results];
    }
    
    // Deduplicate by email ID
    const uniqueEmails = Array.from(
        new Map(allResults.map(e => [e.id, e])).values()
    );
    
    console.log(`\nUnique emails found: ${uniqueEmails.length}`);
    
    if (uniqueEmails.length === 0) {
        return { results: [], summary: 'No results found' };
    }
    
    // Analyze with LLM
    const analysis = await analyzeEmailSummaries(uniqueEmails, searchContext);
    console.log(`After LLM filtering: ${analysis.filtered.length} emails`);
    
    // PHASE 2: Deep read top candidates
    const detailedResults = await deepRead(analysis.filtered, deepReadLimit);
    
    // Sort by relevance score
    detailedResults.sort((a, b) => (b.relevanceScore || 0) - (a.relevanceScore || 0));
    
    return {
        results: detailedResults,
        summary: `Found ${detailedResults.length} relevant emails from ${uniqueEmails.length} matches`,
        metadata: {
            queriesRun: expandedQueries.variants.length,
            initialMatches: uniqueEmails.length,
            afterFiltering: analysis.filtered.length
        }
    };
}

module.exports = {
    searchOutlook,
    broadSearch,
    deepRead,
    analyzeEmailSummaries,
    formatEmailForDisplay,
    formatEmailResult,
    createOutlookLink,
    verifyBridge
};
