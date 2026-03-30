/**
 * Outlook Searcher
 * 
 * Orchestrates multi-phase search through Outlook using outlook-mapi.
 * Implements broad scan → filter → deep read strategy.
 */

const { expandSearchIntent } = require('./query-expander');
const { parseDate } = require('./date-parser');
const { analyzeEmailRelevance, extractRelevantSnippets } = require('./llm');

/**
 * Initialize the MAPI client
 */
async function initMapi() {
    try {
        const mapiModule = require('outlook-mapi');
        return new mapiModule.OutlookMapi();
    } catch (error) {
        console.error('Error initializing Outlook MAPI:', error.message);
        throw error;
    }
}

/**
 * PHASE 1: Broad search - get subjects and snippets
 */
async function broadSearch(mapi, query, options = {}) {
    const { dateRange = null, limit = 50 } = options;
    
    console.log(`\n📬 Phase 1: Broad search for "${query}"`);
    
    try {
        let searchQuery = query;
        
        // Add date filter if provided
        if (dateRange && dateRange.iso) {
            searchQuery += ` AND received/ge '${dateRange.iso.start}' AND received/le '${dateRange.iso.end}'`;
        }
        
        const results = await mapi.searchMailItems({
            query: searchQuery,
            top: limit,
            properties: [
                'subject',
                'senderEmail',
                'senderName',
                'receivedDateTime',
                'isRead',
                'hasAttachments',
                'importance'
            ]
        });
        
        console.log(`  Found ${results.items?.length || 0} emails`);
        return results.items || [];
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
async function deepRead(mapi, topEmails, maxRead = 10) {
    console.log(`\n📖 Phase 2: Deep reading top ${Math.min(topEmails.length, maxRead)} emails...`);
    
    const detailedResults = [];
    
    for (let i = 0; i < Math.min(topEmails.length, maxRead); i++) {
        const email = topEmails[i];
        
        try {
            // Fetch full email content
            const fullEmail = await mapi.getEmailById(email.id, {
                properties: [
                    'subject',
                    'body',
                    'senderEmail',
                    'senderName',
                    'receivedDateTime',
                    'hasAttachments',
                    'attachments',
                    'importance'
                ]
            });
            
            // Extract links from body
            const links = extractLinks(fullEmail.body?.content || '');
            
            // Get attachment info
            const attachments = fullEmail.attachments?.map(a => ({
                name: a.name,
                type: a.type,
                size: a.size
            })) || [];
            
            detailedResults.push({
                ...fullEmail,
                links,
                attachments: attachments.map(a => a.name),
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
 * Extract URLs from email body
 */
function extractLinks(text) {
    if (!text) return [];
    
    // Match http/https URLs
    const httpUrls = text.match(/https?:\/\/[^\s<"')]+[\w\/]/g) || [];
    
    // Deduplicate and return
    return [...new Set(httpUrls)].slice(0, 10);
}

/**
 * Format a single email for display
 */
function formatEmailForDisplay(email) {
    const subject = email.subject || '(no subject)';
    const sender = `${email.senderName || ''} <${email.senderEmail || ''}>`.trim();
    const date = email.receivedDateTime 
        ? new Date(email.receivedDateTime).toLocaleString()
        : 'unknown';
    const readStatus = email.isRead ? '' : '[unread]';
    const importance = email.importance === 'High' ? ' [HIGH]' : '';
    const attachments = email.attachments?.length 
        ? ` (${email.attachments.join(', ')})`
        : '';
    const links = email.links?.length 
        ? `\n  Links: ${email.links.slice(0, 3).join(', ')}${email.links.length > 3 ? '...' : ''}`
        : '';
    
    // Truncate body preview
    const bodyPreview = email.body?.content 
        ? `\n  Preview: ${email.body.content.substring(0, 200).replace(/\n/g, ' ').trim()}...`
        : '';
    
    return `${readStatus}${importance} FROM: ${sender}\n   DATE: ${date}\nSUBJECT: ${subject}${attachments}\n${bodyPreview}\n${links}`;
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
    
    // Initialize MAPI client
    const mapi = await initMapi();
    
    // PHASE 1: Broad search with multiple queries
    let allResults = [];
    for (const queryStr of expandedQueries.variants) {
        const results = await broadSearch(mapi, queryStr, { dateRange: dateRange });
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
    const detailedResults = await deepRead(mapi, analysis.filtered, deepReadLimit);
    
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
    extractLinks
};
