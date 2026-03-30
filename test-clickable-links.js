/**
 * Quick Visa search with CLICKABLE Telegram links
 */

const { searchOutlook, formatEmailResult } = require('./searcher');
const { createEmailLink } = require('./bridge-client');

async function runTest() {
    console.log('\n=== SEARCHING FOR: "Visa statement of work SOW" ===\n');
    
    const results = await searchOutlook({
        query: 'Visa',  // Start simple - complex queries may not work well
        context: 'looking for signed document or detailed description of work to be done',
        dateHint: 'last 4 months'
    }, {
        topN: 10,
        deepReadLimit: 5
    });
    
    console.log('\n' + '='.repeat(70));
    console.log('RESULTS WITH CLICKABLE LINKS');
    console.log('='.repeat(70));
    
    if (results.results.length === 0) {
        // If no results from complex query, fall back to simple "Visa" search
        console.log('\nNo results from filtered search. Showing recent Visa emails...\n');
        
        const { searchEmails } = require('./bridge-client');
        const rawResult = await searchEmails('Visa', {
            folder: 'inbox',
            limit: 10,
            days: 30, // Just last month
            includeFullBody: false
        });
        
        const emails = rawResult.emails || [];
        console.log(`Found ${emails.length} recent emails about "Visa"\n`);
        
        displayEmailsWithLinks(emails);
        return;
    }
    
    console.log(`\n${results.summary}\n`);
    
    displayEmailsWithLinks(results.results);
}

/**
 * Display emails with both raw URLs and Telegram-clickable format
 */
function displayEmailsWithLinks(emails) {
    emails.forEach((email, idx) => {
        // Format with email link
        const formatted = formatEmailResult(email);
        
        // Create Telegram markdown clickable link
        const clickableView = `📧 [Open Email](${formatted.emailLink})`;
        
        console.log(`${idx + 1}. [Score: ${formatted.relevanceScore || 'N/A'}]`);
        console.log(`   Subject: ${formatted.subject}${formatted.attachments}`);
        console.log(`   From: ${formatted.sender}`);
        console.log(`   Date: ${formatted.date}`);
        console.log(`   ${clickableView}`);
        console.log(`   Raw URL: ${formatted.emailLink}`);
        if (formatted.reasoning) {
            console.log(`   Why: ${formatted.reasoning}`);
        }
        console.log('');
    });
}

runTest().catch(e => {
    console.error('Error:', e);
    process.exit(1);
});
