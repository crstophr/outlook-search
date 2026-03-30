/**
 * Quick Visa search with direct Outlook links
 */

const { searchOutlook, formatEmailResult } = require('./searcher');

async function runTest() {
    console.log('\nSearching for: "Visa statement of work SOW"\n');
    
    const results = await searchOutlook({
        query: 'Visa statement of work SOW',
        context: 'looking for signed document or detailed description of work to be done',
        dateHint: 'last 4 months'
    }, {
        topN: 10,
        deepReadLimit: 5
    });
    
    console.log('\n' + '='.repeat(70));
    console.log('RESULTS WITH OUTLOOK LINKS');
    console.log('='.repeat(70));
    
    if (results.results.length === 0) {
        console.log('\nNo relevant results found.');
        return;
    }
    
    console.log(`\n${results.summary}\n`);
    
    results.results.forEach((result, idx) => {
        // Format with Outlook link
        const formatted = formatEmailResult(result);
        
        console.log(`${idx + 1}. [Score: ${formatted.relevanceScore || 'N/A'}]`);
        console.log(`   Subject: ${formatted.subject}${formatted.attachments}`);
        console.log(`   From: ${formatted.sender}`);
        console.log(`   Date: ${formatted.date}`);
        console.log(`   👉 Outlook Link: ${formatted.outlookLink}`);
        if (formatted.reasoning) {
            console.log(`   Reasoning: ${formatted.reasoning}`);
        }
        console.log('');
    });
    
    // Also print the links in a copy-friendly format
    console.log('\n=== QUICK LINKS (copy these to open directly in Outlook) ===');
    results.results.slice(0, 5).forEach((result, idx) => {
        const link = result.id ? createOutlookLink(result.id) : null;
        console.log(`${idx + 1}. ${link}`);
    });
}

function createOutlookLink(entryId) {
    const cleanId = encodeURIComponent(entryId);
    return `outlook://#email/${cleanId}`;
}

runTest().catch(e => {
    console.error('Error:', e);
    process.exit(1);
});
