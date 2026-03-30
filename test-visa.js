/**
 * Quick test script for Visa search
 */

const { searchOutlook, formatEmailForDisplay } = require('./searcher');

async function runTest() {
    console.log('\nSearching for Visa project documentation...\n');
    
    const results = await searchOutlook({
        query: 'Visa statement of work SOW document',
        context: 'looking for signed document or detailed description of work to be done on site installation',
        dateHint: 'last 4 months'
    }, {
        topN: 10,
        deepReadLimit: 5
    });
    
    console.log('\n' + '='.repeat(70));
    console.log('RESULTS');
    console.log('='.repeat(70));
    
    if (results.results.length === 0) {
        console.log('\nNo relevant results found.');
        return;
    }
    
    console.log(`\n${results.summary}\n`);
    
    results.results.forEach((result, idx) => {
        const subject = result.subject || '(no subject)';
        const sender = `${result.sender || ''} <${result.sender_email || ''}>`.trim();
        const date = result.received_datetime ? new Date(result.received_datetime).toLocaleString() : 'unknown';
        const readStatus = result.is_read ? '' : '[unread]';
        const attachments = result.has_attachments ? ` (has ${result.attachments_count} attachment(s))` : '';
        
        console.log(`\n${idx + 1}. [Score: ${result.relevanceScore || 'N/A'}]`);
        console.log(`FROM: ${sender}`);
        console.log(`DATE: ${date}`);
        console.log(`SUBJECT: ${subject}${attachments}`);
        if (result.reasoning) {
            console.log(`REASONING: ${result.reasoning}`);
        }
        if (result.links && result.links.length > 0) {
            console.log(`LINKS: ${result.links.join(', ')}`);
        }
        if (result.body) {
            console.log(`PREVIEW: ${result.body.substring(0, 150).replace(/\n/g, ' ').trim()}...`);
        }
        console.log('-'.repeat(50));
    });
}

runTest().catch(e => {
    console.error('Error:', e);
    process.exit(1);
});
