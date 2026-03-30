/**
 * Search for SOW/statement of work related emails
 */

const { searchEmails } = require('./bridge-client');

async function runTest() {
    console.log('\nSearching for "SOW" and "statement of work"...\n');
    
    // Search for SOW/statement of work in the last 4 months
    const result = await searchEmails('SOW statement of work scope', {
        folder: 'inbox',
        limit: 50,
        unreadOnly: false,
        days: 120,  // ~4 months
        includeFullBody: false
    });
    
    console.log(`Found ${result.count} emails mentioning "SOW" or "statement of work"\n`);
    
    if (result.emails && result.emails.length > 0) {
        result.emails.slice(0, 20).forEach((email, idx) => {
            const subject = email.subject || '(no subject)';
            const sender = `${email.sender || ''} <${email.sender_email || ''}>`.trim();
            const date = email.received_datetime ? new Date(email.received_datetime).toLocaleString() : 'unknown';
            const attachments = email.has_attachments ? ` (has ${email.attachments_count} attachment(s))` : '';
            
            console.log(`${idx + 1}. ${subject}${attachments}`);
            console.log(`   From: ${sender}, Date: ${date}`);
        });
    } else {
        console.log('No results. Let me try searching with "attachment" keyword...');
        
        const result2 = await searchEmails('Visa attachment pdf', {
            folder: 'inbox',
            limit: 30,
            unreadOnly: false,
            days: 120,
            includeFullBody: false
        });
        
        console.log(`\nFound ${result2.count} emails with "Visa" and "attachment/pdf"\n`);
        
        if (result2.emails && result2.emails.length > 0) {
            result2.emails.slice(0, 15).forEach((email, idx) => {
                const subject = email.subject || '(no subject)';
                const sender = `${email.sender || ''} <${email.sender_email || ''}>`.trim();
                const date = email.received_datetime ? new Date(email.received_datetime).toLocaleString() : 'unknown';
                const attachments = email.has_attachments ? ` (has ${email.attachments_count} attachment(s))` : '';
                
                console.log(`${idx + 1}. ${subject}${attachments}`);
                console.log(`   From: ${sender}, Date: ${date}`);
            });
        }
    }
}

runTest().catch(e => {
    console.error('Error:', e);
    process.exit(1);
});
