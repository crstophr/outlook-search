/**
 * Simpler test - basic Visa search
 */

const { searchEmails } = require('./bridge-client');

async function runTest() {
    console.log('\nSimple search for "Visa"...\n');
    
    // Just search for Visa in the last 4 months (~120 days)
    const result = await searchEmails('Visa', {
        folder: 'inbox',
        limit: 50,
        unreadOnly: false,
        days: 120,  // ~4 months
        includeFullBody: false
    });
    
    console.log(`Found ${result.count} emails mentioning "Visa"\n`);
    
    if (result.emails && result.emails.length > 0) {
        result.emails.slice(0, 10).forEach((email, idx) => {
            const subject = email.subject || '(no subject)';
            const sender = `${email.sender || ''} <${email.sender_email || ''}>`.trim();
            const date = email.received_datetime ? new Date(email.received_datetime).toLocaleString() : 'unknown';
            const attachments = email.has_attachments ? ` (has ${email.attachments_count} attachment(s))` : '';
            
            console.log(`${idx + 1}. ${subject}${attachments}`);
            console.log(`   From: ${sender}, Date: ${date}`);
        });
    }
}

runTest().catch(e => {
    console.error('Error:', e);
    process.exit(1);
});
