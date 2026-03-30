/**
 * Get Visa emails with full content for analysis
 */

const { searchEmails } = require('./bridge-client');

async function runTest() {
    console.log('\nGetting all Visa emails with attachments info...\n');
    
    // Get all Visa emails from last 4 months, limit 100
    const result = await searchEmails('Visa', {
        folder: 'inbox',
        limit: 100,
        unreadOnly: false,
        days: 120,
        includeFullBody: true
    });
    
    console.log(`Found ${result.count} total Visa emails\n`);
    
    // Filter for ones with attachments
    const withAttachments = result.emails.filter(e => e.has_attachments && e.attachments_count > 0);
    
    console.log(`=== EMAILS WITH ATTACHMENTS (${withAttachments.length}) ===\n`);
    
    withAttachments.slice(0, 25).forEach((email, idx) => {
        const subject = email.subject || '(no subject)';
        const sender = `${email.sender || ''} <${email.sender_email || ''}>`.trim();
        const date = email.received_datetime ? new Date(email.received_datetime).toLocaleString() : 'unknown';
        
        console.log(`${idx + 1}. ${subject}`);
        console.log(`   From: ${sender}, Date: ${date}`);
        console.log(`   Attachments: ${email.attachments_count}`);
        
        // Check if body mentions documents
        const bodyLower = (email.body || '').toLowerCase();
        if (bodyLower.includes('document') || bodyLower.includes('attachment') || 
            bodyLower.includes('pdf') || bodyLower.includes('signed')) {
            console.log(`   -> Body mentions: document/attachment/pdf/signed`);
        }
        console.log('');
    });
}

runTest().catch(e => {
    console.error('Error:', e);
    process.exit(1);
});
