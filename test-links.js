/**
 * Get Visa emails with Outlook deep links
 */

const { searchEmails } = require('./bridge-client');

/**
 * Create an Outlook deep link from email EntryID
 * Format: outlook://#email/EntryID
 */
function createOutlookLink(entryId, subject = '') {
    // Clean up the entryID for URL usage
    const cleanId = encodeURIComponent(entryId);
    return `outlook://#email/${cleanId}`;
}

/**
 * Create a web link to Outlook Web App (if you use OWA)
 */
function createOwaLink(itemName, itemId) {
    // This works for Outlook on the Web
    return `https://outlook.live.com/mail/0/deeplink/compose?subject=${encodeURIComponent(itemName)}&cc=${encodeURIComponent(itemId)}`;
}

async function runTest() {
    console.log('\nSearching for Visa emails with direct links...\n');
    
    const result = await searchEmails('Visa', {
        folder: 'inbox',
        limit: 50,
        unreadOnly: false,
        days: 180,  // 6 months
        includeFullBody: false
    });
    
    console.log(`Found ${result.count} Visa emails\n`);
    
    // Sort by date (newest first)
    const sorted = result.emails.sort((a, b) => {
        return new Date(b.received_datetime || 0) - new Date(a.received_datetime || 0);
    });
    
    console.log('=== RECENT EMAILS WITH OUTLOOK LINKS ===\n');
    
    sorted.slice(0, 25).forEach((email, idx) => {
        const subject = email.subject || '(no subject)';
        const sender = `${email.sender || ''} <${email.sender_email || ''}>`.trim();
        const date = email.received_datetime ? new Date(email.received_datetime).toLocaleString() : 'unknown';
        const attachments = email.has_attachments ? ` (${email.attachments_count} attachments)` : '';
        
        // Create Outlook deep link
        const outlookLink = createOutlookLink(email.id, subject);
        
        console.log(`${idx + 1}. ${subject}${attachments}`);
        console.log(`   From: ${sender}`);
        console.log(`   Date: ${date}`);
        console.log(`   Outlook Link: ${outlookLink}`);
        console.log(`   EntryID: ${email.id.substring(0, 50)}...`);
        console.log('');
    });
}

runTest().catch(e => {
    console.error('Error:', e);
    process.exit(1);
});
