/**
 * Examine specific Visa email with full body
 */

const { searchEmails, getEmailById } = require('./bridge-client');

async function runTest() {
    console.log('\nLooking for the ORIGINAL Visa email (earliest in thread)...\n');
    
    // Get all Visa emails sorted by date (oldest first)
    const result = await searchEmails('Visa', {
        folder: 'inbox',
        limit: 100,
        days: 180,  // 6 months to be safe
        includeFullBody: true
    });
    
    console.log(`Found ${result.count} Visa emails\n`);
    
    // Sort by date (oldest first)
    const sorted = result.emails.sort((a, b) => {
        const dateA = new Date(a.received_datetime || 0);
        const dateB = new Date(b.received_datetime || 0);
        return dateA - dateB;
    });
    
    // Show the earliest ones (likely to contain original SOW/proposal)
    console.log('=== EARLIEST EMAILS IN CHRONOLOGICAL ORDER ===\n');
    
    for (let i = 0; i < Math.min(15, sorted.length); i++) {
        const email = sorted[i];
        const subject = email.subject || '(no subject)';
        const sender = `${email.sender || ''} <${email.sender_email || ''}>`.trim();
        const date = email.received_datetime ? new Date(email.received_datetime).toLocaleString() : 'unknown';
        const attachments = email.has_attachments ? ` (has ${email.attachments_count} attachment(s))` : '';
        
        console.log(`${i + 1}. ${subject}${attachments}`);
        console.log(`   From: ${sender}, Date: ${date}`);
        
        // Check body for keywords
        const bodyLower = (email.body || '').toLowerCase();
        const keywords = ['statement of work', 'sow', 'scope', 'proposal', 'agreement', 'signed', 
                         'attachment', 'document', 'pdf', 'contract'];
        const foundKeywords = keywords.filter(k => bodyLower.includes(k));
        
        if (foundKeywords.length > 0) {
            console.log(`   Keywords in body: ${foundKeywords.join(', ')}`);
        }
        
        // Show preview of body
        if (email.body && email.body.length > 50) {
            const preview = email.body.substring(0, 250).replace(/\n/g, ' ').trim();
            console.log(`   Preview: ${preview}...`);
        }
        
        // Check for links in body
        const links = (email.html_body || email.body || '').match(/https?:\/\/[^\s<"')]+/g) || [];
        if (links.length > 0) {
            console.log(`   Links: ${links.slice(0, 3).join(', ')}`);
        }
        
        console.log('');
    }
}

runTest().catch(e => {
    console.error('Error:', e);
    process.exit(1);
});
