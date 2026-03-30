/**
 * Test different Outlook link formats
 */

const { searchEmails } = require('./bridge-client');

/**
 * Create Outlook deep link - simple EntryID format
 */
function createOutlookLinkSimple(entryId) {
    // Format: outlook:[EntryID]
    return `outlook:${encodeURIComponent(entryId)}`;
}

/**
 * Create Outlook link with the "#email" path format
 */
function createOutlookLinkWithEmailPath(entryId) {
    // Format: outlook://#email/[EntryID]
    return `outlook://#email/${encodeURIComponent(entryId)}`;
}

/**
 * Create mobile/iOS Outlook link
 */
function createMobileOutlookLink(itemid) {
    // Format for iOS: ms-outlook://mail/0/deeplink?ItemID=
    return `ms-outlook://mail/0/deeplink?ItemID=${encodeURIComponent(itemid)}`;
}

/**
 * Create Outlook Web App link (most compatible)
 */
function createOwaLink(itemId, subject = '') {
    // Works for OWA and can fall back to desktop on some systems
    const encodedSubject = encodeURIComponent(subject || '(Email)');
    return `https://outlook.office.com/mail/0/deeplink?subject=${encodedSubject}&itemid=${encodeURIComponent(itemId)}`;
}

async function runTest() {
    console.log('\nGetting a sample email to test different link formats...\n');
    
    const result = await searchEmails('Visa', {
        folder: 'inbox',
        limit: 1,
        days: 7, // Just recent ones
        includeFullBody: false
    });
    
    if (!result.emails || result.emails.length === 0) {
        console.log('No emails found.');
        return;
    }
    
    const email = result.emails[0];
    const entryId = email.id;
    const subject = email.subject || '(no subject)';
    
    console.log(`=== Testing with email: "${subject}" ===\n`);
    console.log(`EntryID: ${entryId}\n`);
    
    // Generate different link formats
    console.log('--- Different Link Formats ---\n');
    
    const simpleLink = createOutlookLinkSimple(entryId);
    console.log('1. Simple EntryID format (Windows Desktop):');
    console.log(`   ${simpleLink}\n`);
    
    const emailPathLink = createOutlookLinkWithEmailPath(entryId);
    console.log('2. #email path format (experimental):');
    console.log(`   ${emailPathLink}\n`);
    
    const mobileLink = createMobileOutlookLink(entryId);
    console.log('3. Mobile/iOS format:');
    console.log(`   ${mobileLink}\n`);
    
    const owaLink = createOwaLink(entryId, subject);
    console.log('4. Outlook Web App (most compatible):');
    console.log(`   ${owaLink}\n`);
    
    // Telegram-friendly format
    console.log('--- Telegram Clickable Format ---\n');
    console.log('[Open in Outlook](outlook:' + encodeURIComponent(entryId) + ')');
    console.log('\nOr the OWA version:');
    console.log(`[View Email](${owaLink})`);
}

runTest().catch(e => {
    console.error('Error:', e);
    process.exit(1);
});
