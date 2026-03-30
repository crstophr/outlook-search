/**
 * Create an Outlook deep link from email EntryID
 * 
 * For Windows Desktop (Outlook.exe), the protocol is:
 *   outlook:[EntryID]
 * 
 * The EntryID must be URL-encoded for safe transmission.
 */
function createOutlookLink(entryId) {
    if (!entryId) return null;
    
    // URL-encode the EntryID for safety
    const encodedId = encodeURIComponent(entryId);
    
    // Windows desktop format - uses "outlook:" protocol (not outlook://)
    return `outlook:${encodedId}`;
}

/**
 * Create a mobile/OWA link as fallback
 */
function createOwaLink(itemId, subject = '') {
    // For Outlook on the Web / mobile apps
    const encodedSubject = encodeURIComponent(subject);
    const encodedItemId = encodeURIComponent(itemId);
    return `https://outlook.office.com/mail/0/deeplink?subject=${encodedSubject}&itemid=${encodedItemId}`;
}

/**
 * Format a link for Telegram output (markdown clickable)
 */
function formatTelegramLink(entryId, displayName) {
    const outlookUrl = createOutlookLink(entryId);
    if (!outlookUrl) return null;
    
    // Telegram markdown format: [Display Text](URL)
    return `[${displayName}](outlook:${encodeURIComponent(entryId)})`;
}

module.exports = {
    createOutlookLink,
    createOwaLink,
    formatTelegramLink
};
