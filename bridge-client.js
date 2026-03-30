/**
 * Outlook Bridge Client
 * 
 * HTTP client for communicating with the Windows Outlook bridge server.
 */

const BRIDGE_CONFIG = {
    host: process.env.BRIDGE_HOST || 'localhost',
    port: parseInt(process.env.BRIDGE_PORT) || 8765
};

/**
 * Make HTTP request to bridge
 */
async function bridgeRequest(path, options = {}) {
    const url = `http://${BRIDGE_CONFIG.host}:${BRIDGE_CONFIG.port}${path}`;
    
    const defaults = {
        headers: {
            'Content-Type': 'application/json',
        },
    };
    
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 10000);

    try {
        const response = await fetch(url, { ...defaults, ...options, signal: controller.signal });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Bridge returned ${response.status}: ${errorText}`);
        }

        return await response.json();
    } catch (error) {
        if (error.name === 'AbortError') {
            throw new Error('Outlook bridge request timed out after 10 seconds');
        }
        if (error.code === 'ECONNREFUSED') {
            throw new Error(
                'Cannot connect to Outlook bridge server. Make sure outlook-bridge.py is running on Windows:\n' +
                '  python ~/repos/OpenClaw-Outlook-MAPI/scripts/outlook-bridge.py'
            );
        }
        throw error;
    } finally {
        clearTimeout(timeout);
    }
}

/**
 * Health check
 */
async function healthCheck() {
    return bridgeRequest('/health');
}

/**
 * Search for emails
 */
async function searchEmails(query = '', options = {}) {
    const {
        folder = 'inbox',
        limit = 100,
        unreadOnly = false,
        days = null,
        startDate = null,
        endDate = null,
        includeFullBody = false
    } = options;

    return bridgeRequest('/api/email/search', {
        method: 'POST',
        body: JSON.stringify({
            query, folder, limit,
            unread_only: unreadOnly,
            days,
            start_date: startDate ? startDate.toISOString().split('T')[0] : null,
            end_date: endDate ? endDate.toISOString().split('T')[0] : null,
            include_full_body: includeFullBody
        }),
    });
}

/**
 * Get recent emails with richer metadata for downstream processing
 */
async function getRecentEmails(options = {}) {
    const { folder = 'inbox', days = 14, limit = 100, includeFullBody = false } = options;
    
    return bridgeRequest('/api/email/recent', {
        method: 'POST',
        body: JSON.stringify({ folder, days, limit, include_full_body: includeFullBody }),
    });
}

/**
 * Get single email by EntryID (with full body)
 */
async function getEmailById(entryId) {
    return bridgeRequest(`/api/email/${encodeURIComponent(entryId)}`);
}

/**
 * Format Outlook's date format to JavaScript Date
 */
function parseOutlookDate(dateStr) {
    if (!dateStr) return null;
    
    // Handle various formats from Outlook
    try {
        // ISO format (from bridge)
        if (dateStr.includes('T')) {
            return new Date(dateStr);
        }
        
        // MM/DD/YYYY format
        const match = dateStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (match) {
            return new Date(parseInt(match[3]), parseInt(match[1]) - 1, parseInt(match[2]));
        }
        
        return new Date(dateStr);
    } catch (e) {
        return null;
    }
}

/**
 * Extract links from email body
 */
function extractLinks(text) {
    if (!text) return [];
    
    // Match http/https URLs
    const httpUrls = text.match(/https?:\/\/[^\s<"')]+[\w\/]/g) || [];
    
    // Deduplicate and return
    return [...new Set(httpUrls)].slice(0, 10);
}

/**
 * Create a clickable deep link to open a specific email in Outlook
 * 
 * Uses the Outlook Web App (OWA) deeplink format which:
 * - Works in any browser
 * - Often redirects to desktop Outlook app when installed on Windows
 * - Is compatible across platforms and Outlook versions
 * 
 * Note: The native "outlook:" protocol handler is unreliable since Outlook 2013+.
 */
function createEmailLink(entryId, subject = '') {
    // OWA deeplink format
    const encodedSubject = encodeURIComponent(subject || '(Email)');
    return `https://outlook.office.com/mail/0/deeplink?subject=${encodedSubject}&itemid=${encodeURIComponent(entryId)}`;
}

module.exports = {
    bridgeRequest,
    healthCheck,
    searchEmails,
    getRecentEmails,
    getEmailById,
    parseOutlookDate,
    extractLinks,
    createEmailLink,
    BRIDGE_CONFIG
};
