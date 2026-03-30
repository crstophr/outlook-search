const readline = require('readline');

/**
 * Outlook Search Skill
 * 
 * A CLI tool for searching emails and calendar events in Microsoft Outlook.
 * Wraps the outlook-mapi skill to provide unified search capabilities.
 */

// Import outlook-mapi functions
async function initOutlookMapi() {
    try {
        const mapiModule = require('outlook-mapi');
        return new mapiModule.OutlookMapi();
    } catch (error) {
        console.error('Error initializing Outlook MAPI:', error.message);
        process.exit(1);
    }
}

/**
 * Search emails in a specific folder
 */
async function searchEmails(mapi, query, options = {}) {
    const { folder = 'inbox', limit = 10, unreadOnly = false, importance = null } = options;
    
    console.log(`\nSearching emails in ${folder}...`);
    console.log(`Query: ${query || '(all)'}`);
    
    try {
        // Get folder ID
        const folderId = await mapi.getFolderId(folder);
        if (!folderId) {
            console.error(`Folder not found: ${folder}`);
            return [];
        }
        
        // Build search query
        let searchQuery = '';
        if (query) {
            if (query.startsWith('from:')) {
                const sender = query.replace('from:', '');
                searchQuery = `sender:${sender}`;
            } else if (query.startsWith('to:')) {
                const recipient = query.replace('to:', '');
                searchQuery = `recipient:${recipient}`;
            } else {
                searchQuery = query;
            }
        }
        
        // Add unread filter if needed
        if (unreadOnly) {
            searchQuery = searchQuery ? `${searchQuery} AND isRead:false` : 'isRead:false';
        }
        
        // Add importance filter if needed
        if (importance) {
            searchQuery = searchQuery 
                ? `${searchQuery} AND importance:${importance}` 
                : `importance:${importance}`;
        }
        
        // Execute search
        const results = await mapi.searchMailItems({
            folderId,
            query: searchQuery,
            top: limit,
            properties: ['subject', 'senderEmail', 'receivedDateTime', 'isRead', 'importance']
        });
        
        return results.items || [];
    } catch (error) {
        console.error('Error searching emails:', error.message);
        return [];
    }
}

/**
 * Search calendar events
 */
async function searchCalendar(mapi, query, options = {}) {
    const { days = null, limit = 10 } = options;
    
    console.log(`\nSearching calendar...`);
    console.log(`Query: ${query || '(all)'}`);
    
    try {
        // Set date range if specified
        let startDate, endDate;
        if (days) {
            startDate = new Date(Date.now() - days * 24 * 60 * 60 * 1000).toISOString();
            endDate = new Date().toISOString();
        } else {
            // Default: today to next 30 days
            startDate = new Date().toISOString();
            endDate = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();
        }
        
        // Build search query for calendar
        let searchQuery = `
            start/ge '${startDate}' and start/le '${endDate}'
        `;
        
        if (query) {
            searchQuery += ` and (${query})`;
        }
        
        const results = await mapi.searchEvents({
            query: searchQuery,
            top: limit,
            properties: ['subject', 'start', 'end', 'location', 'attendees']
        });
        
        return results.items || [];
    } catch (error) {
        console.error('Error searching calendar:', error.message);
        return [];
    }
}

/**
 * Format and display search results
 */
function formatEmail(email) {
    const subject = email.subject || '(no subject)';
    const sender = email.senderEmail || '?';
    const date = email.receivedDateTime 
        ? new Date(email.receivedDateTime).toLocaleString()
        : 'unknown';
    const readStatus = email.isRead ? '' : '[unread]';
    const importance = email.importance ? ` [${email.importance}]` : '';
    
    return `${readStatus}${importance} ${sender}: ${subject} (${date})`;
}

function formatEvent(event) {
    const subject = event.subject || '(no subject)';
    const start = event.start?.dateTime 
        ? new Date(event.start.dateTime).toLocaleString()
        : 'unknown';
    const location = event.location ? ` @ ${event.location}` : '';
    
    return `${subject}${location} (${start})`;
}

/**
 * Parse command line arguments
 */
function parseArgs(argv) {
    const args = {
        type: 'email',
        query: '',
        folder: 'inbox',
        unreadOnly: false,
        importance: null,
        days: null,
        limit: 10,
        format: 'pretty'
    };
    
    for (let i = 2; i < argv.length; i++) {
        const arg = argv[i];
        
        switch (arg) {
            case '--type':
                args.type = argv[++i] || 'email';
                break;
            case '--query':
            case '-q':
                args.query = argv[++i] || '';
                break;
            case '--folder':
            case '-f':
                args.folder = argv[++i] || 'inbox';
                break;
            case '--unread-only':
            case '-u':
                args.unreadOnly = true;
                break;
            case '--importance':
                args.importance = argv[++i];
                break;
            case '--days':
                args.days = parseInt(argv[++i], 10);
                break;
            case '--limit':
            case '-l':
                args.limit = parseInt(argv[++i], 10);
                break;
            case '--format':
                args.format = argv[++i] || 'pretty';
                break;
            default:
                if (!arg.startsWith('-')) {
                    args.query = arg;
                }
        }
    }
    
    return args;
}

/**
 * Main entry point
 */
async function main() {
    const options = parseArgs(process.argv);
    
    console.log('Outlook Search Tool');
    console.log('==================');
    console.log();
    
    // Initialize MAPI client
    const mapi = await initOutlookMapi();
    
    const allResults = [];
    
    // Search emails if requested
    if (options.type === 'email' || options.type === 'all') {
        console.log('\n--- EMAIL SEARCH ---');
        const emails = await searchEmails(mapi, options.query, {
            folder: options.folder,
            limit: options.limit,
            unreadOnly: options.unreadOnly,
            importance: options.importance
        });
        
        if (options.format === 'json') {
            allResults.emails = emails;
        } else {
            console.log(`\nFound ${emails.length} email(s):\n`);
            emails.forEach(email => {
                console.log(formatEmail(email));
            });
        }
    }
    
    // Search calendar if requested
    if (options.type === 'calendar' || options.type === 'all') {
        console.log('\n--- CALENDAR SEARCH ---');
        const events = await searchCalendar(mapi, options.query, {
            days: options.days,
            limit: options.limit
        });
        
        if (options.format === 'json') {
            allResults.events = events;
        } else {
            console.log(`\nFound ${events.length} event(s):\n`);
            events.forEach(event => {
                console.log(formatEvent(event));
            });
        }
    }
    
    // Output JSON if requested
    if (options.format === 'json') {
        console.log('\n' + JSON.stringify(allResults, null, 2));
    }
}

// Run main function
main().catch(error => {
    console.error('Fatal error:', error);
    process.exit(1);
});
