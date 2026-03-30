/**
 * Date Parser
 * 
 * Converts natural language date references into searchable date ranges
 * with intelligent buffering around uncertain dates.
 */

const MONTH_NAMES = {
    "january": 0, "jan": 0,
    "february": 1, "feb": 1,
    "march": 2, "mar": 2,
    "april": 3, "apr": 3,
    "may": 4,
    "june": 5, "jun": 5,
    "july": 6, "jul": 6,
    "august": 7, "aug": 7,
    "september": 8, "sep": 8, "sept": 8,
    "october": 9, "oct": 9,
    "november": 10, "nov": 10,
    "december": 11, "dec": 11
};

/**
 * Add buffer days to a date range
 */
function addBuffer(startDate, endDate, bufferDays = 14) {
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    start.setDate(start.getDate() - bufferDays);
    end.setDate(end.getDate() + bufferDays);
    
    return { startDate: start, endDate: end };
}

/**
 * Parse a month reference ("February") into a date range with buffer
 */
function parseMonth(monthName) {
    const monthKey = monthName.toLowerCase().trim();
    const monthIndex = MONTH_NAMES[monthKey];
    
    if (monthIndex === undefined) {
        return null;
    }
    
    const now = new Date();
    let year = now.getFullYear();

    // If the month hasn't happened yet this year, use last year
    if (monthIndex > now.getMonth()) {
        year -= 1;
    }

    // Start of the month
    const startDate = new Date(year, monthIndex, 1);
    // End of the month
    const endDate = new Date(year, monthIndex + 1, 0);
    
    // Add 2-week buffer on each side
    return addBuffer(startDate, endDate, 14);
}

/**
 * Parse relative time references ("last week", "a month ago")
 */
function parseRelativeTime(reference) {
    const ref = reference.toLowerCase().trim();
    const now = new Date();
    
    let startDate, endDate;
    
    if (ref.includes("today")) {
        startDate = new Date(now.setHours(0, 0, 0, 0));
        endDate = new Date();
    } else if (ref.includes("yesterday")) {
        startDate = new Date(now);
        startDate.setDate(startDate.getDate() - 1);
        endDate = new Date(startDate);
        endDate.setHours(23, 59, 59, 999);
    } else if (ref.includes("last week") || ref.includes("this week")) {
        const dayOfWeek = now.getDay();
        startDate = new Date(now);
        startDate.setDate(startDate.getDate() - dayOfWeek); // Monday
        endDate = new Date();
    } else if (ref.includes("last month") || ref.includes("this month")) {
        startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        endDate = new Date();
    } else if (ref.match(/(\d+) days? ago/)) {
        const days = parseInt(ref.match(/(\d+)/)[1]);
        startDate = new Date(now);
        startDate.setDate(startDate.getDate() - days - 7); // Add buffer
        endDate = new Date();
    } else if (ref.match(/(\d+) weeks? ago/)) {
        const weeks = parseInt(ref.match(/(\d+)/)[1]);
        startDate = new Date(now);
        startDate.setDate(startDate.getDate() - weeks * 7 - 7);
        endDate = new Date();
    } else if (ref.match(/(\d+) months? ago/)) {
        const months = parseInt(ref.match(/(\d+)/)[1]);
        startDate = new Date(now);
        startDate.setMonth(startDate.getMonth() - months - 1);
        endDate = new Date();
    } else if (ref.includes("recent") || ref.includes("last few days")) {
        startDate = new Date(now);
        startDate.setDate(startDate.getDate() - 7);
        endDate = new Date();
    } else {
        return null;
    }
    
    return { startDate, endDate };
}

/**
 * Parse "last N days" format
 */
function parseDaysAgo(days) {
    const endDate = new Date();
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - days - 7); // Add buffer
    
    return { startDate, endDate };
}

/**
 * Parse a natural language date into Outlook-compatible ISO strings
 */
function parseDateToISO(dateRange) {
    if (!dateRange || !dateRange.startDate) {
        return null;
    }
    
    return {
        start: dateRange.startDate.toISOString().split('T')[0],
        end: dateRange.endDate.toISOString().split('T')[0]
    };
}

/**
 * Main parser: takes natural language and returns a search-friendly date range
 */
function parseDate(naturalLanguage) {
    const text = naturalLanguage.toLowerCase().trim();
    
    // Try exact month first ("February")
    for (const [name, _] of Object.entries(MONTH_NAMES)) {
        if (text === name || text.includes(`in ${name}`) || text.includes(`${name} `)) {
            const range = parseMonth(name);
            return {
                range,
                iso: parseDateToISO(range),
                description: `${range.startDate.toLocaleDateString()} to ${range.endDate.toLocaleDateString()} (with buffer)`
            };
        }
    }
    
    // Try relative time ("last week", "a month ago")
    const relativeResult = parseRelativeTime(text);
    if (relativeResult) {
        return {
            range: relativeResult,
            iso: parseDateToISO(relativeResult),
            description: `${relativeResult.startDate.toLocaleDateString()} to ${relativeResult.endDate.toLocaleDateString()}`
        };
    }
    
    // Try "last N days" pattern
    if (text.match(/last (\d+) days?/)) {
        const days = parseInt(text.match(/last (\d+)/)[1]);
        const range = parseDaysAgo(days);
        return {
            range,
            iso: parseDateToISO(range),
            description: `${range.startDate.toLocaleDateString()} to ${range.endDate.toLocaleDateString()}`
        };
    }
    
    // Default: last 90 days with buffer
    const defaultEnd = new Date();
    const defaultStart = new Date(defaultEnd);
    defaultStart.setDate(defaultStart.getDate() - 90);
    
    return {
        range: { startDate: defaultStart, endDate: defaultEnd },
        iso: parseDateToISO({ startDate: defaultStart, endDate: defaultEnd }),
        description: `Last 90 days (default)`
    };
}

module.exports = {
    parseDate,
    parseMonth,
    parseRelativeTime,
    parseDaysAgo,
    parseDateToISO,
    addBuffer
};
