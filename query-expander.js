/**
 * Query Expander
 * 
 * Takes a natural language search intent and expands it into
 * comprehensive Outlook search queries with synonyms, abbreviations,
 * and related terms.
 */

const SYNONYM_MAP = {
    "statement of work": ["SOW", "scope of work", "work order", "Statement of Work"],
    "contract": ["agreement", "SOW", "terms", "proposal"],
    "invoice": ["bill", "payment", "quote", "purchase order", "PO"],
    "meeting": ["call", "discussion", "conference", "zoom", "teams"],
    "report": ["summary", "analysis", "findings", "results"],
    "budget": ["cost", "estimate", "pricing", "quote", "proposal"],
    "deadline": ["due date", "timeline", "schedule", "milestone"],
    "approval": ["sign-off", "authorize", "confirmed", "approved"],
    "document": ["file", "attachment", "pdf", "docx", "word", "excel"],
    "project": ["engagement", "work", "task", "initiative"]
};

/**
 * Expand a single search term into related terms
 */
function expandTerm(term) {
    const lowerTerm = term.toLowerCase().trim();
    
    // Check if this is a compound phrase
    for (const [key, synonyms] of Object.entries(SYNONYM_MAP)) {
        if (lowerTerm.includes(key) || key.includes(lowerTerm)) {
            return [...new Set([term, ...synonyms])];
        }
    }
    
    // Check individual words
    const words = lowerTerm.split(/\s+/);
    let expanded = new Set([term]);
    
    for (const word of words) {
        for (const [key, synonyms] of Object.entries(SYNONYM_MAP)) {
            if (word === key.toLowerCase()) {
                synonyms.forEach(s => expanded.add(s));
            }
        }
    }
    
    return Array.from(expanded);
}

/**
 * Build an Outlook-compatible search query from expanded terms
 */
function buildSearchQuery(query, options = {}) {
    const { expand: shouldExpand = true } = options;
    
    let terms = [query];
    
    if (shouldExpand) {
        terms = expandTerm(query);
    }
    
    // Build query with OR between synonyms
    const termQuery = terms.map(t => `"${t.replace(/"/g, '\\"')}"`).join(" OR ");
    
    return termQuery;
}

/**
 * Add document type filters to a query
 */
function addDocumentFilters(query, types = ["pdf", "docx", "doc"]) {
    const extensions = types.map(ext => `":.${ext}"`).join(" OR ");
    return `${query} OR (${extensions})`;
}

/**
 * Add common document-related keywords to search
 */
function addDocumentKeywords(query) {
    const docKeywords = [
        "attachment",
        "attached",
        "document",
        "file",
        "pdf",
        "download",
        "see attached",
        "find attached",
        "please find"
    ];
    
    const keywords = docKeywords.map(k => `"${k}"`).join(" OR ");
    return `${query} AND (${keywords})`;
}

/**
 * Main function: expand a search intent into multiple queries
 */
function expandSearchIntent(intent, options = {}) {
    const { includeDocumentFilters = true } = options;
    
    // Base expanded query
    let baseQuery = buildSearchQuery(intent);
    
    // Add document filters if requested
    if (includeDocumentFilters) {
        baseQuery += " OR (\"attachment\" OR \"attached\" OR \"document\" OR \"pdf\")";
    }
    
    return {
        primary: baseQuery,
        variants: [
            baseQuery,
            buildSearchQuery(intent) + " AND (\"sharepoint\" OR \"onedrive\" OR \"microsoft.com\")",
            buildSearchQuery(intent) + " AND (\"see attached\" OR \"find attached\" OR \"please find\")"
        ]
    };
}

module.exports = {
    expandTerm,
    buildSearchQuery,
    addDocumentFilters,
    addDocumentKeywords,
    expandSearchIntent
};
