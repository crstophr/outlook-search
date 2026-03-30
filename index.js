/**
 * Outlook Search - Main Entry Point
 * 
 * Interactive CLI for natural language email and document search.
 */

const readline = require('readline');
const { searchOutlook, formatEmailForDisplay } = require('./searcher');

/**
 * Interactive session manager
 */
class OutlookSearchSession {
    constructor() {
        this.context = {
            query: '',
            additionalContext: '',
            dateHint: null,
            previousResults: [],
            refinementCount: 0
        };
        
        this.rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });
    }
    
    /**
     * Start an interactive search session
     */
    async start() {
        console.log('\n' + '='.repeat(70));
        console.log('  OUTLOOK SEARCH - Interactive Document Finder');
        console.log('='.repeat(70));
        console.log('\nI\'ll help you find emails and documents in your Outlook.');
        console.log('Describe what you\'re looking for in natural language.\n');
        
        // Initial query
        const initialQuery = await this.prompt(
            "What are you looking for? (e.g., 'Visa statement of work document')"
        );
        
        if (!initialQuery.trim()) {
            console.log('\nNo query provided. Exiting.');
            process.exit(0);
        }
        
        this.context.query = initialQuery;
        
        // Optional clarifying questions
        await this.askClarifyingQuestions();
        
        // Execute search
        await this.executeSearch();
        
        // Offer refinement
        await this.offerRefinement();
    }
    
    /**
     * Ask clarifying questions to narrow the search
     */
    async askClarifyingQuestions() {
        console.log('\n' + '-'.repeat(50));
        console.log('Let me ask a few questions to narrow this down:\n');
        
        // Date hint
        const dateAnswer = await this.prompt(
            "Do you remember approximately when this was? (e.g., 'February', 'last week', 'a month ago')"
        );
        
        if (dateAnswer && dateAnswer.trim()) {
            this.context.dateHint = dateAnswer.trim();
            console.log(`  → Noted: around ${this.context.dateHint}`);
        }
        
        // Additional context
        const contextAnswer = await this.prompt(
            "Any other details? (sender name, keywords, project names, etc.) [press enter to skip]"
        );
        
        if (contextAnswer && contextAnswer.trim()) {
            this.context.additionalContext = contextAnswer.trim();
            console.log(`  → Noted: ${this.context.additionalContext}`);
        }
        
        console.log('');
    }
    
    /**
     * Execute the search with gathered context
     */
    async executeSearch() {
        const searchContext = {
            query: this.context.query,
            context: this.context.additionalContext,
            dateHint: this.context.dateHint
        };
        
        console.log('\n' + '='.repeat(70));
        console.log('SEARCHING...');
        console.log('='.repeat(70));
        console.log(`\nLooking for: "${searchContext.query}"`);
        if (searchContext.dateHint) {
            console.log(`Timeframe: ${searchContext.dateHint}`);
        }
        if (searchContext.context) {
            console.log(`Additional context: ${searchContext.context}`);
        }
        console.log('');
        
        // Perform search
        const results = await searchOutlook(searchContext, {
            topN: 10,
            deepReadLimit: 5
        });
        
        // Display results
        this.displayResults(results);
    }
    
    /**
     * Display search results
     */
    displayResults(results) {
        console.log('\n' + '='.repeat(70));
        console.log('RESULTS');
        console.log('='.repeat(70));
        
        if (results.results.length === 0) {
            console.log('\n❌ No relevant results found.');
            console.log(results.summary);
            return;
        }
        
        console.log(`\n${results.summary}\n`);
        
        results.results.forEach((result, idx) => {
            console.log(`${idx + 1}. [Score: ${result.relevanceScore || 'N/A'}]`);
            console.log(formatEmailForDisplay(result));
            console.log('-'.repeat(50));
        });
    }
    
    /**
     * Offer to refine the search
     */
    async offerRefinement() {
        const answer = await this.prompt(
            '\nWant to refine this search? (yes/no)'
        );
        
        if (answer && answer.trim().toLowerCase().startsWith('y')) {
            await this.refineSearch();
        } else {
            console.log('\nThanks for using Outlook Search!');
            process.exit(0);
        }
    }
    
    /**
     * Refine the current search
     */
    async refineSearch() {
        console.log('\n' + '='.repeat(70));
        console.log('REFINEMENT');
        console.log('='.repeat(70));
        console.log('\nHow would you like to refine?');
        console.log('  1. Add more keywords/context');
        console.log('  2. Adjust date range');
        console.log('  3. Start over with new query');
        console.log('');
        
        const choice = await this.prompt('Choice (1/2/3): ');
        
        if (choice === '1') {
            const additional = await this.prompt('What else should I search for? ');
            this.context.query += ' ' + additional;
        } else if (choice === '2') {
            this.context.dateHint = await this.prompt('New date hint (e.g., "March", "last month"): ');
        } else if (choice === '3') {
            this.context.query = await this.prompt('New query: ');
            this.context.additionalContext = '';
            this.context.dateHint = null;
        }
        
        this.context.refinementCount++;
        await this.executeSearch();
        await this.offerRefinement();
    }
    
    /**
     * Helper for prompting user
     */
    prompt(question) {
        return new Promise(resolve => {
            this.rl.question(question + ' ', answer => {
                resolve(answer);
            });
        });
    }
}

/**
 * Main entry point
 */
async function main() {
    // Check if run from OpenClaw or standalone CLI
    if (process.argv.includes('--cli') || process.argv.length > 2) {
        // Interactive CLI mode
        const session = new OutlookSearchSession();
        await session.start();
    } else {
        // Called programmatically by OpenClaw
        console.log('Outlook Search: Ready for programmatic use');
    }
}

// Export for programmatic use
module.exports = {
    OutlookSearchSession,
    searchOutlook: require('./searcher').searchOutlook,
    formatEmailForDisplay
};

// Run if called directly
if (require.main === module) {
    main().catch(error => {
        console.error('Fatal error:', error);
        process.exit(1);
    });
}
