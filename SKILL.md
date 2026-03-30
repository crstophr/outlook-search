---
name: outlook-search
description: Interactive natural language search through Outlook inbox, calendar, and attachments. Use when users want to find specific emails or documents using conversational queries rather than precise keywords. Supports iterative refinement and LLM-powered relevance scoring.
---

# Outlook Search Skill

An intelligent, conversational interface for finding emails and documents in Microsoft Outlook. Instead of wrestling with keyword searches, describe what you're looking for in natural language.

## What It Does

This skill provides **LLM-powered search** that understands vague queries like:

> "I'm looking for the Visa statement of work document - I think it was emailed to me around February"

And intelligently:
1. Expands search terms ("statement of work" → also searches "SOW", "scope of work")
2. Applies date buffers ("February" → mid-January through mid-March)
3. Reads email bodies for context, not just subjects
4. Scores results by relevance using an LLM
5. Extracts links and attachment filenames
6. Supports iterative refinement in conversation

## Usage Examples

### Basic Search

```bash
# From CLI (interactive mode)
node index.js --cli

# Or programmatically from OpenClaw:
const { searchOutlook } = require('./skills/outlook-search');
const results = await searchOutlook({
    query: "Visa statement of work",
    context: "looking for signed document or link to it",
    dateHint: "February"
});
```

### What You Can Ask

**Documents:**
- "Find the Visa project statement of work"
- "Where's that invoice from Acme Corp?"
- "Search for the NDA I signed last month"

**With Context:**
- "Email about the quarterly budget review"  
- "Something from Sarah about the migration timeline"
- "The document with the system requirements - think it had a SharePoint link"

**Date-Based:**
- "That proposal email from last week"
- "Anything about the contract signed in March"
- "Recent emails with attachments from the finance team"

## Interactive Mode Features

When you run `node index.js --cli`, you enter an interactive session:

1. **Describe what you're looking for**
   
   *"Find that Visa statement of work document"*

2. **Answer clarifying questions (optional)**
   
   - "Do you remember approximately when this was?"
   - "Any other details? (sender name, keywords, project names)"

3. **Review results** - sorted by relevance score

4. **Refine if needed:**
   - Add more keywords/context
   - Adjust date range
   - Start over with new query

## Architecture

```
outlook-search (this skill)
    ├── index.js          → Main entry point + interactive CLI
    ├── searcher.js       → Multi-phase search orchestration
    ├── query-expander.js → Natural language → search queries
    └── date-parser.js    → "February" → date range with buffer
            ↓
outlook-mapi (Outlook API client)
            ↓
Outlook Bridge (Python/MAPI)
            ↓
Windows Outlook Desktop App
```

### Search Phases

**Phase 1: Broad Scan**
- Expands query into multiple variants
- Searches with relaxed criteria
- Returns subjects, senders, dates only

**Phase 2: LLM Filtering**
- Analyzes summaries for relevance
- Scores each email (1-10)
- Filters to top candidates

**Phase 3: Deep Read**
- Fetches full email bodies
- Extracts links and attachment names
- Presents detailed results

## Output Format

Each result includes:
- **Subject, Sender, Date**
- **Relevance Score** (1-10)
- **Attachment filenames** (if any)
- **Extracted links** (SharePoint, OneDrive, etc.)
- **Body preview** (first 200 characters)

Example:
```
1. [Score: 9]
FROM: John Smith <john.smith@client.com>
DATE: Feb 15, 2026, 2:34 PM
SUBJECT: Re: Visa Project - Statement of Work (attached: SOW_Visa_Project_v2.pdf)
  Preview: Here's the updated statement of work as discussed. Please review and let me know if you have any questions...
  Links: https://sharepoint.example.com/sites/visa-project/Documents/SOW.docx
```

## Configuration

No special configuration needed - inherits Outlook connection settings from outlook-mapi.

## Limitations

- Requires Windows with Outlook Desktop app
- LLM analysis adds latency (typically 2-5 seconds per phase)
- Deep reading limited to top 5 results by default (configurable)
