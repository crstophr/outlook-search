# Outlook Search

🔍 **Intelligent, conversational search for your Outlook emails and documents**

Stop wrestling with keyword searches. Just describe what you're looking for in natural language.

## Example

```bash
$ node index.js --cli

======================================================================
  OUTLOOK SEARCH - Interactive Document Finder
======================================================================

I'll help you find emails and documents in your Outlook.
Describe what you're looking for in natural language.

What are you looking for? (e.g., 'Visa statement of work document') Visa project statement of work

--------------------------------------------------
Let me ask a few questions to narrow this down:

Do you remember approximately when this was? (e.g., 'February', 'last week', 'a month ago') February
  → Noted: around February

Any other details? (sender name, keywords, project names, etc.) [press enter to skip] signed document
  → Noted: signed document


======================================================================
SEARCHING...
======================================================================

Looking for: "Visa project statement of work"
Timeframe: February
Additional context: signed document

📬 Phase 1: Broad search for "Visa project statement of work"
  Found 47 emails

🧠 Analyzing 47 emails with LLM...
  (LLM call would happen here)
After LLM filtering: 8 emails

📖 Phase 2: Deep reading top 5 emails...

======================================================================
RESULTS
======================================================================

Found 5 relevant emails from 47 matches

1. [Score: 9]
FROM: John Smith <john.smith@client.com>
DATE: Feb 15, 2026, 2:34 PM
SUBJECT: Re: Visa Project - Statement of Work (attached: SOW_Visa_Project_v2.pdf)
  Preview: Here's the updated statement of work as discussed. Please review and let me know if you have any questions...
  Links: https://sharepoint.example.com/sites/visa-project/Documents/SOW.docx
```

## Installation

```bash
cd skills/outlook-search
npm install
```

## Usage

### Interactive Mode (Recommended)

```bash
node index.js --cli
```

You'll be prompted to:
1. Describe what you're looking for
2. Provide approximate timeframe (optional)
3. Add any other details (optional)
4. Review results and optionally refine

### Programmatic Use

```javascript
const { searchOutlook } = require('./skills/outlook-search');

const results = await searchOutlook({
    query: "Visa statement of work",
    context: "looking for signed document or link to it", 
    dateHint: "February"
}, {
    topN: 10,        // Top N results to return
    deepReadLimit: 5 // How many emails to fully read
});

console.log(results.results);
```

## Features

### Natural Language Understanding
- **Query expansion**: "statement of work" → also searches "SOW", "scope of work"
- **Date parsing**: "February" → searches mid-January through mid-March (with buffer)
- **Context awareness**: Uses additional details to refine relevance scoring

### Multi-Phase Search
1. **Broad scan** — lightweight search across multiple query variants
2. **LLM filtering** — analyzes subjects/senders for relevance
3. **Deep read** — fetches full bodies of top candidates

### Smart Extraction
- Attachment filenames (without downloading)
- Links to SharePoint, OneDrive, etc.
- Relevance scoring (1-10)

### Iterative Refinement
Not quite what you wanted? The interactive mode lets you:
- Add more keywords/context
- Adjust the date range
- Start over with a new query

## Requirements

- Windows with Outlook Desktop app installed
- `outlook-mapi` skill configured and accessible
- Bridge server running (typically http://localhost:5000)
