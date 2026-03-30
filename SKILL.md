---
name: outlook-search
description: Intelligent natural language search through Outlook emails with LLM-powered relevance scoring, clickable deep links, and multi-phase analysis.
---

# Outlook Search Skill

An intelligent, conversational interface for finding emails and documents in Microsoft Outlook. Instead of wrestling with keyword searches, describe what you're looking for in natural language.

## What It Does

This skill provides **LLM-powered search** that understands vague queries like:

> "I'm looking for the Visa statement of work document - I think it was emailed to me around February"

And intelligently:
1. **Expands search terms** ("statement of work" → also searches "SOW", "scope of work")
2. **Applies date buffers** ("February" → mid-January through mid-March)
3. **Reads email bodies** for context, not just subjects
4. **Scores results by relevance** using the configured LLM
5. **Extracts links** and attachment information
6. **Provides clickable deep links** that open emails directly in Outlook

## Usage Examples

### Basic Search (CLI)

```bash
# From CLI (interactive mode)
cd skills/outlook-search
node index.js --cli
```

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
```

### Quick Test Script

```bash
node test-visa.js  # Searches for recent Visa-related emails
```

## Features

### 🔍 Smart Query Expansion

Converts natural language into comprehensive search queries:
- "statement of work" → searches `"SOW" OR "scope of work" OR "work order"`+
- "contract" → also searches `"agreement" OR "terms" OR "proposal"`
- Automatically adds document indicators (`"attachment"`, `"pdf"`, etc.)

### 📅 Natural Language Date Parsing

Understands vague time references:
- "February" → Jan 15 to Mar 15 (2-week buffer on each side)
- "last week" → Monday through today
- "a month ago" → ~30 days back with buffer
- "recent" or no date → last 90 days

### 🧠 LLM-Powered Relevance Scoring

Uses your **currently configured OpenClaw model** to:
1. Analyze email subjects and senders for semantic relevance
2. Score each email from 1-10 based on likelihood of matching your intent
3. Filter to top candidates (score ≥ 4 by default)
4. Detect document mentions in context

The LLM integration **automatically uses whatever provider you've configured** in OpenClaw (llamacpp, Anthropic, OpenAI, etc.)

### 🔗 Clickable Deep Links

Each result includes a **[Click here to open email](link)** that:
- Opens directly in Outlook Web App or desktop app
- Works across platforms (Windows, Mac, mobile)
- Renders as clickable in Telegram, Discord, Slack
- Uses OWA deeplinks for maximum compatibility

## Architecture

```
outlook-search (this skill)
    ├── index.js          → Main entry point + interactive CLI
    ├── searcher.js       → Multi-phase search orchestration
    ├── query-expander.js → Natural language → search queries
    ├── date-parser.js    → "February" → date range with buffer
    ├── llm.js            → LLM integration (relevance scoring)
    └── bridge-client.js  → HTTP client for Outlook bridge
            ↓
outlook-mapi / outlook-bridge
            ↓
Windows Outlook Desktop App (via MAPI)
```

### Module Breakdown

| File | Purpose |
|------|--------|
| `query-expander.js` | Converts "statement of work" → expanded search terms |
| `date-parser.js` | Parses natural language dates with intelligent buffering |
| `llm.js` | LLM-powered relevance analysis (uses configured OpenClaw model) |
| `bridge-client.js` | HTTP client for Outlook bridge (`localhost:8765`) |
| `searcher.js` | Orchestrates all phases, formats results |
| `index.js` | Interactive CLI entry point |

### Search Phases

**Phase 1: Broad Scan** 📬
- Expands query into multiple variants
- Searches with relaxed criteria (subject only for speed)
- Returns subjects, senders, dates, attachment flags

**Phase 2: LLM Filtering** 🧠
- Sends email summaries to configured LLM
- Gets relevance scores (1-10) and reasoning
- Filters to emails with score ≥ 4

**Phase 3: Deep Read** 📖
- Fetches full email bodies for top N candidates
- Extracts links from body content
- Presents detailed results with clickbale links

## Output Format

Each result includes:
- **Subject, Sender, Date**
- **Relevance Score** (1-10) with reasoning
- **Attachment info** (count if any)
- **Clickable deep link** to open email directly
- **Body preview** (first 200 characters, if available)

### Example Output

```
1. [Score: 9]
   Subject: Re: Visa MAM Sync Up meeting (3 attachments)
   From: Jesse Ramirez <jesramir@visa.com>
   Date: Mar 30, 2026 at 2:56 PM
   📧 [Open Email](https://outlook.office.com/mail/0/deeplink?...) ← Click this!
   Why: Subject mentions both "Visa" and appears to be in a technical discussion thread
```

## Configuration

No special configuration needed:
- **Outlook bridge**: Auto-connects to `http://localhost:8765` (set via `BRIDGE_HOST`/`BRIDGE_PORT` env vars)
- **LLM provider**: Reads from OpenClaw config at runtime (`agents.defaults.model.primary`)

## Limitations & Notes

- **Requires Windows with Outlook Desktop** running for the bridge to work
- **Bridge server** must be running: `python outlook-bridge.py`
- **LLM analysis adds latency** (typically 2-5 seconds per email batch)
- **Deep reading limited** to top 5 results by default (configurable via `deepReadLimit`)
- **OWA deeplinks** work in browsers and often redirect to desktop app; native `outlook:` protocol is unreliable since Outlook 2013+

## Why OWA Deep Links Instead of `outlook:` Protocol?

The native Outlook desktop protocol handler (`outlook:`) has been **unreliable since Outlook 2013+** due to Microsoft removing many handlers for security reasons.

**OWA deeplinks work better because they:**
- ✅ Work in any browser directly
- ✅ Often auto-redirect to desktop Outlook when installed on Windows
- ✅ Compatible across platforms (Windows, Mac, mobile)
- ✅ Render as clickable links in Telegram, Discord, Slack, etc.
