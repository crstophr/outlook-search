# 🔍 Outlook Search

> **Intelligent, conversational search for your Outlook emails and documents**

Stop wrestling with keyword searches. Just describe what you're looking for in natural language.

## 🎯 Quick Start

```bash
# Clone or navigate to the repo
cd skills/outlook-search

# Install dependencies (links to outlook-mapi)
npm install

# Run a quick test search
node test-visa.js  # Searches for recent Visa-related emails
```

## 📖 Example Output

```bash
$ node test-visa.js

Searching for: "Visa statement of work SOW"

============================================================
OUTLOOK SEARCH
============================================================
Query: Visa
Context: looking for signed document or detailed description...

✓ Bridge connected to Outlook
Date range: Last 90 days (default)

📬 Phase 1: Broad search for "Visa"...
  Found 48 emails

🧠 Analyzing 48 emails with LLM...
  Using model: Qwen3.5-27B-Claude-4.6-Opus... (llamacpp)
  Found 8 relevant emails (score >= 4)

======================================================================
RESULTS WITH CLICKABLE LINKS
======================================================================

Found 8 relevant emails from 48 matches

1. [Score: 9]
   Subject: Re: Visa MAM Sync Up meeting (3 attachments)
   From: Jesse Ramirez <jesramir@visa.com>
   Date: Mar 30, 2026 at 2:56 PM
   📧 [Open Email](https://outlook.office.com/mail/0/deeplink?...) ← CLICK THIS!
   Why: Subject mentions "Visa" and appears to be in a technical thread

2. [Score: 8]
   Subject: Re: Visa MAM Sync Up meeting (3 attachments)
   From: Beny Korotkin <bkorotkin@onediversified.com>
   Date: Mar 17, 2026 at 4:35 PM
   📧 [Open Email](https://outlook.office.com/mail/0/deeplink?...)
   Why: From primary contact (Beny), mentions MAM system configuration
```

## ✨ How It Works

### Natural Language Understanding

You describe what you're looking for:
> "Find the Visa statement of work - I think it was around February"

The skill intelligently expands this to:
- **Query**: `"Visa" AND ("statement of work" OR "SOW" OR "scope of work")`
- **Date range**: Jan 15 – Mar 15, 2026 (with buffer)
- **Document indicators**: Also searches for attachments and document mentions

### Multi-Phase Search

1. **📬 Broad Scan** — Searches with expanded queries (fast, subject-only)
2. **🧠 LLM Filtering** — Analyzes results for relevance (scores 1-10)
3. **📖 Deep Read** — Fetches full bodies of top candidates

### Smart Features

| Feature | How It Helps |
|---------|-------------|
| **Query Expansion** | "statement of work" → also searches "SOW", "scope of work" |
| **Date Buffering** | "February" → actually searches Jan 15 through Mar 15 |
| **Relevance Scoring** | LLM ranks results by likelihood of matching your intent |
| **Clickable Links** | Open emails directly in Outlook with one click |

## 🛠️ Installation & Setup

### Prerequisites

1. **Windows machine** with Outlook Desktop app installed
2. **Outlook bridge running** on port 8765:
   ```bash
   # On Windows, run the bridge server
   python ~/repos/OpenClaw-Outlook-MAPI/scripts/outlook-bridge.py
   ```
3. **Node.js installed** (v14+)

### Install Dependencies

```bash
cd skills/outlook-search
npm install
```

This links to your local `outlook-mapi` repo. Make sure you have it at:
```
/home/openclaw/repos/OpenClaw-Outlook-MAPI/
```

## 📝 Usage Examples

### Simple Search (Test Scripts)

```bash
# Quick Visa search with links
node test-clickable-links.js

# Get recent emails about something
node test-simple.js

# Search for attachments
node test-visa-attach.js
```

### Programmatic Use in Your Code

```javascript
const { searchOutlook } = require('./searcher');

const results = await searchOutlook({
    query: "Visa statement of work",
    context: "looking for signed document or link to it",
    dateHint: "February"  // Optional - adds intelligent buffering
}, {
    topN: 10,         // Top N results
    deepReadLimit: 5  // How many to fully read
});

// Display with clickable links
results.results.forEach(result => {
    const formatted = formatEmailResult(result);
    console.log(`[📧 Open Email](${formatted.emailLink})`);
});
```

## 🔗 Clickable Deep Links

Each result includes a **clickable link** that opens the email directly:

```markdown
📧 [Open Email](https://outlook.office.com/mail/0/deeplink?subject=...&itemid=...)
```

**Why not use `outlook:` protocol?**
- The native `outlook:` protocol handler is unreliable since Outlook 2013+
- OWA deeplinks work in browsers AND often redirect to desktop app
- They render clickable in Telegram, Discord, Slack, etc.

## 🧠 LLM Integration

The skill automatically uses your **currently configured OpenClaw model** (whether you're using llamacpp, Anthropic, OpenAI, or another provider).

It intelligently:
- Analyzes email subjects and senders for semantic relevance
- Scores each email 1-10 based on how likely it matches your intent
- Detects document mentions in context
- Falls back to keyword-based analysis if LLM is unavailable

## 🐛 Troubleshooting

### "Cannot connect to Outlook bridge server"

Make sure the bridge is running:
```bash
python ~/repos/OpenClaw-Outlook-MAPI/scripts/outlook-bridge.py
```

Then verify:
```bash
curl http://localhost:8765/health
```

### "No results found" or low counts

Try simpler queries:
- Start with just the main keyword (e.g., "Visa")
- Don't over-specify - let the LLM help find relevance
- Use a longer date range if needed

### Links don't open in desktop Outlook

The OWA links open in your browser first. If you have Outlook installed and it's set as default mail client, it may automatically redirect. Otherwise, you'll view the email in Outlook Web App.

## 📦 Project Files

| File | Purpose |
|------|--------|
| `SKILL.md` | OpenClaw skill definition & technical docs |
| `README.md` | This file - user documentation |
| `searcher.js` | Multi-phase search orchestration |
| `query-expander.js` | Natural language → search queries |
| `date-parser.js` | "February" → date range with buffer |
| `llm.js` | LLM integration for relevance scoring |
| `bridge-client.js` | HTTP client for Outlook bridge |
| `index.js` | Interactive CLI entry point |

## 🚀 What's Next?

This is a **MVP implementation** with the core architecture in place. Future enhancements could include:
- Attachment downloading/preview
- Calendar event search integration
- More sophisticated query understanding
- Multi-folder/thread-aware search
