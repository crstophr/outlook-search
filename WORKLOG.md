# Work Log - Outlook Search Skill

## Overview

**Project:** OpenClaw skill for intelligent, conversational email search through Microsoft Outlook  
**Started:** March 30, 2026  
**Author:** Christopher (with Agent Smith/Clawd)  
**Repository:** https://github.com/crstophr/outlook-search

---

## Development Timeline

### March 30, 2026

#### 14:12 - Project Initiated
- **Goal:** Create a skill that uses outlook-mapi to search Outlook inbox for information
- **Use case:** Finding project documentation (Visa statement of work) using natural language queries
- **Key requirements:**
  - Search through email bodies, not just subjects
  - Use LLM for intelligent relevance scoring
  - Extract attachment filenames and links
  - Provide clickable deep links to open emails directly

#### 14:12 - Initial Architecture Proposed
Discussed multi-phase search approach:
- **Phase 1:** Broad scan with query expansion (subject/snippet level)
- **Phase 2:** LLM-powered filtering for relevance scoring
- **Phase 3:** Deep read of top candidates, extract links/attachments

Agreed on:
- Single cohesive skill architecture
- Heavy use of LLM for understanding and filtering
- Reading email bodies thoroughly
- Buffer around dates (e.g., "February" → Jan 15-Mar 15)

#### 14:36 - Initial Files Created
Created basic skeleton:
- `SKILL.md` — OpenClaw skill definition
- `index.js` — Basic CLI entry point
- `package.json` — Dependencies (links to outlook-mapi)
- `README.md` — User documentation

#### 14:45 - Module Design Discussion
Proposed architecture:
```javascript
├── query-expander.js   → Natural language → Outlook search queries
├── date-parser.js      → "February" → date range with buffer
├── searcher.js         → Orchestrates multi-phase search
├── email-analyzer.js   → LLM-powered content analysis
└── index.js            → Main entry point + chat interface
```

#### 14:50 - Core Modules Implemented
Created three foundational modules:

**1. `query-expander.js`**
- Maps common phrases to search terms ("statement of work" → "SOW", etc.)
- Builds Outlook-compatible queries with OR clauses
- Supports document type filtering

**2. `date-parser.js`**
- Parses natural language dates ("February", "last week", "a month ago")
- Adds intelligent 2-week buffer on either side
- Returns ISO date strings for Outlook query compatibility

**3. Updated `searcher.js`**
- Multi-phase orchestration (broad → filter → deep)
- Integration points for LLM analysis
- Email formatting and link extraction

#### 14:59 - LLM Integration Added
Created **`llm.js`** — critical module:
- `analyzeEmailRelevance()` — Scores emails 1-10 based on search context
- `extractRelevantSnippets()` — Pulls key excerpts from email bodies
- `callLlm()` — HTTP client for llamacpp API
- **Fallback** to keyword-based analysis if LLM unavailable

Prompt engineering:
```text
The user is searching for: "${query}"
Additional context: ${context}

Score each email from 1-10 based on relevance.
Return JSON array with score, reasoning, hasDocuments fields.
```

#### 15:08 - Dynamic Config Loading
Updated `llm.js` to **read OpenClaw config at runtime**:
- Reads `/home/openclaw/.openclaw/openclaw.json`
- Extracts `agents.defaults.model.primary`
- Supports llamacpp, Anthropic, OpenAI, openai-codex providers
- Logs which model is being used

This ensures the skill **always uses whatever provider Christopher has currently configured**.

#### 15:26 - Bridge Integration Discovery
Realized we need to use the existing `outlook-mapi` bridge client from the local repo, not create a new interface.

Investigated `/home/openclaw/repos/OpenClaw-Outlook-MAPI/` and found:
- Python Flask bridge at `scripts/outlook-bridge.py`
- Runs on port **8765**
- Endpoints: `/api/email/search`, `/api/email/recent`, `/api/email/:id`

#### 15:30 - Bridge Client Created
Created **`bridge-client.js`** — HTTP client for Outlook bridge:
- `searchEmails(query, options)` — Search with query, date filters
- `getRecentEmails(options)` — Get recent emails from folder
- `getEmailById(entryId)` — Fetch full email content
- `healthCheck()` — Verify bridge is running

#### 15:48 - First Search Test
Ran search for "Visa statement of work" and discovered:

**Found:** 48 emails about **Visa MAM project**
- Project: Media Asset Management (MAM) system using Iconik Storage Gateway
- Hardware: 2 Dell Windows workstations with ~2TB SSD each
- Start date: December 17, 2025
- Key people: Beny Korotkin (technical lead), Ernesto Carriel & Unmesh Suryawanshi (Visa contacts)

The actual SOW is referenced as "ISG requirements" in the email thread.

#### 16:06 - Clickable Links Implemented
Added **Telegram-clickable deep links** to search results:
- Format: `📧 [Open Email](https://outlook.office.com/mail/0/deeplink?subject=...&itemid=...)`
- Uses OWA deeplinks for maximum compatibility
- Works in browser AND often redirects to desktop Outlook on Windows
- Renders clickable in Telegram, Discord, Slack, etc.

**Why not `outlook:` protocol?**
The native protocol handler is unreliable since Outlook 2013+ (Microsoft removed handlers for security).

Created `createEmailLink()` function in bridge-client.js:
```javascript
function createEmailLink(entryId, subject = '') {
    const encodedSubject = encodeURIComponent(subject || '(Email)');
    return `https://outlook.office.com/mail/0/deeplink?subject=${encodedSubject}&itemid=${encodeURIComponent(entryId)}`;
}
```

#### 16:11 - Documentation Updated
Updated all documentation files to reflect current state:
- **SKILL.md** — Full technical architecture, module breakdown, API docs
- **README.md** — User-facing guide with examples and troubleshooting
- Added extensive inline comments throughout codebase

---

## Technical Architecture (Final)

```
outlook-search
├── query-expander.js   → "statement of work" → expanded search terms
├── date-parser.js      → "February" → {start, end} with buffer
├── llm.js              → LLM-powered relevance analysis
│   ├── analyzeEmailRelevance()  — Scores emails 1-10
│   ├── extractRelevantSnippets() — Pulls key excerpts
│   └── callLlm()        — HTTP client for configured model
├── bridge-client.js    → HTTP client for Outlook bridge
│   ├── searchEmails()
│   ├── getRecentEmails()
│   ├── getEmailById()
│   └── createEmailLink() — Generates OWA deeplinks
├── searcher.js         → Orchestrates all phases
│   ├── broadSearch()    — Phase 1: fast subject-level search
│   ├── analyzeEmailSummaries() — Phase 2: LLM filtering
│   ├── deepRead()       — Phase 3: full body fetch
│   └── formatEmailResult() — Formats for output
└── index.js            → Main CLI entry point
```

## Key Features Delivered

| Feature | Status |
|---------|--------|
| Natural language query expansion | ✅ |
| Date parsing with buffering | ✅ |
| LLM-powered relevance scoring | ✅ |
| Multi-phase search (broad → filter → deep) | ✅ |
| Attachment detection | ✅ |
| Link extraction from email bodies | ✅ |
| **Clickable Outlook deep links** | ✅ |
| Dynamic model config loading | ✅ |
| Bridge integration (localhost:8765) | ✅ |

## What Works Now

```bash
# Test search with clickable links
node test-clickable-links.js
```

Output includes:
- Relevance scores (1-10)
- Subject, sender, date
- **[Click here to open email](link)** — works in Telegram!

## Outstanding / Future Work

1. **Attachment downloading/preview** — Currently only detects presence
2. **Calendar event search** — Same architecture could apply
3. **Thread-aware search** — Follow conversations intelligently
4. **Conversation refinement loop** — Interactive follow-up questions
5. **More robust test suite** — Unit tests for query expansion, date parsing

## Lessons Learned

1. **OWA deeplinks are more reliable** than `outlook:` protocol for cross-platform compatibility
2. **Simple queries work better** through the bridge — complex boolean logic can fail
3. **LLM adds latency** but significantly improves relevance scoring over keywords alone
4. **Date buffering is essential** — users rarely remember exact dates
5. **Read config at runtime** — user's model preference changes, skill should adapt

---

*Last updated: March 30, 2026 16:11 PDT*
