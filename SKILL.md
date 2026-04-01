---
name: outlook-search
description: Intelligent natural language search through Outlook email using the local Outlook bridge, query expansion, date parsing, LLM-based relevance scoring, and clickable Outlook Web deep links. Use when the user wants to find emails, attachments, message threads, or likely document references in Outlook from vague natural-language prompts.
---

# Outlook Search Skill

This skill repo currently implements **email search** over Outlook mail via the local bridge.

## What the code currently does

- natural-language-ish email search through query expansion
- optional date-hint parsing with buffered date ranges
- bridge health check before searching
- broad search over Outlook email
- LLM-based relevance scoring and filtering
- deep read of top candidate messages
- extraction of HTTP(S) links from message body/html
- generation of clickable Outlook Web deeplinks for results

## Main files

- `index.js` — interactive CLI + exported helpers
- `searcher.js` — main orchestration flow
- `query-expander.js` — synonym/query expansion
- `date-parser.js` — natural-language date parsing
- `llm.js` — model-config loading + relevance analysis
- `bridge-client.js` — HTTP client for bridge routes
- `WORKLOG.md` — project history and lessons

## Bridge routes actually used

- `GET /health`
- `POST /api/email/search`
- `GET /api/email/<id>`

The repo currently does **not** implement calendar search, even though some older metadata/docs implied broader Outlook coverage.

## Run locally

```bash
cd /home/openclaw/repos/outlook-search
npm install
node index.js --cli
```

## Programmatic use

```javascript
const { searchOutlook } = require('./searcher');

const results = await searchOutlook({
  query: 'Visa statement of work',
  context: 'looking for signed document or link to it',
  dateHint: 'February'
});
```

## Important handoff notes

- The repo now lives at `/home/openclaw/repos/outlook-search`.
- If OpenClaw still references the workspace skill path, that can be satisfied by a symlink.
- `package.json` should point to the Outlook bridge via `file:../OpenClaw-Outlook-MAPI` from this repo location.
- `llm.js` now normalizes both snake_case bridge fields and older camelCase variants before scoring/prompts.
- `searcher.js` imports `extractRelevantSnippets` but does not use it in the main path.

## Best files to inspect before editing

1. `searcher.js`
2. `bridge-client.js`
3. `llm.js`
4. `README.md`
5. `WORKLOG.md`