# Outlook Search

Natural-language search over Outlook email using the local Outlook bridge plus LLM-based relevance scoring.

This repo currently focuses on **email search**, not full Outlook search across calendar/tasks/other objects.

## Current location

Canonical repo location:
- `/home/openclaw/repos/outlook-search`

For compatibility, the old skill path can still point here via symlink:
- `/home/openclaw/.openclaw/workspace/skills/outlook-search`

## What the code does today

Core flow in `searcher.js`:
1. verify the Outlook bridge is reachable
2. parse an optional natural-language date hint
3. expand the query into several variants
4. run broad email searches through the bridge
5. deduplicate by email id
6. score/filter results with the configured LLM
7. deep-read the top candidates
8. format results with Outlook Web deep links

## Implemented modules

| File | Purpose |
|---|---|
| `index.js` | Interactive CLI wrapper and exports for programmatic use |
| `searcher.js` | Main orchestration: broad search → LLM filter → deep read |
| `query-expander.js` | Query synonym/variant expansion |
| `date-parser.js` | Natural-language date hints → date ranges |
| `llm.js` | Relevance scoring + snippet extraction via configured model |
| `bridge-client.js` | HTTP client for the Outlook bridge |

## Bridge/API dependency

This repo depends on the Outlook bridge repo:
- `/home/openclaw/repos/OpenClaw-Outlook-MAPI`

Bridge endpoints used by code:
- `GET /health`
- `POST /api/email/search`
- `GET /api/email/<id>`

The code does **not** currently perform calendar search.

## Running it

### Install dependencies
From the repo root:

```bash
cd /home/openclaw/repos/outlook-search
npm install
```

`package.json` links the local Outlook bridge package via:
- `file:../OpenClaw-Outlook-MAPI`

### Run interactive CLI
```bash
node index.js --cli
```

Note: because `index.js` treats any extra CLI args as interactive mode, `node index.js anything` will also start the prompt flow.

## Programmatic use

```javascript
const { searchOutlook } = require('./searcher');

const results = await searchOutlook(
  {
    query: 'Visa statement of work',
    context: 'looking for signed document or link to it',
    dateHint: 'February'
  },
  {
    topN: 10,
    deepReadLimit: 5
  }
);
```

Exports from `index.js`:
- `OutlookSearchSession`
- `searchOutlook`
- `formatEmailForDisplay`

## Search behavior details

### Query expansion
`query-expander.js`:
- expands some common phrases using a synonym map
- generates three query variants
- optionally adds document-related terms like `attachment`, `attached`, `document`, `pdf`

### Date handling
`date-parser.js` supports:
- month names like `February`
- relative phrases like `last week`, `last month`, `3 days ago`, `2 months ago`, `recent`
- default fallback: last 90 days

Implementation note:
- `last week` / `this week` currently subtracts `getDay()` from today even though the comment says Monday; behavior is code-defined, not calendar-perfect.

### LLM scoring
`llm.js`:
- loads OpenClaw model config from `~/.openclaw/openclaw.json` when available
- calls a chat-completions-style endpoint
- scores emails 1–10
- keeps emails with score >= 4
- falls back to keyword scoring if the LLM call fails

### Deep read
`searcher.js` fetches full email data for the top candidates and extracts up to 10 HTTP(S) links from body/html.

## Output

`searchOutlook()` returns an object like:

```javascript
{
  results: [/* detailed email objects */],
  summary: 'Found N relevant emails from M matches',
  metadata: {
    queriesRun: 3,
    initialMatches: 48,
    afterFiltering: 8
  }
}
```

`formatEmailResult()` builds a smaller display-oriented object including:
- `subject`
- `sender`
- `date`
- `attachments`
- `emailLink`
- `entryId`
- `relevanceScore`
- `reasoning`

## Deep links

Email links are generated as Outlook Web deeplinks:

```text
https://outlook.office.com/mail/0/deeplink?subject=...&itemid=...
```

This is what the current code uses; it does not use the native `outlook:` protocol.

## Known code realities

These are worth knowing before another LLM edits this repo:

- The repo is primarily an **email** search tool despite some older references to calendar.
- `bridge-client.js` sends `start_date` and `end_date` fields, but the current Outlook bridge mainly relies on `days` and client-side end-date filtering.
- `llm.js` now normalizes both bridge-style snake_case fields and older camelCase variants before prompting/scoring.
- `searcher.js` imports `extractRelevantSnippets` but does not currently use it in the main search flow.
- There are several ad hoc test/search scripts in the repo; they are not a formal test suite.

## Useful files for handoff

Read in this order:
1. `searcher.js`
2. `bridge-client.js`
3. `llm.js`
4. `query-expander.js`
5. `date-parser.js`
6. `WORKLOG.md`

## Future cleanup ideas exposed by the current code

- unify field naming assumptions between this repo and the Outlook bridge
- decide whether calendar search is in scope or remove the leftover mentions entirely
- add a proper test harness instead of one-off search scripts
- decide whether this should remain a standalone repo or be folded into a broader Outlook skill set