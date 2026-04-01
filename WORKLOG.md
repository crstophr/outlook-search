# Work Log - Outlook Search Skill

## Overview

**Project:** Outlook email search skill with natural-language query expansion, LLM relevance scoring, and clickable Outlook Web deeplinks  
**Started:** March 30, 2026  
**Repository:** https://github.com/crstophr/outlook-search

## Current repo location

Canonical local repo path now:
- `/home/openclaw/repos/outlook-search`

For compatibility with the OpenClaw skill path, the old workspace location may be a symlink:
- `/home/openclaw/.openclaw/workspace/skills/outlook-search`

## What exists in code today

### Core modules
- `index.js` — interactive CLI session flow and exports
- `searcher.js` — main orchestration
- `query-expander.js` — query expansion
- `date-parser.js` — natural-language date parsing
- `llm.js` — model config loading, LLM scoring, snippet extraction, keyword fallback
- `bridge-client.js` — HTTP client and Outlook Web deeplink generator

### Bridge routes used
- `GET /health`
- `POST /api/email/search`
- `GET /api/email/<id>`

### Real scope today
Despite some older wording, this repo is currently focused on **email search**. It does not contain a real calendar-search implementation.

## Development timeline summary

### March 30, 2026
- Project initiated to search Outlook mail for project documentation and related context
- Designed a multi-phase flow: broad scan → LLM filtering → deep read
- Implemented:
  - `query-expander.js`
  - `date-parser.js`
  - `searcher.js`
  - `llm.js`
  - `bridge-client.js`
  - `index.js`
- Added Outlook Web deeplink generation for clickable results
- Wired the repo to the local Outlook bridge on port 8765

## Code realities worth remembering

### 1. Local dependency path changed after repo move
After moving this repo into `/home/openclaw/repos`, the local dependency on `OpenClaw-Outlook-MAPI` must resolve relative to that new location.

Expected local dependency path in `package.json`:
- `file:../OpenClaw-Outlook-MAPI`

### 2. Field-shape compatibility was added in `llm.js`
`llm.js` now normalizes both older camelCase properties and the bridge’s current snake_case response fields, including:
- `sender` / `senderName`
- `sender_email` / `senderEmail`
- `received_datetime` / `receivedDateTime`
- `has_attachments` / `hasAttachments`

That keeps prompt construction and fallback scoring aligned with the current Outlook bridge output.

### 3. Date filtering is partly client-side
`bridge-client.js` sends `start_date` / `end_date`, but current effective behavior depends mostly on:
- `days` sent to the bridge
- client-side end-date filtering in `searcher.js`

### 4. Snippet extraction is implemented but not integrated
`extractRelevantSnippets()` exists in `llm.js`, but `searcher.js` does not currently use it in the main result-building path.

### 5. Test scripts are ad hoc
There are multiple one-off scripts like:
- `test-visa.js`
- `test-clickable-links.js`
- `test-simple.js`
- `search-visa-work.js`
- `direct-visa-search.js`

These are useful probes, not a formal automated test suite.

## Suggested reading order for future work

1. `searcher.js`
2. `bridge-client.js`
3. `llm.js`
4. `README.md`
5. `SKILL.md`

## Good next cleanup targets

1. Align field names with the Outlook bridge response schema.
2. Decide whether calendar search belongs here or should be removed from leftover metadata/docs entirely.
3. Replace ad hoc test scripts with a real test harness.
4. Decide whether snippet extraction should be integrated into displayed results.
5. Tighten docs anytime the bridge contract changes.

---

*Updated after moving repo into `/home/openclaw/repos` and re-auditing docs against code.*