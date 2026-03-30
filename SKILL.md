---
name: outlook-search
description: Search through Outlook inbox, calendar, and other data. Use when users want to find specific emails by keywords, sender, date range, or search for calendar events. Wraps the outlook-mapi skill to provide unified search capabilities across Outlook data.
---

# Outlook Search Skill

A powerful search interface for finding emails and calendar events in Microsoft Outlook.

## What It Does

This skill provides a unified search experience across Outlook data:

### Email Search Capabilities
- **By keyword** — search subject, body, and sender
- **By sender** — find all emails from/to specific people
- **By date range** — limit results to specific time periods
- **By status** — unread only, importance filters
- **By folder** — inbox, sent items, drafts, custom folders
- **Combined queries** — mix multiple criteria together

### Calendar Search Capabilities
- **By keyword** — search event subjects and locations
- **By attendee** — find events with specific people
- **By date range** — search within specific timeframes
- **By type** — meetings vs appointments

## Usage Examples

### Email Searches

```bash
# Find emails by keyword
outlook-search --type email --query "project alpha"

# Find unread emails from a sender
outlook-search --type email --query "from:john.doe" --unread-only

# Search with date filter (last 7 days)
outlook-search --type email --query "budget" --days 7

# Find high importance emails
outlook-search --type email --importance high --limit 20
```

### Calendar Searches

```bash
# Search calendar events by keyword
outlook-search --type calendar --query "meeting" --days 30

# Find events with specific attendee (requires full implementation)
outlook-search --type calendar --query "john.doe"
```

### Combined/Advanced

```bash
# Search both emails and calendar
outlook-search --type all --query "quarterly review" --days 60

# Limit results
outlook-search --type email --query "invoice" --limit 5
```

## Architecture

This skill imports from **outlook-mapi** for data access:

```
outlook-search (this skill)
    ↓ imports from ↓
outlook-mapi (general-purpose Outlook API client)
    ↓ fetches via HTTP ↓
outlook-bridge.py (Python Flask → MAPI COM)
    ↓ uses MAPI ↓
Windows Outlook Desktop App
```

## CLI Interface

### Command Format

```bash
outlook-search [options]
```

### Options

| Option | Default | Description |
|--------|---------|-------------|
| `--type` | `email` | What to search: `email`, `calendar`, or `all` |
| `--query` | `""` | Search query string |
| `--folder` | `inbox` | Email folder (only for --type email) |
| `--unread-only` | `false` | Show only unread emails |
| `--importance` | none | Filter by: `high`, `normal`, `low` |
| `--days` | none | Limit to last N days |
| `--limit` | 10 | Max results to return |
| `--format` | `pretty` | Output format: `pretty`, `json`, `minimal` |

## Implementation Notes

- Requires Windows with Outlook Desktop app installed
- Depends on outlook-mapi skill being available
- Bridge server runs at http://localhost:5000 by default
