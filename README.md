# Outlook Search

A CLI tool and OpenClaw skill for searching emails and calendar events in Microsoft Outlook.

## Installation

```bash
cd skills/outlook-search
npm install
```

## Usage

### Command Line

```bash
# Search emails by keyword
node index.js --query "project alpha"

# Search unread emails only
node index.js --unread-only

# Search calendar events from last 30 days
node index.js --type calendar --query "meeting" --days 30

# Search both emails and calendar
node index.js --type all --query "quarterly review"
```

### Options

| Option | Description |
|--------|-------------|
| `--type [email\|calendar\|all]` | What to search (default: email) |
| `--query, -q` | Search query string |
| `--folder, -f` | Email folder to search (default: inbox) |
| `--unread-only, -u` | Show only unread emails |
| `--importance [high\|normal\|low]` | Filter by importance |
| `--days [n]` | Limit to last N days (for calendar) |
| `--limit, -l [n]` | Max results (default: 10) |
| `--format [pretty\|json]` | Output format |

## Example Outputs

```bash
$ node index.js --query "budget" --unread-only

Outlook Search Tool
==================

--- EMAIL SEARCH ---

Searching emails in inbox...
Query: budget AND isRead:false

Found 3 email(s):

[unread] finance@example.com: Q4 Budget Review (Mar 29, 2026 10:30 AM)
[unread] manager@example.com: Budget Approval Needed - Project Alpha (Mar 28, 2026 2:15 PM)
[unread] [high] accounting@example.com: Budget Variance Report (Mar 27, 2026 4:45 PM)
```

## Requirements

- Windows with Outlook Desktop app installed
- outlook-mapi skill configured and accessible
- Bridge server running (typically http://localhost:5000)
