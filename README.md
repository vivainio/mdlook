# mdlook

Sync Outlook emails to markdown files on Windows.

## Installation

```bash
uv tool install mdlook
```

Or with pip:

```bash
pip install mdlook
```

## Usage

```bash
mdlook                          # sync last 30 days to cache dir
mdlook -f Inbox -s 2026-01-01   # filter by folder and date
mdlook --dry-run                # preview without writing files
mdlook read 20                  # show 20 latest emails
mdlook read --today --unread    # today's unread emails with body
mdlook search "query"           # search and dump matching emails
mdlook list-folders             # list available Outlook folders
```

## Requirements

- Windows with **classic Outlook** installed (not the new Outlook app)
- Must be run as a **non-admin user** (Outlook COM automation does not work when elevated)
- Python 3.11+
