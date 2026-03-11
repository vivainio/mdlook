# mdlook skill

Use mdlook to read, search and sync Outlook emails from the command line.

## Commands

```bash
mdlook                          # sync inbox (last hour or 30 days on first run)
mdlook --external               # also include external senders
mdlook -f "Inbox|Sent"          # sync specific folders (regex)
mdlook -s 2026-01-01            # sync from a specific date
mdlook --dry-run                # preview without writing files

mdlook read 20                  # show 20 latest emails
mdlook read --today             # today's emails with body
mdlook read --unread            # unread emails with body
mdlook read --body              # all emails with body

mdlook search "query"           # search email content, dump matching bodies

mdlook reset                    # delete state + all .md files, force full re-sync
mdlook list-folders             # list available Outlook folders
```

## Output location

Emails are stored as markdown files in:
```
C:\Users\<user>\AppData\Local\mdlook\mails\
```

Folder structure mirrors Outlook: `mails/<FolderName>/<YYYY-MM>/<date>_<subject>_<id>.md`

## Frontmatter fields

Each `.md` file has YAML frontmatter:

```yaml
---
subject: Meeting tomorrow
from: Jane Doe <jane@example.com>
to: Ville Vainio <ville@basware.com>
date: 2026-03-11 09:30:00
unread: true       # only present if unread at sync time
external: true     # only present if sender is outside Exchange org
folder: Inbox
conversation_id: "..."
attachments: report.pdf, photo.jpg
---
```

## Notes

- Requires classic Outlook (not the new Outlook app)
- Must run as non-admin
- External senders (outside the Exchange org) are skipped by default
- State is tracked in `.mdlook_state.json` in the output dir
- On subsequent syncs, only emails since last sync (minus 1h buffer) are checked
- Iteration stops early when an already-seen email is encountered (newest-first sort)
