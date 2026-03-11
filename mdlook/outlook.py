"""Outlook COM interface for fetching emails."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Iterator

import win32com.client


@dataclass
class Attachment:
    name: str
    size: int


@dataclass
class Email:
    entry_id: str
    subject: str
    sender: str
    sender_email: str
    recipients: list[str]
    received_at: datetime
    body_html: str
    body_text: str
    folder_path: str
    attachments: list[Attachment] = field(default_factory=list)
    conversation_id: str = ""
    in_reply_to: str = ""
    unread: bool = False


def _get_folder_path(folder: object) -> str:
    """Build slash-separated folder path from Outlook folder object."""
    parts: list[str] = []
    current = folder
    while True:
        try:
            parts.append(current.Name)
            current = current.Parent
        except Exception:
            break
    parts.reverse()
    # Drop the top-level store name (first element), keep Inbox/Sent/etc.
    return "/".join(parts[1:]) if len(parts) > 1 else parts[0] if parts else ""


def _collect_folders(folder: object, pattern: re.Pattern[str] | None) -> list[object]:
    """Recursively collect Outlook folders matching optional name pattern."""
    results: list[object] = []
    try:
        name = folder.Name
    except Exception:
        return results

    path = _get_folder_path(folder)
    if pattern is None or pattern.search(path):
        results.append(folder)

    try:
        for sub in folder.Folders:
            results.extend(_collect_folders(sub, pattern))
    except Exception:
        pass

    return results


def _safe_str(value: object) -> str:
    try:
        return str(value) if value else ""
    except Exception:
        return ""


def _recipients(mail_item: object) -> list[str]:
    result: list[str] = []
    try:
        for r in mail_item.Recipients:
            try:
                name = _safe_str(r.Name)
                addr = _safe_str(r.Address)
                # Exchange DN addresses start with /o= — use name only
                if addr.startswith("/o=") or not addr:
                    if name:
                        result.append(name)
                else:
                    result.append(f"{name} <{addr}>" if name and name != addr else addr)
            except Exception:
                pass
    except Exception:
        pass
    return result


def _attachments(mail_item: object) -> list[Attachment]:
    result: list[Attachment] = []
    try:
        for att in mail_item.Attachments:
            try:
                result.append(Attachment(name=att.FileName, size=att.Size))
            except Exception:
                pass
    except Exception:
        pass
    return result


def iter_emails(
    folder_pattern: str | None = None,
    since: datetime | None = None,
    account_name: str | None = None,
    skip_ids: set[str] | None = None,
    fetch_body: bool = True,
) -> Iterator[Email]:
    """Yield Email objects from Outlook via COM.

    Args:
        folder_pattern: Regex applied to folder path (e.g. "Inbox|Sent").
        since: Only return emails received at or after this datetime.
        account_name: Match against the root store display name.
        skip_ids: Entry IDs to skip without fetching body (already synced).
        fetch_body: If False, skip HTMLBody/Body fetches (fast, for dry-run).
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    pattern = re.compile(folder_pattern, re.IGNORECASE) if folder_pattern else None

    stores = namespace.Stores
    for store_idx in range(1, stores.Count + 1):
        store = stores.Item(store_idx)
        if account_name and account_name.lower() not in _safe_str(store.DisplayName).lower():
            continue
        try:
            root = store.GetRootFolder()
        except Exception:
            continue

        for folder in _collect_folders(root, pattern):
            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)  # newest first
            except Exception:
                continue

            folder_path = _get_folder_path(folder)

            for item in items:
                try:
                    # Only process MailItem (class 43)
                    if item.Class != 43:
                        continue

                    # Cheap properties first — skip body fetch if already synced
                    entry_id = _safe_str(item.EntryID)
                    if skip_ids and entry_id in skip_ids:
                        continue

                    received: datetime = item.ReceivedTime
                    received_naive = received.replace(tzinfo=None)

                    # Items are sorted newest-first so we can break early
                    if since is not None and received_naive < since:
                        break

                    yield Email(
                        entry_id=entry_id,
                        subject=_safe_str(item.Subject) or "(no subject)",
                        sender=_safe_str(item.SenderName),
                        sender_email=_safe_str(item.SenderEmailAddress),
                        recipients=_recipients(item) if fetch_body else [],
                        received_at=received_naive,
                        body_html=_safe_str(item.HTMLBody) if fetch_body else "",
                        body_text=_safe_str(item.Body) if fetch_body else "",
                        folder_path=folder_path,
                        attachments=_attachments(item) if fetch_body else [],
                        conversation_id=_safe_str(item.ConversationID) if fetch_body else "",
                        in_reply_to=_safe_str(getattr(item, "InReplyTo", "")) if fetch_body else "",
                        unread=bool(item.UnRead),
                    )
                except Exception:
                    continue
