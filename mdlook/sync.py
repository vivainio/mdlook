"""Orchestrate the sync: fetch emails, convert, write to disk."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

from mdlook.convert import email_to_markdown, safe_filename
from mdlook.outlook import Email, iter_emails
from mdlook.state import SyncState


@dataclass
class SyncResult:
    written: int = 0
    skipped: int = 0
    errors: int = 0


def _output_path(email: Email, output_dir: Path, flat: bool) -> Path:
    filename = safe_filename(email.subject, email.received_at, email.entry_id)
    if flat:
        return output_dir / filename
    # Mirror folder structure with month subdir: e.g. mails/Inbox/2025-12/file.md
    month = email.received_at.strftime("%Y-%m")
    sub = Path(*email.folder_path.split("/")) if email.folder_path else Path(".")
    return output_dir / sub / month / filename


def run_sync(
    output_dir: Path,
    folder_pattern: str | None = None,
    since: datetime | None = None,
    account_name: str | None = None,
    include_external: bool = False,
    flat: bool = False,
    dry_run: bool = False,
    state_file: Path | None = None,
    progress_cb: object = None,
) -> SyncResult:
    """Sync Outlook emails to markdown files.

    Args:
        output_dir: Root directory to write .md files into.
        folder_pattern: Regex to filter Outlook folder paths.
        since: Only sync emails received on or after this datetime.
        account_name: Filter by Outlook account/store display name.
        flat: Write all files into output_dir without subdirectories.
        dry_run: Discover emails but do not write any files.
        state_file: Path to JSON state file for deduplication.
        progress_cb: Optional callable(email, status) for progress reporting.
    """
    if state_file is None:
        state_file = output_dir / ".mdlook_state.json"

    state = SyncState(state_file)
    result = SyncResult()

    # Use last_synced_at (minus buffer) as since when not explicitly provided
    if since is None:
        if state.last_synced_at is not None:
            since = state.last_synced_at - timedelta(hours=1)
        else:
            since = datetime.now() - timedelta(days=30)

    if not dry_run:
        output_dir.mkdir(parents=True, exist_ok=True)

    for email in iter_emails(
        folder_pattern=folder_pattern,
        since=since,
        account_name=account_name,
        skip_ids=state.seen,
        fetch_body=not dry_run,
    ):
        try:
            if not include_external and email.external:
                continue

            dest = _output_path(email, output_dir, flat)

            if not dry_run:
                dest.parent.mkdir(parents=True, exist_ok=True)
                md = email_to_markdown(email)
                dest.write_text(md, encoding="utf-8")
                state.mark_synced(email.entry_id)

            result.written += 1
            if progress_cb:
                progress_cb(email, "written" if not dry_run else "dry-run")

        except Exception as exc:
            result.errors += 1
            if progress_cb:
                progress_cb(email, f"error: {exc}")

    if not dry_run:
        state.save()

    return result
