"""CLI entry point for mdlook."""

from __future__ import annotations

import sys

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
if sys.stderr.encoding and sys.stderr.encoding.lower() != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]

from datetime import datetime, timedelta
from pathlib import Path

import click
from platformdirs import user_data_dir

from mdlook.sync import run_sync

DEFAULT_OUTPUT_DIR = str(Path(user_data_dir("mdlook", appauthor=False)) / "mails")


def _parse_date(ctx: object, param: object, value: str | None) -> datetime | None:
    if value is None:
        return None
    for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    raise click.BadParameter(f"Cannot parse date '{value}'. Use YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS.")


@click.group(invoke_without_command=True)
@click.option("--output", "-o", default=DEFAULT_OUTPUT_DIR, type=click.Path(), help="Directory to write .md files into.")
@click.option("--folder", "-f", default=None, help="Regex filter on folder path (e.g. 'Inbox|Sent').")
@click.option("--since", "-s", default=None, callback=_parse_date, is_eager=False, expose_value=True,
              help="Only sync emails received on/after this date (YYYY-MM-DD).")
@click.option("--account", "-a", default=None, help="Filter by Outlook account/store name.")
@click.option("--flat", is_flag=True, default=False, help="Write all files flat (no subdirs).")
@click.option("--dry-run", is_flag=True, default=False, help="Discover emails but do not write files.")
@click.option("--state-file", default=None, type=click.Path(), help="Path to state JSON file.")
@click.pass_context
def main(
    ctx: click.Context,
    output: str,
    folder: str | None,
    since: datetime | None,
    account: str | None,
    flat: bool,
    dry_run: bool,
    state_file: str | None,
) -> None:
    """Sync Outlook emails to markdown files (default output: platformdirs cache)."""
    if ctx.invoked_subcommand is not None:
        return

    out = Path(output)
    sf = Path(state_file) if state_file else None
    if since is None:
        since = datetime.now() - timedelta(days=30)

    click.echo(f"Syncing to: {out.resolve()}")
    if folder:
        click.echo(f"  Folder filter: {folder}")
    if since:
        click.echo(f"  Since: {since.date()}")
    if account:
        click.echo(f"  Account: {account}")
    if dry_run:
        click.echo("  [DRY RUN — no files will be written]")

    def progress(email: object, status: str) -> None:
        marker = "+" if status == "written" else "-" if status == "skipped" else "~" if status == "dry-run" else "!"
        line = f"  [{marker}] {email.received_at.date()} {email.subject[:60]}"  # type: ignore[union-attr]
        click.echo(line.encode("utf-8", errors="replace").decode("utf-8", errors="replace"))

    result = run_sync(
        output_dir=out,
        folder_pattern=folder,
        since=since,
        account_name=account,
        flat=flat,
        dry_run=dry_run,
        state_file=sf,
        progress_cb=progress,
    )

    click.echo("")
    click.echo(f"Done. Written: {result.written}  Skipped: {result.skipped}  Errors: {result.errors}")


def _parse_frontmatter(text: str) -> tuple[dict[str, str], str]:
    """Return (meta dict, body) from a markdown file."""
    lines = text.splitlines()
    meta: dict[str, str] = {}
    body_start = 0
    if lines and lines[0] == "---":
        for i, line in enumerate(lines[1:], 1):
            if line == "---":
                body_start = i + 1
                break
            if ": " in line:
                k, _, v = line.partition(": ")
                meta[k.strip()] = v.strip().strip('"')
    return meta, "\n".join(lines[body_start:]).strip()


@main.command("read")
@click.argument("count", default=10, type=int)
@click.option("--output", "-o", default=DEFAULT_OUTPUT_DIR, type=click.Path(), help="Mail directory.")
@click.option("--today", is_flag=True, default=False, help="Only show emails from today.")
@click.option("--unread", is_flag=True, default=False, help="Only show unread emails (requires re-sync).")
@click.option("--body", is_flag=True, default=False, help="Print full body of each email.")
def read_mails(count: int, output: str, today: bool, unread: bool, body: bool) -> None:
    """Show the COUNT latest emails (default: 10)."""
    from datetime import date as date_type
    mail_dir = Path(output)
    today_str = date_type.today().isoformat()
    show_body = body or unread or today

    files = sorted(mail_dir.rglob("*.md"), key=lambda p: p.name, reverse=True)
    results = []
    for path in files:
        text = path.read_text(encoding="utf-8", errors="replace")
        meta, body_text = _parse_frontmatter(text)
        if today and not meta.get("date", "").startswith(today_str):
            continue
        if unread and meta.get("unread") != "true":
            continue
        results.append((meta, body_text, path))
        if not (today or unread) and len(results) >= count:
            break

    if not results:
        click.echo("No emails found.")
        return

    sep = "=" * 72
    for meta, body_text, path in results:
        subject = meta.get("subject", "?")
        sender = meta.get("from", "?")
        date = meta.get("date", "")[:10]
        if show_body:
            click.echo(sep)
            click.echo(f"{date}  {sender[:50]}")
            click.echo(f"Subject: {subject}")
            click.echo("")
            click.echo(body_text)
        else:
            click.echo(f"{date}  {sender[:30]:<30}  {subject[:60]}")
    if show_body:
        click.echo(sep)


@main.command("search")
@click.argument("query")
@click.option("--output", "-o", default=DEFAULT_OUTPUT_DIR, type=click.Path(), help="Mail directory.")
def search_mails(query: str, output: str) -> None:
    """Search emails by text and print matching bodies."""
    import re
    mail_dir = Path(output)
    pattern = re.compile(query, re.IGNORECASE)
    matches = sorted(
        (p for p in mail_dir.rglob("*.md") if pattern.search(p.read_text(encoding="utf-8", errors="replace"))),
        key=lambda p: p.name,
    )
    if not matches:
        click.echo("No matches.")
        return
    click.echo(f"{len(matches)} match(es) for '{query}'\n")
    sep = "=" * 72
    for path in matches:
        text = path.read_text(encoding="utf-8", errors="replace")
        click.echo(sep)
        click.echo(str(path.resolve()))
        click.echo(text)
    click.echo(sep)


@main.command("list-folders")
@click.option("--account", "-a", default=None, help="Filter by Outlook account/store name.")
def list_folders(account: str | None) -> None:
    """List all Outlook folder paths available for syncing."""
    import win32com.client
    from mdlook.outlook import _collect_folders, _get_folder_path

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    stores = ns.Stores
    for i in range(1, stores.Count + 1):
        store = stores.Item(i)
        name = store.DisplayName
        if account and account.lower() not in name.lower():
            continue
        click.echo(f"\n[{name}]")
        try:
            root = store.GetRootFolder()
            for folder in _collect_folders(root, None):
                path = _get_folder_path(folder)
                click.echo(f"  {path}")
        except Exception as exc:
            click.echo(f"  (error: {exc})")
