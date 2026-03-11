"""Convert Email objects to markdown strings."""

from __future__ import annotations

import re

import html2text

from mdlook.outlook import Email

_h2t = html2text.HTML2Text()
_h2t.ignore_links = False
_h2t.ignore_images = True
_h2t.ignore_tables = True
_h2t.body_width = 0  # no line-wrapping


def _strip_reply_html(html: str) -> str:
    """Remove quoted reply/forward sections from Outlook HTML.

    Outlook wraps the original message in one of these patterns:
    - <div id="divRplyFwdMsg"> (modern Outlook)
    - <div id="x_divRplyFwdMsg"> (Exchange/web variant)
    - <hr> followed by From/Sent/To/Subject block (classic Outlook)
    """
    # Remove reply/forward div and everything after it
    html = re.sub(
        r'<div[^>]+id=["\']x?divRplyFwdMsg["\'][^>]*>.*',
        "",
        html,
        flags=re.IGNORECASE | re.DOTALL,
    )
    # Remove <blockquote> (quoted text)
    html = re.sub(r"<blockquote[^>]*>.*?</blockquote>", "", html, flags=re.IGNORECASE | re.DOTALL)
    return html


def _strip_reply_plain(text: str) -> str:
    """Remove quoted reply from plain text emails."""
    # Common separators used by Outlook in plain text
    separators = [
        r"^-{3,}\s*Original Message\s*-{3,}",
        r"^From:.*\nSent:.*\nTo:",
        r"^_{3,}",
    ]
    for sep in separators:
        m = re.search(sep, text, flags=re.IGNORECASE | re.MULTILINE)
        if m:
            text = text[: m.start()].rstrip()
            break
    return text


_SAFELINK_RE = re.compile(
    r'https://eur\w+\.safelinks\.protection\.outlook\.com/\?url=([^&">\s]+)[^">\s)]*',
    re.IGNORECASE,
)
_CONFIDENTIALITY_RE = re.compile(
    r'(Confidentiality Notice|NOTICE OF CONFIDENTIALITY).*',
    re.IGNORECASE | re.DOTALL,
)


def _clean_markdown(text: str) -> str:
    # Unwrap safelinks to actual URLs
    text = _SAFELINK_RE.sub(lambda m: __import__('urllib.parse', fromlist=['unquote']).unquote(m.group(1)), text)
    # Strip mailto links: [foo](mailto:...) -> foo
    text = re.sub(r'\[([^\]]+)\]\(mailto:[^)]+\)', r'\1', text)
    # Strip empty links: [](url) or [ ](url)
    text = re.sub(r'\[\s*\]\([^)]*\)', '', text)
    # Strip tel: links: [+123](tel:...) -> +123
    text = re.sub(r'\[([^\]]+)\]\(tel:[^)]+\)', r'\1', text)
    # Strip "Classified as Private" lines
    text = re.sub(r'\n?\s*Classified as Private\s*\n?', '\n', text, flags=re.IGNORECASE)
    # Strip confidentiality notice blocks
    text = _CONFIDENTIALITY_RE.sub('', text)
    # Strip zero-width / invisible characters used as email spacers
    text = re.sub(r'[\u200c\u034f\u00ad\ufeff]+', '', text)
    text = re.sub(r'(\s*[\u200c\u034f]\s*){2,}', '', text)
    # Strip [Image] and [image NNN] artifacts
    text = re.sub(r'\[image\d*\]', '', text, flags=re.IGNORECASE)
    # Truncate long URLs in markdown links: [text](very-long-url) -> [text](url…)
    def _shorten_url(m: re.Match) -> str:
        label, url = m.group(1), m.group(2)
        short = url[:80] + "…" if len(url) > 80 else url
        return f"[{label}]({short})"
    text = re.sub(r'\[([^\]]*)\]\((https?://[^)]+)\)', _shorten_url, text)
    # Truncate bare long URLs
    def _shorten_bare(m: re.Match) -> str:
        url = m.group(0)
        return url[:80] + "…" if len(url) > 80 else url
    text = re.sub(r'(?<!\()https?://\S+', _shorten_bare, text)
    # Collapse 3+ blank lines to 2
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def _html_to_md(html: str) -> str:
    if not html.strip():
        return ""
    html = _strip_reply_html(html)
    return _clean_markdown(_h2t.handle(html).strip())


def _plain_to_md(text: str) -> str:
    return _strip_reply_plain(text.strip())


def email_to_markdown(email: Email) -> str:
    """Return a markdown string representing the email."""
    lines: list[str] = []

    # YAML-style front matter
    lines.append("---")
    lines.append(f"subject: {_yaml_str(email.subject)}")
    sender_email = email.sender_email if not email.sender_email.lower().startswith("/o=") else ""
    from_val = f"{email.sender} <{sender_email}>" if sender_email else email.sender
    lines.append(f"from: {_yaml_str(from_val)}")
    if email.recipients:
        lines.append(f"to: {_yaml_str(', '.join(email.recipients))}")
    lines.append(f"date: {email.received_at.strftime('%Y-%m-%d %H:%M:%S')}")
    if email.unread:
        lines.append("unread: true")
    if email.external:
        lines.append("external: true")
    lines.append(f"folder: {_yaml_str(email.folder_path)}")
    if email.conversation_id:
        lines.append(f"conversation_id: {_yaml_str(email.conversation_id)}")
    if email.attachments:
        att_names = ", ".join(a.name for a in email.attachments)
        lines.append(f"attachments: {_yaml_str(att_names)}")
    lines.append("---")
    lines.append("")

    # Title
    lines.append(f"# {email.subject}")
    lines.append("")

    # Body — prefer HTML conversion, fall back to plain text
    if email.body_html.strip():
        body = _html_to_md(email.body_html)
    else:
        body = _plain_to_md(email.body_text)

    lines.append(body)

    return "\n".join(lines)


def _yaml_str(value: str) -> str:
    """Wrap value in quotes if it contains special YAML characters."""
    if any(c in value for c in (':', '#', '[', ']', '{', '}', '>', '|', '&', '*', '!')):
        escaped = value.replace('"', '\\"')
        return f'"{escaped}"'
    return value


def safe_filename(subject: str, received_at: object, entry_id: str) -> str:
    """Build a filesystem-safe filename for the email."""
    date_str = received_at.strftime("%Y%m%d_%H%M%S")
    slug = re.sub(r"[^\w\s-]", "", subject).strip()
    slug = re.sub(r"[\s_]+", "_", slug)[:60]
    short_id = entry_id[-8:] if len(entry_id) >= 8 else entry_id
    return f"{date_str}_{slug}_{short_id}.md"
