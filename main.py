"""detach — Email PDF attachment archiver.

Connects to an IMAP server, scans matching folders for emails from matching
senders, downloads PDF attachments (removing passwords if configured), saves
them locally, and optionally deletes the processed emails.
"""

from __future__ import annotations

import argparse
import email
import email.utils
import fnmatch
import imaplib
import logging
import os
import re
import socket
import ssl
import sys
import tomllib
from dataclasses import dataclass, field
from email.message import Message
from io import BytesIO
from pathlib import Path

import pikepdf
import socks

log = logging.getLogger("detach")


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

@dataclass
class ImapConfig:
    server: str
    username: str
    password: str
    port: int = 993
    use_ssl: bool = True
    proxy_url: str | None = None  # e.g. "socks5://proxy.host:1080"


@dataclass
class FiltersConfig:
    folder_patterns: list[str] = field(default_factory=lambda: ["INBOX"])
    sender_patterns: list[str] = field(default_factory=list)


@dataclass
class Config:
    imap: ImapConfig
    filters: FiltersConfig
    output_folder: str
    pdf_password: str = ""
    delete_after_archive: bool = False


def load_config(path: str, output_override: str | None = None) -> Config:
    """Load configuration from a TOML file.

    Output folder priority: output_override > env DETACH_OUTPUT_DIR > config value.
    """
    with open(path, "rb") as f:
        raw = tomllib.load(f)

    imap_section = raw.get("imap", {})
    imap = ImapConfig(
        server=imap_section["server"],
        username=imap_section["username"],
        password=imap_section["password"],
        port=imap_section.get("port", 993),
        use_ssl=imap_section.get("use_ssl", True),
        proxy_url=(
            imap_section.get("proxy")
            or os.environ.get("all_proxy")
            or os.environ.get("ALL_PROXY")
            or os.environ.get("http_proxy")
            or os.environ.get("HTTP_PROXY")
            or None
        ),
    )

    filters_section = raw.get("filters", {})
    filters = FiltersConfig(
        folder_patterns=filters_section.get("folder_patterns", ["INBOX"]),
        sender_patterns=filters_section.get("sender_patterns", []),
    )

    config_output = raw.get("output", {}).get("folder", ".")
    env_output = os.environ.get("DETACH_OUTPUT_DIR")
    output_folder = output_override or env_output or config_output
    output_folder = str(Path(output_folder).expanduser())

    pdf_password = raw.get("pdf", {}).get("password", "")
    delete_after = raw.get("behavior", {}).get("delete_after_archive", False)

    return Config(
        imap=imap,
        filters=filters,
        output_folder=output_folder,
        pdf_password=pdf_password,
        delete_after_archive=delete_after,
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="detach",
        description="Email PDF attachment archiver",
    )
    parser.add_argument(
        "-c", "--config",
        default=os.environ.get("DETACH_CONFIG", "config.toml"),
        help="Config file path (env: DETACH_CONFIG, default: config.toml)",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Output dir (env: DETACH_OUTPUT_DIR, overrides config)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Log actions without making changes",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Debug logging",
    )
    return parser.parse_args(argv)


# ---------------------------------------------------------------------------
# IMAP helpers
# ---------------------------------------------------------------------------

def _parse_proxy(proxy_url: str) -> tuple[int, str, int, str | None, str | None]:
    """Parse a proxy URL into (proxy_type, host, port, username, password)."""
    from urllib.parse import urlparse
    parsed = urlparse(proxy_url)
    scheme = parsed.scheme.lower()
    type_map = {"socks4": socks.SOCKS4, "socks5": socks.SOCKS5, "http": socks.HTTP}
    if scheme not in type_map:
        raise ValueError(f"Unsupported proxy scheme: {scheme!r}")
    port = parsed.port or (1080 if "socks" in scheme else 8080)
    return (type_map[scheme], parsed.hostname, port,
            parsed.username or None, parsed.password or None)


def _make_proxy_socket(
    host: str,
    port: int,
    proxy_url: str,
    timeout: float | None,
) -> socket.socket:
    """Create a TCP socket connected through the given proxy."""
    proxy_type, proxy_host, proxy_port, proxy_user, proxy_pass = _parse_proxy(proxy_url)
    sock = socks.create_connection(
        (host, port),
        proxy_type=proxy_type,
        proxy_addr=proxy_host,
        proxy_port=proxy_port,
        proxy_username=proxy_user,
        proxy_password=proxy_pass,
        timeout=timeout,
    )
    return sock


class _ProxiedIMAP4(imaplib.IMAP4):
    """IMAP4 subclass that routes the connection through a proxy."""

    def __init__(self, host: str, port: int, proxy_url: str) -> None:
        self._proxy_url = proxy_url
        super().__init__(host, port)

    def _create_socket(self, timeout):
        return _make_proxy_socket(self.host, self.port, self._proxy_url, timeout)


class _ProxiedIMAP4SSL(imaplib.IMAP4_SSL):
    """IMAP4_SSL subclass that routes the connection through a proxy."""

    def __init__(
        self,
        host: str,
        port: int,
        ssl_context: ssl.SSLContext,
        proxy_url: str,
    ) -> None:
        self._proxy_url = proxy_url
        super().__init__(host, port, ssl_context=ssl_context)

    def _create_socket(self, timeout):
        sock = _make_proxy_socket(self.host, self.port, self._proxy_url, timeout)
        return self.ssl_context.wrap_socket(sock, server_hostname=self.host)


def connect_imap(config: ImapConfig) -> imaplib.IMAP4 | imaplib.IMAP4_SSL:
    """Connect and authenticate to the IMAP server."""
    if config.proxy_url:
        log.debug("Connecting via proxy: %s", config.proxy_url)
        if config.use_ssl:
            conn = _ProxiedIMAP4SSL(
                config.server, config.port,
                ssl.create_default_context(), config.proxy_url,
            )
        else:
            conn = _ProxiedIMAP4(config.server, config.port, config.proxy_url)
    elif config.use_ssl:
        conn = imaplib.IMAP4_SSL(config.server, config.port)
    else:
        conn = imaplib.IMAP4(config.server, config.port)
    conn.login(config.username, config.password)
    log.info("Logged in to %s as %s", config.server, config.username)
    return conn


# Pattern for parsing IMAP LIST response lines.
_LIST_RE = re.compile(
    r'\((?P<flags>[^)]*)\)\s+"(?P<delimiter>[^"]+)"\s+(?P<name>.+)'
)


def list_matching_folders(
    conn: imaplib.IMAP4,
    patterns: list[str],
) -> list[str]:
    """Return folder names matching any of the given glob patterns."""
    status, data = conn.list()
    if status != "OK":
        log.error("Failed to list folders: %s", status)
        return []

    folders: list[str] = []
    for item in data:
        if item is None:
            continue
        line = item.decode() if isinstance(item, bytes) else item
        m = _LIST_RE.match(line)
        if not m:
            continue
        name = m.group("name").strip().strip('"')
        folders.append(name)

    matched = []
    for folder in folders:
        for pattern in patterns:
            if fnmatch.fnmatch(folder, pattern):
                matched.append(folder)
                break

    log.debug("All folders: %s", folders)
    log.info("Matched folders: %s", matched)
    return matched


def _match_sender(sender_addr: str, patterns: list[str]) -> bool:
    """Check if a sender email address matches any of the glob patterns."""
    sender_lower = sender_addr.lower()
    return any(fnmatch.fnmatch(sender_lower, p.lower()) for p in patterns)


def _extract_search_terms(patterns: list[str]) -> list[str]:
    """Extract literal substrings from glob patterns for IMAP FROM search.

    IMAP SEARCH FROM does substring matching, so we strip glob wildcards
    to get a useful server-side filter. For example:
      "*@bank.com"  -> "bank.com"
      "billing@*"   -> "billing"
      "*@*.com"     -> ".com"
    """
    terms: list[str] = []
    for pattern in patterns:
        # Remove leading/trailing wildcards and split on remaining ones.
        stripped = pattern.strip("*")
        if stripped:
            # Split on wildcards, strip stray '@', take the longest segment.
            parts = [p.strip("@") for p in stripped.split("*") if p.strip("@")]
            if parts:
                terms.append(max(parts, key=len))
    return terms


def fetch_matching_emails(
    conn: imaplib.IMAP4,
    folder: str,
    sender_patterns: list[str],
) -> list[tuple[str, Message]]:
    """Fetch emails in *folder* whose sender matches any pattern.

    Returns list of (uid, email.message.Message) tuples.
    Uses IMAP server-side SEARCH FROM for initial filtering, then applies
    precise fnmatch on the results.
    """
    status, _data = conn.select(f'"{folder}"', readonly=False)
    if status != "OK":
        log.warning("Could not select folder %s: %s", folder, status)
        return []

    if not sender_patterns:
        # No sender filter — search all.
        status, msg_ids = conn.uid("search", None, "ALL")
        if status != "OK" or not msg_ids[0]:
            log.debug("No messages in %s", folder)
            return []
        candidate_uids = set(msg_ids[0].split())
    else:
        # Use server-side FROM search for each extracted term.
        search_terms = _extract_search_terms(sender_patterns)
        candidate_uids: set[bytes] = set()
        for term in search_terms:
            status, msg_ids = conn.uid("search", None, "FROM", f'"{term}"')
            if status == "OK" and msg_ids[0]:
                candidate_uids.update(msg_ids[0].split())
        log.debug(
            "Server-side search returned %d candidates in %s",
            len(candidate_uids), folder,
        )

    if not candidate_uids:
        log.info("Folder %s: no matching emails", folder)
        return []

    # Fetch FROM headers for candidates and apply precise fnmatch filtering.
    matching_uids: list[bytes] = []
    for uid in candidate_uids:
        if sender_patterns:
            status, header_data = conn.uid(
                "fetch", uid, "(BODY.PEEK[HEADER.FIELDS (FROM)])"
            )
            if status != "OK":
                continue
            raw_header = header_data[0][1]
            header_msg = email.message_from_bytes(raw_header)
            _name, addr = email.utils.parseaddr(header_msg.get("From", ""))
            if not _match_sender(addr, sender_patterns):
                continue
        matching_uids.append(uid)

    log.info(
        "Folder %s: %d emails match sender filters",
        folder, len(matching_uids),
    )

    # Fetch full message for each match.
    results: list[tuple[str, Message]] = []
    for uid in matching_uids:
        status, msg_data = conn.uid("fetch", uid, "(RFC822)")
        if status != "OK":
            log.warning("Failed to fetch UID %s", uid)
            continue
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)
        results.append((uid.decode() if isinstance(uid, bytes) else uid, msg))

    return results


# ---------------------------------------------------------------------------
# PDF extraction & saving
# ---------------------------------------------------------------------------

def extract_pdf_attachments(msg: Message) -> list[tuple[str, bytes]]:
    """Extract PDF attachments from an email message.

    Returns list of (filename, pdf_bytes) tuples.
    """
    pdfs: list[tuple[str, bytes]] = []
    for part in msg.walk():
        content_type = part.get_content_type()
        disposition = str(part.get("Content-Disposition", ""))

        if content_type != "application/pdf" and not (
            "attachment" in disposition
            and (part.get_filename() or "").lower().endswith(".pdf")
        ):
            continue

        filename = part.get_filename()
        if not filename:
            filename = "unnamed.pdf"
        payload = part.get_payload(decode=True)
        if payload:
            pdfs.append((filename, payload))

    return pdfs


def remove_pdf_password(pdf_bytes: bytes, password: str) -> bytes:
    """Return decrypted PDF bytes. Raises pikepdf.PasswordError on failure."""
    with pikepdf.open(BytesIO(pdf_bytes), password=password) as pdf:
        out = BytesIO()
        pdf.save(out)
        return out.getvalue()


_UNSAFE_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def make_safe_filename(
    date_str: str,
    subject: str,
    attachment_name: str,
    folder: str,
) -> str:
    """Build a collision-safe filename: YYYY-MM-DD_subject_name.pdf.

    The folder is used as a prefix subdirectory.
    """
    # Parse the date.
    parsed = email.utils.parsedate_to_datetime(date_str) if date_str else None
    if parsed:
        date_prefix = parsed.strftime("%Y-%m-%d")
    else:
        date_prefix = "unknown-date"

    # Sanitize subject: keep first 50 chars, replace unsafe chars.
    safe_subject = _UNSAFE_CHARS.sub("_", subject or "no-subject")[:50].strip("_ ")

    # Sanitize attachment name.
    stem = Path(attachment_name).stem
    safe_name = _UNSAFE_CHARS.sub("_", stem)[:50].strip("_ ")

    filename = f"{date_prefix}_{safe_subject}_{safe_name}.pdf"

    # Sanitize folder for use as subdirectory.
    safe_folder = _UNSAFE_CHARS.sub("_", folder)

    return os.path.join(safe_folder, filename)


def save_pdf(
    pdf_bytes: bytes,
    path: str,
    password: str,
    dry_run: bool,
) -> bool:
    """Save PDF to *path*, removing password if configured.

    Returns True on success.
    """
    if password:
        try:
            pdf_bytes = remove_pdf_password(pdf_bytes, password)
            log.info("Removed password from %s", path)
        except pikepdf.PasswordError:
            log.warning("Wrong PDF password for %s, saving as-is", path)
        except Exception:
            log.exception("Error decrypting %s, saving as-is", path)

    if dry_run:
        log.info("[DRY RUN] Would save %s (%d bytes)", path, len(pdf_bytes))
        return True

    dest = Path(path)
    dest.parent.mkdir(parents=True, exist_ok=True)

    # Avoid overwriting: append a counter if file exists.
    final_path = dest
    counter = 1
    while final_path.exists():
        final_path = dest.with_stem(f"{dest.stem}_{counter}")
        counter += 1

    final_path.write_bytes(pdf_bytes)
    log.info("Saved %s (%d bytes)", final_path, len(pdf_bytes))
    return True


# ---------------------------------------------------------------------------
# Email deletion
# ---------------------------------------------------------------------------

def delete_email(
    conn: imaplib.IMAP4,
    uid: str,
    dry_run: bool,
) -> bool:
    """Flag an email for deletion by UID. Returns True on success."""
    if dry_run:
        log.info("[DRY RUN] Would delete UID %s", uid)
        return True
    typ, data = conn.uid("store", uid, "+FLAGS", "(\\Deleted)")
    if typ != "OK":
        log.warning("Failed to mark UID %s for deletion: %s %s", uid, typ, data)
        return False
    log.info("Marked UID %s for deletion", uid)
    return True


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def process_folder(
    conn: imaplib.IMAP4,
    folder: str,
    config: Config,
    dry_run: bool,
) -> tuple[int, int]:
    """Process a single folder. Returns (saved_count, email_count)."""
    emails = fetch_matching_emails(conn, folder, config.filters.sender_patterns)
    saved_count = 0
    email_count = len(emails)

    for uid, msg in emails:
        pdfs = extract_pdf_attachments(msg)
        if not pdfs:
            log.debug("UID %s: no PDF attachments", uid)
            continue

        all_saved = True
        for filename, pdf_bytes in pdfs:
            rel_path = make_safe_filename(
                msg.get("Date", ""),
                msg.get("Subject", ""),
                filename,
                folder,
            )
            full_path = os.path.join(config.output_folder, rel_path)

            if not save_pdf(pdf_bytes, full_path, config.pdf_password, dry_run):
                all_saved = False

            saved_count += 1

        if config.delete_after_archive and all_saved:
            delete_email(conn, uid, dry_run)

    if config.delete_after_archive and not dry_run:
        typ, data = conn.expunge()
        if typ != "OK":
            log.warning("EXPUNGE failed for folder %s: %s %s", folder, typ, data)
        else:
            log.debug("Expunged deleted messages in %s", folder)

    return saved_count, email_count


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)-8s %(message)s",
    )

    try:
        config = load_config(args.config, args.output)
    except FileNotFoundError:
        log.error("Config file not found: %s", args.config)
        sys.exit(1)
    except (KeyError, ValueError) as exc:
        log.error("Config error: %s", exc)
        sys.exit(1)

    log.info("Output folder: %s", config.output_folder)
    log.info("Dry run: %s", args.dry_run)

    conn = connect_imap(config.imap)

    try:
        folders = list_matching_folders(conn, config.filters.folder_patterns)
        total_saved = 0
        total_emails = 0

        for folder in folders:
            saved, emails = process_folder(conn, folder, config, args.dry_run)
            total_saved += saved
            total_emails += emails

        log.info(
            "Done: %d PDFs saved from %d emails across %d folders",
            total_saved, total_emails, len(folders),
        )
    finally:
        try:
            conn.logout()
        except Exception:
            pass


if __name__ == "__main__":
    main()
