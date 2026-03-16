"""Tests for detach — Email PDF attachment archiver."""

from __future__ import annotations

import email
import os
import textwrap
from email.message import Message
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO
from pathlib import Path
from unittest.mock import MagicMock, patch

import pikepdf
import pytest

from main import (
    Config,
    FiltersConfig,
    ImapConfig,
    _extract_search_terms,
    _match_sender,
    extract_pdf_attachments,
    fetch_matching_emails,
    list_matching_folders,
    load_config,
    make_safe_filename,
    parse_args,
    process_folder,
    remove_pdf_password,
    save_pdf,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def sample_config_file(tmp_path: Path) -> Path:
    config = tmp_path / "config.toml"
    config.write_text(textwrap.dedent("""\
        [imap]
        server = "imap.test.com"
        username = "user@test.com"
        password = "pass123"
        port = 993
        use_ssl = true

        [filters]
        folder_patterns = ["INBOX", "Bills/*"]
        sender_patterns = ["*@bank.com"]

        [output]
        folder = "/tmp/detach-test"

        [pdf]
        password = "secret"

        [behavior]
        delete_after_archive = true
    """))
    return config


@pytest.fixture()
def minimal_config_file(tmp_path: Path) -> Path:
    config = tmp_path / "config.toml"
    config.write_text(textwrap.dedent("""\
        [imap]
        server = "imap.test.com"
        username = "user@test.com"
        password = "pass123"
    """))
    return config


def _make_pdf_bytes() -> bytes:
    """Create a minimal valid PDF in memory."""
    buf = BytesIO()
    with pikepdf.new() as pdf:
        pdf.save(buf)
    return buf.getvalue()


def _make_encrypted_pdf_bytes(password: str) -> bytes:
    """Create a password-protected PDF in memory."""
    buf = BytesIO()
    with pikepdf.new() as pdf:
        pdf.save(buf, encryption=pikepdf.Encryption(owner=password, user=password))
    return buf.getvalue()


def _make_email_with_pdf(
    sender: str = "billing@bank.com",
    subject: str = "Your Statement",
    date: str = "Tue, 15 Mar 2026 10:00:00 +0000",
    pdf_name: str = "statement.pdf",
    pdf_bytes: bytes | None = None,
) -> Message:
    """Build a multipart email with a PDF attachment."""
    msg = MIMEMultipart()
    msg["From"] = f"Billing <{sender}>"
    msg["Subject"] = subject
    msg["Date"] = date

    msg.attach(MIMEText("Please find your statement attached."))

    if pdf_bytes is None:
        pdf_bytes = _make_pdf_bytes()
    att = MIMEApplication(pdf_bytes, _subtype="pdf")
    att.add_header("Content-Disposition", "attachment", filename=pdf_name)
    msg.attach(att)

    return msg


# ---------------------------------------------------------------------------
# Config loading
# ---------------------------------------------------------------------------

class TestLoadConfig:
    def test_full_config(self, sample_config_file: Path) -> None:
        config = load_config(str(sample_config_file))
        assert config.imap.server == "imap.test.com"
        assert config.imap.username == "user@test.com"
        assert config.imap.password == "pass123"
        assert config.imap.port == 993
        assert config.imap.use_ssl is True
        assert config.filters.folder_patterns == ["INBOX", "Bills/*"]
        assert config.filters.sender_patterns == ["*@bank.com"]
        assert config.output_folder == "/tmp/detach-test"
        assert config.pdf_password == "secret"
        assert config.delete_after_archive is True

    def test_minimal_config_defaults(self, minimal_config_file: Path) -> None:
        config = load_config(str(minimal_config_file))
        assert config.imap.port == 993
        assert config.imap.use_ssl is True
        assert config.filters.folder_patterns == ["INBOX"]
        assert config.filters.sender_patterns == []
        assert config.pdf_password == ""
        assert config.delete_after_archive is False

    def test_output_override(self, sample_config_file: Path) -> None:
        config = load_config(str(sample_config_file), output_override="/override")
        assert config.output_folder == "/override"

    def test_env_output_override(
        self, sample_config_file: Path, monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        monkeypatch.setenv("DETACH_OUTPUT_DIR", "/from-env")
        config = load_config(str(sample_config_file))
        assert config.output_folder == "/from-env"

    def test_cli_overrides_env(
        self, sample_config_file: Path, monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        monkeypatch.setenv("DETACH_OUTPUT_DIR", "/from-env")
        config = load_config(str(sample_config_file), output_override="/from-cli")
        assert config.output_folder == "/from-cli"

    def test_missing_config_raises(self) -> None:
        with pytest.raises(FileNotFoundError):
            load_config("/nonexistent/config.toml")

    def test_missing_required_field_raises(self, tmp_path: Path) -> None:
        config = tmp_path / "bad.toml"
        config.write_text('[imap]\nserver = "x"\n')
        with pytest.raises(KeyError):
            load_config(str(config))


# ---------------------------------------------------------------------------
# CLI parsing
# ---------------------------------------------------------------------------

class TestParseArgs:
    def test_defaults(self) -> None:
        args = parse_args([])
        assert args.config == "config.toml"
        assert args.output is None
        assert args.dry_run is False
        assert args.verbose is False

    def test_all_flags(self) -> None:
        args = parse_args(["-c", "my.toml", "-o", "/out", "--dry-run", "-v"])
        assert args.config == "my.toml"
        assert args.output == "/out"
        assert args.dry_run is True
        assert args.verbose is True

    def test_config_from_env(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("DETACH_CONFIG", "/env/config.toml")
        args = parse_args([])
        assert args.config == "/env/config.toml"


# ---------------------------------------------------------------------------
# Sender matching
# ---------------------------------------------------------------------------

class TestMatchSender:
    def test_exact_domain(self) -> None:
        assert _match_sender("billing@bank.com", ["*@bank.com"])

    def test_wildcard_user(self) -> None:
        assert _match_sender("user@bank.com", ["*@bank.com"])

    def test_wildcard_domain(self) -> None:
        assert _match_sender("billing@anywhere.com", ["billing@*"])

    def test_no_match(self) -> None:
        assert not _match_sender("other@other.com", ["*@bank.com"])

    def test_case_insensitive(self) -> None:
        assert _match_sender("BILLING@BANK.COM", ["*@bank.com"])

    def test_empty_patterns(self) -> None:
        assert not _match_sender("a@b.com", [])


# ---------------------------------------------------------------------------
# Filename generation
# ---------------------------------------------------------------------------

class TestMakeSafeFilename:
    def test_normal(self) -> None:
        result = make_safe_filename(
            "Tue, 15 Mar 2026 10:00:00 +0000",
            "Your Statement",
            "statement.pdf",
            "INBOX",
        )
        assert result == os.path.join("INBOX", "2026-03-15_Your Statement_statement.pdf")

    def test_unsafe_chars_sanitized(self) -> None:
        result = make_safe_filename(
            "Tue, 15 Mar 2026 10:00:00 +0000",
            'Bad: "chars" <here>',
            "file?.pdf",
            "INBOX",
        )
        assert "<" not in result
        assert ">" not in result
        assert '"' not in result
        assert "?" not in result

    def test_missing_date(self) -> None:
        result = make_safe_filename("", "Subject", "file.pdf", "INBOX")
        assert "unknown-date" in result

    def test_long_subject_truncated(self) -> None:
        long_subject = "A" * 100
        result = make_safe_filename(
            "Tue, 15 Mar 2026 10:00:00 +0000",
            long_subject,
            "file.pdf",
            "INBOX",
        )
        # Subject part should be at most 50 chars.
        parts = Path(result).name.split("_", 1)
        subject_and_rest = parts[1]  # after date prefix
        subject_part = subject_and_rest.rsplit("_", 1)[0]
        assert len(subject_part) <= 50


# ---------------------------------------------------------------------------
# PDF extraction
# ---------------------------------------------------------------------------

class TestExtractPdfAttachments:
    def test_extracts_pdf(self) -> None:
        msg = _make_email_with_pdf()
        pdfs = extract_pdf_attachments(msg)
        assert len(pdfs) == 1
        assert pdfs[0][0] == "statement.pdf"
        assert len(pdfs[0][1]) > 0

    def test_no_attachments(self) -> None:
        msg = MIMEText("Just text, no attachments")
        msg["From"] = "a@b.com"
        pdfs = extract_pdf_attachments(msg)
        assert pdfs == []

    def test_multiple_pdfs(self) -> None:
        msg = MIMEMultipart()
        msg["From"] = "a@b.com"
        msg.attach(MIMEText("body"))
        for name in ["a.pdf", "b.pdf"]:
            att = MIMEApplication(_make_pdf_bytes(), _subtype="pdf")
            att.add_header("Content-Disposition", "attachment", filename=name)
            msg.attach(att)
        pdfs = extract_pdf_attachments(msg)
        assert len(pdfs) == 2
        assert {p[0] for p in pdfs} == {"a.pdf", "b.pdf"}


# ---------------------------------------------------------------------------
# PDF password removal
# ---------------------------------------------------------------------------

class TestRemovePdfPassword:
    def test_decrypt_success(self) -> None:
        encrypted = _make_encrypted_pdf_bytes("mypass")
        decrypted = remove_pdf_password(encrypted, "mypass")
        # Should be valid PDF bytes.
        with pikepdf.open(BytesIO(decrypted)) as pdf:
            assert pdf is not None

    def test_wrong_password_raises(self) -> None:
        encrypted = _make_encrypted_pdf_bytes("mypass")
        with pytest.raises(pikepdf.PasswordError):
            remove_pdf_password(encrypted, "wrongpass")


# ---------------------------------------------------------------------------
# PDF saving
# ---------------------------------------------------------------------------

class TestSavePdf:
    def test_save_basic(self, tmp_path: Path) -> None:
        pdf_bytes = _make_pdf_bytes()
        path = str(tmp_path / "INBOX" / "test.pdf")
        assert save_pdf(pdf_bytes, path, "", dry_run=False)
        assert Path(path).exists()
        assert Path(path).read_bytes() == pdf_bytes

    def test_save_dry_run(self, tmp_path: Path) -> None:
        path = str(tmp_path / "test.pdf")
        assert save_pdf(b"data", path, "", dry_run=True)
        assert not Path(path).exists()

    def test_save_collision(self, tmp_path: Path) -> None:
        pdf_bytes = _make_pdf_bytes()
        path = str(tmp_path / "test.pdf")
        Path(path).write_bytes(b"existing")
        save_pdf(pdf_bytes, path, "", dry_run=False)
        assert (tmp_path / "test_1.pdf").exists()

    def test_save_with_password_removal(self, tmp_path: Path) -> None:
        encrypted = _make_encrypted_pdf_bytes("pass")
        path = str(tmp_path / "test.pdf")
        assert save_pdf(encrypted, path, "pass", dry_run=False)
        # Saved file should be decrypted (openable without password).
        with pikepdf.open(path) as pdf:
            assert pdf is not None


# ---------------------------------------------------------------------------
# Folder listing
# ---------------------------------------------------------------------------

class TestListMatchingFolders:
    def test_matches_patterns(self) -> None:
        conn = MagicMock()
        conn.list.return_value = (
            "OK",
            [
                b'(\\HasNoChildren) "/" "INBOX"',
                b'(\\HasNoChildren) "/" "Bills/Electric"',
                b'(\\HasNoChildren) "/" "Bills/Water"',
                b'(\\HasNoChildren) "/" "Trash"',
            ],
        )
        result = list_matching_folders(conn, ["INBOX", "Bills/*"])
        assert result == ["INBOX", "Bills/Electric", "Bills/Water"]

    def test_no_match(self) -> None:
        conn = MagicMock()
        conn.list.return_value = (
            "OK",
            [b'(\\HasNoChildren) "/" "Spam"'],
        )
        result = list_matching_folders(conn, ["INBOX"])
        assert result == []


# ---------------------------------------------------------------------------
# Folder selection quoting
# ---------------------------------------------------------------------------

class TestExtractSearchTerms:
    def test_wildcard_prefix(self) -> None:
        assert _extract_search_terms(["*@bank.com"]) == ["bank.com"]

    def test_wildcard_suffix(self) -> None:
        assert _extract_search_terms(["billing@*"]) == ["billing"]

    def test_both_wildcards(self) -> None:
        assert _extract_search_terms(["*@*.com"]) == [".com"]

    def test_no_wildcards(self) -> None:
        assert _extract_search_terms(["user@bank.com"]) == ["user@bank.com"]

    def test_multiple_patterns(self) -> None:
        terms = _extract_search_terms(["*@bank.com", "billing@*"])
        assert set(terms) == {"bank.com", "billing"}

    def test_all_wildcards_skipped(self) -> None:
        assert _extract_search_terms(["*"]) == []


class TestFetchMatchingEmailsSelectQuoting:
    def test_folder_with_special_chars_is_quoted(self) -> None:
        """Folder names like '[Gmail]/All Mail' must be quoted in SELECT."""
        conn = MagicMock()
        conn.select.return_value = ("OK", [b"0"])
        conn.uid.return_value = ("OK", [b""])

        fetch_matching_emails(conn, "[Gmail]/All Mail", ["*@x.com"])

        conn.select.assert_called_once_with('"[Gmail]/All Mail"', readonly=False)

    def test_simple_folder_is_still_quoted(self) -> None:
        conn = MagicMock()
        conn.select.return_value = ("OK", [b"0"])
        conn.uid.return_value = ("OK", [b""])

        fetch_matching_emails(conn, "INBOX", ["*@x.com"])

        conn.select.assert_called_once_with('"INBOX"', readonly=False)

    def test_uses_server_side_search_from(self) -> None:
        """Should use SEARCH FROM instead of fetching all headers."""
        msg = _make_email_with_pdf(sender="billing@bank.com")
        raw = msg.as_bytes()
        from_header = f"From: {msg['From']}\r\n".encode()

        conn = MagicMock()
        conn.select.return_value = ("OK", [b"1"])

        def uid_side_effect(command, *args):
            if command == "search":
                # Should be called with FROM, not ALL.
                assert "FROM" in args, (
                    f"Expected server-side FROM search, got: search {args}"
                )
                return ("OK", [b"1"])
            if command == "fetch":
                what = args[1]
                if "HEADER" in what:
                    return ("OK", [(b"1 (BODY[HEADER.FIELDS (FROM)]", from_header)])
                return ("OK", [(b"1 (RFC822", raw)])
            return ("OK", [])

        conn.uid.side_effect = uid_side_effect

        results = fetch_matching_emails(conn, "INBOX", ["*@bank.com"])
        assert len(results) == 1


# ---------------------------------------------------------------------------
# Integration: process_folder with mocked IMAP
# ---------------------------------------------------------------------------

class TestProcessFolder:
    def test_process_folder_saves_pdfs(self, tmp_path: Path) -> None:
        msg = _make_email_with_pdf()
        raw = msg.as_bytes()

        conn = MagicMock()
        conn.select.return_value = ("OK", [b"1"])
        conn.uid.side_effect = self._make_uid_side_effect(raw)

        config = Config(
            imap=ImapConfig(server="x", username="x", password="x"),
            filters=FiltersConfig(
                folder_patterns=["INBOX"],
                sender_patterns=["*@bank.com"],
            ),
            output_folder=str(tmp_path),
            pdf_password="",
            delete_after_archive=False,
        )

        saved, emails = process_folder(conn, "INBOX", config, dry_run=False)
        assert emails == 1
        assert saved == 1

        # Check file was created.
        pdf_files = list(tmp_path.rglob("*.pdf"))
        assert len(pdf_files) == 1

    def test_process_folder_dry_run(self, tmp_path: Path) -> None:
        msg = _make_email_with_pdf()
        raw = msg.as_bytes()

        conn = MagicMock()
        conn.select.return_value = ("OK", [b"1"])
        conn.uid.side_effect = self._make_uid_side_effect(raw)

        config = Config(
            imap=ImapConfig(server="x", username="x", password="x"),
            filters=FiltersConfig(sender_patterns=["*@bank.com"]),
            output_folder=str(tmp_path),
        )

        saved, emails = process_folder(conn, "INBOX", config, dry_run=True)
        assert saved == 1
        assert list(tmp_path.rglob("*.pdf")) == []

    def test_process_folder_deletes_email(self, tmp_path: Path) -> None:
        msg = _make_email_with_pdf()
        raw = msg.as_bytes()

        conn = MagicMock()
        conn.select.return_value = ("OK", [b"1"])
        conn.uid.side_effect = self._make_uid_side_effect(raw)

        config = Config(
            imap=ImapConfig(server="x", username="x", password="x"),
            filters=FiltersConfig(sender_patterns=["*@bank.com"]),
            output_folder=str(tmp_path),
            delete_after_archive=True,
        )

        process_folder(conn, "INBOX", config, dry_run=False)
        # Should have called store with \Deleted flag.
        store_calls = [
            c for c in conn.uid.call_args_list
            if c[0][0] == "store"
        ]
        assert len(store_calls) == 1
        conn.expunge.assert_called_once()

    @staticmethod
    def _make_uid_side_effect(raw_email: bytes):
        """Create a side_effect function for conn.uid() that handles
        search, fetch header, and fetch full message calls."""
        from_header = email.message_from_bytes(raw_email)["From"]
        header_bytes = f"From: {from_header}\r\n".encode()

        def side_effect(command, *args):
            if command == "search":
                return ("OK", [b"1"])
            if command == "fetch":
                uid = args[0]
                what = args[1]
                if "HEADER" in what:
                    return ("OK", [(b"1 (BODY[HEADER.FIELDS (FROM)]", header_bytes)])
                # Full fetch.
                return ("OK", [(b"1 (RFC822", raw_email)])
            if command == "store":
                return ("OK", [b"1"])
            return ("OK", [])

        return side_effect
