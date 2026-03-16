# detach

Email PDF attachment archiver. Connects to an IMAP server, scans matching folders for emails from matching senders, downloads PDF attachments, optionally removes passwords, saves them locally, and optionally deletes the processed emails.

## Install

Requires Python 3.13+ and [uv](https://docs.astral.sh/uv/).

```sh
uv sync
```

## Usage

```
uv run detach [-c CONFIG] [-o OUTPUT] [--dry-run] [-v]
```

| Flag | Description |
|---|---|
| `-c`, `--config` | Config file path (env: `DETACH_CONFIG`, default: `config.toml`) |
| `-o`, `--output` | Output directory (env: `DETACH_OUTPUT_DIR`, overrides config) |
| `--dry-run` | Log actions without saving files or deleting emails |
| `-v`, `--verbose` | Debug logging |

### Quick start

```sh
cp config.example.toml config.toml
# Edit config.toml with your IMAP credentials and filters
uv run detach --dry-run -v   # preview what would happen
uv run detach                # run for real
```

## Configuration

See [`config.example.toml`](config.example.toml) for a complete example.

```toml
[imap]
server = "imap.example.com"
username = "user@example.com"
password = "secret"
port = 993        # optional, default 993
use_ssl = true    # optional, default true
# proxy = "socks5://proxy.host:1080"   # optional

[filters]
folder_patterns = ["INBOX", "Bills/*"]       # glob patterns
sender_patterns = ["*@bankofamerica.com"]     # glob patterns

[output]
folder = "~/Documents/attachments"

[pdf]
password = ""     # optional, for unlocking encrypted PDFs

[behavior]
delete_after_archive = false
```

### Proxy support

Route IMAP connections through a SOCKS4, SOCKS5, or HTTP CONNECT proxy.

Set it explicitly in the config (takes priority over environment variables):

```toml
[imap]
proxy = "socks5://user:pass@proxy.host:1080"
```

Or rely on the standard proxy environment variables as a fallback:

```sh
all_proxy=socks5://127.0.0.1:1080 uv run detach --dry-run
```

Supported env vars (checked in order): `all_proxy`, `ALL_PROXY`, `http_proxy`, `HTTP_PROXY`.
Supported schemes: `socks4`, `socks5`, `http`.

### Output folder priority

1. CLI `--output` flag
2. `DETACH_OUTPUT_DIR` environment variable
3. `output.folder` in config file

### Pattern matching

Folder and sender patterns use glob syntax (`fnmatch`):

- `*` matches any sequence of characters
- `?` matches any single character
- `[seq]` matches any character in *seq*

Examples: `INBOX`, `Bills/*`, `*@bank.com`, `billing@*`.

## Saved file layout

PDFs are saved under the output folder, organized by IMAP folder name:

```
output/
  INBOX/
    2026-03-15_Your Statement_statement.pdf
  Bills/Electric/
    2026-03-01_March Invoice_invoice.pdf
```

Duplicate filenames are resolved by appending `_1`, `_2`, etc.

## Tests

```sh
uv run python -m pytest test_main.py -v
```
