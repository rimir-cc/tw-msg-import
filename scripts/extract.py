#!/usr/bin/env python3
"""
extract.py — Outlook .msg / RFC 5322 .eml → markdown+frontmatter for
rimir/msg-import.

Invoked by file-pipeline via runner-actions.json. Three modes:

  --mode=body
    Parse <input>, write inline images (cid:) into <attach-dir>, pipe the
    HTML body through pandoc, emit YAML frontmatter + markdown to stdout.
    The pipeline captures stdout into a `.email` tiddler (type
    text/x-frontmattered-markdown).

  --mode=attachments
    Parse <input>, write all non-inline attachments into <attach-dir>. The
    pipeline picks them up via scanDir and creates artifact tiddlers.

  --mode=thumb
    Render a small PNG preview of the message (subject text in a caption)
    via ImageMagick.

Format dispatch is by filename extension:

  *.msg  → parsed via the `extract-msg` library (Outlook binary format)
  *.eml  → parsed via Python's stdlib `email` module (RFC 5322 plain text)

Both paths share the same body / attachments / thumb emission logic — the
`.eml` parser projects an `EmailMessage` through the subset of the
extract-msg API this script uses (subject / sender / to / cc / bcc / date
/ message-id / in-reply-to / htmlBody / body / attachments[]).

Security: realpath of --input is asserted to live under $TW_WIKI_PATH (cwd
fallback). Attachment filenames are sanitized (no traversal, no control
chars, ASCII-safe). Executable extensions are quarantined unless
--allow-executables is passed.

Dependencies: pandoc (system), and `extract-msg` (pip) for the `.msg` path
only. The `.eml` path uses Python's stdlib `email` module — no extra pip
install required.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path

EXECUTABLE_EXTS = {
    ".exe", ".bat", ".cmd", ".ps1", ".vbs", ".vbe", ".js", ".jse",
    ".jar", ".scr", ".com", ".msi", ".pif", ".reg", ".lnk", ".hta",
    ".cpl", ".msc", ".ws", ".wsf", ".wsh",
}

SAFE_FILENAME_RE = re.compile(r"[^A-Za-z0-9._-]+")
CTRL_RE = re.compile(r"[\x00-\x1f\x7f]")


def die(msg: str, code: int = 1) -> None:
    sys.stderr.write("extract_msg: " + msg + "\n")
    sys.exit(code)


def assert_path_under_wiki(input_path: Path, wiki_path: Path) -> None:
    try:
        real_input = input_path.resolve(strict=True)
        real_wiki = wiki_path.resolve(strict=True)
    except FileNotFoundError as exc:
        die(f"path does not exist: {exc.filename}", 2)
    try:
        real_input.relative_to(real_wiki)
    except ValueError:
        die(f"refusing to operate on path outside wiki: {real_input}", 2)


def sanitize_filename(name: str) -> str:
    name = CTRL_RE.sub("", name)
    name = name.replace("\\", "/").rsplit("/", 1)[-1]
    name = name.lstrip(".")
    name = SAFE_FILENAME_RE.sub("_", name)
    name = name.strip("._-") or "file"
    if len(name) > 120:
        stem, dot, ext = name.rpartition(".")
        if dot and len(ext) <= 8:
            name = stem[: 120 - len(ext) - 1] + "." + ext
        else:
            name = name[:120]
    return name


def is_executable(name: str) -> bool:
    ext = os.path.splitext(name)[1].lower()
    return ext in EXECUTABLE_EXTS


def unique_path(dir_path: Path, name: str) -> Path:
    candidate = dir_path / name
    if not candidate.exists():
        return candidate
    stem, dot, ext = name.rpartition(".")
    base = stem if dot else name
    suffix = ext if dot else ""
    digest = hashlib.sha1(os.urandom(8)).hexdigest()[:8]
    new_name = f"{base}-{digest}" + (f".{suffix}" if suffix else "")
    return dir_path / new_name


def format_addr(entry) -> str:
    if not entry:
        return ""
    if isinstance(entry, (list, tuple)):
        return ", ".join(format_addr(e) for e in entry if e)
    return str(entry).strip()


def parse_recipients(field) -> list[str]:
    if not field:
        return []
    if isinstance(field, str):
        parts = [p.strip() for p in re.split(r"[;,]", field) if p.strip()]
        return parts
    if isinstance(field, (list, tuple)):
        out: list[str] = []
        for item in field:
            if isinstance(item, str):
                out.extend(parse_recipients(item))
            else:
                s = str(item).strip()
                if s:
                    out.append(s)
        return out
    return [str(field).strip()]


def yaml_quote(value: str) -> str:
    needs_quote = (
        not value
        or value[0] in "!&*?{[|>%@`'\""
        or value.lower() in {"yes", "no", "true", "false", "null", "~"}
        or ":" in value
        or "#" in value
        or value.strip() != value
    )
    if needs_quote:
        return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'
    return value


def yaml_scalar(value) -> str:
    if value is None:
        return '""'
    if isinstance(value, bool):
        return "yes" if value else "no"
    return yaml_quote(str(value))


def yaml_list(values: list[str]) -> str:
    if not values:
        return "[]"
    return "\n" + "\n".join(f"  - {yaml_quote(v)}" for v in values)


def emit_frontmatter(meta: dict) -> str:
    lines = ["---"]
    for key, value in meta.items():
        if isinstance(value, list):
            lines.append(f"{key}:{yaml_list(value)}")
        else:
            lines.append(f"{key}: {yaml_scalar(value)}")
    lines.append("---")
    return "\n".join(lines) + "\n"


def html_to_markdown(html: str) -> str:
    proc = subprocess.run(
        ["pandoc", "-f", "html", "-t", "markdown", "--wrap=none"],
        input=html.encode("utf-8"),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    if proc.returncode != 0:
        die("pandoc failed: " + proc.stderr.decode("utf-8", "replace"), 3)
    return proc.stdout.decode("utf-8", "replace")


def text_to_markdown(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def isoformat_date(value) -> str:
    if not value:
        return ""
    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.replace(tzinfo=timezone.utc).isoformat()
        return value.astimezone(timezone.utc).isoformat()
    return str(value)


def attach_target_as_dir(value: str) -> Path:
    """Body mode: --attach-dir is always a directory path (mkdir if needed).

    file-pipeline's executor calls `path.resolve()` on the step's `output`,
    which strips trailing slashes — so `_derived/{{basename}}/` arrives here
    as `/abs/_derived/{{basename}}` without a slash. Treat the whole path
    as the destination directory in body mode (we don't need a filename
    prefix; inline images are named `cid_<sanitized>.<ext>`).
    """
    return Path(value)


def split_attach_target_for_scan(value: str) -> tuple[Path, str]:
    """Attachments mode: split --attach-dir into (dirname, basename-prefix).

    The attachments step's output template is `_derived/{{basename}}/att_`,
    which becomes `/abs/_derived/{{basename}}/att_` after path.resolve. The
    leaf `att_` is the file-prefix scanDir matches against. file-pipeline's
    executor creates `path.dirname(outputPath)`, so the parent already
    exists when we get here.
    """
    p = Path(value)
    return p.parent, p.name


def derived_url_for(filename: str, source_uri: str) -> str:
    """URL for a file in `_derived/<basename>/` given the source's canonical URI.

    The wiki may serve binaries from any location (default `/files/`, but
    orga-apps and friends use custom locations with their own `uriPrefix` →
    `basePath` mapping). The filesystem path is NOT a served URL — we must
    derive the URL from the source tiddler's `_canonical_uri` instead.

    `source_uri` looks like `/work/mgm/.../Fortify Ergebnisse.msg`. The
    derived files live at `<dirname>/_derived/<basename>/<filename>`,
    served via the same location's route. Percent-encode the result so
    spaces and `:` etc. survive CommonMark URL parsing.
    """
    slash = source_uri.rfind("/")
    if slash < 0:
        dir_part = ""
        base = source_uri
    else:
        dir_part = source_uri[: slash + 1]
        base = source_uri[slash + 1 :]
    url = f"{dir_part}_derived/{base}/{filename}"
    return urllib.parse.quote(url, safe="/")


def _detect_format(input_path: Path) -> str:
    """Pick a parser based on filename extension.

    `.eml` → RFC 5322 (stdlib `email` module); everything else (including
    `.msg`) is routed to `extract-msg`. The magic-byte sniff would be more
    robust but adds little here — file-upload's dropzone already validates
    by MIME, and the pipeline only fires for the two registered types.
    """
    if input_path.suffix.lower() == ".eml":
        return "eml"
    return "msg"


class _EmlAttachment:
    """Minimal attachment shape compatible with `extract-msg`'s Attachment.

    Exposes the subset used downstream: `.cid` (lowercase, brackets
    stripped), `.longFilename` / `.shortFilename` / `.displayName` (all
    the same string here — RFC 5322 doesn't have the Outlook split), and
    `.data` (bytes).
    """

    __slots__ = ("cid", "longFilename", "shortFilename", "displayName", "data")

    def __init__(self, cid: str, name: str, data: bytes) -> None:
        self.cid = cid
        self.longFilename = name
        self.shortFilename = name
        self.displayName = name
        self.data = data


class _EmlMsg:
    """Adapter projecting Python's `email.message.EmailMessage` through the
    subset of extract-msg's `Message` API this script uses.

    Subject / from / to / cc / bcc / date / message-id / in-reply-to /
    htmlBody / body / attachments[] are read from the parsed message and
    exposed as the attributes `build_meta` and the body / attachments
    runners read. Lets `run_body`, `run_attachments` and
    `render_thumbnail` handle `.msg` and `.eml` inputs without branching.
    """

    def __init__(self, source) -> None:
        self.subject = (source.get("Subject") or "").strip()
        self.sender = (source.get("From") or "").strip()
        self.senderName = self.sender
        self.to = source.get("To") or ""
        self.cc = source.get("Cc") or ""
        self.bcc = source.get("Bcc") or ""
        date_str = source.get("Date") or ""
        parsed_dt = None
        if date_str:
            try:
                from email.utils import parsedate_to_datetime
                parsed_dt = parsedate_to_datetime(date_str)
            except (TypeError, ValueError):
                parsed_dt = None
        self.date = parsed_dt
        self.parsedDate = parsed_dt
        self.messageId = (source.get("Message-ID") or "").strip()
        self.inReplyTo = (source.get("In-Reply-To") or "").strip()
        html_part = source.get_body(preferencelist=("html",))
        plain_part = source.get_body(preferencelist=("plain",))
        self.htmlBody = html_part.get_content() if html_part is not None else None
        self.html = self.htmlBody
        self.body = plain_part.get_content() if plain_part is not None else ""

        # Collect every non-body leaf, including inline images attached
        # to the html body via `multipart/related` — Python's
        # iter_attachments() treats those as part of the body group and
        # skips them. The downstream `collect_attachments` /
        # `run_attachments` code distinguishes inline vs. non-inline via
        # the `cid` field, so we just need to surface BOTH here.
        body_ids = {id(p) for p in (html_part, plain_part) if p is not None}
        self.attachments = []
        seen = set()
        for part in source.walk():
            if part.is_multipart():
                continue
            if id(part) in body_ids or id(part) in seen:
                continue
            seen.add(id(part))
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            name = part.get_filename() or "attachment"
            cid_raw = part.get("Content-ID", "") or ""
            cid = cid_raw.strip().strip("<>").lower()
            self.attachments.append(_EmlAttachment(cid, name, bytes(payload)))


def parse_eml(input_path: Path) -> _EmlMsg:
    """Parse an RFC 5322 .eml file via Python's stdlib `email` module."""
    import email
    import email.policy
    with open(input_path, "rb") as f:
        raw = email.message_from_binary_file(f, policy=email.policy.default)
    return _EmlMsg(raw)


def open_message(input_path: Path):
    """Dispatch to the correct parser based on filename extension."""
    if _detect_format(input_path) == "eml":
        return parse_eml(input_path)
    try:
        import extract_msg
    except ImportError:
        die("missing dependency: pip install extract-msg", 4)
    return extract_msg.openMsg(str(input_path))


def collect_attachments(msg, attach_dir: Path, source_uri: str,
                        allow_executables: bool, quarantine_log: list[dict]
                        ) -> tuple[list[dict], dict[str, str]]:
    """Walk msg.attachments for inline images (those with a CID).

    Writes inline images only. Real attachments are handled in
    attachments mode. Returns (files_written, cid_to_url).
    """
    written: list[dict] = []
    cid_map: dict[str, str] = {}
    for att in getattr(msg, "attachments", []) or []:
        cid = (getattr(att, "cid", "") or "").strip().strip("<>").lower()
        if not cid:
            continue
        raw_name = (
            getattr(att, "longFilename", None)
            or getattr(att, "shortFilename", None)
            or getattr(att, "displayName", None)
            or f"cid-{cid}"
        )
        data = getattr(att, "data", None)
        if data is None or not isinstance(data, (bytes, bytearray)):
            continue
        ext = os.path.splitext(raw_name)[1] or ".bin"
        safe = sanitize_filename(f"cid_{cid}{ext}")
        if is_executable(safe) and not allow_executables:
            quarantine_log.append({
                "original": raw_name,
                "reason": "executable extension blocked",
            })
            continue
        target = unique_path(attach_dir, safe)
        target.write_bytes(bytes(data))
        url = derived_url_for(target.name, source_uri)
        written.append({
            "filename": target.name,
            "path": str(target),
            "url": url,
            "cid": cid,
            "original": raw_name,
        })
        cid_map[cid] = url
    return written, cid_map


def rewrite_cid_refs(markdown: str, cid_map: dict[str, str]) -> str:
    if not cid_map:
        return markdown

    def repl(match: re.Match) -> str:
        cid = match.group(1).lower()
        url = cid_map.get(cid)
        return url if url else match.group(0)

    return re.sub(r"cid:([A-Za-z0-9._@\-]+)", repl, markdown)


def build_meta(msg, has_attachments: bool, quarantined: list[dict]) -> dict:
    sender = format_addr(getattr(msg, "sender", "") or "")
    if not sender:
        sender = format_addr(getattr(msg, "senderName", "") or "")
    meta: dict = {
        "type": "text/x-frontmattered-markdown",
        "msg-subject": (getattr(msg, "subject", "") or "").strip(),
        "msg-from": sender,
        "msg-to": parse_recipients(getattr(msg, "to", "")),
        "msg-cc": parse_recipients(getattr(msg, "cc", "")),
        "msg-bcc": parse_recipients(getattr(msg, "bcc", "")),
        "msg-date": isoformat_date(getattr(msg, "date", None) or getattr(msg, "parsedDate", None)),
        "msg-message-id": (getattr(msg, "messageId", "") or "").strip(),
        "msg-in-reply-to": (getattr(msg, "inReplyTo", "") or "").strip(),
        "msg-has-attachments": has_attachments,
        "tags": "email",
    }
    if quarantined:
        meta["msg-quarantined-attachments"] = [q["original"] for q in quarantined]
    return meta


def run_body(args: argparse.Namespace) -> None:
    input_path = Path(args.input)
    attach_dir = attach_target_as_dir(args.attach_dir)
    attach_dir.mkdir(parents=True, exist_ok=True)

    msg = open_message(input_path)
    quarantined: list[dict] = []
    files, cid_map = collect_attachments(
        msg, attach_dir, args.source_uri, args.allow_executables, quarantined,
    )
    has_real_attachments = any(
        not ((getattr(a, "cid", "") or "").strip().strip("<>"))
        for a in (getattr(msg, "attachments", []) or [])
    )
    html = getattr(msg, "htmlBody", None) or getattr(msg, "html", None)
    if html and isinstance(html, bytes):
        html = html.decode("utf-8", "replace")
    if html:
        body_md = html_to_markdown(html)
    else:
        body_md = text_to_markdown(getattr(msg, "body", "") or "")
    body_md = rewrite_cid_refs(body_md, cid_map)

    meta = build_meta(
        msg,
        has_attachments=has_real_attachments or bool(files),
        quarantined=quarantined,
    )
    subject = meta["msg-subject"] or "(no subject)"
    sys.stdout.write(emit_frontmatter(meta))
    sys.stdout.write("\n# " + subject + "\n\n")
    sys.stdout.write(body_md.rstrip() + "\n")


def truncate(value: str, limit: int) -> str:
    value = (value or "").strip()
    if len(value) <= limit:
        return value
    return value[: limit - 1] + "…"  # ellipsis


def render_thumbnail(input_path: Path, output_path: Path, size: int) -> None:
    """Render a PNG previewing the email — subject only, large and wrapped.

    Uses ImageMagick's `magick` binary (already present per CLAUDE.md).
    `caption:` auto-wraps the text to fit the canvas; pointsize is chosen
    big enough to stay readable inside an attachment tile.
    """
    msg = open_message(input_path)

    subject = truncate(getattr(msg, "subject", "") or "(no subject)", 200)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    cmd = [
        "magick",
        "-size", f"{size}x{size}",
        "-background", "#fafafa",
        "-fill", "#222",
        "-gravity", "Center",
        "-pointsize", "16",
        f"caption:{subject}",
        "-bordercolor", "#ddd",
        "-border", "1",
        str(output_path),
    ]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if proc.returncode != 0:
        die("magick failed: " + proc.stderr.decode("utf-8", "replace"), 5)


def run_thumb(args: argparse.Namespace) -> None:
    size = int(getattr(args, "size", 200) or 200)
    render_thumbnail(Path(args.input), Path(args.output), size)


def run_attachments(args: argparse.Namespace) -> None:
    input_path = Path(args.input)
    attach_dir, attach_prefix = split_attach_target_for_scan(args.attach_dir)
    attach_dir.mkdir(parents=True, exist_ok=True)

    msg = open_message(input_path)
    written: list[dict] = []
    quarantined: list[dict] = []
    for att in getattr(msg, "attachments", []) or []:
        cid = (getattr(att, "cid", "") or "").strip().strip("<>").lower()
        # Inline images are handled in body mode; skip here.
        if cid:
            continue
        raw_name = (
            getattr(att, "longFilename", None)
            or getattr(att, "shortFilename", None)
            or getattr(att, "displayName", None)
            or "attachment"
        )
        data = getattr(att, "data", None)
        if data is None or not isinstance(data, (bytes, bytearray)):
            continue
        safe = sanitize_filename(f"{attach_prefix}{raw_name}")
        if is_executable(safe) and not args.allow_executables:
            quarantined.append({"original": raw_name})
            continue
        target = unique_path(attach_dir, safe)
        target.write_bytes(bytes(data))
        written.append({
            "filename": target.name,
            "url": derived_url_for(target.name, args.source_uri),
            "original": raw_name,
        })

    summary = {"written": written, "quarantined": quarantined}
    sys.stdout.write(json.dumps(summary, indent=2) + "\n")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Extract Outlook .msg to markdown+frontmatter, attachments, or a thumbnail PNG")
    parser.add_argument("--mode", choices=("body", "attachments", "thumb"), required=True)
    parser.add_argument("--input", required=True, help="absolute path to .msg file")
    parser.add_argument("--attach-dir", help="absolute path to attachment output dir (body / attachments modes)")
    parser.add_argument("--source-uri",
                        help="canonical_uri of the parent .msg (body / attachments modes); used to derive served URLs for written files")
    parser.add_argument("--output", help="absolute path of the PNG to write (thumb mode)")
    parser.add_argument("--size", type=int, default=200,
                        help="thumb edge length in px (thumb mode, default 200)")
    parser.add_argument("--allow-executables", action="store_true",
                        help="do not quarantine .exe/.bat/etc. attachments")
    args = parser.parse_args(argv)

    wiki_path = Path(os.environ.get("TW_WIKI_PATH", os.getcwd()))
    assert_path_under_wiki(Path(args.input), wiki_path)

    if args.mode == "thumb":
        if not args.output:
            die("--output is required for --mode=thumb", 2)
        # Output path's parent must resolve under the wiki.
        probe = Path(args.output).parent
        while probe and not probe.exists():
            probe = probe.parent
        try:
            probe.resolve().relative_to(wiki_path.resolve())
        except ValueError:
            die(f"output is outside wiki: {args.output}", 2)
        run_thumb(args)
        return 0

    if not args.attach_dir:
        die(f"--attach-dir is required for --mode={args.mode}", 2)
    if not args.source_uri:
        die(f"--source-uri is required for --mode={args.mode}", 2)

    # attach-dir's effective directory must resolve under the wiki.
    if args.mode == "body":
        attach_dir_check = attach_target_as_dir(args.attach_dir)
    else:
        attach_dir_check, _prefix = split_attach_target_for_scan(args.attach_dir)
    try:
        probe = attach_dir_check
        while probe and not probe.exists():
            probe = probe.parent
        probe.resolve().relative_to(wiki_path.resolve())
    except ValueError:
        die(f"attach-dir is outside wiki: {args.attach_dir}", 2)

    if args.mode == "body":
        run_body(args)
    else:
        run_attachments(args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
