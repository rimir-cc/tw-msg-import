# tw-msg-import

Drop Outlook `.msg` **or** RFC 5322 `.eml` files into a TiddlyWiki and get a clean markdown tiddler with YAML frontmatter, plus the attachments as artifact child tiddlers. The original blob is preserved as the canonical asset.

Format dispatch is by filename extension: `.msg` parsed via the `extract-msg` Python library, `.eml` parsed via Python's stdlib `email` module. Both paths produce the same output tiddler shape.

## Features

- Drag-and-drop import into any `<$file-dropzone>` — handles both `.msg` and `.eml`.
- Frontmatter metadata (subject, from, to, cc, bcc, date, message-id, in-reply-to) extracted into YAML fields and round-tripped to disk by `rimir/frontmatter`.
- Email body converted from HTML to markdown via pandoc (stdin, no shell, no filename arg).
- Each attachment becomes a child tiddler with `_artifact_source` linking back to the parent; rename and delete cascade for free.
- Inline images preserved: `cid:` references in the body are rewritten to wiki-relative URLs pointing at the saved inline-image files.
- Optional LLM summary step gated by `$:/config/rimir/msg-import/llm-summary`.
- Executable attachments (`.exe`, `.bat`, `.ps1`, `.vbs`, `.jar`, etc.) are quarantined by default.

## Prerequisites

- `rimir/file-upload` (≥ 0.1.24 — ships both the `application/vnd.ms-outlook` and `message/rfc822` MIME entries and `email/` subfolder routing for both)
- `rimir/file-pipeline`
- `rimir/runner`
- `rimir/frontmatter`
- Python 3 — stdlib only for `.eml`; `pip install extract-msg` only if you intend to drop `.msg` files
- `pandoc` on `PATH`

## Install

1. Drop the plugin folder under `dev-wiki/plugins/rimir/msg-import/`.
2. Add `"rimir/msg-import"` to your wiki's `tiddlywiki.info` plugin list.
3. Merge the two entries from `runner-actions/msg.json` into `dev-wiki/runner-actions.json` (restart required — runner-actions are read at boot).
4. `pip install -r dev-wiki/plugins/rimir/msg-import/scripts/requirements.txt`.

## Output

For an imported `meeting.msg` (or `meeting.eml`), three tiddler groups are created:

- `meeting.msg` — original binary, `type: application/vnd.ms-outlook` (or `message/rfc822` for `.eml`)
- `meeting.msg.email` — markdown body, `type: text/x-frontmattered-markdown`
- `meeting.msg.attachments/<filename>` — one tiddler per attachment

## License

MIT — see `LICENSE.md`.
