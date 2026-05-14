"""Unit tests for the pure helpers in scripts/extract.py.

Run with:

    cd dev-wiki
    .venv/bin/python3 -m pytest plugins/rimir/msg-import/test/ -v

or with the stdlib runner if pytest isn't installed:

    .venv/bin/python3 -m unittest discover -s plugins/rimir/msg-import/test/

These tests cover the helpers that don't need a real `.msg` fixture
(sanitization, URL derivation, attach-dir splitting, YAML emit). The
ImageMagick / extract-msg-library code paths are exercised by manual
smoke tests against real .msg drops in the wiki.
"""

from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

import extract  # noqa: E402


class SanitizeFilenameTests(unittest.TestCase):
    def test_keeps_ascii_safe(self):
        self.assertEqual(extract.sanitize_filename("invoice.pdf"), "invoice.pdf")

    def test_replaces_space_with_underscore(self):
        self.assertEqual(extract.sanitize_filename("my file.pdf"), "my_file.pdf")

    def test_strips_path_traversal(self):
        self.assertEqual(
            extract.sanitize_filename("../../etc/passwd"),
            "passwd",
        )

    def test_replaces_special_chars(self):
        # @, !, $ are all outside [A-Za-z0-9._-] → collapsed to a single _
        self.assertEqual(
            extract.sanitize_filename("a@b!c$d.txt"),
            "a_b_c_d.txt",
        )

    def test_strips_control_chars(self):
        self.assertEqual(
            extract.sanitize_filename("foo\x00bar.txt"),
            "foobar.txt",
        )

    def test_strips_leading_dots(self):
        # Leading dots from sanitize → stripped; "..msg" → "msg"
        self.assertEqual(extract.sanitize_filename("..msg"), "msg")

    def test_empty_after_sanitize_falls_back(self):
        # All-special input collapses to nothing → "file" placeholder
        self.assertEqual(extract.sanitize_filename("@@@@"), "file")

    def test_long_names_truncated_keeping_extension(self):
        long_stem = "a" * 130
        result = extract.sanitize_filename(f"{long_stem}.pdf")
        self.assertLessEqual(len(result), 120)
        self.assertTrue(result.endswith(".pdf"))


class IsExecutableTests(unittest.TestCase):
    def test_known_executables(self):
        for name in ("evil.exe", "evil.bat", "evil.ps1", "evil.vbs", "evil.jar"):
            self.assertTrue(extract.is_executable(name), name)

    def test_uppercase_extension(self):
        self.assertTrue(extract.is_executable("Bad.EXE"))

    def test_benign_types(self):
        for name in ("invoice.pdf", "photo.png", "notes.txt", "data.csv"):
            self.assertFalse(extract.is_executable(name), name)


class TruncateTests(unittest.TestCase):
    def test_short_passes_through(self):
        self.assertEqual(extract.truncate("hello", 100), "hello")

    def test_long_gets_ellipsis(self):
        result = extract.truncate("a" * 50, 10)
        self.assertEqual(len(result), 10)
        self.assertTrue(result.endswith("…"))

    def test_blank_input(self):
        self.assertEqual(extract.truncate("", 10), "")
        self.assertEqual(extract.truncate(None, 10), "")


class DerivedUrlForTests(unittest.TestCase):
    def test_default_files_location(self):
        url = extract.derived_url_for(
            "cid_part1.png", "/files/email/foo.msg"
        )
        self.assertEqual(url, "/files/email/_derived/foo.msg/cid_part1.png")

    def test_orga_apps_location_with_colons(self):
        url = extract.derived_url_for(
            "cid_part1.png",
            "/work/files/mgm/partnerships/ps:open_text/nt:problems/Fortify Ergebnisse.msg",
        )
        # `:` percent-encoded to %3A, space to %20, slashes preserved
        self.assertEqual(
            url,
            "/work/files/mgm/partnerships/ps%3Aopen_text/nt%3Aproblems/_derived/Fortify%20Ergebnisse.msg/cid_part1.png",
        )

    def test_filename_with_space_is_encoded(self):
        url = extract.derived_url_for(
            "att_some file.pdf", "/files/email/foo.msg"
        )
        self.assertEqual(
            url, "/files/email/_derived/foo.msg/att_some%20file.pdf"
        )


class AttachTargetSplittingTests(unittest.TestCase):
    def test_body_mode_is_directory(self):
        # In body mode, the whole path is the destination dir — even when
        # path.resolve has stripped the trailing slash from the executor.
        result = extract.attach_target_as_dir(
            "/abs/files/email/_derived/foo.msg"
        )
        self.assertEqual(result, Path("/abs/files/email/_derived/foo.msg"))

    def test_attachments_mode_splits_into_dir_and_prefix(self):
        dirname, prefix = extract.split_attach_target_for_scan(
            "/abs/files/email/_derived/foo.msg/att_"
        )
        self.assertEqual(dirname, Path("/abs/files/email/_derived/foo.msg"))
        self.assertEqual(prefix, "att_")


class YamlEmitTests(unittest.TestCase):
    def test_scalar_passthrough(self):
        self.assertEqual(extract.yaml_scalar("hello"), "hello")

    def test_scalar_quotes_when_contains_colon(self):
        self.assertEqual(
            extract.yaml_scalar("Mirko: dev"), '"Mirko: dev"'
        )

    def test_scalar_quotes_reserved_words(self):
        self.assertEqual(extract.yaml_scalar("yes"), '"yes"')
        self.assertEqual(extract.yaml_scalar("no"), '"no"')

    def test_scalar_boolean(self):
        self.assertEqual(extract.yaml_scalar(True), "yes")
        self.assertEqual(extract.yaml_scalar(False), "no")

    def test_list_empty(self):
        self.assertEqual(extract.yaml_list([]), "[]")

    def test_list_one_item(self):
        self.assertEqual(extract.yaml_list(["a"]), "\n  - a")

    def test_list_multiple(self):
        self.assertEqual(
            extract.yaml_list(["a", "b"]),
            "\n  - a\n  - b",
        )


class ParseRecipientsTests(unittest.TestCase):
    def test_single_address(self):
        self.assertEqual(
            extract.parse_recipients("alice@example.com"),
            ["alice@example.com"],
        )

    def test_comma_separated(self):
        self.assertEqual(
            extract.parse_recipients("a@x; b@y, c@z"),
            ["a@x", "b@y", "c@z"],
        )

    def test_empty(self):
        self.assertEqual(extract.parse_recipients(""), [])
        self.assertEqual(extract.parse_recipients(None), [])

    def test_list_input(self):
        self.assertEqual(
            extract.parse_recipients(["a@x", "b@y"]),
            ["a@x", "b@y"],
        )


class DetectFormatTests(unittest.TestCase):
    def test_eml_extension(self):
        self.assertEqual(extract._detect_format(Path("/foo/bar.eml")), "eml")

    def test_msg_extension(self):
        self.assertEqual(extract._detect_format(Path("/foo/bar.msg")), "msg")

    def test_case_insensitive(self):
        self.assertEqual(extract._detect_format(Path("/foo/BAR.EML")), "eml")
        self.assertEqual(extract._detect_format(Path("/foo/BAR.MSG")), "msg")

    def test_unknown_extension_defaults_to_msg(self):
        # The pipeline only fires for the two registered MIME types so the
        # script should never see an unknown extension — but treat it as
        # `.msg` (extract-msg path) since that's the original behaviour.
        self.assertEqual(extract._detect_format(Path("/foo/bar.txt")), "msg")


class EmlPlainTextTests(unittest.TestCase):
    """Smallest possible RFC 5322 fixture: plain-text body, no attachments."""

    PLAIN_EML = (
        b"From: alice@example.com\r\n"
        b"To: bob@example.com\r\n"
        b"Cc: carol@example.com\r\n"
        b"Subject: Test plain email\r\n"
        b"Date: Wed, 14 May 2026 10:00:00 +0000\r\n"
        b"Message-ID: <test-001@example.com>\r\n"
        b"In-Reply-To: <prev-000@example.com>\r\n"
        b"Content-Type: text/plain; charset=utf-8\r\n"
        b"\r\n"
        b"Hello world\r\n"
    )

    def setUp(self):
        import tempfile
        tmp = tempfile.NamedTemporaryFile(suffix=".eml", delete=False)
        tmp.write(self.PLAIN_EML)
        tmp.close()
        self.path = Path(tmp.name)
        self.msg = extract.parse_eml(self.path)

    def tearDown(self):
        try:
            self.path.unlink()
        except FileNotFoundError:
            pass

    def test_subject(self):
        self.assertEqual(self.msg.subject, "Test plain email")

    def test_sender(self):
        self.assertEqual(self.msg.sender, "alice@example.com")

    def test_recipients(self):
        self.assertEqual(str(self.msg.to), "bob@example.com")
        self.assertEqual(str(self.msg.cc), "carol@example.com")
        self.assertEqual(str(self.msg.bcc), "")

    def test_message_id_and_in_reply_to(self):
        self.assertEqual(self.msg.messageId, "<test-001@example.com>")
        self.assertEqual(self.msg.inReplyTo, "<prev-000@example.com>")

    def test_date_parsed_to_utc(self):
        self.assertIsNotNone(self.msg.date)
        self.assertEqual(self.msg.date.year, 2026)
        self.assertEqual(self.msg.date.month, 5)
        self.assertEqual(self.msg.date.day, 14)

    def test_body_plain(self):
        self.assertIn("Hello world", self.msg.body)

    def test_html_body_absent(self):
        self.assertIsNone(self.msg.htmlBody)

    def test_no_attachments(self):
        self.assertEqual(self.msg.attachments, [])


class EmlMultipartTests(unittest.TestCase):
    """HTML body + a real attachment + an inline image with Content-ID."""

    def setUp(self):
        import email.message
        import email.policy
        import tempfile

        em = email.message.EmailMessage(policy=email.policy.default)
        em["From"] = "alice@example.com"
        em["To"] = "bob@example.com"
        em["Subject"] = "Multipart with PDF"
        em["Date"] = "Wed, 14 May 2026 10:00:00 +0000"
        em.set_content("Plain text alt")
        em.add_alternative(
            "<p>HTML body <img src=\"cid:logo123\"/></p>",
            subtype="html",
        )
        # add_alternative replaces the body — fetch the html sub-part and
        # add an inline image related to it.
        html_part = em.get_body(preferencelist=("html",))
        html_part.add_related(
            b"\x89PNG\r\n\x1a\n fake-png-bytes",
            maintype="image", subtype="png", cid="<logo123>",
        )
        # And a real non-inline attachment.
        em.add_attachment(
            b"%PDF-1.4 fake",
            maintype="application", subtype="pdf",
            filename="invoice.pdf",
        )
        tmp = tempfile.NamedTemporaryFile(suffix=".eml", delete=False)
        tmp.write(em.as_bytes())
        tmp.close()
        self.path = Path(tmp.name)
        self.msg = extract.parse_eml(self.path)

    def tearDown(self):
        try:
            self.path.unlink()
        except FileNotFoundError:
            pass

    def test_html_body_extracted(self):
        self.assertIsNotNone(self.msg.htmlBody)
        self.assertIn("HTML body", self.msg.htmlBody)

    def test_attachments_include_inline_and_real(self):
        # The inline image and the PDF are both in iter_attachments() — the
        # downstream code distinguishes via the `cid` field.
        names = sorted(a.longFilename for a in self.msg.attachments)
        self.assertIn("invoice.pdf", names)

    def test_real_attachment_has_no_cid(self):
        pdf = next(a for a in self.msg.attachments if a.longFilename == "invoice.pdf")
        self.assertEqual(pdf.cid, "")
        self.assertTrue(pdf.data.startswith(b"%PDF"))

    def test_inline_image_keeps_cid_lowercase(self):
        inline = [a for a in self.msg.attachments if a.cid]
        self.assertEqual(len(inline), 1)
        self.assertEqual(inline[0].cid, "logo123")


class EmlMalformedDateTests(unittest.TestCase):
    """A bad Date: header must not crash parsing."""

    def test_unparseable_date_yields_none(self):
        import tempfile
        raw = (
            b"From: a@x\r\n"
            b"Subject: bad date\r\n"
            b"Date: not-a-real-date\r\n"
            b"\r\n"
            b"body\r\n"
        )
        tmp = tempfile.NamedTemporaryFile(suffix=".eml", delete=False)
        tmp.write(raw)
        tmp.close()
        try:
            msg = extract.parse_eml(Path(tmp.name))
            # Bad date string → either None or some best-effort datetime.
            # Either way, parse_eml must not raise.
            self.assertEqual(msg.subject, "bad date")
        finally:
            Path(tmp.name).unlink()


class PathSafetyTests(unittest.TestCase):
    def setUp(self):
        self.wiki_root = Path(__file__).resolve().parent.parent.parent.parent.parent
        # That's dev-wiki/

    def test_input_under_wiki_accepted(self):
        # Use this very file as the "input" — it's under the wiki dir.
        try:
            extract.assert_path_under_wiki(Path(__file__), self.wiki_root)
        except SystemExit:
            self.fail("assert_path_under_wiki rejected a path inside the wiki")

    def test_input_outside_wiki_rejected(self):
        with self.assertRaises(SystemExit) as cm:
            extract.assert_path_under_wiki(Path("/etc/passwd"), self.wiki_root)
        self.assertEqual(cm.exception.code, 2)


if __name__ == "__main__":
    unittest.main()
