/*\
title: $:/plugins/rimir/msg-import/test/test-email-cascade.js
type: application/javascript
tags: [[$:/tags/test-spec]]

Filter pins for `cascade/email.tid` and the attachment filter inside
`body/email.tid`. The wikitext templates themselves can't be unit-rendered
here without a navigator + appify stand-in, but the filter shapes that
drive them are testable in isolation.

\*/
"use strict";

describe("msg-import: cascade/email dispatches to body/email", function () {

	var CASCADE_FILTER = $tw.wiki.getTiddlerText("$:/plugins/rimir/msg-import/cascade/email") || "";
	var EMAIL_BODY = "$:/plugins/rimir/msg-import/body/email";

	var added = [];

	function add(title, fields) {
		var base = {title: title};
		for(var k in fields) base[k] = fields[k];
		$tw.wiki.addTiddler(new $tw.Tiddler(base));
		added.push(title);
	}

	beforeEach(function() { added = []; });
	afterEach(function() {
		for(var i = 0; i < added.length; i++) $tw.wiki.deleteTiddler(added[i]);
	});

	function evalCascade(title) {
		// Cascade is a filter expression with `then[...]`; with currentTiddler
		// set to the target tiddler, evaluating it should yield the body
		// template title (or nothing, if the tiddler doesn't match).
		return $tw.wiki.filterTiddlers(CASCADE_FILTER.trim(), {
			getVariable: function(name) {
				if(name === "currentTiddler") return title;
				return undefined;
			}
		});
	}

	it("ships a non-empty cascade filter that mentions body/email", function () {
		expect(CASCADE_FILTER.length).toBeGreaterThan(0);
		expect(CASCADE_FILTER.indexOf("body/email")).toBeGreaterThan(-1);
	});

	it("dispatches a frontmattered-markdown tiddler with msg-subject to body/email", function () {
		add("$:/test/mi/email-A", {
			type: "text/x-frontmattered-markdown",
			"msg-subject": "Project kickoff",
			"msg-from": "alice@example.com",
			text: "# Project kickoff"
		});
		expect(evalCascade("$:/test/mi/email-A")).toEqual([EMAIL_BODY]);
	});

	it("does NOT dispatch a frontmattered-markdown tiddler without msg-subject", function () {
		// A generic markdown-with-frontmatter tiddler (not an imported email)
		// must fall through to whatever later cascade entries pick it up.
		add("$:/test/mi/generic-doc", {
			type: "text/x-frontmattered-markdown",
			caption: "Just some notes",
			text: "# Hello"
		});
		expect(evalCascade("$:/test/mi/generic-doc")).toEqual([]);
	});

	it("does NOT dispatch tiddlers of other types even if they happen to have msg-subject", function () {
		add("$:/test/mi/weird", {
			type: "text/vnd.tiddlywiki",
			"msg-subject": "spoof",
			text: ""
		});
		expect(evalCascade("$:/test/mi/weird")).toEqual([]);
	});
});

describe("msg-import: cascade/eml dispatches the .eml parent to body/msg", function () {

	var CASCADE_FILTER = $tw.wiki.getTiddlerText("$:/plugins/rimir/msg-import/cascade/eml") || "";
	var MSG_BODY = "$:/plugins/rimir/msg-import/body/msg";

	var added = [];

	function add(title, fields) {
		var base = {title: title};
		for(var k in fields) base[k] = fields[k];
		$tw.wiki.addTiddler(new $tw.Tiddler(base));
		added.push(title);
	}

	beforeEach(function() { added = []; });
	afterEach(function() {
		for(var i = 0; i < added.length; i++) $tw.wiki.deleteTiddler(added[i]);
	});

	function evalCascade(title) {
		return $tw.wiki.filterTiddlers(CASCADE_FILTER.trim(), {
			getVariable: function(name) {
				if(name === "currentTiddler") return title;
				return undefined;
			}
		});
	}

	it("ships a non-empty cascade filter that mentions body/msg", function () {
		expect(CASCADE_FILTER.length).toBeGreaterThan(0);
		expect(CASCADE_FILTER.indexOf("body/msg")).toBeGreaterThan(-1);
	});

	it("dispatches a message/rfc822 tiddler to body/msg", function () {
		add("$:/test/mi/parent.eml", {
			type: "message/rfc822",
			_canonical_uri: "/files/email/parent.eml"
		});
		expect(evalCascade("$:/test/mi/parent.eml")).toEqual([MSG_BODY]);
	});

	it("does NOT dispatch unrelated MIME types", function () {
		add("$:/test/mi/parent.txt", {type: "text/plain"});
		expect(evalCascade("$:/test/mi/parent.txt")).toEqual([]);
	});

	it("does NOT dispatch application/vnd.ms-outlook (handled by cascade/msg)", function () {
		add("$:/test/mi/parent.msg", {type: "application/vnd.ms-outlook"});
		expect(evalCascade("$:/test/mi/parent.msg")).toEqual([]);
	});
});

describe("msg-import: body/email attachment filter", function () {

	var EMAIL = "$:/test/mi/eml-" + Date.now() + ".msg.email";
	var MSG = "$:/test/mi/eml-" + Date.now() + ".msg";
	var added = [];

	function add(title, fields) {
		var base = {title: title};
		for(var k in fields) base[k] = fields[k];
		$tw.wiki.addTiddler(new $tw.Tiddler(base));
		added.push(title);
	}

	beforeEach(function() {
		added = [];
		add(MSG, {type: "application/vnd.ms-outlook"});
		add(EMAIL, {
			type: "text/x-frontmattered-markdown",
			"msg-subject": "test",
			"_artifact_source": MSG,
			"_artifact_type": "conversion"
		});
		add(MSG + ".attachments/att_a.pdf", {
			_artifact_source: MSG,
			_artifact_type: "attachment",
			type: "application/pdf"
		});
		add(MSG + ".attachments/att_b.jpg", {
			_artifact_source: MSG,
			_artifact_type: "attachment",
			type: "image/jpeg"
		});
		// Non-attachment artifacts of the same parent — must NOT be listed.
		add(MSG + ".summary", {
			_artifact_source: MSG,
			_artifact_type: "summary",
			type: "text/x-markdown"
		});
	});

	afterEach(function() {
		for(var i = 0; i < added.length; i++) $tw.wiki.deleteTiddler(added[i]);
	});

	it("starting from the .email, finds the parent's attachments via _artifact_source", function () {
		// This is the literal filter inside body/email.tid, evaluated against
		// the .email's _artifact_source (which points at the .msg).
		var msgParent = $tw.wiki.getTiddler(EMAIL).fields._artifact_source;
		var results = $tw.wiki.filterTiddlers(
			"[_artifact_source[" + msgParent + "]_artifact_type[attachment]]"
		);
		expect(results.sort()).toEqual([
			MSG + ".attachments/att_a.pdf",
			MSG + ".attachments/att_b.jpg"
		]);
	});

	it("excludes the .summary and .conversion siblings", function () {
		var msgParent = $tw.wiki.getTiddler(EMAIL).fields._artifact_source;
		var results = $tw.wiki.filterTiddlers(
			"[_artifact_source[" + msgParent + "]_artifact_type[attachment]]"
		);
		expect(results.indexOf(MSG + ".summary")).toBe(-1);
		expect(results.indexOf(EMAIL)).toBe(-1);
	});
});
