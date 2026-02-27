# docx-easy-comments

A [Claude Code](https://claude.ai/code) plugin marketplace for annotating Word (.docx) documents with comments, tracked changes, and insertions.

## Installation

In Claude Code, add this marketplace by URL:

```
https://github.com/<owner>/docs_comments_skill.git
```

Then install the plugin:

```
/plugin install docx-easy-comments@docx-easy-comments
```

## What it does

Adds an `/annotate-docx` slash command that lets Claude:

- **Comment** — add review comments anchored to specific text
- **Replace** — suggest tracked replacements
- **Delete** — suggest tracked deletions
- **Insert** — insert new text before or after existing passages

All annotations support **inline markdown** (`**bold**`, `*italic*`, `~~strikethrough~~`) that renders natively in Word.

## How it works

```
┌──────────────┐        HTTP multipart        ┌──────────────────┐
│ annotate.mjs │  ──── POST /api/annotate ──► │  C# service      │
│ (CLI wrapper) │ ◄── JSON { ok, file, errors} │  (Open XML SDK)  │
└──────────────┘                              └──────────────────┘
```

The plugin ships a zero-dependency Node.js CLI (`annotate.mjs`) that sends the document and annotations to a hosted annotation service. The service handles all Open XML manipulation and returns the annotated document.

## Requirements

- **Node.js** ≥ 18 (uses built-in `node:http`, `node:fs`, `node:path`)
- Internet access to reach the annotation service at `https://comments-service.edumagick.com/`

## Example

Ask Claude:

> Review this contract and add comments about any concerning clauses: `./contract.docx`

Claude will read the document, craft annotations, and produce an annotated copy with Word-native comments and tracked changes.

## License

MIT
