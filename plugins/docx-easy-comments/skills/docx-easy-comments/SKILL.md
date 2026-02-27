---
description: Annotate a Word (.docx) document with comments, tracked replacements, deletions, and insertions
argument-hint: <path-to-document.docx>
allowed-tools: Read, Write, Bash
---

# Annotate Word Document

You have access to a tool that adds comments, tracked replacements, deletions, and insertions to Word (.docx) documents. The tool is a Node.js CLI at:

```
${CLAUDE_PLUGIN_ROOT}/annotate.mjs
```

## When to use this

Use this skill when the user wants to:
- Add review comments to a Word document
- Suggest tracked changes (replace, delete, or insert text)
- Annotate a .docx file with feedback, edits, or notes

## How to use the tool

```bash
node "${CLAUDE_PLUGIN_ROOT}/annotate.mjs" <input.docx> <output.docx> --json <annotations.json>
```

Or with inline JSON:

```bash
node "${CLAUDE_PLUGIN_ROOT}/annotate.mjs" <input.docx> <output.docx> '<annotations_json>'
```

### Options

| Flag               | Default                                          | Description                          |
|--------------------|--------------------------------------------------|--------------------------------------|
| `--url <url>`      | `https://comments-service.edumagick.com/`        | Annotation service URL               |
| `--author <name>`  | `AI Assistant`                                   | Author shown in Word                 |
| `--initials <str>` | `AI`                                             | Initials shown in comment bubbles    |
| `--json <file>`    | —                                                | Read annotations from a file instead of inline |

## Annotation JSON format

The annotations JSON is an **array** of objects. Each entry targets text using the `find` field.

### Comment

```json
{
  "type": "Comment",
  "find": "The Seller shall indemnify the Buyer against all losses",
  "comment": "**Issue:** This indemnification is *overly broad*.\n\n**Recommendation:** Limit to ~~all losses~~ ***direct*** losses only."
}
```

### Replace (tracked change)

```json
{
  "type": "Replace",
  "find": "best efforts",
  "replacement": "**commercially reasonable** efforts"
}
```

### Delete (tracked change)

```json
{
  "type": "Delete",
  "find": "including but not limited to consequential damages"
}
```

### Insert (tracked change)

```json
{
  "type": "Insert",
  "find": "terminate on December 31, 2026.",
  "text": " Either party may extend by providing *30 days* written notice.",
  "position": "after"
}
```

`position` is `"after"` (default) or `"before"`.

## Markdown formatting in annotation text

Comment bodies, replacement text, and insertion text support inline markdown that renders natively in Word:

| Markdown             | Renders as         |
|----------------------|--------------------|
| `**bold**`           | **bold**           |
| `*italic*`           | *italic*           |
| `***bold italic***`  | ***bold italic***  |
| `~~strikethrough~~`  | ~~strikethrough~~  |

Newlines (`\n`) in comment text create separate paragraphs in the comment bubble.

## Text matching

The `find` field locates text using a three-pass search:

1. **Exact match** — literal string comparison
2. **Normalized match** — collapses whitespace, smart quotes → ASCII, dashes normalized
3. **Dehyphenated match** — strips hyphens between lowercase letters (handles PDF-to-Word artifacts)

Use `"occurrence": 2` to target the second occurrence (1-based, defaults to 1).

## Important guidelines

1. **Always write annotations JSON to a temp file** and use `--json <file>` rather than inline JSON to avoid shell quoting issues.
2. **Use long, unique `find` strings** — copy enough surrounding text to uniquely identify the target passage. Short strings may match in the wrong place.
3. **Copy `find` text exactly** from the document. Do not paraphrase or rephrase — the matching is literal.
4. **Check the output carefully.** The tool returns structured JSON to stdout.

## Output format

The tool always outputs JSON to stdout.

**Success** (exit code 0):
```json
{ "ok": true, "output": "/absolute/path/to/output.docx" }
```

**Partial success** — file written but some annotations failed (exit code 2):
```json
{
  "ok": true,
  "output": "/absolute/path/to/output.docx",
  "errors": [
    {
      "index": 1,
      "message": "Text not found in document",
      "find": "shall use best efforts",
      "closestMatches": [
        {
          "paragraph": 52,
          "text": "The Receiving Party shall use commercially reasonable efforts to...",
          "matchedFragment": "shall use"
        }
      ]
    }
  ]
}
```

**Fatal error** — no file written (exit code 1):
```json
{ "ok": false, "error": "Cannot connect to annotation service..." }
```

## Error handling

- If you get partial errors with `closestMatches`, use those hints to fix the `find` text and retry only the failed annotations.
- If the service is unreachable, inform the user.
- Exit code 2 means the output file was still written — report both the success and the failures to the user.

## Reading document content with pandoc

To read the text of a .docx file before annotating, you are recommented to use **pandoc** to convert it to markdown:

```bash
pandoc "input.docx" -t markdown -o /tmp/input_preview.md
```

Then read the resulting markdown file to understand the document's structure and content.

**Why this matters:**
- You need to see the actual text to craft accurate `find` strings that match exactly.
- Markdown output preserves headings, lists, tables, and emphasis — giving you a faithful view of the document's structure.
- This avoids blind guessing and dramatically reduces partial-match errors.

> **Note:** If `pandoc` is not installed, try something else

## Workflow

1. Ask the user which .docx file to annotate (or use the one they provided).
2. **Convert the .docx to markdown with pandoc** and read the result to understand the document content.
3. Build the annotations array based on the user's request, copying `find` strings verbatim from the pandoc output.
4. Write the annotations to a temporary JSON file.
5. Run the tool with `--json`.
6. Report results. If there are partial errors, offer to retry the failed ones.
7. Tell the user where the output file is.
