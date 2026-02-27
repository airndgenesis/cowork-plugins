---
description: Annotate a Word (.docx) document with comments, tracked replacements, deletions, and insertions
argument-hint: <path-to-document.docx>
allowed-tools: Read, Write, Bash
---

# Annotate Word Document

You have access to a tool that adds comments, tracked replacements, deletions, and insertions to Word (.docx) documents. The tool is a Python CLI at:

```bash
${SKILL_DIR}/annotate.py
```

## When to use this

Use this skill when the user wants to:
- Add review comments to a Word document
- Suggest tracked changes (replace, delete, or insert text)
- Annotate a .docx file with feedback, edits, or notes

## How to use the tool

```bash
python3 "${SKILL_DIR}/annotate.py" <input.docx> <output.docx> --json <annotations.json>
```

Or with inline JSON:

```bash
python3 "${SKILL_DIR}/annotate.py" <input.docx> <output.docx> '<annotations_json>'
```

### Options

| Flag               | Default                                          | Description                          |
|--------------------|--------------------------------------------------|--------------------------------------|
| `--url <url>`      | `https://comments-service.edumagick.com/`        | Annotation service URL               |
| `--author <name>`  | `AI Assistant`                                   | Author shown in Word                 |
| `--initials <str>` | `AI`                                             | Initials shown in comment bubbles    |
| `--json <file>`    | —                                                | Read annotations from a file instead of inline |

## Annotation JSON format

The annotations JSON is an **array** of objects.

### Preferred targeting: `find`

```json
{
  "type": "Comment",
  "find": "The Seller shall indemnify the Buyer against all losses",
  "comment": "**Issue:** This indemnification is *overly broad*."
}
```

You can disambiguate repeated matches with:

```json
{ "occurrence": 2 }
```

### Fallback targeting: paragraph/sentence indexes

If needed, you can target by position instead of text:

```json
{
  "type": "Comment",
  "paragraphIndex": 12,
  "sentenceIndex": 0,
  "comment": "Please verify this statement."
}
```

- `paragraphIndex` is zero-based across the document.
- `sentenceIndex` is zero-based within the paragraph.
- Omit `sentenceIndex` (or set `-1`) to target the whole paragraph.

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

## Text matching behavior (`find`)

The service resolves `find` using multiple passes:

1. Exact match (ordinal)
2. Normalized match (whitespace, smart quotes, dash variants, invisible characters)
3. Dehyphenated match (handles PDF-to-Word artifacts like `con-cepts`)
4. Case-insensitive dehyphenated match
5. Cross-paragraph search (windowed, including hyphenated paragraph boundary joins)

If no match is found, the service may return `closestMatches` hints.

## Important guidelines

1. **Prefer writing annotations JSON to a temp file** and using `--json <file>` to avoid shell quoting issues.
2. **Use long, unique `find` strings** to avoid matching the wrong place.
3. **Copy `find` text exactly** from the document whenever possible.
4. **Check tool output JSON** and handle errors explicitly.

## Output format

The CLI always outputs JSON to stdout.

**Success** (exit code 0):
```json
{ "ok": true, "output": "/absolute/path/to/output.docx" }
```

**Annotation validation/apply failure from API** (exit code 2, HTTP 422 upstream):
```json
{
  "ok": false,
  "error": "One or more annotations could not be applied.",
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

**Fatal error** (exit code 1):
```json
{ "ok": false, "error": "Cannot connect to annotation service..." }
```

## Error handling

- If you get `errors` with `closestMatches`, adjust `find` text and retry.
- If the service is unreachable, inform the user.
- The API applies annotations atomically for resolve failures: if any target cannot be resolved, no annotated output file is returned.

## Reading document content with pandoc

To inspect a .docx before annotating, use **pandoc** to convert it to markdown:

```bash
pandoc "input.docx" -t markdown -o /tmp/input_preview.md
```

Then read the markdown file to craft accurate `find` strings.

> If `pandoc` is not installed, use another extraction approach.

## Workflow

1. Ask the user which .docx file to annotate (or use the one provided).
2. Convert the .docx to markdown with pandoc and inspect content.
3. Build annotations, preferring `find` targeting and exact copied snippets.
4. Write annotations to a temporary JSON file.
5. Run the tool with `--json`.
6. Report results; if failures occur, fix and retry.
7. Tell the user where the output file is.

