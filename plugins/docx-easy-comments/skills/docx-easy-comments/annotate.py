#!/usr/bin/env python3
"""
annotate.py — CLI wrapper for the annotation service.

Usage:
  python3 annotate.py <input.docx> <output.docx> '<json>'
  python3 annotate.py <input.docx> <output.docx> --json <file.json>
  python3 annotate.py <input.docx> <output.docx> '<json>' --url http://host:port

The JSON is an array of annotations:
  [
    { "type": "Comment",  "find": "some text", "comment": "my note" },
    { "type": "Replace",  "find": "old text",  "replacement": "new text" },
    { "type": "Delete",   "find": "remove this" },
    { "type": "Insert",   "find": "anchor text", "text": "new clause", "position": "after" }
  ]

Outputs structured JSON to stdout:
  { "ok": true,  "output": "/path/to/out.docx" }
  { "ok": false, "error": "...", "errors": [...] }

Requires: requests  (pip install requests)
"""

import json
import os
import sys

try:
    import requests
except ImportError:
    print(json.dumps({"ok": False, "error": "Missing dependency: pip install requests"}))
    sys.exit(1)


def bail(msg):
    print(json.dumps({"ok": False, "error": msg}))
    sys.exit(1)


# ── Parse args ───────────────────────────────────────────────

args = sys.argv[1:]

input_path = None
output_path = None
annotations_json = None
service_url = "https://comments-service.edumagick.com/"
author = "AI Assistant"
initials = "AI"

positional = []
i = 0
while i < len(args):
    arg = args[i]
    if arg == "--json":
        i += 1
        try:
            with open(args[i], "r", encoding="utf-8") as f:
                annotations_json = f.read()
        except (IndexError, FileNotFoundError, IOError) as e:
            bail(f"Cannot read JSON file: {e}")
    elif arg == "--url":
        i += 1
        try:
            service_url = args[i]
        except IndexError:
            bail("Missing value for --url")
    elif arg == "--author":
        i += 1
        try:
            author = args[i]
        except IndexError:
            bail("Missing value for --author")
    elif arg == "--initials":
        i += 1
        try:
            initials = args[i]
        except IndexError:
            bail("Missing value for --initials")
    elif arg in ("--help", "-h"):
        print(
            "Usage: python3 annotate.py <input.docx> <output.docx> '<json>' [options]\n"
            "\n"
            "Arguments:\n"
            "  input.docx       Source Word document\n"
            "  output.docx      Where to write the annotated document\n"
            "  <json>           Annotations array as inline JSON string\n"
            "\n"
            "Options:\n"
            "  --json <file>    Read annotations from a JSON file instead of inline\n"
            "  --url <url>      Service URL (default: https://comments-service.edumagick.com/)\n"
            "  --author <name>  Author name (default: AI Assistant)\n"
            "  --initials <i>   Author initials (default: AI)\n"
            "\n"
            "Annotation types:\n"
            '  { "type": "Comment",  "find": "...", "comment": "..." }\n'
            '  { "type": "Replace",  "find": "...", "replacement": "..." }\n'
            '  { "type": "Delete",   "find": "..." }\n'
            '  { "type": "Insert",   "find": "...", "text": "...", "position": "after"|"before" }'
        )
        sys.exit(0)
    else:
        positional.append(arg)
    i += 1

input_path = positional[0] if len(positional) > 0 else None
output_path = positional[1] if len(positional) > 1 else None
if annotations_json is None and len(positional) > 2:
    annotations_json = positional[2]

if not input_path or not output_path or not annotations_json:
    bail("Usage: python3 annotate.py <input.docx> <output.docx> '<json>' [--json FILE] [--url URL] [--author NAME] [--initials XX]")

# Validate the JSON early so the agent gets a clear error
try:
    annotations = json.loads(annotations_json)
    if not isinstance(annotations, list):
        raise ValueError("Expected a JSON array of annotations")
except (json.JSONDecodeError, ValueError) as e:
    bail(f"Invalid annotations JSON: {e}")

# ── Read input file ──────────────────────────────────────────

input_abs = os.path.abspath(input_path)
try:
    with open(input_abs, "rb") as f:
        file_bytes = f.read()
except (FileNotFoundError, IOError) as e:
    bail(f"Cannot read input file: {e}")

# ── Build and send request ───────────────────────────────────

url = service_url.rstrip("/") + "/api/annotate"
request_body = json.dumps({"author": author, "initials": initials, "annotations": annotations})

try:
    resp = requests.post(
        url,
        files={
            "file": (os.path.basename(input_path), file_bytes,
                     "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        },
        data={"request": request_body},
        headers={"Accept": "application/json"},
        timeout=120,
    )
except requests.RequestException as e:
    bail(f"Cannot connect to annotation service at {service_url}: {e}")

# ── Handle response ──────────────────────────────────────────

if resp.status_code != 200:
    try:
        err = resp.json()
    except (json.JSONDecodeError, ValueError):
        bail(resp.text[:500])

    # API contract: 422 means one or more annotations could not be resolved/applied.
    # No output file is returned in this case.
    if resp.status_code == 422 and isinstance(err, dict) and isinstance(err.get("errors"), list):
        print(json.dumps({
            "ok": False,
            "error": "One or more annotations could not be applied.",
            "errors": err["errors"],
        }, indent=2))
        sys.exit(2)

    bail(err.get("error", resp.text[:500]))

try:
    import base64

    envelope = resp.json()
    out_path = os.path.abspath(output_path)
    with open(out_path, "wb") as f:
        f.write(base64.b64decode(envelope["file"]))

    print(json.dumps({"ok": True, "output": out_path}, indent=2))
    sys.exit(0)

except (KeyError, json.JSONDecodeError, ValueError, IOError) as e:
    bail(f"Unexpected response from service: {e}")

