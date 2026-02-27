#!/usr/bin/env node
//
// annotate.mjs — Zero-dependency CLI wrapper for the annotation service.
//
// Usage:
//   node annotate.mjs <input.docx> <output.docx> '<json>'
//   node annotate.mjs <input.docx> <output.docx> --json <file.json>
//   node annotate.mjs <input.docx> <output.docx> '<json>' --url http://host:port
//
// The JSON is an array of annotations:
//   [
//     { "type": "Comment",  "find": "some text", "comment": "my note" },
//     { "type": "Replace",  "find": "old text",  "replacement": "new text" },
//     { "type": "Delete",   "find": "remove this" },
//     { "type": "Insert",   "find": "anchor text", "text": "new clause", "position": "after" }
//   ]
//
// Outputs structured JSON to stdout:
//   { "ok": true,  "output": "/path/to/out.docx", "errors": [] }
//   { "ok": false, "error": "..." }
//

import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { request } from "node:http";

// ── Parse args ──────────────────────────────────────────────

const args = process.argv.slice(2);

function bail(msg) {
  console.log(JSON.stringify({ ok: false, error: msg }));
  process.exit(1);
}

let inputPath, outputPath, annotationsJson;
let serviceUrl = "https://comments-service.edumagick.com/";
let author = "AI Assistant";
let initials = "AI";

const positional = [];
for (let i = 0; i < args.length; i++) {
  switch (args[i]) {
    case "--json":
      annotationsJson = readFileSync(args[++i], "utf-8");
      break;
    case "--url":
      serviceUrl = args[++i];
      break;
    case "--author":
      author = args[++i];
      break;
    case "--initials":
      initials = args[++i];
      break;
    case "--help":
    case "-h":
      console.log(
        [
          "Usage: node annotate.mjs <input.docx> <output.docx> '<json>' [options]",
          "",
          "Arguments:",
          "  input.docx       Source Word document",
          "  output.docx      Where to write the annotated document",
          "  <json>           Annotations array as inline JSON string",
          "",
          "Options:",
          '  --json <file>    Read annotations from a JSON file instead of inline',
          "  --url <url>      Service URL (default: $ANNOTATE_SERVICE_URL or http://localhost:8080)",
          "  --author <name>  Author name (default: AI Assistant)",
          '  --initials <i>   Author initials (default: AI)',
          "",
          "Annotation types:",
          '  { "type": "Comment",  "find": "...", "comment": "..." }',
          '  { "type": "Replace",  "find": "...", "replacement": "..." }',
          '  { "type": "Delete",   "find": "..." }',
          '  { "type": "Insert",   "find": "...", "text": "...", "position": "after"|"before" }',
        ].join("\n")
      );
      process.exit(0);
      break;
    default:
      positional.push(args[i]);
  }
}

inputPath = positional[0];
outputPath = positional[1];
if (!annotationsJson && positional[2]) annotationsJson = positional[2];

if (!inputPath || !outputPath || !annotationsJson) {
  bail(
    "Usage: node annotate.mjs <input.docx> <output.docx> '<json>' [--url URL]"
  );
}

// Validate the JSON early so the agent gets a clear error
let annotations;
try {
  annotations = JSON.parse(annotationsJson);
  if (!Array.isArray(annotations))
    throw new Error("Expected a JSON array of annotations");
} catch (e) {
  bail(`Invalid annotations JSON: ${e.message}`);
}

// ── Build multipart/form-data ───────────────────────────────

const boundary = `----NodeFormBoundary${Date.now()}${Math.random().toString(36).slice(2)}`;
const CRLF = "\r\n";

function multipartField(name, value) {
  return (
    `--${boundary}${CRLF}` +
    `Content-Disposition: form-data; name="${name}"${CRLF}${CRLF}` +
    value +
    CRLF
  );
}

function multipartFile(name, filename, buf) {
  const header =
    `--${boundary}${CRLF}` +
    `Content-Disposition: form-data; name="${name}"; filename="${filename}"${CRLF}` +
    `Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document${CRLF}${CRLF}`;
  return Buffer.concat([
    Buffer.from(header),
    buf,
    Buffer.from(CRLF),
  ]);
}

let fileBuffer;
try {
  fileBuffer = readFileSync(resolve(inputPath));
} catch (e) {
  bail(`Cannot read input file: ${e.message}`);
}

const requestBody = JSON.stringify({ author, initials, annotations });

const parts = Buffer.concat([
  multipartFile("file", inputPath.split("/").pop(), fileBuffer),
  Buffer.from(multipartField("request", requestBody)),
  Buffer.from(`--${boundary}--${CRLF}`),
]);

// ── Send request ────────────────────────────────────────────

const url = new URL("/api/annotate", serviceUrl);

const req = request(
  {
    hostname: url.hostname,
    port: url.port,
    path: url.pathname,
    method: "POST",
    headers: {
      "Content-Type": `multipart/form-data; boundary=${boundary}`,
      "Content-Length": parts.length,
      Accept: "application/json",
    },
  },
  (res) => {
    const chunks = [];
    res.on("data", (c) => chunks.push(c));
    res.on("end", () => {
      const body = Buffer.concat(chunks);

      // Non-200 → the body is a JSON error
      if (res.statusCode !== 200) {
        try {
          const err = JSON.parse(body.toString());
          console.log(
            JSON.stringify({ ok: false, error: err.error || body.toString() })
          );
        } catch {
          console.log(
            JSON.stringify({ ok: false, error: body.toString().slice(0, 500) })
          );
        }
        process.exit(1);
      }

      // 200 → JSON envelope with base64 file + errors
      try {
        const envelope = JSON.parse(body.toString());
        const outPath = resolve(outputPath);
        writeFileSync(outPath, Buffer.from(envelope.file, "base64"));

        const result = { ok: true, output: outPath };
        if (envelope.errors && envelope.errors.length > 0)
          result.errors = envelope.errors;

        console.log(JSON.stringify(result, null, 2));
        process.exit(envelope.errors?.length ? 2 : 0);
      } catch (e) {
        bail(`Unexpected response from service: ${e.message}`);
      }
    });
  }
);

req.on("error", (e) => {
  bail(`Cannot connect to annotation service at ${serviceUrl}: ${e.message}`);
});

req.write(parts);
req.end();
