# Mammoth: semantic reading, not visual rendering

This skill uses the installed **Node/npm Mammoth**. It maps Word styles to semantic HTML, supports text extraction, and includes images inline by default. It intentionally ignores many layout details, including table borders. It cannot edit DOCX or generate DOCX.

```powershell
node scripts/mammoth_extract.cjs input.docx semantic-preview.html
node scripts/mammoth_extract.cjs input.docx content.txt --text
```

The wrapper resolves modules from the skill script's location or Windows global npm installations, returns conversion messages as JSON, and refuses existing output paths. For a project-local package, set `MAMMOTH_MODULE` to its absolute package directory; changing the working directory does not change Node's module search base. HTML is a full UTF-8 document with a restrictive CSP; this does **not** make it sanitized. It is a semantic reading artifact, not evidence of pagination or fidelity.

## Library use for custom extraction

```javascript
const mammoth = require(process.env.MAMMOTH_MODULE || 'mammoth');
const result = await mammoth.convertToHtml({path: 'input.docx'}, {
  externalFileAccess: false,
  includeEmbeddedStyleMap: false,
  styleMap: [
    "p[style-name='Section Title'] => h1:fresh",
    "p[style-name='Aside Text'] => p.aside:fresh"
  ]
});
console.log(result.messages);
```

If `require('mammoth')` fails despite `npm list -g` showing it, see [Windows module discovery](windows.md). A globally installed CLI and a resolvable local Node module are different things.

`extractRawText({path})` ignores formatting and separates paragraphs with two newlines. It is useful for search/review but not an exhaustive OOXML audit. It should not be the only inventory when headers, comments, revision history, or hidden/alternate content matters.

`convertToHtml` supports images, headings, lists, links, tables, notes, and style maps. Text-box content can move to a separate paragraph after its anchor. Comments require deliberate mapping/inspection depending on the requested representation. Retain `result.messages` and review unsupported-content warnings.

## Untrusted documents

Mammoth explicitly does not sanitize source documents. HTML links may contain unsafe schemes, and file references need special care. Keep `externalFileAccess: false` unless a trusted workflow needs a specific external resource. The wrapper ignores embedded style maps for predictable extraction.

Before embedding output in an application, apply the application's established HTML sanitizer with URL policy. A CSP or string replacement is not a sanitizer. Do not automatically click extracted links or upload converted document contents.

For Markdown, use HTML extraction followed by a converter such as Pandoc when that semantic transformation is wanted. Mammoth's own Markdown output is deprecated. Do not claim Mammoth -> HTML -> Pandoc -> DOCX is a lossless editing path.

Source: [official npm implementation README](https://github.com/mwilliamson/mammoth.js), including CLI, Library, Style maps, and Security; also checked against the installed 1.12.3 README.
