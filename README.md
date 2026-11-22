# Docx to HTML Converter

Convert Word (`.docx`) documents into paginated, self-contained HTML files —
one HTML file per page — with embedded fonts, extracted images, and an
`info.json` describing the document.

Node.js port of the C# `DocxToHtml` tool; the output is byte-for-byte identical
to the reference implementation (apart from the `exportedAt` timestamp).

## Install

```bash
npm install docx-to-html-conv
```

Requires Node.js >= 18.

### Releasing to npm

For maintainers, from a clean `main`:

```bash
npm test
npm pack --dry-run                          # check which files go in the package
npm whoami                                  # run `npm login` if this fails
npm version patch --tag-version-prefix=""   # or minor / major
npm publish
git push origin main --follow-tags
npm view docx-to-html-conv version          # confirm the new version is live
```

## Usage

ESM:

```js
import { exportDocx, getPagesCount } from 'docx-to-html-conv';

exportDocx('document.docx', 'out/document');
console.log(getPagesCount('document.docx'));
```

CommonJS:

```js
const { exportDocx, getPagesCount } = require('docx-to-html-conv');

exportDocx('document.docx', 'out/document');
```

TypeScript (declarations ship with the package — nothing extra to install):

```ts
import { exportDocx, getPagesCount, getVersion } from 'docx-to-html-conv';

const pages: number = getPagesCount('document.docx');
exportDocx('document.docx', 'out/document');
```

## CLI

```bash
npx docx-to-html-conv --file <input.docx> [--out <output-folder>] [--default-font-family <family>]
```

- `--file` — path to the `.docx` file (required)
- `--out` — destination folder (default: `out/<document-name>`)
- `--default-font-family` — the font family (or comma-separated families) of
  the page the document is shown in (optional). Text in a face the computer
  lacks is drawn in it rather than in the browser's generic font.

## API

| Export | Description |
| --- | --- |
| `exportDocx(docxPath, outputDirectory, defaultFontFamily?)` | Convert a document; creates the output folder if missing. `defaultFontFamily` is the host page's font, used for any face the computer lacks. |
| `getPagesCount(docxPath)` | Number of pages the conversion would produce. |
| `getVersion()` | Tool version, as written into `info.json`. |
| `DocxConverter` | The converter class, for driving a run directly. |

Pages are split on explicit page breaks, Word's rendered page-break hints
(`w:lastRenderedPageBreak`), and section boundaries; each page carries its
section's page size, margins, columns, and header/footer.

Supported: heading/quote styles, alignment, spacing and indents, paragraph
borders and shading, full run formatting (bold/italic/underline styles/strike/
size/color/font/highlight/caps/super-subscript), nested lists from
`numbering.xml`, tables with `colspan`/`rowspan`/shading, `PAGE`/`NUMPAGES`
fields, inline and anchored images, external and internal hyperlinks,
default/first/even headers and footers, and embedded + system font extraction
with `@font-face` embedding.

## License

MIT
