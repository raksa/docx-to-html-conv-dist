# docx-to-html

Convert Word (`.docx`) documents into paginated, self-contained HTML files —
one HTML file per page — with embedded fonts, extracted images, and an
`info.json` describing the document.

Node.js port of the C# `DocxToHtml` tool; the output is byte-for-byte identical
to the reference implementation (apart from the `exportedAt` timestamp).

## Install

```bash
npm install docx-to-html
```

Requires Node.js >= 18.

## Usage

ESM:

```js
import { exportDocx, getPagesCount } from 'docx-to-html';

exportDocx('document.docx', 'out/document');
console.log(getPagesCount('document.docx'));
```

CommonJS:

```js
const { exportDocx, getPagesCount } = require('docx-to-html');

exportDocx('document.docx', 'out/document');
```

TypeScript (declarations ship with the package — nothing extra to install):

```ts
import { exportDocx, getPagesCount, getVersion } from 'docx-to-html';

const pages: number = getPagesCount('document.docx');
exportDocx('document.docx', 'out/document');
```

## CLI

```bash
npx docx-to-html --file <input.docx> [--out <output-folder>]
```

- `--file` — path to the `.docx` file (required)
- `--out` — destination folder (default: `out/<document-name>`)

## API

| Export | Description |
| --- | --- |
| `exportDocx(docxPath, outputDirectory)` | Convert a document; creates the output folder if missing. |
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
