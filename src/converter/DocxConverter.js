'use strict';

// Port of the C# `static partial class DocxConverter`. The C# version keeps
// per-run state in static fields; here each conversion uses one instance whose
// prototype is composed from the ported method modules (mirroring the partial
// class split across files).

const { CIMap, CISet } = require('../util/collections');

class DocxConverter {
  constructor() {
    // Output directory for the current run.
    this._outputDir = '';
    // Unique image filename counter.
    this._imageCounter = 0;
    // When true, rendering runs purely in memory (no image/font/page files written).
    this._skipAssetWrites = false;
    // Deduplication cache: part uri -> saved relative path.
    this._savedImagePaths = new Map();

    // All fonts referenced across the document (sorted, case-insensitive).
    this._allFonts = new CISet();
    // Embedded font data extracted from the DOCX FontTablePart (name -> bytes).
    this._embeddedDocxFonts = new CIMap();
    // Cache of @font-face CSS rules (font name -> css string or null).
    this._fontFaceCache = new CIMap();
    // Fonts that could not be found anywhere.
    this._missingFonts = new CISet();
    // Family names of fonts embedded in the DOCX.
    this._embeddedFontFamilyNames = new CISet();

    // Theme fonts (defaults match the C# converter's fallbacks).
    this._themeBodyFont = 'Cambria';
    this._themeHeadingFont = 'Calibri Light';

    // Numbering / styles parts and a style-by-id cache.
    this._numberingPart = null;
    this._stylesPart = null;
    this._styleCache = new CIMap();
  }
}

// Compose the prototype from the ported method modules (order-independent).
Object.assign(
  DocxConverter.prototype,
  require('./run'),
  require('./sections'),
  require('./pageBoundaries'),
  require('./paragraphs'),
  require('./runs'),
  require('./inlineElements'),
  require('./lists'),
  require('./tables'),
  require('./headerFooter'),
  require('./pageHtml'),
  require('./utilities'),
  require('./fontHelper')
);

module.exports = { DocxConverter };
