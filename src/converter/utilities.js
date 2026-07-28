'use strict';

// Port of DocxToHtml/docx/DocxConverter.Utilities.cs — misc helpers that touch
// per-run state (page-break detection, style lookup, image saving, info.json).

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { firstChild, elements, childElements, attr, innerText, NS } = require('../ooxml/dom');
const { serializeIndented } = require('../util/dotnetJson');

module.exports = {
  isPageBreakOnly(para) {
    let foundBreak = false;
    for (const run of elements(para, 'r', NS.W)) {
      for (const child of childElements(run)) {
        if (child.namespaceURI !== NS.W) continue;
        if (child.localName === 'rPr') continue;
        if (child.localName === 'br' && attr(child, 'type') === 'page') {
          foundBreak = true;
          continue;
        }
        if (child.localName === 't' && innerText(child).trim() !== '') return false;
        if (child.localName === 'drawing') return false;
      }
    }
    if (elements(para, 'hyperlink', NS.W).length > 0) return false;
    return foundBreak;
  },

  getStyleById(styleId) {
    if (this._styleCache.has(styleId)) return this._styleCache.get(styleId);
    const stylesRoot = this._stylesPart != null ? this._stylesPart.root : null;
    let style = null;
    if (stylesRoot != null) {
      style = elements(stylesRoot, 'style', NS.W).find((s) => attr(s, 'styleId') === styleId) || null;
    }
    if (style != null) this._styleCache.set(styleId, style);
    return style;
  },

  saveImageFile(part) {
    const partUri = part.uri;
    if (this._savedImagePaths.has(partUri)) return this._savedImagePaths.get(partUri);

    const mime = part.contentType || 'image/png';
    let ext;
    switch (mime) {
      case 'image/jpeg': ext = '.jpg'; break;
      case 'image/gif': ext = '.gif'; break;
      case 'image/bmp': ext = '.bmp'; break;
      case 'image/svg+xml': ext = '.svg'; break;
      case 'image/tiff': ext = '.tiff'; break;
      case 'image/webp': ext = '.webp'; break;
      default: ext = '.png';
    }

    this._imageCounter++;
    const fileName = `image${this._imageCounter}${ext}`;
    if (!this._skipAssetWrites) {
      const bytes = part.getBytes();
      const filePath = path.join(this._outputDir, 'images', fileName);
      fs.writeFileSync(filePath, Buffer.from(bytes));
    }

    const relativePath = `images/${fileName}`;
    this._savedImagePaths.set(partUri, relativePath);
    return relativePath;
  },

  writeInfoJson(outputDir, inputPath, toolName, toolVersion, exportedAt, pageInfos) {
    const fileBytes = fs.readFileSync(inputPath);
    const md5 = crypto.createHash('md5').update(fileBytes).digest('hex');
    const sha256 = crypto.createHash('sha256').update(fileBytes).digest('hex');

    const info = {
      toolName,
      toolVersion,
      exportedAt,
      docxFileName: path.basename(inputPath),
      checksum: { sha256, md5 },
      fontFamily: this._allFonts.toArray(),
      embeddedFontFamily: this._embeddedFontFamilyNames.toArray(),
      missingFontFamily: this._missingFonts.toArray(),
      pages: pageInfos,
    };

    const json = serializeIndented(info);
    fs.writeFileSync(path.join(outputDir, 'info.json'), '﻿' + json + '\n', 'utf8');
  },
};
