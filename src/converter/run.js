'use strict';

// Port of DocxToHtml/docx/DocxConverter.cs — the Run entry point, theme
// loading, and per-section page-layout reading.

const fs = require('fs');
const path = require('path');
const { firstChild, elements, childElements, attr, attrInt, NS } = require('../ooxml/dom');
const { themePart } = require('../ooxml/parts');
const { StringBuilder } = require('../util/stringbuilder');
const { f0, bankRound } = require('../util/format');
const { twipsToPt } = require('./helpers');
const {
  PageBreakMarker,
  SectionMarkerPrefix,
  SectionMarkerSuffix,
} = require('./constants');

// Direct children of the body that carry content (paragraphs and tables).
function contentChildren(body) {
  return childElements(body).filter(
    (e) => (e.localName === 'p' || e.localName === 'tbl') && e.namespaceURI === NS.W
  );
}

module.exports = {
  run(mainPart, outputDir, inputPath, title, toolName, toolVersion, exportedAt) {
    this.resetRunState(outputDir);

    fs.mkdirSync(path.join(outputDir, 'fonts'), { recursive: true });
    fs.mkdirSync(path.join(outputDir, 'images'), { recursive: true });

    const { numberingPart, stylesPart } = require('../ooxml/parts');
    this._numberingPart = numberingPart(mainPart);
    this._stylesPart = stylesPart(mainPart);

    console.log('Extracting embedded fonts...');
    this.extractEmbeddedFonts(mainPart);
    console.log(`  ${this._embeddedDocxFonts.size} embedded font(s) found.`);

    this.loadTheme(mainPart);

    const body = firstChild(mainPart.root, 'body', NS.W);
    const bodySectPr = elements(body, 'sectPr', NS.W).pop() || null;

    // Pre-compute per-section page dimensions and headers/footers.
    const contentElements = contentChildren(body);
    const sections = this.identifySections(body, contentElements);
    const sectionDims = [];
    const sectionHdrFtrs = [];
    let prevHdrFtr = null;
    for (const section of sections) {
      const props = section.props || bodySectPr;
      sectionDims.push(this.readPageLayout(props));

      const hf = this.extractSectionHdrFtr(props, mainPart, prevHdrFtr);
      sectionHdrFtrs.push(hf);
      prevHdrFtr = hf;
    }

    const defaultDims = sectionDims[sectionDims.length - 1];
    console.log(
      `Page:   ${f0((defaultDims.widthPx * 72.0) / 96.0)}x${f0((defaultDims.heightPx * 72.0) / 96.0)}pt ` +
        `(${f0(defaultDims.widthPx)}x${f0(defaultDims.heightPx)}px)`
    );

    // Render body content with page-break and section markers.
    const bodySb = new StringBuilder();
    this.renderBodyContent(bodySb, body, mainPart);
    const bodyContent = bodySb.toString();

    // Split into pages and resolve per-page section.
    const pages = this.splitIntoPages(bodyContent);

    console.log(`Pages:  ${pages.length}`);
    console.log();

    const fontFaceCss = this.buildAllFontFaceCss();

    if (this._missingFonts.size > 0)
      console.error(`Missing fonts (${this._missingFonts.size}): ${this._missingFonts.join(', ')}`);

    const hasAnyHdrFtr = sectionHdrFtrs.some(
      (h) =>
        h.defaultHeaderHtml != null || h.firstHeaderHtml != null || h.evenHeaderHtml != null ||
        h.defaultFooterHtml != null || h.firstFooterHtml != null || h.evenFooterHtml != null
    );
    if (hasAnyHdrFtr) console.log('Headers/footers detected.');

    const pageInfos = [];
    for (let i = 0; i < pages.length; i++) {
      const secIdx = Math.min(pages[i].sectionIdx, sectionDims.length - 1);
      const dims = sectionDims[secIdx];
      const hf = sectionHdrFtrs[secIdx];

      const isFirstOfSection = i === 0 || pages[i].sectionIdx !== pages[i - 1].sectionIdx;
      const isEvenPage = (i + 1) % 2 === 0;
      const { header: headerHtml, footer: footerHtml } = this.resolvePageHdrFtr(hf, isFirstOfSection, isEvenPage);

      const css = this.buildCss(fontFaceCss, dims);
      const fileName = `page${i + 1}.html`;
      let html = this.wrapPageHtml(css, pages[i].content, headerHtml, footerHtml, title, i + 1, dims);

      html = html.split('{{PAGE_NUMBER}}').join(String(i + 1)).split('{{TOTAL_PAGES}}').join(String(pages.length));

      const outPath = path.join(outputDir, fileName);
      fs.writeFileSync(outPath, '﻿' + html, 'utf8');
      console.log(
        `  [${i + 1}/${pages.length}] ${fileName} (${Math.trunc(html.length / 1024)} KB) ` +
          `${f0(dims.widthPx)}x${f0(dims.heightPx)}px`
      );
      pageInfos.push({
        htmlFileName: fileName,
        width: bankRound(dims.widthPx),
        height: bankRound(dims.heightPx),
      });
    }

    console.log();

    this.writeInfoJson(outputDir, inputPath, toolName, toolVersion, exportedAt, pageInfos);
    console.log('Metadata written → info.json');
    console.log(`Done. ${pages.length} page(s) converted to ${outputDir}`);
  },

  // Count the pages run() would emit for this document, without writing any files.
  countPages(mainPart) {
    this.resetRunState('');
    this._skipAssetWrites = true;
    try {
      const { numberingPart, stylesPart } = require('../ooxml/parts');
      this._numberingPart = numberingPart(mainPart);
      this._stylesPart = stylesPart(mainPart);

      this.loadTheme(mainPart);

      const bodySb = new StringBuilder();
      this.renderBodyContent(bodySb, firstChild(mainPart.root, 'body', NS.W), mainPart);

      return this.splitIntoPages(bodySb.toString()).length;
    } finally {
      this._skipAssetWrites = false;
    }
  },

  // Reset the mutable per-run state shared by run() and countPages().
  resetRunState(outputDir) {
    this._outputDir = outputDir;
    this._imageCounter = 0;
    this._skipAssetWrites = false;
    this._savedImagePaths.clear();
    this._allFonts.clear();
    this._embeddedDocxFonts.clear();
    this._fontFaceCache.clear();
    this._missingFonts.clear();
    this._embeddedFontFamilyNames.clear();
    this._styleCache.clear();
  },

  // Split rendered body content on page-break markers, consuming the section
  // markers to track which section each page belongs to.
  splitIntoPages(bodyContent) {
    const rawSegments = bodyContent.split(PageBreakMarker);
    let currentSectionIdx = 0;
    const pages = []; // { content, sectionIdx }
    for (const raw of rawSegments) {
      let segment = raw;
      let markerPos = segment.indexOf(SectionMarkerPrefix);
      while (markerPos >= 0) {
        const endPos = segment.indexOf(SectionMarkerSuffix, markerPos + SectionMarkerPrefix.length);
        if (endPos < 0) break;
        const idxStr = segment.slice(markerPos + SectionMarkerPrefix.length, endPos);
        if (/^-?\d+$/.test(idxStr)) currentSectionIdx = parseInt(idxStr, 10);
        segment = segment.slice(0, markerPos) + segment.slice(endPos + SectionMarkerSuffix.length);
        markerPos = segment.indexOf(SectionMarkerPrefix);
      }
      const trimmed = segment.trim();
      pages.push({
        content: trimmed === '' ? '<p>&nbsp;</p>' : trimmed,
        sectionIdx: currentSectionIdx,
      });
    }

    // Drop trailing blank page (trailing page break with no content after).
    if (pages.length > 1 && pages[pages.length - 1].content === '<p>&nbsp;</p>')
      pages.pop();

    if (pages.length === 0) pages.push({ content: '<p>&nbsp;</p>', sectionIdx: 0 });

    return pages;
  },

  loadTheme(mainPart) {
    const theme = themePart(mainPart);
    const themeRoot = theme ? theme.root : null;
    const themeElements = firstChild(themeRoot, 'themeElements', NS.A);
    const fontScheme = firstChild(themeElements, 'fontScheme', NS.A);
    const minorLatin = firstChild(firstChild(fontScheme, 'minorFont', NS.A), 'latin', NS.A);
    const majorLatin = firstChild(firstChild(fontScheme, 'majorFont', NS.A), 'latin', NS.A);
    const minor = minorLatin ? attr(minorLatin, 'typeface') : null;
    const major = majorLatin ? attr(majorLatin, 'typeface') : null;
    this._themeBodyFont = minor || 'Cambria';
    this._themeHeadingFont = major || 'Calibri Light';
    this._allFonts.add(this._themeBodyFont);
    this._allFonts.add(this._themeHeadingFont);
  },

  readPageLayout(sectPr) {
    let pageW = 612, pageH = 792;
    let mTop = 72, mBottom = 72, mLeft = 72, mRight = 72;
    let hDist = 36, fDist = 36;

    if (sectPr != null) {
      const pgSz = firstChild(sectPr, 'pgSz', NS.W);
      if (pgSz != null) {
        const w = attrInt(pgSz, 'w');
        const h = attrInt(pgSz, 'h');
        if (w != null) pageW = twipsToPt(w);
        if (h != null) pageH = twipsToPt(h);
      }

      const pgMar = firstChild(sectPr, 'pgMar', NS.W);
      if (pgMar != null) {
        const top = attrInt(pgMar, 'top');
        const bottom = attrInt(pgMar, 'bottom');
        const left = attrInt(pgMar, 'left');
        const right = attrInt(pgMar, 'right');
        const header = attrInt(pgMar, 'header');
        const footer = attrInt(pgMar, 'footer');
        if (top != null) mTop = twipsToPt(top);
        if (bottom != null) mBottom = twipsToPt(bottom);
        if (left != null) mLeft = twipsToPt(left);
        if (right != null) mRight = twipsToPt(right);
        if (header != null) hDist = twipsToPt(header);
        if (footer != null) fDist = twipsToPt(footer);
      }
    }

    return {
      widthPx: (pageW * 96.0) / 72.0,
      heightPx: (pageH * 96.0) / 72.0,
      marginTop: mTop,
      marginBottom: mBottom,
      marginLeft: mLeft,
      marginRight: mRight,
      headerDist: hDist,
      footerDist: fDist,
    };
  },
};
