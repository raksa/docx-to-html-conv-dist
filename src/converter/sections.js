'use strict';

// Port of DocxToHtml/docx/DocxConverter.Sections.cs — body content rendering
// with section handling and section identification.

const { firstChild, elements, childElements, attr, attrInt, NS } = require('../ooxml/dom');
const { PageBreakMarker, SectionMarkerPrefix, SectionMarkerSuffix } = require('./constants');

function contentChildren(body) {
  return childElements(body).filter(
    (e) => (e.localName === 'p' || e.localName === 'tbl') && e.namespaceURI === NS.W
  );
}

module.exports = {
  renderBodyContent(sb, body, part) {
    const contentElements = contentChildren(body);

    const sections = this.identifySections(body, contentElements);
    let sectionIdx = 0;

    for (const section of sections) {
      sb.append(`${SectionMarkerPrefix}${sectionIdx}${SectionMarkerSuffix}`);
      sectionIdx++;
      let numCols = 1;
      if (section.props != null) {
        const cols = firstChild(section.props, 'cols', NS.W);
        const cc = cols != null ? attrInt(cols, 'num') : null;
        if (cc != null) numCols = cc;
      }

      if (numCols > 1) sb.appendLine(`<div style="column-count:${numCols};column-gap:36pt;">`);

      const listStack = [];

      for (let i = section.startIdx; i <= section.endIdx && i < contentElements.length; i++) {
        const el = contentElements[i];
        if (el.localName === 'p') {
          if (this.shouldInsertRenderedPageBreak(contentElements, i)) {
            this.closeAllLists(sb, listStack);
            sb.append(PageBreakMarker);
          }

          const numInfo = this.getEffectiveNumbering(el);
          if (numInfo != null) this.handleListItem(sb, listStack, el, numInfo, part);
          else {
            this.closeAllLists(sb, listStack);
            this.renderParagraph(sb, el, part);
          }
        } else if (el.localName === 'tbl') {
          this.closeAllLists(sb, listStack);
          this.renderTable(sb, el, part);
        }
      }

      this.closeAllLists(sb, listStack);

      if (numCols > 1) sb.appendLine('</div>');

      // Page break between sections.
      if (section.props != null) {
        const sectionType = firstChild(section.props, 'type', NS.W);
        const isPageBreak = sectionType == null || attr(sectionType, 'val') !== 'continuous';
        if (isPageBreak && section.endIdx < contentElements.length - 1) sb.append(PageBreakMarker);
      }
    }
  },

  identifySections(body, elementsList) {
    const sections = [];
    let currentStart = 0;

    for (let i = 0; i < elementsList.length; i++) {
      const el = elementsList[i];
      if (el.localName === 'p') {
        const sectPr = firstChild(firstChild(el, 'pPr', NS.W), 'sectPr', NS.W);
        if (sectPr != null) {
          sections.push({ startIdx: currentStart, endIdx: i, props: sectPr });
          currentStart = i + 1;
        }
      }
    }

    // Last section uses body-level sectPr.
    sections.push({ startIdx: currentStart, endIdx: elementsList.length - 1, props: null });

    return sections;
  },
};
