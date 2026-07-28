'use strict';

// Port of DocxToHtml/docx/DocxConverter.HeaderFooter.cs — header/footer
// extraction and per-page resolution.

const { firstChild, elements, childElements, attr, NS } = require('../ooxml/dom');
const { StringBuilder } = require('../util/stringbuilder');

const isRelNs = (ns) => ns === NS.R || (ns || '').toLowerCase().includes('relationships');

function newHdrFtr() {
  return {
    defaultHeaderHtml: null,
    firstHeaderHtml: null,
    evenHeaderHtml: null,
    defaultFooterHtml: null,
    firstFooterHtml: null,
    evenFooterHtml: null,
    hasTitlePage: false,
  };
}

module.exports = {
  extractSectionHdrFtr(props, mainPart, previous) {
    // struct-copy semantics: inherit from the previous section, or start fresh.
    const result = previous ? Object.assign({}, previous) : newHdrFtr();
    if (props == null) return result;

    result.hasTitlePage = firstChild(props, 'titlePg', NS.W) != null;

    for (const headerRef of elements(props, 'headerReference', NS.W)) {
      const id = attr(headerRef, 'id', isRelNs);
      if (id == null) continue;
      try {
        const headerPart = mainPart.getPartById(id);
        if (headerPart == null || headerPart.root == null) continue;
        const html = this.renderHeaderFooterContent(headerPart.root, headerPart);
        if (html.trim() !== '') {
          const hfType = attr(headerRef, 'type');
          if (hfType === 'first') result.firstHeaderHtml = html;
          else if (hfType === 'even') result.evenHeaderHtml = html;
          else result.defaultHeaderHtml = html;
        }
      } catch (ex) {
        console.error(`  Warning: could not read header: ${ex.message}`);
      }
    }

    for (const footerRef of elements(props, 'footerReference', NS.W)) {
      const id = attr(footerRef, 'id', isRelNs);
      if (id == null) continue;
      try {
        const footerPart = mainPart.getPartById(id);
        if (footerPart == null || footerPart.root == null) continue;
        const html = this.renderHeaderFooterContent(footerPart.root, footerPart);
        if (html.trim() !== '') {
          const hfType = attr(footerRef, 'type');
          if (hfType === 'first') result.firstFooterHtml = html;
          else if (hfType === 'even') result.evenFooterHtml = html;
          else result.defaultFooterHtml = html;
        }
      } catch (ex) {
        console.error(`  Warning: could not read footer: ${ex.message}`);
      }
    }

    return result;
  },

  renderHeaderFooterContent(container, part) {
    const sb = new StringBuilder();
    for (const child of childElements(container)) {
      if (child.namespaceURI !== NS.W) continue;
      if (child.localName === 'p') this.renderParagraph(sb, child, part);
      else if (child.localName === 'tbl') this.renderTable(sb, child, part);
    }
    return sb.toString().trim();
  },

  resolvePageHdrFtr(hf, isFirstPageOfSection, isEvenPage) {
    let header, footer;

    if (isFirstPageOfSection && hf.hasTitlePage) header = hf.firstHeaderHtml ?? hf.defaultHeaderHtml;
    else if (isEvenPage && hf.evenHeaderHtml != null) header = hf.evenHeaderHtml;
    else header = hf.defaultHeaderHtml;

    if (isFirstPageOfSection && hf.hasTitlePage) footer = hf.firstFooterHtml ?? hf.defaultFooterHtml;
    else if (isEvenPage && hf.evenFooterHtml != null) footer = hf.evenFooterHtml;
    else footer = hf.defaultFooterHtml;

    return { header, footer };
  },
};
