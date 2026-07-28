'use strict';

// Port of DocxToHtml/docx/DocxConverter.InlineElements.cs — image (Drawing)
// and hyperlink rendering.

const { firstChild, childElements, attr, attrLong, NS } = require('../ooxml/dom');
const { f1 } = require('../util/format');
const { encode, emuToPt } = require('./helpers');

const isRelNs = (ns) => ns === NS.R || (ns || '').toLowerCase().includes('relationships');

module.exports = {
  renderDrawing(sb, drawing, part) {
    try {
      const inline = firstChild(drawing, 'inline', NS.WP);
      const anchor = firstChild(drawing, 'anchor', NS.WP);

      const container = inline || anchor;
      if (container == null) return;

      // Dimensions
      let cx = 0, cy = 0;
      if (inline != null) {
        const ext = firstChild(inline, 'extent', NS.WP);
        cx = (ext != null ? attrLong(ext, 'cx') : null) || 0;
        cy = (ext != null ? attrLong(ext, 'cy') : null) || 0;
      } else if (anchor != null) {
        const ext = firstChild(anchor, 'extent', NS.WP);
        cx = (ext != null ? attrLong(ext, 'cx') : null) || 0;
        cy = (ext != null ? attrLong(ext, 'cy') : null) || 0;
      }

      // Navigate: Graphic → GraphicData → Picture → BlipFill → Blip
      const graphic = firstChild(container, 'graphic', NS.A);
      const graphicData = firstChild(graphic, 'graphicData', NS.A);
      const pic = firstChild(graphicData, 'pic', NS.PIC);
      const blipFill = firstChild(pic, 'blipFill', NS.PIC);
      const blip = firstChild(blipFill, 'blip', NS.A);

      const embed = blip != null ? attr(blip, 'embed', isRelNs) : null;
      if (embed == null) return;

      const imagePart = part.getPartById(embed);
      if (imagePart == null) return;
      const imgPath = this.saveImageFile(imagePart);

      const styleFragments = [];
      if (cx > 0) styleFragments.push(`width:${f1(emuToPt(cx))}pt`);
      if (cy > 0) styleFragments.push(`height:${f1(emuToPt(cy))}pt`);

      const imgStyle = styleFragments.length > 0 ? ` style="${styleFragments.join(';')}"` : '';

      sb.append(`<img src="${imgPath}"${imgStyle}/>`);
    } catch (ex) {
      console.error(`  Warning: could not extract image: ${ex.message}`);
    }
  },

  renderHyperlink(sb, hyperlink, part, paraStyleId) {
    let url = null;

    const id = attr(hyperlink, 'id', isRelNs);
    const anchor = attr(hyperlink, 'anchor');
    if (id != null) {
      const rel = part.hyperlinkRelationships().find((r) => r.id === id);
      url = rel ? rel.uri : null;
    } else if (anchor != null) {
      url = `#${anchor}`;
    }

    if (url != null) sb.append(`<a href="${encode(url)}">`);

    for (const child of childElements(hyperlink)) {
      if (child.localName === 'r' && child.namespaceURI === NS.W) this.renderRun(sb, child, part, paraStyleId);
    }

    if (url != null) sb.append('</a>');
  },
};
