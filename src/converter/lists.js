'use strict';

// Port of DocxToHtml/docx/DocxConverter.Lists.cs — list detection and
// nested <ul>/<ol> rendering.

const { firstChild, elements, attr, attrInt, NS } = require('../ooxml/dom');

module.exports = {
  handleListItem(sb, stack, para, numInfo, part) {
    const isBullet = this.isBulletList(numInfo.numId, numInfo.ilvl);
    const tag = isBullet ? 'ul' : 'ol';
    const targetLevel = numInfo.ilvl;

    // Close deeper levels
    while (stack.length > targetLevel + 1) {
      const last = stack[stack.length - 1];
      sb.appendLine(`</li></${last.tag}>`);
      stack.pop();
    }

    if (stack.length === targetLevel + 1) {
      // Same level: close previous li
      sb.appendLine('</li>');
    } else {
      // Open new levels
      while (stack.length <= targetLevel) {
        let olType = '';
        if (!isBullet) {
          const fmt = this.getNumberFormat(numInfo.numId, stack.length);
          if (fmt === 'lowerLetter') olType = ' type="a"';
          else if (fmt === 'upperLetter') olType = ' type="A"';
          else if (fmt === 'lowerRoman') olType = ' type="i"';
          else if (fmt === 'upperRoman') olType = ' type="I"';
        }
        sb.appendLine(`<${tag}${olType}>`);
        stack.push({ tag, numId: numInfo.numId, level: stack.length });
      }
    }

    // Build li style from paragraph formatting (alignment only)
    let liStyle = '';
    const jc = firstChild(firstChild(para, 'pPr', NS.W), 'jc', NS.W);
    const jcv = jc != null ? attr(jc, 'val') : null;
    if (jcv === 'center') liStyle = ' style="text-align:center"';
    else if (jcv === 'right') liStyle = ' style="text-align:right"';

    sb.append(`<li${liStyle}>`);
    this.renderParagraphInlineContent(sb, para, part);
  },

  closeAllLists(sb, stack) {
    while (stack.length > 0) {
      const last = stack[stack.length - 1];
      sb.appendLine(`</li></${last.tag}>`);
      stack.pop();
    }
  },

  getEffectiveNumbering(para) {
    const pPr = firstChild(para, 'pPr', NS.W);

    // Direct numPr
    const numPr = firstChild(pPr, 'numPr', NS.W);
    if (numPr != null) {
      const numId = attrInt(firstChild(numPr, 'numId', NS.W), 'val', 0);
      const ilvl = attrInt(firstChild(numPr, 'ilvl', NS.W), 'val', 0);
      if (numId > 0) return { numId, ilvl };
    }

    // Style's numPr
    const styleId = attr(firstChild(pPr, 'pStyle', NS.W), 'val');
    if (styleId != null) {
      const style = this.getStyleById(styleId);
      const styleNumPr = firstChild(firstChild(style, 'pPr', NS.W), 'numPr', NS.W);
      if (styleNumPr != null) {
        const numId = attrInt(firstChild(styleNumPr, 'numId', NS.W), 'val', 0);
        const ilvl = attrInt(firstChild(styleNumPr, 'ilvl', NS.W), 'val', 0);
        if (numId > 0) return { numId, ilvl };
      }
    }

    return null;
  },

  isBulletList(numId, ilvl) {
    return this.getNumberFormat(numId, ilvl) === 'bullet';
  },

  getNumberFormat(numId, ilvl) {
    const numbering = this._numberingPart != null ? this._numberingPart.root : null;
    if (numbering == null) return null;

    const numInstance = elements(numbering, 'num', NS.W).find((n) => attrInt(n, 'numId') === numId);
    if (numInstance == null) return null;

    const abstractNumId = attrInt(firstChild(numInstance, 'abstractNumId', NS.W), 'val');
    if (abstractNumId == null) return null;

    const abstractNum = elements(numbering, 'abstractNum', NS.W).find(
      (a) => attrInt(a, 'abstractNumId') === abstractNumId
    );
    if (abstractNum == null) return null;

    const level = elements(abstractNum, 'lvl', NS.W).find((l) => attrInt(l, 'ilvl') === ilvl);
    if (level == null) return null;

    return attr(firstChild(level, 'numFmt', NS.W), 'val');
  },
};
