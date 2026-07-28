'use strict';

// Port of DocxToHtml/docx/DocxConverter.Paragraphs.cs — paragraph rendering,
// tag selection, and paragraph-level style building.

const { firstChild, attr, NS } = require('../ooxml/dom');
const { f1, f2 } = require('../util/format');
const { twipsToPt } = require('./helpers');
const { PageBreakMarker } = require('./constants');

// uint.TryParse — non-negative integer only.
function parseUintStr(s) {
  if (s == null) return null;
  const t = s.trim();
  if (!/^\+?\d+$/.test(t)) return null;
  const n = parseInt(t, 10);
  return Number.isFinite(n) ? n : null;
}

// int.TryParse — signed integer.
function parseIntStr(s) {
  if (s == null) return null;
  const t = s.trim();
  if (!/^[+-]?\d+$/.test(t)) return null;
  const n = parseInt(t, 10);
  return Number.isFinite(n) ? n : null;
}

// (uint) cast of a signed 32-bit value (matches C# `(uint)value`).
const toUint32 = (n) => n >>> 0;

module.exports = {
  renderParagraph(sb, para, part) {
    if (this.isPageBreakOnly(para)) {
      sb.append(PageBreakMarker);
      return;
    }

    if (this.paragraphHasInlinePageBoundaries(para)) {
      this.renderParagraphWithPageBoundaries(sb, para, part);
      return;
    }

    const tag = this.getHtmlTag(para);
    const style = this.buildParagraphStyle(para);
    const styleAttr = style !== '' ? ` style="${style}"` : '';
    sb.append(`<${tag}${styleAttr}>`);

    const hasContent = this.renderParagraphInlineContent(sb, para, part);
    if (!hasContent) sb.append('&nbsp;');

    sb.appendLine(`</${tag}>`);
  },

  getHtmlTag(para) {
    const styleId = attr(firstChild(firstChild(para, 'pPr', NS.W), 'pStyle', NS.W), 'val');
    if (styleId == null) return 'p';

    const headingLevel = this.getHeadingLevel(styleId);
    if (headingLevel >= 1 && headingLevel <= 6) return `h${headingLevel}`;
    if (headingLevel >= 7) return 'h6';

    const lower = styleId.toLowerCase();
    if (lower === 'intensequote' || lower === 'quote') return 'blockquote';
    if (lower === 'title' || lower === 'subtitle') return 'p';

    return 'p';
  },

  getHeadingLevel(styleId) {
    if (
      styleId.toLowerCase().startsWith('heading') &&
      styleId.length > 7 &&
      /^\d+$/.test(styleId.slice(7))
    ) {
      return parseInt(styleId.slice(7), 10);
    }

    // Check the style's outline level as fallback.
    const style = this.getStyleById(styleId);
    const stylePPr = firstChild(style, 'pPr', NS.W);
    const outlineLvl = firstChild(stylePPr, 'outlineLvl', NS.W);
    const val = outlineLvl != null ? parseIntStr(attr(outlineLvl, 'val')) : null;
    if (val != null) return val + 1;

    return -1;
  },

  buildParagraphStyle(para) {
    const pPr = firstChild(para, 'pPr', NS.W);
    const styleId = attr(firstChild(pPr, 'pStyle', NS.W), 'val');
    let stylePPr = null;
    if (styleId != null) stylePPr = firstChild(this.getStyleById(styleId), 'pPr', NS.W);

    const parts = [];

    // Alignment
    const jc = firstChild(pPr, 'jc', NS.W) || firstChild(stylePPr, 'jc', NS.W);
    const jcv = jc != null ? attr(jc, 'val') : null;
    if (jcv != null) {
      let ta = '';
      if (jcv === 'center') ta = 'center';
      else if (jcv === 'right') ta = 'right';
      else if (jcv === 'both') ta = 'justify';
      if (ta !== '') parts.push(`text-align:${ta}`);
    }

    // Spacing
    const spacing = firstChild(pPr, 'spacing', NS.W) || firstChild(stylePPr, 'spacing', NS.W);
    if (spacing != null) {
      const bef = parseUintStr(attr(spacing, 'before'));
      if (bef != null) parts.push(`margin-top:${f1(twipsToPt(bef))}pt`);
      const aft = parseUintStr(attr(spacing, 'after'));
      if (aft != null) parts.push(`margin-bottom:${f1(twipsToPt(aft))}pt`);
      const lineVal = parseIntStr(attr(spacing, 'line'));
      if (lineVal != null) {
        const lineRule = attr(spacing, 'lineRule');
        if (lineRule === 'auto' || lineRule == null) {
          const ratio = lineVal / 240.0;
          parts.push(`line-height:${f2(ratio)}`);
        } else {
          parts.push(`line-height:${f1(twipsToPt(toUint32(lineVal)))}pt`);
        }
      }
    }

    // Indentation
    const ind = firstChild(pPr, 'ind', NS.W) || firstChild(stylePPr, 'ind', NS.W);
    if (ind != null) {
      const leftTw = parseIntStr(attr(ind, 'left'));
      if (leftTw != null) parts.push(`margin-left:${f1(twipsToPt(toUint32(leftTw)))}pt`);
      const rightTw = parseIntStr(attr(ind, 'right'));
      if (rightTw != null) parts.push(`margin-right:${f1(twipsToPt(toUint32(rightTw)))}pt`);
      const flTw = parseIntStr(attr(ind, 'firstLine'));
      if (flTw != null) parts.push(`text-indent:${f1(twipsToPt(toUint32(flTw)))}pt`);
      const hgTw = parseIntStr(attr(ind, 'hanging'));
      if (hgTw != null) parts.push(`text-indent:-${f1(twipsToPt(toUint32(hgTw)))}pt`);
    }

    // Borders
    const pBdr = firstChild(pPr, 'pBdr', NS.W) || firstChild(stylePPr, 'pBdr', NS.W);
    if (pBdr != null) {
      this.appendBorderCss(parts, firstChild(pBdr, 'top', NS.W), 'border-top', 'padding-top');
      this.appendBorderCss(parts, firstChild(pBdr, 'bottom', NS.W), 'border-bottom', 'padding-bottom');
      this.appendBorderCss(parts, firstChild(pBdr, 'left', NS.W), 'border-left', 'padding-left');
      this.appendBorderCss(parts, firstChild(pBdr, 'right', NS.W), 'border-right', 'padding-right');
    }

    // Shading
    const shading = firstChild(pPr, 'shd', NS.W) || firstChild(stylePPr, 'shd', NS.W);
    const fill = shading != null ? attr(shading, 'fill') : null;
    if (fill != null && fill !== 'auto') {
      parts.push(`background-color:#${fill}`);
      if (!parts.some((p) => p.startsWith('padding'))) parts.push('padding:4pt 8pt');
    }

    return parts.join(';');
  },

  appendBorderCss(parts, border, cssBorder, cssPadding) {
    if (border == null) return;
    const val = attr(border, 'val');
    if (val == null) return;
    if (val === 'none' || val === 'nil') return;

    let style = 'solid';
    if (val === 'double') style = 'double';
    else if (val === 'dashed' || val === 'dashSmallGap') style = 'dashed';
    else if (val === 'dotted') style = 'dotted';

    const size = parseIntStr(attr(border, 'sz'));
    const widthPt = (size != null ? size : 4) / 8.0;
    let color = attr(border, 'color');
    if (color == null || color === 'auto') color = '000000';

    parts.push(`${cssBorder}:${f1(widthPt)}pt ${style} #${color}`);
    const space = parseIntStr(attr(border, 'space'));
    if (space != null && space > 0) parts.push(`${cssPadding}:${space}pt`);
  },
};
