'use strict';

// Port of DocxToHtml/docx/DocxConverter.Runs.cs — run rendering and
// run-level (character) style building.

const { firstChild, childElements, attr, innerText, NS } = require('../ooxml/dom');
const { encode, isBooleanOn, highlightColorToHex } = require('./helpers');

// C# `double.ToString("G", InvariantCulture)` for the small half-point font
// sizes produced here (always X.0 or X.5): shortest round-trippable form.
function dotnetG(value) {
  return String(value);
}

// First non-null of two attribute values (mirrors the `?? ` chain in C#,
// where an empty string counts as present).
const firstNonNull = (a, b) => (a != null ? a : b);

module.exports = {
  renderRun(sb, run, part, paraStyleId) {
    let rendered = false;
    const rPr = firstChild(run, 'rPr', NS.W);

    for (const child of childElements(run)) {
      if (child.localName === 'rPr' && child.namespaceURI === NS.W) continue;

      if (child.localName === 't' && child.namespaceURI === NS.W) {
        this.renderStyledText(sb, innerText(child), rPr, paraStyleId);
        rendered = true;
      } else if (child.localName === 'br' && child.namespaceURI === NS.W) {
        if (attr(child, 'type') !== 'page') {
          sb.append('<br/>');
          rendered = true;
        }
      } else if (child.localName === 'tab' && child.namespaceURI === NS.W) {
        sb.append('<span style="display:inline-block;width:36pt;"></span>');
        rendered = true;
      } else if (child.localName === 'drawing' && child.namespaceURI === NS.W) {
        this.renderDrawing(sb, child, part);
        rendered = true;
      } else if (child.localName === 'cr' && child.namespaceURI === NS.W) {
        sb.append('<br/>');
        rendered = true;
      }
    }

    return rendered;
  },

  renderStyledText(sb, text, directRPr, paraStyleId) {
    if (text == null || text === '') return;

    let styleRPr = null;
    if (paraStyleId != null) styleRPr = firstChild(this.getStyleById(paraStyleId), 'rPr', NS.W);

    const style = this.buildRunStyle(directRPr, styleRPr);
    const needsSpan = style !== '';

    const vertAlign = firstChild(directRPr, 'vertAlign', NS.W) || firstChild(styleRPr, 'vertAlign', NS.W);
    const va = vertAlign != null ? attr(vertAlign, 'val') : null;
    const isSup = va === 'superscript';
    const isSub = va === 'subscript';

    if (isSup) sb.append('<sup>');
    else if (isSub) sb.append('<sub>');
    if (needsSpan) sb.append(`<span style="${style}">`);
    sb.append(encode(text));
    if (needsSpan) sb.append('</span>');
    if (isSup) sb.append('</sup>');
    else if (isSub) sb.append('</sub>');
  },

  buildRunStyle(d, s) {
    const parts = [];

    // Bold
    const bold = firstChild(d, 'b', NS.W) || firstChild(s, 'b', NS.W);
    if (isBooleanOn(bold)) parts.push('font-weight:bold');

    // Italic
    const italic = firstChild(d, 'i', NS.W) || firstChild(s, 'i', NS.W);
    if (isBooleanOn(italic)) parts.push('font-style:italic');

    // Underline + strikethrough → combined text-decoration
    const underline = firstChild(d, 'u', NS.W) || firstChild(s, 'u', NS.W);
    const strike = firstChild(d, 'strike', NS.W) || firstChild(s, 'strike', NS.W);
    const dstrike = firstChild(d, 'dstrike', NS.W) || firstChild(s, 'dstrike', NS.W);
    const uval = underline != null ? attr(underline, 'val') : null;

    const deco = [];
    if (uval != null && uval !== 'none') deco.push('underline');
    if (isBooleanOn(strike) || isBooleanOn(dstrike)) deco.push('line-through');
    if (deco.length > 0) parts.push(`text-decoration:${deco.join(' ')}`);

    // Underline style
    if (uval != null) {
      let uStyle = null;
      if (uval === 'double') uStyle = 'double';
      else if (uval === 'wave' || uval === 'wavyDouble' || uval === 'wavyHeavy') uStyle = 'wavy';
      else if (uval === 'dotted' || uval === 'dottedHeavy') uStyle = 'dotted';
      else if (
        uval === 'dash' || uval === 'dashedHeavy' || uval === 'dashLong' ||
        uval === 'dashLongHeavy' || uval === 'dashDotDotHeavy' || uval === 'dashDotHeavy'
      )
        uStyle = 'dashed';
      if (uStyle != null) parts.push(`text-decoration-style:${uStyle}`);
      if (uval === 'thick') parts.push('text-decoration-thickness:2px');
    }

    // Font size (half-points → points)
    const fontSize = firstChild(d, 'sz', NS.W) || firstChild(s, 'sz', NS.W);
    const szVal = fontSize != null ? attr(fontSize, 'val') : null;
    if (szVal != null && /^[+-]?(\d+\.?\d*|\.\d+)$/.test(szVal.trim())) {
      const hp = parseFloat(szVal);
      const pt = hp / 2.0;
      parts.push(`font-size:${dotnetG(pt)}pt`);
    }

    // Font color
    const color = firstChild(d, 'color', NS.W) || firstChild(s, 'color', NS.W);
    const colorVal = color != null ? attr(color, 'val') : null;
    if (colorVal != null && colorVal !== 'auto') parts.push(`color:#${colorVal}`);

    // Font family
    const fonts = firstChild(d, 'rFonts', NS.W) || firstChild(s, 'rFonts', NS.W);
    const fontName = fonts != null ? firstNonNull(attr(fonts, 'ascii'), attr(fonts, 'hAnsi')) : null;
    if (fontName != null) {
      this._allFonts.add(fontName);
      parts.push(`font-family:'${fontName}'`);
    }

    // Highlight
    const highlight = firstChild(d, 'highlight', NS.W) || firstChild(s, 'highlight', NS.W);
    const hlVal = highlight != null ? attr(highlight, 'val') : null;
    if (hlVal != null && hlVal !== 'none') parts.push(`background-color:${highlightColorToHex(hlVal)}`);

    // Run-level shading (e.g., code background)
    const shading = firstChild(d, 'shd', NS.W) || firstChild(s, 'shd', NS.W);
    const shFill = shading != null ? attr(shading, 'fill') : null;
    if (shFill != null && shFill !== 'auto' && highlight == null) parts.push(`background-color:#${shFill}`);

    // Small caps
    const smallCaps = firstChild(d, 'smallCaps', NS.W) || firstChild(s, 'smallCaps', NS.W);
    if (isBooleanOn(smallCaps)) parts.push('font-variant:small-caps');

    // All caps
    const caps = firstChild(d, 'caps', NS.W) || firstChild(s, 'caps', NS.W);
    if (isBooleanOn(caps)) parts.push('text-transform:uppercase');

    return parts.join(';');
  },
};
