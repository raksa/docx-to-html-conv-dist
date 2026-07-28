'use strict';

// Pure, state-free helpers ported from DocxConverter.Utilities.cs and
// FontHelper.cs. Kept as module functions (they touch no instance state).

const { attr } = require('../ooxml/dom');
const { htmlEncode } = require('../util/format');

/** System.Net.WebUtility.HtmlEncode (Encode() in the C# converter). */
const encode = (text) => htmlEncode(text == null ? '' : String(text));

/** TwipsToPt: 1/20 pt per twip. */
const twipsToPt = (twips) => twips / 20.0;

/** EmuToPt: 12700 EMU per pt. */
const emuToPt = (emu) => emu / 12700.0;

/**
 * IsBooleanOn(OnOffType?) — an on/off toggle element (w:b, w:i, w:strike, …):
 * absent -> false; present with no w:val -> true; else the parsed w:val.
 * OOXML on/off values: "1"/"true"/"on" -> true, "0"/"false"/"off" -> false.
 */
function isBooleanOn(el) {
  if (el == null) return false;
  const v = attr(el, 'val');
  if (v == null) return true;
  const lv = v.toLowerCase();
  return v === '1' || lv === 'true' || lv === 'on';
}

/** HighlightColorToHex — maps a w:highlight value to a CSS hex color. */
function highlightColorToHex(color) {
  switch (color) {
    case 'yellow': return '#FFFF00';
    case 'green': return '#00FF00';
    case 'cyan': return '#00FFFF';
    case 'magenta': return '#FF00FF';
    case 'blue': return '#0000FF';
    case 'red': return '#FF0000';
    case 'darkBlue': return '#00008B';
    case 'darkCyan': return '#008B8B';
    case 'darkGreen': return '#006400';
    case 'darkMagenta': return '#8B008B';
    case 'darkRed': return '#8B0000';
    case 'darkYellow': return '#808000';
    case 'darkGray': return '#A9A9A9';
    case 'lightGray': return '#D3D3D3';
    case 'black': return '#000000';
    default: return '#FFFF00';
  }
}

/**
 * .NET Guid.ToByteArray() byte order for a GUID string. The first three
 * groups are little-endian, the last two are stored as-is. Used for OOXML
 * embedded-font XOR deobfuscation. Returns a 16-byte Buffer, or null if the
 * string is not a valid GUID.
 */
function guidToByteArray(guidStr) {
  if (guidStr == null) return null;
  const hex = guidStr.replace(/[{}\-\s]/g, '');
  if (!/^[0-9a-fA-F]{32}$/.test(hex)) return null;
  const b = Buffer.alloc(16);
  const byteAt = (i) => parseInt(hex.substr(i * 2, 2), 16);
  // Data1 (4 bytes, little-endian)
  b[0] = byteAt(3); b[1] = byteAt(2); b[2] = byteAt(1); b[3] = byteAt(0);
  // Data2 (2 bytes, little-endian)
  b[4] = byteAt(5); b[5] = byteAt(4);
  // Data3 (2 bytes, little-endian)
  b[6] = byteAt(7); b[7] = byteAt(6);
  // Data4 (8 bytes, as-is)
  for (let i = 8; i < 16; i++) b[i] = byteAt(i);
  return b;
}

module.exports = { encode, twipsToPt, emuToPt, isBooleanOn, highlightColorToHex, guidToByteArray };
