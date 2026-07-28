'use strict';

// Named part-navigation helpers mirroring the typed part relationships the
// C# converter walks (MainDocumentPart.NumberingDefinitionsPart /
// StyleDefinitionsPart / ThemePart / FontTablePart).

const { REL_TYPE } = require('./package');

function numberingPart(mainPart) {
  return mainPart ? mainPart.relatedPart(REL_TYPE.NUMBERING) : null;
}

function stylesPart(mainPart) {
  return mainPart ? mainPart.relatedPart(REL_TYPE.STYLES) : null;
}

function themePart(mainPart) {
  return mainPart ? mainPart.relatedPart(REL_TYPE.THEME) : null;
}

function fontTablePart(mainPart) {
  return mainPart ? mainPart.relatedPart(REL_TYPE.FONT_TABLE) : null;
}

module.exports = {
  numberingPart,
  stylesPart,
  themePart,
  fontTablePart,
};
