'use strict';

// Tool identity and shared marker strings (from DocxToHtml.cs / DocxConverter.cs).

const pkg = require('../../package.json');

module.exports = {
  // `pkg.name` is the npm id ("docx-to-html-conv"); info.json must carry the same
  // tool name the C# port writes, so it stays a literal here.
  ToolName: 'DocxToHtml',
  ToolVersion: pkg.version,

  PageBreakMarker: '<!--PB-->',
  SectionMarkerPrefix: '<!--SEC:',
  SectionMarkerSuffix: '-->',
};
