'use strict';

// Port of DocxToHtml/docx/DocxConverter.PageHtml.cs — per-page CSS and the
// full HTML document wrapper.

const { StringBuilder } = require('../util/stringbuilder');
const { f0 } = require('../util/format');
const { encode } = require('./helpers');

module.exports = {
  buildCss(fontFaceCss, dims) {
    const sb = new StringBuilder();
    if (fontFaceCss !== '') sb.append(fontFaceCss);
    sb.appendLine('* { margin: 0; padding: 0; box-sizing: border-box; }');
    sb.appendLine(
      `.page-header { position: absolute; top: ${f0(dims.headerDist)}pt; left: ${f0(dims.marginLeft)}pt; right: ${f0(dims.marginRight)}pt; overflow: hidden; }`
    );
    sb.appendLine('.page-header p { margin-bottom: 0; }');
    sb.appendLine(
      `.document { position: absolute; top: ${f0(dims.marginTop)}pt; right: ${f0(dims.marginRight)}pt; bottom: ${f0(dims.marginBottom)}pt; left: ${f0(dims.marginLeft)}pt; overflow: hidden; }`
    );
    sb.appendLine(
      `.page-footer { position: absolute; bottom: ${f0(dims.footerDist)}pt; left: ${f0(dims.marginLeft)}pt; right: ${f0(dims.marginRight)}pt; overflow: hidden; }`
    );
    sb.appendLine('.page-footer p { margin-bottom: 0; }');
    sb.appendLine(
      `h1, h2, h3, h4, h5, h6 { font-family: '${encode(this._themeHeadingFont)}', 'Calibri Light', 'Calibri', sans-serif; margin-bottom: 0; }`
    );
    sb.appendLine('h1 { font-size: 16pt; color: #2F5496; margin-top: 12pt; font-weight: normal; }');
    sb.appendLine('h2 { font-size: 13pt; color: #2F5496; margin-top: 2pt; font-weight: normal; }');
    sb.appendLine('h3 { font-size: 12pt; color: #1F3763; margin-top: 2pt; font-weight: normal; }');
    sb.appendLine('h4 { font-size: 11pt; color: #2F5496; font-style: italic; margin-top: 2pt; font-weight: normal; }');
    sb.appendLine('h5 { font-size: 11pt; color: #2F5496; margin-top: 2pt; font-weight: normal; }');
    sb.appendLine('h6 { font-size: 11pt; color: #1F3763; margin-top: 2pt; font-weight: normal; }');
    sb.appendLine('p { margin-bottom: 8pt; line-height: 1.15; }');
    sb.appendLine('table { border-collapse: collapse; margin: 8pt 0; }');
    sb.appendLine('td, th { border: 1px solid #bfbfbf; padding: 3pt 5pt; vertical-align: top; }');
    sb.appendLine('td p, th p { margin-bottom: 2pt; }');
    sb.appendLine('td p:last-child, th p:last-child { margin-bottom: 0; }');
    sb.appendLine('ul, ol { margin-left: 36pt; margin-bottom: 8pt; }');
    sb.appendLine('li { margin-bottom: 2pt; line-height: 1.15; }');
    sb.appendLine('img { max-width: 100%; height: auto; display: block; margin: 4pt auto; }');
    sb.appendLine('a { color: #0563C1; text-decoration: underline; }');
    sb.appendLine('blockquote { border-left: 4px solid #2F5496; padding: 8pt 12pt; margin: 8pt 0; color: #2F5496; font-style: italic; }');
    return sb.toString();
  },

  wrapPageHtml(css, bodyContent, headerHtml, footerHtml, title, pageNumber, dims) {
    const sb = new StringBuilder();
    sb.appendLine('<!DOCTYPE html>');
    sb.appendLine('<html lang="en">');
    sb.appendLine('<head>');
    sb.appendLine('<meta charset="UTF-8">');
    sb.appendLine(`<title>${encode(title)} – Page ${pageNumber}</title>`);
    sb.appendLine('</head>');
    sb.appendLine(
      `<body style="font-family: '${encode(this._themeBodyFont)}', 'Cambria', 'Georgia', serif; font-size: 11pt; line-height: 1.15; color: #000; width: ${f0(dims.widthPx)}px; height: ${f0(dims.heightPx)}px; overflow: hidden; margin: 0; padding: 0; position: relative;">`
    );
    sb.appendLine('<style>');
    sb.append(css);
    sb.appendLine('</style>');
    if (headerHtml != null && headerHtml !== '') {
      sb.appendLine('<div class="page-header">');
      sb.appendLine(headerHtml);
      sb.appendLine('</div>');
    }
    sb.appendLine('<div class="document">');
    sb.appendLine(bodyContent);
    sb.appendLine('</div>');
    if (footerHtml != null && footerHtml !== '') {
      sb.appendLine('<div class="page-footer">');
      sb.appendLine(footerHtml);
      sb.appendLine('</div>');
    }
    sb.appendLine('</body>');
    sb.appendLine('</html>');
    return sb.toString();
  },
};
