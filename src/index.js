'use strict';

// Port of DocxToHtml/DocxToHtml.cs — the shared entry point for converting a
// .docx document to paginated HTML assets.

const fs = require('fs');
const path = require('path');
const { openDocx } = require('./ooxml/package');
const { firstChild, NS } = require('./ooxml/dom');
const { DocxConverter } = require('./converter/DocxConverter');
const { ToolName, ToolVersion } = require('./converter/constants');

function pad2(n) {
  return String(n).padStart(2, '0');
}

function formatExportedAt(date) {
  return (
    `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())} ` +
    `${pad2(date.getHours())}:${pad2(date.getMinutes())}:${pad2(date.getSeconds())}`
  );
}

/**
 * Convert a .docx document to paginated HTML assets.
 * @param {string} docxPath Path to the input .docx file.
 * @param {string} outputDirectory Folder to write HTML and assets (created if missing).
 */
function exportDocx(docxPath, outputDirectory) {
  if (docxPath == null || String(docxPath).trim() === '') throw new Error('docxPath is required.');
  if (outputDirectory == null || String(outputDirectory).trim() === '')
    throw new Error('outputDirectory is required.');

  if (!fs.existsSync(docxPath)) throw new Error(`File not found: ${docxPath}`);

  fs.mkdirSync(outputDirectory, { recursive: true });

  const { pkg, mainPart } = openDocx(docxPath);
  const body = mainPart != null ? firstChild(mainPart.root, 'body', NS.W) : null;
  if (body == null) throw new Error('Invalid or empty document.');

  let title = '';
  try {
    title = pkg.coreTitle() || '';
  } catch {
    /* ignore missing/malformed core properties */
  }
  if (title === '') title = path.basename(docxPath, path.extname(docxPath));

  console.log(`Input:  ${docxPath}`);
  console.log(`Output: ${outputDirectory}`);

  const exportedAt = formatExportedAt(new Date());

  const converter = new DocxConverter();
  converter.run(mainPart, outputDirectory, docxPath, title, ToolName, ToolVersion, exportedAt);
}

/**
 * Count the pages exportDocx would produce for a .docx document.
 * @param {string} docxPath Path to the input .docx file.
 * @returns {number} The number of pages.
 */
function getPagesCount(docxPath) {
  if (docxPath == null || String(docxPath).trim() === '') throw new Error('docxPath is required.');

  if (!fs.existsSync(docxPath)) throw new Error(`File not found: ${docxPath}`);

  const { mainPart } = openDocx(docxPath);
  const body = mainPart != null ? firstChild(mainPart.root, 'body', NS.W) : null;
  if (body == null) throw new Error('Invalid or empty document.');

  const converter = new DocxConverter();
  return converter.countPages(mainPart);
}

function getVersion() {
  return ToolVersion;
}

module.exports = { exportDocx, getPagesCount, getVersion, DocxConverter };
