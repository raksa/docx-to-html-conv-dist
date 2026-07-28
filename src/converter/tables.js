'use strict';

// Port of DocxToHtml/docx/DocxConverter.Tables.cs — table rendering with
// grid-span and vertical-merge handling.

const { firstChild, elements, childElements, attr, attrInt, NS } = require('../ooxml/dom');

function gridSpanOf(cell) {
  const tcPr = firstChild(cell, 'tcPr', NS.W);
  return attrInt(firstChild(tcPr, 'gridSpan', NS.W), 'val', 1);
}

module.exports = {
  renderTable(sb, table, part) {
    const rows = elements(table, 'tr', NS.W);
    if (rows.length === 0) return;

    const mergeMap = this.calculateVerticalMerges(rows);

    // Table alignment
    const tblPr = firstChild(table, 'tblPr', NS.W);
    const jc = firstChild(tblPr, 'jc', NS.W);
    let tblStyle = '';
    if (jc != null && attr(jc, 'val') === 'center')
      tblStyle = ' style="margin-left:auto;margin-right:auto;"';

    sb.appendLine(`<table${tblStyle}>`);

    for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
      sb.append('<tr>');
      const cells = elements(rows[rowIdx], 'tc', NS.W);

      let logicalCol = 0;
      for (let cellIdx = 0; cellIdx < cells.length; cellIdx++) {
        const cell = cells[cellIdx];
        const tcPr = firstChild(cell, 'tcPr', NS.W);

        const colspan = gridSpanOf(cell);

        // Vertical merge: skip continuation cells
        const key = `${rowIdx},${logicalCol}`;
        const rowspan = mergeMap.has(key) ? mergeMap.get(key) : 1;
        if (rowspan === 0) {
          logicalCol += colspan;
          continue;
        }

        let attrs = '';
        if (colspan > 1) attrs += ` colspan="${colspan}"`;
        if (rowspan > 1) attrs += ` rowspan="${rowspan}"`;

        const cellStyle = [];

        // Shading
        const shading = firstChild(tcPr, 'shd', NS.W);
        const fill = shading != null ? attr(shading, 'fill') : null;
        if (fill != null && fill !== 'auto') cellStyle.push(`background-color:#${fill}`);

        // Vertical alignment
        const vAlign = firstChild(tcPr, 'vAlign', NS.W);
        const vav = vAlign != null ? attr(vAlign, 'val') : null;
        if (vav != null) {
          let va = 'top';
          if (vav === 'center') va = 'middle';
          else if (vav === 'bottom') va = 'bottom';
          if (va !== 'top') cellStyle.push(`vertical-align:${va}`);
        }

        const sAttr = cellStyle.length > 0 ? ` style="${cellStyle.join(';')}"` : '';

        sb.append(`<td${attrs}${sAttr}>`);

        for (const child of childElements(cell)) {
          if (child.namespaceURI !== NS.W) continue;
          if (child.localName === 'p') this.renderParagraph(sb, child, part);
          else if (child.localName === 'tbl') this.renderTable(sb, child, part);
        }

        sb.append('</td>');
        logicalCol += colspan;
      }

      sb.appendLine('</tr>');
    }

    sb.appendLine('</table>');
  },

  calculateVerticalMerges(rows) {
    const result = new Map(); // "row,col" -> span

    // Build logical-column mapping per row (accounting for GridSpan)
    const colMaps = [];
    for (let r = 0; r < rows.length; r++) {
      const cells = elements(rows[r], 'tc', NS.W);
      const map = new Array(cells.length);
      let logicalCol = 0;
      for (let c = 0; c < cells.length; c++) {
        map[c] = logicalCol;
        logicalCol += gridSpanOf(cells[c]);
      }
      colMaps.push(map);
    }

    // Track active merges: logicalCol → { row, col }
    const mergeStarts = new Map();

    for (let r = 0; r < rows.length; r++) {
      const cells = elements(rows[r], 'tc', NS.W);
      const activeLogicals = new Set();

      for (let c = 0; c < cells.length; c++) {
        const logicalCol = colMaps[r][c];
        activeLogicals.add(logicalCol);
        const vMerge = firstChild(firstChild(cells[c], 'tcPr', NS.W), 'vMerge', NS.W);

        if (vMerge != null) {
          if (attr(vMerge, 'val') === 'restart') {
            mergeStarts.set(logicalCol, { row: r, col: logicalCol });
            result.set(`${r},${logicalCol}`, 1);
          } else {
            if (mergeStarts.has(logicalCol)) {
              const start = mergeStarts.get(logicalCol);
              const startKey = `${start.row},${start.col}`;
              result.set(startKey, (result.has(startKey) ? result.get(startKey) : 1) + 1);
              result.set(`${r},${logicalCol}`, 0);
            }
          }
        } else {
          mergeStarts.delete(logicalCol);
        }
      }

      // End merges for columns not present in this row
      for (const lc of [...mergeStarts.keys()]) {
        if (!activeLogicals.has(lc)) mergeStarts.delete(lc);
      }
    }

    return result;
  },
};
