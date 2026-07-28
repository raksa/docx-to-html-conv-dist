'use strict';

// Port of DocxToHtml/docx/DocxConverter.PageBoundaries.cs — paragraph/run
// rendering that is aware of inline page boundaries (explicit page breaks and
// Word's LastRenderedPageBreak markers), plus PAGE/NUMPAGES field handling.

const { firstChild, childElements, descendants, attr, innerText, NS } = require('../ooxml/dom');
const { PageBreakMarker } = require('./constants');

// LeadingPageBoundary enum
const LB = { None: 'none', Explicit: 'explicit', Rendered: 'rendered' };

module.exports = {
  renderParagraphWithPageBoundaries(sb, para, part) {
    const tag = this.getHtmlTag(para);
    const style = this.buildParagraphStyle(para);
    const styleAttr = style !== '' ? ` style="${style}"` : '';
    const paraStyleId = attr(firstChild(firstChild(para, 'pPr', NS.W), 'pStyle', NS.W), 'val');

    let paragraphOpen = false;
    let segmentHasContent = false;
    let hasAnyContent = false;

    const ensureParagraphOpen = () => {
      if (paragraphOpen) return;
      sb.append(`<${tag}${styleAttr}>`);
      paragraphOpen = true;
      segmentHasContent = false;
    };

    const markContentRendered = () => {
      segmentHasContent = true;
      hasAnyContent = true;
    };

    const closeParagraphSegment = (addPlaceholder) => {
      if (!paragraphOpen) return;
      if (addPlaceholder && !segmentHasContent) sb.append('&nbsp;');
      sb.appendLine(`</${tag}>`);
      paragraphOpen = false;
      segmentHasContent = false;
    };

    const emitPageBreak = (_dedupeWithPreviousBoundary) => {
      closeParagraphSegment(false);
      sb.append(PageBreakMarker);
    };

    let fieldDepth = 0;
    let fieldInstr = '';
    let afterSeparator = false;

    for (const child of childElements(para)) {
      if (child.localName === 'r' && child.namespaceURI === NS.W) {
        const run = child;
        const fldChar = firstChild(run, 'fldChar', NS.W);
        if (fldChar != null) {
          const fct = attr(fldChar, 'fldCharType');
          if (fct === 'begin') {
            fieldDepth++;
            fieldInstr = '';
            afterSeparator = false;
            continue;
          } else if (fct === 'separate') {
            afterSeparator = true;
            continue;
          } else if (fct === 'end') {
            fieldDepth--;
            if (fieldDepth <= 0) {
              fieldDepth = 0;
              afterSeparator = false;
            }
            continue;
          }
        }

        if (fieldDepth > 0 && !afterSeparator) {
          const instrText = firstChild(run, 'instrText', NS.W);
          if (instrText != null) fieldInstr += innerText(instrText);
          continue;
        }

        if (fieldDepth > 0 && afterSeparator) {
          const instr = fieldInstr.trim().toUpperCase();
          if (instr === 'PAGE' || instr.startsWith('PAGE ')) {
            ensureParagraphOpen();
            sb.append('{{PAGE_NUMBER}}');
            markContentRendered();
            continue;
          }
          if (instr === 'NUMPAGES' || instr.startsWith('NUMPAGES ')) {
            ensureParagraphOpen();
            sb.append('{{TOTAL_PAGES}}');
            markContentRendered();
            continue;
          }

          this.renderRunWithPageBoundaries(sb, run, part, paraStyleId, ensureParagraphOpen, markContentRendered, emitPageBreak);
          continue;
        }

        this.renderRunWithPageBoundaries(sb, run, part, paraStyleId, ensureParagraphOpen, markContentRendered, emitPageBreak);
      } else if (child.localName === 'fldSimple' && child.namespaceURI === NS.W) {
        const simpleField = child;
        const instr = (attr(simpleField, 'instr') || '').trim().toUpperCase();
        if (instr === 'PAGE' || instr.startsWith('PAGE ')) {
          ensureParagraphOpen();
          sb.append('{{PAGE_NUMBER}}');
          markContentRendered();
        } else if (instr === 'NUMPAGES' || instr.startsWith('NUMPAGES ')) {
          ensureParagraphOpen();
          sb.append('{{TOTAL_PAGES}}');
          markContentRendered();
        } else {
          for (const sfChild of childElements(simpleField)) {
            if (sfChild.localName === 'r' && sfChild.namespaceURI === NS.W) {
              ensureParagraphOpen();
              if (this.renderRun(sb, sfChild, part, paraStyleId)) markContentRendered();
            }
          }
        }
      } else if (child.localName === 'hyperlink' && child.namespaceURI === NS.W) {
        ensureParagraphOpen();
        this.renderHyperlink(sb, child, part, paraStyleId);
        markContentRendered();
      }
    }

    if (paragraphOpen) {
      closeParagraphSegment(!segmentHasContent && !hasAnyContent);
    } else if (!hasAnyContent) {
      sb.append(`<${tag}${styleAttr}>&nbsp;</${tag}>\n`);
    }
  },

  renderRunWithPageBoundaries(sb, run, part, paraStyleId, ensureParagraphOpen, markContentRendered, emitPageBreak) {
    const rPr = firstChild(run, 'rPr', NS.W);
    for (const child of childElements(run)) {
      if (child.namespaceURI !== NS.W) continue;
      if (child.localName === 'rPr') continue;

      if (child.localName === 't') {
        const text = innerText(child);
        if (text.length === 0) continue;
        ensureParagraphOpen();
        this.renderStyledText(sb, text, rPr, paraStyleId);
        markContentRendered();
      } else if (child.localName === 'br') {
        if (attr(child, 'type') === 'page') {
          emitPageBreak(false);
        } else {
          ensureParagraphOpen();
          sb.append('<br/>');
          markContentRendered();
        }
      } else if (child.localName === 'tab') {
        ensureParagraphOpen();
        sb.append('<span style="display:inline-block;width:36pt;"></span>');
        markContentRendered();
      } else if (child.localName === 'drawing') {
        ensureParagraphOpen();
        const beforeLength = sb.length;
        this.renderDrawing(sb, child, part);
        if (sb.length > beforeLength) markContentRendered();
      } else if (child.localName === 'cr') {
        ensureParagraphOpen();
        sb.append('<br/>');
        markContentRendered();
      }
    }
  },

  paragraphHasInlinePageBoundaries(para) {
    return descendants(para, 'br', NS.W).some((b) => attr(b, 'type') === 'page');
  },

  renderParagraphInlineContent(sb, para, part) {
    let hasContent = false;
    const paraStyleId = attr(firstChild(firstChild(para, 'pPr', NS.W), 'pStyle', NS.W), 'val');

    let fieldDepth = 0;
    let fieldInstr = '';
    let afterSeparator = false;

    for (const child of childElements(para)) {
      if (child.localName === 'r' && child.namespaceURI === NS.W) {
        const run = child;
        const fldChar = firstChild(run, 'fldChar', NS.W);
        if (fldChar != null) {
          const fct = attr(fldChar, 'fldCharType');
          if (fct === 'begin') {
            fieldDepth++;
            fieldInstr = '';
            afterSeparator = false;
            continue;
          } else if (fct === 'separate') {
            afterSeparator = true;
            continue;
          } else if (fct === 'end') {
            fieldDepth--;
            if (fieldDepth <= 0) {
              fieldDepth = 0;
              afterSeparator = false;
            }
            continue;
          }
        }

        if (fieldDepth > 0 && !afterSeparator) {
          const instrText = firstChild(run, 'instrText', NS.W);
          if (instrText != null) fieldInstr += innerText(instrText);
          continue;
        }

        if (fieldDepth > 0 && afterSeparator) {
          const instr = fieldInstr.trim().toUpperCase();
          if (instr === 'PAGE' || instr.startsWith('PAGE ')) {
            sb.append('{{PAGE_NUMBER}}');
            hasContent = true;
            continue;
          }
          if (instr === 'NUMPAGES' || instr.startsWith('NUMPAGES ')) {
            sb.append('{{TOTAL_PAGES}}');
            hasContent = true;
            continue;
          }
          if (this.renderRun(sb, run, part, paraStyleId)) hasContent = true;
          continue;
        }

        if (this.renderRun(sb, run, part, paraStyleId)) hasContent = true;
      } else if (child.localName === 'fldSimple' && child.namespaceURI === NS.W) {
        const simpleField = child;
        const instr = (attr(simpleField, 'instr') || '').trim().toUpperCase();
        if (instr === 'PAGE' || instr.startsWith('PAGE ')) {
          sb.append('{{PAGE_NUMBER}}');
          hasContent = true;
        } else if (instr === 'NUMPAGES' || instr.startsWith('NUMPAGES ')) {
          sb.append('{{TOTAL_PAGES}}');
          hasContent = true;
        } else {
          for (const sfChild of childElements(simpleField)) {
            if (sfChild.localName === 'r' && sfChild.namespaceURI === NS.W && this.renderRun(sb, sfChild, part, paraStyleId))
              hasContent = true;
          }
        }
      } else if (child.localName === 'hyperlink' && child.namespaceURI === NS.W) {
        this.renderHyperlink(sb, child, part, paraStyleId);
        hasContent = true;
      }
    }

    return hasContent;
  },

  shouldInsertRenderedPageBreak(elementsList, index) {
    const para = elementsList[index];
    if (!(para.localName === 'p' && para.namespaceURI === NS.W)) return false;

    if (this.getLeadingPageBoundary(para) !== LB.Rendered) return false;

    if (index === 0) return false;

    return !this.previousElementStartsNewPage(elementsList[index - 1]);
  },

  getLeadingPageBoundary(para) {
    for (const child of childElements(para)) {
      if (child.localName === 'r' && child.namespaceURI === NS.W) {
        for (const runChild of childElements(child)) {
          if (runChild.namespaceURI !== NS.W) continue;
          if (runChild.localName === 'rPr') continue;

          if (runChild.localName === 'br' && attr(runChild, 'type') === 'page') return LB.Explicit;
          if (runChild.localName === 'lastRenderedPageBreak') return LB.Rendered;
          if (runChild.localName === 't' && innerText(runChild).length > 0) return LB.None;
          if (runChild.localName === 'tab' || runChild.localName === 'drawing' || runChild.localName === 'cr')
            return LB.None;
        }
      } else if (
        (child.localName === 'fldSimple' || child.localName === 'hyperlink') &&
        child.namespaceURI === NS.W
      ) {
        return LB.None;
      }
    }

    return LB.None;
  },

  previousElementStartsNewPage(element) {
    if (!(element.localName === 'p' && element.namespaceURI === NS.W)) return false;

    if (this.isPageBreakOnly(element)) return true;

    const sectPr = firstChild(firstChild(element, 'pPr', NS.W), 'sectPr', NS.W);
    if (sectPr == null) return false;

    const sectionType = firstChild(sectPr, 'type', NS.W);
    return sectionType == null || attr(sectionType, 'val') !== 'continuous';
  },
};
