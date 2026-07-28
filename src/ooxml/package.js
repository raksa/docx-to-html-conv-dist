'use strict';

// Open Packaging Conventions (OPC) reader. Replaces the pieces of
// DocumentFormat.OpenXml.Packaging that the converter relies on:
//   WordprocessingDocument.Open / MainDocumentPart /
//   NumberingDefinitionsPart / StyleDefinitionsPart / ThemePart /
//   FontTablePart / HeaderPart / FooterPart / GetPartById /
//   HyperlinkRelationships / part.Uri / ContentType / GetStream /
//   PackageProperties.Title.

const fs = require('fs');
const { unzipSync } = require('fflate');
const { DOMParser } = require('@xmldom/xmldom');
const { NS, elements, attr, firstChild, innerText } = require('./dom');

const REL_TYPE = {
  OFFICE_DOCUMENT: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  STYLES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  NUMBERING: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  THEME: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  FONT_TABLE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  HEADER: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  FOOTER: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  HYPERLINK: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  IMAGE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  FONT: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font',
  CORE_PROPERTIES: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
};

const _domParser = new DOMParser({
  onError: () => {
    // Ignore non-fatal parse warnings; malformed parts are handled by callers.
  },
});

function normalizePartName(name) {
  return name.startsWith('/') ? name : '/' + name;
}

function dirName(partName) {
  const idx = partName.lastIndexOf('/');
  return idx <= 0 ? '/' : partName.slice(0, idx);
}

function baseName(partName) {
  const idx = partName.lastIndexOf('/');
  return idx < 0 ? partName : partName.slice(idx + 1);
}

/**
 * Approximate System.Uri(target).ToString() for external hyperlink targets:
 * an absolute URL with an empty path is canonicalized with a trailing slash,
 * matching how the C# tool serializes HyperlinkRelationship.Uri. Falls back to
 * the raw target when it is not a parseable absolute URL (e.g. mailto with odd
 * syntax, or a relative target).
 */
function dotnetUriToString(target) {
  if (target == null) return target;
  try {
    const u = new URL(target);
    return u.href;
  } catch {
    return target;
  }
}

/** Resolve an OPC relationship target (possibly relative, with ../) to an absolute part name. */
function resolveRelativeTarget(baseDir, target) {
  if (target.startsWith('/')) return target;
  const stack = baseDir === '/' ? [] : baseDir.split('/').filter(Boolean);
  for (const segment of target.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') stack.pop();
    else stack.push(segment);
  }
  return '/' + stack.join('/');
}

class Part {
  constructor(pkg, uri, contentType) {
    this.pkg = pkg;
    this.uri = uri; // e.g. "/word/document.xml"
    this.contentType = contentType;
    this._root = null;
    this._rels = null;
  }

  /** Raw bytes of the part (Uint8Array). part.GetStream() -> bytes. */
  getBytes() {
    return this.pkg.files[this.uri.slice(1)];
  }

  /** Decoded UTF-8 text of the part. */
  getText() {
    const bytes = this.getBytes();
    if (!bytes) return '';
    let text = Buffer.from(bytes).toString('utf8');
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    return text;
  }

  /** Parsed document element (root) of an XML part. Cached. */
  get root() {
    if (this._root === null) {
      const text = this.getText();
      if (!text) return null;
      try {
        this._root = _domParser.parseFromString(text, 'text/xml').documentElement;
      } catch {
        this._root = undefined;
      }
    }
    return this._root || null;
  }

  /** Relationships defined for this part, keyed by Id. */
  relationships() {
    if (this._rels !== null) return this._rels;
    const rels = new Map();
    const relsPartName = dirName(this.uri) + '/_rels/' + baseName(this.uri) + '.rels';
    const bytes = this.pkg.files[relsPartName.slice(1)];
    if (bytes) {
      let text = Buffer.from(bytes).toString('utf8');
      if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
      try {
        const relsRoot = _domParser.parseFromString(text, 'text/xml').documentElement;
        for (const rel of elements(relsRoot, 'Relationship')) {
          const id = attr(rel, 'Id');
          const type = attr(rel, 'Type');
          const target = attr(rel, 'Target');
          const mode = attr(rel, 'TargetMode');
          if (!id) continue;
          const external = mode === 'External';
          const absolute = external ? target : resolveRelativeTarget(dirName(this.uri), target);
          rels.set(id, { id, type, target, external, absolute });
        }
      } catch {
        /* ignore malformed .rels */
      }
    }
    this._rels = rels;
    return rels;
  }

  /** Resolve a relationship id to the target Part. GetPartById(relId). */
  getPartById(relId) {
    const rel = this.relationships().get(relId);
    if (!rel || rel.external) return null;
    return this.pkg.getPart(rel.absolute);
  }

  /** First related internal part of a given relationship type. */
  relatedPart(relType) {
    for (const rel of this.relationships().values()) {
      if (rel.type === relType && !rel.external) return this.pkg.getPart(rel.absolute);
    }
    return null;
  }

  /**
   * Hyperlink relationships: [{ id, uri, external }].
   * Mirrors OpenXmlPart.HyperlinkRelationships. `uri` matches System.Uri's
   * ToString() canonicalization (e.g. an authority-only URL gains a trailing
   * slash: "https://example.com" -> "https://example.com/").
   */
  hyperlinkRelationships() {
    const out = [];
    for (const rel of this.relationships().values()) {
      if (rel.type === REL_TYPE.HYPERLINK) {
        const raw = rel.external ? rel.target : rel.absolute;
        out.push({ id: rel.id, uri: dotnetUriToString(raw), external: rel.external });
      }
    }
    return out;
  }
}

class OpcPackage {
  constructor(files) {
    this.files = files; // { "word/document.xml": Uint8Array, ... } (no leading slash)
    this._parts = new Map();
    this._contentTypes = this._readContentTypes();
  }

  _readContentTypes() {
    const defaults = new Map(); // extension(lower) -> contentType
    const overrides = new Map(); // partName -> contentType
    const bytes = this.files['[Content_Types].xml'];
    if (bytes) {
      let text = Buffer.from(bytes).toString('utf8');
      if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
      try {
        const root = _domParser.parseFromString(text, 'text/xml').documentElement;
        for (const d of elements(root, 'Default')) {
          const ext = (attr(d, 'Extension') || '').toLowerCase();
          defaults.set(ext, attr(d, 'ContentType'));
        }
        for (const o of elements(root, 'Override')) {
          overrides.set(attr(o, 'PartName'), attr(o, 'ContentType'));
        }
      } catch {
        /* ignore */
      }
    }
    return { defaults, overrides };
  }

  contentTypeFor(partName) {
    const override = this._contentTypes.overrides.get(partName);
    if (override) return override;
    const dot = partName.lastIndexOf('.');
    if (dot >= 0) {
      const ext = partName.slice(dot + 1).toLowerCase();
      const def = this._contentTypes.defaults.get(ext);
      if (def) return def;
    }
    return null;
  }

  /** Get a Part by absolute part name (e.g. "/word/document.xml"). */
  getPart(partName) {
    partName = normalizePartName(partName);
    if (this._parts.has(partName)) return this._parts.get(partName);
    if (!(partName.slice(1) in this.files)) {
      this._parts.set(partName, null);
      return null;
    }
    const part = new Part(this, partName, this.contentTypeFor(partName));
    this._parts.set(partName, part);
    return part;
  }

  /** Package-level relationships (from /_rels/.rels). */
  rootRelationships() {
    const rels = new Map();
    const bytes = this.files['_rels/.rels'];
    if (!bytes) return rels;
    let text = Buffer.from(bytes).toString('utf8');
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    try {
      const root = _domParser.parseFromString(text, 'text/xml').documentElement;
      for (const rel of elements(root, 'Relationship')) {
        const id = attr(rel, 'Id');
        const type = attr(rel, 'Type');
        const target = attr(rel, 'Target');
        const mode = attr(rel, 'TargetMode');
        if (!id) continue;
        const external = mode === 'External';
        rels.set(id, {
          id,
          type,
          target,
          external,
          absolute: external ? target : resolveRelativeTarget('/', target),
        });
      }
    } catch {
      /* ignore */
    }
    return rels;
  }

  /** The main document part (word/document.xml). MainDocumentPart. */
  mainDocumentPart() {
    for (const rel of this.rootRelationships().values()) {
      if (rel.type === REL_TYPE.OFFICE_DOCUMENT && !rel.external) {
        return this.getPart(rel.absolute);
      }
    }
    return this.getPart('/word/document.xml');
  }

  /** PackageProperties.Title — the dc:title from the core-properties part. */
  coreTitle() {
    let corePart = null;
    for (const rel of this.rootRelationships().values()) {
      if (rel.type === REL_TYPE.CORE_PROPERTIES && !rel.external) {
        corePart = this.getPart(rel.absolute);
        break;
      }
    }
    if (corePart == null) corePart = this.getPart('/docProps/core.xml');
    if (corePart == null || corePart.root == null) return '';
    const titleEl = firstChild(corePart.root, 'title', NS.DC);
    return titleEl ? innerText(titleEl) : '';
  }
}

/** Open a .docx file from disk. Returns { pkg, mainPart }. */
function openDocx(docxPath) {
  const buffer = fs.readFileSync(docxPath);
  const files = unzipSync(new Uint8Array(buffer));
  const pkg = new OpcPackage(files);
  return { pkg, mainPart: pkg.mainDocumentPart() };
}

module.exports = {
  REL_TYPE,
  Part,
  OpcPackage,
  openDocx,
  NS,
};
