# CLAUDE.md — xlsx-populate Codebase Guide

## Project Overview

**xlsx-populate** is a JavaScript library for parsing and generating Excel XLSX files with dual Node.js and browser support. Its key design philosophy is **preservation**: it manipulates the underlying XML directly so existing workbook features (charts, macros, styles) are preserved even when not explicitly supported by the library.

- **Version**: 1.21.0
- **License**: MIT
- **Main entry**: `lib/XlsxPopulate.js`
- **Browser entry**: `browser/xlsx-populate.js`
- **Node.js minimum**: v4+

---

## Repository Structure

```
xlsx-populate/
├── lib/                    # Core library source (29 ES6 modules)
├── test/
│   ├── unit/              # Jasmine unit tests (one spec per lib module)
│   ├── e2e-generate/      # Integration tests: generating workbooks
│   ├── e2e-parse/         # Integration tests: parsing existing files
│   ├── e2e-browser/       # Browser-specific integration tests
│   ├── helpers/           # Custom Jasmine matchers and async utilities
│   └── files/             # Test fixture XLSX files
├── examples/              # Runnable usage examples (basic, browser, encryption, ranges, styles)
├── browser/               # Pre-built browser bundles (committed artifacts)
├── blank/                 # Default blank workbook template (base64-encoded)
├── docs/                  # JSDoc output (generated)
├── gulpfile.js            # Build automation
├── .eslintrc.json         # Strict ESLint configuration
└── package.json
```

---

## Core Architecture

### How XLSX Files Work

XLSX is a ZIP archive of XML files. The library uses **jszip** to open/write the archive and **sax** for XML parsing. Key paths inside an XLSX:

| Path | Contents |
|------|----------|
| `xl/workbook.xml` | Sheet list, defined names, workbook properties |
| `xl/worksheets/sheet{n}.xml` | Cell data for each sheet |
| `xl/styles.xml` | Style definitions (fonts, fills, borders, cell formats) |
| `xl/sharedStrings.xml` | Deduplicated string table |
| `xl/relationships/` | Internal link registry |
| `[Content_Types].xml` | MIME type registry |

### Module Map

#### Public API Layer

| Module | Role |
|--------|------|
| `XlsxPopulate.js` | Entry point / static factory methods |
| `Workbook.js` | Workbook container — sheets, metadata, I/O |
| `Sheet.js` | Worksheet — cells, rows, columns, ranges |
| `Cell.js` | Individual cell — value, formula, style, hyperlink |
| `Range.js` | Multi-cell operations — bulk set/get, iterators |
| `Row.js` | Row — height, hidden, style |
| `Column.js` | Column — width, hidden, style |

#### Styling

| Module | Role |
|--------|------|
| `Style.js` | Per-cell style object (font, fill, border, alignment) |
| `StyleSheet.js` | Style deduplication/caching backed by `xl/styles.xml` |
| `RichText.js` | Rich text value (multiple fragments with mixed styles) |
| `RichTextFragment.js` | Single styled fragment within a rich text value |

#### XML Infrastructure

| Module | Role |
|--------|------|
| `XmlParser.js` | SAX-based XML → JSON-like object tree |
| `XmlBuilder.js` | JSON-like object tree → XML string |
| `xmlq.js` | Query/mutate XML node trees (find, append, remove) |
| `SharedStrings.js` | Shared string table (index ↔ string lookup) |
| `ContentTypes.js` | `[Content_Types].xml` wrapper |
| `Relationships.js` | Relationship XML wrapper |

#### Utilities

| Module | Role |
|--------|------|
| `ArgHandler.js` | Runtime method overloading by argument types |
| `addressConverter.js` | A1 ↔ row/col index conversions |
| `dateConverter.js` | Excel serial date ↔ JS Date conversions |
| `FormulaError.js` | Typed formula error values (e.g. `#REF!`) |
| `colorIndexes.js` | Legacy indexed color lookup table |
| `regexify.js` | Convert glob patterns to RegExp |
| `externals.js` | Dependency injection shim (overridable Promise, etc.) |
| `blank.js` | Base64-encoded minimal blank workbook |

#### Optional / Advanced

| Module | Role |
|--------|------|
| `Encryptor.js` | CFB-based XLSX encryption/decryption (password protection) |
| `AppProperties.js` | `xl/app.xml` — application-level properties |
| `CoreProperties.js` | `xl/core.xml` — document metadata (author, title, etc.) |
| `PageBreaks.js` | Row/column page break management |

---

## Key Design Patterns

### 1. XML Preservation

Rather than fully deserializing to rich objects (losing unsupported attributes), the library parses XML into plain JSON-like node trees and mutates them in-place. Only the touched portions are understood by the library; everything else is passed through on re-serialization.

### 2. Fluent / Method Chaining API

Setters return `this`; getters return values. This enables jQuery-style chaining:

```javascript
workbook.sheet(0).cell("A1").value("Hello").style("bold", true);
```

### 3. ArgHandler — Method Overloading

`ArgHandler` matches argument signatures at runtime to route to the correct handler. Virtually every public method that acts as both getter and setter uses this. When modifying method signatures, add/update entries in the `ArgHandler` chain rather than using manual `if` branches.

```javascript
// Inside a class method:
return new ArgHandler("Cell.value", arguments)
    .case([], () => /* getter */)
    .case(['*'], value => /* setter */)
    .handle();
```

### 4. Private Members

- Private instance fields: `_fieldName`
- Private methods: `_methodName`
- Constructor logic often delegated to `_init()`

### 5. Dual Environment (Node + Browser)

- Node.js: Full file system access via `toFileAsync`/`fromFileAsync`.
- Browser: These file methods are absent; use `toDataAsync`/`fromDataAsync` with `Blob` or `ArrayBuffer`.
- `process.browser` is checked internally to gate file operations.
- Two browser bundles: with and without `Encryptor` (for smaller size).

### 6. Lazy I/O

Parsed XML nodes are kept in memory. Changes accumulate. The workbook is only re-serialized (XML rebuilt, ZIP written) when `toFileAsync()` or `toDataAsync()` is called.

---

## Development Workflows

### Setup

```bash
npm install
```

### Running Tests

```bash
# Unit tests (fast, no file I/O)
npm test

# End-to-end: generating workbooks
npm run e2e-generate

# End-to-end: parsing existing files
npm run e2e-parse

# All tests (Windows-compatible)
npm run test:windows
```

### Continuous Development (watch mode)

```bash
gulp          # Watches lib/ and test/, runs lint + unit tests on change
```

### Building Browser Bundles

```bash
gulp build    # Produces browser/xlsx-populate.js and browser/xlsx-populate-no-encryption.js
```

Build pipeline: **Browserify → Babelify (babel-preset-env) → Uglify**

### Linting

```bash
gulp lint     # ESLint on lib/**/*.js
```

### Generating Docs

```bash
npm run docs  # JSDoc → docs/
```

### Browser Tests

```bash
gulp karma    # Karma test runner with Chrome/Firefox/IE
```

---

## Code Conventions

### Language

- **ES6** throughout: `const`/`let` (never `var`), arrow functions, template literals, classes.
- No Babel for Node source — relies on native Node ES6 support.

### Naming

| Thing | Convention | Example |
|-------|-----------|---------|
| Classes | PascalCase | `StyleSheet`, `XmlParser` |
| Methods/properties | camelCase | `activeSheet()`, `toFileAsync()` |
| Private fields/methods | `_` prefix | `_node`, `_init()` |
| Constants | UPPER_SNAKE_CASE | `MIME_TYPE` |

### ESLint

The `.eslintrc.json` is very strict (218+ rules). Notable rules:
- `eqeqeq`: always use `===`
- `no-var`: enforced
- `indent`: 4 spaces
- `quotes`: double quotes
- `comma-dangle`: never

Always run `gulp lint` before committing.

### JSDoc

Every public method must have:
- `@param` with type and description for each argument
- `@returns` with type and description
- `@throws` where applicable

### Tests

- One spec file per lib module: `test/unit/Foo.spec.js` mirrors `lib/Foo.js`
- Use **Jasmine 3.5** syntax (`describe`/`it`/`expect`)
- Use `proxyquire` to mock dependencies
- Custom matchers in `test/helpers/matchers.js`:
  - `toEqualJson(expected)` — deep JSON diff
  - `toEqualUInt8Array(expected)` — byte-level buffer comparison

---

## Common Tasks

### Adding a New Style Property

1. Add getter/setter logic in `lib/Style.js` using the ArgHandler pattern.
2. Map the property name in `lib/StyleSheet.js` if it requires custom XML serialization.
3. Add unit tests in `test/unit/Style.spec.js`.

### Adding a New Cell Feature

1. Implement in `lib/Cell.js` (getter/setter via ArgHandler).
2. If the feature involves a new XML namespace or relationship type, update `lib/Relationships.js` and `lib/ContentTypes.js`.
3. Write unit tests and, if reading/writing files is involved, add an e2e fixture.

### Supporting a New XML Element

1. Parse in the appropriate `_init()` method using `xmlq` helpers.
2. Serialize back in the appropriate `_toXml()` / `toObject()` method.
3. Do not strip unknown attributes — preserve them for round-trip fidelity.

### Modifying the Browser Bundle

- Edit `gulpfile.js` to adjust Browserify entry points or Babel targets.
- After `gulp build`, commit the updated files in `browser/`.

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `jszip` | ZIP read/write (XLSX container) |
| `sax` | Streaming XML parser |
| `lodash` | Utility functions |
| `cfb` | Compound File Binary format (encryption) |
| `n-readlines` | Line-by-line file reading (Node only) |

Dev tools: `gulp`, `jasmine`, `karma`, `browserify`, `babelify`, `uglify-js`, `proxyquire`, `eslint`.

---

## Known Limitations / Backlog

Features explicitly **not yet supported** (see `backlog.md`):

- Conditional formatting
- Charts
- Cell comments/notes
- Frozen rows/columns
- Images
- Array formulas
- Full defined names support
- Shared formula relocation when the reference cell is deleted

When working in these areas, preserve unknown XML rather than discarding it.

---

## CI

Travis CI (`.travis.yml`) runs:
- `npm test` (unit tests)
- `npm run e2e-generate`
- `npm run e2e-parse`

All three must pass before merging.
