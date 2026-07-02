# Changelog

Notable changes to `@mitresthen/excelents`. The format follows
[Keep a Changelog](https://keepachangelog.com/); pre-1.0, minor versions may contain breaking
changes. GitHub Releases carry the full generated notes for each tag.

## [Unreleased]

### Added

- Browser test suite (Vitest browser mode, Chromium) run in CI.
- `bench/` — reproducible excelents-vs-exceljs benchmark (`pnpm bench`); results in the README.
- ExcelJS migration guide (`docs/MIGRATING-FROM-EXCELJS.md`).
- npm keywords, rewritten README, changelog.

## [0.3.0] — 2026-07-02

### Added

- Embedded images on write: `workbook.addImage` (png/jpeg/gif, bytes or base64) +
  `worksheet.placeImage` with cell anchor and pixel size.
- `worksheet.setAutoFilter(ref)` — autoFilter over a range (write).
- `worksheet.freeze({ rows, cols })` — frozen panes (write).
- `alignment.indent` cell style (write).

## [0.2.0] — 2026-07-01

### Changed

- **Breaking:** ESM-only — the CommonJS build was dropped.

## [0.1.0] — 2026-07-01

First release of the ground-up rewrite (previous npm releases under the unscoped `excelents` name
were the ExcelJS-era codebase). TypeScript-first, zero runtime dependencies, web-standard core
(Node 24+, browsers, edge).

### Added

- Object model: `createWorkbook`, worksheets, rows, columns, cells with typed values
  (string/number/boolean/date/formula/rich text/hyperlink) and styles
  (font/fill/border/alignment/number format).
- xlsx codecs: `writeXlsx` / `readXlsx` with merges, row/column sizing, defined names,
  data validation, and tables — conformance-tested against `exceljs@4.4.0` over real fixtures.
- CSV (RFC 4180): `writeCsv` / `readCsv` on the `./csv` entry.
- Streaming on the `./stream` entry: `writeXlsxStream`, `createXlsxStreamWriter`,
  `readXlsxRows`, `writeCsvStream`, `readCsvRows` — bounded memory over Web Streams.
- Node adapter on the `./node` entry: `nodeFileSystem`.
- npm trusted publishing (OIDC) with provenance.

[Unreleased]: https://github.com/mitresthen/excelents/compare/v0.3.0...HEAD
[0.3.0]: https://github.com/mitresthen/excelents/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/mitresthen/excelents/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/mitresthen/excelents/releases/tag/v0.1.0
