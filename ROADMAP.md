# Toward a complete BME reference workspace

This iteration repairs indexing and delivers a local HTML workspace. The stages
below are planned, not implemented features. Prioritize them with feedback from
engineers using real document libraries.

## 1. Dependable library maintenance

- OCR for image-only PDFs with language selection, progress, cancellation, and
  explicit dependencies. Preserve originals and record extraction provenance.
- File identity and folder relocation to reconnect moved documents while retaining
  notes, favorites, and links; preview changes before applying them.
- Per-document extraction diagnostics, limits for unusually large files, scan
  cancellation, and streaming for very large libraries.
- Content-hash duplicate detection with a reviewable merge workflow.
- Validated backup restore in the interface with a rollback copy.

## 2. Complete the reference workflow

- Embedded PDF reader with outline, highlighting, multiple tabs, page bookmarks,
  and persistent reading position.
- Bring document linking and batch editing from the desktop viewer into HTML.
- Saved searches, equipment families, manufacturer aliases, and structured tags.
- Revision comparison and superseded-document relationships that retain history.
- Work packs grouping manuals, SOPs, notes, and test-equipment references for a
  service task, with export and document provenance.

## 3. Broader access

- Signed Windows installer, notarized macOS app, and Linux packages.
- Explicit x64/ARM64 distribution targets and real-device acceptance checks.
- Translation support, Arabic right-to-left layouts, and a full keyboard and
  screen-reader accessibility review.
- A licensed demonstration library and contribution guide for public references;
  do not redistribute copyrighted manuals without permission.

## 4. Teams and assisted discovery

- Optional shared libraries after authentication, authorization, encrypted
  transport, concurrent editing, audit history, and backup policies are in place.
- Optional question answering with exact document/page citations, local processing
  options, and explicit consent for any external service.
- Approved-source links for manufacturer publications and safety notices, with
  reviewable updates and jurisdiction-specific source selection.

## Current acceptance gaps

- Regression suites, portable builds, and packaged PDF/DOCX indexing/search checks
  pass on Windows, macOS, and Linux GitHub runners. Real-user hardware and broader
  architecture/distribution coverage still need acceptance testing.
- Browser interaction checks use clearly marked sample PDFs, not real equipment
  documentation. There is no formal load test or full accessibility audit yet.
- OCR, signed installers, team sharing, automatic file relocation, and the
  workflow features above remain future work.
- The original desktop viewer shares the repaired scanner; its remaining large
  GUI module needs a separate refactoring and interaction review.
