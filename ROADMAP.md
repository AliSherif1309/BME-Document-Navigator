# Toward a complete BME reference workspace

The local HTML workspace now includes reliable indexing, saved searches, batch
metadata previews, verified folder reconnection, an in-app PDF reader, page
bookmarks, saved reading positions, related-document links, and optional local OCR
with language selection, progress, cancellation, and recognition provenance. Bulk operations
create a backup, reject stale previews, and preserve document history. The stages below
are future work. Prioritize them with feedback from engineers using real libraries.

## 1. Dependable library maintenance

- Extend OCR with standalone-image support, deskewing, a persistent document
  queue, searchable-PDF export, and representative multilingual acceptance tests.
- Extend folder reconnection to individual file renames and reorganized layouts;
  offer reviewable matches when identical content appears at a different path.
- Per-document extraction diagnostics, limits for unusually large files, scan
  cancellation, and streaming for very large libraries.
- Content-hash duplicate detection with a reviewable merge workflow.
- Validated backup restore in the interface with a rollback copy.

## 2. Complete the reference workflow

- Multiple reader tabs, back/forward history, and text selection directly on the
  rendered page. Outline navigation, highlighting, bookmarks, and saved reading
  position are already available in the single-document reader.
- Equipment families, manufacturer aliases, and structured tags.
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
- Signed installers, team sharing, automatic file matching, and the
  workflow features above remain future work.
- The original desktop viewer shares the repaired scanner; its remaining large
  GUI module needs a separate refactoring and interaction review.
