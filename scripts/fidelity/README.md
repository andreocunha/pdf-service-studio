# Editor/PDF fidelity regression

The fixtures run the real `NotionEditor` (editable) and `PdfRenderClient`
(readonly export) against **the same immutable snapshot**. They do not save
changes to documents. Keep snapshots and rendered documents outside Git.

Each snapshot is `UUID.json` containing `id`, `workspaceId`, `title`, `blocks`,
`meta`, `effectiveDocumentId`, `rawClauses`, and raw `design_block_templates`
rows in `templates`. Reconstruct blocks with the app's `deserializeDocument`,
including every pending update. Mirror the export route's library resolution:
include all templates owned by the document and its source (including navigation
chrome resolved by name), plus templates referenced by blocks and clauses.
Paginate database reads; do not truncate them to 1,000 records. Preserve the
export route's workspace access boundaries: including a referenced template
from another workspace can make a fixture falsely pass while the real export
correctly cannot access that asset. Templates readable through anonymous RLS can be listed in `publicTemplateIds`
after a real public-client lookup; never infer that list from the admin lookup.
Missing/private inaccessible templates must fail export.

Prepare a disposable copy of the app, then start it on an unused port:

```sh
node scripts/fidelity/prepare-fixture-app.mjs ../lex-studio-v2 /tmp/pdf-fidelity-app
cd /tmp/pdf-fidelity-app
FIDELITY_FIXTURE_DIR=/tmp/pdf-fixtures npx next dev --hostname localhost --port 3002
```

From pdf-service, for each fixture UUID:

```sh
FIDELITY_BASE_URL=http://localhost:3002 FIDELITY_OUTPUT_DIR=/tmp/pdf-results npm run test:pdf-fidelity -- UUID
FIDELITY_OUTPUT_DIR=/tmp/pdf-results python3 scripts/fidelity/compare-layout.py UUID
python3 scripts/check-pdf-regression.py /tmp/pdf-results/UUID-before.pdf /tmp/pdf-results/UUID-after.pdf
npm run test:pdf-font-failures -- 'http://localhost:3002/render-pdf/validation/UUID?mode=export'
```

The audit starts fresh browser contexts and delays the font catalog by 1.5s.
Set `FIDELITY_FONT_DELAY_MS=0` to cover the ordinary loading order too.
Always also exercise `renderDocumentPdf` against the authenticated route: a
fixture alone does not validate server-side library fetching.
It fails on resource errors or unfinished layout. It records every block, visible image, navigation target and rendered text run,
including fonts, wrapping and page-relative rectangles. Image readiness is checked
separately. Set `FIDELITY_SCREENSHOTS=1` to capture each page.
The comparator requires identical pages, block order, text and font styles,
and coordinates within 0.001 CSS px. The before/after PDFs separately verify
that mask conversion cannot move content or replace fonts/links.

The failure test blocks the catalog, font metrics manifest and original font
files in turn. Each must report `__PDF_ERROR` without `__PDF_READY`. Pass optional resource URL
substrings after the URL to test image failures too. Use a fixture without external
images for the font tests, so an unrelated network failure cannot mask the result.

External image providers can rate-limit repeated cold runs. Honor `Retry-After`
and rerun the failed case; do not intercept missing images with substitutes or
accept a partially loaded document. Run editor/export modes separately if needed.
A numerical link-rectangle tolerance of `--link-coordinate-tolerance 0.0001`
PDF points is allowed for Chromium float serialization; destinations must stay exact.

Use `CHROMIUM_EXECUTABLE_PATH` for a specific browser. Mac defaults to installed
Chrome; Linux defaults to the bundled Chromium. The same tests should run in
the deployment image before release. Generated PDFs here are **raw renders**;
use the normal compression pipeline before comparing final download size and
review those final files in Quartz/Preview and Poppler/Chrome. System-only
fonts/emoji, soft masks and exotic CSS remain separate compatibility cases;
passing these fixtures is not a universal PDF viewer guarantee.

The real-route audit uses the service configuration from `.env`. Override
`RENDER_BASE_URL` to the disposable app and set `CHROMIUM_EXECUTABLE_PATH`
when running locally. It does not save document changes:

```sh
FIDELITY_FIXTURE_DIR=/tmp/pdf-fixtures FIDELITY_OUTPUT_DIR=/tmp/pdf-results RENDER_BASE_URL=http://localhost:3002 node --env-file=.env --import tsx scripts/fidelity/audit-live.mts UUID
node --import tsx scripts/fidelity/test-late-image-failure.mts 'http://localhost:3002/render-pdf/validation/UUID?mode=export'
```

Compare `UUID-live.pdf` against `UUID-after.pdf` using the regression checker.
For the missing-template guard, copy a fixture under a new UUID, remove its
referenced templates, and pass `missing-template` as the failure-test resource.
