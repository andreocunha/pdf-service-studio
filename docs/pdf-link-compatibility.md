# Internal PDF links: known reader differences

Decision: the user explicitly chose **top alignment on Chrome/Edge and centered
destinations in Google Drive on Android**, after testing both alternatives.
Use explicit GoTo/XYZ with null horizontal position and null zoom (the tested
v2 encoding). Preserve the exact target page/y. This is an accepted compromise,
not a claim that every reader honors top alignment.

## Device observations

The user tested Chrome/Edge on desktop and Google Drive on a Samsung Galaxy S26.
The Drive app version was not recorded. These observations do not establish the
behavior of every Android app, iOS, or future Drive versions.

- Explicit GoTo actions made the previously broken mobile menu clickable.
- GoTo/FitH produced the preferred mobile framing, but Chrome/Edge did not honor
  the exact destination height.
- GoTo/XYZ corrected desktop positioning but moved the title to the middle of
  the screen in Drive. Changing the horizontal coordinate from 0 to null did
  not change Drive's centering. The user subsequently selected the XYZ behavior.
- macOS PDFKit tests did not predict the Drive behavior.

## Six-way diagnostic

A separate three-page PDF used six links to the same point, 794.88 PDF points
above the bottom of page 2 (page height 842.88). A third page avoided an end-of-
document scroll limit. The user reported the following outcomes and supplied
screenshots:

| Button | Encoding | Drive outcome |
| --- | --- | --- |
| A | GoTo with explicit FitH | Page 2 appears centered, putting its heading above screen center |
| B | GoTo with explicit XYZ, x=0, zoom=null | Destination appears near screen center |
| C | B followed by a GoTo/FitH action in Next | Same as B |
| D | GoTo with a string destination in Names/Dests, resolving to XYZ | Same as B |
| E | Direct annotation Dest containing XYZ | Same as B |
| F | Annotation Dest with a string destination in Names/Dests, resolving to XYZ | Same as B |

The screenshots support page centering for A and destination-point centering for
B-F: the page-2 bounds in A are approximately y=383..1335, while its green target
starts at y=437; the target in B-F starts at approximately y=858. Even A does
**not** put the heading at the screen top. The final choice was XYZ so desktop
positioning remains exact, accepting Drive's destination-point centering.

PDFium extracted the exact target page/y for B-F. That verifies destination
parsing, not the viewport behavior of the Drive application. The Next action in
C did not produce the desired fallback in this device test. All six variants
used valid, mutually exclusive A or Dest entries; do not combine both to exploit
different reader precedence.

## Evidence and limits

Chromium's link-click path uses FPDFDest_GetLocationInPage, whose PDFium
implementation extracts coordinates from XYZ destinations. The AndroidX viewer
source centers points passed by its link-click handler. AndroidX illustrates the
behavior but is not proof of which implementation/version the Drive app uses.

- [Chromium link handling](https://github.com/chromium/chromium/blob/main/pdf/pdfium/pdfium_page.cc)
- [PDFium destination parsing](https://github.com/chromium/pdfium/blob/main/core/fpdfdoc/cpdf_dest.cpp)
- [AndroidX viewport and link handling](https://github.com/androidx/androidx/blob/androidx-main/pdf/pdf-viewer/src/main/kotlin/androidx/pdf/view/PdfView.kt)

The unit tests check PDF objects, target coordinates, external links, and
idempotence. They cannot assert top alignment in native readers. A future fix
needs validation in both Chrome/Edge and the actual Drive Android viewer,
including multiple sections on one page, without changing the document layout.
No single encoding with top alignment in both readers has been verified so far;
the current XYZ behavior is the user's accepted result.
