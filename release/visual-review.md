# Community-renderer visual baseline review

Reviewed on 2026-09-03 against the locally available pinned community source
`dc4d7e5927864498913b7ba42d0da06cc7cf628e` and the five existing 720 × 900
split-view PNGs. The manifest records their exact bytes and permitted source paths.

These are inherited **host-adapted community renderer regression baselines**, not
screenshots of an unmodified upstream Electron application. Their original capture
time and browser profile were not retained; this review does not invent those facts.
The new evidence bundle records the current packaged host and browser profile.
No enterprise source or prohibited application shell was used for this review.

Observed baseline surfaces:

- DOCX: native ribbon, blank paginated document, zoom and status controls.
- Markdown: native ribbon, empty editable document and status strip.
- XLSX: native ribbon, formula bar, Sheet1 grid, A1 selection and sheet/status tabs.
- PPTX: native ribbon, slide thumbnail, blank slide canvas and zoom controls.
- PDF: native ribbon, page thumbnail and rendered PDF fixture. Its inherited fixture
  text says “GenOffice Lite PDF host gate”; the current fixture uses “TandemFolio”.

The source paths were checked at the pinned commit. Only the permitted renderer
areas were inspected. Current host additions and removed upstream features are
described by the ADRs and `docs/migration/provenance.md`; the manifest does not
claim pixel identity to the original application. All masks remain empty and the
existing 3% pixel-difference ceiling is unchanged. A new capture must still pass
the visual and runtime gates; restoring these references alone is not approval.

The current five images in `release/artifacts/` were also visually inspected after
the approved 2026-09-04T02:51:06.592Z capture. The document surfaces, native ribbons, sheet/slide
and PDF thumbnails remain present. Visible differences include the accepted host
file-location status control and the renamed
PDF fixture. No current screenshot was substituted for an inherited reference.
All comparisons are within 3%, with no masks or substituted references. The
source-current performance and release gates also pass. See [verification results](validation.md).
