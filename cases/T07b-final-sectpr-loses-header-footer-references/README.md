# T07b — Header/footer references of every `sectPr` after an unchanged open + save

**Ticket:** T07 · Package tolerance and export fidelity · **Type:** Bug (not reproduced on current versions; kept as a regression check)

## Why we do this
Our contracts use two sections with a distinct first page (`titlePg`) and different headers and footers per section. An earlier comparison suggested that the final, body-level `<w:sectPr>` lost its default header and footer references after a save. We wrote this check to confirm that before reporting it.

## What the script does
1. Builds two synthetic two-section documents. Section 1 has `titlePg` plus first and default header and footer references. The final section has its own default header and footer. In one variant the final section has body content; in the other it owns no paragraphs, which is the shape of our real template.
2. Also runs on `fixtures/sanitized/po-locked-footer.docx` when it is present.
3. Opens each document, changes nothing, and saves.
4. Compares every `sectPr`, before and after, by reference type and by the part each reference points at. `r:id` values may be renumbered, so they are not compared directly.

```bash
node cases/T07b-final-sectpr-loses-header-footer-references/repro.mjs 2.15.0
```

## Expected
Every `sectPr` keeps the same header and footer references and `titlePg`.

## Actual
Nothing is lost on the versions we ran: the script prints `FIXED?`. The earlier observation compared SuperDoc 2.x output with our old 1.x output, not with the input. On the input itself, the final `sectPr` of our real template carries no references at all, and the save preserves that. We do not consider this a SuperDoc defect. See [RESULTS.md](../../RESULTS.md).

## Done when
It already passes. The case stays in the kit as a regression check for multi-section templates.

## Our workaround today
None needed.

## Question for SuperDoc
None. This is listed only so you can see we checked.
