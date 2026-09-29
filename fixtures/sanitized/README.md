# Sanitized real templates

These five packages are real contract templates from our customers' workflow, passed through
`tools/sanitize.py`. They keep every structure the cases depend on and none of the content.

| File | What it is | Used by |
| --- | --- | --- |
| `po-locked-footer.docx` | Purchase order exported by a Google-Docs-based authoring flow: two sections with `titlePg`, footers built as tables, the first-page footer wrapped in a `contentLocked` block SDT (`goog_rdk_0`), legacy placeholders `{{scope_contract_number}}`, `[seller_initials]`, `[buyer_initials]` in the footers | T03b, T07b |
| `work-order-footer-initials.docx` | Work order whose footer reads `Subcontractor Initials: [seller_initials}` · tab · `Contractor Initials: [buyer_initials]` in one paragraph (the `}` typo is in the real template) | T01c, T06b |
| `po-header-numbering.docx` | Purchase order exported by SuperDoc 1.45.x: `word/numbering.xml` uses `w15:` / `w16cid:` attributes without declaring the prefixes; 13 body placeholders | T07a |
| `paragraph-116-codes.docx` | A single real paragraph with 110 `{{placeholders}}` separated by `<w:br/>` (with a leading break), as authored in our field-catalog template | T01d |
| `redline-01.docx` | A purchase order with six tracked changes (one run-property change), two authors | T09a |

What the sanitizer changes: every word outside a placeholder token is replaced by a pseudo-word of the
same length and case (so character offsets and run boundaries are unchanged), digits are remapped,
authors become "Author A/B", document properties and custom XML are blanked, external links point to
example.com, embedded fonts and the thumbnail are removed, images become neutral placeholders of the
same size. Paragraph ids (`w14:paraId`), SDT tags/ids/locks, fields, numbering, sections and tracked
changes are untouched. `SHA256SUMS` lists the exact bytes.
