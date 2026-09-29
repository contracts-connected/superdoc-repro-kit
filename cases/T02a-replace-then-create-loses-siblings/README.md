# T02a — `replace()` then `create.contentControl` on the range `replace()` reported loses sibling text

**Ticket:** T02 · range reported by `replace()` / coordinate space · **Type:** Bug (silent data loss) on 2.13.0; refusal from 2.15.0 on

## Why we do this
Our first port wrote the value and then wrapped it: `replace()` the placeholder with its value, then
`create.contentControl` on the range the `replace()` receipt reports. That is the natural sequence, and on 2.13.0 it
silently deletes text around the placeholder. We now create first and fill afterwards, which is clean. We report it
because nothing warns a developer who writes the natural sequence.

## What the script does
Paragraph `00000001`: `Retention: {{retainage_note}} of the sum.`, token at offsets 11..29.

| Fixture | Runs |
| --- | --- |
| `one-run-sentence.docx` | one run for the whole sentence |
| `three-run-sentence.docx` | bold `Retention:` · ` {{retainage_note}} ` · `of the sum.` |

**replace-first** (fresh open):
1. `doc.replace({ target: <11..29>, text: 'Retainage Note' })`
2. Range to wrap: `receipt.effects.insertedText[0].selectionTarget`, else `receipt.resolution.selectionTarget`
   (if the receipt carries neither, the script re-locates the text with `doc.query.match`)
3. `doc.create.contentControl({ kind: 'inline', controlType: 'text', tag: 'retainage_note', alias: 'Retainage Note', at: <that range> })`

**create-first** (control, fresh open):
1. `doc.create.contentControl({ …, at: <11..29> })` on the untouched document
2. `doc.contentControls.selectByTag({ tag: 'retainage_note' })` → `doc.contentControls.text.setValue({ target, value: 'Retainage Note' })`

Both save, and the script reads the paragraph text, run count and the bold label from `word/document.xml`.

## Expected
Both sequences export `Retention: Retainage Note of the sum.` with the bold label intact.

## Actual
- **2.13.0:** replace-first returns `success:true` for both calls. One run: the paragraph is reduced to `Retainage Note`,
  and `Retention: ` and ` of the sum.` are gone. Three runs: `Retention:Retainage Noteof the sum.` (the spaces inside the
  token's run are gone). create-first is clean.
- **2.15.0 and 2.16.0-next.17:** `replace()` succeeds and leaves the run as three `<w:t>` nodes; `create.contentControl` on
  the reported range throws `CAPABILITY_UNAVAILABLE Content-control selection spans unsupported inline structures.`.
  Nothing is lost. create-first is clean.

See RESULTS.md for every version and platform.

## Done when
The script prints `FIXED?` (no text lost). It does on 2.15.0+.

## Our workaround today
Create first on the untouched range, then `text.setValue`. We never wrap a range that `replace()` just produced.

## Question for SuperDoc
Is a range from the `replace()` receipt meant to be reusable as a `create.contentControl` target? If not, a note in the
documentation would save the next team from this.
