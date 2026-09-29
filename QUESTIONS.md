# Questions and confirmations

Items with no failing behaviour to reproduce. A one-line answer per item is enough; we will link your
answer from the matching ticket.

## Q1 · Surface we could not find a v2 shape for

We treat each of these as permanently unsupported unless you tell us otherwise:

| # | v1 surface | What we did instead |
| --- | --- | --- |
| a | `setCheckboxSymbolPair` (choose the checked / unchecked glyphs) | nothing; `checkbox.setState` only toggles |
| b | `setEditable(false)` / `presentationEditor` | read-only previews can only warn, not lock |
| c | a contenteditable / DOM input path (a DOM `Range` + keyboard events changes nothing) | all writes go through the Document API |
| d | `superdoc/super-editor` and a Node opener inside the browser package | headless work moved to `@superdoc/sdk` |
| e | `ctx.visuals.highlight` with phrase-level precision | our own overlay for "find this clause" |
| f | `annotations: true` config key | signing fields moved to content controls |

## Q2 · Bulk mutations in the browser (resolved on 2.16.0, informational)

On `superdoc@2.14.0`, ~500 sequential Document API calls (importing a template with 116 placeholders)
tripped `render.scheduler-degraded` ("current unpainted target reached 1017ms (limit 1000ms)"); the page
never painted again and later `replace()` calls failed with `target must be a SelectionTarget or body
StoryLocator`. On 2.16.0 the same import paints throughout, so we removed our workaround (a double
`requestAnimationFrame` yield after every mutation). Two questions remain: is `doc.plan.execute` the
recommended path for hundreds of mutations, and is the 1000 ms watchdog configurable?

## Q3 · `capabilities.resolve` as a planning tool

Across T03a, T06a and T08b, `capabilities.resolve` (and `capabilities.get`) answer `supported` for calls
that then refuse, and `unavailable` for a call that works. We would like to plan with it instead of
trying and catching. Is it meant to be authoritative per target (story, control kind, paragraph
content), or only per operation?
