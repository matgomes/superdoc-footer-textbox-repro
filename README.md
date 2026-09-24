# Loading overlay stuck at 96%: `wrap="none"` text box inside a footer table cell

## Summary

A DOCX whose footer contains a **text box with `<wps:bodyPr wrap="none">`**, anchored in a **table cell**, never finishes loading. The "Loading document / Almost ready… 96%" overlay stays on screen indefinitely and covers the editor.

- The footer pre-layout throws `remeasureParagraph: maxWidth must be a positive number, got Infinity`.
- `onException` then fires with `page furniture (headers/footers) could not be resolved`.
- `onReady` **does** fire, so the host has no signal that anything is wrong. Only the overlay never goes away.

We hit this in production with customer contract templates, where every footer holding a logo text box in a table cell fails this way. Real documents can have a dozen such footers, one per section.

## Reproduce

```bash
npm install
npm run dev
# open http://localhost:5174/?doc=A-textbox-nowrap-in-pct-table.docx
```

The page header shows `onReady`, the overlay state and any exceptions. Use the dropdown to switch fixtures.

## Fixtures

Each fixture is one page with one footer. They differ only in the footer. `fixtures/build_fixtures.py` regenerates them into `public/` (Python standard library only).

| Fixture | Footer | Result |
|---|---|---|
| **A** | `wrap="none"` text box in a table cell; table width `5000 pct` | ❌ overlay stuck at 96% |
| **C** | `wrap="none"` text box in a table cell; table width fixed `dxa` | ❌ overlay stuck at 96% |
| **B** | same `wrap="none"` text box, **no table** | ✅ loads |
| **D** | text box with `wrap="square"` in a table cell | ✅ loads |

So the trigger is `wrap="none"` combined with a table cell. The table's width type doesn't matter.

The text box is the standard Word shape: `wp:anchor` with `layoutInCell="1"`, `wps:wsp` holding `wps:txbx`, and `wps:bodyPr wrap="none"` with `a:spAutoFit`. That is what Word writes for an auto-fitting label or logo box.

## Versions tested

All with `@superdoc/docx-engine` as pinned by each release:

| superdoc | docx-engine | A | C |
|---|---|---|---|
| 2.14.0 | 0.13.0 | ❌ | ❌ |
| 2.16.0 | 0.15.0 | ❌ | ❌ |
| 2.17.0 | 0.16.0 | ❌ | ❌ |
| 2.18.0-next.8 | 0.17.0-next.9 | ❌ | ❌ |

B and D load correctly on every version. The same result holds with and without `v2Collaboration`.

## Console output (fixture A)

```
[Layout] Footer pre-layout failed: Error: remeasureParagraph: maxWidth must be a positive number, got Infinity
[v2-page-furniture] Failed to layout footer rId=rId1: Error: remeasureParagraph: maxWidth must be a positive number, got Infinity
```

## Expected

- The footer lays out. A `wrap="none"` text box should size to its content, bounded by the page or cell.
- Failing that, the loading overlay should be dismissed once the document itself is ready, with only the footer degraded.

## Likely cause (from the outside)

`wrap="none"` means the text box's paragraph width is unconstrained. Inside a table cell, that unconstrained width seems to reach `remeasureParagraph` as `Infinity` instead of being resolved from the cell or page. The failure in page-furniture layout then blocks the overlay's final step, even though the body has laid out and `onReady` has fired.

## Screenshots

- `screenshot-A-textbox-nowrap-in-pct-table.png`: the stuck overlay
- `screenshot-B-textbox-nowrap-no-table.png`: the control, loading correctly
