---
"@betteroffice/docx": minor
"@betteroffice/docx-react": minor
---

DOCX plugins can opt into `reservations` to lay out inline boxes, blocks and hidden ranges, then draw in them through `getAnchorGeometry({ kind: 'reservation', id })`. Saves and paged exports leave them out.
