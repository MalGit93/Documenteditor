# Executive Brief Designer

A single-page, client-side editor for producing themed executive briefs. It supports block-based editing, quality checks, export to DOCX, and now standalone HTML files that retain the designed layout.

## Running locally

Open `Maineditor (4) (1).html` in any modern browser. No build step is required, but you can also serve the directory via a lightweight server, e.g.

```bash
python3 -m http.server 8000
```

## Exporting deliverables

- **DOCX** – Uses a custom generator so you can download a Word file that mirrors the designer layout.
- **HTML** – Creates a self-contained, read-only HTML document with inline styles and lightweight footnote navigation for sharing interactive reports without the editor chrome.
- **JSON** – Save or load your working state for later editing.

> The previous editable PDF option was removed because html-to-PDF rendering could not preserve the strict A4 layout. Use the print dialog or the standalone HTML export when a PDF is still required.

## Embedded content

Use the new **Embed** block (Insert → Embed) to paste iframe-based visualisations, videos, or dashboards. The editor sanitises the snippet, keeps it read-only inside the canvas, and ensures each embed carries a caption so recipients know what they are viewing. Embedded content is preserved in DOCX/HTML exports (DOCX receives a descriptive placeholder, while HTML keeps the live embed).

## Accessibility & quality checks

The “Check brief” button scans for placeholder text, missing alt text on images, incomplete embed captions, and unresolved footnote markers. Fixing the reported items before exporting keeps the shared documents production-ready.
