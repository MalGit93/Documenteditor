# Executive Brief Designer

A single-page, client-side editor for producing themed executive briefs. It supports block-based editing, quality checks, export to DOCX, and now standalone HTML files that retain the designed layout.

## Running locally

Open `index.html` in any modern browser. No build step is required, but you can also serve the directory via a lightweight server, e.g.

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

## Publishing on GitHub Pages

GitHub Pages (or any static host) will automatically serve `index.html` from the repository root, so the editor works out of the box once you push the repo to GitHub and enable Pages. If you prefer to keep the project in a subfolder (for example `/docs`), move `index.html`, `styles.css`, and `app.js` into that folder, update any asset paths if needed, and point GitHub Pages at it. Because everything is client-side, no additional build configuration is required—just make sure the published branch/folder contains `index.html` so the live site renders the designer instead of the README preview.

You can upload any additional HTML (for example `/charts/q1-scenario.html` or AI-generated dashboards) alongside the editor. Once Pages publishes the files:

1. Drop the exported HTML in the repository, keeping the path relative to the published folder (e.g. `charts/q1-scenario.html`).
2. Open the designer, click **Share link**, and use the new **Standalone HTML** section to paste that relative path.
3. Copy the published URL or iframe snippet that appears immediately in the modal—perfect for intranet embeds, wiki documentation, or customer updates that reuse the hosted dashboards.

The helper still exposes the deployed editor URL for the main experience, and it reminds you that anything under `/charts` can be embedded back into the brief with relative URLs once Pages republishes the changes.

## Accessibility & quality checks

The “Check brief” button scans for placeholder text, missing alt text on images, incomplete embed captions, and unresolved footnote markers. Fixing the reported items before exporting keeps the shared documents production-ready.
