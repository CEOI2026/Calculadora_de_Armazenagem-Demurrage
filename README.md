# Demurrage and Storage Calculator

Static browser app for calculating storage and detention charges for single containers and batch Excel uploads.

## Run

Open `index.html` in a browser, or serve the repository with any static file server.

## Structure

- `index.html`: UI markup and script loading
- `styles.css`: page styles and print styles
- `pricing-engine.js`: pure pricing rules for storage and detention
- `app.js`: DOM wiring, rendering, local storage persistence, Excel import/export, batch UI

## Notes

- Batch Excel features depend on SheetJS loaded from the CDN declared in `index.html`.
- PDF export is browser print flow, not server-side PDF generation.
- User configuration is stored in browser `localStorage`.
