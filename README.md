# Concept Map Builder

Concept Map Builder is a local-first web app for creating well-formatted clinical concept maps from case-log PDFs. It extracts the best available information from the uploaded case log, fills the concept-map fields, flags anything that needs review, and exports a formatted PDF layered onto the original concept-map template.

The output depends on the quality of the case log. Clearer case logs produce better filled maps, and every field remains editable so users can correct, complete, or replace anything before export.

## What It Does

- Upload a case-log PDF.
- Parse key clinical details in the browser.
- Auto-fill concept-map fields.
- Flag generated or uncertain values in the review form.
- Let users manually edit every field, including nursing diagnosis content.
- Require Nursing Diagnosis completion before concept-map PDF export.
- Export a formatted concept-map PDF.
- Export with nursing priority sections appended in the PDF.

## Privacy Note

This app is currently local-first. PDF text extraction and form filling run in the browser, and no backend service is required for the core workflow.

Avoid uploading protected health information unless the deployment and access controls are appropriate for your use case.

## Run Locally

```bash
npm install
npm run dev
```

Then open the local URL shown in the terminal, usually:

```text
http://localhost:5173
```

## Build

```bash
npm run build
```

The production build uses relative asset paths so the app can be hosted from a root domain or from a GitHub Pages project path.

## GitHub Pages

The included GitHub Actions workflow builds and deploys the app to GitHub Pages on every push to `main`. In the GitHub repository, set Pages to use GitHub Actions as the deployment source.

## Project Structure

- `src/App.jsx` - main React app, parser, review UI, and PDF export logic
- `index.html` - app shell and styling
- `vite.config.js` - Vite/React build configuration
- `public/templates/` - concept-map PDF and related templates
- `electron/` - early desktop wrapper files, not required for the web app

## Current Status

This is a working web app ready for testing and pilot use. Nursing diagnosis content is user-authored in the form and required for concept-map PDF export.
