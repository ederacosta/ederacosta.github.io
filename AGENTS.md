# AGENTS.md

## Cursor Cloud specific instructions

This is a **static HTML/CSS/JS portfolio website** (no build step, no package manager, no backend). It uses the "Strata" template by HTML5 UP.

### Running the site

Serve the site locally with any static HTTP server from the repo root:

```bash
python3 -m http.server 8080
```

Then open `http://localhost:8080/` in a browser.

### Key notes

- There is **no build step**, no `package.json`, no dependency manager. All CSS and JS assets are pre-compiled and committed.
- SCSS source files live in `assets/sass/` but the compiled CSS (`assets/css/main.css`) is already committed, so a Sass compiler is only needed if editing `.scss` files.
- The site is hosted on GitHub Pages with custom domain `ederacosta.com` (configured via `CNAME` file).
- There are no automated tests, no linter configuration, and no CI/CD pipeline in the repo.
