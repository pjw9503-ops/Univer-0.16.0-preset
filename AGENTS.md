# AGENTS.md

## Cursor Cloud specific instructions

### Overview

This is a **Univer 0.16.0 Preset mode** example/documentation repository. It contains a Vite-powered client-side spreadsheet application using the `@univerjs/presets` and `@univerjs/preset-sheets-core` packages. No backend, database, or Docker is needed.

### Services

| Service | Command | Port |
|---------|---------|------|
| Vite dev server | `npm run dev` | 5173 |

### Key commands

- **Dev server**: `npm run dev` (Vite on port 5173)
- **Build**: `npm run build` (outputs to `dist/`)
- **Lint**: `npm run lint` (ESLint on `src/`)
- **Preview built app**: `npm run preview`

### Non-obvious notes

- The Univer library is entirely client-side; there are no external service dependencies.
- The `examples/` directory contains standalone JS snippet files (not importable modules). The runnable entry point is `src/main.js`.
- The Vite build produces a ~10 MB JS bundle due to the Univer engine. The chunk size warning is expected and can be ignored.
- Univer uses canvas-based rendering internally — Chrome is required for GUI testing (Firefox may have rendering issues).
