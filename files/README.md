# Client Master — ADAS Globus

A multi-entity client management application (entities, individuals, relationships,
services, payroll, deadlines, audit log, role-based access) built as a single
React component and packaged as a deployable Vite app.

## Quick start

```bash
npm install      # install dependencies
npm run dev      # start dev server at http://localhost:5173
npm run build    # production build into dist/
npm run preview  # preview the production build locally
```

## Project structure

```
adas-client-master/
├── index.html                 # Vite HTML entry
├── package.json
├── vite.config.js             # Vite + @vitejs/plugin-react
├── tailwind.config.js         # Tailwind v3 content paths
├── postcss.config.js          # Tailwind + autoprefixer
├── vercel.json                # SPA rewrite (all routes -> /)
├── .gitignore
├── README.md
└── src/
    ├── main.jsx               # React entry; installs storage polyfill, mounts App
    ├── App.jsx                # Renders <ClientMaster />
    ├── index.css              # Tailwind directives + base styles
    ├── lib/
    │   └── storage.js         # window.storage polyfill (localStorage-backed)
    └── components/
        └── ClientMaster.jsx   # Your original component, unchanged
```

## Tech stack

- React 18
- Vite 5
- Tailwind CSS 3 (PostCSS + autoprefixer)
- lucide-react (icons)
- xlsx / SheetJS (CSV & Excel import/export)

## About the storage polyfill

`ClientMaster.jsx` persists data through an async `window.storage` key-value API
(`get` / `set` / `delete`). That API is provided by the original host environment
but does **not** exist in a normal browser. `src/lib/storage.js` supplies a drop-in
replacement backed by `localStorage`, installed in `main.jsx` before the app mounts,
so all save/load behaviour works on Vercel exactly as before. Data persists per
browser; it does not sync across devices. Use the built-in **Export JSON (full
backup)** feature for portable backups.

The component file itself was copied in verbatim — no logic was changed.

## Deploy to Vercel

1. Push this folder to a Git repository (GitHub/GitLab/Bitbucket).
2. In Vercel, **Add New → Project** and import the repo.
3. Vercel auto-detects Vite. Defaults are correct:
   - Build command: `npm run build`
   - Output directory: `dist`
   - Install command: `npm install`
4. Deploy.

Or via CLI:

```bash
npm i -g vercel
vercel        # preview deploy
vercel --prod # production deploy
```

## Notes

- The build prints a "chunk larger than 500 kB" advisory — this is informational
  (the `xlsx` library is large) and does not affect correctness. Code-splitting can
  be added later if desired.
- `npm install` may report advisories for the `xlsx` package; these come from the
  upstream library version published to npm and do not affect the build.
