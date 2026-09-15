# Schedule Conflict Checker

A browser-only schedule conflict checker. The workbook is read locally in the browser; no schedule data is uploaded to a server.

## Develop locally

Install the development dependency once:

```sh
npm install
```

Start the local server:

```sh
npm run dev
```

Open the URL Vite prints (normally `http://localhost:5173`). The default workbook is served from `public/` and loads automatically.

## Test and build

```sh
npm test
npm run build
```

`npm run build` creates `dist/`, a self-contained static site ready for deployment. Preview that production build locally with:

```sh
npm run preview
```

## Deploy

Deploy the contents of `dist/` to any static host, such as Netlify, Vercel, Cloudflare Pages, or GitHub Pages.

Use these settings where the host asks for them:

| Setting | Value |
| --- | --- |
| Build command | `npm run build` |
| Publish/output directory | `dist` |
| Node version | 20 or newer |

The included default `.xlsm` workbook is a public static asset at `public/CESS CAIE Schedule Fall 2026.xlsm`, and is copied to the root of `dist/` during the build. Anyone who can open the deployed app can download that workbook, so remove it from `public/` if it should not be publicly accessible. Users can still select a workbook manually from their own computer.
