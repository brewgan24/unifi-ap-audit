# UniFi AP Audit

A local-first daily Wi-Fi access point audit workspace. It accepts tab-separated rows copied from UniFi, screenshots from Ookla Speedtest, and produces a printable report for all locations.

## Run locally

```bash
npm install
npm run dev
```

The audit data is stored in the browser's local storage. No UniFi credentials or client data are sent to a server.

## Deploy to Vercel

```bash
npm install -g vercel
vercel login
vercel --prod
```

When asked for the project settings, keep the detected Vite defaults. The `vercel.json` rewrite keeps the single-page app working on direct routes.
