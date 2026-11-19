# Production Schedule Manager
Requires Node.js 18+.
    npm install
    npm run dev      # open the URL shown (usually http://localhost:5173)
Data is saved automatically to data/production-data.xlsx (and LocalStorage). Keep `npm run dev` running.

## Deploy on Vercel (same data on every device)
1. Push this folder to GitHub and import it in Vercel (Vite is detected automatically).
2. In the Vercel project: Storage / Marketplace -> add **Upstash Redis** and connect it to the project.
3. Settings -> Environment Variables -> add `APP_PASSWORD` (a long password you choose).
4. Deploy (or Redeploy after adding variables), then log in with that password on each device.
