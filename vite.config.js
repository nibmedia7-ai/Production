import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'fs'
import path from 'path'

// Local Excel storage: GET reads data/production-data.xlsx, POST writes it.
const FILE = path.resolve('data/production-data.xlsx')
const excelApi = () => ({
  name: 'excel-api',
  configureServer(server) {
    fs.mkdirSync(path.dirname(FILE), { recursive: true })
    server.middlewares.use('/api/excel', (req, res) => {
      const send = (code, obj) => { res.statusCode = code; res.setHeader('Content-Type', 'application/json'); res.end(JSON.stringify(obj)) }
      try {
        if (req.method === 'GET') {
          if (!fs.existsSync(FILE)) return send(200, { exists: false })
          const { mtimeMs } = fs.statSync(FILE)
          if (req.url.includes('meta')) return send(200, { exists: true, mtime: mtimeMs })
          return send(200, { exists: true, mtime: mtimeMs, data: fs.readFileSync(FILE).toString('base64') })
        }
        if (req.method === 'POST') {
          const chunks = []
          req.on('data', c => chunks.push(c))
          req.on('end', () => {
            try { fs.writeFileSync(FILE, Buffer.concat(chunks)); send(200, { mtime: fs.statSync(FILE).mtimeMs }) }
            catch (e) { send(500, { error: e.code || 'write failed' }) }
          })
          return
        }
        send(405, { error: 'method' })
      } catch (e) { send(500, { error: e.code || 'error' }) }
    })
  },
})

export default defineConfig({
  plugins: [react(), excelApi()],
  server: { watch: { ignored: ['**/data/**'] } },
})
