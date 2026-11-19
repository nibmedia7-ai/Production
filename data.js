// Vercel serverless function: stores all productions in Upstash Redis (shared across devices).
import { timingSafeEqual } from 'crypto'

const URL_ = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL
const TOKEN = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN

const redis = async cmds => {
  const r = await fetch(URL_ + '/pipeline', { method: 'POST', headers: { Authorization: 'Bearer ' + TOKEN, 'Content-Type': 'application/json' }, body: JSON.stringify(cmds) })
  if (!r.ok) throw new Error('redis')
  return r.json()
}
const same = (a, b) => { const x = Buffer.from(String(a)), y = Buffer.from(String(b)); return x.length === y.length && timingSafeEqual(x, y) }

export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store')
  const pw = process.env.APP_PASSWORD
  if (!pw || !URL_ || !TOKEN) return res.status(500).json({ error: 'Server not configured: set APP_PASSWORD and connect Upstash Redis in Vercel.' })
  if (!same(req.headers['x-app-password'] || '', pw)) return res.status(401).json({ error: 'Wrong password' })
  try {
    if (req.method === 'GET') {
      if (req.query.meta) { const [r] = await redis([['GET', 'psm:rev']]); return res.json({ rev: r.result ? Number(r.result) : 0 }) }
      const [d, r] = await redis([['GET', 'psm:data'], ['GET', 'psm:rev']])
      return res.json({ rev: r.result ? Number(r.result) : 0, data: d.result ? JSON.parse(d.result) : null })
    }
    if (req.method === 'PUT') {
      const data = req.body && req.body.data
      if (!Array.isArray(data)) return res.status(400).json({ error: 'Invalid data' })
      const rev = Date.now()
      await redis([['SET', 'psm:data', JSON.stringify(data)], ['SET', 'psm:rev', String(rev)]])
      return res.json({ rev })
    }
    return res.status(405).json({ error: 'Method not allowed' })
  } catch {
    return res.status(500).json({ error: 'Storage error' })
  }
}
