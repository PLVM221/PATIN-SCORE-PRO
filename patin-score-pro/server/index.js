import { createServer } from 'node:http'
import { readFileSync, existsSync, mkdirSync } from 'node:fs'
import { extname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { DatabaseSync } from 'node:sqlite'

const __dirname = fileURLToPath(new URL('.', import.meta.url))
const rootDir = resolve(__dirname, '..')
const dataDir = join(__dirname, 'data')
const distDir = join(rootDir, 'dist')
const dbPath = join(dataDir, 'patin-score-pro.sqlite')
const port = Number(process.env.PORT || 4174)

mkdirSync(dataDir, { recursive: true })

const db = new DatabaseSync(dbPath)
db.exec(readFileSync(join(__dirname, 'schema.sql'), 'utf8'))

const getState = db.prepare('SELECT data_json, updated_at FROM app_state WHERE id = 1')
const saveState = db.prepare(`
  INSERT INTO app_state (id, data_json, updated_at)
  VALUES (1, ?, CURRENT_TIMESTAMP)
  ON CONFLICT(id) DO UPDATE SET data_json = excluded.data_json, updated_at = CURRENT_TIMESTAMP
`)
const audit = db.prepare('INSERT INTO audit_log (action) VALUES (?)')

function json(res, status, body) {
  const payload = JSON.stringify(body)
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(payload),
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET,PUT,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  })
  res.end(payload)
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = ''
    req.on('data', (chunk) => {
      body += chunk
      if (body.length > 80 * 1024 * 1024) {
        reject(new Error('Payload demasiado grande'))
        req.destroy()
      }
    })
    req.on('end', () => resolve(body))
    req.on('error', reject)
  })
}

function serveStatic(req, res) {
  const url = new URL(req.url, `http://${req.headers.host}`)
  const requested = url.pathname === '/' ? '/index.html' : url.pathname
  const filePath = resolve(distDir, `.${requested}`)
  const target = filePath.startsWith(distDir) && existsSync(filePath)
    ? filePath
    : join(distDir, 'index.html')

  if (!existsSync(target)) {
    json(res, 404, { ok: false, error: 'Primero ejecutá npm run build para generar dist.' })
    return
  }

  const types = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.svg': 'image/svg+xml',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.ico': 'image/x-icon',
  }
  const bytes = readFileSync(target)
  res.writeHead(200, { 'Content-Type': types[extname(target)] || 'application/octet-stream' })
  res.end(bytes)
}

const server = createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`)

    if (req.method === 'OPTIONS') {
      json(res, 204, {})
      return
    }

    if (url.pathname === '/api/health') {
      json(res, 200, { ok: true, db: dbPath })
      return
    }

    if (url.pathname === '/api/state' && req.method === 'GET') {
      const row = getState.get()
      json(res, 200, {
        ok: true,
        data: row ? JSON.parse(row.data_json) : null,
        updatedAt: row?.updated_at || null,
      })
      return
    }

    if (url.pathname === '/api/state' && req.method === 'PUT') {
      const body = JSON.parse(await readBody(req))
      if (!body || typeof body !== 'object') {
        json(res, 400, { ok: false, error: 'Estado inválido' })
        return
      }
      saveState.run(JSON.stringify(body))
      audit.run('state_saved')
      json(res, 200, { ok: true })
      return
    }

    if (url.pathname === '/api/audit' && req.method === 'GET') {
      const rows = db.prepare('SELECT id, action, created_at FROM audit_log ORDER BY id DESC LIMIT 100').all()
      json(res, 200, { ok: true, rows })
      return
    }

    serveStatic(req, res)
  } catch (error) {
    json(res, 500, { ok: false, error: error.message })
  }
})

server.listen(port, '0.0.0.0', () => {
  console.log(`Patin Score Pro SQL listo en http://localhost:${port}`)
  console.log(`SQLite: ${dbPath}`)
})
