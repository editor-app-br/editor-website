#!/usr/bin/env node
/**
 * Pack Next static export (./out) into JusIntegra Workspace editor-bundle slot layout.
 * Usage: node scripts/pack-workspace-editor-slot.mjs [slotDir]
 */
import fs from 'node:fs'
import path from 'node:path'
import zlib from 'node:zlib'
import { fileURLToPath } from 'node:url'

const UPSTREAM = 'https://editor.app.br'
const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..')
const outDir = path.join(root, 'out')
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'))
const version = String(pkg.version || '0.0.0')
const defaultSlot = path.join(
  process.env.HOME || '',
  'Library/Application Support/com.workspace.platform/editor-bundle/slots',
  version
)
const slotRoot = path.resolve(process.argv[2] || defaultSlot)

function isWasmMagic(buf) {
  return buf.length >= 4 && buf[0] === 0 && buf[1] === 0x61 && buf[2] === 0x73 && buf[3] === 0x6d
}

function prepareBody(urlPath, body) {
  if (urlPath !== '/x2t/x2t.wasm' || isWasmMagic(body)) return body
  try {
    const out = zlib.brotliDecompressSync(body)
    if (isWasmMagic(out)) return out
  } catch {
    /* keep original */
  }
  return body
}

function contentTypeFor(urlPath) {
  if (urlPath === '/embed' || urlPath.startsWith('/embed/')) return 'text/html; charset=utf-8'
  if (urlPath.endsWith('.wasm')) return 'application/wasm'
  if (urlPath.endsWith('.js')) return 'application/javascript'
  if (urlPath.endsWith('.css')) return 'text/css'
  if (urlPath.endsWith('.html')) return 'text/html'
  if (urlPath.endsWith('.json')) return 'application/json'
  if (urlPath.endsWith('.svg')) return 'image/svg+xml'
  if (urlPath.endsWith('.png')) return 'image/png'
  if (urlPath.endsWith('.woff2')) return 'font/woff2'
  return 'application/octet-stream'
}

function cachePaths(slot, urlPath) {
  let dir = slot
  const trimmed = urlPath.replace(/^\//, '')
  if (!trimmed) {
    throw new Error(`invalid path ${urlPath}`)
  }
  for (const part of trimmed.split('/')) {
    dir = path.join(dir, part)
  }
  return { body: path.join(dir, 'body'), meta: path.join(dir, 'meta.json') }
}

function writeBody(urlPath, rawBody) {
  const body = prepareBody(urlPath, rawBody)
  const entry = cachePaths(slotRoot, urlPath)
  fs.mkdirSync(path.dirname(entry.body), { recursive: true })
  fs.writeFileSync(entry.body, body)
  fs.writeFileSync(
    entry.meta,
    `${JSON.stringify({ status: 200, content_type: contentTypeFor(urlPath) }, null, 2)}\n`
  )
}

function outFileForUrl(urlPath) {
  if (urlPath === '/embed') {
    const flat = path.join(outDir, 'embed.html')
    if (fs.existsSync(flat)) return flat
    return path.join(outDir, 'embed', 'index.html')
  }
  const rel = urlPath.replace(/^\//, '')
  const inOut = path.join(outDir, rel)
  if (fs.existsSync(inOut)) return inOut
  return path.join(root, 'public', rel)
}

function collectNextStatic(html) {
  const found = []
  for (const m of html.matchAll(/\/_next\/static\/[A-Za-z0-9_./-]+/g)) {
    let p = m[0].replace(/[),.;]+$/, '')
    if (!p.includes('/chunks/') && !p.endsWith('.css') && !p.endsWith('.js')) continue
    if (!found.includes(p)) found.push(p)
  }
  return found
}

async function loadBytes(urlPath) {
  const filePath = outFileForUrl(urlPath)
  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    return fs.readFileSync(filePath)
  }
  const res = await fetch(`${UPSTREAM}${urlPath}`, { redirect: 'follow' })
  if (!res.ok) {
    throw new Error(`upstream ${res.status} ${urlPath}`)
  }
  return Buffer.from(await res.arrayBuffer())
}

if (!fs.existsSync(outDir) || !fs.statSync(outDir).isDirectory()) {
  console.error(`Missing ${outDir} — run pnpm build first`)
  process.exit(1)
}

const manifestPath = path.join(outDir, 'embed-offline-manifest.json')
if (!fs.existsSync(manifestPath)) {
  console.error(`Missing ${manifestPath}`)
  process.exit(1)
}

const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
let paths = [...(manifest.paths || [])]

const embedFile = outFileForUrl('/embed')
if (fs.existsSync(embedFile)) {
  const html = fs.readFileSync(embedFile, 'utf8')
  for (const extra of collectNextStatic(html)) {
    if (!paths.includes(extra)) paths.push(extra)
  }
}

if (fs.existsSync(slotRoot)) {
  fs.rmSync(slotRoot, { recursive: true, force: true })
}
fs.mkdirSync(slotRoot, { recursive: true })

let ok = 0
let miss = 0
for (const urlPath of paths) {
  if (urlPath.startsWith('/embed?')) continue
  try {
    const bytes = await loadBytes(urlPath)
    writeBody(urlPath, bytes)
    ok += 1
  } catch (err) {
    console.warn(`skip ${urlPath}: ${err instanceof Error ? err.message : err}`)
    miss += 1
  }
}

console.log(`Packed editor v${version} -> ${slotRoot} (${ok} ok, ${miss} missing)`)
if (miss > 0) process.exitCode = 1
