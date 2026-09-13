#!/usr/bin/env node
/**
 * Writes embed-release.json + embed-offline-manifest.json for packaged workspace offline cache.
 */
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..')
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'))
const appRootJson = JSON.parse(fs.readFileSync(path.join(root, 'public/app-root.json'), 'utf8'))
const appRoot = String(appRootJson.appRoot || '/v9.3.1-2').replace(/\/$/, '')
const version = String(pkg.version || '0.0.0')

const editorShells = [
  '/web-apps/apps/documenteditor/main/index.html',
  '/web-apps/apps/spreadsheeteditor/main/index.html',
  '/web-apps/apps/presentationeditor/main/index.html',
  '/web-apps/apps/pdfeditor/main/index.html',
]

const paths = [
  '/embed',
  '/embed-partner.json',
  '/app-root.json',
  `${appRoot}/web-apps/apps/api/documents/api.js`,
  `${appRoot}/web-apps/apps/api/documents/preload.html`,
  `${appRoot}/sdkjs-plugins/v1/plugins.js`,
  `${appRoot}/plugins.json`,
  ...editorShells.map((p) => `${appRoot}${p}`),
  '/x2t/x2t.js',
  '/x2t/x2t.wasm',
  '/office-plugins/agent/config.json',
  '/office-plugins/agent/index.html',
  '/office-plugins/agent/plugin.js',
]

const publicDir = path.join(root, 'public')
fs.writeFileSync(
  path.join(publicDir, 'embed-release.json'),
  `${JSON.stringify({ version, manifest: '/embed-offline-manifest.json' }, null, 2)}\n`
)
fs.writeFileSync(
  path.join(publicDir, 'embed-offline-manifest.json'),
  `${JSON.stringify({ version, appRoot, paths }, null, 2)}\n`
)

console.log(`embed offline manifest v${version} (${paths.length} paths, appRoot=${appRoot})`)
