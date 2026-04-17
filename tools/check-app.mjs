import { spawn } from 'node:child_process';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const { chromium } = require('playwright');

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const nodePath =
  'C:\\Users\\cleyt\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\bin\\node.exe';

const server = spawn(nodePath, ['tools/local-server.mjs'], {
  cwd: root,
  stdio: ['ignore', 'pipe', 'pipe'],
});

try {
  await waitForServer('http://localhost:5173');

  const browser = await chromium.launch({
    executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    headless: true,
  });
  const page = await browser.newPage({ viewport: { width: 1440, height: 950 } });
  await page.goto('http://localhost:5173', { waitUntil: 'networkidle' });
  await page.locator('#loginForm button[type="submit"]').click();
  await page.waitForFunction(() => document.querySelector('#appView')?.hidden === false);
  await page.waitForSelector('.module.active table tbody tr');
  await page.screenshot({ path: path.resolve(root, 'erp-preview-desktop.png'), fullPage: false });

  const result = await page.evaluate(() => ({
    module: document.querySelector('.nav-link.active')?.textContent?.trim(),
    total: document.querySelector('#module-acompanhamento .metric-card strong')?.textContent?.trim(),
    firstStudent: document.querySelector('#module-acompanhamento [data-open-student]')?.textContent?.trim(),
  }));

  await page.setViewportSize({ width: 390, height: 900 });
  await page.screenshot({ path: path.resolve(root, 'erp-preview-mobile.png'), fullPage: false });

  console.log(JSON.stringify(result));
  await browser.close();
} finally {
  server.kill();
}

async function waitForServer(url) {
  const started = Date.now();
  while (Date.now() - started < 30000) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      await new Promise((resolve) => setTimeout(resolve, 300));
    }
  }

  throw new Error('Local server did not start');
}
