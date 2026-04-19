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

let server = null;

const errors = [];

try {
  if (!(await serverIsReady('http://localhost:5173'))) {
    server = spawn(nodePath, ['tools/local-server.mjs'], {
      cwd: root,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
  }
  await waitForServer('http://localhost:5173');
  const browser = await chromium.launch({
    executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    headless: true,
  });
  const page = await browser.newPage({ viewport: { width: 1440, height: 920 } });
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error' && !message.text().includes('Failed to load resource')) {
      errors.push(message.text());
    }
  });

  await page.goto('http://localhost:5173', { waitUntil: 'networkidle' });
  await page.locator('#loginForm button[type="submit"]').click();
  await page.waitForSelector('#module-inteligencia.active');

  const moduleIds = await page.$$eval('.module', (modules) => modules.map((item) => item.id.replace('module-', '')));
  const routes = new Set();
  const clicked = [];

  await clickModuleButtons(page, '.top-nav [data-module]:not([hidden])', routes, clicked);
  await clickMegaMenuButtons(page, routes, clicked);

  await page.locator('#profileSelect').selectOption('consultor');
  await page.locator('#megaMenuButton').click();
  const restrictedVisible = await page.$$eval(
    '#megaMenu [data-admin-only], #megaMenu [data-admin-module], #megaMenu [data-quick-module="financeiro"]',
    (items) => items.filter((item) => !item.hidden && getComputedStyle(item).display !== 'none').map((item) => item.textContent.trim()),
  );
  if (restrictedVisible.length) {
    throw new Error(`Restritos visiveis para Consultor: ${restrictedVisible.join(', ')}`);
  }

  await browser.close();

  if (errors.length) {
    throw new Error(errors.join('\n'));
  }

  const expectedRouted = [
    'acompanhamento',
    'agenda',
    'avaliacoes',
    'bi',
    'cursos',
    'fila',
    'financeiro',
    'inteligencia',
    'matriculas',
    'metas',
    'repasse',
    'retencao',
    'seguranca',
  ];
  const missingRoutes = expectedRouted.filter((module) => !routes.has(module));
  const missingModules = [...routes].filter((module) => !moduleIds.includes(module));
  if (missingRoutes.length || missingModules.length) {
    throw new Error(
      JSON.stringify({
        missingRoutes,
        missingModules,
      }),
    );
  }

  console.log(
    JSON.stringify({
      buttonsClicked: clicked.length,
      routes: [...routes].sort(),
      sample: clicked.slice(0, 12),
    }),
  );
} finally {
  if (server) server.kill();
}

async function serverIsReady(url) {
  try {
    const response = await fetch(url);
    return response.ok;
  } catch {
    return false;
  }
}

async function clickModuleButtons(page, selector, routes, clicked, adminMenu = false, originModule = '') {
  const candidates = await page.$$eval(selector, (items) =>
    items.map((item) => ({
      label: item.textContent.trim(),
      module: item.dataset.module || item.dataset.quickModule || item.dataset.adminModule || '',
      view: item.dataset.view || '',
      hidden: item.hidden || getComputedStyle(item).display === 'none',
    })),
  );
  for (const data of candidates) {
    if (!data.module || data.hidden) continue;
    if (originModule) {
      await page.locator(`[data-module="${originModule}"]`).first().click();
      await page.waitForSelector(`#module-${originModule}.active`);
    }
    if (adminMenu) {
      const menuIsOpen = await page.$eval('[data-admin-menu]', (menu) => menu.open).catch(() => false);
      if (!menuIsOpen) await page.locator('[data-admin-menu] summary').click();
    }
    const dataAttr = selector.includes('data-admin-module')
      ? 'data-admin-module'
      : selector.includes('data-quick-module')
        ? 'data-quick-module'
        : 'data-module';
    const viewPart = data.view ? `[data-view="${data.view}"]` : '';
    const button = page.locator(`${selector}[${dataAttr}="${data.module}"]${viewPart}`).first();
    routes.add(data.module);
    await button.click();
    await page.waitForSelector(`#module-${data.module}.active`);
    clicked.push(`${data.label}:${data.module}`);
  }
}

async function clickMegaMenuButtons(page, routes, clicked) {
  await page.locator('#megaMenuButton').click();
  const candidates = await page.$$eval('#megaMenu [data-quick-module], #megaMenu [data-admin-module]', (items) =>
    items.map((item) => ({
      label: item.textContent.trim(),
      module: item.dataset.quickModule || item.dataset.adminModule || '',
      view: item.dataset.view || '',
      attr: item.dataset.quickModule ? 'data-quick-module' : 'data-admin-module',
      hidden: item.hidden || getComputedStyle(item).display === 'none',
    })),
  );
  await page.locator('#megaMenuClose').click();

  for (const data of candidates) {
    if (!data.module || data.hidden) continue;
    await page.locator('#megaMenuButton').click();
    const viewPart = data.view ? `[data-view="${data.view}"]` : '';
    const button = page.locator(`#megaMenu [${data.attr}="${data.module}"]${viewPart}`).first();
    routes.add(data.module);
    await button.click();
    await page.waitForSelector(`#module-${data.module}.active`);
    clicked.push(`${data.label}:${data.module}`);
  }
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
