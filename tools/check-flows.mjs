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

const errors = [];

try {
  await waitForServer('http://localhost:5173');
  const browser = await chromium.launch({
    executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    headless: true,
  });
  const page = await browser.newPage({ viewport: { width: 1366, height: 850 } });
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error' && !message.text().includes('Failed to load resource')) {
      errors.push(message.text());
    }
  });

  await page.goto('http://localhost:5173', { waitUntil: 'networkidle' });
  await page.locator('#loginForm button[type="submit"]').click();
  await page.waitForSelector('.module.active table tbody tr');
  await page.locator('.metric-card[data-dashboard-focus="risco"]').click();
  await page.waitForSelector('text=Alunos que precisam de atenção');
  await page.locator('#module-inteligencia [data-open-student]').first().click();
  await page.waitForSelector('.health-score-card');
  await page.waitForSelector('.checklist-panel');
  await page.waitForSelector('.timeline-panel');
  await page.locator('#closeDrawer').click();

  await page.locator('#megaMenuButton').click();
  await page.locator('#megaMenu [data-quick-module="financeiro"]').click();
  await page.waitForSelector('#module-financeiro .table-panel');
  const financeTitle = await page.locator('#module-financeiro h1').textContent();

  await page.locator('#profileSelect').selectOption('consultor');
  await page.locator('#megaMenuButton').click();
  await page.waitForFunction(() => document.querySelector('#megaMenu [data-quick-module="financeiro"]')?.hidden === true);
  const locked = await page.locator('#megaMenu [data-quick-module="financeiro"]').getAttribute('title');
  await page.locator('#megaMenuClose').click();

  await page.locator('#profileSelect').selectOption('admin');
  await page.locator('[data-module="retencao"]').click();
  await page.waitForSelector('#module-retencao .table-panel');
  await page.locator('#module-retencao [data-open-retention]').first().click();
  await page.waitForSelector('form[data-form="retention-contact"]');
  await page.locator('form[data-form="retention-contact"] select[name="contacted"]').selectOption('true');
  await page.locator('form[data-form="retention-contact"] textarea[name="note"]').fill('Teste automatizado de contato de retenção.');
  await page.locator('form[data-form="retention-contact"] button[type="submit"]').click();
  await page.waitForSelector('.toast.show');
  const retentionToast = await page.locator('.toast').textContent();
  await page.locator('#closeDrawer').click();

  await page.locator('[data-module="matriculas"][data-view="crm"]').click();
  await page.waitForSelector('#module-matriculas.active');
  await page.locator('#megaMenuButton').click();
  await page.locator('#megaMenu [data-quick-module="matriculas"][data-view="matricula"]').click();
  await page.waitForSelector('#module-matriculas.active');

  await page.locator('#megaMenuButton').click();
  await page.locator('#megaMenu [data-quick-module="inteligencia"]').click();
  await page.waitForSelector('#module-inteligencia.active');

  await page.locator('[data-module="agenda"]').click();
  await page.waitForSelector('#module-agenda.active');

  await page.locator('#megaMenuButton').click();
  await page.locator('#megaMenu [data-quick-module="financeiro"]').click();
  await page.waitForSelector('#module-financeiro.active');

  for (const module of ['cursos', 'seguranca', 'repasse']) {
    await page.locator('#megaMenuButton').click();
    await page.locator(`#megaMenu [data-admin-module="${module}"]`).first().click();
    await page.waitForSelector(`#module-${module}.active`);
  }

  await browser.close();

  if (errors.length) {
    throw new Error(errors.join('\n'));
  }

  console.log(
    JSON.stringify({
      financeTitle,
      locked,
      retentionToast,
      modulesChecked: 8,
    }),
  );
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
