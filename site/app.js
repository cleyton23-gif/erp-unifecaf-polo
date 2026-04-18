const SHEET_ID = '1AoZ9KCNIaIzTEW17MyNFv5O7lVmcoaa8_bFKI8dBY8Q';
const GOOGLE_CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv`;
const STORAGE_KEY = 'unifecaf-erp-state-v2';

const fields = {
  id: 'ID_Matricula',
  unit: 'Unidade',
  name: 'Nome',
  ra: 'RA',
  cpf: 'CPF',
  phone: 'Celular',
  email: 'E-mail',
  enrollmentDate: 'Data Matrícula',
  course: 'Curso',
  startPeriod: 'Período Ingressante',
  lastPeriod: 'Ult. Período',
  status: 'Status Ult. Período',
  debt: 'Inadimplente',
  debtValue: 'Valor Devido',
  entry: 'Forma de Ingresso',
  currentPeriod: 'Período Vig.',
  currentClass: 'Turma Vig.',
  origin: 'Origem',
  oneAccess: 'Acessou One?',
  oneDays: 'Dias Ult. Acesso One',
  avaAccess: 'Acessou AVA?',
  avaDays: 'Dias Ult. Acesso AVA',
  avaLastAccessDate: 'Data Último Acesso AVA',
  plan: 'Plano',
  address: 'Endereço',
  overdueMonths: 'Meses Inadimplentes',
};

const STATUS_OPTIONS = ['ATIVO', 'PREMAT', 'TRANCADO', 'ABANDONADO', 'CANCELADO', 'DESISTENTE'];
const LEAD_STAGES = ['Lead', 'Visita', 'Em negociação', 'Gerar contrato', 'Matriculado', 'Perdido'];
const MONTHS = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

const state = {
  profile: 'admin',
  currentUser: null,
  module: 'inteligencia',
  segment: 'todos',
  allRows: [],
  filteredRows: [],
  selectedKey: '',
  store: loadStore(),
  users: [],
  remoteState: false,
  remoteSaveTimer: 0,
  remoteWarningShown: false,
};

const els = {
  loginView: document.querySelector('#loginView'),
  loginForm: document.querySelector('#loginForm'),
  loginUser: document.querySelector('#loginUser'),
  loginPassword: document.querySelector('#loginPassword'),
  appView: document.querySelector('#appView'),
  sourceStatus: document.querySelector('#sourceStatus'),
  userStatus: document.querySelector('#userStatus'),
  syncStamp: document.querySelector('#syncStamp'),
  profileSelect: document.querySelector('#profileSelect'),
  refreshButton: document.querySelector('#refreshButton'),
  logoutButton: document.querySelector('#logoutButton'),
  searchInput: document.querySelector('#searchInput'),
  statusFilter: document.querySelector('#statusFilter'),
  courseFilter: document.querySelector('#courseFilter'),
  cohortFilter: document.querySelector('#cohortFilter'),
  unitFilter: document.querySelector('#unitFilter'),
  debtFilter: document.querySelector('#debtFilter'),
  csvInput: document.querySelector('#csvInput'),
  alerts: document.querySelector('#alerts'),
  drawer: document.querySelector('#studentDrawer'),
  drawerContent: document.querySelector('#drawerContent'),
  closeDrawer: document.querySelector('#closeDrawer'),
  toast: document.querySelector('#toast'),
  modules: {
    inteligencia: document.querySelector('#module-inteligencia'),
    bi: document.querySelector('#module-bi'),
    acompanhamento: document.querySelector('#module-acompanhamento'),
    retencao: document.querySelector('#module-retencao'),
    financeiro: document.querySelector('#module-financeiro'),
    cursos: document.querySelector('#module-cursos'),
    metas: document.querySelector('#module-metas'),
    agenda: document.querySelector('#module-agenda'),
    avaliacoes: document.querySelector('#module-avaliacoes'),
    fila: document.querySelector('#module-fila'),
    seguranca: document.querySelector('#module-seguranca'),
  },
};

const collator = new Intl.Collator('pt-BR', { sensitivity: 'base' });

bootstrap();

function bootstrap() {
  seedStore();
  bindEvents();

  if (localStorage.getItem('unifecaf-erp-session') === 'active') {
    const savedUser = JSON.parse(localStorage.getItem('unifecaf-erp-user') || 'null');
    state.currentUser = savedUser;
    state.profile = savedUser?.perfil || 'admin';
    showApp();
  }
}

function bindEvents() {
  els.loginForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    const ok = await authenticateUser(els.loginUser.value, els.loginPassword.value);
    if (!ok) {
      toast('Usuário ou senha inválidos, ou usuário inativo.');
      return;
    }
    localStorage.setItem('unifecaf-erp-session', 'active');
    showApp();
  });

  els.logoutButton.addEventListener('click', () => {
    localStorage.removeItem('unifecaf-erp-session');
    localStorage.removeItem('unifecaf-erp-user');
    state.currentUser = null;
    els.appView.hidden = true;
    els.loginView.hidden = false;
  });

  els.profileSelect.addEventListener('change', () => {
    if (state.currentUser && normalize(state.currentUser.perfil) !== 'admin') {
      state.profile = state.currentUser.perfil || 'consultor';
      updateUserStatus();
      toast('Somente Admin pode alternar a visão de perfil.');
      return;
    }

    state.profile = els.profileSelect.value;
    updateUserStatus();
    render();
  });

  document.querySelectorAll('.nav-link').forEach((button) => {
    button.addEventListener('click', () => activateModule(button.dataset.module));
  });

  [els.searchInput, els.statusFilter, els.courseFilter, els.cohortFilter, els.unitFilter, els.debtFilter].forEach(
    (input) => input.addEventListener('input', render),
  );

  document.querySelectorAll('.segment').forEach((button) => {
    button.addEventListener('click', () => {
      state.segment = button.dataset.segment;
      document.querySelectorAll('.segment').forEach((item) => item.classList.remove('active'));
      button.classList.add('active');
      render();
    });
  });

  els.refreshButton.addEventListener('click', loadRemoteSheet);
  els.closeDrawer.addEventListener('click', () => els.drawer.classList.remove('open'));

  els.csvInput.addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    hydrateRows(await file.text(), `CSV importado: ${file.name}`);
  });

  document.addEventListener('submit', handleSubmit);
  document.addEventListener('click', handleClick);
  document.addEventListener('change', handleChange);
  document.addEventListener('dragstart', handleDragStart);
  document.addEventListener('dragover', handleDragOver);
  document.addEventListener('dragleave', handleDragLeave);
  document.addEventListener('drop', handleDrop);
  document.addEventListener('dragend', clearLeadDragState);
}

async function showApp() {
  els.loginView.hidden = true;
  els.appView.hidden = false;
  updateUserStatus();
  activateModule(state.module, false);
  await loadOperationalStore();
  await loadRemoteSheet();
}

function activateModule(module, shouldRender = true) {
  state.module = module || 'inteligencia';
  document.querySelectorAll('.nav-link').forEach((item) => item.classList.toggle('active', item.dataset.module === state.module));
  Object.entries(els.modules).forEach(([name, element]) => {
    element.classList.toggle('active', name === state.module);
  });
  if (shouldRender) render();
}

function applyAccessControl() {
  const financeButton = document.querySelector('[data-module="financeiro"]');
  if (!financeButton) return;
  financeButton.classList.toggle('restricted', !canSeeFinancial());
  financeButton.title = canSeeFinancial() ? 'Acessar Financeiro' : 'Acesso restrito a Admin e Financeiro';
}

function updateUserStatus() {
  const actualProfile = normalize(state.currentUser?.perfil || state.profile || 'consultor');
  const canSwitchProfile = actualProfile === 'admin';

  if (!canSwitchProfile) {
    state.profile = actualProfile || 'consultor';
  }

  els.profileSelect.value = state.profile;
  els.profileSelect.disabled = !canSwitchProfile;
  applyAccessControl();

  if (state.currentUser?.nome) {
    const viewLabel =
      canSwitchProfile && actualProfile !== normalize(state.profile)
        ? ` · visão ${roleLabel(state.profile)}`
        : '';
    els.userStatus.textContent = `${state.currentUser.nome} (${roleLabel(actualProfile)}${viewLabel})`;
    return;
  }

  els.userStatus.textContent = `Perfil ${roleLabel(state.profile)}`;
}

async function authenticateUser(username, password) {
  const users = await loadUsers();
  const normalizedUser = normalize(username);
  const user = users.find(
    (item) =>
      normalize(item.usuario) === normalizedUser &&
      String(item.senha || '') === String(password || '') &&
      normalize(item.ativo || 'SIM') !== 'nao',
  );

  if (!user) return false;

  state.currentUser = {
    usuario: user.usuario,
    nome: user.nome || user.usuario,
    perfil: normalize(user.perfil || 'consultor'),
  };
  state.profile = state.currentUser.perfil;
  localStorage.setItem('unifecaf-erp-user', JSON.stringify(state.currentUser));
  return true;
}

async function loadUsers() {
  try {
    const response = await fetch('/.netlify/functions/state?action=users', { cache: 'no-store' });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    const users = Array.isArray(payload.users) ? payload.users : [];
    state.users = users.length ? users : defaultUsers();
  } catch {
    state.users = defaultUsers();
  }
  return state.users;
}

async function loadOperationalStore() {
  try {
    const response = await fetch('/.netlify/functions/state', { cache: 'no-store' });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    if (!payload.ok || !payload.state) throw new Error(payload.message || 'Estado remoto indisponivel');

    state.store = normalizeStore(payload.state);
    state.users = Array.isArray(payload.users) ? payload.users : state.users;
    seedStore();
    persistLocal();
    state.remoteState = true;
    toast('Estado operacional carregado da planilha ERP.');
  } catch (error) {
    state.remoteState = false;
    state.store = normalizeStore(state.store);
    seedStore();
    els.syncStamp.textContent = `Estado local ativo: ${error.message}`;
  }
}

async function loadRemoteSheet() {
  setSourceStatus('Sincronizando planilha');
  const sources = ['/.netlify/functions/sheet', GOOGLE_CSV_URL];
  let lastError = null;

  for (const source of sources) {
    try {
      const response = await fetch(source, { cache: 'no-store' });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      hydrateRows(await response.text(), source.includes('netlify') ? 'Planilha via Netlify' : 'Planilha Google');
      return;
    } catch (error) {
      lastError = error;
    }
  }

  setSourceStatus('Importe um CSV para continuar');
  toast(`Não foi possível carregar a planilha: ${lastError?.message ?? 'erro desconhecido'}`);
  render();
}

function hydrateRows(csvText, label) {
  const parsed = parseCsv(csvText);
  state.allRows = parsed.rows.map((row, index) => enrichRow(row, index));

  populateFilter(els.statusFilter, uniqueValues(state.allRows, (row) => row.localStatus), 'Todos');
  populateFilter(els.courseFilter, uniqueValues(state.allRows, (row) => row.course), 'Todos');
  populateFilter(els.cohortFilter, uniqueValues(state.allRows, (row) => row.startPeriod), 'Todos');
  populateFilter(els.unitFilter, uniqueValues(state.allRows, (row) => row.unit), 'Todas');

  setSourceStatus(`${label} · ${state.allRows.length.toLocaleString('pt-BR')} registros`);
  els.syncStamp.textContent = `Última sincronização: ${new Date().toLocaleString('pt-BR')}`;
  render();
}

function enrichRow(row, index) {
  const key = rowKey(row, index);
  const override = state.store.overrides[key] || {};
  const localStatus = override.status || cleanText(row[fields.status]) || 'SEM STATUS';
  const debtValue = parseMoney(row[fields.debtValue]);
  const enriched = {
    raw: row,
    key,
    index,
    id: cleanText(row[fields.id]),
    unit: cleanText(row[fields.unit]),
    name: cleanText(row[fields.name]) || 'Sem nome',
    ra: cleanText(row[fields.ra]) || key,
    cpf: cleanText(row[fields.cpf]),
    phone: cleanText(override.contactPhone) || cleanText(row[fields.phone]),
    email: cleanText(override.contactEmail) || cleanText(row[fields.email]),
    enrollmentDate: cleanText(row[fields.enrollmentDate]),
    course: cleanText(row[fields.course]) || 'Sem curso',
    startPeriod: cleanText(row[fields.startPeriod]) || cleanText(row[fields.currentPeriod]) || '-',
    currentClass: cleanText(row[fields.currentClass]),
    sourceStatus: cleanText(row[fields.status]) || 'SEM STATUS',
    localStatus,
    statusOverridden: Boolean(override.status),
    contactOverridden: Boolean(override.contactPhone || override.contactEmail || override.followStatus),
    isDebt: normalize(row[fields.debt]) === 'sim',
    debtValue,
    entry: cleanText(row[fields.entry]),
    origin: cleanText(row[fields.origin]),
    oneAccess: cleanText(row[fields.oneAccess]),
    oneDays: cleanText(row[fields.oneDays]),
    avaAccess: cleanText(row[fields.avaAccess]),
    avaDays: cleanText(row[fields.avaDays]),
    avaLastAccessDate: getFirstValue(row, [
      fields.avaLastAccessDate,
      'Último Acesso AVA',
      'Ult. Acesso AVA',
      'Data Ult. Acesso AVA',
      'Data Últ. Acesso AVA',
      'Último acesso AVA',
    ]),
    plan: cleanText(row[fields.plan]),
    address: cleanText(row[fields.address]),
    overdueMonths: cleanText(row[fields.overdueMonths]) || 'Não informado na sede',
    override,
    retention: state.store.retention[key] || {},
  };
  enriched.avaDaysNumber = getAvaDaysNumber(enriched);
  enriched.avaAlert = getAvaAlert(enriched.avaDaysNumber, enriched.avaAccess);
  enriched.risk = getRisk(enriched);
  enriched.cohort = parseCohort(enriched.startPeriod);
  return enriched;
}

function render() {
  state.filteredRows = applyFilters(state.allRows);
  renderAlerts();
  renderInteligencia();
  renderBI();
  renderAcompanhamento();
  renderRetencao();
  renderFinanceiro();
  renderCursos();
  renderMetas();
  renderAgenda();
  renderAvaliacoes();
  renderFila();
  renderSeguranca();
}

function applyFilters(rows) {
  const query = normalize(els.searchInput.value);
  const status = els.statusFilter.value;
  const course = els.courseFilter.value;
  const cohort = els.cohortFilter.value;
  const unit = els.unitFilter.value;
  const debt = els.debtFilter.value;

  return rows
    .filter((row) => {
      if (status && row.localStatus !== status) return false;
      if (course && row.course !== course) return false;
      if (cohort && row.startPeriod !== cohort) return false;
      if (unit && row.unit !== unit) return false;
      if (debt === 'SIM' && !row.isDebt) return false;
      if (debt === 'NÃO' && row.isDebt) return false;

      if (query) {
        const haystack = normalize([row.name, row.ra, row.cpf, row.phone, row.email, row.course, row.unit].join(' '));
        if (!haystack.includes(query)) return false;
      }

      if (state.segment === 'risco' && row.risk.level !== 'Alto') return false;
      if (state.segment === 'override' && !row.statusOverridden) return false;
      if (state.segment === 'sem-acesso' && !hasNoAccess(row)) return false;

      return true;
    })
    .sort((a, b) => b.risk.score - a.risk.score || collator.compare(a.name, b.name));
}

function renderAlerts() {
  const lateFollowups = state.filteredRows.filter((row) => row.risk.level === 'Alto').slice(0, 2);
  const noContact = state.filteredRows.filter((row) => !row.override.lastContact && row.risk.score >= 35).slice(0, 2);
  const boletoPending = state.store.leads
    .filter((lead) => lead.stage === 'Matriculado' && !lead.boletoSent)
    .slice(0, 2);
  const alerts = [
    ...boletoPending.map((lead) => ({
      type: 'warning',
      text: `Boleto pendente: ${lead.name} foi matriculado e precisa de envio do boleto de matrícula.`,
    })),
    ...lateFollowups.map((row) => ({
      type: 'danger',
      text: `Prioridade: ${row.name} está com risco alto e status ${row.localStatus}.`,
    })),
    ...noContact.map((row) => ({
      type: 'warning',
      text: `Sem registro recente: ${row.name} precisa de anotação de acompanhamento.`,
    })),
  ];

  els.alerts.innerHTML = alerts.length
    ? alerts.map((item) => `<div class="alert ${item.type}">${escapeHtml(item.text)}</div>`).join('')
    : '<div class="alert ok">Operação carregada. Use os módulos para acompanhar o polo.</div>';
}

function renderInteligencia() {
  const rows = state.filteredRows.length ? state.filteredRows : state.allRows;
  const totals = getStudentTotals(rows);
  const debtRows = rows.filter((row) => row.isDebt);
  const activeRows = rows.filter(isActive);
  const retention = rows.length ? Math.round((activeRows.length / rows.length) * 100) : 0;
  const debtTotal = sumBy(debtRows, (row) => row.debtValue);
  const queue = dailyQueue();
  const target = Number(state.store.settings.monthlyTarget || 65);
  const matriculas = activeRows.length + state.store.localStudents.length;
  const gap = matriculas - target;
  const totalComputers = Number(state.store.settings.computersTotal || 24);
  const maintenance = Number(state.store.settings.computersMaintenance || 0);
  const availableComputers = Math.max(0, totalComputers - maintenance - currentExamReservations());
  const signals = decisionSignals({ rows, debtRows, debtTotal, retention, gap, availableComputers, queue });
  const matriculatedLeads = state.store.leads.filter((lead) => lead.stage === 'Matriculado').length;

  els.modules.inteligencia.innerHTML = `
    ${moduleTitle('Painel de acesso rápido', 'Quatro frentes centrais para operar o polo com menos cliques.')}
    <section class="quick-access-grid" aria-label="Módulos de acesso rápido">
      ${quickAccessCard('Gestão Financeira', 'Carteira, mensalidades, boleto e matrícula', 'financeiro', canSeeFinancial() ? formatMoney(debtTotal) : 'Blindado', canSeeFinancial() ? 'Em atraso' : 'Acesso por perfil', 'green')}
      ${quickAccessCard('Retenção de Alunos', 'Alertas AVA e registro de contato', 'retencao', totals.highRisk, 'Risco alto', 'yellow')}
      ${quickAccessCard('Matrículas', 'CRM de leads com Kanban arrastável', 'cursos', matriculatedLeads, 'Matriculados no funil', 'cyan')}
      ${quickAccessCard('Agendamentos', 'Aulas, provas e fila diária', 'agenda', state.store.schedule.length + state.store.exams.length, 'Reservas ativas', 'red')}
    </section>
    <section class="metric-grid">
      ${metricCard('Retenção filtrada', `${retention}%`, `${activeRows.length.toLocaleString('pt-BR')} ativos`, retention >= 70 ? 'green' : 'yellow')}
      ${metricCard('Risco alto', totals.highRisk, 'Intervenção acadêmica', totals.highRisk ? 'yellow' : 'green')}
      ${metricCard('Gap da meta', gap, `Meta mensal: ${target}`, gap >= 0 ? 'green' : 'red')}
      ${metricCard('Recursos livres', availableComputers, 'Computadores úteis agora', availableComputers > 4 ? 'cyan' : 'red')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Sinais para decisão', 'Prioridades automáticas')}</div>
          <span>${signals.length} sinais</span>
        </div>
        <div class="decision-list">
          ${signals
            .map(
              (signal) => `
                <div class="decision-signal ${signal.tone}">
                  <strong>${escapeHtml(signal.title)}</strong>
                  <span>${escapeHtml(signal.text)}</span>
                </div>
              `,
            )
            .join('')}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Plano de ação', 'Decisões registradas na planilha')}</div>
          <span>${state.store.decisions.length} ações</span>
        </div>
        <form class="stack-form" data-form="decision">
          <select name="area" required>
            <option>Acadêmico</option>
            <option>Financeiro</option>
            <option>Comercial</option>
            <option>Logística</option>
            <option>Infraestrutura</option>
          </select>
          <input name="title" placeholder="Decisão ou ação definida" required />
          <div class="two-cols">
            <input name="owner" placeholder="Responsável" required />
            <input name="due" type="date" required />
          </div>
          <select name="status">
            <option>Aberta</option>
            <option>Em andamento</option>
            <option>Concluída</option>
            <option>Bloqueada</option>
          </select>
          <button type="submit">Registrar decisão</button>
        </form>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Governança de decisões', 'Rastro para crescimento do polo')}</div>
        <span>${state.store.decisions.length} registros</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Área</th>
              <th>Decisão</th>
              <th>Responsável</th>
              <th>Prazo</th>
              <th>Status</th>
              <th>Criada em</th>
            </tr>
          </thead>
          <tbody>
            ${state.store.decisions
              .slice()
              .reverse()
              .map(
                (item) => `
                  <tr>
                    <td><span class="badge cyan">${escapeHtml(item.area)}</span></td>
                    <td>${escapeHtml(item.title)}</td>
                    <td>${escapeHtml(item.owner)}</td>
                    <td>${escapeHtml(item.due)}</td>
                    <td><span class="badge ${item.status === 'Concluída' ? 'green' : item.status === 'Bloqueada' ? 'red' : 'yellow'}">${escapeHtml(item.status)}</span></td>
                    <td>${escapeHtml(item.createdAt ? new Date(item.createdAt).toLocaleString('pt-BR') : '-')}</td>
                  </tr>
                `,
              )
              .join('') || emptyRow('Nenhuma decisão registrada ainda.')}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderBI() {
  const filter = getBiFilter();
  const rows = filterRowsByPeriod(state.allRows, filter);
  const censusRows = rows.length ? rows : state.allRows;
  const census = academicCensus(censusRows);
  const commercial = commercialKpis(filter);
  const finance = financeKpis(filter);
  const monthlyTarget = Number(state.store.settings.monthlyTarget || 65);
  const annualTarget = Number(state.store.settings.annualTarget || monthlyTarget * 12);
  const monthlyProgress = percent(commercial.monthlyMatriculations, monthlyTarget);
  const annualProgress = percent(commercial.annualMatriculations, annualTarget);

  els.modules.bi.innerHTML = `
    ${moduleTitle('Business Intelligence', 'Central de indicadores para tomada de decisão do polo.')}
    <section class="metric-grid">
      ${metricCard('Meta mensal', `${monthlyProgress}%`, `${commercial.monthlyMatriculations}/${monthlyTarget} matrículas`, monthlyProgress >= 100 ? 'green' : 'yellow')}
      ${metricCard('Meta anual', `${annualProgress}%`, `${commercial.annualMatriculations}/${annualTarget} matrículas`, annualProgress >= 100 ? 'green' : 'cyan')}
      ${metricCard('Ativos no censo', census.active, `${censusRows.length.toLocaleString('pt-BR')} registros analisados`, 'green')}
      ${metricCard('Inativos no censo', census.inactive, 'Trancado, abandonado e cancelado', census.inactive ? 'yellow' : 'green')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Funil comercial', 'Metas mensal e anual')}</div>
          <span>${escapeHtml(periodLabel(filter))}</span>
        </div>
        <form class="inline-form bi-controls" data-form="bi-filter">
          <select name="biMonth">${monthOptions(filter.month)}</select>
          <select name="biYear">${yearOptions(filter.year)}</select>
          <button type="submit">Filtrar</button>
        </form>
        <div class="progress-list">
          ${progressLine('Meta mensal', monthlyProgress, `${commercial.monthlyMatriculations} matrículas no período`)}
          ${progressLine('Meta anual', annualProgress, `${commercial.annualMatriculations} matrículas no ano`)}
        </div>
        <form class="inline-form bi-controls" data-form="bi-goals">
          <input type="number" min="1" name="monthlyTarget" value="${monthlyTarget}" placeholder="Meta mensal" />
          <input type="number" min="1" name="annualTarget" value="${annualTarget}" placeholder="Meta anual" />
          <button type="submit">Atualizar metas</button>
        </form>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Censo acadêmico interno', 'Status da base por período')}</div>
          <span>${censusRows.length.toLocaleString('pt-BR')} alunos</span>
        </div>
        <div class="status-chart">
          ${census.statuses
            .map(
              (item) => `
                <div class="status-chart-item ${item.className}">
                  <span style="height:${item.height}%"></span>
                  <strong>${item.count.toLocaleString('pt-BR')}</strong>
                  <em>${escapeHtml(item.label)}</em>
                </div>
              `,
            )
            .join('')}
        </div>
      </article>
    </section>
    <section class="split-grid">
      <article class="table-panel">
        <div class="panel-heading">
          <div>${smallTitle('Tabela do censo', 'Ativo, abandono, trancado e cancelado')}</div>
          <span>${escapeHtml(periodLabel(filter))}</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Status</th>
                <th>Quantidade</th>
                <th>Participação</th>
              </tr>
            </thead>
            <tbody>
              ${census.statuses
                .map(
                  (item) => `
                    <tr>
                      <td><span class="badge ${item.className}">${escapeHtml(item.label)}</span></td>
                      <td>${item.count.toLocaleString('pt-BR')}</td>
                      <td>${percent(item.count, Math.max(censusRows.length, 1))}%</td>
                    </tr>
                  `,
                )
                .join('')}
            </tbody>
          </table>
        </div>
      </article>
      <article class="table-panel">
        <div class="panel-heading">
          <div>${smallTitle('Saúde financeira', 'Taxa de matrícula e mensalidades recorrentes')}</div>
          <span>${canSeeFinancial() ? 'Comparativo operacional' : 'Blindado pelo RBAC'}</span>
        </div>
        ${
          canSeeFinancial()
            ? `
              <form class="inline-form bi-controls" data-form="bi-ticket">
                <input type="number" min="0" step="0.01" name="monthlyTicket" value="${Number(state.store.settings.monthlyTicket || 299)}" placeholder="Ticket mensalidade" />
                <button type="submit">Atualizar ticket</button>
              </form>
              <div class="finance-kpi-list">
                ${financeKpiItem('Taxa de Matrícula', finance.enrollment.current, finance.enrollment.previousMonth, finance.enrollment.previousYear)}
                ${financeKpiItem('Mensalidades Recorrentes', finance.recurring.current, finance.recurring.previousMonth, finance.recurring.previousYear)}
                ${financeKpiItem('Acumulado do Ano', finance.ytd.current, finance.ytd.previousMonth, finance.ytd.previousYear)}
              </div>
            `
            : '<div class="locked-panel"><strong>Acesso blindado pelo RBAC</strong><p>Fluxo de caixa disponível somente para Admin e Financeiro.</p></div>'
        }
      </article>
    </section>
  `;
}

function renderAcompanhamento() {
  const rows = state.filteredRows;
  const totals = getStudentTotals(rows);
  const cohorts = getCohortStats(rows).slice(0, 7);
  els.modules.acompanhamento.innerHTML = `
    ${moduleTitle('Cadastro e acompanhamento de alunos', 'Master Data com contato local soberano, status e visão 360°.')}
    <section class="metric-grid">
      ${metricCard('Base filtrada', rows.length, 'Registros na visão atual')}
      ${metricCard('Ativos', totals.active, 'Status operacional ativo', 'green')}
      ${metricCard('Overrides locais', totals.overrides, 'Soberanos nas sincronizações', 'cyan')}
      ${metricCard('Risco alto', totals.highRisk, 'Fila de intervenção', 'yellow')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Análise de safra', 'Período de início')}</div>
          <span>${cohorts.length} coortes</span>
        </div>
        <div class="bars">
          ${cohorts
            .map(
              (item) => `
                <div class="bar-line">
                  <strong>${escapeHtml(item.period)}</strong>
                  <div class="bar-track"><span style="width:${item.retention}%"></span></div>
                  <em>${item.retention}% retenção</em>
                </div>
              `,
            )
            .join('')}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Ação rápida', 'Linha operacional')}</div>
          <span>Contato, status e matrícula</span>
        </div>
        <div class="kanban-row">
          ${STATUS_OPTIONS.slice(0, 5)
            .map((status) => {
              const count = rows.filter((row) => normalize(row.localStatus).includes(normalize(status))).length;
              return `<div class="stage-card"><span>${escapeHtml(status)}</span><strong>${count}</strong></div>`;
            })
            .join('')}
        </div>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Grade soberana', 'Nome, CPF, RA, curso e período de início')}</div>
        <span>${rows.length.toLocaleString('pt-BR')} registros</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Aluno</th>
              <th>CPF</th>
              <th>RA</th>
              <th>Curso</th>
              <th>Período inicial</th>
              <th>Status local</th>
              <th>Ações</th>
            </tr>
          </thead>
          <tbody>
            ${rows.slice(0, 220).map(studentRowTemplate).join('') || emptyRow('Nenhum aluno encontrado.')}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderRetencao() {
  const rows = state.filteredRows
    .filter((row) => isActive(row) || row.avaAlert.level !== 'ok')
    .sort((a, b) => b.avaAlert.weight - a.avaAlert.weight || b.avaDaysNumber - a.avaDaysNumber);
  const yellow = rows.filter((row) => row.avaAlert.level === 'yellow');
  const red = rows.filter((row) => row.avaAlert.level === 'red');
  const contacted = rows.filter((row) => row.retention.contacted).length;
  const pendingContact = rows.filter((row) => row.avaAlert.level !== 'ok' && !row.retention.contacted).length;

  els.modules.retencao.innerHTML = `
    ${moduleTitle('Retenção AVA', 'Controle dos responsáveis por acesso, alerta e contato com alunos sem atividade.')}
    <section class="metric-grid">
      ${metricCard('Monitorados', rows.length, 'Alunos ativos ou em alerta')}
      ${metricCard('Alerta amarelo', yellow.length, '5 a 7 dias sem acesso', 'yellow')}
      ${metricCard('Alerta vermelho', red.length, '8 dias ou mais sem acesso', 'red')}
      ${metricCard('Contatos registrados', contacted, `${pendingContact} pendentes`, contacted ? 'green' : 'cyan')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Regra de alerta', 'Governança de retenção')}</div>
          <span>AVA</span>
        </div>
        <div class="decision-list">
          <div class="decision-signal success">
            <strong>Até 4 dias</strong>
            <span>Aluno dentro do acompanhamento normal.</span>
          </div>
          <div class="decision-signal warning">
            <strong>5 a 7 dias</strong>
            <span>Acende alerta amarelo para contato preventivo.</span>
          </div>
          <div class="decision-signal danger">
            <strong>8 dias ou mais</strong>
            <span>Acende alerta vermelho e entra na prioridade de retenção.</span>
          </div>
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Fila de responsáveis', 'Contato e motivo de ausência')}</div>
          <span>${pendingContact} pendentes</span>
        </div>
        <div class="reason-list">
          ${['Sem internet', 'Dificuldade no AVA', 'Problema financeiro', 'Sem tempo', 'Não localizado']
            .map((reason) => {
              const count = Object.values(state.store.retention).filter((item) => item.reason === reason).length;
              return `<div class="reason-item"><strong>${reason}</strong><span class="badge ${count ? 'yellow' : 'green'}">${count}</span></div>`;
            })
            .join('')}
        </div>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Janela de retenção', 'Controle de acesso ao AVA e contato')}</div>
        <span>${rows.length.toLocaleString('pt-BR')} alunos</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Aluno</th>
              <th>Curso</th>
              <th>Acesso AVA</th>
              <th>Dias sem acesso</th>
              <th>Alerta</th>
              <th>Contato</th>
              <th>Ação</th>
            </tr>
          </thead>
          <tbody>
            ${rows.slice(0, 260).map(retentionRowTemplate).join('') || emptyRow('Nenhum aluno em monitoramento nesta visão.')}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderFinanceiro() {
  const locked = !canSeeFinancial();
  if (locked) {
    els.modules.financeiro.innerHTML = accessDenied('Financeiro', 'O perfil Consultor possui blindagem total de valores financeiros.');
    return;
  }

  const rows = state.filteredRows;
  const debtRows = rows.filter((row) => row.isDebt);
  const paidRows = rows.filter((row) => !row.isDebt);
  const debtTotal = sumBy(debtRows, (row) => row.debtValue);
  const overrideValues = Object.values(state.store.overrides);
  const boletoSent = overrideValues.filter((item) => item.boletoSent).length;
  const exemptEnrollments = overrideValues.filter((item) => item.enrollmentExempt).length;
  const receivedEnrollment = sumBy(Object.values(state.store.overrides), (item) =>
    item.enrollmentPaid ? Number(item.enrollmentFee || 0) : 0,
  );
  const expectedEnrollment = rows.length * Number(state.store.settings.enrollmentFee || 99);
  const pie = pieChart(paidRows.length, debtRows.length);
  const evolution = debtEvolution(debtRows);

  els.modules.financeiro.innerHTML = `
    ${moduleTitle('Ecossistema financeiro', 'Matrículas sob controle do polo e mensalidades da sede em modo leitura.')}
    <section class="metric-grid">
      ${metricCard('Em dia', paidRows.length, 'Carteira sem atraso', 'green')}
      ${metricCard('Em atraso', debtRows.length, 'Alunos inadimplentes', 'red')}
      ${metricCard('Dívida total', formatMoney(debtTotal), 'Valor exato informado pela sede', 'yellow')}
      ${metricCard('Matrículas baixadas', formatMoney(receivedEnrollment), `Previsto: ${formatMoney(expectedEnrollment)}`, 'cyan')}
      ${metricCard('Boletos enviados', boletoSent, 'Matrícula do polo', 'green')}
      ${metricCard('Isentos', exemptEnrollments, 'Sem cobrança de matrícula', 'cyan')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Composição de carteira', 'Em dia vs. em atraso')}</div>
          <span>${rows.length.toLocaleString('pt-BR')} alunos</span>
        </div>
        <div class="chart-row">
          ${pie}
          <div class="legend-list">
            <span><i class="dot green"></i> Em dia: ${paidRows.length.toLocaleString('pt-BR')}</span>
            <span><i class="dot red"></i> Em atraso: ${debtRows.length.toLocaleString('pt-BR')}</span>
          </div>
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Evolução e volume da dívida', 'Agrupamento por safra')}</div>
          <span>${formatMoney(debtTotal)}</span>
        </div>
        <div class="mini-columns">
          ${evolution
            .map(
              (item) => `
                <div class="column-item">
                  <span style="height:${item.height}%"></span>
                  <em>${escapeHtml(item.label)}</em>
                  <strong>${formatCompactMoney(item.value)}</strong>
                </div>
              `,
            )
            .join('')}
        </div>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Controle analítico de inadimplência', 'Valor total e meses em atraso')}</div>
        <span>${debtRows.length.toLocaleString('pt-BR')} devedores</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Aluno</th>
              <th>Curso</th>
              <th>Status</th>
              <th>Valor em atraso</th>
              <th>Meses inadimplentes</th>
              <th>Matrícula polo</th>
            </tr>
          </thead>
          <tbody>
            ${debtRows.slice(0, 220).map(financialRowTemplate).join('') || emptyRow('Nenhum inadimplente nesta visão.')}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderCursos() {
  const catalog = getCourseCatalog();
  const leads = state.store.leads;
  const signed = leads.filter((lead) => lead.stage === 'Matriculado').length;
  const boletoPending = leads.filter((lead) => lead.stage === 'Matriculado' && !lead.boletoSent).length;
  const paymentOk = leads.filter(
    (lead) => lead.stage === 'Matriculado' && ['Pago', 'Isento'].includes(lead.enrollmentPaymentStatus),
  ).length;
  els.modules.cursos.innerHTML = `
    ${moduleTitle('Matrículas e CRM de Leads', 'Funil visual com Kanban arrastável, contrato, boleto e status de pagamento.')}
    <section class="metric-grid">
      ${metricCard('Leads ativos', leads.length, 'Base comercial local')}
      ${metricCard('Contratos assinados', signed, 'Status Matriculado', 'green')}
      ${metricCard('Boletos pendentes', boletoPending, 'Alerta automático', boletoPending ? 'yellow' : 'green')}
      ${metricCard('Matrículas baixadas', paymentOk, 'Pago ou isento', 'cyan')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Catálogo central', 'Inclusão, alteração e exclusão local')}</div>
          <span>${catalog.length} cursos</span>
        </div>
        <form class="inline-form" data-form="course">
          <input name="name" placeholder="Novo curso" required />
          <select name="modality">
            <option>EAD</option>
            <option>Semipresencial</option>
            <option>Presencial</option>
            <option>Pós-graduação</option>
          </select>
          <button type="submit">Adicionar</button>
        </form>
        <div class="catalog-list">
          ${catalog
            .slice(0, 80)
            .map(
              (course) => `
                <div class="catalog-item">
                  <div>
                    <strong>${escapeHtml(course.name)}</strong>
                    <span>${escapeHtml(course.modality)} · ${course.students} alunos</span>
                  </div>
                  ${course.local ? `<button type="button" data-delete-course="${escapeHtml(course.name)}">Excluir</button>` : '<span class="badge">Sede</span>'}
                </div>
              `,
            )
            .join('')}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Novo lead', 'Vinculação obrigatória ao catálogo')}</div>
          <span>${leads.length} registros locais</span>
        </div>
        <form class="stack-form" data-form="lead">
          <input name="name" placeholder="Nome do candidato" required />
          <input name="phone" placeholder="Telefone" />
          <select name="course" required>
            <option value="">Curso validado</option>
            ${catalog.map((course) => `<option>${escapeHtml(course.name)}</option>`).join('')}
          </select>
          <input name="origin" placeholder="Origem comercial" />
          <select name="stage">${LEAD_STAGES.map((stage) => `<option>${stage}</option>`).join('')}</select>
          <button type="submit">Salvar lead</button>
        </form>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Funil de captação', 'Conversão em matrícula gera registro local')}</div>
        <span>${leads.length} leads</span>
      </div>
      <div class="kanban-board">
        ${LEAD_STAGES.map((stage) => leadColumn(stage)).join('')}
      </div>
    </section>
  `;
}

function renderMetas() {
  const leads = state.store.leads.length;
  const matriculas = state.allRows.filter((row) => isActive(row)).length + state.store.localStudents.length;
  const target = Number(state.store.settings.monthlyTarget || 65);
  const conversion = leads + matriculas ? Math.round((matriculas / (leads + matriculas)) * 100) : 0;
  const gap = matriculas - target;
  const monthProgress = Math.min(100, Math.round((new Date().getDate() / daysInCurrentMonth()) * 100));
  const projection = Math.round(matriculas / Math.max(monthProgress, 1) * 100);

  els.modules.metas.innerHTML = `
    ${moduleTitle('Metas e performance comercial', 'KPIs dinâmicos com meta mensal editável.')}
    <section class="metric-grid">
      ${metricCard('Meta mensal', target, 'Editável pelo polo', 'cyan')}
      ${metricCard('Matrículas', matriculas, 'Realizado acumulado', 'green')}
      ${metricCard('Gap', gap, gap >= 0 ? 'Acima da meta' : 'Abaixo da meta', gap >= 0 ? 'green' : 'red')}
      ${metricCard('Conversão', `${conversion}%`, 'Leads vs. matrículas', 'yellow')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Input dinâmico de metas', 'Sazonalidade mensal')}</div>
          <span>${monthProgress}% do mês</span>
        </div>
        <form class="inline-form" data-form="goal">
          <input type="number" name="monthlyTarget" min="1" value="${target}" />
          <button type="submit">Atualizar meta</button>
        </form>
        <div class="progress-card">
          <span>Projeção de fechamento</span>
          <strong>${projection}</strong>
          <div class="bar-track"><span style="width:${Math.min(100, Math.round((projection / target) * 100))}%"></span></div>
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Ranking de consultores', 'Base local do funil')}</div>
          <span>${state.store.leads.length} leads</span>
        </div>
        <div class="rank-list">
          ${consultantRanking().map((item, index) => `<div><strong>${index + 1}. ${escapeHtml(item.name)}</strong><span>${item.total} registros · ${item.conversion}% conversão</span></div>`).join('')}
        </div>
      </article>
    </section>
  `;
}

function renderAgenda() {
  const events = state.store.schedule.slice().sort((a, b) => new Date(a.start) - new Date(b.start));
  els.modules.agenda.innerHTML = `
    ${moduleTitle('Agendamentos de aulas e provas', 'Agenda única com validação anti-conflito e acesso rápido às reservas de prova.')}
    <section class="quick-access-grid agenda-shortcuts">
      ${quickAccessCard('Agendar Aula', 'Docente, sala, horário e lotação', 'agenda', events.length, 'Aulas ativas', 'cyan')}
      ${quickAccessCard('Reservar Prova', 'Disciplinas, tempo e estações de TI', 'avaliacoes', state.store.exams.length, 'Provas ativas', 'yellow')}
      ${quickAccessCard('Fila do Dia', 'Check-in, atendimento e finalização', 'fila', dailyQueue().length, 'Eventos hoje', 'green')}
      ${quickAccessCard('Infraestrutura', 'Computadores disponíveis e manutenção', 'avaliacoes', currentExamReservations(), 'Estações reservadas', 'red')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Novo agendamento', 'Docente, assunto, horário, sala e lotação')}</div>
          <span>Bloqueio em tempo real</span>
        </div>
        <form class="stack-form" data-form="schedule">
          <input name="teacher" placeholder="Docente" required />
          <input name="subject" placeholder="Assunto" required />
          <div class="two-cols">
            <input name="start" type="datetime-local" required />
            <input name="end" type="datetime-local" required />
          </div>
          <div class="two-cols">
            <input name="room" placeholder="Sala" required />
            <input name="capacity" type="number" min="1" placeholder="Lotação prevista" required />
          </div>
          <button type="submit">Agendar aula</button>
        </form>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Transparência', 'Agenda única do polo')}</div>
          <span>${events.length} eventos</span>
        </div>
        <div class="timeline-list">
          ${events.map(scheduleTemplate).join('') || '<p class="muted">Nenhuma aula agendada.</p>'}
        </div>
      </article>
    </section>
  `;
}

function renderAvaliacoes() {
  const settings = state.store.settings;
  const total = Number(settings.computersTotal || 24);
  const maintenance = Number(settings.computersMaintenance || 2);
  const reserved = currentExamReservations();
  const available = Math.max(0, total - maintenance - reserved);
  const exams = state.store.exams.slice().sort((a, b) => new Date(a.start) - new Date(b.start));

  els.modules.avaliacoes.innerHTML = `
    ${moduleTitle('Avaliações e infraestrutura TI', 'Reserva de provas alinhada ao inventário real de computadores.')}
    <section class="metric-grid">
      ${metricCard('Computadores', total, 'Inventário total', 'cyan')}
      ${metricCard('Manutenção', maintenance, 'Retirados da capacidade', 'yellow')}
      ${metricCard('Reservados agora', reserved, 'Slots ocupados', 'red')}
      ${metricCard('Disponíveis', available, 'Capacidade útil', 'green')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Inventário dinâmico', 'Prevenção de overbooking')}</div>
          <span>${available} livres</span>
        </div>
        <form class="inline-form" data-form="inventory">
          <input type="number" name="computersTotal" min="1" value="${total}" />
          <input type="number" name="computersMaintenance" min="0" value="${maintenance}" />
          <button type="submit">Atualizar</button>
        </form>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Agendar prova', 'Tempo, disciplina e estação')}</div>
          <span>Liberação por duração</span>
        </div>
        <form class="stack-form" data-form="exam">
          <input name="student" placeholder="Aluno" required />
          <input name="discipline" placeholder="Disciplina/prova" required />
          <div class="two-cols">
            <input name="start" type="datetime-local" required />
            <input name="duration" type="number" min="15" value="90" required />
          </div>
          <input name="machines" type="number" min="1" value="1" />
          <button type="submit">Reservar prova</button>
        </form>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Mapa de provas', 'Gargalos e manutenção fora do pico')}</div>
        <span>${exams.length} reservas</span>
      </div>
      <div class="timeline-list">
        ${exams.map(examTemplate).join('') || '<p class="muted">Nenhuma prova agendada.</p>'}
      </div>
    </section>
  `;
}

function renderFila() {
  const events = dailyQueue();
  els.modules.fila.innerHTML = `
    ${moduleTitle('Execução diária', 'Fila FIFO com check-in, atendimento e check-out.')}
    <section class="metric-grid">
      ${metricCard('Na fila', events.length, 'Eventos de hoje')}
      ${metricCard('Próximos 30 min', events.filter((event) => event.soon).length, 'Preparação do ambiente', 'yellow')}
      ${metricCard('Em atendimento', events.filter((event) => event.status === 'presente').length, 'Aguardando finalização', 'cyan')}
      ${metricCard('Arquivados', state.store.archive.length, 'Concluídos', 'green')}
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Fila cronológica', 'O próximo evento assume o topo')}</div>
        <span>${new Date().toLocaleDateString('pt-BR')}</span>
      </div>
      <div class="queue-list">
        ${events.map(queueTemplate).join('') || '<p class="muted">Nenhum evento para hoje.</p>'}
      </div>
    </section>
  `;
}

function renderSeguranca() {
  const configuredUsers = state.users.length ? state.users : defaultUsers();
  els.modules.seguranca.innerHTML = `
    ${moduleTitle('Governança de acessos e segurança', 'RBAC com segmentação de dados e trilha local.')}
    <section class="metric-grid">
      ${metricCard('Persistência', state.remoteState ? 'Planilha' : 'Local', state.remoteState ? 'Google Sheets operacional' : 'Fallback neste navegador', state.remoteState ? 'green' : 'yellow')}
      ${metricCard('Overrides', Object.keys(state.store.overrides).length, 'Status soberanos')}
      ${metricCard('Decisões', state.store.decisions.length, 'Plano de ação gerencial', 'cyan')}
      ${metricCard('Auditoria', state.remoteState ? 'Ativa' : 'Local', 'Rastro de alterações', state.remoteState ? 'green' : 'yellow')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Perfis RBAC', 'Escopo de visualização')}</div>
          <span>Perfil atual: ${escapeHtml(roleLabel(state.profile))}</span>
        </div>
        <div class="role-grid">
          ${roleCard('Admin', 'Visão integral, faturamento global e saúde financeira macro.', state.profile === 'admin')}
          ${roleCard('Financeiro', 'Valores individuais, cobrança e baixa manual de matrícula.', state.profile === 'financeiro')}
          ${roleCard('Consultor', 'Dados cadastrais, comerciais e acadêmicos sem valores financeiros.', state.profile === 'consultor')}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Soberania local', 'Override não sobrescrito pela sede')}</div>
          <span>${Object.keys(state.store.overrides).length} registros</span>
        </div>
        <div class="audit-list">
          ${auditItems().join('') || '<p class="muted">Nenhuma alteração local registrada.</p>'}
        </div>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Usuários do sistema', 'Cadastro mantido na aba ERP_Usuarios')}</div>
        <span>${configuredUsers.length} usuários configurados</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Usuário</th>
              <th>Nome</th>
              <th>Perfil</th>
              <th>Status</th>
              <th>Onde alterar</th>
            </tr>
          </thead>
          <tbody>${configuredUsers.map(userRowTemplate).join('')}</tbody>
        </table>
      </div>
    </section>
  `;
}

function studentRowTemplate(row) {
  return `
    <tr>
      <td>
        <button class="text-button" type="button" data-open-student="${escapeHtml(row.key)}">
          ${escapeHtml(row.name)}
          <span>${escapeHtml(shortUnit(row.unit))}${row.contactOverridden ? ' · contato local' : ''}${row.override.followStatus ? ` · ${escapeHtml(row.override.followStatus)}` : ''}</span>
        </button>
      </td>
      <td>${escapeHtml(maskCpf(row.cpf))}</td>
      <td>${escapeHtml(row.ra)}</td>
      <td>${escapeHtml(row.course)}</td>
      <td><span class="badge cyan">${escapeHtml(row.startPeriod)}</span></td>
      <td>
        <select class="status-select" data-status-key="${escapeHtml(row.key)}">
          ${STATUS_OPTIONS.map((status) => `<option ${normalize(status) === normalize(row.localStatus) ? 'selected' : ''}>${status}</option>`).join('')}
        </select>
        ${row.statusOverridden ? '<span class="badge yellow">Local</span>' : '<span class="badge">Sede</span>'}
      </td>
      <td>
        <button class="mini-button" type="button" data-open-student="${escapeHtml(row.key)}">Editar</button>
        ${normalizePhone(row.phone) ? `<a class="mini-link" href="${whatsAppUrl(row)}" target="_blank" rel="noreferrer">WhatsApp</a>` : ''}
      </td>
    </tr>
  `;
}

function financialRowTemplate(row) {
  const override = state.store.overrides[row.key] || {};
  return `
    <tr>
      <td>${escapeHtml(row.name)}<span class="subtext">RA ${escapeHtml(row.ra)}</span></td>
      <td>${escapeHtml(row.course)}</td>
      <td><span class="badge ${statusClass(row.localStatus)}">${escapeHtml(row.localStatus)}</span></td>
      <td><strong>${formatMoney(row.debtValue)}</strong></td>
      <td>${escapeHtml(row.overdueMonths)}</td>
      <td>
        <span class="badge ${override.enrollmentExempt ? 'cyan' : override.enrollmentPaid ? 'green' : 'yellow'}">
          ${override.enrollmentExempt ? 'Isento' : override.enrollmentPaid ? 'Pago' : 'Pendente'}
        </span>
        <span class="subtext">Boleto: ${override.boletoSent ? 'enviado' : 'não enviado'}</span>
        <label class="check-line">
          <input type="checkbox" data-enrollment-paid="${escapeHtml(row.key)}" ${override.enrollmentPaid ? 'checked' : ''} />
          Baixada pelo polo
        </label>
      </td>
    </tr>
  `;
}

function retentionRowTemplate(row) {
  const retention = row.retention || {};
  const contactText = retention.contacted
    ? `${retention.contactDate || '-'} · ${retention.reason || 'sem motivo'}`
    : 'Sem contato registrado';
  return `
    <tr>
      <td>
        <button class="text-button" type="button" data-open-retention="${escapeHtml(row.key)}">
          ${escapeHtml(row.name)}
          <span>RA ${escapeHtml(row.ra)} · ${escapeHtml(maskCpf(row.cpf))}</span>
        </button>
      </td>
      <td>${escapeHtml(row.course)}</td>
      <td>${escapeHtml(row.avaAccess || '-')}<span class="subtext">${escapeHtml(row.avaLastAccessDate || row.avaDays || 'Sem data exata')}</span></td>
      <td><strong>${row.avaDaysNumber >= 0 ? `${row.avaDaysNumber} dias` : 'Sem dado'}</strong></td>
      <td><span class="badge ${row.avaAlert.className}">${escapeHtml(row.avaAlert.label)}</span></td>
      <td>${escapeHtml(contactText)}${retention.note ? `<span class="subtext">${escapeHtml(retention.note)}</span>` : ''}</td>
      <td><button class="mini-button" type="button" data-open-retention="${escapeHtml(row.key)}">Registrar contato</button></td>
    </tr>
  `;
}

function showDrawer(row) {
  state.selectedKey = row.key;
  const override = state.store.overrides[row.key] || {};
  els.drawerContent.className = 'drawer-content';
  els.drawerContent.innerHTML = `
    <p class="eyebrow">Visão 360°</p>
    <h2>${escapeHtml(row.name)}</h2>
    <div class="badge-row">
      <span class="badge ${statusClass(row.localStatus)}">${escapeHtml(row.localStatus)}</span>
      <span class="badge ${row.risk.level === 'Alto' ? 'red' : row.risk.level === 'Médio' ? 'yellow' : 'green'}">Risco ${escapeHtml(row.risk.level)}</span>
      ${row.statusOverridden ? '<span class="badge yellow">Override local</span>' : '<span class="badge">Sede</span>'}
    </div>
    <form class="stack-form" data-form="student-override" data-student-key="${escapeHtml(row.key)}">
      <label>Status local soberano
        <select name="status">
          ${STATUS_OPTIONS.map((status) => `<option ${normalize(status) === normalize(row.localStatus) ? 'selected' : ''}>${status}</option>`).join('')}
        </select>
      </label>
      <div class="two-cols">
        <label>Telefone/WhatsApp local
          <input name="contactPhone" value="${escapeHtml(row.phone)}" placeholder="Telefone atualizado pelo polo" />
        </label>
        <label>E-mail local
          <input name="contactEmail" type="email" value="${escapeHtml(row.email)}" placeholder="E-mail atualizado pelo polo" />
        </label>
      </div>
      <label>Status de acompanhamento
        <select name="followStatus">
          ${['Sem acompanhamento', 'Contato pendente', 'Em acompanhamento', 'Resolvido', 'Não localizado']
            .map((status) => `<option ${status === (override.followStatus || 'Sem acompanhamento') ? 'selected' : ''}>${status}</option>`)
            .join('')}
        </select>
      </label>
      <label>Registro de contato
        <input name="contact" placeholder="Ex: Ligação, WhatsApp, presencial" />
      </label>
      <label>Anotação
        <textarea name="note" rows="4" placeholder="Registre o andamento...">${escapeHtml(override.note || '')}</textarea>
      </label>
      ${canSeeFinancial() ? `
        <label>Taxa de matrícula paga?
          <select name="enrollmentPaid">
            <option value="false" ${override.enrollmentPaid ? '' : 'selected'}>Não</option>
            <option value="true" ${override.enrollmentPaid ? 'selected' : ''}>Sim</option>
          </select>
        </label>
        <label>Aluno isento de matrícula?
          <select name="enrollmentExempt">
            <option value="false" ${override.enrollmentExempt ? '' : 'selected'}>Não</option>
            <option value="true" ${override.enrollmentExempt ? 'selected' : ''}>Sim</option>
          </select>
        </label>
        <label>Boleto de matrícula enviado?
          <select name="boletoSent">
            <option value="false" ${override.boletoSent ? '' : 'selected'}>Não</option>
            <option value="true" ${override.boletoSent ? 'selected' : ''}>Sim</option>
          </select>
        </label>
        <label>Valor da matrícula do polo
          <input name="enrollmentFee" type="number" min="0" step="0.01" value="${Number(override.enrollmentFee || state.store.settings.enrollmentFee || 99)}" />
        </label>
      ` : ''}
      <button type="submit">Salvar alteração local</button>
    </form>
    <div class="detail-list">
      ${detailRow('CPF', maskCpf(row.cpf))}
      ${detailRow('RA', row.ra)}
      ${detailRow('Curso', row.course)}
      ${detailRow('Período início', row.startPeriod)}
      ${detailRow('Status sede', row.sourceStatus)}
      ${detailRow('One', `${row.oneAccess || '-'} ${row.oneDays || ''}`)}
      ${detailRow('AVA', `${row.avaAccess || '-'} ${row.avaDays || ''}`)}
      ${detailRow('Alerta AVA', `${row.avaAlert.label} · ${row.avaDaysNumber >= 0 ? `${row.avaDaysNumber} dias` : 'sem dado'}`)}
      ${canSeeFinancial() ? detailRow('Financeiro', `${row.isDebt ? 'Em atraso' : 'Em dia'} · ${formatMoney(row.debtValue)}`) : detailRow('Financeiro', 'Blindado pelo perfil')}
      ${canSeeFinancial() ? detailRow('Matrícula polo', `${override.enrollmentExempt ? 'Isento' : override.enrollmentPaid ? 'Pago' : 'Pendente'} · boleto ${override.boletoSent ? 'enviado' : 'não enviado'}`) : ''}
      ${detailRow('Telefone', row.phone)}
      ${detailRow('E-mail', row.email)}
      ${detailRow('Status acompanhamento', override.followStatus || 'Sem acompanhamento')}
      ${detailRow('Endereço', row.address)}
    </div>
  `;
  els.drawer.classList.add('open');
}

function showRetentionDrawer(row) {
  state.selectedKey = row.key;
  const retention = row.retention || {};
  els.drawerContent.className = 'drawer-content';
  els.drawerContent.innerHTML = `
    <p class="eyebrow">Retenção AVA</p>
    <h2>${escapeHtml(row.name)}</h2>
    <div class="badge-row">
      <span class="badge ${row.avaAlert.className}">${escapeHtml(row.avaAlert.label)}</span>
      <span class="badge cyan">${row.avaDaysNumber >= 0 ? `${row.avaDaysNumber} dias sem acesso` : 'sem dado de dias'}</span>
      <span class="badge ${retention.contacted ? 'green' : 'yellow'}">${retention.contacted ? 'Contato registrado' : 'Contato pendente'}</span>
    </div>
    <form class="stack-form" data-form="retention-contact" data-student-key="${escapeHtml(row.key)}">
      <label>Foi entrado em contato?
        <select name="contacted">
          <option value="false" ${retention.contacted ? '' : 'selected'}>Não</option>
          <option value="true" ${retention.contacted ? 'selected' : ''}>Sim</option>
        </select>
      </label>
      <div class="two-cols">
        <label>Data do contato
          <input name="contactDate" type="date" value="${toDateInputValue(retention.contactDate)}" />
        </label>
        <label>Canal
          <select name="channel">
            ${['WhatsApp', 'Ligação', 'Presencial', 'E-mail', 'Não localizado']
              .map((channel) => `<option ${retention.channel === channel ? 'selected' : ''}>${channel}</option>`)
              .join('')}
          </select>
        </label>
      </div>
      <label>Motivo informado
        <select name="reason">
          ${['Sem internet', 'Dificuldade no AVA', 'Problema financeiro', 'Sem tempo', 'Não localizado', 'Outro']
            .map((reason) => `<option ${retention.reason === reason ? 'selected' : ''}>${reason}</option>`)
            .join('')}
        </select>
      </label>
      <label>Responsável pela retenção
        <input name="responsible" value="${escapeHtml(retention.responsible || roleLabel(state.profile))}" />
      </label>
      <label>Observação
        <textarea name="note" rows="4" placeholder="Explique o motivo e o encaminhamento...">${escapeHtml(retention.note || '')}</textarea>
      </label>
      <button type="submit">Salvar contato de retenção</button>
    </form>
    <div class="detail-list">
      ${detailRow('RA', row.ra)}
      ${detailRow('CPF', maskCpf(row.cpf))}
      ${detailRow('Curso', row.course)}
      ${detailRow('Status', row.localStatus)}
      ${detailRow('Último acesso AVA', row.avaLastAccessDate || row.avaDays || '-')}
      ${detailRow('Telefone', row.phone)}
      ${detailRow('E-mail', row.email)}
    </div>
  `;
  els.drawer.classList.add('open');
}

function handleSubmit(event) {
  const form = event.target.closest('form[data-form]');
  if (!form) return;
  event.preventDefault();
  const data = Object.fromEntries(new FormData(form).entries());
  const formType = form.dataset.form;
  const resetAfterSubmit = !['student-override', 'retention-contact', 'bi-filter', 'bi-goals', 'bi-ticket'].includes(formType);

  if (formType === 'student-override') saveStudentOverride(form.dataset.studentKey, data);
  if (formType === 'retention-contact') saveRetentionContact(form.dataset.studentKey, data);
  if (formType === 'course') addCourse(data);
  if (formType === 'lead') addLead(data);
  if (formType === 'goal') updateSettings({ monthlyTarget: Number(data.monthlyTarget || 0) });
  if (formType === 'bi-filter') updateSettings({ biMonth: Number(data.biMonth || 0), biYear: Number(data.biYear || new Date().getFullYear()) });
  if (formType === 'bi-goals') updateSettings({
    monthlyTarget: Number(data.monthlyTarget || 0),
    annualTarget: Number(data.annualTarget || 0),
  });
  if (formType === 'bi-ticket') updateSettings({ monthlyTicket: Number(data.monthlyTicket || 0) });
  if (formType === 'decision') addDecision(data);
  if (formType === 'schedule') addSchedule(data);
  if (formType === 'inventory') updateSettings({
    computersTotal: Number(data.computersTotal || 0),
    computersMaintenance: Number(data.computersMaintenance || 0),
  });
  if (formType === 'exam') addExam(data);

  if (resetAfterSubmit) form.reset();
  persist();
  rerichRows();
  render();
}

function handleClick(event) {
  const quickModule = event.target.closest('[data-quick-module]');
  if (quickModule) {
    if (quickModule.dataset.quickModule === 'financeiro' && !canSeeFinancial()) {
      toast('Financeiro restrito a Admin e Financeiro.');
      return;
    }
    activateModule(quickModule.dataset.quickModule);
    return;
  }

  const openButton = event.target.closest('[data-open-student]');
  if (openButton) {
    const row = state.allRows.find((item) => item.key === openButton.dataset.openStudent);
    if (row) showDrawer(row);
    return;
  }

  const retentionButton = event.target.closest('[data-open-retention]');
  if (retentionButton) {
    const row = state.allRows.find((item) => item.key === retentionButton.dataset.openRetention);
    if (row) showRetentionDrawer(row);
    return;
  }

  const deleteCourse = event.target.closest('[data-delete-course]');
  if (deleteCourse) {
    state.store.courses = state.store.courses.filter((course) => course.name !== deleteCourse.dataset.deleteCourse);
    persist();
    render();
    return;
  }

  const convertLead = event.target.closest('[data-convert-lead]');
  if (convertLead) {
    updateLeadStage(convertLead.dataset.convertLead, 'Matriculado');
    return;
  }

  const queueButton = event.target.closest('[data-queue-action]');
  if (queueButton) {
    queueAction(queueButton.dataset.queueAction, queueButton.dataset.queueId, queueButton.dataset.queueType);
  }
}

function handleChange(event) {
  const statusSelect = event.target.closest('[data-status-key]');
  if (statusSelect) {
    saveStudentOverride(statusSelect.dataset.statusKey, { status: statusSelect.value });
    return;
  }

  const paidCheck = event.target.closest('[data-enrollment-paid]');
  if (paidCheck) {
    if (!canSeeFinancial()) {
      paidCheck.checked = !paidCheck.checked;
      toast('Baixa de pagamento restrita a Admin e Financeiro.');
      return;
    }
    saveStudentOverride(paidCheck.dataset.enrollmentPaid, { enrollmentPaid: paidCheck.checked ? 'true' : 'false' });
    return;
  }

  const boletoCheck = event.target.closest('[data-lead-boleto]');
  if (boletoCheck) {
    if (!canSeeFinancial()) {
      boletoCheck.checked = !boletoCheck.checked;
      toast('Controle de boleto restrito a Admin e Financeiro.');
      return;
    }
    updateLeadFinance(boletoCheck.dataset.leadBoleto, { boletoSent: boletoCheck.checked });
    return;
  }

  const leadPayment = event.target.closest('[data-lead-payment]');
  if (leadPayment) {
    if (!canSeeFinancial()) {
      toast('Baixa de pagamento restrita a Admin e Financeiro.');
      render();
      return;
    }
    updateLeadFinance(leadPayment.dataset.leadPayment, { enrollmentPaymentStatus: leadPayment.value });
  }
}

function handleDragStart(event) {
  const card = event.target.closest('[data-lead-id]');
  if (!card) return;
  event.dataTransfer.effectAllowed = 'move';
  event.dataTransfer.setData('text/plain', card.dataset.leadId);
  card.classList.add('dragging');
}

function handleDragOver(event) {
  const column = event.target.closest('[data-lead-stage]');
  if (!column) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = 'move';
  column.classList.add('drag-over');
}

function handleDragLeave(event) {
  const column = event.target.closest('[data-lead-stage]');
  if (column && !column.contains(event.relatedTarget)) {
    column.classList.remove('drag-over');
  }
}

function handleDrop(event) {
  const column = event.target.closest('[data-lead-stage]');
  if (!column) return;
  event.preventDefault();
  document.querySelectorAll('.kanban-column.drag-over').forEach((item) => item.classList.remove('drag-over'));
  document.querySelectorAll('.lead-card.dragging').forEach((item) => item.classList.remove('dragging'));
  const leadId = event.dataTransfer.getData('text/plain');
  if (leadId) updateLeadStage(leadId, column.dataset.leadStage);
}

function clearLeadDragState() {
  document.querySelectorAll('.kanban-column.drag-over').forEach((item) => item.classList.remove('drag-over'));
  document.querySelectorAll('.lead-card.dragging').forEach((item) => item.classList.remove('dragging'));
}

function saveStudentOverride(key, data) {
  const existing = state.store.overrides[key] || {};
  state.store.overrides[key] = {
    ...existing,
    status: data.status || existing.status,
    contactPhone: data.contactPhone ?? existing.contactPhone,
    contactEmail: data.contactEmail ?? existing.contactEmail,
    followStatus: data.followStatus ?? existing.followStatus,
    note: data.note ?? existing.note,
    lastContact: data.contact ? `${new Date().toLocaleString('pt-BR')} · ${data.contact}` : existing.lastContact,
    enrollmentPaid:
      data.enrollmentPaid !== undefined
        ? data.enrollmentPaid === 'true' || data.enrollmentPaid === true
        : Boolean(existing.enrollmentPaid),
    enrollmentExempt:
      data.enrollmentExempt !== undefined
        ? data.enrollmentExempt === 'true' || data.enrollmentExempt === true
        : Boolean(existing.enrollmentExempt),
    boletoSent:
      data.boletoSent !== undefined
        ? data.boletoSent === 'true' || data.boletoSent === true
        : Boolean(existing.boletoSent),
    boletoSentAt:
      data.boletoSent === 'true' && !existing.boletoSentAt
        ? new Date().toISOString()
        : existing.boletoSentAt || '',
    enrollmentFee: data.enrollmentFee ?? existing.enrollmentFee,
    updatedAt: new Date().toISOString(),
  };
  persist();
  rerichRows();
  toast('Alteração local salva e preservada contra sincronizações.');
  render();
}

function saveRetentionContact(key, data) {
  state.store.retention[key] = {
    ...(state.store.retention[key] || {}),
    contacted: data.contacted === 'true',
    contactDate: data.contactDate || new Date().toLocaleDateString('pt-BR'),
    channel: data.channel || '',
    responsible: data.responsible || roleLabel(state.profile),
    reason: data.reason || '',
    note: data.note || '',
    updatedAt: new Date().toISOString(),
  };
  persist();
  rerichRows();
  toast('Contato de retenção registrado.');
  render();
}

function addCourse(data) {
  const name = cleanText(data.name);
  if (!name) return;
  if (getCourseCatalog().some((course) => normalize(course.name) === normalize(name))) {
    toast('Curso já existe no catálogo.');
    return;
  }
  state.store.courses.push({ name, modality: data.modality || 'EAD', local: true });
  toast('Curso adicionado ao catálogo local.');
}

function addLead(data) {
  const lead = normalizeLead({
    id: cryptoId(),
    name: cleanText(data.name),
    phone: cleanText(data.phone),
    course: cleanText(data.course),
    origin: cleanText(data.origin) || 'Não informado',
    stage: data.stage || 'Lead',
    consultant: roleLabel(state.profile),
    createdAt: new Date().toISOString(),
  });
  if (lead.stage === 'Matriculado') {
    ensureLocalStudentForLead(lead);
    lead.matriculatedAt = new Date().toISOString();
  }
  state.store.leads.push(lead);
  toast(lead.stage === 'Matriculado' ? matriculationAlertMessage(lead) : 'Lead salvo no funil.');
}

function convertLeadToStudent(id) {
  updateLeadStage(id, 'Matriculado');
}

function updateLeadStage(id, stage) {
  const lead = state.store.leads.find((item) => item.id === id);
  if (!lead) return;
  const oldStage = lead.stage;
  lead.stage = stage;
  lead.updatedAt = new Date().toISOString();
  if (stage === 'Matriculado') {
    if (oldStage !== 'Matriculado') lead.matriculatedAt = new Date().toISOString();
    ensureLocalStudentForLead(lead);
    toast(matriculationAlertMessage(lead));
  } else {
    toast(`Lead movido para ${stage}.`);
  }
  persist();
  render();
}

function updateLeadFinance(id, changes) {
  const lead = state.store.leads.find((item) => item.id === id);
  if (!lead) return;
  if (changes.boletoSent !== undefined) {
    lead.boletoSent = Boolean(changes.boletoSent);
    lead.boletoSentAt = lead.boletoSent ? lead.boletoSentAt || new Date().toISOString() : '';
  }
  if (changes.enrollmentPaymentStatus !== undefined) {
    lead.enrollmentPaymentStatus = changes.enrollmentPaymentStatus || 'Pendente';
  }
  lead.updatedAt = new Date().toISOString();
  persist();
  toast('Status da matrícula atualizado no CRM.');
  render();
}

function ensureLocalStudentForLead(lead) {
  if (lead.localStudentId && state.store.localStudents.some((student) => student.id === lead.localStudentId)) return;
  const localStudent = {
    id: cryptoId(),
    name: lead.name,
    phone: lead.phone,
    course: lead.course,
    startPeriod: currentCohortCode(),
    status: 'PREMAT',
    createdAt: new Date().toISOString(),
  };
  state.store.localStudents.push(localStudent);
  lead.localStudentId = localStudent.id;
}

function matriculationAlertMessage(lead) {
  return lead.boletoSent
    ? 'Lead matriculado: contrato assinado e boleto já marcado como enviado.'
    : `Lead matriculado: confira o envio do boleto de matrícula para ${lead.name}.`;
}

function addDecision(data) {
  state.store.decisions.push({
    id: cryptoId(),
    area: cleanText(data.area),
    title: cleanText(data.title),
    owner: cleanText(data.owner),
    due: cleanText(data.due),
    status: cleanText(data.status) || 'Aberta',
    createdAt: new Date().toISOString(),
  });
  toast('Decisão registrada para acompanhamento gerencial.');
}

function addSchedule(data) {
  const eventItem = {
    id: cryptoId(),
    teacher: cleanText(data.teacher),
    subject: cleanText(data.subject),
    start: data.start,
    end: data.end,
    room: cleanText(data.room),
    capacity: Number(data.capacity || 0),
    status: 'agendado',
  };
  const conflict = scheduleConflict(eventItem);
  if (conflict) {
    toast(conflict);
    return;
  }
  state.store.schedule.push(eventItem);
  toast('Aula agendada na agenda única.');
}

function addExam(data) {
  const exam = {
    id: cryptoId(),
    student: cleanText(data.student),
    discipline: cleanText(data.discipline),
    start: data.start,
    duration: Number(data.duration || 90),
    machines: Number(data.machines || 1),
    status: 'agendado',
  };
  const conflict = examConflict(exam);
  if (conflict) {
    toast(conflict);
    return;
  }
  state.store.exams.push(exam);
  toast('Prova reservada sem overbooking.');
}

function queueAction(action, id, type) {
  const collection = type === 'exam' ? state.store.exams : state.store.schedule;
  const item = collection.find((entry) => entry.id === id);
  if (!item) return;

  if (action === 'present') item.status = 'presente';
  if (action === 'absent') item.status = 'ausente';
  if (action === 'cancel') {
    const reason = window.prompt('Informe a justificativa para cancelar/remanejar:');
    if (!reason) return;
    item.status = 'cancelada';
    item.reason = reason;
  }
  if (action === 'finish') {
    state.store.archive.push({ ...item, type, finishedAt: new Date().toISOString() });
    const index = collection.findIndex((entry) => entry.id === id);
    collection.splice(index, 1);
  }

  persist();
  render();
}

function parseCsv(text) {
  const rows = [];
  let current = [];
  let value = '';
  let inQuotes = false;

  for (let index = 0; index < text.length; index++) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && inQuotes && next === '"') {
      value += '"';
      index++;
      continue;
    }
    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (char === ',' && !inQuotes) {
      current.push(value);
      value = '';
      continue;
    }
    if ((char === '\n' || char === '\r') && !inQuotes) {
      if (char === '\r' && next === '\n') index++;
      current.push(value);
      rows.push(current);
      current = [];
      value = '';
      continue;
    }
    value += char;
  }

  if (value || current.length) {
    current.push(value);
    rows.push(current);
  }

  const headers = rows.shift()?.map(cleanText) ?? [];
  return {
    headers,
    rows: rows
      .filter((row) => row.some(Boolean))
      .map((row) =>
        headers.reduce((record, header, index) => {
          record[header] = cleanText(row[index] ?? '');
          return record;
        }, {}),
      ),
  };
}

function rowKey(row, index) {
  return cleanText(row[fields.ra]) || cleanText(row[fields.cpf]) || `row-${index}`;
}

function getRisk(row) {
  let score = 0;
  if (row.isDebt) score += 35;
  if (hasNoAccess(row)) score += 28;
  if (hasLongNoAccess(row)) score += 18;
  if (row.debtValue >= 500) score += 10;
  if (isInactiveStatus(row)) score += 14;
  return {
    score,
    level: score >= 55 ? 'Alto' : score >= 25 ? 'Médio' : 'Baixo',
  };
}

function hasNoAccess(row) {
  return isNo(row.oneAccess) || isNo(row.avaAccess) || noValue(row.oneAccess) || noValue(row.avaAccess);
}

function hasLongNoAccess(row) {
  return normalize(`${row.oneDays} ${row.avaDays}`).includes('mais de 60') ||
    normalize(`${row.oneDays} ${row.avaDays}`).includes('31 a 60');
}

function isInactiveStatus(row) {
  return ['tranca', 'cancel', 'desist', 'aband'].some((item) => normalize(row.localStatus).includes(item));
}

function isActive(row) {
  return ['ativo', 'premat'].some((item) => normalize(row.localStatus).includes(item));
}

function canSeeFinancial() {
  return state.profile === 'admin' || state.profile === 'financeiro';
}

function getBiFilter() {
  const today = new Date();
  return {
    month: Number(state.store.settings.biMonth || today.getMonth() + 1),
    year: Number(state.store.settings.biYear || today.getFullYear()),
  };
}

function monthOptions(selected) {
  return [
    '<option value="0">Ano inteiro</option>',
    ...MONTHS.map((month, index) => `<option value="${index + 1}" ${Number(selected) === index + 1 ? 'selected' : ''}>${month}</option>`),
  ].join('');
}

function yearOptions(selected) {
  const current = new Date().getFullYear();
  return Array.from({ length: 6 }, (_, index) => current - 3 + index)
    .map((year) => `<option value="${year}" ${Number(selected) === year ? 'selected' : ''}>${year}</option>`)
    .join('');
}

function periodLabel(filter) {
  return filter.month ? `${MONTHS[filter.month - 1]}/${String(filter.year).slice(-2)}` : `Ano ${filter.year}`;
}

function filterRowsByPeriod(rows, filter) {
  return rows.filter((row) => {
    const date = rowPeriodDate(row);
    if (!date) return filter.month === 0;
    return matchesPeriod(date, filter);
  });
}

function rowPeriodDate(row) {
  return parseBrazilianDate(row.enrollmentDate) || cohortDate(row.startPeriod);
}

function cohortDate(value) {
  const cohort = parseCohort(value);
  if (!cohort.year) return null;
  const month = cohort.term === 2 ? 6 : 0;
  return new Date(cohort.year, month, 1);
}

function matchesPeriod(date, filter) {
  if (!date || Number.isNaN(date.getTime())) return false;
  if (date.getFullYear() !== Number(filter.year)) return false;
  return !Number(filter.month) || date.getMonth() + 1 === Number(filter.month);
}

function previousMonthFilter(filter) {
  const month = Number(filter.month || new Date().getMonth() + 1);
  const year = Number(filter.year);
  return month === 1 ? { month: 12, year: year - 1 } : { month: month - 1, year };
}

function previousYearFilter(filter) {
  return { month: Number(filter.month || 0), year: Number(filter.year) - 1 };
}

function academicCensus(rows) {
  const definitions = [
    { key: 'Ativo', label: 'Ativo', className: 'green' },
    { key: 'Abandonado', label: 'Abandonado', className: 'yellow' },
    { key: 'Trancado', label: 'Trancado', className: 'cyan' },
    { key: 'Cancelado', label: 'Cancelado', className: 'red' },
  ];
  const counts = definitions.map((definition) => ({
    ...definition,
    count: rows.filter((row) => academicStatus(row) === definition.key).length,
  }));
  const max = Math.max(...counts.map((item) => item.count), 1);
  return {
    active: counts.find((item) => item.key === 'Ativo')?.count || 0,
    inactive: counts.filter((item) => item.key !== 'Ativo').reduce((total, item) => total + item.count, 0),
    statuses: counts.map((item) => ({ ...item, height: Math.max(8, Math.round((item.count / max) * 100)) })),
  };
}

function academicStatus(row) {
  const status = normalize(row.localStatus);
  if (status.includes('tranca')) return 'Trancado';
  if (status.includes('aband') || status.includes('desist')) return 'Abandonado';
  if (status.includes('cancel')) return 'Cancelado';
  return isActive(row) ? 'Ativo' : 'Cancelado';
}

function commercialKpis(filter) {
  const monthlyFilter = { month: Number(filter.month || new Date().getMonth() + 1), year: Number(filter.year) };
  return {
    monthlyMatriculations: matriculationsInPeriod(monthlyFilter),
    annualMatriculations: matriculationsInPeriod({ month: 0, year: filter.year }),
  };
}

function matriculationsInPeriod(filter) {
  const leadCount = state.store.leads.filter((lead) => {
    if (lead.stage !== 'Matriculado') return false;
    const date = parseBrazilianDate(lead.matriculatedAt || lead.updatedAt || lead.createdAt);
    return matchesPeriod(date, filter);
  }).length;
  const localStudentCount = state.store.localStudents.filter((student) => {
    const date = parseBrazilianDate(student.createdAt);
    return matchesPeriod(date, filter);
  }).length;
  return leadCount + localStudentCount;
}

function financeKpis(filter) {
  const previousMonth = previousMonthFilter(filter);
  const previousYear = previousYearFilter(filter);
  return {
    enrollment: {
      current: enrollmentRevenue(filter),
      previousMonth: enrollmentRevenue(previousMonth),
      previousYear: enrollmentRevenue(previousYear),
    },
    recurring: {
      current: recurringRevenue(filter),
      previousMonth: recurringRevenue(previousMonth),
      previousYear: recurringRevenue(previousYear),
    },
    ytd: {
      current: enrollmentRevenue({ month: 0, year: filter.year }) + recurringRevenue({ month: 0, year: filter.year }),
      previousMonth: enrollmentRevenue(previousMonth) + recurringRevenue(previousMonth),
      previousYear: enrollmentRevenue({ month: 0, year: filter.year - 1 }) + recurringRevenue({ month: 0, year: filter.year - 1 }),
    },
  };
}

function enrollmentRevenue(filter) {
  const fee = Number(state.store.settings.enrollmentFee || 99);
  const overrideRevenue = Object.values(state.store.overrides).reduce((total, item) => {
    if (!item.enrollmentPaid) return total;
    const date = parseBrazilianDate(item.updatedAt || item.boletoSentAt);
    if (!matchesPeriod(date, filter)) return total;
    return total + Number(item.enrollmentFee || fee);
  }, 0);
  const leadRevenue = state.store.leads.reduce((total, lead) => {
    if (lead.stage !== 'Matriculado' || lead.enrollmentPaymentStatus !== 'Pago') return total;
    const date = parseBrazilianDate(lead.updatedAt || lead.matriculatedAt || lead.createdAt);
    if (!matchesPeriod(date, filter)) return total;
    return total + fee;
  }, 0);
  return overrideRevenue + leadRevenue;
}

function recurringRevenue(filter) {
  const ticket = Number(state.store.settings.monthlyTicket || 299);
  return filterRowsByPeriod(state.allRows, filter).filter((row) => isActive(row) && !row.isDebt).length * ticket;
}

function percent(value, target) {
  return Math.min(999, Math.round((Number(value || 0) / Math.max(Number(target || 0), 1)) * 100));
}

function progressLine(label, progress, detail) {
  return `
    <div class="progress-line">
      <div>
        <strong>${escapeHtml(label)}</strong>
        <span>${escapeHtml(detail)}</span>
      </div>
      <em>${progress}%</em>
      <div class="bar-track"><span style="width:${Math.min(progress, 100)}%"></span></div>
    </div>
  `;
}

function financeKpiItem(label, current, previousMonth, previousYear) {
  return `
    <div class="finance-kpi-item">
      <div>
        <strong>${escapeHtml(label)}</strong>
        <span>Atual: ${formatMoney(current)}</span>
      </div>
      <span class="badge ${current >= previousMonth ? 'green' : 'yellow'}">Mês anterior: ${signedPercentChange(current, previousMonth)}</span>
      <span class="badge ${current >= previousYear ? 'green' : 'red'}">Ano anterior: ${signedPercentChange(current, previousYear)}</span>
    </div>
  `;
}

function signedPercentChange(current, previous) {
  if (!previous) return current ? '+100%' : '0%';
  const change = Math.round(((current - previous) / Math.abs(previous)) * 100);
  return `${change >= 0 ? '+' : ''}${change}%`;
}

function getStudentTotals(rows) {
  return {
    active: rows.filter(isActive).length,
    overrides: rows.filter((row) => row.statusOverridden || row.contactOverridden).length,
    highRisk: rows.filter((row) => row.risk.level === 'Alto').length,
  };
}

function getCohortStats(rows) {
  const grouped = groupBy(rows, (row) => row.startPeriod || '-');
  return Object.entries(grouped)
    .map(([period, items]) => {
      const active = items.filter(isActive).length;
      return { period, total: items.length, retention: items.length ? Math.round((active / items.length) * 100) : 0 };
    })
    .sort((a, b) => b.total - a.total);
}

function getCourseCatalog() {
  const sheetCourses = Object.entries(groupBy(state.allRows, (row) => row.course)).map(([name, rows]) => ({
    name,
    modality: inferModality(name),
    students: rows.length,
    local: false,
  }));
  const localCourses = state.store.courses.map((course) => ({
    ...course,
    students: state.allRows.filter((row) => normalize(row.course) === normalize(course.name)).length,
    local: true,
  }));
  const merged = new Map();
  [...sheetCourses, ...localCourses].forEach((course) => merged.set(normalize(course.name), course));
  return [...merged.values()].sort((a, b) => collator.compare(a.name, b.name));
}

function leadColumn(stage) {
  const leads = state.store.leads.filter((lead) => lead.stage === stage);
  return `
    <div class="kanban-column" data-lead-stage="${escapeHtml(stage)}">
      <div class="kanban-head"><strong>${escapeHtml(stage)}</strong><span>${leads.length}</span></div>
      ${leads
        .map(
          (lead) => `
            <div class="lead-card ${lead.stage === 'Matriculado' && !lead.boletoSent ? 'needs-boleto' : ''}" draggable="true" data-lead-id="${escapeHtml(lead.id)}">
              <strong>${escapeHtml(lead.name)}</strong>
              <span>${escapeHtml(lead.course)}</span>
              <small>${escapeHtml(lead.origin)} · ${escapeHtml(lead.consultant || 'Consultor')}</small>
              ${
                stage === 'Matriculado'
                  ? canSeeFinancial()
                    ? `
                    <div class="lead-finance-box">
                      <label class="check-line">
                        <input type="checkbox" data-lead-boleto="${escapeHtml(lead.id)}" ${lead.boletoSent ? 'checked' : ''} />
                        Boleto enviado
                      </label>
                      <label class="lead-payment-field">
                        Pagamento da matrícula
                        <select data-lead-payment="${escapeHtml(lead.id)}">
                          ${paymentOptions(lead.enrollmentPaymentStatus)}
                        </select>
                      </label>
                      <small>${lead.boletoSentAt ? `Enviado em ${escapeHtml(new Date(lead.boletoSentAt).toLocaleString('pt-BR'))}` : 'Alerta ativo até o envio do boleto.'}</small>
                    </div>
                  `
                    : `
                    <div class="lead-finance-box">
                      <span class="badge ${lead.boletoSent ? 'green' : 'yellow'}">Boleto ${lead.boletoSent ? 'enviado' : 'pendente'}</span>
                      <span class="badge">Pagamento ${escapeHtml(lead.enrollmentPaymentStatus)}</span>
                      <small>Baixa financeira restrita a Admin e Financeiro.</small>
                    </div>
                  `
                  : `<button type="button" data-convert-lead="${escapeHtml(lead.id)}">Matricular</button>`
              }
            </div>
          `,
        )
        .join('')}
    </div>
  `;
}

function paymentOptions(selected = 'Pendente') {
  return ['Pendente', 'Pago', 'Isento']
    .map((status) => `<option ${status === selected ? 'selected' : ''}>${status}</option>`)
    .join('');
}

function scheduleTemplate(item) {
  return `
    <div class="timeline-item">
      <div>
        <strong>${escapeHtml(item.subject)}</strong>
        <span>${formatDateTime(item.start)} - ${formatTime(item.end)} · ${escapeHtml(item.room)}</span>
      </div>
      <span class="badge ${item.status === 'cancelada' ? 'red' : 'cyan'}">${escapeHtml(item.teacher)}</span>
    </div>
  `;
}

function examTemplate(item) {
  return `
    <div class="timeline-item">
      <div>
        <strong>${escapeHtml(item.student)}</strong>
        <span>${escapeHtml(item.discipline)} · ${formatDateTime(item.start)} · ${item.duration} min</span>
      </div>
      <span class="badge yellow">${item.machines} estação</span>
    </div>
  `;
}

function queueTemplate(item) {
  return `
    <div class="queue-item ${item.soon ? 'soon' : ''}">
      <div>
        <span class="queue-time">${formatTime(item.start)}</span>
        <strong>${escapeHtml(item.title)}</strong>
        <small>${escapeHtml(item.subtitle)}</small>
      </div>
      <div class="queue-actions">
        <button type="button" data-queue-action="present" data-queue-id="${escapeHtml(item.id)}" data-queue-type="${item.type}">Presente</button>
        <button type="button" data-queue-action="absent" data-queue-id="${escapeHtml(item.id)}" data-queue-type="${item.type}">Ausente</button>
        <button type="button" data-queue-action="cancel" data-queue-id="${escapeHtml(item.id)}" data-queue-type="${item.type}">Cancelar</button>
        <button type="button" data-queue-action="finish" data-queue-id="${escapeHtml(item.id)}" data-queue-type="${item.type}">Finalizar</button>
      </div>
    </div>
  `;
}

function dailyQueue() {
  const now = new Date();
  const today = now.toDateString();
  const lessons = state.store.schedule.map((item) => ({
    id: item.id,
    type: 'schedule',
    title: item.subject,
    subtitle: `${item.teacher} · ${item.room}`,
    start: item.start,
    status: item.status,
  }));
  const exams = state.store.exams.map((item) => ({
    id: item.id,
    type: 'exam',
    title: item.student,
    subtitle: `${item.discipline} · ${item.duration} min`,
    start: item.start,
    status: item.status,
  }));
  return [...lessons, ...exams]
    .filter((item) => new Date(item.start).toDateString() === today)
    .map((item) => {
      const minutes = (new Date(item.start) - now) / 60000;
      return { ...item, soon: minutes >= 0 && minutes <= 30 };
    })
    .sort((a, b) => new Date(a.start) - new Date(b.start));
}

function scheduleConflict(candidate) {
  if (new Date(candidate.end) <= new Date(candidate.start)) return 'O fim deve ser posterior ao início.';
  if (candidate.capacity <= 0) return 'Informe uma lotação válida.';
  const conflicts = state.store.schedule.filter((item) => overlaps(candidate.start, candidate.end, item.start, item.end));
  if (conflicts.some((item) => normalize(item.teacher) === normalize(candidate.teacher))) return 'Conflito de docência: professor já alocado neste horário.';
  if (conflicts.some((item) => normalize(item.room) === normalize(candidate.room))) return 'Conflito de espaço: sala já ocupada neste horário.';
  return '';
}

function examConflict(candidate) {
  const total = Number(state.store.settings.computersTotal || 0);
  const maintenance = Number(state.store.settings.computersMaintenance || 0);
  const end = addMinutes(candidate.start, candidate.duration);
  const reserved = state.store.exams
    .filter((exam) => overlaps(candidate.start, end, exam.start, addMinutes(exam.start, exam.duration)))
    .reduce((sum, exam) => sum + Number(exam.machines || 1), 0);
  if (reserved + candidate.machines > total - maintenance) return 'Overbooking bloqueado: computadores insuficientes neste horário.';
  return '';
}

function currentExamReservations() {
  const now = new Date();
  return state.store.exams
    .filter((exam) => now >= new Date(exam.start) && now <= new Date(addMinutes(exam.start, exam.duration)))
    .reduce((sum, exam) => sum + Number(exam.machines || 1), 0);
}

function updateSettings(values) {
  state.store.settings = { ...state.store.settings, ...values };
  toast('Parâmetro atualizado.');
}

function consultantRanking() {
  const grouped = groupBy(state.store.leads, (lead) => lead.consultant || 'Consultor');
  const rows = Object.entries(grouped).map(([name, leads]) => {
    const converted = leads.filter((lead) => lead.stage === 'Matriculado').length;
    return { name, total: leads.length, conversion: leads.length ? Math.round((converted / leads.length) * 100) : 0 };
  });
  return rows.length ? rows.sort((a, b) => b.total - a.total) : [{ name: 'Sem dados locais', total: 0, conversion: 0 }];
}

function decisionSignals({ rows, debtRows, debtTotal, retention, gap, availableComputers, queue }) {
  const signals = [];
  const highRisk = rows.filter((row) => row.risk.level === 'Alto').length;
  const noAccess = rows.filter(hasNoAccess).length;

  if (highRisk) {
    signals.push({
      tone: 'warning',
      title: 'Retenção exige intervenção',
      text: `${highRisk.toLocaleString('pt-BR')} alunos estão em risco alto. Priorize contato acadêmico e financeiro.`,
    });
  }

  if (debtRows.length && canSeeFinancial()) {
    signals.push({
      tone: 'danger',
      title: 'Carteira em atraso impacta caixa',
      text: `${debtRows.length.toLocaleString('pt-BR')} alunos somam ${formatMoney(debtTotal)} em atraso.`,
    });
  }

  if (noAccess) {
    signals.push({
      tone: 'warning',
      title: 'Acesso digital fragilizado',
      text: `${noAccess.toLocaleString('pt-BR')} alunos sem uso pleno de One/AVA. Acione apoio de acesso.`,
    });
  }

  if (gap < 0) {
    signals.push({
      tone: 'danger',
      title: 'Meta comercial abaixo do necessário',
      text: `Faltam ${Math.abs(gap).toLocaleString('pt-BR')} matrículas para atingir a meta mensal.`,
    });
  }

  if (availableComputers <= 4) {
    signals.push({
      tone: 'danger',
      title: 'Gargalo de infraestrutura',
      text: `Há apenas ${availableComputers} computadores úteis. Replaneje provas ou manutenção.`,
    });
  }

  if (queue.some((item) => item.soon)) {
    signals.push({
      tone: 'info',
      title: 'Operação diária próxima',
      text: 'Existem eventos nos próximos 30 minutos. Prepare sala, docente ou estação de prova.',
    });
  }

  if (retention >= 75 && gap >= 0 && !signals.length) {
    signals.push({
      tone: 'success',
      title: 'Operação saudável',
      text: 'Retenção, meta e recursos estão em patamar confortável nesta visão.',
    });
  }

  return signals.length
    ? signals
    : [
        {
          tone: 'info',
          title: 'Sem sinal crítico',
          text: 'Continue acompanhando safra, carteira e agenda para antecipar riscos.',
        },
      ];
}

function debtEvolution(rows) {
  const grouped = groupBy(rows, (row) => row.cohort.year ? `${row.cohort.year}.${row.cohort.term}` : 'Sem safra');
  const entries = Object.entries(grouped)
    .map(([label, items]) => ({ label, value: sumBy(items, (row) => row.debtValue) }))
    .sort((a, b) => collator.compare(a.label, b.label))
    .slice(-7);
  const max = Math.max(...entries.map((item) => item.value), 1);
  return entries.map((item) => ({ ...item, height: Math.max(8, Math.round((item.value / max) * 100)) }));
}

function pieChart(paid, debt) {
  const total = Math.max(1, paid + debt);
  const debtPercent = Math.round((debt / total) * 100);
  const paidPercent = 100 - debtPercent;
  return `
    <svg class="pie-chart" viewBox="0 0 36 36" aria-label="Composição de carteira">
      <circle cx="18" cy="18" r="15.9" fill="transparent" stroke="#dff5ec" stroke-width="4"></circle>
      <circle cx="18" cy="18" r="15.9" fill="transparent" stroke="#d84f45" stroke-width="4" stroke-dasharray="${debtPercent} ${paidPercent}" stroke-dashoffset="25"></circle>
      <text x="18" y="19" text-anchor="middle">${debtPercent}%</text>
    </svg>
  `;
}

function accessDenied(title, message) {
  return `
    ${moduleTitle(title, message)}
    <section class="locked-panel">
      <strong>Acesso blindado pelo RBAC</strong>
      <p>${escapeHtml(message)}</p>
    </section>
  `;
}

function moduleTitle(title, subtitle) {
  return `
    <header class="module-title">
      <div>
        <p class="eyebrow">ERP Educacional</p>
        <h1>${escapeHtml(title)}</h1>
        <span>${escapeHtml(subtitle)}</span>
      </div>
    </header>
  `;
}

function metricCard(label, value, subtitle, tone = '') {
  return `
    <article class="metric-card ${tone ? `tone-${tone}` : ''}">
      <span>${escapeHtml(label)}</span>
      <strong>${typeof value === 'number' ? value.toLocaleString('pt-BR') : escapeHtml(value)}</strong>
      <small>${escapeHtml(subtitle)}</small>
    </article>
  `;
}

function quickAccessCard(title, subtitle, module, value, label, tone) {
  return `
    <button class="quick-access-card tone-${tone}" type="button" data-quick-module="${escapeHtml(module)}">
      <span>${escapeHtml(title)}</span>
      <strong>${typeof value === 'number' ? value.toLocaleString('pt-BR') : escapeHtml(value)}</strong>
      <small>${escapeHtml(label)}</small>
      <em>${escapeHtml(subtitle)}</em>
    </button>
  `;
}

function smallTitle(eyebrow, title) {
  return `<p class="eyebrow">${escapeHtml(eyebrow)}</p><h2>${escapeHtml(title)}</h2>`;
}

function detailRow(label, value) {
  return `<div class="detail-row"><span>${escapeHtml(label)}</span><strong>${escapeHtml(cleanText(value) || '-')}</strong></div>`;
}

function roleCard(title, description, active) {
  return `<div class="role-card ${active ? 'active' : ''}"><strong>${escapeHtml(title)}</strong><span>${escapeHtml(description)}</span></div>`;
}

function userRowTemplate(user) {
  const active = normalize(user.ativo || 'SIM') !== 'nao';
  return `
    <tr>
      <td><strong>${escapeHtml(user.usuario || '-')}</strong></td>
      <td>${escapeHtml(user.nome || '-')}</td>
      <td><span class="badge cyan">${escapeHtml(roleLabel(normalize(user.perfil || 'consultor')))}</span></td>
      <td><span class="badge ${active ? 'green' : 'red'}">${active ? 'Ativo' : 'Inativo'}</span></td>
      <td>Aba ERP_Usuarios</td>
    </tr>
  `;
}

function auditItems() {
  return Object.entries(state.store.overrides)
    .slice(-12)
    .reverse()
    .map(([key, value]) => {
      const row = state.allRows.find((item) => item.key === key);
      return `<div><strong>${escapeHtml(row?.name || key)}</strong><span>${escapeHtml(value.updatedAt ? new Date(value.updatedAt).toLocaleString('pt-BR') : '-')} · ${escapeHtml(value.status || 'status preservado')}</span></div>`;
    });
}

function emptyRow(message) {
  return `<tr><td colspan="8"><div class="empty-state">${escapeHtml(message)}</div></td></tr>`;
}

function parseCohort(value) {
  const match = cleanText(value).match(/^([A-Z])\.(\d{4})\.(\d)([A-Z])?/i);
  return match
    ? { modality: match[1].toUpperCase(), year: Number(match[2]), term: Number(match[3]), wave: match[4] || '' }
    : { modality: '', year: 0, term: 0, wave: '' };
}

function getFirstValue(row, names) {
  for (const name of names) {
    if (row[name]) return cleanText(row[name]);
  }
  return '';
}

function getAvaDaysNumber(row) {
  const exactDate = parseBrazilianDate(row.avaLastAccessDate);
  if (exactDate) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    exactDate.setHours(0, 0, 0, 0);
    return Math.max(0, Math.floor((today - exactDate) / 86400000));
  }

  const text = normalize(row.avaDays);
  if (!text || text === '-') return isNo(row.avaAccess) ? 999 : -1;
  if (text.includes('0 a 07') || text.includes('0 a 7')) return 4;
  if (text.includes('08 a 30') || text.includes('8 a 30')) return 8;
  if (text.includes('31 a 60')) return 31;
  if (text.includes('mais de 60')) return 61;

  const number = text.match(/\d+/);
  return number ? Number(number[0]) : -1;
}

function getAvaAlert(days, accessText) {
  if (isNo(accessText) || days >= 8) {
    return { level: 'red', label: 'Vermelho', className: 'red', weight: 3 };
  }
  if (days >= 5) {
    return { level: 'yellow', label: 'Amarelo', className: 'yellow', weight: 2 };
  }
  if (days >= 0) {
    return { level: 'ok', label: 'Ok', className: 'green', weight: 1 };
  }
  return { level: 'unknown', label: 'Sem dado', className: 'cyan', weight: 0 };
}

function parseBrazilianDate(value) {
  const text = cleanText(value);
  if (!text || text === '-') return null;

  const iso = new Date(text);
  if (!Number.isNaN(iso.getTime())) return iso;

  const numeric = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);
  if (numeric) {
    const year = Number(numeric[3].length === 2 ? `20${numeric[3]}` : numeric[3]);
    return new Date(year, Number(numeric[2]) - 1, Number(numeric[1]));
  }

  const monthNames = {
    jan: 0,
    fev: 1,
    mar: 2,
    abr: 3,
    mai: 4,
    jun: 5,
    jul: 6,
    ago: 7,
    set: 8,
    out: 9,
    nov: 10,
    dez: 11,
  };
  const longMatch = normalize(text).match(/(\d{1,2})\s+de\s+([a-z]{3})\.?\s+de\s+(\d{4})/);
  if (longMatch && monthNames[longMatch[2]] !== undefined) {
    return new Date(Number(longMatch[3]), monthNames[longMatch[2]], Number(longMatch[1]));
  }

  return null;
}

function inferModality(name) {
  const normalized = normalize(name);
  if (normalized.includes('semipresencial')) return 'Semipresencial';
  if (normalized.includes('pos')) return 'Pós-graduação';
  if (normalized.includes('ead')) return 'EAD';
  return 'Presencial';
}

function currentCohortCode() {
  const now = new Date();
  const term = now.getMonth() < 6 ? 1 : 2;
  return `E.${now.getFullYear()}.${term}L`;
}

function populateFilter(select, values, defaultLabel) {
  const current = select.value;
  select.innerHTML = `<option value="">${defaultLabel}</option>${values.map((value) => `<option>${escapeHtml(value)}</option>`).join('')}`;
  if (values.includes(current)) select.value = current;
}

function uniqueValues(rows, getter) {
  return [...new Set(rows.map(getter).filter(Boolean))].sort(collator.compare).slice(0, 400);
}

function rerichRows() {
  state.allRows = state.allRows.map((row, index) => enrichRow(row.raw, index));
}

function seedStore() {
  state.store = normalizeStore(state.store);
  if (!state.store.settings.enrollmentFee) state.store.settings.enrollmentFee = 99;
  if (!state.store.settings.monthlyTarget) state.store.settings.monthlyTarget = 65;
  if (!state.store.settings.annualTarget) state.store.settings.annualTarget = Number(state.store.settings.monthlyTarget || 65) * 12;
  if (!state.store.settings.monthlyTicket) state.store.settings.monthlyTicket = 299;
  if (!state.store.settings.computersTotal) state.store.settings.computersTotal = 24;
  if (state.store.settings.computersMaintenance === undefined) state.store.settings.computersMaintenance = 2;

  if (!state.store.schedule.length) {
    const today = new Date();
    today.setHours(15, 0, 0, 0);
    const end = new Date(today);
    end.setHours(17, 0, 0, 0);
    state.store.schedule.push({
      id: cryptoId(),
      teacher: 'Carlos Andrade',
      subject: 'Aula presencial - Administração',
      start: toLocalInputValue(today),
      end: toLocalInputValue(end),
      room: 'Sala 02',
      capacity: 35,
      status: 'agendado',
    });
  }
}

function loadStore() {
  try {
    return normalizeStore(JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}'));
  } catch {
    return defaultStore();
  }
}

function persist() {
  persistLocal();
  queueRemotePersist();
}

function persistLocal() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.store));
}

function queueRemotePersist() {
  if (!state.remoteState) return;
  window.clearTimeout(state.remoteSaveTimer);
  state.remoteSaveTimer = window.setTimeout(saveOperationalStore, 450);
}

async function saveOperationalStore() {
  if (!state.remoteState) return;
  try {
    const response = await fetch('/.netlify/functions/state', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ actor: roleLabel(state.profile), state: state.store }),
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    if (!payload.ok) throw new Error(payload.message || 'Falha ao salvar');
    els.syncStamp.textContent = `Estado salvo na planilha: ${new Date().toLocaleString('pt-BR')}`;
  } catch (error) {
    state.remoteState = false;
    if (!state.remoteWarningShown) {
      state.remoteWarningShown = true;
      toast(`Falha ao salvar na planilha. Mantendo fallback local: ${error.message}`);
    }
  }
}

function normalizeStore(input = {}) {
  return {
    ...defaultStore(),
    ...input,
    overrides: input.overrides || {},
    retention: input.retention || {},
    courses: Array.isArray(input.courses) ? input.courses : [],
    leads: Array.isArray(input.leads) ? input.leads.map(normalizeLead) : [],
    localStudents: Array.isArray(input.localStudents) ? input.localStudents : [],
    schedule: Array.isArray(input.schedule) ? input.schedule : [],
    exams: Array.isArray(input.exams) ? input.exams : [],
    archive: Array.isArray(input.archive) ? input.archive : [],
    decisions: Array.isArray(input.decisions) ? input.decisions : [],
    settings: input.settings || {},
  };
}

function normalizeLead(lead = {}) {
  const stage = LEAD_STAGES.includes(lead.stage) ? lead.stage : 'Lead';
  const paymentStatus = ['Pendente', 'Pago', 'Isento'].includes(lead.enrollmentPaymentStatus)
    ? lead.enrollmentPaymentStatus
    : 'Pendente';
  return {
    ...lead,
    id: lead.id || cryptoId(),
    name: cleanText(lead.name),
    phone: cleanText(lead.phone),
    course: cleanText(lead.course),
    origin: cleanText(lead.origin) || 'Não informado',
    stage,
    consultant: cleanText(lead.consultant) || 'Consultor',
    boletoSent: lead.boletoSent === true || String(lead.boletoSent).toLowerCase() === 'true',
    boletoSentAt: cleanText(lead.boletoSentAt),
    enrollmentPaymentStatus: paymentStatus,
    matriculatedAt: cleanText(lead.matriculatedAt),
    localStudentId: cleanText(lead.localStudentId),
    createdAt: cleanText(lead.createdAt) || new Date().toISOString(),
    updatedAt: cleanText(lead.updatedAt),
  };
}

function defaultStore() {
  return {
    overrides: {},
    retention: {},
    courses: [],
    leads: [],
    localStudents: [],
    schedule: [],
    exams: [],
    archive: [],
    decisions: [],
    settings: {},
  };
}

function defaultUsers() {
  return [
    { usuario: 'admin', senha: 'admin', nome: 'Administrador', perfil: 'admin', ativo: 'SIM' },
    {
      usuario: 'financeiro',
      senha: '123456',
      nome: 'Responsável Financeiro',
      perfil: 'financeiro',
      ativo: 'SIM',
    },
    {
      usuario: 'retencao',
      senha: '123456',
      nome: 'Responsável Retenção',
      perfil: 'consultor',
      ativo: 'SIM',
    },
    {
      usuario: 'consultor',
      senha: '123456',
      nome: 'Consultor Comercial',
      perfil: 'consultor',
      ativo: 'SIM',
    },
  ];
}

function setSourceStatus(text) {
  els.sourceStatus.textContent = text;
}

function toast(message) {
  els.toast.textContent = message;
  els.toast.classList.add('show');
  window.clearTimeout(toast.timer);
  toast.timer = window.setTimeout(() => els.toast.classList.remove('show'), 3200);
}

function rowByKey(key) {
  return state.allRows.find((row) => row.key === key);
}

function normalizePhone(value) {
  const digits = cleanText(value).replace(/\D/g, '');
  if (!digits) return '';
  return digits.startsWith('55') ? digits : `55${digits}`;
}

function whatsAppUrl(row) {
  const message = `Olá, ${firstName(row.name)}! Tudo bem? Aqui é da UniFECAF. Estou entrando em contato para acompanhar sua matrícula no curso ${row.course}.`;
  return `https://wa.me/${normalizePhone(row.phone)}?text=${encodeURIComponent(message)}`;
}

function firstName(value) {
  return cleanText(value).split(/\s+/)[0] || '';
}

function maskCpf(value) {
  const digits = cleanText(value).replace(/\D/g, '');
  if (digits.length < 5) return '-';
  return `***.${digits.slice(-5, -2)}.${digits.slice(-2)}`;
}

function statusClass(status) {
  const normalized = normalize(status);
  if (normalized.includes('ativo') || normalized.includes('premat')) return 'green';
  if (normalized.includes('tranca') || normalized.includes('aband')) return 'yellow';
  if (normalized.includes('cancel') || normalized.includes('desist')) return 'red';
  return 'cyan';
}

function isNo(value) {
  return normalize(value).startsWith('nao');
}

function noValue(value) {
  return !cleanText(value) || cleanText(value) === '-';
}

function parseMoney(value) {
  const text = cleanText(value);
  if (!text || text === '-') return 0;
  return Number.parseFloat(text.replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.')) || 0;
}

function formatMoney(value) {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(Number(value) || 0);
}

function formatCompactMoney(value) {
  return new Intl.NumberFormat('pt-BR', { notation: 'compact', style: 'currency', currency: 'BRL' }).format(Number(value) || 0);
}

function formatDateTime(value) {
  if (!value) return '-';
  return new Date(value).toLocaleString('pt-BR', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
}

function formatTime(value) {
  if (!value) return '-';
  return new Date(value).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
}

function toLocalInputValue(date) {
  const pad = (number) => String(number).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function toDateInputValue(value) {
  const parsed = parseBrazilianDate(value) || (value ? new Date(value) : new Date());
  if (Number.isNaN(parsed.getTime())) return '';
  const pad = (number) => String(number).padStart(2, '0');
  return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}`;
}

function addMinutes(value, minutes) {
  const date = new Date(value);
  date.setMinutes(date.getMinutes() + Number(minutes || 0));
  return date.toISOString();
}

function overlaps(startA, endA, startB, endB) {
  return new Date(startA) < new Date(endB) && new Date(endA) > new Date(startB);
}

function daysInCurrentMonth() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
}

function groupBy(rows, getter) {
  return rows.reduce((groups, row) => {
    const key = getter(row) || '-';
    groups[key] = groups[key] || [];
    groups[key].push(row);
    return groups;
  }, {});
}

function sumBy(rows, getter) {
  return rows.reduce((sum, row) => sum + Number(getter(row) || 0), 0);
}

function roleLabel(role) {
  return { admin: 'Admin', financeiro: 'Financeiro', consultor: 'Consultor' }[role] || role;
}

function shortUnit(value) {
  return cleanText(value).split('[')[0].trim() || '-';
}

function cleanText(value) {
  return String(value ?? '').trim();
}

function normalize(value) {
  return cleanText(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function escapeHtml(value) {
  return cleanText(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function cryptoId() {
  if (crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}
