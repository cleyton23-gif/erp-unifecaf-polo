const SHEET_ID = '1AoZ9KCNIaIzTEW17MyNFv5O7lVmcoaa8_bFKI8dBY8Q';
const GOOGLE_CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv`;
const STORAGE_KEY = 'unifecaf-erp-state-v2';
const USERS_STORAGE_KEY = 'unifecaf-erp-users-v1';

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

const STATUS_OPTIONS = ['MATRICULADO', 'ATIVO', 'PREMAT', 'TRANCADO', 'ABANDONADO', 'CANCELADO', 'DESISTENTE'];
const LEAD_STAGES = ['Lead', 'Visita', 'Em negociação', 'Gerar contrato', 'Matriculado', 'Perdido'];
const MONTHS = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
const MONTHS_FULL = ['Janeiro', 'Fevereiro', 'Marco', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
const COURSE_DISCOUNT_RATES = [10, 20, 30, 40, 50, 60];
const CLASSROOM_OPTIONS = ['Laboratório Saúde', 'Sala de Saúde', 'Laboratório de Informática', 'Sala de Aula'];
const ACOMPANHAMENTO_NEW_ENROLLMENT_START = new Date(2026, 3, 1);
const ACOMPANHAMENTO_MATRICULA_CUTOFF_LABEL = '31/03/2026';
const ACOMPANHAMENTO_NEW_ENROLLMENT_LABEL = '01/04/2026';

const state = {
  profile: 'admin',
  currentUser: null,
  module: 'inteligencia',
  moduleView: '',
  dashboardFocus: '',
  segment: 'todos',
  allRows: [],
  filteredRows: [],
  selectedKey: '',
  store: loadStore(),
  users: [],
  remoteState: false,
  remoteSaveTimer: 0,
  remoteWarningShown: false,
  pendingImport: null,
  lastStudentCsvText: '',
  lastStudentCsvLabel: '',
  lastImportUndo: null,
};

let whatsappWindowRef = null;

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
  changePasswordButton: document.querySelector('#changePasswordButton'),
  logoutButton: document.querySelector('#logoutButton'),
  megaMenu: document.querySelector('#megaMenu'),
  megaMenuButton: document.querySelector('#megaMenuButton'),
  megaMenuClose: document.querySelector('#megaMenuClose'),
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
  modalOverlay: document.querySelector('#modalOverlay'),
  loadingOverlay: document.querySelector('#loadingOverlay'),
  modules: {
    inteligencia: document.querySelector('#module-inteligencia'),
    bi: document.querySelector('#module-bi'),
    acompanhamento: document.querySelector('#module-acompanhamento'),
    retencao: document.querySelector('#module-retencao'),
    atendimento: document.querySelector('#module-atendimento'),
    financeiro: document.querySelector('#module-financeiro'),
    repasse: document.querySelector('#module-repasse'),
    cursos: document.querySelector('#module-cursos'),
    matriculas: document.querySelector('#module-matriculas'),
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

  els.logoutButton.addEventListener('click', () => endSession());

  els.megaMenuButton?.addEventListener('click', openMegaMenu);
  els.megaMenuClose?.addEventListener('click', closeMegaMenu);

  els.profileSelect.addEventListener('change', () => {
    if (state.currentUser && normalize(state.currentUser.perfil) !== 'admin') {
      state.profile = state.currentUser.perfil || 'consultor';
      updateUserStatus();
      toast('Somente o Administrador pode alternar a visão de perfil.');
      return;
    }

    state.profile = els.profileSelect.value;
    updateUserStatus();
    render();
  });

  document.querySelectorAll('.nav-link').forEach((button) => {
    button.addEventListener('click', (event) => {
      const module = button.dataset.module || button.dataset.quickModule;
      if (!module) return;
      if (button.dataset.quickModule) event.stopPropagation();
      if (button.dataset.view) state.moduleView = button.dataset.view;
      if (isRestrictedModule(module)) {
        toast(restrictedMessage(module));
        return;
      }
      activateModule(module);
    });
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
  els.changePasswordButton?.addEventListener('click', showPasswordModal);
  els.closeDrawer.addEventListener('click', () => els.drawer.classList.remove('open'));

  els.csvInput.addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    startGuidedImport('students', await file.text(), file.name);
    els.csvInput.value = '';
  });

  document.addEventListener('submit', handleSubmit);
  document.addEventListener('click', handleClick);
  document.addEventListener('keydown', handleKeydown);
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
  const sessionValid = await validateCurrentSession();
  if (!sessionValid) {
    endSession('Usuário removido ou desativado. Faça login com uma conta ativa.');
    return;
  }
  updateUserStatus();
  activateModule(state.module, false);
  setLoading(true, 'Carregando dados do polo...');
  try {
    await loadOperationalStore(false);
    await loadRemoteSheet(false);
  } finally {
    setLoading(false);
  }
}

async function validateCurrentSession() {
  if (!state.currentUser?.usuario) return true;
  const users = await loadUsers();
  const activeUser = users.find(
    (user) => normalize(user.usuario) === normalize(state.currentUser.usuario) && normalize(user.ativo || 'SIM') !== 'nao',
  );
  if (!activeUser) return false;
  state.currentUser = {
    usuario: activeUser.usuario,
    nome: activeUser.nome || activeUser.usuario,
    perfil: normalizeProfile(activeUser.perfil || 'consultor'),
  };
  state.profile = state.currentUser.perfil;
  localStorage.setItem('unifecaf-erp-user', JSON.stringify(state.currentUser));
  return true;
}

function endSession(message = '') {
  localStorage.removeItem('unifecaf-erp-session');
  localStorage.removeItem('unifecaf-erp-user');
  state.currentUser = null;
  els.appView.hidden = true;
  els.loginView.hidden = false;
  if (message) toast(message);
}

function activateModule(module, shouldRender = true) {
  if (isRestrictedModule(module)) {
    state.module = 'inteligencia';
    state.moduleView = '';
    toast(restrictedMessage(module));
  } else {
    state.module = module || 'inteligencia';
    if (state.module !== 'matriculas') state.moduleView = '';
    if (state.module === 'matriculas' && !state.moduleView) state.moduleView = 'crm';
  }
  document.querySelectorAll('.nav-link').forEach((item) => {
    const moduleName = item.dataset.module || item.dataset.quickModule;
    const sameModule = moduleName === state.module;
    const sameView = !item.dataset.view || item.dataset.view === state.moduleView;
    item.classList.toggle('active', sameModule && sameView);
  });
  Object.entries(els.modules).forEach(([name, element]) => {
    element.classList.toggle('active', name === state.module);
  });
  els.appView.classList.toggle('home-mode', state.module === 'inteligencia');
  if (shouldRender) render();
}

function applyAccessControl() {
  document.querySelectorAll('[data-module="financeiro"], [data-quick-module="financeiro"], [data-admin-module="financeiro"]').forEach((button) => {
    const locked = !canSeeFinancial();
    button.hidden = locked;
    button.classList.toggle('restricted', locked);
    button.title = locked ? 'Acesso restrito ao Administrador ou Financeiro' : 'Acessar módulo';
  });
  document.querySelectorAll('[data-module="seguranca"], [data-admin-module="seguranca"], [data-admin-module="cursos"]').forEach((button) => {
    const locked = !canAccessAdminModules();
    button.hidden = locked;
    button.classList.toggle('restricted', locked);
    button.title = locked ? 'Acesso restrito ao Administrador' : 'Acessar módulo';
  });
  document.querySelectorAll('[data-module="repasse"], [data-admin-module="repasse"]').forEach((button) => {
    const locked = !canSeeFinancial();
    button.hidden = locked;
    button.classList.toggle('restricted', locked);
    button.title = locked ? 'Acesso restrito ao Administrador ou Financeiro' : 'Acessar módulo';
  });
  document.querySelectorAll('[data-admin-menu]').forEach((menu) => {
    menu.hidden = !canAccessAdminModules();
  });
  document.querySelectorAll('[data-financial-only]').forEach((element) => {
    element.hidden = !canSeeFinancial();
  });
  document.querySelectorAll('[data-admin-only]').forEach((element) => {
    element.hidden = !canAccessAdminModules();
  });
  if (isRestrictedModule(state.module)) activateModule('inteligencia', false);
}

function openMegaMenu() {
  if (!els.megaMenu) return;
  applyAccessControl();
  els.megaMenu.hidden = false;
  els.megaMenuButton?.setAttribute('aria-expanded', 'true');
  document.body.classList.add('mega-menu-open');
  requestAnimationFrame(() => {
    els.megaMenu.querySelector('button:not([hidden])')?.focus();
  });
}

function closeMegaMenu() {
  if (!els.megaMenu || els.megaMenu.hidden) return;
  els.megaMenu.hidden = true;
  els.megaMenuButton?.setAttribute('aria-expanded', 'false');
  document.body.classList.remove('mega-menu-open');
  els.megaMenuButton?.focus();
}

function updateUserStatus() {
  const actualProfile = normalizeProfile(state.currentUser?.perfil || state.profile || 'consultor');
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
    els.userStatus.textContent = `Bem-vindo, ${state.currentUser.nome} (${roleLabel(actualProfile)}${viewLabel})`;
    return;
  }

  els.userStatus.textContent = `Bem-vindo, ${roleLabel(state.profile)}`;
}

async function authenticateUser(username, password) {
  const users = await loadUsers();
  const normalizedUser = normalize(username);
  let user = users.find(
    (item) =>
      normalize(item.usuario) === normalizedUser &&
      String(item.senha || '') === String(password || '') &&
      normalize(item.ativo || 'SIM') !== 'nao',
  );

  if (!user) return false;

  user.lastAccess = new Date().toISOString();
  state.users = users.map((item) => (normalize(item.usuario) === normalizedUser ? normalizeUser(user) : normalizeUser(item)));
  persistUsersLocal();

  state.currentUser = {
    usuario: user.usuario,
    nome: user.nome || user.usuario,
    perfil: normalizeProfile(user.perfil || 'consultor'),
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
    state.users = loadLocalUsers();
  }
  return state.users;
}

function loadLocalUsers() {
  try {
    const users = JSON.parse(localStorage.getItem(USERS_STORAGE_KEY) || '[]');
    return Array.isArray(users) && users.length ? mergeUsersByLogin(users.map(normalizeUser)) : defaultUsers();
  } catch {
    return defaultUsers();
  }
}

function mergeUsersByLogin(users = []) {
  const map = new Map();
  users.map(normalizeUser).forEach((user) => {
    if (user.usuario) map.set(normalize(user.usuario), user);
  });
  return [...map.values()];
}

function normalizeUser(user = {}) {
  const profile = normalizeProfile(user.perfil);
  return {
    usuario: cleanText(user.usuario),
    senha: cleanText(user.senha),
    nome: cleanText(user.nome) || cleanText(user.usuario),
    perfil: profile,
    ativo: normalize(user.ativo || 'SIM') === 'nao' ? 'NAO' : 'SIM',
    lastAccess: cleanText(user.lastAccess || user.ultimoAcesso),
  };
}

function normalizeProfile(value) {
  const profile = normalize(value);
  if (profile.includes('admin') || profile.includes('administrador')) return 'admin';
  if (profile.includes('financ')) return 'financeiro';
  return 'consultor';
}

function persistUsersLocal() {
  localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(state.users.map(normalizeUser)));
}

function mergeLocalUserHistory(users) {
  const localUsers = loadLocalUsers();
  return users.map((user) => {
    const normalized = normalizeUser(user);
    const local = localUsers.find((item) => normalize(item.usuario) === normalize(normalized.usuario));
    return {
      ...normalized,
      lastAccess: mostRecentIso(normalized.lastAccess, local?.lastAccess),
    };
  });
}

function syncLoggedUserAccess() {
  const username = normalize(state.currentUser?.usuario);
  if (!username) return;
  const local = loadLocalUsers().find((user) => normalize(user.usuario) === username);
  const index = state.users.findIndex((user) => normalize(user.usuario) === username);
  if (index < 0 || !local?.lastAccess) return;
  state.users[index] = { ...normalizeUser(state.users[index]), lastAccess: mostRecentIso(state.users[index].lastAccess, local.lastAccess) };
  persistUsersLocal();
  saveUsersRemote();
}

function changeCurrentPassword(data) {
  const username = normalize(state.currentUser?.usuario);
  const currentPassword = cleanText(data.currentPassword);
  const newPassword = cleanText(data.newPassword);
  const confirmPassword = cleanText(data.confirmPassword);
  const users = (state.users.length ? state.users : loadLocalUsers()).map(normalizeUser);
  const index = users.findIndex((user) => normalize(user.usuario) === username);

  if (!username || index < 0) {
    toast('Faça login novamente para alterar a senha.');
    return;
  }
  if (users[index].senha !== currentPassword) {
    toast('Senha atual incorreta.');
    return;
  }
  if (newPassword.length < 4) {
    toast('A nova senha precisa ter pelo menos 4 caracteres.');
    return;
  }
  if (newPassword !== confirmPassword) {
    toast('A confirmação da nova senha não confere.');
    return;
  }

  users[index] = { ...users[index], senha: newPassword };
  state.users = mergeUsersByLogin(users);
  persistUsersLocal();
  saveUsersRemote();
  recordAudit('Senha alterada', state.currentUser?.usuario || 'usuario logado');
  closeModal();
  toast('Senha alterada com sucesso.');
}

function mostRecentIso(first, second) {
  const firstTime = Date.parse(first || '');
  const secondTime = Date.parse(second || '');
  if (!Number.isNaN(firstTime) && (Number.isNaN(secondTime) || firstTime >= secondTime)) return first;
  return Number.isNaN(secondTime) ? first || '' : second;
}

async function saveUsersRemote() {
  if (!state.remoteState) return;
  try {
    const response = await fetch('/.netlify/functions/state', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'users', actor: roleLabel(state.profile), users: state.users.map(normalizeUser) }),
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    if (!payload.ok) throw new Error(payload.message || 'Falha ao salvar usuarios');
    els.syncStamp.textContent = `Usuarios salvos na planilha: ${new Date().toLocaleString('pt-BR')}`;
  } catch (error) {
    toast(`Usuarios salvos localmente. Falha na planilha: ${error.message}`);
  }
}

function saveConfiguredUser(data) {
  if (!canAccessAdminModules()) {
    toast('Gestão de usuários disponível apenas para Administrador.');
    return;
  }
  const user = normalizeUser(data);
  if (!user.usuario || !user.senha) {
    toast('Informe usuario e senha.');
    return;
  }
  const index = state.users.findIndex((item) => normalize(item.usuario) === normalize(user.usuario));
  if (index >= 0) {
    state.users[index] = user;
  } else {
    state.users.push(user);
  }
  persistUsersLocal();
  recordAudit('Usuario salvo', `${user.usuario} - ${roleLabel(user.perfil)}`);
  saveUsersRemote();
  toast('Usuario salvo.');
}

function deleteConfiguredUser(username) {
  if (!canAccessAdminModules()) {
    toast('Gestão de usuários disponível apenas para Administrador.');
    return;
  }
  const normalized = normalize(username);
  if (normalize(state.currentUser?.usuario) === normalized) {
    toast('O usuario logado nao pode excluir a propria conta.');
    return;
  }
  const before = state.users.length;
  state.users = state.users.filter((user) => normalize(user.usuario) !== normalized);
  if (state.users.length === before) return;
  persistUsersLocal();
  recordAudit('Usuario excluido', username);
  saveUsersRemote();
  persist();
  render();
  toast('Usuario excluido.');
}

async function loadOperationalStore(showLoader = true) {
  if (showLoader) setLoading(true, 'Carregando estado operacional...');
  try {
    const response = await fetch('/.netlify/functions/state', { cache: 'no-store' });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    if (!payload.ok || !payload.state) throw new Error(payload.message || 'Estado remoto indisponivel');

    state.store = normalizeStore(payload.state);
    state.users = Array.isArray(payload.users) ? mergeLocalUserHistory(payload.users) : state.users;
    seedStore();
    persistLocal();
    state.remoteState = true;
    syncLoggedUserAccess();
    toast('Estado operacional carregado da planilha ERP.');
  } catch (error) {
    state.remoteState = false;
    state.store = normalizeStore(state.store);
    seedStore();
    els.syncStamp.textContent = `Estado local ativo: ${error.message}`;
  } finally {
    if (showLoader) setLoading(false);
  }
}

async function loadRemoteSheet(showLoader = true) {
  if (showLoader) setLoading(true, 'Sincronizando planilha da sede...');
  setSourceStatus('Sincronizando planilha');
  const sources = ['/.netlify/functions/sheet', GOOGLE_CSV_URL];
  let lastError = null;

  for (const source of sources) {
    try {
      const response = await fetch(source, { cache: 'no-store' });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      hydrateRows(await response.text(), source.includes('netlify') ? 'Planilha via Netlify' : 'Planilha Google');
      if (showLoader) setLoading(false);
      return;
    } catch (error) {
      lastError = error;
    }
  }

  setSourceStatus('Importe um CSV para continuar');
  toast(`Não foi possível carregar a planilha: ${lastError?.message ?? 'erro desconhecido'}`);
  render();
  if (showLoader) setLoading(false);
}

function hydrateRows(csvText, label) {
  const parsed = parseCsv(csvText);
  state.allRows = parsed.rows.map((row, index) => enrichRow(row, index));
  state.lastStudentCsvText = csvText;
  state.lastStudentCsvLabel = label;

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
  const sourceStatus = cleanText(row[fields.status]) || 'SEM STATUS';
  const sourceStartPeriod = cleanText(row[fields.startPeriod]) || cleanText(row[fields.currentPeriod]) || '-';
  const sourceEnrollmentDate = parseBrazilianDate(cleanText(row[fields.enrollmentDate]));
  const sourceStartDate = cohortDate(sourceStartPeriod);
  const importedAfterNewRule = isAcompanhamentoNewEnrollmentByDates(sourceEnrollmentDate, sourceStartDate);
  const localStatus = override.status || (importedAfterNewRule ? 'MATRICULADO' : sourceStatus);
  const debtValue = parseMoney(row[fields.debtValue]);
  const enriched = {
    raw: row,
    sourceType: 'acompanhamento',
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
    startPeriod: sourceStartPeriod,
    currentClass: cleanText(row[fields.currentClass]),
    sourceStatus,
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
  syncDecisionFoundation();
  renderAlerts();
  renderInteligencia();
  renderBI();
  renderAcompanhamento();
  renderRetencao();
  renderAtendimento();
  renderFinanceiro();
  renderRepasse();
  renderCursos();
  renderMatriculas();
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
  const boletoPending = enrollmentActionItems(state.filteredRows.length ? state.filteredRows : state.allRows).slice(0, 2);
  const alerts = [
    ...boletoPending.map((item) => ({
      type: 'warning',
      text: item.boletoOk
        ? `Baixa a confirmar: ${item.name} precisa da confirmação de pagamento ou isenção.`
        : `Boleto a confirmar: ${item.name} foi matriculado e precisa da confirmação de envio.`,
    })),
    ...lateFollowups.map((row) => ({
      type: 'danger',
      text: `Prioridade: ${row.name} está com atenção alta e status ${row.localStatus}.`,
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
  const noAvaRows = rows.filter((row) => ['red', 'yellow'].includes(row.avaAlert.level));
  const boletoPending = enrollmentActionItems(rows);
  const retention = rows.length ? Math.round((activeRows.length / rows.length) * 100) : 0;
  const debtTotal = sumBy(debtRows, (row) => row.debtValue);
  const queue = dailyQueue();
  const target = Number(state.store.settings.monthlyTarget || 65);
  const matriculas = activeRows.length + confirmedNewMatriculations();
  const gap = matriculas - target;
  const totalComputers = Number(state.store.settings.computersTotal || 24);
  const maintenance = Number(state.store.settings.computersMaintenance || 0);
  const availableComputers = Math.max(0, totalComputers - maintenance - currentExamReservations());
  const signals = decisionSignals({ rows, debtRows, debtTotal, retention, gap, availableComputers, queue });
  const matriculatedLeads = confirmedNewMatriculations();
  const focus = dashboardFocusData(rows, { debtRows, noAvaRows, boletoPending, activeRows });
  const autoTasks = automaticTaskItems(rows);
  const openTasks = autoTasks.filter((task) => task.status !== 'Concluída');
  const quality = dataQualityIssues(rows);
  const qualityScore = dataQualityScore(quality, rows.length);

  els.modules.inteligencia.innerHTML = `
    ${portalHome({
      rows,
      activeRows,
      totals,
      debtRows,
      debtTotal,
      boletoPending,
      noAvaRows,
      target,
      gap,
      availableComputers,
      matriculatedLeads,
      retention,
    })}
    <section class="metric-grid">
      ${dashboardMetricCard('Alunos ativos', `${retention}%`, `${activeRows.length.toLocaleString('pt-BR')} ativos`, retention >= 70 ? 'green' : 'yellow', 'ativos')}
      ${dashboardMetricCard('Precisam de atenção', totals.highRisk, 'Ação acadêmica ou financeira', totals.highRisk ? 'yellow' : 'green', 'risco')}
      ${dashboardMetricCard('Sem acesso ao AVA', noAvaRows.length, 'Alerta amarelo ou vermelho', noAvaRows.length ? 'red' : 'green', 'ava')}
      ${dashboardMetricCard('Matrículas a regularizar', boletoPending.length, 'Boleto, baixa ou isenção', boletoPending.length ? 'yellow' : 'green', 'boletos')}
      ${dashboardMetricCard('Pagamentos em atraso', formatMoney(debtTotal), `${debtRows.length} alunos`, debtRows.length ? 'red' : 'green', 'inadimplencia')}
      ${dashboardMetricCard('Resultado da meta', gap, `Meta mensal: ${target}`, gap >= 0 ? 'green' : 'red', 'metas')}
      ${metricCard('Recursos livres', availableComputers, 'Computadores úteis agora', availableComputers > 4 ? 'cyan' : 'red')}
      ${metricCard('Tarefas abertas', openTasks.length, 'Geradas pelos alertas', openTasks.length ? 'yellow' : 'green')}
      ${metricCard('Qualidade dos dados', `${qualityScore}%`, `${quality.reduce((total, item) => total + item.count, 0)} pontos`, qualityScore >= 85 ? 'green' : 'yellow')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Prioridades de hoje', 'O que fazer agora')}</div>
          <span>${nextBestActions(rows).length} prioridades</span>
        </div>
        <div class="decision-list">
          ${nextBestActions(rows)
            .map(
              (action) => `
                <button class="decision-signal ${action.tone} action-signal" type="button" data-dashboard-focus="${escapeHtml(action.focus)}">
                  <strong>${escapeHtml(action.title)}</strong>
                  <span>${escapeHtml(action.text)}</span>
                </button>
              `,
            )
            .join('')}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Trabalho da equipe', 'Vendas e acompanhamento')}</div>
          <span>Operação do polo</span>
        </div>
        <div class="productivity-grid">
          ${productivityItems()
            .map(
              (item) => `
                <div class="productivity-item">
                  <strong>${escapeHtml(item.name)}</strong>
                  <span>${escapeHtml(item.detail)}</span>
                  <em>${escapeHtml(item.value)}</em>
                </div>
              `,
            )
            .join('')}
        </div>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Central de execução', 'Alertas que viraram tarefa, dono e prazo')}</div>
        <div class="panel-actions">
          <button class="mini-button" type="button" data-export="tarefas">Exportar tarefas</button>
          <button class="mini-button" type="button" data-report="executivo">Relatório PDF</button>
          <button class="mini-button" type="button" data-quick-module="bi">Ver indicadores</button>
          <span>${openTasks.length.toLocaleString('pt-BR')} abertas</span>
        </div>
      </div>
      <div class="table-wrap">
        ${taskTable(autoTasks.slice(0, 80))}
      </div>
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Saúde da base', 'Confiabilidade para decidir')}</div>
          <span>${qualityScore}%</span>
        </div>
        <div class="decision-list">
          ${quality.slice(0, 5).map(qualitySignalTemplate).join('') || '<div class="decision-signal success"><strong>Base sem alerta crítico</strong><span>Não há inconsistências relevantes nesta visão.</span></div>'}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Histórico mensal', 'Último retrato congelado')}</div>
          <div class="panel-actions">
            <button class="mini-button" type="button" data-undo-import>Desfazer importação</button>
            <span>${state.store.snapshots.length} snapshots</span>
          </div>
        </div>
        ${snapshotSummaryCard(latestSnapshot())}
      </article>
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Previsões simples', 'Tendência para agir antes do fechamento')}</div>
          <span>Projeção</span>
        </div>
        <div class="decision-list">
          ${forecastSignals(rows).map((item) => `<div class="decision-signal ${item.tone}"><strong>${escapeHtml(item.title)}</strong><span>${escapeHtml(item.text)}</span></div>`).join('')}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Relatórios executivos', 'Reunião de gestão em um clique')}</div>
          <span>PDF e Excel</span>
        </div>
        <div class="decision-list">
          <button class="decision-signal info action-signal" type="button" data-report="executivo"><strong>Relatório executivo</strong><span>Abre uma versão pronta para imprimir ou salvar como PDF.</span></button>
          <button class="decision-signal success action-signal" type="button" data-export="executivo"><strong>Resumo para Excel</strong><span>Exporta KPIs, qualidade, tarefas e previsão em CSV.</span></button>
          <button class="decision-signal warning action-signal" type="button" data-export="qualidade"><strong>Qualidade dos dados</strong><span>Lista os problemas que afetam a tomada de decisão.</span></button>
        </div>
      </article>
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Pontos de atenção', 'Avisos importantes')}</div>
          <span>${signals.length} avisos</span>
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
          <div>${smallTitle('Plano de ação', 'Acompanhamentos salvos')}</div>
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
    <section class="table-panel student-access-panel">
      <div class="panel-heading">
        <div>${smallTitle(focus.title, focus.subtitle)}</div>
        <div class="panel-actions">
          ${state.dashboardFocus ? '<button class="mini-button" type="button" data-dashboard-focus="">Limpar filtro</button>' : ''}
          <button class="mini-button" type="button" data-quick-module="acompanhamento">Ver lista completa</button>
          <span>${focus.rows.length.toLocaleString('pt-BR')} registros</span>
        </div>
      </div>
      <div class="table-wrap">
        ${dashboardFocusTable(focus)}
      </div>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Histórico das ações', 'Registro do que foi combinado')}</div>
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

function portalHome({
  rows,
  activeRows,
  totals,
  debtRows,
  debtTotal,
  boletoPending,
  noAvaRows,
  target,
  gap,
  availableComputers,
  matriculatedLeads,
  retention,
}) {
  const userName = state.currentUser?.nome || roleLabel(state.profile);
  const groups = [
    {
      title: 'Acadêmico',
      icon: '▦',
      links: [
        { label: 'Lista de alunos', detail: `${rows.length.toLocaleString('pt-BR')} registros`, attrs: 'data-quick-module="acompanhamento"' },
        { label: 'Acompanhamento do AVA', detail: `${noAvaRows.length.toLocaleString('pt-BR')} em alerta`, attrs: 'data-quick-module="retencao"' },
        { label: 'Atendimento ao aluno', detail: `${state.store.serviceTickets.length.toLocaleString('pt-BR')} solicitações`, attrs: 'data-quick-module="atendimento"' },
        { label: 'Alunos ativos', detail: `${activeRows.length.toLocaleString('pt-BR')} ativos`, attrs: 'data-dashboard-focus="ativos"' },
        { label: 'Precisam de atenção', detail: `${totals.highRisk.toLocaleString('pt-BR')} alunos`, attrs: 'data-dashboard-focus="risco"' },
      ],
    },
    {
      title: 'Planejamento',
      icon: '▥',
      links: [
        { label: 'Indicadores do polo', detail: `${retention}% de permanência`, attrs: 'data-quick-module="bi"' },
        { label: 'Metas do polo', detail: `Meta mensal: ${target}`, attrs: 'data-quick-module="metas"' },
        { label: 'Resultado da meta', detail: gap >= 0 ? `${gap} acima` : `${Math.abs(gap)} faltando`, attrs: 'data-dashboard-focus="metas"' },
        { label: 'Plano de ação', detail: `${state.store.decisions.length} ações`, attrs: 'data-dashboard-focus="base"' },
      ],
    },
    {
      title: 'Captação',
      icon: '▤',
      links: [
        { label: 'Novo candidato', detail: `${state.store.leads.length} candidatos`, attrs: 'data-quick-module="matriculas" data-view="crm"' },
        { label: 'Funil de captação', detail: `${matriculatedLeads} confirmados`, attrs: 'data-quick-module="matriculas" data-view="crm"' },
        { label: 'Fechamento', detail: `${boletoPending.length} pendências`, attrs: 'data-quick-module="matriculas" data-view="matricula"' },
        { label: 'Matrículas a regularizar', detail: `${boletoPending.length} pendentes`, attrs: 'data-dashboard-focus="boletos"' },
      ],
    },
    {
      title: 'Agenda',
      icon: '▧',
      links: [
        { label: 'Aulas e provas', detail: `${state.store.schedule.length + state.store.exams.length} reservas`, attrs: 'data-quick-module="agenda"' },
        { label: 'Fila do dia', detail: `${dailyQueue().length} eventos hoje`, attrs: 'data-quick-module="fila"' },
        { label: 'Infraestrutura', detail: `${availableComputers} computadores livres`, attrs: 'data-quick-module="avaliacoes"' },
        { label: 'Agendar aula', detail: 'Sala, horário e docente', attrs: 'data-quick-module="agenda"' },
      ],
    },
  ];

  if (canAccessAdminModules()) {
    groups.push(
      {
        title: 'Financeiro',
        icon: '▨',
        links: [
          { label: 'Pagamentos', detail: `${debtRows.length.toLocaleString('pt-BR')} em atraso`, attrs: 'data-quick-module="financeiro"' },
          { label: 'Total em atraso', detail: formatMoney(debtTotal), attrs: 'data-dashboard-focus="inadimplencia"' },
          { label: 'Valores repassados', detail: `${state.store.repasses.length} linhas`, attrs: 'data-admin-module="repasse"' },
          { label: 'Importar planilhas', detail: 'Faturamento e recebimento', attrs: 'data-admin-module="financeiro"' },
        ],
      },
      {
        title: 'Configurações',
        icon: '▩',
        links: [
          { label: 'Cursos', detail: 'Catálogo do polo', attrs: 'data-admin-module="cursos"' },
          { label: 'Usuários e acessos', detail: `${state.users.length || defaultUsers().length} usuários`, attrs: 'data-admin-module="seguranca"' },
          { label: 'Histórico de alterações', detail: `${state.store.auditTrail.length} registros`, attrs: 'data-admin-module="seguranca"' },
          { label: 'Dados do polo', detail: `${Object.keys(state.store.overrides).length} ajustes`, attrs: 'data-dashboard-focus="base"' },
        ],
      },
    );
  }

  return `
    <section class="portal-home" aria-label="Portal do polo">
      <header class="portal-home-header">
        <div>
          <p class="portal-kicker">Portal do Polo</p>
          <h1>MENDES &amp; FLOR EDUCACIONAL</h1>
          <span>Polo UniFECAF - bem-vindo, ${escapeHtml(userName)}</span>
        </div>
        <div class="portal-home-summary">
          <strong>${activeRows.length.toLocaleString('pt-BR')}</strong>
          <span>alunos ativos</span>
        </div>
      </header>
      <div class="portal-module-grid">
        ${groups.map(portalGroup).join('')}
      </div>
      <footer class="portal-home-footer">
        <span>Dados da sede + dados do polo</span>
        <span>${new Date().getFullYear()} - Gestão educacional</span>
      </footer>
    </section>
  `;
}

function portalGroup(group) {
  return `
    <article class="portal-group">
      <div class="portal-group-title">
        <span aria-hidden="true">${escapeHtml(group.icon)}</span>
        <strong>${escapeHtml(group.title)}</strong>
      </div>
      <div class="portal-link-list">
        ${group.links
          .map(
            (link) => `
              <button class="portal-link" type="button" ${link.attrs}>
                <span>${escapeHtml(link.label)}</span>
                <small>${escapeHtml(link.detail)}</small>
              </button>
            `,
          )
          .join('')}
      </div>
    </article>
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
  const decisionSeries = receiptRepasseSeries(filter);
  const indicatorRows = indicatorDictionaryRows({ rows: censusRows, census, commercial, finance, monthlyProgress, annualProgress });
  const quality = dataQualityIssues(censusRows);
  const qualityScore = dataQualityScore(quality, censusRows.length);
  const snapshots = state.store.snapshots.slice().sort((a, b) => collator.compare(b.periodKey, a.periodKey));

  els.modules.bi.innerHTML = `
    ${moduleTitle('Indicadores', 'Números principais para decidir com segurança.')}
    <section class="bi-filter-bar">
      <form class="inline-form bi-controls" data-form="bi-filter">
        <select name="biMonth">${monthOptions(filter.month)}</select>
        <select name="biYear">${yearOptions(filter.year)}</select>
        <button type="submit">Atualizar indicadores</button>
      </form>
      <span>${escapeHtml(periodLabel(filter))}</span>
    </section>
    <section class="bi-decision-grid">
      ${gaugeWidget('Meta do mês', monthlyProgress, `${commercial.monthlyMatriculations}/${monthlyTarget} matrículas`)}
      ${donutWidget('Situação dos alunos', census.active, census.inactive)}
      ${lineWidget('Recebido x Repassado', decisionSeries)}
    </section>
    <section class="metric-grid">
      ${metricCard('Meta mensal', `${monthlyProgress}%`, `${commercial.monthlyMatriculations}/${monthlyTarget} matrículas`, monthlyProgress >= 100 ? 'green' : 'yellow')}
      ${metricCard('Meta anual', `${annualProgress}%`, `${commercial.annualMatriculations}/${annualTarget} matrículas`, annualProgress >= 100 ? 'green' : 'cyan')}
      ${metricCard('Alunos ativos', census.active, `${censusRows.length.toLocaleString('pt-BR')} registros analisados`, 'green')}
      ${metricCard('Alunos inativos', census.inactive, 'Trancado, abandonado e cancelado', census.inactive ? 'yellow' : 'green')}
      ${metricCard('Qualidade da base', `${qualityScore}%`, `${quality.length} tipos de alerta`, qualityScore >= 85 ? 'green' : 'yellow')}
      ${metricCard('Histórico salvo', snapshots.length, 'Snapshots mensais', snapshots.length ? 'cyan' : 'yellow')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Funil de captação', 'Metas mensal e anual')}</div>
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
          <div>${smallTitle('Situação da base', 'Status dos alunos por período')}</div>
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
          <div>${smallTitle('Resumo por status', 'Ativo, abandono, trancado e cancelado')}</div>
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
          <div>${smallTitle('Pagamentos', 'Taxa de matrícula e mensalidades recorrentes')}</div>
          <span>${canSeeFinancial() ? 'Comparativo do período' : 'Disponível apenas para Administrador'}</span>
        </div>
        ${
          canSeeFinancial()
            ? `
              <form class="inline-form bi-controls" data-form="bi-ticket">
                <input type="number" min="0" step="0.01" name="monthlyTicket" value="${Number(state.store.settings.monthlyTicket || 299)}" placeholder="Valor médio da mensalidade" />
                <button type="submit">Atualizar valor</button>
              </form>
              <div class="finance-kpi-list">
                ${financeKpiItem('Taxa de matrícula', finance.enrollment.current, finance.enrollment.previousMonth, finance.enrollment.previousYear)}
                ${financeKpiItem('Mensalidades recebidas', finance.recurring.current, finance.recurring.previousMonth, finance.recurring.previousYear)}
                ${financeKpiItem('Acumulado do Ano', finance.ytd.current, finance.ytd.previousMonth, finance.ytd.previousYear)}
              </div>
            `
            : '<div class="locked-panel"><strong>Acesso protegido</strong><p>Fluxo de caixa disponível somente para Administrador e Financeiro.</p></div>'
        }
      </article>
    </section>
    <section class="table-panel">
        <div class="panel-heading">
          <div>${smallTitle('Dicionário de indicadores', 'Fórmula, fonte, dono e valor atual')}</div>
          <div class="panel-actions">
            <button class="mini-button" type="button" data-export="indicadores">Exportar Excel</button>
            <span>${indicatorRows.length} métricas oficiais</span>
          </div>
        </div>
      <div class="table-wrap">
        ${indicatorDictionaryTable(indicatorRows)}
      </div>
    </section>
    <section class="split-grid">
      <article class="table-panel">
        <div class="panel-heading">
          <div>${smallTitle('Histórico mensal congelado', 'Comparação da evolução do polo')}</div>
          <div class="panel-actions">
            <button class="mini-button" type="button" data-refresh-snapshot>Atualizar snapshot do mês</button>
            <button class="mini-button" type="button" data-export="snapshots">Exportar histórico</button>
            <button class="mini-button" type="button" data-undo-import>Desfazer importação</button>
            <span>${snapshots.length} períodos</span>
          </div>
        </div>
        <div class="table-wrap">
          ${snapshotTable(snapshots)}
        </div>
      </article>
      <article class="table-panel">
        <div class="panel-heading">
          <div>${smallTitle('Qualidade dos dados', 'Erros que afetam a decisão')}</div>
          <div class="panel-actions">
            <button class="mini-button" type="button" data-export="qualidade">Exportar Excel</button>
            <span>${qualityScore}% de confiança</span>
          </div>
        </div>
        <div class="table-wrap">
          ${qualityTable(quality)}
        </div>
      </article>
    </section>
  `;
}

function renderAcompanhamento() {
  const rows = state.filteredRows;
  const totals = getStudentTotals(rows);
  const cohorts = getCohortStats(rows).slice(0, 7);
  els.modules.acompanhamento.innerHTML = `
    ${moduleTitle('Lista de alunos', 'Dados da sede com atualizações feitas pelo polo.')}
    <section class="metric-grid">
      ${metricCard('Alunos filtrados', rows.length, 'Registros na visão atual')}
      ${metricCard('Ativos', totals.active, 'Status operacional ativo', 'green')}
      ${metricCard('Dados do polo', totals.overrides, 'Atualizações locais preservadas', 'cyan')}
      ${metricCard('Atenção alta', totals.highRisk, 'Fila de contato', 'yellow')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Análise de período', 'Início das aulas')}</div>
          <span>${cohorts.length} períodos</span>
        </div>
        <div class="bars">
          ${cohorts
            .map(
              (item) => `
                <div class="bar-line">
                  <strong>${escapeHtml(item.period)}</strong>
                  <div class="bar-track"><span style="width:${item.retention}%"></span></div>
                  <em>${item.retention}% permanência</em>
                </div>
              `,
            )
            .join('')}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Ação rápida', 'Ajustes direto na linha')}</div>
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
        <div>${smallTitle('Lista principal', 'Nome, CPF, RA, curso e período de início')}</div>
        <div class="panel-actions">
          <button class="mini-button" type="button" data-export="ativos">Exportar ativos</button>
          <span>${rows.length.toLocaleString('pt-BR')} registros</span>
        </div>
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
              <th>Atenção</th>
              <th>Status do polo</th>
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
    .filter(isRetentionEligible)
    .sort((a, b) => b.avaAlert.weight - a.avaAlert.weight || b.avaDaysNumber - a.avaDaysNumber);
  const yellow = rows.filter((row) => row.avaAlert.level === 'yellow');
  const red = rows.filter((row) => row.avaAlert.level === 'red');
  const contacted = rows.filter((row) => row.retention.contacted).length;
  const pendingContact = rows.filter((row) => row.avaAlert.level !== 'ok' && !row.retention.contacted).length;

  els.modules.retencao.innerHTML = `
    ${moduleTitle('Acompanhamento AVA', 'Alunos ativos, pré-matriculados e transferências para contato prioritário.')}
    <section class="metric-grid">
      ${metricCard('Monitorados', rows.length, 'Ativos, pré-matriculados e transferências')}
      ${metricCard('Alerta amarelo', yellow.length, '5 a 7 dias sem acesso', 'yellow')}
      ${metricCard('Alerta vermelho', red.length, '8 dias ou mais sem acesso', 'red')}
      ${metricCard('Contatos registrados', contacted, `${pendingContact} pendentes`, contacted ? 'green' : 'cyan')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Regra de alerta', 'Acompanhamento de acesso')}</div>
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
            <span>Acende alerta vermelho e entra na prioridade de contato.</span>
          </div>
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Fila de contato', 'Motivo informado pelo aluno')}</div>
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
        <div>${smallTitle('Lista de acompanhamento', 'Acesso ao AVA e registro de contato')}</div>
        <div class="panel-actions">
          <button class="mini-button" type="button" data-export="retencao">Exportar lista</button>
          <span>${rows.length.toLocaleString('pt-BR')} alunos</span>
        </div>
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
            ${rows.slice(0, 260).map(retentionRowTemplate).join('') || emptyRow('Fila zerada. Ótimo momento para revisar contatos resolvidos e preparar ações preventivas.')}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderAtendimento() {
  const tickets = state.store.serviceTickets.slice().sort((a, b) => new Date(b.requestedAt || b.createdAt) - new Date(a.requestedAt || a.createdAt));
  const open = tickets.filter((ticket) => !['Finalizado', 'Respondido'].includes(ticket.status));
  const late = tickets.filter((ticket) => ticket.deadline && new Date(ticket.deadline) < startOfToday() && !['Finalizado', 'Respondido'].includes(ticket.status));
  const headquarters = tickets.filter((ticket) => normalize(ticket.status).includes('sede') || normalize(ticket.sector).includes('sede'));

  els.modules.atendimento.innerHTML = `
    ${moduleTitle('Atendimento ao aluno', 'Protocolos, solicitações à sede e respostas acompanhadas pelo polo.')}
    <section class="metric-grid">
      ${metricCard('Solicitações abertas', open.length, 'Acompanhar até a resposta', open.length ? 'yellow' : 'green')}
      ${metricCard('Com prazo vencido', late.length, 'Prioridade de cobrança', late.length ? 'red' : 'green')}
      ${metricCard('Enviadas à sede', headquarters.length, 'Dependem de retorno externo', 'cyan')}
      ${metricCard('Finalizadas', tickets.filter((ticket) => ['Finalizado', 'Respondido'].includes(ticket.status)).length, 'Histórico resolvido', 'green')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Nova solicitação', 'Problema, protocolo e responsável')}</div>
          <span>Gestão de atendimento</span>
        </div>
        <form class="stack-form" data-form="service-ticket">
          <input name="student" list="studentLookupOptions" placeholder="Buscar aluno por nome, CPF ou RA" required />
          ${studentLookupDatalist()}
          <div class="two-cols">
            <input name="protocol" placeholder="Protocolo da sede ou interno" />
            <select name="status">${serviceStatusOptions('Novo')}</select>
          </div>
          <div class="two-cols">
            <input name="requestedAt" type="date" value="${new Date().toISOString().slice(0, 10)}" required />
            <input name="deadline" type="date" placeholder="Prazo de resposta" />
          </div>
          <div class="two-cols">
            <input name="attendant" placeholder="Quem atendeu no polo" value="${escapeHtml(state.currentUser?.nome || '')}" />
            <input name="sector" placeholder="Setor responsável. Ex: Secretaria, Financeiro, Sede" />
          </div>
          <textarea name="problem" rows="3" placeholder="Descreva o problema do aluno" required></textarea>
          <textarea name="response" rows="3" placeholder="Resposta, orientação ou retorno recebido"></textarea>
          <button type="submit">Salvar atendimento</button>
        </form>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Fila de prioridade', 'Vencidos e pendentes primeiro')}</div>
          <span>${open.length} abertas</span>
        </div>
        <div class="decision-list">
          ${servicePriorityItems(tickets).join('') || '<div class="empty-state">Nenhuma solicitação aberta. Atendimento em dia.</div>'}
        </div>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Histórico de atendimentos', 'Aluno, problema, resposta e situação')}</div>
        <div class="panel-actions">
          <button class="mini-button" type="button" data-export="atendimentos">Exportar Excel</button>
          <span>${tickets.length.toLocaleString('pt-BR')} registros</span>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Aluno</th>
              <th>Protocolo</th>
              <th>Problema</th>
              <th>Solicitação</th>
              <th>Atendente/Setor</th>
              <th>Resposta</th>
              <th>Situação</th>
            </tr>
          </thead>
          <tbody>${tickets.map(serviceTicketRowTemplate).join('') || emptyRow('Nenhuma solicitação cadastrada.')}</tbody>
        </table>
      </div>
    </section>
  `;
}

function renderFinanceiro() {
  const locked = !canSeeFinancial();
  if (locked) {
    els.modules.financeiro.innerHTML = accessDenied('Financeiro', 'Módulo disponível apenas para Administrador ou Financeiro.');
    return;
  }

  const rows = state.filteredRows;
  const usingLocalFallback = !state.remoteState;
  const sourceFinance = buildFinancialPosition(rows);
  const hasFinancialSource = hasBillingReceiptSource();
  const debtRows = sourceFinance.debtRows;
  const paidRows = sourceFinance.paidRows;
  const debtTotal = sumBy(debtRows, (row) => row.debtValue);
  const defaultEnrollmentFee = Number(state.store.settings.enrollmentFee || 99);
  const leadEnrollments = state.store.leads.filter((lead) => lead.stage === 'Matriculado' && !isLeadShadowedByLegacyAcompanhamento(lead));
  const sourcePendingEnrollments = acompanhamentoPendingEnrollmentRows(rows, { excludeLeadMatches: true });
  const receivedEnrollment = sumBy(leadEnrollments, (lead) =>
    lead.enrollmentPaymentStatus === 'Pago' ? Number(lead.enrollmentFee || defaultEnrollmentFee) : 0,
  ) + sumBy(rows.filter(isAcompanhamentoNewEnrollment), (row) => {
    const enrollment = enrollmentSettlementForStudent(row);
    if (enrollment.lead || !enrollment.paymentOk || enrollment.paymentLabel === 'Isento') return 0;
    return Number((state.store.overrides[row.key] || {}).enrollmentFee || defaultEnrollmentFee);
  });
  const expectedEnrollment = sumBy(leadEnrollments.filter(leadNeedsEnrollmentAction), (lead) =>
    lead.enrollmentPaymentStatus === 'Isento' ? 0 : Number(lead.enrollmentFee || defaultEnrollmentFee),
  ) + sumBy(sourcePendingEnrollments, (row) => Number((state.store.overrides[row.key] || {}).enrollmentFee || defaultEnrollmentFee));
  const sourceSettledRows = rows.filter(isAcompanhamentoLegacyEnrollment).length;
  const pie = pieChart(paidRows.length, debtRows.length);
  const evolution = monthlyDebtEvolution();
  const filter = getBiFilter();
  const billedFiltered = totalFinancialRecords(state.store.billing, filter);
  const receivedFiltered = totalFinancialRecords(state.store.receipts, filter);
  const billedYear = totalFinancialRecords(state.store.billing, { month: 0, year: filter.year });
  const receivedYear = totalFinancialRecords(state.store.receipts, { month: 0, year: filter.year });
  const billedLifetime = totalFinancialRecords(state.store.billing, allHistoryFilter());
  const receivedLifetime = totalFinancialRecords(state.store.receipts, allHistoryFilter());
  const repasseLifetime = totalFinancialRecords(state.store.repasses, allHistoryFilter());
  const financeMissingData = !hasFinancialLedgerData();

  els.modules.financeiro.innerHTML = `
    ${moduleTitle('Financeiro e valores repassados', hasFinancialSource ? 'Mensalidades calculadas pelas planilhas de faturamento e recebimento.' : 'Importe faturamento e recebimento para ativar os indicadores financeiros.')}
    <section class="finance-subnav">
      <span>Dados lidos diretamente das abas Faturamento e Recebimento do banco.</span>
      <form class="inline-form bi-controls" data-form="bi-filter">
        <select name="biMonth">${monthOptions(filter.month)}</select>
        <select name="biYear">${yearOptions(filter.year)}</select>
        <button type="submit">Atualizar período</button>
      </form>
      <button class="mini-button" type="button" data-reset-bi-filter>Todo histórico</button>
      <button class="mini-button" type="button" data-refresh-operational>Atualizar banco</button>
      <button class="mini-button" type="button" data-admin-module="repasse">Ver valores repassados</button>
      <button class="mini-button" type="button" data-export="executivo">Exportar resumo</button>
    </section>
    ${financeMissingData ? `
      <section class="decision-signal warning">
        <strong>Financeiro sem base carregada</strong>
        <span>As abas Faturamento, Recebimento e Repasse ainda não entraram no estado operacional. Conecte o Apps Script ou importe os CSVs para visualizar os valores reais.</span>
      </section>
    ` : ''}
    ${usingLocalFallback ? `
      <section class="decision-signal warning">
        <strong>Fonte local em uso</strong>
        <span>O sistema ainda não está sincronizado com a sede. Os números abaixo usam o histórico salvo neste navegador até a conexão com o Apps Script ficar ativa.</span>
      </section>
    ` : ''}
    <section class="metric-grid">
      ${metricCard('Em dia', paidRows.length, 'Carteira sem atraso', 'green')}
      ${metricCard('Em atraso', debtRows.length, 'Alunos com pagamento em aberto', 'red')}
      ${metricCard('Total em atraso', formatMoney(debtTotal), 'Faturado menos recebido', 'yellow')}
      ${metricCard('Faturamento filtrado', formatMoney(billedFiltered), periodLabel(filter), 'green')}
      ${metricCard('Recebimento filtrado', formatMoney(receivedFiltered), periodLabel(filter), 'cyan')}
      ${metricCard('Faturado no ano', formatMoney(billedYear), yearlyLabel(filter), 'green')}
      ${metricCard('Recebido no ano', formatMoney(receivedYear), yearlyLabel(filter), 'cyan')}
      ${metricCard('Faturado histórico', formatMoney(billedLifetime), 'Desde a fundação do polo', 'green')}
      ${metricCard('Recebido histórico', formatMoney(receivedLifetime), 'Desde a fundação do polo', 'cyan')}
      ${metricCard('Repassado histórico', formatMoney(repasseLifetime), 'Repasse final da sede', 'yellow')}
      ${metricCard('Base sem boleto', sourceSettledRows, `Regularizada ate ${ACOMPANHAMENTO_MATRICULA_CUTOFF_LABEL}`, 'green')}
      ${metricCard('Novas a regularizar', sourcePendingEnrollments.length, `Desde ${ACOMPANHAMENTO_NEW_ENROLLMENT_LABEL}`, sourcePendingEnrollments.length ? 'yellow' : 'green')}
      ${metricCard('Taxas novas confirmadas', formatMoney(receivedEnrollment), `A confirmar: ${formatMoney(expectedEnrollment)}`, 'cyan')}
      ${metricCard('Faturamento', state.store.billing.length, 'Linhas históricas acumuladas', 'green')}
      ${metricCard('Recebimento', state.store.receipts.length, 'Linhas históricas acumuladas', 'cyan')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Faturamento da sede', 'Lido direto do banco')}</div>
          <span>${state.store.billing.length.toLocaleString('pt-BR')} linhas</span>
        </div>
        <p class="muted">Use as colunas <strong>Valor Faturado</strong> e <strong>DataFaturado</strong>. Novas linhas coladas na aba Faturamento entram no cálculo ao atualizar o banco.</p>
        <div class="finance-kpi-list">
          ${progressLine('Total filtrado', percent(billedFiltered, Math.max(billedLifetime, 1)), `${formatMoney(billedFiltered)} no período / ${formatMoney(billedLifetime)} desde a fundação`)}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Recebimentos da sede', 'Lido direto do banco')}</div>
          <span>${state.store.receipts.length.toLocaleString('pt-BR')} linhas</span>
        </div>
        <p class="muted">Use as colunas <strong>Valor Pago</strong> e <strong>Data Pagamento</strong>. O recebimento real é calculado pelo valor pago registrado pela sede.</p>
        <div class="finance-kpi-list">
          ${progressLine('Total filtrado', percent(receivedFiltered, Math.max(receivedLifetime, 1)), `${formatMoney(receivedFiltered)} no período / ${formatMoney(receivedLifetime)} desde a fundação`)}
        </div>
      </article>
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Situação dos pagamentos', 'Em dia e em atraso')}</div>
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
          <div>${smallTitle('Evolução dos atrasos', 'Faturado menos recebido por mês')}</div>
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
        <div>${smallTitle('Alunos com pagamento em atraso', 'Valor total e meses em aberto')}</div>
        <div class="panel-actions">
          <button class="mini-button" type="button" data-export="inadimplentes">Exportar Excel</button>
          <span>${debtRows.length.toLocaleString('pt-BR')} alunos</span>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Aluno</th>
              <th>Curso</th>
              <th>Status</th>
              <th>Valor em atraso</th>
              <th>Meses em aberto</th>
              <th>Ação</th>
            </tr>
          </thead>
          <tbody>
            ${debtRows.slice(0, 220).map(financialRowTemplate).join('') || emptyRow('Nenhum aluno com pagamento em atraso nesta visão.')}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderRepasse() {
  if (!canSeeFinancial()) {
    els.modules.repasse.innerHTML = accessDenied('Valores repassados', 'Acesso disponível apenas para Administrador ou Financeiro.');
    return;
  }

  const filter = getBiFilter();
  const usingLocalFallback = !state.remoteState;
  const rows = repasseRows(filter);
  const currentTotal = sumBy(rows, (row) => row.amount);
  const yearTotal = sumBy(repasseRows({ month: 0, year: filter.year }), (row) => row.amount);
  const lifetimeTotal = totalFinancialRecords(state.store.repasses, allHistoryFilter());
  const previousMonthTotal = Number(filter.year || 0) ? sumBy(repasseRows(previousMonthFilter(filter)), (row) => row.amount) : 0;
  const previousYearTotal = Number(filter.year || 0) ? sumBy(repasseRows(previousYearFilter(filter)), (row) => row.amount) : 0;
  const grouped = repasseEvolution();
  const reconciliation = repasseReconciliation(filter, currentTotal);
  const repasseMissingData = !state.store.repasses.length;
  const repasseControls = `
    <section class="finance-subnav">
      <span>Repasse final importado da sede. O sistema exibe o valor real e confere a faixa esperada sobre o recebimento anterior.</span>
      <form class="inline-form bi-controls" data-form="bi-filter">
        <select name="biMonth">${monthOptions(filter.month)}</select>
        <select name="biYear">${yearOptions(filter.year)}</select>
        <button type="submit">Atualizar repasse</button>
      </form>
      <button class="mini-button" type="button" data-reset-bi-filter>Todo histórico</button>
      <button class="mini-button" type="button" data-refresh-operational>Atualizar banco</button>
      <button class="mini-button" type="button" data-export="repasse">Exportar repasse</button>
    </section>
  `;

  els.modules.repasse.innerHTML = `
    ${moduleTitle('Valores repassados', 'Histórico dos valores enviados pela sede ao polo.')}
    ${repasseControls}
    ${repasseMissingData ? `
      <section class="decision-signal warning">
        <strong>Repasse sem base carregada</strong>
        <span>Não há linhas de Repasse no estado atual. Quando a aba da sede for conectada, o histórico e o acumulado vão aparecer aqui.</span>
      </section>
    ` : ''}
    ${usingLocalFallback ? `
      <section class="decision-signal warning">
        <strong>Fonte local em uso</strong>
        <span>O repasse ainda não está sendo lido do Apps Script. Os valores exibidos abaixo dependem do histórico salvo neste navegador.</span>
      </section>
    ` : ''}
    <section class="metric-grid">
      ${metricCard('Valor filtrado', formatMoney(currentTotal), periodLabel(filter), 'green')}
      ${metricCard('Acumulado do ano', formatMoney(yearTotal), yearlyLabel(filter), 'cyan')}
      ${metricCard('Base do repasse', formatMoney(reconciliation.baseReceived), 'Recebido no mes anterior', 'cyan')}
      ${metricCard('% efetivo', reconciliation.rateLabel, 'Faixa esperada: 35% a 40%', reconciliation.tone)}
      ${metricCard('Repasse histórico', formatMoney(lifetimeTotal), 'Desde a fundação do polo', 'green')}
      ${metricCard('Linhas no banco', state.store.repasses.length, 'Histórico consolidado', 'yellow')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Valores repassados da sede', 'Lido direto do banco')}</div>
          <span>${state.store.repasses.length.toLocaleString('pt-BR')} linhas</span>
        </div>
        <p class="muted">Cole os novos repasses na aba Repasse ou Valores Repassados do banco. O sistema usa o valor final informado pela sede para decisão.</p>
        <div class="finance-kpi-list">
          ${progressLine('Repasse consolidado', percent(currentTotal, Math.max(lifetimeTotal, 1)), `${formatMoney(currentTotal)} filtrado / ${formatMoney(lifetimeTotal)} desde a fundação`)}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Comparativo', 'Mês anterior e ano anterior')}</div>
          <span>${escapeHtml(periodLabel(filter))}</span>
        </div>
        <div class="finance-kpi-list">
          ${financeKpiItem('Valor recebido', currentTotal, previousMonthTotal, previousYearTotal)}
          ${financeKpiItem('Base recebida anterior', reconciliation.baseReceived, 0, 0)}
          ${progressLine('Faixa esperada', percent(currentTotal, Math.max(reconciliation.expectedMax, 1)), `${formatMoney(reconciliation.expectedMin)} a ${formatMoney(reconciliation.expectedMax)}`)}
          ${progressLine('Participação no ano', percent(currentTotal, Math.max(yearTotal, 1)), `${formatMoney(currentTotal)} de ${formatMoney(yearTotal)}`)}
        </div>
      </article>
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Evolução dos repasses', 'Histórico mensal')}</div>
          <span>${grouped.length} competências</span>
        </div>
        <div class="mini-columns">
          ${grouped
            .map(
              (item) => `
                <div class="column-item">
                  <span style="height:${item.height}%"></span>
                  <em>${escapeHtml(item.label)}</em>
                  <strong>${formatCompactMoney(item.value)}</strong>
                </div>
              `,
            )
            .join('') || '<p class="muted">Importe o primeiro CSV de repasse.</p>'}
        </div>
      </article>
      <article class="table-panel">
        <div class="panel-heading">
          <div>${smallTitle('Extrato de valores repassados', 'Linhas importadas da sede')}</div>
          <div class="panel-actions">
            <button class="mini-button" type="button" data-export="repasse">Exportar Excel</button>
            <span>${rows.length.toLocaleString('pt-BR')} linhas</span>
          </div>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Competência</th>
                <th>Data</th>
                <th>Descrição</th>
                <th>Curso/Origem</th>
                <th>Valor</th>
                <th>Arquivo</th>
              </tr>
            </thead>
            <tbody>${rows.map(repasseRowTemplate).join('') || emptyRow('Nenhum repasse importado neste filtro.')}</tbody>
          </table>
        </div>
      </article>
    </section>
  `;
}

function renderCursos() {
  const catalog = getCourseCatalog();
  const priced = catalog.filter((course) => Number(course.monthlyFee || 0) > 0);
  const withAuthorization = catalog.filter((course) => cleanText(course.authorization)).length;
  const avgMonthly = priced.length ? sumBy(priced, (course) => Number(course.monthlyFee || 0)) / priced.length : 0;
  els.modules.cursos.innerHTML = `
    ${moduleTitle('Cursos', 'Catalogo comercial com mensalidade, descontos e portaria.')}
    <section class="metric-grid">
      ${metricCard('Cursos cadastrados', catalog.length, 'Sede + inclusoes locais', 'cyan')}
      ${metricCard('Mensalidade media', formatMoney(avgMonthly), `${priced.length} curso(s) com valor`, 'green')}
      ${metricCard('Portarias', withAuthorization, 'Reconhecimento/autorizacao preenchidos', withAuthorization === catalog.length ? 'green' : 'yellow')}
      ${metricCard('Atualizados no polo', catalog.filter((course) => course.local).length, 'Camada local preservada', 'cyan')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Salvar curso', 'Inclua ou atualize os dados comerciais')}</div>
          <span>Valores editaveis</span>
        </div>
        <form class="stack-form" data-form="course">
          <input name="name" placeholder="Curso" required />
          <select name="modality">
            <option>EAD</option>
            <option>Semipresencial</option>
            <option>Presencial</option>
            <option>Pos-graduacao</option>
          </select>
          <input name="habilitation" placeholder="Habilitacao. Ex: Bacharelado, Licenciatura" />
          <input name="duration" placeholder="Duracao. Ex: 8 semestres" />
          <input name="monthlyFee" type="number" min="0" step="0.01" placeholder="Mensalidade" />
          <div class="course-discount-grid">
            ${COURSE_DISCOUNT_RATES.map((rate) => `<input name="discount${rate}" type="number" min="0" step="0.01" placeholder="${rate}% desc." />`).join('')}
          </div>
          <input name="authorization" placeholder="Portaria de Reconhecimento/Autorizacao" />
          <button type="submit">Salvar curso</button>
        </form>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Atualizacao periodica', 'Importe a aba Cursos quando houver reajuste')}</div>
          <span>CSV por curso</span>
        </div>
        <label class="upload-box financial-upload">
          Importar CSV da aba Cursos
          <input type="file" accept=".csv,text/csv" data-course-import />
        </label>
        <p class="muted">Colunas aceitas: Curso, Modalidade, Habilitacao, Duracao, Mensalidade, Desconto de 10% ate 60% e Portaria de Reconhecimento/Autorizacao.</p>
        <button class="mini-button" type="button" data-export="cursos">Exportar catalogo</button>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Tabela de precos e descontos', 'Dados locais tem prioridade sobre a base importada')}</div>
        <span>${catalog.length} cursos</span>
      </div>
      <div class="table-wrap">
        <table class="course-table">
          <thead>
            <tr>
              <th>Curso</th>
              <th>Modalidade</th>
              <th>Habilitacao</th>
              <th>Duracao</th>
              <th>Mensalidade</th>
              <th>Descontos 10% a 60%</th>
              <th>Portaria</th>
              <th>Acao</th>
            </tr>
          </thead>
          <tbody>
            ${catalog.slice(0, 220).map(courseRowTemplate).join('') || emptyRow('Cadastre ou importe a primeira tabela de cursos.')}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderMatriculas() {
  const mode = state.moduleView === 'matricula' ? 'matricula' : 'crm';
  const catalog = getCourseCatalog();
  const leads = state.store.leads;
  const crmStages = LEAD_STAGES.filter((stage) => stage !== 'Matriculado');
  const visibleStages = mode === 'matricula' ? ['Matriculado'] : crmStages;
  const visibleLeads = leads.filter((lead) => visibleStages.includes(lead.stage));
  const signed = leads.filter((lead) => lead.stage === 'Matriculado').length;
  const sourcePendingEnrollments = acompanhamentoPendingEnrollmentRows(state.filteredRows.length ? state.filteredRows : state.allRows, { excludeLeadMatches: true });
  const boletoPending = pendingEnrollmentActionCount(state.filteredRows.length ? state.filteredRows : state.allRows);
  const paymentOk = leads.filter((lead) => isLeadEnrollmentSettled(lead) && !acompanhamentoRowForLead(lead)).length;
  const variableFees = leads.filter((lead) => Number(lead.enrollmentFee || 0) !== Number(state.store.settings.enrollmentFee || 99)).length;

  els.modules.matriculas.innerHTML = `
    ${moduleTitle(mode === 'matricula' ? 'Fechamento de matrícula' : 'Captação de alunos', mode === 'matricula' ? 'Contrato assinado, boleto de matrícula e confirmação manual pelo polo.' : 'Acompanhe candidatos do primeiro contato até o contrato.')}
    <section class="metric-grid">
      ${metricCard(mode === 'matricula' ? 'Matriculados' : 'Candidatos ativos', mode === 'matricula' ? signed : visibleLeads.length, mode === 'matricula' ? 'Contratos assinados' : 'Base de captação local')}
      ${metricCard('Contratos assinados', signed, 'Status Matriculado', 'green')}
      ${metricCard('Matrículas a regularizar', boletoPending, 'Boleto ou baixa pendente', boletoPending ? 'yellow' : 'green')}
      ${metricCard('Taxas combinadas', variableFees, 'Curso, bolsa e desconto', 'cyan')}
      ${metricCard('Matrículas confirmadas', paymentOk, 'Boleto pago ou isento', 'green')}
    </section>
    <section class="split-grid">
      <article class="panel ${mode === 'matricula' ? 'muted-panel' : ''}">
        <div class="panel-heading">
          <div>${smallTitle('Novo candidato', 'Escolha um curso cadastrado')}</div>
          <span>${catalog.length} cursos validados</span>
        </div>
        <form class="stack-form" data-form="lead">
          <input name="name" placeholder="Nome do candidato" required />
          <div class="two-cols">
            <input name="phone" placeholder="Telefone/WhatsApp" />
            <input name="origin" placeholder="Origem comercial" />
          </div>
          <select name="course" required>
            <option value="">Curso validado</option>
            ${catalog.map((course) => `<option>${escapeHtml(course.name)}</option>`).join('')}
          </select>
          <div class="two-cols">
            <select name="stage">${LEAD_STAGES.map((stage) => `<option value="${escapeHtml(stage)}">${escapeHtml(leadStageLabel(stage))}</option>`).join('')}</select>
            <input name="enrollmentFee" type="number" min="0" step="0.01" value="${Number(state.store.settings.enrollmentFee || 99)}" placeholder="Taxa de matrícula combinada" />
          </div>
          <button type="submit">Salvar candidato</button>
        </form>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Depois da assinatura', 'Boleto e taxa de matrícula')}</div>
          <span>Confirmação pelo polo</span>
        </div>
        <div class="decision-list">
          <div class="decision-signal warning"><strong>Boleto a enviar</strong><span>Aparece quando o contrato chega em Matriculado sem confirmação de envio.</span></div>
          <div class="decision-signal success"><strong>Taxa confirmada</strong><span>Usado quando a taxa de matrícula foi paga ou marcada como isenta.</span></div>
        </div>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle(mode === 'matricula' ? 'Conferência da matrícula' : 'Funil de captação', mode === 'matricula' ? 'Confira boleto, taxa acordada e confirmação pelo polo' : 'Arraste ou use os botões para mudar de etapa')}</div>
        <div class="panel-actions">
          <button class="mini-button" type="button" data-export="leads">Exportar Excel</button>
          <span>${visibleLeads.length} registros</span>
        </div>
      </div>
      <div class="kanban-board">
        ${visibleStages.map((stage) => leadColumn(stage)).join('')}
      </div>
    </section>
    ${
      mode === 'matricula'
        ? `
          <section class="table-panel">
            <div class="panel-heading">
              <div>${smallTitle('Acompanhamento desde 01/04/2026', 'Matriculados na sede com boleto ou baixa pendente')}</div>
              <span>${sourcePendingEnrollments.length} pendentes sem lead duplicado</span>
            </div>
            <div class="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Aluno</th>
                    <th>Curso</th>
                    <th>Boleto</th>
                    <th>Pagamento</th>
                    <th>Ação</th>
                  </tr>
                </thead>
                <tbody>
                  ${sourcePendingEnrollments.map(sourceEnrollmentRowTemplate).join('') || emptyRow('Nenhum aluno novo da base Acompanhamento pendente de boleto ou baixa.')}
                </tbody>
              </table>
            </div>
          </section>
        `
        : ''
    }
  `;
}

function renderMetas() {
  const leads = state.store.leads.length;
  const matriculas = state.allRows.filter((row) => isActive(row)).length + confirmedNewMatriculations();
  const target = Number(state.store.settings.monthlyTarget || 65);
  const conversion = leads + matriculas ? Math.round((matriculas / (leads + matriculas)) * 100) : 0;
  const gap = matriculas - target;
  const monthProgress = Math.min(100, Math.round((new Date().getDate() / daysInCurrentMonth()) * 100));
  const projection = Math.round(matriculas / Math.max(monthProgress, 1) * 100);
  const planning = planningContext();
  const sellers = sellerPerformanceRows();

  els.modules.metas.innerHTML = `
    ${moduleTitle('Metas do polo', 'Acompanhe meta, resultado e previsão de fechamento.')}
    <section class="metric-grid">
      ${metricCard('Meta mensal', target, 'Editável pelo polo', 'cyan')}
      ${metricCard('Matrículas', matriculas, 'Realizado acumulado', 'green')}
      ${metricCard('Falta ou sobra', gap, gap >= 0 ? 'Acima da meta' : 'Abaixo da meta', gap >= 0 ? 'green' : 'red')}
      ${metricCard('Viraram matrícula', `${conversion}%`, 'Candidatos vs. matrículas', 'yellow')}
      ${metricCard('5W2H próximo mês', `${planning.monthlyProgress}%`, `${planning.monthly.length} ações`, planning.monthlyProgress >= 80 ? 'green' : 'cyan')}
      ${metricCard('Semana vendedores', `${planning.weeklyProgress}%`, `${planning.weekly.length} ações atuais`, planning.weeklyProgress >= 80 ? 'green' : 'yellow')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Editar meta', 'Ajuste conforme o mês')}</div>
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
          <div>${smallTitle('Equipe de captação', 'Base local do funil')}</div>
          <div class="panel-actions">
            <button class="mini-button" type="button" data-export="vendedores">Exportar vendedores</button>
            <span>${state.store.leads.length} candidatos</span>
          </div>
        </div>
        <div class="rank-list">
          ${consultantRanking().map((item, index) => `<div><strong>${index + 1}. ${escapeHtml(item.name)}</strong><span>${item.total} registros · ${item.conversion}% conversão</span></div>`).join('')}
        </div>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Painel individual dos vendedores', 'Esforço, funil, conversão e semana atual')}</div>
        <div class="panel-actions">
          <button class="mini-button" type="button" data-report="vendedores">Relatório PDF</button>
          <button class="mini-button" type="button" data-export="vendedores">Exportar Excel</button>
          <span>${sellers.length} vendedores</span>
        </div>
      </div>
      <div class="table-wrap">
        ${sellerPerformanceTable(sellers)}
      </div>
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Previsão de meta', 'Ritmo atual do mês')}</div>
          <span>${monthProgress}% do mês</span>
        </div>
        <div class="decision-list">
          ${forecastSignals(state.allRows).map((item) => `<div class="decision-signal ${item.tone}"><strong>${escapeHtml(item.title)}</strong><span>${escapeHtml(item.text)}</span></div>`).join('')}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Relatórios comerciais', 'Gestão e cobrança semanal')}</div>
          <span>PDF e Excel</span>
        </div>
        <div class="decision-list">
          <button class="decision-signal success action-signal" type="button" data-export="5w2h"><strong>5W2H para Excel</strong><span>Exporta ações do próximo mês e da semana atual.</span></button>
          <button class="decision-signal info action-signal" type="button" data-report="vendedores"><strong>Relatório dos vendedores</strong><span>Abre uma versão pronta para reunião e impressão.</span></button>
        </div>
      </article>
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Plano 5W2H do próximo mês', 'Ação clara, dono único, prazo, KPI e custo')}</div>
          <span>${escapeHtml(planning.nextMonth.label)}</span>
        </div>
        <form class="stack-form" data-form="planning-5w2h">
          <input type="hidden" name="planType" value="monthly" />
          <input type="hidden" name="periodKey" value="${escapeHtml(planning.nextMonth.key)}" />
          <select name="area" required>
            <option>Comercial</option>
            <option>Retenção</option>
            <option>Acadêmico</option>
            <option>Financeiro</option>
            <option>Infraestrutura</option>
          </select>
          <input name="what" placeholder="O quê será feito? Ex: campanha de rematrícula ativa" required />
          <textarea name="why" rows="2" placeholder="Por quê? Resultado esperado ou problema que resolve" required></textarea>
          <div class="two-cols">
            <input name="where" placeholder="Onde? Canal, bairro, escola, WhatsApp, polo" required />
            <input name="when" type="date" value="${escapeHtml(planning.nextMonth.end)}" required />
          </div>
          <div class="two-cols">
            <select name="who" required>${sellerOptions()}</select>
            <input name="kpi" placeholder="Indicador/KPI. Ex: 30 leads, 12 matrículas" required />
          </div>
          <textarea name="how" rows="2" placeholder="Como será executado? Passos concretos, não intenção genérica" required></textarea>
          <div class="two-cols">
            <input name="howMuch" placeholder="Quanto custa? Ex: R$ 300 ou sem custo" required />
            <select name="status">
              <option>Planejada</option>
              <option>Em andamento</option>
              <option>Concluída</option>
              <option>Bloqueada</option>
            </select>
          </div>
          <button type="submit">Adicionar ao plano do próximo mês</button>
        </form>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Semana atual dos vendedores', 'O que cada vendedor precisa executar agora')}</div>
          <span>${escapeHtml(planning.week.label)}</span>
        </div>
        <form class="stack-form" data-form="planning-5w2h">
          <input type="hidden" name="planType" value="weeklySeller" />
          <input type="hidden" name="weekStart" value="${escapeHtml(planning.week.start)}" />
          <input type="hidden" name="area" value="Comercial" />
          <select name="who" required>${sellerOptions()}</select>
          <input name="what" placeholder="O quê nesta semana? Ex: resgatar leads sem resposta" required />
          <textarea name="why" rows="2" placeholder="Por quê? Ex: aumentar comparecimento e fechamento" required></textarea>
          <div class="two-cols">
            <input name="where" placeholder="Onde? WhatsApp, ligação, escola, balcão" required />
            <input name="when" type="date" value="${escapeHtml(planning.week.today)}" min="${escapeHtml(planning.week.start)}" max="${escapeHtml(planning.week.end)}" required />
          </div>
          <textarea name="how" rows="2" placeholder="Como? Sequência prática: contato, follow-up, oferta, registro no CRM" required></textarea>
          <div class="two-cols">
            <input name="kpi" placeholder="Meta da semana. Ex: 40 contatos, 8 visitas" required />
            <input name="howMuch" placeholder="Custo/recurso. Ex: sem custo" required />
          </div>
          <select name="status">
            <option>Planejada</option>
            <option>Em andamento</option>
            <option>Concluída</option>
            <option>Bloqueada</option>
          </select>
          <button type="submit">Adicionar plano da semana</button>
        </form>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('5W2H do próximo mês', 'Revisão mensal antes da execução')}</div>
        <div class="panel-actions">
          <span>${planning.monthly.length} ações · ${planning.monthlyProgress}% concluído</span>
        </div>
      </div>
      <div class="table-wrap">
        ${planningTable(planning.monthly, 'Nenhuma ação 5W2H cadastrada para o próximo mês.')}
      </div>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Semana atual por vendedor', 'Rotina comercial da semana corrente')}</div>
        <div class="panel-actions">
          <span>${planning.weekly.length} ações · ${planning.weeklyProgress}% concluído</span>
        </div>
      </div>
      <div class="table-wrap">
        ${planningTable(planning.weekly, 'Nenhuma ação cadastrada para a semana atual dos vendedores.')}
      </div>
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Boas práticas 5W2H', 'Para gerar execução, não só intenção')}</div>
          <span>Gestão</span>
        </div>
        <div class="decision-list">
          ${planningBestPractices().map((item) => `<div class="decision-signal ${item.tone}"><strong>${escapeHtml(item.title)}</strong><span>${escapeHtml(item.text)}</span></div>`).join('')}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Ritual recomendado', 'Cadência simples')}</div>
          <span>Mensal + semanal</span>
        </div>
        <div class="decision-list">
          <div class="decision-signal info"><strong>Sexta-feira</strong><span>Revisar a semana dos vendedores e travar os próximos compromissos.</span></div>
          <div class="decision-signal warning"><strong>Última semana do mês</strong><span>Preencher o 5W2H do mês seguinte com responsáveis e indicadores.</span></div>
          <div class="decision-signal success"><strong>Segunda-feira</strong><span>Começar a semana com cada vendedor sabendo o que deve entregar.</span></div>
        </div>
      </article>
    </section>
  `;
}

function planningContext() {
  const nextMonth = nextMonthPlanningPeriod();
  const week = currentWeekPeriod();
  const monthly = state.store.decisions
    .filter((item) => item.planType === 'monthly' && item.periodKey === nextMonth.key)
    .sort((a, b) => timelineTime(a.when) - timelineTime(b.when));
  const weekly = state.store.decisions
    .filter((item) => item.planType === 'weeklySeller' && item.weekStart === week.start)
    .sort((a, b) => collator.compare(a.who || '', b.who || '') || timelineTime(a.when) - timelineTime(b.when));
  return {
    nextMonth,
    week,
    monthly,
    weekly,
    monthlyProgress: planningProgress(monthly),
    weeklyProgress: planningProgress(weekly),
  };
}

function nextMonthPlanningPeriod(baseDate = new Date()) {
  const start = new Date(baseDate.getFullYear(), baseDate.getMonth() + 1, 1);
  const end = new Date(baseDate.getFullYear(), baseDate.getMonth() + 2, 0);
  return {
    key: `${start.getFullYear()}-${String(start.getMonth() + 1).padStart(2, '0')}`,
    label: start.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' }),
    start: isoDate(start),
    end: isoDate(end),
  };
}

function currentWeekPeriod(baseDate = new Date()) {
  const dayIndex = (baseDate.getDay() + 6) % 7;
  const start = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate() - dayIndex);
  const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6);
  return {
    start: isoDate(start),
    end: isoDate(end),
    today: isoDate(baseDate),
    label: `${start.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' })} a ${end.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' })}`,
  };
}

function isoDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function sellerNames() {
  const names = new Set();
  (state.users.length ? state.users : defaultUsers())
    .filter((user) => user.ativo !== 'NAO' && normalizeProfile(user.perfil) === 'consultor')
    .forEach((user) => names.add(user.nome || user.usuario));
  state.store.leads.forEach((lead) => names.add(lead.consultant || 'Atendimento'));
  names.add('Atendimento');
  return [...names].filter(Boolean).sort(collator.compare);
}

function sellerOptions(selected = '') {
  return sellerNames()
    .map((name) => `<option value="${escapeHtml(name)}" ${name === selected ? 'selected' : ''}>${escapeHtml(name)}</option>`)
    .join('');
}

function planningProgress(actions) {
  if (!actions.length) return 0;
  const done = actions.filter((item) => normalize(item.status).includes('conclu')).length;
  return Math.round((done / actions.length) * 100);
}

function planningTable(actions, emptyMessage) {
  return `
    <table>
      <thead>
        <tr>
          <th>O quê</th>
          <th>Quem</th>
          <th>Quando</th>
          <th>Onde</th>
          <th>Como</th>
          <th>KPI</th>
          <th>Custo</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        ${actions.map(planningRowTemplate).join('') || emptyRow(emptyMessage)}
      </tbody>
    </table>
  `;
}

function planningRowTemplate(item) {
  return `
    <tr>
      <td>
        <strong>${escapeHtml(item.what || item.title)}</strong>
        <span class="subtext">${escapeHtml(item.why || 'Sem justificativa registrada')}</span>
      </td>
      <td>${escapeHtml(item.who || '-')}</td>
      <td>${escapeHtml(formatShortDate(item.when))}</td>
      <td>${escapeHtml(item.where || '-')}</td>
      <td>${escapeHtml(item.how || '-')}</td>
      <td>${escapeHtml(item.kpi || '-')}</td>
      <td>${escapeHtml(item.howMuch || '-')}</td>
      <td>
        <select class="status-select ${planningStatusClass(item.status)}" data-plan-status="${escapeHtml(item.id)}">
          ${['Planejada', 'Em andamento', 'Concluída', 'Bloqueada'].map((status) => `<option ${normalize(status) === normalize(item.status) ? 'selected' : ''}>${status}</option>`).join('')}
        </select>
      </td>
    </tr>
  `;
}

function planningStatusClass(status) {
  const text = normalize(status);
  if (text.includes('conclu')) return 'green';
  if (text.includes('bloque')) return 'red';
  if (text.includes('andamento')) return 'yellow';
  return 'cyan';
}

function formatShortDate(value) {
  const date = parseBrazilianDate(value);
  return date ? date.toLocaleDateString('pt-BR') : cleanText(value) || '-';
}

function planningBestPractices() {
  return [
    {
      tone: 'success',
      title: 'Um dono por ação',
      text: 'Cada ação precisa ter uma pessoa responsável, evitando tarefas sem comando claro.',
    },
    {
      tone: 'warning',
      title: 'Prazo e KPI obrigatórios',
      text: 'Toda ação deve ter data e indicador mensurável para permitir cobrança semanal.',
    },
    {
      tone: 'info',
      title: 'Como e quanto sem vazio',
      text: 'O plano deve dizer como será executado e qual recurso ou custo será usado.',
    },
    {
      tone: 'danger',
      title: 'Evite verbos vagos',
      text: 'Troque “melhorar vendas” por ações concretas como ligar, visitar, recuperar ou converter.',
    },
  ];
}

function forecastSignals(rows = state.allRows) {
  const target = Number(state.store.settings.monthlyTarget || 65);
  const confirmed = confirmedNewMatriculations();
  const progress = Math.max(1, Math.round((new Date().getDate() / daysInCurrentMonth()) * 100));
  const projection = Math.round((confirmed / progress) * 100);
  const highRisk = rows.filter((row) => row.risk.level === 'Alto').length;
  const noAva = rows.filter((row) => ['red', 'yellow'].includes(row.avaAlert.level)).length;
  const debtRows = hasBillingReceiptSource() ? buildFinancialPosition(rows).debtRows : rows.filter((row) => row.isDebt);
  const latest = latestSnapshot();
  const previous = state.store.snapshots.slice().sort((a, b) => collator.compare(b.periodKey, a.periodKey))[1];
  const debtTrend = latest && previous ? Number(latest.debtTotal || 0) - Number(previous.debtTotal || 0) : 0;
  const neededPerWeek = Math.max(0, Math.ceil((target - confirmed) / Math.max(weeksLeftInMonth(), 1)));
  return [
    {
      tone: projection >= target ? 'success' : projection >= target * 0.8 ? 'warning' : 'danger',
      title: 'Fechamento da meta',
      text: `Ritmo atual projeta ${projection} matrícula(s) confirmadas. Faltam ${Math.max(0, target - confirmed)} no mês, cerca de ${neededPerWeek} por semana.`,
    },
    {
      tone: highRisk ? 'warning' : 'success',
      title: 'Risco de evasão',
      text: `${highRisk} aluno(s) em risco alto e ${noAva} com alerta de acesso. Priorize contato antes de virar evasão.`,
    },
    {
      tone: debtRows.length ? 'danger' : 'success',
      title: 'Tendência financeira',
      text: debtTrend
        ? `Inadimplência variou ${formatMoney(debtTrend)} frente ao snapshot anterior. Carteira atual tem ${debtRows.length} pendência(s).`
        : `Carteira atual tem ${debtRows.length} pendência(s). Use snapshots mensais para comparar tendência.`,
    },
  ];
}

function weeksLeftInMonth(date = new Date()) {
  const last = new Date(date.getFullYear(), date.getMonth() + 1, 0);
  return Math.max(1, Math.ceil((last.getDate() - date.getDate() + 1) / 7));
}

function sellerPerformanceRows() {
  const sellers = sellerNames();
  const planning = planningContext();
  return sellers.map((seller) => {
    const leads = state.store.leads.filter((lead) => normalize(lead.consultant || 'Atendimento') === normalize(seller));
    const confirmed = leads.filter(isLeadEnrollmentSettled).length;
    const pending = leads.filter(leadRequiresVisibleEnrollmentAction).length;
    const visits = leads.filter((lead) => normalize(lead.stage).includes('visita')).length;
    const contracts = leads.filter((lead) => normalize(lead.stage).includes('contrato')).length;
    const active = leads.filter((lead) => !['Matriculado', 'Perdido'].includes(lead.stage)).length;
    const weeklyActions = planning.weekly.filter((item) => normalize(item.who) === normalize(seller));
    const weeklyDone = weeklyActions.filter((item) => normalize(item.status).includes('conclu')).length;
    return {
      seller,
      total: leads.length,
      active,
      visits,
      contracts,
      confirmed,
      pending,
      conversion: leads.length ? Math.round((confirmed / leads.length) * 100) : 0,
      avgTicket: confirmed ? Math.round(sumBy(leads.filter(isLeadEnrollmentSettled), (lead) => lead.enrollmentFee || state.store.settings.enrollmentFee || 99) / confirmed) : 0,
      weeklyActions: weeklyActions.length,
      weeklyDone,
      weeklyProgress: weeklyActions.length ? Math.round((weeklyDone / weeklyActions.length) * 100) : 0,
    };
  }).sort((a, b) => b.confirmed - a.confirmed || b.conversion - a.conversion || collator.compare(a.seller, b.seller));
}

function sellerPerformanceTable(rows) {
  return `
    <table>
      <thead>
        <tr>
          <th>Vendedor</th>
          <th>Leads</th>
          <th>Ativos</th>
          <th>Visitas</th>
          <th>Contratos</th>
          <th>Matrículas</th>
          <th>Pendências</th>
          <th>Conversão</th>
          <th>Semana 5W2H</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map((row) => `
          <tr>
            <td><strong>${escapeHtml(row.seller)}</strong><span class="subtext">Ticket médio ${formatMoney(row.avgTicket)}</span></td>
            <td>${row.total}</td>
            <td>${row.active}</td>
            <td>${row.visits}</td>
            <td>${row.contracts}</td>
            <td><span class="badge green">${row.confirmed}</span></td>
            <td><span class="badge ${row.pending ? 'yellow' : 'green'}">${row.pending}</span></td>
            <td>${row.conversion}%</td>
            <td><span class="badge ${row.weeklyProgress >= 80 ? 'green' : row.weeklyActions ? 'yellow' : 'cyan'}">${row.weeklyDone}/${row.weeklyActions}</span></td>
          </tr>
        `).join('') || emptyRow('Nenhum vendedor com dados de funil ainda.')}
      </tbody>
    </table>
  `;
}

function syncDecisionFoundation() {
  if (!state.allRows.length) return;
  let changed = upsertCurrentSnapshot(false);
  changed = ensureAutomaticTaskRecords(state.allRows) || changed;
  if (changed) persist();
}

function currentPeriodKey(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function periodFilterFromKey(periodKey) {
  const [year, month] = cleanText(periodKey).split('-').map(Number);
  return { month: Number(month || 0), year: Number(year || new Date().getFullYear()) };
}

function periodLabelFromKey(periodKey) {
  const filter = periodFilterFromKey(periodKey);
  if (!filter.month) return periodKey || '-';
  return `${MONTHS[filter.month - 1]}/${filter.year}`;
}

function upsertCurrentSnapshot(force = false) {
  if (!state.allRows.length) return false;
  const key = currentPeriodKey();
  const index = state.store.snapshots.findIndex((item) => item.periodKey === key);
  if (index >= 0 && !force) return false;
  const snapshot = buildManagementSnapshot(key);
  if (index >= 0) state.store.snapshots[index] = snapshot;
  else state.store.snapshots.push(snapshot);
  state.store.snapshots = state.store.snapshots
    .slice()
    .sort((a, b) => collator.compare(a.periodKey, b.periodKey))
    .slice(-36);
  return true;
}

function buildManagementSnapshot(periodKey = currentPeriodKey()) {
  const rows = state.allRows;
  const filter = periodFilterFromKey(periodKey);
  const financial = buildFinancialPosition(rows);
  const activeRows = rows.filter(isActive);
  const noAvaRows = rows.filter((row) => ['red', 'yellow'].includes(row.avaAlert.level));
  const quality = dataQualityIssues(rows);
  return normalizeSnapshot({
    periodKey,
    active: activeRows.length,
    totalStudents: rows.length,
    highRisk: rows.filter((row) => row.risk.level === 'Alto').length,
    noAva: noAvaRows.length,
    retention: rows.length ? Math.round((activeRows.length / rows.length) * 100) : 0,
    leads: state.store.leads.length,
    confirmedMatriculations: confirmedNewMatriculations(),
    pendingEnrollments: pendingEnrollmentActionCount(rows),
    debtCount: financial.debtRows.length || rows.filter((row) => row.isDebt).length,
    debtTotal: sumBy(financial.debtRows.length ? financial.debtRows : rows.filter((row) => row.isDebt), (row) => row.debtValue),
    enrollmentRevenue: enrollmentRevenue(filter),
    recurringRevenue: recurringRevenue(filter),
    repasse: totalFinancialRecords(state.store.repasses, filter),
    qualityScore: dataQualityScore(quality, rows.length),
    qualityIssues: quality.reduce((total, item) => total + item.count, 0),
    createdAt: new Date().toISOString(),
  });
}

function latestSnapshot() {
  return state.store.snapshots.slice().sort((a, b) => collator.compare(b.periodKey, a.periodKey))[0] || null;
}

function snapshotSummaryCard(snapshot) {
  if (!snapshot) return '<div class="empty-state">Ainda não há snapshot mensal. Abra Indicadores para criar o primeiro retrato.</div>';
  return `
    <div class="finance-kpi-list">
      <div class="finance-kpi-item"><div><strong>Período</strong><span>${escapeHtml(periodLabelFromKey(snapshot.periodKey))}</span></div><span class="badge cyan">Snapshot</span></div>
      <div class="finance-kpi-item"><div><strong>Alunos ativos</strong><span>${snapshot.active.toLocaleString('pt-BR')} de ${snapshot.totalStudents.toLocaleString('pt-BR')}</span></div><span class="badge green">${snapshot.retention}% retenção</span></div>
      <div class="finance-kpi-item"><div><strong>Pendências críticas</strong><span>${snapshot.highRisk} risco alto · ${snapshot.pendingEnrollments} matrículas</span></div><span class="badge ${snapshot.qualityScore >= 85 ? 'green' : 'yellow'}">${snapshot.qualityScore}% dados</span></div>
    </div>
  `;
}

function snapshotTable(snapshots) {
  return `
    <table>
      <thead>
        <tr>
          <th>Período</th>
          <th>Ativos</th>
          <th>Retenção</th>
          <th>Risco alto</th>
          <th>Matrículas pendentes</th>
          <th>Inadimplência</th>
          <th>Recebido</th>
          <th>Qualidade</th>
        </tr>
      </thead>
      <tbody>
        ${snapshots
          .slice(0, 18)
          .map(
            (item) => `
              <tr>
                <td><span class="badge cyan">${escapeHtml(periodLabelFromKey(item.periodKey))}</span></td>
                <td>${Number(item.active || 0).toLocaleString('pt-BR')}</td>
                <td>${Number(item.retention || 0)}%</td>
                <td>${Number(item.highRisk || 0).toLocaleString('pt-BR')}</td>
                <td>${Number(item.pendingEnrollments || 0).toLocaleString('pt-BR')}</td>
                <td>${formatMoney(item.debtTotal)}</td>
                <td>${formatMoney(Number(item.enrollmentRevenue || 0) + Number(item.recurringRevenue || 0))}</td>
                <td><span class="badge ${Number(item.qualityScore || 0) >= 85 ? 'green' : 'yellow'}">${Number(item.qualityScore || 0)}%</span></td>
              </tr>
            `,
          )
          .join('') || emptyRow('Nenhum snapshot mensal registrado ainda.')}
      </tbody>
    </table>
  `;
}

function indicatorDictionaryRows(context) {
  const rows = context.rows || state.allRows;
  const financial = buildFinancialPosition(rows);
  const pending = pendingEnrollmentActionCount(rows);
  const noAva = rows.filter((row) => ['red', 'yellow'].includes(row.avaAlert.level)).length;
  const active = rows.filter(isActive).length;
  const planning = planningContext();
  return [
    {
      name: 'Matrícula confirmada',
      formula: 'Lead matriculado com boleto pago ou isento + aluno ativo da base Acompanhamento',
      source: 'CRM, Acompanhamento e overrides do polo',
      owner: 'Gestão comercial',
      current: `${(context.commercial?.monthlyMatriculations || 0).toLocaleString('pt-BR')} no período`,
      health: pending ? 'Atenção' : 'Ok',
    },
    {
      name: 'Aluno ativo',
      formula: 'Status Ativo, Pré-matriculado, Matriculado ou Transferência, respeitando override local',
      source: 'Acompanhamento + Dados do Polo',
      owner: 'Secretaria / Retenção',
      current: `${active.toLocaleString('pt-BR')} alunos`,
      health: active ? 'Ok' : 'Atenção',
    },
    {
      name: 'Retenção acadêmica',
      formula: 'Alunos ativos ÷ total da base filtrada',
      source: 'Acompanhamento',
      owner: 'Retenção',
      current: `${rows.length ? Math.round((active / rows.length) * 100) : 0}%`,
      health: rows.length && active / rows.length >= 0.7 ? 'Ok' : 'Atenção',
    },
    {
      name: 'Risco AVA',
      formula: 'Alunos com 5+ dias sem acesso ou sem registro confiável de acesso',
      source: 'Último acesso AVA / One',
      owner: 'Retenção',
      current: `${noAva.toLocaleString('pt-BR')} alunos`,
      health: noAva ? 'Atenção' : 'Ok',
    },
    {
      name: 'Inadimplência real',
      formula: 'Faturamento da sede menos Recebimento da sede por aluno e mês',
      source: 'Faturamento e Recebimento append-only',
      owner: 'Financeiro',
      current: formatMoney(sumBy(financial.debtRows, (row) => row.debtValue)),
      health: financial.debtRows.length ? 'Atenção' : 'Ok',
    },
    {
      name: 'Receita do polo',
      formula: 'Taxas de matrícula pagas + mensalidades recebidas no período',
      source: 'CRM, overrides, Recebimento',
      owner: 'Administrador',
      current: formatMoney((context.finance?.enrollment.current || 0) + (context.finance?.recurring.current || 0)),
      health: 'Ok',
    },
    {
      name: 'Meta comercial',
      formula: 'Matrículas confirmadas ÷ meta mensal/anual cadastrada',
      source: 'CRM, Acompanhamento e metas',
      owner: 'Gestão comercial',
      current: `${context.monthlyProgress || 0}% mês / ${context.annualProgress || 0}% ano`,
      health: (context.monthlyProgress || 0) >= 80 ? 'Ok' : 'Atenção',
    },
    {
      name: 'Execução 5W2H',
      formula: 'Ações concluídas ÷ ações planejadas para próximo mês e semana atual',
      source: 'Plano 5W2H',
      owner: 'Gestor do polo',
      current: `${planning.monthlyProgress}% mensal / ${planning.weeklyProgress}% semanal`,
      health: planning.monthlyProgress >= 70 || planning.weeklyProgress >= 70 ? 'Ok' : 'Atenção',
    },
  ];
}

function indicatorDictionaryTable(rows) {
  return `
    <table>
      <thead>
        <tr>
          <th>Indicador</th>
          <th>Fórmula oficial</th>
          <th>Fonte</th>
          <th>Dono</th>
          <th>Valor atual</th>
          <th>Saúde</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (item) => `
              <tr>
                <td><strong>${escapeHtml(item.name)}</strong></td>
                <td>${escapeHtml(item.formula)}</td>
                <td>${escapeHtml(item.source)}</td>
                <td>${escapeHtml(item.owner)}</td>
                <td>${escapeHtml(item.current)}</td>
                <td><span class="badge ${item.health === 'Ok' ? 'green' : 'yellow'}">${escapeHtml(item.health)}</span></td>
              </tr>
            `,
          )
          .join('')}
      </tbody>
    </table>
  `;
}

function dataQualityIssues(rows = state.allRows) {
  const issues = [];
  const addIssue = (severity, title, count, recommendation) => {
    if (count > 0) issues.push({ severity, title, count, recommendation });
  };
  const cpfGroups = duplicateGroups(rows, (row) => cleanText(row.cpf).replace(/\D/g, ''), (key) => key.length === 11);
  const raGroups = duplicateGroups(rows, (row) => normalize(row.ra), Boolean);
  const invalidPhones = rows.filter((row) => {
    const digits = cleanText(row.phone).replace(/\D/g, '');
    return !digits || digits.length < 10;
  }).length;
  const invalidEmails = rows.filter((row) => !cleanText(row.email) || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanText(row.email))).length;
  const missingPeriod = rows.filter((row) => noValue(row.startPeriod)).length;
  const missingCourse = rows.filter((row) => noValue(row.course) || normalize(row.course).includes('sem curso')).length;
  const leadDuplicates = state.store.leads.filter(acompanhamentoRowForLead).length;
  const billingUnmatched = state.store.billing.filter((record) => !findStudentForFinancial(record)).length;
  const receiptUnmatched = state.store.receipts.filter((record) => !findStudentForFinancial(record)).length;
  const enrollmentPending = enrollmentActionItems(rows).length;

  addIssue('red', 'CPF duplicado na base', cpfGroups.length, 'Conferir documentos antes de cobrar, matricular ou exportar relatório.');
  addIssue('red', 'RA duplicado na base', raGroups.length, 'Validar registro acadêmico para evitar aluno contado duas vezes.');
  addIssue('yellow', 'CRM e Acompanhamento com possível duplicidade', leadDuplicates, 'Vincular pelo mesmo telefone/nome e manter uma única cobrança.');
  addIssue('yellow', 'Telefone ausente ou inválido', invalidPhones, 'Atualizar WhatsApp local para garantir contato e retenção.');
  addIssue('yellow', 'E-mail ausente ou inválido', invalidEmails, 'Corrigir e-mail para comunicação oficial e documentos.');
  addIssue('yellow', 'Período inicial sem padrão', missingPeriod, 'Padronizar safra para análise de retenção.');
  addIssue('yellow', 'Curso ausente ou inválido', missingCourse, 'Vincular todos os alunos a um curso cadastrado.');
  addIssue('red', 'Faturamento sem aluno identificado', billingUnmatched, 'Conferir CPF, RA ou nome antes de analisar inadimplência.');
  addIssue('red', 'Recebimento sem aluno identificado', receiptUnmatched, 'Conferir recebimentos para não subestimar baixa financeira.');
  addIssue('yellow', 'Matrículas novas sem boleto ou baixa', enrollmentPending, 'Confirmar envio de boleto, pagamento ou isenção.');

  return issues.sort((a, b) => severityWeight(b.severity) - severityWeight(a.severity) || b.count - a.count);
}

function duplicateGroups(rows, getter, valid = Boolean) {
  return Object.values(groupBy(rows, getter)).filter((items) => valid(getter(items[0])) && items.length > 1);
}

function severityWeight(value) {
  return value === 'red' ? 3 : value === 'yellow' ? 2 : 1;
}

function dataQualityScore(issues, totalRows = state.allRows.length) {
  const penalty = issues.reduce((total, issue) => total + issue.count * severityWeight(issue.severity), 0);
  const base = Math.max(Number(totalRows || 0), 1);
  return Math.max(0, Math.min(100, 100 - Math.round((penalty / base) * 4)));
}

function qualitySignalTemplate(item) {
  return `
    <div class="decision-signal ${item.severity === 'red' ? 'danger' : 'warning'}">
      <strong>${escapeHtml(item.title)} (${item.count.toLocaleString('pt-BR')})</strong>
      <span>${escapeHtml(item.recommendation)}</span>
    </div>
  `;
}

function qualityTable(issues) {
  return `
    <table>
      <thead>
        <tr>
          <th>Severidade</th>
          <th>Problema</th>
          <th>Qtd.</th>
          <th>Ação recomendada</th>
        </tr>
      </thead>
      <tbody>
        ${issues
          .map(
            (item) => `
              <tr>
                <td><span class="badge ${item.severity}">${item.severity === 'red' ? 'Crítico' : 'Atenção'}</span></td>
                <td><strong>${escapeHtml(item.title)}</strong></td>
                <td>${item.count.toLocaleString('pt-BR')}</td>
                <td>${escapeHtml(item.recommendation)}</td>
              </tr>
            `,
          )
          .join('') || emptyRow('Nenhum problema relevante de qualidade encontrado.')}
      </tbody>
    </table>
  `;
}

function automaticTaskItems(rows = state.allRows) {
  const today = isoDate(new Date());
  const upcoming = new Date();
  upcoming.setDate(upcoming.getDate() + 3);
  const tasks = [];
  enrollmentActionItems(rows).forEach((item) => {
    tasks.push({
      id: `enrollment:${item.kind}:${item.id || item.key}`,
      area: 'Matrícula',
      priority: item.boletoOk ? 'Alta' : 'Crítica',
      owner: item.boletoOk ? 'Financeiro' : 'Vendas',
      title: item.boletoOk ? `Confirmar baixa de ${item.name}` : `Enviar boleto para ${item.name}`,
      due: today,
      source: item.origin,
    });
  });
  rows
    .filter((row) => isRetentionEligible(row) && ['red', 'yellow'].includes(row.avaAlert.level) && !row.retention.contacted)
    .slice(0, 120)
    .forEach((row) => {
      tasks.push({
        id: `retention:${row.key}`,
        area: 'Retenção',
        priority: row.avaAlert.level === 'red' ? 'Crítica' : 'Alta',
        owner: 'Retenção',
        title: `Contatar ${row.name} sobre acesso ao AVA`,
        due: today,
        source: `${row.avaDaysNumber >= 0 ? `${row.avaDaysNumber} dias` : 'sem dado'} sem acesso`,
      });
    });
  const financialRows = hasBillingReceiptSource() ? buildFinancialPosition(rows).debtRows : rows.filter((row) => row.isDebt);
  financialRows.slice(0, 80).forEach((row) => {
    tasks.push({
      id: `finance:${row.key}`,
      area: 'Financeiro',
      priority: Number(row.debtValue || 0) >= 500 ? 'Crítica' : 'Alta',
      owner: 'Financeiro',
      title: `Acompanhar pagamento de ${row.name}`,
      due: today,
      source: canSeeFinancial() ? formatMoney(row.debtValue) : 'inadimplência registrada',
    });
  });
  state.store.decisions
    .filter((item) => !normalize(item.status).includes('conclu'))
    .filter((item) => {
      const date = parseBrazilianDate(item.when || item.due);
      return date && date <= upcoming;
    })
    .forEach((item) => {
      tasks.push({
        id: `plan:${item.id}`,
        area: item.area || 'Plano',
        priority: parseBrazilianDate(item.when || item.due) < new Date() ? 'Crítica' : 'Alta',
        owner: item.who || item.owner || 'Gestor',
        title: `Executar plano: ${item.what || item.title}`,
        due: item.when || item.due || today,
        source: item.planType === 'weeklySeller' ? '5W2H semanal' : 'Plano de ação',
      });
    });
  dataQualityIssues(rows)
    .filter((item) => item.severity === 'red')
    .slice(0, 4)
    .forEach((item) => {
      tasks.push({
        id: `quality:${normalize(item.title)}`,
        area: 'Dados',
        priority: 'Crítica',
        owner: 'Administrador',
        title: item.title,
        due: today,
        source: `${item.count.toLocaleString('pt-BR')} ocorrência(s)`,
      });
    });
  return tasks
    .map((task) => ({ ...task, status: taskStatus(task.id), sort: taskPriorityWeight(task.priority) }))
    .sort((a, b) => b.sort - a.sort || timelineTime(a.due) - timelineTime(b.due) || collator.compare(a.title, b.title));
}

function ensureAutomaticTaskRecords(rows = state.allRows) {
  if (!state.store.taskStatus || typeof state.store.taskStatus !== 'object') state.store.taskStatus = {};
  let changed = false;
  automaticTaskItems(rows).forEach((task) => {
    if (!state.store.taskStatus[task.id]) {
      state.store.taskStatus[task.id] = { status: 'Aberta', createdAt: new Date().toISOString(), updatedAt: '' };
      changed = true;
    }
  });
  return changed;
}

function taskStatus(id) {
  return cleanText(state.store.taskStatus?.[id]?.status) || 'Aberta';
}

function updateAutomaticTaskStatus(id, status) {
  state.store.taskStatus[id] = {
    ...(state.store.taskStatus[id] || {}),
    status: cleanText(status) || 'Aberta',
    updatedAt: new Date().toISOString(),
  };
  recordAudit('Tarefa automática atualizada', `${id}: ${status}`);
  persist();
  render();
}

function taskPriorityWeight(priority) {
  return priority === 'Crítica' ? 3 : priority === 'Alta' ? 2 : 1;
}

function taskTable(tasks) {
  return `
    <table>
      <thead>
        <tr>
          <th>Prioridade</th>
          <th>Tarefa</th>
          <th>Dono</th>
          <th>Prazo</th>
          <th>Origem</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        ${tasks
          .map(
            (task) => `
              <tr>
                <td><span class="badge ${task.priority === 'Crítica' ? 'red' : 'yellow'}">${escapeHtml(task.priority)}</span></td>
                <td><strong>${escapeHtml(task.title)}</strong><span class="subtext">${escapeHtml(task.area)}</span></td>
                <td>${escapeHtml(task.owner)}</td>
                <td>${escapeHtml(formatShortDate(task.due))}</td>
                <td>${escapeHtml(task.source)}</td>
                <td>
                  <select class="status-select ${planningStatusClass(task.status)}" data-task-status="${escapeHtml(task.id)}">
                    ${['Aberta', 'Em andamento', 'Concluída', 'Bloqueada'].map((status) => `<option ${normalize(status) === normalize(task.status) ? 'selected' : ''}>${status}</option>`).join('')}
                  </select>
                </td>
              </tr>
            `,
          )
          .join('') || emptyRow('Nenhuma tarefa automática aberta.')}
      </tbody>
    </table>
  `;
}

function renderAgenda() {
  const events = state.store.schedule.slice().sort((a, b) => new Date(a.start) - new Date(b.start));
  const catalog = getCourseCatalog();
  const teachers = state.store.teachers || [];
  const currentYear = new Date().getFullYear();
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
          <input name="teacher" list="teacherOptions" placeholder="Docente" required />
          ${teacherDatalist()}
          <input name="subject" placeholder="Assunto" required />
          <select name="course" required>
            <option value="">Curso da turma</option>
            ${catalog.map((course) => `<option>${escapeHtml(course.name)}</option>`).join('')}
          </select>
          <div class="two-cols">
            <input name="cohortYear" type="number" min="2020" max="2035" value="${currentYear}" placeholder="Ano de ingresso" required />
            <select name="semester" required>
              <option value="1">1º semestre: A a F</option>
              <option value="2">2º semestre: A a F</option>
            </select>
          </div>
          <div class="two-cols">
            <input name="start" type="datetime-local" required />
            <input name="end" type="datetime-local" required />
          </div>
          <div class="two-cols">
            <select name="room" required>
              <option value="">Ambiente da aula</option>
              ${CLASSROOM_OPTIONS.map((room) => `<option>${escapeHtml(room)}</option>`).join('')}
            </select>
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
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Professores', 'Contato para aviso de aula')}</div>
          <span>${teachers.length} cadastrado(s)</span>
        </div>
        <form class="stack-form" data-form="teacher">
          <input name="name" placeholder="Nome do professor" required />
          <input name="education" placeholder="Formacao. Ex: Graduacao, pos-graduacao" />
          <input name="phone" placeholder="Telefone/WhatsApp" />
          <input name="email" type="email" placeholder="E-mail" />
          <button type="submit">Salvar professor</button>
        </form>
      </article>
      <article class="table-panel">
        <div class="panel-heading">
          <div>${smallTitle('Lista de professores', 'WhatsApp e e-mail')}</div>
          <span>${teachers.length} registros</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Professor</th><th>Formacao</th><th>Contato</th><th>Acao</th></tr></thead>
            <tbody>${teachers.map(teacherRowTemplate).join('') || emptyRow('Cadastre o primeiro professor.')}</tbody>
          </table>
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
          <input name="student" list="studentLookupOptions" placeholder="Nome, CPF ou RA do aluno" autocomplete="off" required />
          ${studentLookupDatalist()}
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
    ${moduleTitle('Fila do dia', 'Eventos em ordem de horário, da chegada até a finalização.')}
    <section class="metric-grid">
      ${metricCard('Na fila', events.length, 'Eventos de hoje')}
      ${metricCard('Próximos 30 min', events.filter((event) => event.soon).length, 'Preparação do ambiente', 'yellow')}
      ${metricCard('Em atendimento', events.filter((event) => event.status === 'presente').length, 'Aguardando finalização', 'cyan')}
      ${metricCard('Arquivados', state.store.archive.length, 'Concluídos', 'green')}
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Próximos atendimentos', 'O evento mais próximo fica no topo')}</div>
        <span>${new Date().toLocaleDateString('pt-BR')}</span>
      </div>
      <div class="queue-list">
        ${events.map(queueTemplate).join('') || '<p class="muted">Nenhum evento para hoje.</p>'}
      </div>
    </section>
  `;
}

function renderSeguranca() {
  if (!canAccessAdminModules()) {
    els.modules.seguranca.innerHTML = accessDenied('Usuários e acessos', 'Gestão de usuários e histórico são disponíveis apenas para Administrador.');
    return;
  }
  const configuredUsers = state.users.length ? state.users : defaultUsers();
  els.modules.seguranca.innerHTML = `
    ${moduleTitle('Usuários e acessos', 'Perfis de acesso e histórico das alterações do polo.')}
    <section class="metric-grid">
      ${metricCard('Persistência', state.remoteState ? 'Planilha' : 'Local', state.remoteState ? 'Google Sheets operacional' : 'Fallback neste navegador', state.remoteState ? 'green' : 'yellow')}
      ${metricCard('Dados do polo', Object.keys(state.store.overrides).length, 'Alterações locais preservadas')}
      ${metricCard('Decisões', state.store.decisions.length, 'Plano de ação gerencial', 'cyan')}
      ${metricCard('Histórico', state.store.auditTrail.length, 'Alterações importantes registradas', state.remoteState ? 'green' : 'yellow')}
    </section>
    <section class="split-grid">
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Perfis de acesso', 'O que cada pessoa pode ver')}</div>
          <span>Perfil atual: ${escapeHtml(roleLabel(state.profile))}</span>
        </div>
        <div class="role-grid">
          ${roleCard('Administrador', 'Visão completa, financeiro global e configurações.', state.profile === 'admin')}
          ${roleCard('Financeiro', 'Valores individuais, cobrança e confirmação manual de matrícula.', state.profile === 'financeiro')}
          ${roleCard('Atendimento', 'Dados cadastrais, comerciais e acadêmicos sem valores financeiros.', state.profile === 'consultor')}
        </div>
      </article>
      <article class="panel">
        <div class="panel-heading">
          <div>${smallTitle('Dados do polo', 'Atualizações que a sede não apaga')}</div>
          <span>${Object.keys(state.store.overrides).length} registros</span>
        </div>
        <div class="audit-list">
          ${auditItems().join('') || '<p class="muted">Nenhuma alteração local registrada.</p>'}
        </div>
      </article>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Usuários do sistema', 'Quem pode acessar o app')}</div>
        <span>${configuredUsers.length} usuários configurados</span>
      </div>
      <form class="inline-form user-form" data-form="user">
        <input name="usuario" placeholder="Usuario" required />
        <input name="senha" placeholder="Senha" required />
        <input name="nome" placeholder="Nome completo" required />
        <select name="perfil">
          <option value="consultor">Atendimento</option>
          <option value="financeiro">Financeiro</option>
          <option value="admin">Administrador</option>
        </select>
        <select name="ativo">
          <option>SIM</option>
          <option>NAO</option>
        </select>
        <button type="submit">Salvar usuário</button>
      </form>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Nome</th>
              <th>Usuário</th>
              <th>Perfil</th>
              <th>Último acesso</th>
              <th>Ação</th>
            </tr>
          </thead>
          <tbody>${configuredUsers.map(userRowTemplate).join('')}</tbody>
        </table>
      </div>
    </section>
    <section class="table-panel">
      <div class="panel-heading">
        <div>${smallTitle('Governança e LGPD', 'Controle de acesso, exportação e rastreabilidade')}</div>
        <div class="panel-actions">
          <button class="mini-button" type="button" data-export="governanca">Exportar governança</button>
          <span>${state.store.auditTrail.length} logs</span>
        </div>
      </div>
      <div class="table-wrap">
        ${governanceTable()}
      </div>
    </section>
  `;
}

function studentRowTemplate(row) {
  const enrollment = enrollmentSettlementForStudent(row);
  return `
    <tr>
      <td>
        <div class="student-name-line">
          <button class="text-button" type="button" data-open-student="${escapeHtml(row.key)}">
            ${escapeHtml(row.name)}
            <span>${escapeHtml(shortUnit(row.unit))}${row.contactOverridden ? ' · contato local' : ''}${row.override.followStatus ? ` · ${escapeHtml(row.override.followStatus)}` : ''}</span>
          </button>
          ${normalizePhone(row.phone) ? `<a class="whatsapp-inline" href="${whatsAppUrl(row)}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp</a>` : ''}
        </div>
        ${isAcompanhamentoReopeningEnrollment(row) ? '<span class="badge cyan">Reabertura do curso</span>' : ''}
        ${enrollment.newEnrollment ? `<span class="badge ${enrollment.requiresAction ? 'yellow' : 'green'}">${enrollment.requiresAction ? 'Matrícula nova pendente' : 'Matrícula nova confirmada'}</span>` : ''}
      </td>
      <td>${escapeHtml(maskCpf(row.cpf))}</td>
      <td>${escapeHtml(row.ra)}</td>
      <td>${escapeHtml(row.course)}</td>
      <td><span class="badge cyan">${escapeHtml(humanizeStartPeriod(row.startPeriod))}</span></td>
      <td><span class="score-pill ${riskScoreClass(row.risk.score)}">${row.risk.score}/100</span></td>
      <td>
        <select class="status-select" data-status-key="${escapeHtml(row.key)}">
          ${STATUS_OPTIONS.map((status) => `<option ${normalize(status) === normalize(row.localStatus) ? 'selected' : ''}>${status}</option>`).join('')}
        </select>
        ${row.statusOverridden ? '<span class="badge yellow">Local</span>' : '<span class="badge">Sede</span>'}
      </td>
      <td>
        <button class="mini-button" type="button" data-open-student="${escapeHtml(row.key)}">Editar</button>
      </td>
    </tr>
  `;
}

function sourceEnrollmentRowTemplate(row) {
  const enrollment = enrollmentSettlementForStudent(row);
  return `
    <tr>
      <td>
        <button class="text-button" type="button" data-open-student="${escapeHtml(row.key)}">
          ${escapeHtml(row.name)}
          <span>RA ${escapeHtml(row.ra || '-')} - base Acompanhamento</span>
        </button>
      </td>
      <td>${escapeHtml(row.course)}</td>
      <td><span class="badge ${enrollment.boletoOk ? 'green' : 'yellow'}">${escapeHtml(enrollment.boletoLabel)}</span></td>
      <td><span class="badge ${enrollment.paymentOk ? 'green' : 'yellow'}">${escapeHtml(enrollment.paymentLabel)}</span></td>
      <td>
        <button class="mini-button" type="button" data-open-student="${escapeHtml(row.key)}">Regularizar</button>
        ${
          normalizePhone(row.phone)
            ? `<a class="mini-link" href="${boletoWhatsAppUrl(row)}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp boleto</a>`
            : ''
        }
      </td>
    </tr>
  `;
}

function courseRowTemplate(course) {
  const discounts = COURSE_DISCOUNT_RATES.map((rate) => {
    const value = courseDiscountValue(course, rate);
    return `
      <label>
        ${rate}%
        <input type="number" min="0" step="0.01" value="${Number(value || 0)}" data-course-field="discount${rate}" data-course-key="${escapeHtml(course.key)}" />
      </label>
    `;
  }).join('');
  return `
    <tr>
      <td>
        <strong>${escapeHtml(course.name)}</strong>
        <span class="subtext">${course.local ? 'Dados do polo' : 'Vindo da base de alunos'}</span>
      </td>
      <td><input class="table-input" value="${escapeHtml(course.modality)}" data-course-field="modality" data-course-key="${escapeHtml(course.key)}" /></td>
      <td><input class="table-input" value="${escapeHtml(course.habilitation)}" data-course-field="habilitation" data-course-key="${escapeHtml(course.key)}" /></td>
      <td><input class="table-input" value="${escapeHtml(course.duration)}" data-course-field="duration" data-course-key="${escapeHtml(course.key)}" /></td>
      <td><input class="table-input money-input" type="number" min="0" step="0.01" value="${Number(course.monthlyFee || 0)}" data-course-field="monthlyFee" data-course-key="${escapeHtml(course.key)}" /></td>
      <td><div class="course-discount-grid compact">${discounts}</div></td>
      <td><textarea class="table-input" rows="2" data-course-field="authorization" data-course-key="${escapeHtml(course.key)}">${escapeHtml(course.authorization)}</textarea></td>
      <td>
        <span class="badge ${course.local ? 'cyan' : 'yellow'}">${course.local ? 'Editavel' : 'Criar override'}</span>
        ${course.local ? `<button type="button" data-delete-course="${escapeHtml(course.key)}">Excluir</button>` : ''}
      </td>
    </tr>
  `;
}

function financialRowTemplate(row) {
  const matchedStudent = row.sourceFinance && row.matchedKey ? rowByKey(row.matchedKey) : null;
  const contactRow = matchedStudent || row;
  if (row.sourceFinance) {
    return `
      <tr>
        <td>
          <button class="text-button" type="button" data-open-financial="${escapeHtml(row.key)}">
            ${escapeHtml(row.name)}
            <span>RA ${escapeHtml(row.ra || '-')} - Faturado ${formatMoney(row.billed)} / recebido ${formatMoney(row.received)}</span>
          </button>
        </td>
        <td>${escapeHtml(row.course)}</td>
        <td><span class="badge ${row.isDebt ? 'red' : 'green'}">${row.isDebt ? 'Em atraso' : 'Em dia'}</span></td>
        <td><strong>${formatMoney(row.debtValue)}</strong></td>
        <td>${escapeHtml(row.overdueMonths)}</td>
        <td>
          <span class="badge cyan">Sede</span>
          <span class="subtext">Faturado e recebido</span>
          ${
            row.isDebt && normalizePhone(contactRow.phone)
              ? `<a class="mini-link" href="${overdueWhatsAppUrl({ ...row, phone: contactRow.phone })}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp cobrança</a>`
              : ''
          }
        </td>
      </tr>
    `;
  }
  const override = state.store.overrides[row.key] || {};
  const enrollment = enrollmentSettlementForStudent(row);
  return `
    <tr>
      <td>
        <button class="text-button" type="button" data-open-financial="${escapeHtml(row.key)}">
          ${escapeHtml(row.name)}
          <span>RA ${escapeHtml(row.ra)} - histórico de pagamentos</span>
        </button>
      </td>
      <td>${escapeHtml(row.course)}</td>
      <td><span class="badge ${statusClass(row.localStatus)}">${escapeHtml(row.localStatus)}</span></td>
      <td><strong>${formatMoney(row.debtValue)}</strong></td>
      <td>${escapeHtml(row.overdueMonths)}</td>
      <td>
        <span class="badge ${enrollment.paymentOk ? 'green' : 'yellow'}">
          ${escapeHtml(enrollment.paymentLabel)}
        </span>
        <span class="subtext">${escapeHtml(enrollment.boletoLabel)}</span>
        ${
          enrollment.legacySettled
            ? `<span class="subtext">${escapeHtml(enrollment.detail)}</span>`
            : `
              <label class="check-line">
                <input type="checkbox" data-enrollment-paid="${escapeHtml(row.key)}" ${enrollment.paymentLabel === 'Pago' ? 'checked' : ''} />
                Pagamento confirmado
              </label>
            `
        }
        ${
          row.isDebt && normalizePhone(row.phone)
            ? `<a class="mini-link" href="${overdueWhatsAppUrl(row)}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp cobrança</a>`
            : ''
        }
      </td>
    </tr>
  `;
}

function repasseRowTemplate(row) {
  return `
    <tr>
      <td><span class="badge cyan">${escapeHtml(financialPeriodLabel(row.periodKey || row.competence))}</span></td>
      <td>${escapeHtml(row.date || row.paymentDate || '-')}</td>
      <td>${escapeHtml(row.description || row.studentName || 'Repasse da sede')}</td>
      <td>${escapeHtml(row.course || '-')}</td>
      <td><strong>${formatMoney(row.amount)}</strong></td>
      <td>${escapeHtml(row.importFile || '-')}</td>
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
      <td>
        <button class="mini-button" type="button" data-open-retention="${escapeHtml(row.key)}">Registrar contato</button>
        ${normalizePhone(row.phone) ? `<a class="mini-link" href="${whatsAppUrl(row)}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp</a>` : ''}
      </td>
    </tr>
  `;
}

function showDrawer(row) {
  state.selectedKey = row.key;
  const override = state.store.overrides[row.key] || {};
  const enrollment = enrollmentSettlementForStudent(row);
  els.drawerContent.className = 'drawer-content';
  els.drawerContent.innerHTML = `
    <p class="eyebrow">Visão 360°</p>
    <h2>${escapeHtml(row.name)}</h2>
    <div class="badge-row">
      <span class="badge ${statusClass(row.localStatus)}">${escapeHtml(row.localStatus)}</span>
      <span class="badge ${row.risk.level === 'Alto' ? 'red' : row.risk.level === 'Médio' ? 'yellow' : 'green'}">Atenção ${escapeHtml(row.risk.level)}</span>
      <span class="badge ${riskScoreClass(row.risk.score)}">${row.risk.score}/100</span>
      ${row.statusOverridden ? '<span class="badge yellow">Dados do polo</span>' : '<span class="badge">Sede</span>'}
    </div>
    <section class="drawer-section health-score-card">
      <div>
        <span>Nível de atenção</span>
        <strong>${row.risk.score}/100</strong>
        <small>${escapeHtml(riskReason(row))}</small>
      </div>
      <div class="score-meter" style="--score:${row.risk.score}"></div>
    </section>
    ${checklistTemplate(enrollmentChecklist(row))}
    <form class="stack-form" data-form="student-override" data-student-key="${escapeHtml(row.key)}">
      <label>Status do polo
        <select name="status">
          ${STATUS_OPTIONS.map((status) => `<option ${normalize(status) === normalize(row.localStatus) ? 'selected' : ''}>${status}</option>`).join('')}
        </select>
      </label>
      <div class="two-cols">
        <label>Telefone/WhatsApp do polo
          <input name="contactPhone" value="${escapeHtml(row.phone)}" placeholder="Telefone atualizado pelo polo" />
        </label>
        <label>E-mail do polo
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
      ${canSeeFinancial() && enrollment.legacySettled ? `
        <div class="decision-signal success">
          <strong>Base Acompanhamento regularizada</strong>
          <span>${escapeHtml(enrollment.detail)}</span>
        </div>
      ` : ''}
      ${canSeeFinancial() && enrollment.newEnrollment ? `
        <div class="decision-signal warning">
          <strong>Matrícula nova da sede</strong>
          <span>${escapeHtml(enrollment.detail)}${enrollment.lead ? ' Dados financeiros vinculados ao CRM local quando houver correspondência.' : ''}</span>
        </div>
      ` : ''}
      ${canSeeFinancial() && !enrollment.legacySettled ? `
        <label>Taxa de matrícula paga?
          <select name="enrollmentPaid">
            <option value="false" ${enrollment.paymentLabel === 'Pago' ? '' : 'selected'}>Não</option>
            <option value="true" ${enrollment.paymentLabel === 'Pago' ? 'selected' : ''}>Sim</option>
          </select>
        </label>
        <label>Aluno isento da taxa?
          <select name="enrollmentExempt">
            <option value="false" ${enrollment.paymentLabel === 'Isento' ? '' : 'selected'}>Não</option>
            <option value="true" ${enrollment.paymentLabel === 'Isento' ? 'selected' : ''}>Sim</option>
          </select>
        </label>
        <label>Boleto enviado?
          <select name="boletoSent">
            <option value="false" ${enrollment.boletoOk && enrollment.paymentLabel !== 'Isento' ? '' : 'selected'}>Não</option>
            <option value="true" ${enrollment.boletoOk && enrollment.paymentLabel !== 'Isento' ? 'selected' : ''}>Sim</option>
          </select>
        </label>
        <label>Valor combinado da taxa
          <input name="enrollmentFee" type="number" min="0" step="0.01" value="${Number(override.enrollmentFee || enrollment.lead?.enrollmentFee || state.store.settings.enrollmentFee || 99)}" />
        </label>
      ` : ''}
      <button type="submit">Salvar dados do polo</button>
    </form>
    <div class="detail-list">
      ${detailRow('CPF', maskCpf(row.cpf))}
      ${detailRow('RA', row.ra)}
      ${detailRow('Curso', row.course)}
      ${detailRow('Período início', humanizeStartPeriod(row.startPeriod))}
      ${detailRow('Status na sede', row.sourceStatus)}
      ${detailRow('One', `${row.oneAccess || '-'} ${row.oneDays || ''}`)}
      ${detailRow('AVA', `${row.avaAccess || '-'} ${row.avaDays || ''}`)}
      ${detailRow('Alerta AVA', `${row.avaAlert.label} · ${row.avaDaysNumber >= 0 ? `${row.avaDaysNumber} dias` : 'sem dado'}`)}
      ${canSeeFinancial() ? detailRow('Pagamento', `${row.isDebt ? 'Em atraso' : 'Em dia'} · ${formatMoney(row.debtValue)}`) : detailRow('Pagamento', 'Disponível apenas para Administrador')}
      ${canSeeFinancial() ? detailRow('Taxa do polo', enrollment.sourceBase ? enrollment.detail : `${override.enrollmentExempt ? 'Isento' : override.enrollmentPaid ? 'Pago' : 'Pendente'} · boleto ${override.boletoSent ? 'enviado' : 'não enviado'}`) : ''}
      ${detailRow('Telefone', row.phone)}
      ${detailRow('E-mail', row.email)}
      ${detailRow('Situação do acompanhamento', override.followStatus || 'Sem acompanhamento')}
      ${detailRow('Endereço', row.address)}
    </div>
    ${timelineTemplate(studentTimeline(row))}
  `;
  els.drawer.classList.add('open');
}

function showRetentionDrawer(row) {
  state.selectedKey = row.key;
  const retention = row.retention || {};
  els.drawerContent.className = 'drawer-content';
  els.drawerContent.innerHTML = `
    <p class="eyebrow">Acompanhamento AVA</p>
    <h2>${escapeHtml(row.name)}</h2>
    <div class="badge-row">
      <span class="badge ${row.avaAlert.className}">${escapeHtml(row.avaAlert.label)}</span>
      <span class="badge cyan">${row.avaDaysNumber >= 0 ? `${row.avaDaysNumber} dias sem acesso` : 'sem dado de dias'}</span>
      <span class="badge ${retention.contacted ? 'green' : 'yellow'}">${retention.contacted ? 'Contato registrado' : 'Contato pendente'}</span>
    </div>
    ${normalizePhone(row.phone) ? `<a class="solid-button whatsapp-button" href="${whatsAppUrl(row)}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>Abrir WhatsApp</a>` : ''}
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
      <label>Responsável pelo acompanhamento
        <input name="responsible" value="${escapeHtml(retention.responsible || roleLabel(state.profile))}" />
      </label>
      <label>Observação
        <textarea name="note" rows="4" placeholder="Explique o motivo e o encaminhamento...">${escapeHtml(retention.note || '')}</textarea>
      </label>
      <button type="submit">Salvar contato</button>
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

function showFinancialDrawer(row) {
  if (row.sourceFinance) {
    const matchedStudent = row.matchedKey ? rowByKey(row.matchedKey) : null;
    const enrollment = matchedStudent ? enrollmentSettlementForStudent(matchedStudent) : null;
    els.drawerContent.className = 'drawer-content';
    els.drawerContent.innerHTML = `
      <p class="eyebrow">Histórico de pagamentos</p>
      <h2>${escapeHtml(row.name)}</h2>
      <div class="badge-row">
        <span class="badge ${row.isDebt ? 'red' : 'green'}">${row.isDebt ? 'Em atraso' : 'Em dia'}</span>
        <span class="badge yellow">${escapeHtml(row.overdueMonths || 'Sem mes em aberto')}</span>
        <span class="badge cyan">${formatMoney(row.debtValue)}</span>
      </div>
      <div class="detail-list">
        ${detailRow('CPF', maskCpf(row.cpf))}
        ${detailRow('RA', row.ra)}
        ${detailRow('Curso', row.course)}
        ${detailRow('Fonte das mensalidades', 'Faturamento e recebimento')}
        ${detailRow('Total faturado', formatMoney(row.billed))}
        ${detailRow('Total recebido', formatMoney(row.received))}
        ${row.interestOrLatePayment ? detailRow('Recebido acima do faturado', `${formatMoney(row.interestOrLatePayment)} (possivel juros/regularizacao posterior)`) : ''}
        ${detailRow('Saldo em atraso', formatMoney(row.debtValue))}
        ${detailRow('Meses em aberto', row.overdueMonths)}
        ${enrollment ? detailRow('Taxa de matrícula', enrollment.sourceBase ? enrollment.detail : enrollment.paymentLabel) : ''}
        ${detailRow('Registros de faturamento', row.billingRecords?.length || 0)}
        ${detailRow('Registros de recebimento', row.receiptRecords?.length || 0)}
      </div>
    `;
    els.drawer.classList.add('open');
    return;
  }
  const override = state.store.overrides[row.key] || {};
  const enrollment = enrollmentSettlementForStudent(row);
  els.drawerContent.className = 'drawer-content';
  els.drawerContent.innerHTML = `
    <p class="eyebrow">Histórico de pagamentos</p>
    <h2>${escapeHtml(row.name)}</h2>
    <div class="badge-row">
      <span class="badge ${row.isDebt ? 'red' : 'green'}">${row.isDebt ? 'Em atraso' : 'Em dia'}</span>
      <span class="badge yellow">${escapeHtml(row.overdueMonths || 'Meses nao informados')}</span>
      <span class="badge cyan">${formatMoney(row.debtValue)}</span>
    </div>
    <div class="detail-list">
      ${detailRow('CPF', maskCpf(row.cpf))}
      ${detailRow('RA', row.ra)}
      ${detailRow('Curso', row.course)}
      ${detailRow('Fonte das mensalidades', 'Faturamento/Recebimento da sede quando configurado')}
      ${detailRow('Valor total em atraso', formatMoney(row.debtValue))}
      ${detailRow('Meses em aberto', row.overdueMonths)}
      ${detailRow('Taxa de matrícula', enrollment.sourceBase ? enrollment.feeLabel : formatMoney(override.enrollmentFee || state.store.settings.enrollmentFee || 99))}
      ${detailRow('Boleto da matrícula', enrollment.boletoLabel)}
      ${detailRow('Pagamento da taxa', enrollment.paymentLabel)}
      ${detailRow('Última alteração do polo', override.updatedAt ? new Date(override.updatedAt).toLocaleString('pt-BR') : '-')}
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
  if (formType === 'service-ticket') saveServiceTicket(data);
  if (formType === 'course') addCourse(data);
  if (formType === 'teacher') saveTeacher(data);
  if (formType === 'change-password') changeCurrentPassword(data);
  if (formType === 'lead') addLead(data);
  if (formType === 'user') saveConfiguredUser(data);
  if (formType === 'goal') updateSettings({ monthlyTarget: Number(data.monthlyTarget || 0) });
  if (formType === 'bi-filter') updateSettings({ biMonth: Number(data.biMonth || 0), biYear: Number(data.biYear || new Date().getFullYear()) });
  if (formType === 'bi-goals') updateSettings({
    monthlyTarget: Number(data.monthlyTarget || 0),
    annualTarget: Number(data.annualTarget || 0),
  });
  if (formType === 'bi-ticket') updateSettings({ monthlyTicket: Number(data.monthlyTicket || 0) });
  if (formType === 'decision') addDecision(data);
  if (formType === 'planning-5w2h') addPlanningAction(data);
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
  const whatsappLink = event.target.closest('[data-whatsapp-link]');
  if (whatsappLink) {
    event.preventDefault();
    openWhatsAppWindow(whatsappLink.href);
    return;
  }

  if (event.target.closest('[data-mega-close]')) {
    closeMegaMenu();
    return;
  }

  const dashboardFocus = event.target.closest('[data-dashboard-focus]');
  if (dashboardFocus) {
    state.dashboardFocus = dashboardFocus.dataset.dashboardFocus || '';
    activateModule('inteligencia');
    closeMegaMenu();
    return;
  }

  const snapshotButton = event.target.closest('[data-refresh-snapshot]');
  if (snapshotButton) {
    upsertCurrentSnapshot(true);
    persist();
    toast('Snapshot mensal atualizado para comparação histórica.');
    render();
    return;
  }

  const resetBiFilter = event.target.closest('[data-reset-bi-filter]');
  if (resetBiFilter) {
    updateSettings({ biMonth: 0, biYear: 0 });
    persist();
    render();
    toast('Filtro financeiro voltado para o histórico completo.');
    return;
  }

  const adminModule = event.target.closest('[data-admin-module]');
  if (adminModule) {
    const module = adminModule.dataset.adminModule;
    if (isRestrictedModule(module)) {
      toast(restrictedMessage(module));
      return;
    }
    const menu = adminModule.closest('[data-admin-menu]');
    if (menu) menu.removeAttribute('open');
    activateModule(module);
    closeMegaMenu();
    return;
  }

  const modalClose = event.target.closest('[data-modal-close]');
  if (modalClose) {
    closeModal();
    return;
  }

  const modalBoleto = event.target.closest('[data-modal-boleto]');
  if (modalBoleto) {
    updateLeadFinance(modalBoleto.dataset.modalBoleto, { boletoSent: true });
    closeModal();
    return;
  }

  const modalPaid = event.target.closest('[data-modal-paid]');
  if (modalPaid) {
    updateLeadFinance(modalPaid.dataset.modalPaid, { boletoSent: true, enrollmentPaymentStatus: 'Pago' });
    closeModal();
    return;
  }

  const importConfirm = event.target.closest('[data-import-confirm]');
  if (importConfirm) {
    confirmPendingImport();
    return;
  }

  const undoImport = event.target.closest('[data-undo-import]');
  if (undoImport) {
    undoLastImport();
    return;
  }

  const refreshOperational = event.target.closest('[data-refresh-operational]');
  if (refreshOperational) {
    refreshOperationalData();
    return;
  }

  const reportButton = event.target.closest('[data-report]');
  if (reportButton) {
    openExecutiveReport(reportButton.dataset.report);
    return;
  }

  const openMenu = event.target.closest('[data-open-mega-menu]');
  if (openMenu) {
    openMegaMenu();
    return;
  }

  const quickModule = event.target.closest('[data-quick-module]');
  if (quickModule) {
    if (isRestrictedModule(quickModule.dataset.quickModule)) {
      toast(restrictedMessage(quickModule.dataset.quickModule));
      return;
    }
    if (quickModule.dataset.quickModule === 'matriculas') state.moduleView = quickModule.dataset.view || 'crm';
    activateModule(quickModule.dataset.quickModule);
    closeMegaMenu();
    return;
  }

  const openButton = event.target.closest('[data-open-student]');
  if (openButton) {
    const row = state.allRows.find((item) => item.key === openButton.dataset.openStudent);
    if (row) showDrawer(row);
    return;
  }

  const financialButton = event.target.closest('[data-open-financial]');
  if (financialButton) {
    if (!canSeeFinancial()) {
      toast('Histórico financeiro disponível apenas para Administrador e Financeiro.');
      return;
    }
    const row =
      buildFinancialPosition(state.filteredRows).rows.find((item) => item.key === financialButton.dataset.openFinancial) ||
      state.allRows.find((item) => item.key === financialButton.dataset.openFinancial);
    if (row) showFinancialDrawer(row);
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
    state.store.courses = state.store.courses.filter((course) => course.key !== deleteCourse.dataset.deleteCourse && course.name !== deleteCourse.dataset.deleteCourse);
    recordAudit('Curso excluido', deleteCourse.dataset.deleteCourse);
    persist();
    render();
    return;
  }

  const deleteUser = event.target.closest('[data-delete-user]');
  if (deleteUser) {
    deleteConfiguredUser(deleteUser.dataset.deleteUser);
    return;
  }

  const deleteTeacher = event.target.closest('[data-delete-teacher]');
  if (deleteTeacher) {
    deleteConfiguredTeacher(deleteTeacher.dataset.deleteTeacher);
    return;
  }

  const exportButton = event.target.closest('[data-export]');
  if (exportButton) {
    exportDataset(exportButton.dataset.export);
    return;
  }

  const convertLead = event.target.closest('[data-convert-lead]');
  if (convertLead) {
    updateLeadStage(convertLead.dataset.convertLead, 'Matriculado');
    return;
  }

  const leadMove = event.target.closest('[data-lead-move]');
  if (leadMove) {
    updateLeadStage(leadMove.dataset.leadMove, leadMove.dataset.stage);
    return;
  }

  const queueButton = event.target.closest('[data-queue-action]');
  if (queueButton) {
    queueAction(queueButton.dataset.queueAction, queueButton.dataset.queueId, queueButton.dataset.queueType);
  }
}

function handleKeydown(event) {
  if (event.key === 'Escape' && !els.megaMenu?.hidden) {
    closeMegaMenu();
  }
}

function handleChange(event) {
  const courseImport = event.target.closest('[data-course-import]');
  if (courseImport) {
    const file = courseImport.files?.[0];
    if (!file) return;
    file.text().then((text) => importCourseCsv(text, file.name));
    courseImport.value = '';
    return;
  }

  const financialImport = event.target.closest('[data-financial-import]');
  if (financialImport) {
    const file = financialImport.files?.[0];
    if (!file) return;
    file.text().then((text) => startGuidedImport(financialImport.dataset.financialImport, text, file.name));
    financialImport.value = '';
    return;
  }

  const courseField = event.target.closest('[data-course-field]');
  if (courseField) {
    updateCourseField(courseField.dataset.courseKey, courseField.dataset.courseField, courseField.value);
    return;
  }

  const statusSelect = event.target.closest('[data-status-key]');
  if (statusSelect) {
    saveStudentOverride(statusSelect.dataset.statusKey, { status: statusSelect.value });
    return;
  }

  const serviceStatus = event.target.closest('[data-service-status]');
  if (serviceStatus) {
    updateServiceTicketField(serviceStatus.dataset.serviceStatus, 'status', serviceStatus.value);
    return;
  }

  const serviceResponse = event.target.closest('[data-service-response]');
  if (serviceResponse) {
    updateServiceTicketField(serviceResponse.dataset.serviceResponse, 'response', serviceResponse.value);
    return;
  }

  const paidCheck = event.target.closest('[data-enrollment-paid]');
  if (paidCheck) {
    if (!canSeeFinancial()) {
      paidCheck.checked = !paidCheck.checked;
      toast('Confirmação de pagamento disponível apenas para Administrador e Financeiro.');
      return;
    }
    saveStudentOverride(paidCheck.dataset.enrollmentPaid, { enrollmentPaid: paidCheck.checked ? 'true' : 'false' });
    return;
  }

  const boletoCheck = event.target.closest('[data-lead-boleto]');
  if (boletoCheck) {
    if (!canSeeFinancial()) {
      boletoCheck.checked = !boletoCheck.checked;
      toast('Controle de boleto disponível apenas para Administrador e Financeiro.');
      return;
    }
    updateLeadFinance(boletoCheck.dataset.leadBoleto, { boletoSent: boletoCheck.checked });
    return;
  }

  const leadPayment = event.target.closest('[data-lead-payment]');
  if (leadPayment) {
    if (!canSeeFinancial()) {
      toast('Confirmação de pagamento disponível apenas para Administrador e Financeiro.');
      render();
      return;
    }
    updateLeadFinance(leadPayment.dataset.leadPayment, { enrollmentPaymentStatus: leadPayment.value });
  }

  const leadFee = event.target.closest('[data-lead-fee]');
  if (leadFee) {
    updateLeadFinance(leadFee.dataset.leadFee, { enrollmentFee: Number(leadFee.value || 0) });
  }

  const planStatus = event.target.closest('[data-plan-status]');
  if (planStatus) {
    updatePlanningStatus(planStatus.dataset.planStatus, planStatus.value);
  }

  const taskStatus = event.target.closest('[data-task-status]');
  if (taskStatus) {
    updateAutomaticTaskStatus(taskStatus.dataset.taskStatus, taskStatus.value);
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
  const row = rowByKey(key);
  syncLeadFinanceFromStudentOverride(row, data);
  recordAudit('Dados do aluno atualizados', `${row?.name || key}: contato, status ou financeiro atualizado`);
  persist();
  rerichRows();
  toast('Dados do polo salvos e preservados nas sincronizações.');
  render();
}

function syncLeadFinanceFromStudentOverride(row, data) {
  if (!row) return;
  const lead = leadForStudent(row);
  if (!lead || data.enrollmentPaid === undefined) return;
  if (isAcompanhamentoNewEnrollment(row)) {
    lead.stage = 'Matriculado';
    lead.matriculatedAt = lead.matriculatedAt || new Date().toISOString();
  }
  if (data.boletoSent !== undefined) {
    lead.boletoSent = data.boletoSent === 'true' || data.boletoSent === true;
    lead.boletoSentAt = lead.boletoSent ? lead.boletoSentAt || new Date().toISOString() : '';
  }
  if (data.enrollmentExempt === 'true' || data.enrollmentExempt === true) {
    lead.enrollmentPaymentStatus = 'Isento';
  } else if (data.enrollmentPaid === 'true' || data.enrollmentPaid === true) {
    lead.enrollmentPaymentStatus = 'Pago';
  } else {
    lead.enrollmentPaymentStatus = 'Pendente';
  }
  if (data.enrollmentFee !== undefined) lead.enrollmentFee = Number(data.enrollmentFee || lead.enrollmentFee || state.store.settings.enrollmentFee || 99);
  lead.updatedAt = new Date().toISOString();
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
  recordAudit('Contato de acompanhamento', `${rowByKey(key)?.name || key}: ${data.reason || 'sem motivo informado'}`);
  persist();
  rerichRows();
  toast('Contato registrado.');
  render();
}

function saveServiceTicket(data) {
  const student = resolveStudentSelection(data.student) || rowByKey(data.student);
  if (!student) {
    toast('Selecione o aluno correto por nome, CPF ou RA.');
    return;
  }
  const ticket = normalizeServiceTicket({
    ...data,
    studentKey: student.key,
    studentName: student.name,
    cpf: student.cpf,
    ra: student.ra,
    course: student.course,
  });
  state.store.serviceTickets.push(ticket);
  recordAudit('Atendimento registrado', `${ticket.studentName}: ${ticket.problem}`);
  persist();
  render();
  toast('Solicitação registrada para acompanhamento.');
}

function updateServiceTicketField(id, field, value) {
  const ticket = state.store.serviceTickets.find((item) => item.id === id);
  if (!ticket) return;
  ticket[field] = cleanText(value);
  ticket.updatedAt = new Date().toISOString();
  recordAudit('Atendimento atualizado', `${ticket.studentName}: ${field}`);
  persist();
  renderAtendimento();
  toast('Atendimento atualizado.');
}

function normalizeFinancialCollection(records, kind) {
  if (!Array.isArray(records)) return [];
  return records
    .map((item, index) => normalizeStoredFinancialRecord(item, kind, index))
    .filter((record) => isValidFinancialRecord(record, kind));
}

function appendFinancialCsv(kind, csvText, fileName) {
  if (!canSeeFinancial()) {
    toast('Importação financeira disponível apenas para Administrador e Financeiro.');
    return;
  }
  const collectionMap = { billing: 'billing', receipts: 'receipts', repasses: 'repasses' };
  const labelMap = { billing: 'Faturamento', receipts: 'Recebimento', repasses: 'Valores repassados' };
  const collectionName = collectionMap[kind] || 'billing';
  const label = labelMap[collectionName] || 'Faturamento';
  const parsed = parseCsv(csvText);
  const existingIds = new Set(state.store[collectionName].map((item) => item.id));
  const importedAt = new Date().toISOString();
  const records = parsed.rows
    .map((row, index) => normalizeFinancialRecord(row, collectionName, index, fileName, importedAt))
    .filter((record) => isValidFinancialRecord(record, collectionName));
  const fresh = records.filter((record) => !existingIds.has(record.id));

  state.store[collectionName].push(...fresh);
  state.store.importHistory.push({
    id: cryptoId(),
    kind: collectionName,
    label,
    fileName,
    importedAt,
    recordIds: fresh.map((record) => record.id),
  });
  state.store.importHistory = state.store.importHistory.slice(-20);
  recordAudit(`Importação ${label}`, `${fresh.length} novas linhas de ${fileName}`);
  persist();
  render();
  toast(`${label}: ${fresh.length} novas linhas adicionadas ao histórico.`);
}

function startGuidedImport(kind, csvText, fileName) {
  const review = buildImportReview(kind, csvText, fileName);
  state.pendingImport = { kind, csvText, fileName, review };
  showImportReviewModal(review);
}

function buildImportReview(kind, csvText, fileName) {
  const parsed = parseCsv(csvText);
  if (kind === 'students') {
    const currentKeys = new Set(state.allRows.map((row) => row.key));
    const newRows = parsed.rows.map((row, index) => ({ row, key: rowKey(row, index) }));
    const keys = newRows.map((item) => item.key);
    const duplicateKeys = new Set(keys.filter((key, index) => key && keys.indexOf(key) !== index));
    const changed = newRows.filter((item) => currentKeys.has(item.key)).length;
    const localConflicts = newRows.filter((item) => state.store.overrides[item.key]).length;
    const missingCpf = parsed.rows.filter((row) => !cleanText(row[fields.cpf])).length;
    const missingCourse = parsed.rows.filter((row) => !cleanText(row[fields.course])).length;
    return {
      kind,
      title: 'Importação guiada do Acompanhamento',
      fileName,
      total: parsed.rows.length,
      fresh: Math.max(0, parsed.rows.length - changed),
      changed,
      duplicates: duplicateKeys.size,
      conflicts: localConflicts,
      errors: missingCpf + missingCourse,
      warnings: [
        localConflicts ? `${localConflicts} aluno(s) possuem dados soberanos do polo e serão preservados.` : '',
        missingCpf ? `${missingCpf} linha(s) sem CPF.` : '',
        missingCourse ? `${missingCourse} linha(s) sem curso.` : '',
      ].filter(Boolean),
      sample: parsed.rows.slice(0, 5).map((row) => ({
        name: cleanText(row[fields.name]) || 'Sem nome',
        course: cleanText(row[fields.course]) || 'Sem curso',
        status: cleanText(row[fields.status]) || 'Sem status',
      })),
    };
  }
  const collectionMap = { billing: 'billing', receipts: 'receipts', repasses: 'repasses' };
  const labelMap = { billing: 'Faturamento', receipts: 'Recebimento', repasses: 'Valores repassados' };
  const collectionName = collectionMap[kind] || 'billing';
  const existingIds = new Set(state.store[collectionName].map((item) => item.id));
  const importedAt = new Date().toISOString();
  const records = parsed.rows
    .map((row, index) => normalizeFinancialRecord(row, collectionName, index, fileName, importedAt))
    .filter((record) => isValidFinancialRecord(record, collectionName));
  const fresh = records.filter((record) => !existingIds.has(record.id));
  const unmatched = fresh.filter((record) => collectionName !== 'repasses' && !findStudentForFinancial(record)).length;
  const zeroAmount = fresh.filter((record) => !Number(record.amount || 0)).length;
  return {
    kind,
    title: `Importação guiada de ${labelMap[collectionName]}`,
    fileName,
    total: records.length,
    fresh: fresh.length,
    changed: records.length - fresh.length,
    duplicates: records.length - fresh.length,
    conflicts: unmatched,
    errors: zeroAmount,
    warnings: [
      unmatched ? `${unmatched} linha(s) não foram vinculadas a aluno por CPF, RA ou nome.` : '',
      zeroAmount ? `${zeroAmount} linha(s) com valor zerado.` : '',
      records.length - fresh.length ? `${records.length - fresh.length} linha(s) já existiam e serão ignoradas.` : '',
    ].filter(Boolean),
    records: fresh,
    sample: fresh.slice(0, 5).map((record) => ({
      name: record.studentName || record.description || 'Sem identificação',
      course: record.course || record.competence || '-',
      status: formatMoney(record.amount),
    })),
  };
}

function showImportReviewModal(review) {
  if (!els.modalOverlay) return;
  els.modalOverlay.hidden = false;
  els.modalOverlay.innerHTML = `
    <article class="modal-card wide-modal" role="dialog" aria-modal="true" aria-label="Importação guiada">
      <button class="icon-button modal-close" type="button" data-modal-close aria-label="Fechar">x</button>
      <p class="eyebrow">Validação antes de salvar</p>
      <h2>${escapeHtml(review.title)}</h2>
      <p>Arquivo: ${escapeHtml(review.fileName)}. Confira os números antes de gravar no sistema.</p>
      <section class="metric-grid compact-metrics">
        ${metricCard('Linhas lidas', review.total, 'Total do arquivo', 'cyan')}
        ${metricCard('Novas', review.fresh, 'Entrarão no sistema', 'green')}
        ${metricCard('Já existentes', review.changed, 'Sem duplicar', review.changed ? 'yellow' : 'green')}
        ${metricCard('Alertas', review.conflicts + review.errors + review.duplicates, 'Conferir antes', review.conflicts + review.errors + review.duplicates ? 'yellow' : 'green')}
      </section>
      <div class="decision-list">
        ${review.warnings.map((warning) => `<div class="decision-signal warning"><strong>Atenção</strong><span>${escapeHtml(warning)}</span></div>`).join('') || '<div class="decision-signal success"><strong>Validação limpa</strong><span>Nenhum alerta crítico encontrado.</span></div>'}
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Nome/Descrição</th><th>Curso/Competência</th><th>Status/Valor</th></tr></thead>
          <tbody>
            ${review.sample.map((item) => `<tr><td>${escapeHtml(item.name)}</td><td>${escapeHtml(item.course)}</td><td>${escapeHtml(item.status)}</td></tr>`).join('') || emptyRow('Sem linhas válidas para visualizar.')}
          </tbody>
        </table>
      </div>
      <div class="modal-actions">
        <button class="ghost-button" type="button" data-modal-close>Cancelar</button>
        <button class="solid-button" type="button" data-import-confirm>Confirmar importação</button>
      </div>
    </article>
  `;
}

function confirmPendingImport() {
  const pending = state.pendingImport;
  if (!pending) return;
  if (pending.kind === 'students') {
    state.lastImportUndo = state.lastStudentCsvText
      ? { kind: 'students', csvText: state.lastStudentCsvText, label: state.lastStudentCsvLabel || 'Base anterior' }
      : null;
    hydrateRows(pending.csvText, `CSV importado: ${pending.fileName}`);
    recordAudit('Importação Acompanhamento', `${pending.review.total} linhas de ${pending.fileName}`);
    closeModal();
    state.pendingImport = null;
    toast('Acompanhamento importado com dados locais preservados.');
    return;
  }
  appendFinancialCsv(pending.kind, pending.csvText, pending.fileName);
  closeModal();
  state.pendingImport = null;
}

function undoLastImport() {
  if (state.lastImportUndo?.kind === 'students') {
    const undo = state.lastImportUndo;
    state.lastImportUndo = null;
    hydrateRows(undo.csvText, `Importação desfeita: ${undo.label}`);
    recordAudit('Desfazer importação', 'Acompanhamento voltou para a base anterior da sessão');
    toast('Última importação do Acompanhamento foi desfeita nesta sessão.');
    return;
  }
  const last = state.store.importHistory.slice(-1)[0];
  if (!last) {
    toast('Nenhuma importação recente para desfazer.');
    return;
  }
  const ids = new Set(last.recordIds || []);
  state.store[last.kind] = state.store[last.kind].filter((record) => !ids.has(record.id));
  state.store.importHistory.pop();
  recordAudit('Desfazer importação', `${last.label}: ${ids.size} linha(s) removidas de ${last.fileName}`);
  persist();
  render();
  toast(`Importação desfeita: ${ids.size} linha(s) removidas.`);
}

function normalizeFinancialRecord(row, kind, index, fileName, importedAt) {
  const studentName = getLooseValue(row, ['Nome', 'Aluno', 'Nome Aluno', 'Nome do Aluno', 'Discente', 'Cliente', 'Sacado', 'Solicitante']);
  const description = getLooseValue(row, ['Descricao', 'Descrição', 'Historico', 'Histórico', 'Lancamento', 'Lançamento', 'Origem', 'Categoria', 'Tipo']);
  const cpf = getLooseValue(row, ['CPF', 'CPF do Aluno', 'Documento', 'CPF Aluno', 'CPF/CNPJ']);
  const ra = getLooseValue(row, ['RA', 'RA do Aluno', 'Registro Academico', 'Registro Acadêmico', 'Matricula', 'Matrícula']);
  const course = getLooseValue(row, ['Curso', 'Nome Curso', 'Produto']);
  const installmentId = getLooseValue(row, [
    'ID da Parcela',
    'Id da Parcela',
    'ID Parcela',
    'Id Parcela',
    'Identificador Parcela',
    'Codigo Parcela',
    'CÃ³digo Parcela',
    'ID Titulo',
    'ID TÃ­tulo',
    'Id Titulo',
    'Id TÃ­tulo',
    'Nosso Numero',
    'Nosso NÃºmero',
    'Numero Documento',
    'NÃºmero Documento',
    'Documento Parcela',
  ]);
  const competence = getLooseValue(row, [
    'Competencia',
    'Competência',
    'Mes/Ano',
    'Mês/Ano',
    'Mes Ano',
    'Mês Ano',
    'Mes / Ano',
    'Mês / Ano',
    'Periodo - Mes',
    'Período - Mês',
    'Mes',
    'Mês',
    'Mes Referencia',
    'Mês Referência',
    'Referencia',
    'Referência',
    'Periodo',
    'Período',
  ]) || cleanText(row.__contextPeriod) || inferFinancialCompetenceFromFile(fileName);
  const date = getLooseValue(row, ['Data', 'DataFaturado', 'Data Faturado', 'Data do Faturamento', 'Data Faturamento', 'Emissao', 'Emissão']);
  const dueDate = getLooseValue(row, ['Vencimento', 'Data Vencimento', 'Data de Vencimento']);
  const paymentDate = getLooseValue(row, ['Pagamento', 'Data Pagamento', 'Data Recebimento', 'Recebimento', 'Data de Recebimento']);
  const amount = financialAmount(row, kind);
  const normalized = {
    kind,
    importedAt,
    importFile: fileName || '',
    competence: competence || inferFinancialCompetence(date || dueDate || paymentDate),
    date,
    dueDate,
    paymentDate,
    description,
    studentName,
    cpf,
    ra,
    course,
    installmentId,
    amount,
    rawJson: JSON.stringify(row),
  };
  normalized.id = financialRecordId(kind, normalized, index);
  return normalized;
}

function financialAmount(row, kind) {
  const common = ['Valor', 'Valor Total', 'Total', 'Total Geral', 'Total geral'];
  const byKind = {
    billing: ['Valor Faturado', 'ValorFaturado', 'Faturado', 'Valor Mensalidade', 'Mensalidade', 'Parcela', 'Valor Parcela', ...common],
    receipts: ['Valor Pago', 'ValorPago', 'Total Geral', 'Total geral', 'Valor Recebido', 'Recebido', 'Pago', 'Banco', 'Cartao', 'Cartão', ...common],
    repasses: ['Repasse Final', 'Valor Repasse Final', 'Valor Repasse', 'Repasse', 'Valor Repassado', 'Repassado', 'Repasse Polo', ...common],
  };
  const direct = parseMoney(getLooseValue(row, byKind[kind] || common));
  if (direct) return direct;
  const bank = parseMoney(getLooseValue(row, ['Banco']));
  const card = parseMoney(getLooseValue(row, ['Cartao', 'Cartão']));
  return bank + card;
}

function isValidFinancialRecord(record, kind) {
  if (!Number(record.amount || 0)) return false;
  const aggregate = [record.studentName, record.description, record.ra, record.cpf].some((value) => {
    const normalized = normalize(value);
    return normalized === 'total' || normalized === 'total geral' || normalized === 'resumo' || normalized === 'analitico';
  });
  if (aggregate) return false;
  if (kind === 'repasses') return true;
  return Boolean(record.cpf || record.ra || record.studentName);
}

function getLooseValue(row, candidates) {
  const entries = Object.entries(row);
  for (const candidate of candidates) {
    const normalizedCandidate = normalize(candidate);
    const found = entries.find(([key]) => normalize(key) === normalizedCandidate);
    if (found && cleanText(found[1])) return cleanText(found[1]);
  }
  for (const candidate of candidates) {
    const normalizedCandidate = normalize(candidate);
    const found = entries.find(([key]) => normalize(key).includes(normalizedCandidate));
    if (found && cleanText(found[1])) return cleanText(found[1]);
  }
  return '';
}

function financialRecordId(kind, record, index) {
  return hashText(
    [
      kind,
      financialStudentKey(record),
      normalizeInstallmentId(record.installmentId),
      normalize(record.competence),
      normalize(record.date || record.dueDate || record.paymentDate),
      Number(record.amount || 0).toFixed(2),
      index,
    ].join('|'),
  );
}

function inferFinancialCompetence(value) {
  const date = parseBrazilianDate(value);
  return date ? `${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}` : '';
}

function inferFinancialCompetenceFromFile(fileName = '') {
  const parsed = parseFinancialPeriod(fileName);
  return parsed ? `${String(parsed.getMonth() + 1).padStart(2, '0')}/${parsed.getFullYear()}` : '';
}

function financialStudentKey(record) {
  const identity = financialIdentityKey(record);
  return identity;
}

function financialIdentityKey(record) {
  const ra = normalizeRa(record.ra);
  if (ra) return `ra:${ra}`;
  const cpf = cleanText(record.cpf).replace(/\D/g, '');
  if (cpf) return `cpf:${cpf}`;
  const name = normalize(record.studentName);
  return name ? `nome:${name}` : '';
}

function rowFinancialIdentityKey(row) {
  const ra = normalizeRa(row.ra);
  if (ra) return `ra:${ra}`;
  const cpf = cleanText(row.cpf).replace(/\D/g, '');
  if (cpf) return `cpf:${cpf}`;
  const name = normalize(row.name);
  return name ? `nome:${name}` : '';
}

function normalizeInstallmentId(value) {
  return normalize(value).replace(/[^a-z0-9]/g, '');
}

function financialCourseKey(record) {
  const course = normalize(record.course);
  if (isPlaceholderValue(record.course) || ['sem curso', 'todo', 'todos', 'tudo'].includes(course)) return '';
  return course;
}

function normalizeRa(value) {
  const text = cleanText(value);
  if (isPlaceholderValue(text)) return '';
  const digits = text.replace(/\D/g, '');
  return digits ? digits.replace(/^0+/, '') || '0' : normalize(text);
}

function isPlaceholderValue(value) {
  const normalized = normalize(value);
  return !normalized || ['-', 'nulo', 'null', 'sem informacao', 'sem informação', 'nao informado', 'não informado'].includes(normalized);
}

function addCourse(data) {
  const name = cleanText(data.name);
  if (!name) return;
  if (getCourseCatalog().some((course) => normalize(course.name) === normalize(name))) {
    toast('Curso já existe no catálogo.');
    return;
  }
  state.store.courses.push({ name, modality: data.modality || 'EAD', local: true });
  recordAudit('Curso criado', name);
  toast('Curso adicionado ao catálogo local.');
}

function addCourse(data) {
  const course = normalizeCourse(data);
  if (!course.name) return;
  upsertCourse(course);
  recordAudit('Curso salvo', `${course.name} - ${formatMoney(course.monthlyFee)}`);
  toast('Curso salvo no catalogo local.');
}

function updateCourseField(key, field, value) {
  const base = getCourseCatalog().find((course) => course.key === key);
  if (!base) return;
  const patch = { ...base, local: true };
  if (field === 'monthlyFee' || field.startsWith('discount')) patch[field] = Number(value || 0);
  else patch[field] = cleanText(value);
  upsertCourse(patch);
  recordAudit('Curso atualizado', `${base.name}: ${field}`);
  persist();
  renderCursos();
}

function importCourseCsv(csvText, fileName = 'Cursos.csv') {
  const parsed = parseCsv(csvText);
  const courses = parsed.rows.map(courseFromCsvRow).filter((course) => course.name);
  courses.forEach(upsertCourse);
  recordAudit('Cursos importados', `${fileName}: ${courses.length} curso(s)`);
  persist();
  render();
  toast(`${courses.length} curso(s) importado(s) ou atualizado(s).`);
}

function courseFromCsvRow(row) {
  return normalizeCourse({
    name: getLooseValue(row, ['Curso', 'Nome Curso', 'Produto']),
    modality: getLooseValue(row, ['Modalidade']),
    habilitation: getLooseValue(row, ['Habilitacao', 'HabilitaÃ§Ã£o', 'Grau', 'Tipo']),
    duration: getLooseValue(row, ['Duracao', 'DuraÃ§Ã£o', 'Tempo']),
    monthlyFee: getLooseValue(row, ['Mensalidade', 'Valor Mensalidade', 'Valor do Curso', 'Valor Curso']),
    discount10: getLooseValue(row, ['Desconto de 10%', 'Desconto 10%', '10%']),
    discount20: getLooseValue(row, ['Desconto de 20%', 'Desconto 20%', '20%']),
    discount30: getLooseValue(row, ['Desconto de 30%', 'Desconto 30%', '30%']),
    discount40: getLooseValue(row, ['Desconto de 40%', 'Desconto 40%', '40%']),
    discount50: getLooseValue(row, ['Desconto de 50%', 'Desconto 50%', '50%']),
    discount60: getLooseValue(row, ['Desconto de 60%', 'Desconto 60%', '60%']),
    authorization: getLooseValue(row, ['Portaria de Reconhecimento/Autorizacao', 'Portaria de Reconhecimento/AutorizaÃ§Ã£o', 'Portaria', 'Reconhecimento', 'Autorizacao', 'AutorizaÃ§Ã£o']),
  });
}

function upsertCourse(course) {
  const normalized = normalizeCourse(course);
  const index = state.store.courses.findIndex((item) => item.key === normalized.key || normalize(item.name) === normalize(normalized.name));
  if (index >= 0) state.store.courses[index] = normalizeCourse({ ...state.store.courses[index], ...normalized, local: true });
  else state.store.courses.push({ ...normalized, local: true });
}

function addLead(data) {
  const lead = normalizeLead({
    id: cryptoId(),
    name: cleanText(data.name),
    phone: cleanText(data.phone),
    course: cleanText(data.course),
    origin: cleanText(data.origin) || 'Não informado',
    stage: data.stage || 'Lead',
    enrollmentFee: Number(data.enrollmentFee || state.store.settings.enrollmentFee || 99),
    consultant: roleLabel(state.profile),
    createdAt: new Date().toISOString(),
  });
  if (lead.stage === 'Matriculado') {
    ensureLocalStudentForLead(lead);
    lead.matriculatedAt = new Date().toISOString();
  }
  state.store.leads.push(lead);
  recordAudit('Candidato criado', `${lead.name} - ${lead.course}`);
  toast(lead.stage === 'Matriculado' ? matriculationAlertMessage(lead) : 'Candidato salvo no funil.');
}

function convertLeadToStudent(id) {
  updateLeadStage(id, 'Matriculado');
}

function updateLeadStage(id, stage) {
  const lead = state.store.leads.find((item) => item.id === id);
  if (!lead) return;
  const oldStage = lead.stage;
  const shouldShowMatriculationModal = stage === 'Matriculado' && oldStage !== 'Matriculado';
  lead.stage = stage;
  lead.updatedAt = new Date().toISOString();
  if (stage === 'Matriculado') {
    if (oldStage !== 'Matriculado') lead.matriculatedAt = new Date().toISOString();
    ensureLocalStudentForLead(lead);
    toast(matriculationAlertMessage(lead));
  } else {
    toast(`Candidato movido para ${leadStageLabel(stage)}.`);
  }
  recordAudit('Status do candidato', `${lead.name}: ${leadStageLabel(oldStage)} -> ${leadStageLabel(stage)}`);
  persist();
  render();
  if (shouldShowMatriculationModal) showMatriculationModal(lead);
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
  if (changes.enrollmentFee !== undefined) {
    lead.enrollmentFee = Number(changes.enrollmentFee || 0);
  }
  lead.updatedAt = new Date().toISOString();
  recordAudit('Pagamento do candidato', `${lead.name}: boleto, pagamento ou taxa atualizado`);
  persist();
  toast('Status da matrícula atualizado.');
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
    ? 'Candidato matriculado: contrato assinado e boleto já marcado como enviado.'
    : `Candidato matriculado: confira o envio do boleto de matrícula para ${lead.name}.`;
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

function addPlanningAction(data) {
  const planType = cleanText(data.planType) || 'monthly';
  const action = normalizeDecision({
    id: cryptoId(),
    planType,
    area: cleanText(data.area) || (planType === 'weeklySeller' ? 'Comercial' : 'Gestão'),
    title: cleanText(data.what),
    what: cleanText(data.what),
    why: cleanText(data.why),
    where: cleanText(data.where),
    when: cleanText(data.when),
    who: cleanText(data.who),
    how: cleanText(data.how),
    howMuch: cleanText(data.howMuch),
    kpi: cleanText(data.kpi),
    status: cleanText(data.status) || 'Planejada',
    periodKey: cleanText(data.periodKey),
    weekStart: cleanText(data.weekStart),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });
  state.store.decisions.push(action);
  recordAudit('Plano 5W2H criado', `${action.who}: ${action.what}`);
  toast(planType === 'weeklySeller' ? 'Plano semanal do vendedor registrado.' : 'Plano 5W2H do próximo mês registrado.');
}

function updatePlanningStatus(id, status) {
  const action = state.store.decisions.find((item) => item.id === id);
  if (!action) return;
  action.status = cleanText(status) || action.status;
  action.updatedAt = new Date().toISOString();
  recordAudit('Status 5W2H atualizado', `${action.what || action.title}: ${action.status}`);
  persist();
  render();
}

function saveTeacher(data) {
  const teacher = normalizeTeacher(data);
  if (!teacher.name) {
    toast('Informe o nome do professor.');
    return;
  }
  if (!Array.isArray(state.store.teachers)) state.store.teachers = [];
  const index = state.store.teachers.findIndex((item) => normalize(item.name) === normalize(teacher.name));
  if (index >= 0) state.store.teachers[index] = { ...state.store.teachers[index], ...teacher, updatedAt: new Date().toISOString() };
  else state.store.teachers.push(teacher);
  recordAudit('Professor salvo', `${teacher.name} - ${teacher.education || 'formacao nao informada'}`);
  toast('Professor salvo na agenda.');
}

function deleteConfiguredTeacher(name) {
  const normalized = normalize(name);
  const before = state.store.teachers.length;
  state.store.teachers = state.store.teachers.filter((teacher) => normalize(teacher.name) !== normalized);
  if (state.store.teachers.length === before) return;
  recordAudit('Professor excluido', name);
  persist();
  render();
  toast('Professor excluido.');
}

function addSchedule(data) {
  const course = cleanText(data.course);
  const cohortYear = Number(data.cohortYear || new Date().getFullYear());
  const semester = Number(data.semester || 1);
  const room = cleanText(data.room);
  const teacher = teacherByName(data.teacher);
  if (!CLASSROOM_OPTIONS.some((item) => normalize(item) === normalize(room))) {
    toast('Escolha um dos ambientes cadastrados para aula.');
    return;
  }
  const students = lessonStudents(course, cohortYear, semester);
  if (!students.length) {
    toast('Nenhum aluno ativo encontrado para esse curso, ano e semestre de ingresso.');
    return;
  }
  const eventItem = {
    id: cryptoId(),
    teacher: cleanText(data.teacher),
    teacherPhone: teacher?.phone || '',
    teacherEmail: teacher?.email || '',
    teacherEducation: teacher?.education || '',
    subject: cleanText(data.subject),
    course,
    cohortYear,
    semester,
    studentKeys: students.map((row) => row.key),
    studentCount: students.length,
    start: data.start,
    end: data.end,
    room,
    capacity: Number(data.capacity || 0),
    status: 'agendado',
  };
  const conflict = scheduleConflict(eventItem);
  if (conflict) {
    toast(conflict);
    return;
  }
  state.store.schedule.push(eventItem);
  toast(`Aula agendada para ${students.length} aluno(s) da turma.`);
}

function addExam(data) {
  const student = resolveStudentSelection(data.student);
  if (!student) {
    toast('Escolha o aluno correto na lista, principalmente quando houver mesmo RA em cursos diferentes.');
    return;
  }
  const exam = {
    id: cryptoId(),
    student: student.name,
    studentKey: student.key,
    cpf: student.cpf,
    ra: student.ra,
    course: student.course,
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
  const delimiter = detectCsvDelimiter(text);

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
    if (char === delimiter && !inQuotes) {
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

  const headerIndex = detectHeaderIndex(rows);
  const contextPeriod = inferCsvContextPeriod(rows.slice(0, headerIndex));
  const headers = (rows[headerIndex] || []).map(cleanText);
  return {
    headers,
    rows: rows
      .slice(headerIndex + 1)
      .filter((row) => row.some(Boolean))
      .map((row) =>
        headers.reduce((record, header, index) => {
          record[header] = cleanText(row[index] ?? '');
          record.__contextPeriod = contextPeriod;
          return record;
        }, {}),
      ),
  };
}

function detectHeaderIndex(rows) {
  let best = 0;
  let bestScore = -1;
  rows.forEach((row, index) => {
    const cells = row.map(cleanText).filter(Boolean);
    if (cells.length < 2) return;
    const normalized = cells.map(normalize);
    const score = normalized.reduce((total, cell) => {
      if (['ra', 'ra do aluno', 'cpf', 'cpf do aluno', 'nome', 'nome do aluno', 'curso', 'parcela', 'valor', 'total geral'].includes(cell)) return total + 3;
      if (cell.includes('aluno') || cell.includes('curso') || cell.includes('competencia') || cell.includes('periodo') || cell.includes('receb') || cell.includes('fatur') || cell.includes('repasse')) return total + 1;
      return total;
    }, 0);
    if (score > bestScore || (score === bestScore && cells.length > rows[best]?.filter(Boolean).length)) {
      best = index;
      bestScore = score;
    }
  });
  return bestScore >= 3 ? best : 0;
}

function inferCsvContextPeriod(rows) {
  const values = rows.flat().map(cleanText).filter(Boolean);
  for (const value of values) {
    if (parseFinancialPeriod(value)) return value;
  }
  return '';
}

function detectCsvDelimiter(text) {
  const lines = cleanText(text).split(/\r?\n/).filter(Boolean).slice(0, 20);
  const score = { ';': 0, ',': 0, '\t': 0 };
  lines.forEach((line) => {
    score[';'] += (line.match(/;/g) || []).length;
    score[','] += (line.match(/,/g) || []).length;
    score['\t'] += (line.match(/\t/g) || []).length;
  });
  const best = Object.entries(score).sort((a, b) => b[1] - a[1])[0];
  return best?.[1] ? best[0] : ',';
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
  const capped = Math.min(100, score);
  return {
    score: capped,
    level: capped >= 55 ? 'Alto' : capped >= 25 ? 'Médio' : 'Baixo',
  };
}

function riskScoreClass(score) {
  if (score >= 55) return 'red';
  if (score >= 25) return 'yellow';
  return 'green';
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
  return ['ativo', 'premat', 'matriculado'].some((item) => normalize(row.localStatus).includes(item));
}

function isRetentionEligible(row) {
  const status = normalize(row.localStatus);
  if (['tranca', 'cancel', 'desist', 'aband'].some((item) => status.includes(item))) return false;
  return ['ativo', 'premat', 'pre-mat', 'matriculado', 'transfer'].some((item) => status.includes(item));
}

function canSeeFinancial() {
  return state.profile === 'admin' || state.profile === 'financeiro';
}

function canAccessAdminModules() {
  return state.profile === 'admin';
}

function isAcompanhamentoBaseStudent(row) {
  return Boolean(row && row.sourceType === 'acompanhamento');
}

function isAcompanhamentoNewEnrollment(row) {
  if (!isAcompanhamentoBaseStudent(row)) return false;
  const enrollmentDate = parseBrazilianDate(row.enrollmentDate);
  const startDate = cohortDate(row.startPeriod);
  return isAcompanhamentoNewEnrollmentByDates(enrollmentDate, startDate);
}

function isAcompanhamentoNewEnrollmentByDates(enrollmentDate, startDate) {
  const hasEnrollmentDate = isValidDate(enrollmentDate);
  const hasStartDate = isValidDate(startDate);
  if (hasStartDate && startDate < ACOMPANHAMENTO_NEW_ENROLLMENT_START) return false;
  if (hasEnrollmentDate) return enrollmentDate >= ACOMPANHAMENTO_NEW_ENROLLMENT_START;
  return Boolean(hasStartDate && startDate >= ACOMPANHAMENTO_NEW_ENROLLMENT_START);
}

function isAcompanhamentoReopeningEnrollment(row) {
  if (!isAcompanhamentoBaseStudent(row)) return false;
  const enrollmentDate = parseBrazilianDate(row.enrollmentDate);
  const startDate = cohortDate(row.startPeriod);
  return Boolean(
    isValidDate(enrollmentDate) &&
      enrollmentDate >= ACOMPANHAMENTO_NEW_ENROLLMENT_START &&
      isValidDate(startDate) &&
      startDate < ACOMPANHAMENTO_NEW_ENROLLMENT_START,
  );
}

function isValidDate(date) {
  return date instanceof Date && !Number.isNaN(date.getTime());
}

function isAcompanhamentoLegacyEnrollment(row) {
  return isAcompanhamentoBaseStudent(row) && !isAcompanhamentoNewEnrollment(row);
}

function isLeadEnrollmentSettled(lead) {
  if (!lead || lead.stage !== 'Matriculado') return false;
  if (lead.enrollmentPaymentStatus === 'Isento') return true;
  return Boolean(lead.boletoSent && lead.enrollmentPaymentStatus === 'Pago');
}

function leadNeedsEnrollmentAction(lead) {
  return Boolean(lead && lead.stage === 'Matriculado' && !isLeadEnrollmentSettled(lead));
}

function leadNeedsEnrollmentBoleto(lead) {
  return Boolean(lead && lead.stage === 'Matriculado' && lead.enrollmentPaymentStatus !== 'Isento' && !lead.boletoSent);
}

function isLeadShadowedByLegacyAcompanhamento(lead) {
  const row = acompanhamentoRowForLead(lead);
  return Boolean(row && isAcompanhamentoLegacyEnrollment(row));
}

function leadRequiresVisibleEnrollmentAction(lead) {
  return leadNeedsEnrollmentAction(lead) && !isLeadShadowedByLegacyAcompanhamento(lead);
}

function leadRequiresVisibleEnrollmentBoleto(lead) {
  return leadNeedsEnrollmentBoleto(lead) && !isLeadShadowedByLegacyAcompanhamento(lead);
}

function confirmedNewMatriculations() {
  return state.store.leads.filter((lead) => isLeadEnrollmentSettled(lead) && !acompanhamentoRowForLead(lead)).length;
}

function studentEnrollmentPaymentState(row, lead = leadForStudent(row)) {
  const override = state.store.overrides[row.key] || {};
  const exempt = Boolean(override.enrollmentExempt || lead?.enrollmentPaymentStatus === 'Isento');
  const boletoOk = Boolean(exempt || override.boletoSent || lead?.boletoSent);
  const paymentOk = Boolean(exempt || override.enrollmentPaid || isLeadEnrollmentSettled(lead));
  return { override, lead, exempt, boletoOk, paymentOk };
}

function acompanhamentoPendingEnrollmentRows(rows = state.allRows, options = {}) {
  return rows.filter((row) => {
    if (!isAcompanhamentoNewEnrollment(row)) return false;
    const enrollment = enrollmentSettlementForStudent(row);
    if (!enrollment.requiresAction) return false;
    return !(options.excludeLeadMatches && enrollment.lead);
  });
}

function pendingEnrollmentActionCount(rows = state.allRows) {
  const pendingLeads = state.store.leads.filter((lead) => leadNeedsEnrollmentAction(lead) && !isLeadShadowedByLegacyAcompanhamento(lead));
  return pendingLeads.length + acompanhamentoPendingEnrollmentRows(rows, { excludeLeadMatches: true }).length;
}

function enrollmentActionItems(rows = state.allRows) {
  const leadItems = state.store.leads
    .filter((lead) => leadNeedsEnrollmentAction(lead) && !isLeadShadowedByLegacyAcompanhamento(lead))
    .map((lead) => {
      const matchedRow = acompanhamentoRowForLead(lead);
      const boletoOk = Boolean(lead.boletoSent || lead.enrollmentPaymentStatus === 'Isento');
      return {
        kind: 'lead',
        id: lead.id,
        name: lead.name,
        course: lead.course,
        origin: matchedRow ? 'CRM + Acompanhamento' : 'CRM',
        boletoOk,
        paymentOk: isLeadEnrollmentSettled(lead),
        boletoLabel: boletoOk ? 'Boleto enviado ou dispensado' : 'Boleto pendente',
        paymentLabel: lead.enrollmentPaymentStatus || 'Pendente',
        feeLabel: formatMoney(lead.enrollmentFee || state.store.settings.enrollmentFee || 99),
        actionLabel: 'Abrir fechamento',
      };
    });
  const studentItems = acompanhamentoPendingEnrollmentRows(rows, { excludeLeadMatches: true }).map((row) => {
    const enrollment = enrollmentSettlementForStudent(row);
    return {
      kind: 'student',
      key: row.key,
      name: row.name,
      course: row.course,
      origin: 'Acompanhamento',
      boletoOk: enrollment.boletoOk,
      paymentOk: enrollment.paymentOk,
      boletoLabel: enrollment.boletoLabel,
      paymentLabel: enrollment.paymentLabel,
      feeLabel: enrollment.feeLabel,
      actionLabel: 'Abrir aluno',
    };
  });
  return [...leadItems, ...studentItems];
}

function enrollmentSettlementForStudent(row) {
  const { override, lead, exempt, boletoOk, paymentOk } = studentEnrollmentPaymentState(row);
  if (isAcompanhamentoLegacyEnrollment(row)) {
    const reopening = isAcompanhamentoReopeningEnrollment(row);
    return {
      sourceBase: true,
      legacySettled: true,
      newEnrollment: false,
      requiresAction: false,
      lead,
      contractSigned: true,
      boletoOk: true,
      paymentOk: true,
      paymentLabel: reopening ? 'Reabertura' : 'Base regularizada',
      boletoLabel: reopening ? 'Sem boleto de nova matricula' : `Sem boleto retroativo ate ${ACOMPANHAMENTO_MATRICULA_CUTOFF_LABEL}`,
      feeLabel: 'Sem cobranca retroativa',
      detail: reopening
        ? `Aluno retornou ao curso com periodo inicial ${humanizeStartPeriod(row.startPeriod)}. O app trata como reabertura, nao como nova matricula, e nao gera boleto de matricula automaticamente.`
        : `Aluno da base Acompanhamento: matricula considerada confirmada e sem boleto retroativo ate ${ACOMPANHAMENTO_MATRICULA_CUTOFF_LABEL}, inclusive.`,
    };
  }
  if (isAcompanhamentoNewEnrollment(row)) {
    const requiresAction = !(boletoOk && paymentOk);
    return {
      sourceBase: true,
      legacySettled: false,
      newEnrollment: true,
      requiresAction,
      lead,
      contractSigned: true,
      boletoOk,
      paymentOk,
      paymentLabel: exempt ? 'Isento' : paymentOk ? 'Pago' : boletoOk ? 'Baixa pendente' : 'Pendente',
      boletoLabel: boletoOk ? 'Boleto enviado ou dispensado' : `Boleto pendente desde ${ACOMPANHAMENTO_NEW_ENROLLMENT_LABEL}`,
      feeLabel: formatMoney(override.enrollmentFee || lead?.enrollmentFee || state.store.settings.enrollmentFee || 99),
      detail: requiresAction
        ? `Matriculado na base Acompanhamento a partir de ${ACOMPANHAMENTO_NEW_ENROLLMENT_LABEL}: confirmar boleto e baixa do pagamento.`
        : 'Boleto e pagamento confirmados pelo polo.',
    };
  }
  return {
    sourceBase: false,
    legacySettled: false,
    newEnrollment: false,
    requiresAction: !(boletoOk && paymentOk),
    lead,
    contractSigned: Boolean(lead?.stage === 'Matriculado' || normalize(row.localStatus).includes('premat') || normalize(row.localStatus).includes('ativo') || normalize(row.localStatus).includes('matriculado')),
    boletoOk,
    paymentOk,
    paymentLabel: override.enrollmentExempt || lead?.enrollmentPaymentStatus === 'Isento' ? 'Isento' : paymentOk ? 'Pago' : 'Pendente',
    boletoLabel: boletoOk ? 'Boleto enviado ou dispensado' : 'Boleto pendente',
    feeLabel: formatMoney(override.enrollmentFee || state.store.settings.enrollmentFee || 99),
    detail: paymentOk ? 'Pago ou isento' : 'Aguardando boleto e baixa manual',
  };
}

function isRestrictedModule(module) {
  if (['financeiro', 'repasse'].includes(module)) return !canSeeFinancial();
  if (['seguranca', 'cursos'].includes(module)) return !canAccessAdminModules();
  return false;
}

function restrictedMessage(module) {
  if (['financeiro', 'repasse'].includes(module)) return 'Acesso disponível apenas para Administrador ou Financeiro.';
  if (['seguranca', 'cursos'].includes(module)) return 'Acesso disponível apenas para Administrador.';
  return 'Acesso restrito.';
}

function getBiFilter() {
  if (!hasFinancialLedgerData()) {
    return allHistoryFilter();
  }
  const today = new Date();
  const settings = state.store.settings || {};
  return {
    month: settings.biMonth === undefined || settings.biMonth === '' ? today.getMonth() + 1 : Number(settings.biMonth),
    year: settings.biYear === undefined || settings.biYear === '' ? today.getFullYear() : Number(settings.biYear),
  };
}

function monthOptions(selected) {
  return [
    `<option value="0" ${Number(selected) === 0 ? 'selected' : ''}>Todos os meses</option>`,
    ...MONTHS.map((month, index) => `<option value="${index + 1}" ${Number(selected) === index + 1 ? 'selected' : ''}>${month}</option>`),
  ].join('');
}

function yearOptions(selected) {
  const current = new Date().getFullYear();
  return [
    `<option value="0" ${Number(selected) === 0 ? 'selected' : ''}>Todo histórico</option>`,
    ...Array.from({ length: 8 }, (_, index) => current - 5 + index)
    .map((year) => `<option value="${year}" ${Number(selected) === year ? 'selected' : ''}>${year}</option>`)
  ].join('');
}

function periodLabel(filter) {
  const month = Number(filter.month || 0);
  const year = Number(filter.year || 0);
  if (!year && !month) return 'Todo histórico';
  if (!year && month) return `${MONTHS[month - 1]} - todos os anos`;
  return month ? `${MONTHS[month - 1]}/${String(year).slice(-2)}` : `Ano ${year}`;
}

function yearlyLabel(filter) {
  return Number(filter.year || 0) ? `Ano ${filter.year}` : 'Todo histórico';
}

function allHistoryFilter() {
  return { month: 0, year: 0 };
}

function hasFinancialLedgerData() {
  return Boolean(state.store.billing.length || state.store.receipts.length || state.store.repasses.length);
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
  if (cohort.year) return new Date(cohort.year, cohortMonthIndex(cohort), 1);
  return monthYearDate(value);
}

function cohortMonthIndex(cohort) {
  const wave = cleanText(cohort.wave).toUpperCase();
  const firstTermMonths = { A: 0, B: 1, C: 2, D: 3, E: 4, F: 5 };
  const secondTermMonths = { A: 6, B: 7, C: 8, D: 9, E: 10, F: 11 };
  if (Number(cohort.term) === 2) return secondTermMonths[wave] ?? 7;
  return firstTermMonths[wave] ?? 1;
}

function monthYearDate(value) {
  const text = normalize(value);
  if (!text) return null;
  const numeric = text.match(/^(\d{1,2})[\/-](\d{4})$/);
  if (numeric) return new Date(Number(numeric[2]), Number(numeric[1]) - 1, 1);
  const yearFirst = text.match(/^(\d{4})[\/-](\d{1,2})$/);
  if (yearFirst) return new Date(Number(yearFirst[1]), Number(yearFirst[2]) - 1, 1);
  const monthNames = {
    janeiro: 0,
    fevereiro: 1,
    marco: 2,
    abril: 3,
    maio: 4,
    junho: 5,
    julho: 6,
    agosto: 7,
    setembro: 8,
    outubro: 9,
    novembro: 10,
    dezembro: 11,
  };
  const named = text.match(/^([a-z]+)[\s\/-]+(\d{4})$/);
  if (named && monthNames[named[1]] !== undefined) return new Date(Number(named[2]), monthNames[named[1]], 1);
  return null;
}

function matchesPeriod(date, filter) {
  if (!date || Number.isNaN(date.getTime())) return false;
  if (Number(filter.year || 0) && date.getFullYear() !== Number(filter.year)) return false;
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
  return state.store.leads.filter((lead) => {
    if (!isLeadEnrollmentSettled(lead)) return false;
    const date = parseBrazilianDate(lead.matriculatedAt || lead.updatedAt || lead.createdAt);
    return matchesPeriod(date, filter);
  }).length;
}

function hasBillingReceiptSource() {
  return state.store.billing.length > 0 || state.store.receipts.length > 0;
}

function buildFinancialPosition(visibleRows = state.allRows) {
  const visibleIdentityKeys = new Set(visibleRows.map(rowFinancialIdentityKey).filter(Boolean));
  const receiptGroups = groupFinancialRecordsByIdentity(state.store.receipts);
  const billingGroups = groupFinancialRecordsByIdentity(state.store.billing);
  const rows = [];

  billingGroups.forEach((group, identityKey) => {
    const receipts = receiptGroups.get(identityKey) || { amount: 0, records: [], periods: {}, courses: [] };
    const first = group.records[0] || {};
    const matched = findStudentForFinancial(first);
    if (visibleIdentityKeys.size && !visibleIdentityKeys.has(identityKey)) return;
    const reconciliation = reconcileFinancialGroup(group, receipts);
    const periods = reconciliation.periods;
    const billed = reconciliation.billed;
    const received = reconciliation.received;
    const receivedApplied = reconciliation.receivedApplied;
    const duePeriods = reconciliation.duePeriods;
    const due = reconciliation.due;
    const interestOrLatePayment = reconciliation.interestOrLatePayment;
    const courses = [...new Set([...(group.courses || []), ...(receipts.courses || [])].filter(Boolean))];
    rows.push({
      key: `finance-${hashText(identityKey)}`,
      sourceFinance: true,
      matchedKey: matched?.key || '',
      name: matched?.name || first.studentName || 'Aluno sem nome',
      cpf: matched?.cpf || first.cpf || '',
      ra: matched?.ra || first.ra || '',
      course: courses.length > 1 ? `${courses.length} cursos` : matched?.course || courses[0] || first.course || '-',
      localStatus: matched?.localStatus || 'Fonte sede',
      debtValue: due,
      isDebt: due > 0.009,
      overdueMonths: duePeriods.map((item) => `${financialPeriodLabel(item.period)} (${formatMoney(item.due)})`).join(', '),
      billed,
      received,
      receivedApplied,
      interestOrLatePayment,
      periods: periods.map((item) => item.period),
      duePeriods,
      billingRecords: group.records,
      receiptRecords: receipts.records,
    });
  });

  return {
    rows,
    debtRows: rows.filter((row) => row.isDebt).sort((a, b) => b.debtValue - a.debtValue),
    paidRows: rows.filter((row) => !row.isDebt),
  };
}

function groupFinancialRecordsByStudent(records) {
  return records.reduce((map, record) => {
    const studentKey = financialStudentKey(record);
    const period = financialPeriodKey(record);
    const group = map.get(studentKey) || { amount: 0, records: [], periods: {} };
    const periodGroup = group.periods[period] || { amount: 0, records: [] };
    const amount = Number(record.amount || 0);
    group.amount += amount;
    group.records.push(record);
    periodGroup.amount += amount;
    periodGroup.records.push(record);
    group.periods[period] = periodGroup;
    map.set(studentKey, group);
    return map;
  }, new Map());
}

function groupFinancialRecordsByIdentity(records) {
  return records.reduce((map, record) => {
    const identityKey = financialIdentityKey(record);
    if (!identityKey) return map;
    const period = financialPeriodKey(record);
    const group = map.get(identityKey) || { amount: 0, records: [], periods: {}, courses: [] };
    const periodGroup = group.periods[period] || { amount: 0, records: [] };
    const amount = Number(record.amount || 0);
    const course = financialCourseKey(record);
    group.amount += amount;
    group.records.push(record);
    if (course && !group.courses.includes(cleanText(record.course))) group.courses.push(cleanText(record.course));
    periodGroup.amount += amount;
    periodGroup.records.push(record);
    group.periods[period] = periodGroup;
    map.set(identityKey, group);
    return map;
  }, new Map());
}

function reconcileFinancialGroup(billingGroup, receiptGroup) {
  const billingWithInstallment = billingGroup.records.filter((record) => normalizeInstallmentId(record.installmentId));
  if (!billingWithInstallment.length) {
    return reconcileByPeriodAmounts(billingGroup.periods, receiptGroup.periods || {}, billingGroup.amount, receiptGroup.amount || 0);
  }

  const receiptByInstallment = groupRecordsByInstallment(receiptGroup.records || []);
  const duePeriods = [];
  let receivedApplied = 0;

  billingWithInstallment.forEach((billingRecord) => {
    const installmentKey = normalizeInstallmentId(billingRecord.installmentId);
    const receiptMatch = receiptByInstallment.get(installmentKey);
    const paid = Number(receiptMatch?.amount || 0);
    const billed = Number(billingRecord.amount || 0);
    receivedApplied += Math.min(paid, billed);
    const due = Math.max(0, billed - paid);
    if (due > 0.009) {
      addDuePeriod(duePeriods, financialPeriodKey(billingRecord), due);
    }
  });

  const billingWithoutInstallment = billingGroup.records.filter((record) => !normalizeInstallmentId(record.installmentId));
  const receiptWithoutInstallment = (receiptGroup.records || []).filter((record) => !normalizeInstallmentId(record.installmentId));
  if (billingWithoutInstallment.length) {
    const fallback = reconcileByPeriodAmounts(
      periodGroupsFromRecords(billingWithoutInstallment),
      periodGroupsFromRecords(receiptWithoutInstallment),
      sumBy(billingWithoutInstallment, (record) => record.amount),
      sumBy(receiptWithoutInstallment, (record) => record.amount),
    );
    fallback.duePeriods.forEach((item) => addDuePeriod(duePeriods, item.period, item.due));
    receivedApplied += fallback.receivedApplied;
  }

  const billed = Number(billingGroup.amount || 0);
  const received = Number(receiptGroup.amount || 0);
  const periods = Object.entries(billingGroup.periods)
    .map(([period, periodGroup]) => ({ period, ...periodGroup }))
    .sort((a, b) => collator.compare(a.period, b.period));
  const due = sumBy(duePeriods, (item) => item.due);
  return {
    periods,
    billed,
    received,
    receivedApplied: Math.min(receivedApplied, billed),
    duePeriods: duePeriods.sort((a, b) => collator.compare(a.period, b.period)),
    due,
    interestOrLatePayment: Math.max(0, received - billed),
  };
}

function reconcileByPeriodAmounts(billingPeriods = {}, receiptPeriods = {}, billed = 0, received = 0) {
  const periods = Object.entries(billingPeriods)
    .map(([period, periodGroup]) => ({ period, ...periodGroup }))
    .sort((a, b) => collator.compare(a.period, b.period));
  let receiptBalance = Object.entries(receiptPeriods).reduce((balance, [period, receiptPeriod]) => {
    if (billingPeriods[period]) return balance;
    return balance + Number(receiptPeriod.amount || 0);
  }, 0);
  let receivedApplied = 0;
  const duePeriods = [];
  const provisionalDue = periods.map((periodGroup) => {
    const samePeriodReceived = Number(receiptPeriods[periodGroup.period]?.amount || 0);
    const samePeriodApplied = Math.min(samePeriodReceived, periodGroup.amount);
    receiptBalance += Math.max(0, samePeriodReceived - periodGroup.amount);
    receivedApplied += samePeriodApplied;
    return {
      ...periodGroup,
      due: Math.max(0, periodGroup.amount - samePeriodApplied),
    };
  });
  provisionalDue.forEach((periodGroup) => {
    const extraApplied = Math.min(receiptBalance, periodGroup.due);
    receiptBalance -= extraApplied;
    receivedApplied += extraApplied;
    const due = Math.max(0, periodGroup.due - extraApplied);
    if (due > 0.009) addDuePeriod(duePeriods, periodGroup.period, due);
  });
  return {
    periods,
    billed: Number(billed || 0),
    received: Number(received || 0),
    receivedApplied: Math.min(receivedApplied, Number(billed || 0)),
    duePeriods,
    due: sumBy(duePeriods, (item) => item.due),
    interestOrLatePayment: Math.max(0, Number(received || 0) - Number(billed || 0)),
  };
}

function periodGroupsFromRecords(records) {
  return records.reduce((periods, record) => {
    const period = financialPeriodKey(record);
    const group = periods[period] || { amount: 0, records: [] };
    group.amount += Number(record.amount || 0);
    group.records.push(record);
    periods[period] = group;
    return periods;
  }, {});
}

function groupRecordsByInstallment(records) {
  return records.reduce((map, record) => {
    const key = normalizeInstallmentId(record.installmentId);
    if (!key) return map;
    const group = map.get(key) || { amount: 0, records: [] };
    group.amount += Number(record.amount || 0);
    group.records.push(record);
    map.set(key, group);
    return map;
  }, new Map());
}

function addDuePeriod(duePeriods, period, due) {
  const existing = duePeriods.find((item) => item.period === period);
  if (existing) existing.due += Number(due || 0);
  else duePeriods.push({ period, due: Number(due || 0) });
}

function financialPeriodKey(record) {
  const date = financialRecordDate(record);
  return date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}` : normalize(record.competence || 'sem-periodo');
}

function groupFinancialRecords(records) {
  return records.reduce((map, record) => {
    const date = financialRecordDate(record);
    const period = date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}` : normalize(record.competence || 'sem-periodo');
    const key = `${financialStudentKey(record)}|${period}`;
    const group = map.get(key) || { amount: 0, records: [], periods: [] };
    group.amount += Number(record.amount || 0);
    group.records.push(record);
    if (!group.periods.includes(period)) group.periods.push(period);
    map.set(key, group);
    return map;
  }, new Map());
}

function findStudentForFinancial(record) {
  const cpf = cleanText(record.cpf).replace(/\D/g, '');
  if (cpf) {
    const byCpf = chooseStudentCandidate(state.allRows.filter((row) => cleanText(row.cpf).replace(/\D/g, '') === cpf), record);
    if (byCpf) return byCpf;
  }
  if (record.ra) {
    const normalizedRa = normalizeRa(record.ra);
    const byRa = chooseStudentCandidate(state.allRows.filter((row) => normalizeRa(row.ra) === normalizedRa), record);
    if (byRa) return byRa;
  }
  if (record.studentName) {
    return chooseStudentCandidate(state.allRows.filter((row) => normalize(row.name) === normalize(record.studentName)), record);
  }
  return null;
}

function chooseStudentCandidate(candidates, record = {}) {
  const unique = [...new Map(candidates.map((row) => [row.key, row])).values()];
  if (!unique.length) return null;
  const course = normalize(record.course);
  if (course) {
    const sameCourse = unique.filter((row) => normalize(row.course) === course);
    if (sameCourse.length === 1) return sameCourse[0];
    if (sameCourse.length > 1) return null;
  }
  return unique.length === 1 ? unique[0] : null;
}

function financialRecordDate(record, preferredKind = '') {
  const kind = preferredKind || record.kind || record.sourceKind || '';
  if (kind === 'receipts') {
    return parseBrazilianDate(record.paymentDate) ||
      parseBrazilianDate(record.date) ||
      parseFinancialPeriod(record.competence) ||
      parseBrazilianDate(record.dueDate);
  }
  if (kind === 'billing') {
    return parseBrazilianDate(record.date) ||
      parseFinancialPeriod(record.competence) ||
      parseBrazilianDate(record.dueDate) ||
      parseBrazilianDate(record.paymentDate);
  }
  if (kind === 'repasses') {
    return parseBrazilianDate(record.paymentDate) ||
      parseBrazilianDate(record.date) ||
      parseFinancialPeriod(record.competence) ||
      parseBrazilianDate(record.dueDate);
  }
  return parseFinancialPeriod(record.competence) ||
    parseBrazilianDate(record.date) ||
    parseBrazilianDate(record.dueDate) ||
    parseBrazilianDate(record.paymentDate);
}

function parseFinancialPeriod(value) {
  const text = cleanText(value);
  if (!text) return null;
  const compact = text.match(/^(\d{4})(\d{2})$/);
  if (compact) return new Date(Number(compact[1]), Number(compact[2]) - 1, 1);
  const yearMonth = text.match(/^(\d{4})[-/](\d{1,2})$/);
  if (yearMonth) return new Date(Number(yearMonth[1]), Number(yearMonth[2]) - 1, 1);
  const monthYear = text.match(/^(\d{1,2})[-/](\d{2,4})$/);
  if (monthYear) {
    const year = Number(monthYear[2].length === 2 ? `20${monthYear[2]}` : monthYear[2]);
    return new Date(year, Number(monthYear[1]) - 1, 1);
  }
  const monthNames = {
    jan: 0,
    janeiro: 0,
    fev: 1,
    fevereiro: 1,
    mar: 2,
    marco: 2,
    março: 2,
    abr: 3,
    abril: 3,
    mai: 4,
    maio: 4,
    jun: 5,
    junho: 5,
    jul: 6,
    julho: 6,
    ago: 7,
    agosto: 7,
    set: 8,
    setembro: 8,
    out: 9,
    outubro: 9,
    nov: 10,
    novembro: 10,
    dez: 11,
    dezembro: 11,
  };
  const readableText = normalize(text)
    .replace(/\./g, '')
    .replace(/\s+de\s+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const named = readableText.match(/([a-z]+)[\s_/-]+(\d{2,4})/);
  if (named && monthNames[named[1]] !== undefined) {
    const year = Number(named[2].length === 2 ? `20${named[2]}` : named[2]);
    return new Date(year, monthNames[named[1]], 1);
  }
  return null;
}

function financialPeriodLabel(period) {
  const match = cleanText(period).match(/^(\d{4})-(\d{2})$/);
  if (!match) return period || '-';
  return `${MONTHS[Number(match[2]) - 1]}/${String(match[1]).slice(-2)}`;
}

function repasseRows(filter) {
  return state.store.repasses
    .map((record) => {
      const date = financialRecordDate(record);
      const periodKey = date
        ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`
        : normalize(record.competence || 'sem-periodo');
      return { ...record, periodDate: date, periodKey };
    })
    .filter((record) => matchesPeriod(record.periodDate, filter) || (!record.periodDate && !Number(filter.month)))
    .sort((a, b) => {
      const dateA = a.periodDate ? a.periodDate.getTime() : 0;
      const dateB = b.periodDate ? b.periodDate.getTime() : 0;
      return dateB - dateA;
    });
}

function repasseReconciliation(filter, currentRepasseTotal = 0) {
  if (!Number(filter.year || 0)) {
    return {
      baseFilter: allHistoryFilter(),
      baseReceived: totalFinancialRecords(state.store.receipts, allHistoryFilter()),
      expectedMin: 0,
      expectedMax: 0,
      rate: 0,
      rateLabel: 'Histórico',
      tone: 'cyan',
    };
  }
  const month = Number(filter.month || new Date().getMonth() + 1);
  const monthFilter = { month, year: Number(filter.year || new Date().getFullYear()) };
  const baseFilter = previousMonthFilter(monthFilter);
  const baseReceived = totalFinancialRecords(state.store.receipts, baseFilter);
  const expectedMin = baseReceived * 0.35;
  const expectedMax = baseReceived * 0.4;
  const rate = baseReceived ? (Number(currentRepasseTotal || 0) / baseReceived) * 100 : 0;
  const tone = !baseReceived ? 'yellow' : rate >= 35 && rate <= 40 ? 'green' : 'yellow';
  return {
    baseFilter,
    baseReceived,
    expectedMin,
    expectedMax,
    rate,
    rateLabel: baseReceived ? `${rate.toFixed(1).replace('.', ',')}%` : 'Sem base',
    tone,
  };
}

function repasseEvolution() {
  const grouped = groupBy(state.store.repasses, (record) => {
    const date = financialRecordDate(record);
    return date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}` : normalize(record.competence || 'sem-periodo');
  });
  const entries = Object.entries(grouped)
    .map(([period, rows]) => ({
      label: financialPeriodLabel(period),
      sortKey: period,
      value: sumBy(rows, (row) => row.amount),
    }))
    .sort((a, b) => collator.compare(a.sortKey, b.sortKey))
    .slice(-8);
  const max = Math.max(...entries.map((item) => item.value), 1);
  return entries.map((item) => ({ ...item, height: Math.max(8, Math.round((item.value / max) * 100)) }));
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
  const overrideRevenue = Object.entries(state.store.overrides).reduce((total, [key, item]) => {
    const row = rowByKey(key);
    if (row && isAcompanhamentoLegacyEnrollment(row)) return total;
    if (row && isAcompanhamentoNewEnrollment(row) && leadForStudent(row)) return total;
    if (!item.enrollmentPaid) return total;
    const date = parseBrazilianDate(item.updatedAt || item.boletoSentAt);
    if (!matchesPeriod(date, filter)) return total;
    return total + Number(item.enrollmentFee || fee);
  }, 0);
  const leadRevenue = state.store.leads.reduce((total, lead) => {
    if (isLeadShadowedByLegacyAcompanhamento(lead)) return total;
    if (!isLeadEnrollmentSettled(lead) || lead.enrollmentPaymentStatus !== 'Pago') return total;
    const date = parseBrazilianDate(lead.updatedAt || lead.matriculatedAt || lead.createdAt);
    if (!matchesPeriod(date, filter)) return total;
    return total + Number(lead.enrollmentFee || fee);
  }, 0);
  return overrideRevenue + leadRevenue;
}

function recurringRevenue(filter) {
  return state.store.receipts.reduce((total, record) => {
    const date = financialRecordDate(record);
    return matchesPeriod(date, filter) ? total + Number(record.amount || 0) : total;
  }, 0);
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

function receiptRepasseSeries(filter) {
  const baseMonth = Number(filter.month || new Date().getMonth() + 1);
  const base = new Date(Number(filter.year), baseMonth - 1, 1);
  return Array.from({ length: 6 }, (_, index) => {
    const date = new Date(base.getFullYear(), base.getMonth() - (5 - index), 1);
    const monthFilter = { month: date.getMonth() + 1, year: date.getFullYear() };
    return {
      label: `${MONTHS[date.getMonth()]}/${String(date.getFullYear()).slice(-2)}`,
      receipts: totalFinancialRecords(state.store.receipts, monthFilter),
      repasses: totalFinancialRecords(state.store.repasses, monthFilter),
    };
  });
}

function totalFinancialRecords(records, filter) {
  return records.reduce((total, record) => {
    const date = financialRecordDate(record);
    if (!date && !Number(filter.year || 0) && !Number(filter.month || 0)) return total + Number(record.amount || 0);
    return matchesPeriod(date, filter) ? total + Number(record.amount || 0) : total;
  }, 0);
}

function gaugeWidget(label, value, detail) {
  const capped = Math.min(Math.max(Number(value) || 0, 0), 100);
  return `
    <article class="decision-widget gauge-widget">
      <div>
        <span>${escapeHtml(label)}</span>
        <strong>${value}%</strong>
        <small>${escapeHtml(detail)}</small>
      </div>
      <div class="gauge-ring" style="--progress:${capped}">
        <em>${value}%</em>
      </div>
    </article>
  `;
}

function donutWidget(label, active, inactive) {
  const total = Math.max(Number(active) + Number(inactive), 1);
  const activePercent = Math.round((Number(active) / total) * 100);
  return `
    <article class="decision-widget donut-widget">
      <div>
        <span>${escapeHtml(label)}</span>
        <strong>${activePercent}% ativos</strong>
        <small>${Number(active).toLocaleString('pt-BR')} ativos / ${Number(inactive).toLocaleString('pt-BR')} inativos</small>
      </div>
      <div class="donut-ring" style="--active:${activePercent}">
        <em>${activePercent}%</em>
      </div>
    </article>
  `;
}

function lineWidget(label, series) {
  const width = 420;
  const height = 170;
  const pad = 28;
  const max = Math.max(...series.flatMap((item) => [item.receipts, item.repasses]), 1);
  const pointsFor = (key) =>
    series
      .map((item, index) => {
        const x = pad + (index * (width - pad * 2)) / Math.max(series.length - 1, 1);
        const y = height - pad - (Number(item[key] || 0) / max) * (height - pad * 2);
        return `${x.toFixed(1)},${y.toFixed(1)}`;
      })
      .join(' ');
  return `
    <article class="decision-widget line-widget">
      <div class="decision-heading">
        <div>
          <span>${escapeHtml(label)}</span>
          <strong>${formatCompactMoney(max)}</strong>
        </div>
        <div class="chart-legend">
          <span><i class="dot cyan"></i> Recebimento</span>
          <span><i class="dot green"></i> Repasse</span>
        </div>
      </div>
      <svg class="line-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(label)}">
        <polyline points="${pointsFor('receipts')}" fill="none" stroke="#0078d4" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"></polyline>
        <polyline points="${pointsFor('repasses')}" fill="none" stroke="#16a34a" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"></polyline>
        ${series
          .map((item, index) => {
            const x = pad + (index * (width - pad * 2)) / Math.max(series.length - 1, 1);
            return `<text x="${x.toFixed(1)}" y="${height - 6}" text-anchor="middle">${escapeHtml(item.label)}</text>`;
          })
          .join('')}
      </svg>
    </article>
  `;
}

function signedPercentChange(current, previous) {
  if (!previous) return current ? '+100%' : '0%';
  const change = Math.round(((current - previous) / Math.abs(previous)) * 100);
  return `${change >= 0 ? '+' : ''}${change}%`;
}

function dashboardFocusData(rows, context) {
  const focus = state.dashboardFocus || 'base';
  const leadRows = state.store.leads;
  const definitions = {
    base: {
      type: 'students',
      title: 'Lista de alunos',
      subtitle: 'Abra o aluno ou ajuste o status na linha',
      rows,
    },
    ativos: {
      type: 'students',
      title: 'Alunos ativos',
      subtitle: 'Base operacional com status ativo, pré-matriculado ou transferência',
      rows: context.activeRows,
    },
    risco: {
      type: 'students',
      title: 'Alunos que precisam de atenção',
      subtitle: 'Priorize contato, orientação acadêmica ou apoio financeiro',
      rows: rows.filter((row) => row.risk.level === 'Alto'),
    },
    ava: {
      type: 'students',
      title: 'Sem acesso ao AVA',
      subtitle: 'Alunos com alerta amarelo ou vermelho',
      rows: context.noAvaRows,
    },
    inadimplencia: {
      type: 'students',
      title: 'Pagamentos em atraso',
      subtitle: 'Alunos com parcelas em aberto na base atual',
      rows: context.debtRows,
    },
    boletos: {
      type: 'enrollmentActions',
      title: 'Matrículas novas a regularizar',
      subtitle: 'Contrato assinado com boleto, pagamento ou isenção ainda pendente',
      rows: context.boletoPending,
    },
    metas: {
      type: 'leads',
      title: 'Funil da meta',
      subtitle: 'Candidatos que ainda não chegaram em Matriculado',
      rows: leadRows.filter((lead) => lead.stage !== 'Matriculado'),
    },
  };
  return definitions[focus] || definitions.base;
}

function dashboardFocusTable(focus) {
  if (focus.type === 'enrollmentActions') {
    return `
      <table>
        <thead>
          <tr>
            <th>Aluno</th>
            <th>Origem</th>
            <th>Curso</th>
            <th>Taxa</th>
            <th>Boleto</th>
            <th>Pagamento</th>
            <th>Ação</th>
          </tr>
        </thead>
        <tbody>
          ${focus.rows
            .slice(0, 100)
            .map(
              (item) => `
                <tr>
                  <td><strong>${escapeHtml(item.name)}</strong></td>
                  <td><span class="badge cyan">${escapeHtml(item.origin)}</span></td>
                  <td>${escapeHtml(item.course || '-')}</td>
                  <td>${escapeHtml(item.feeLabel)}</td>
                  <td><span class="badge ${item.boletoOk ? 'green' : 'yellow'}">${escapeHtml(item.boletoLabel)}</span></td>
                  <td><span class="badge ${item.paymentOk ? 'green' : 'yellow'}">${escapeHtml(item.paymentLabel)}</span></td>
                  <td>${
                    item.kind === 'student'
                      ? `<button class="mini-button" type="button" data-open-student="${escapeHtml(item.key)}">${escapeHtml(item.actionLabel)}</button>`
                      : `<button class="mini-button" type="button" data-quick-module="matriculas" data-view="matricula">${escapeHtml(item.actionLabel)}</button>`
                  }</td>
                </tr>
              `,
            )
            .join('') || emptyRow('Nenhuma matrícula nova pendente de boleto ou baixa.')}
        </tbody>
      </table>
    `;
  }
  if (focus.type === 'leads') {
    return `
      <table>
        <thead>
          <tr>
            <th>Candidato</th>
            <th>Curso</th>
            <th>Status</th>
            <th>Taxa</th>
            <th>Boleto</th>
            <th>Pagamento</th>
          </tr>
        </thead>
        <tbody>
          ${focus.rows
            .slice(0, 100)
            .map(
              (lead) => `
                <tr>
                  <td><strong>${escapeHtml(lead.name)}</strong><span class="subtext">${escapeHtml(lead.phone || '-')}</span></td>
                  <td>${escapeHtml(lead.course || '-')}</td>
                  <td><span class="badge cyan">${escapeHtml(leadStageLabel(lead.stage))}</span></td>
                  <td>${formatMoney(lead.enrollmentFee || state.store.settings.enrollmentFee || 99)}</td>
                  <td><span class="badge ${lead.boletoSent ? 'green' : 'yellow'}">${lead.boletoSent ? 'Enviado' : 'Pendente'}</span></td>
                  <td><span class="badge ${isLeadEnrollmentSettled(lead) ? 'green' : 'yellow'}">${isLeadEnrollmentSettled(lead) ? 'Confirmada' : escapeHtml(lead.enrollmentPaymentStatus || 'Pendente')}</span></td>
                </tr>
              `,
            )
            .join('') || emptyRow('Nenhum candidato encontrado para este filtro.')}
        </tbody>
      </table>
    `;
  }
  return `
    <table>
      <thead>
        <tr>
          <th>Aluno</th>
          <th>CPF</th>
          <th>RA</th>
          <th>Curso</th>
          <th>Início</th>
          <th>Atenção</th>
          <th>Status do polo</th>
          <th>Ações</th>
        </tr>
      </thead>
      <tbody>
        ${focus.rows.slice(0, 100).map(studentRowTemplate).join('') || emptyRow('Nenhum aluno encontrado para este filtro.')}
      </tbody>
    </table>
  `;
}

function nextBestActions(rows) {
  const highRisk = rows.filter((row) => row.risk.level === 'Alto');
  const noAva = rows.filter((row) => ['red', 'yellow'].includes(row.avaAlert.level));
  const noContact = rows.filter((row) => !row.override.lastContact && row.risk.score >= 35);
  const boletos = enrollmentActionItems(rows);
  const actions = [];
  if (boletos.length) {
    actions.push({
      focus: 'boletos',
      tone: 'warning',
      title: 'Regularizar matrículas novas',
      text: `${boletos.length} matrícula(s) novas precisam de boleto, baixa ou isenção.`,
    });
  }
  if (highRisk.length) {
    actions.push({
      focus: 'risco',
      tone: 'danger',
      title: 'Cuidar dos alunos com atenção alta',
      text: `${highRisk.length} aluno(s) precisam de contato ou orientação.`,
    });
  }
  if (noAva.length) {
    actions.push({
      focus: 'ava',
      tone: 'warning',
      title: 'Acompanhar acesso ao AVA',
      text: `${noAva.length} aluno(s) em alerta amarelo ou vermelho de acesso.`,
    });
  }
  if (noContact.length) {
    actions.push({
      focus: 'risco',
      tone: 'warning',
      title: 'Registrar contato pendente',
      text: `${noContact.length} aluno(s) em risco ainda sem registro local recente.`,
    });
  }
  return actions.length
    ? actions.slice(0, 4)
    : [
        {
          focus: 'ativos',
          tone: 'success',
          title: 'Operação sob controle',
          text: 'Use a base ativa para ações preventivas e manutenção da qualidade.',
        },
      ];
}

function productivityItems() {
  const ranking = consultantRanking()[0];
  const retentionRecords = Object.values(state.store.retention);
  const contacted = retentionRecords.filter((item) => item.contacted).length;
  const pending = state.filteredRows.filter((row) => isRetentionEligible(row) && !row.retention.contacted).length;
  const overrides = Object.keys(state.store.overrides).length;
  return [
    {
      name: 'Destaque da captação',
      value: ranking ? `${ranking.total} candidatos` : 'Sem candidatos',
      detail: ranking ? `${ranking.name} · ${ranking.conversion}% conversão` : 'Cadastre candidatos para iniciar o acompanhamento',
    },
    {
      name: 'Contatos registrados',
      value: `${contacted} contatos`,
      detail: `${pending} pendentes na fila operacional`,
    },
    {
      name: 'Dados do polo',
      value: `${overrides} ajustes`,
      detail: 'Status, telefone, e-mail ou pagamento preservados pelo polo',
    },
  ];
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
    key: courseCatalogKey({ name, modality: inferModality(name), habilitation: '' }),
    name,
    modality: inferModality(name),
    habilitation: '',
    duration: '',
    monthlyFee: 0,
    authorization: '',
    students: rows.length,
    local: false,
  }));
  const localCourses = state.store.courses.map((course) => {
    const normalized = normalizeCourse(course);
    return {
      ...normalized,
      students: state.allRows.filter((row) => normalize(row.course) === normalize(normalized.name)).length,
      local: true,
    };
  });
  const merged = new Map();
  [...sheetCourses, ...localCourses].forEach((course) => merged.set(normalize(course.name), course));
  return [...merged.values()].sort((a, b) => collator.compare(a.name, b.name));
}

function normalizeCourse(course = {}) {
  const name = cleanText(course.name || course.Curso);
  const modality = cleanText(course.modality || course.Modalidade) || inferModality(name);
  const habilitation = cleanText(course.habilitation || course.habilitacao || course['HabilitaÃ§Ã£o'] || course.Habilitacao);
  const normalized = {
    key: cleanText(course.key),
    name,
    modality,
    habilitation,
    duration: cleanText(course.duration || course.duracao || course['DuraÃ§Ã£o'] || course.Duracao),
    monthlyFee: parseMoney(course.monthlyFee ?? course.mensalidade ?? course.Mensalidade),
    authorization: cleanText(course.authorization || course.portaria || course.Portaria),
    local: course.local !== false,
    updatedAt: cleanText(course.updatedAt) || new Date().toISOString(),
  };
  COURSE_DISCOUNT_RATES.forEach((rate) => {
    const field = `discount${rate}`;
    normalized[field] = parseMoney(course[field] ?? course[`Desconto de ${rate}%`] ?? course[`Desconto ${rate}%`]);
  });
  normalized.key = normalized.key || courseCatalogKey(normalized);
  return normalized;
}

function courseCatalogKey(course = {}) {
  return hashText([course.name, course.modality, course.habilitation].map(normalize).join('|'));
}

function courseDiscountValue(course, rate) {
  const saved = Number(course[`discount${rate}`] || 0);
  if (saved > 0) return saved;
  const monthly = Number(course.monthlyFee || 0);
  return monthly ? Math.round(monthly * (1 - rate / 100) * 100) / 100 : 0;
}

function leadStageLabel(stage) {
  return {
    Lead: 'Novo candidato',
    Visita: 'Visita agendada',
    'Em negociação': 'Em negociação',
    'Gerar contrato': 'Gerar contrato',
    Matriculado: 'Matriculado',
    Perdido: 'Não avançou',
  }[stage] || stage || '-';
}

function leadColumn(stage) {
  const leads = state.store.leads.filter((lead) => lead.stage === stage);
  return `
    <div class="kanban-column" data-lead-stage="${escapeHtml(stage)}">
      <div class="kanban-head"><strong>${escapeHtml(leadStageLabel(stage))}</strong><span>${leads.length}</span></div>
      ${leads
        .map(
          (lead) => `
            <div class="lead-card ${leadToneClass(lead.stage)} ${leadRequiresVisibleEnrollmentAction(lead) ? 'needs-boleto' : ''}" draggable="true" data-lead-id="${escapeHtml(lead.id)}">
              <strong>${escapeHtml(lead.name)}</strong>
              <span>${escapeHtml(lead.course)}</span>
              <small class="lead-last-record">Tempo na coluna: ${escapeHtml(timeInColumn(lead.updatedAt || lead.createdAt))}</small>
              <div class="lead-badges">
                <span class="badge cyan">Taxa ${formatMoney(lead.enrollmentFee || state.store.settings.enrollmentFee || 99)}</span>
                ${isLeadShadowedByLegacyAcompanhamento(lead) ? '<span class="badge green">Base regularizada</span>' : ''}
                ${leadRequiresVisibleEnrollmentBoleto(lead) ? '<span class="badge yellow">Boleto a enviar</span>' : ''}
                ${lead.stage === 'Matriculado' && lead.boletoSent && lead.enrollmentPaymentStatus === 'Pendente' && !isLeadShadowedByLegacyAcompanhamento(lead) ? '<span class="badge yellow">Baixa pendente</span>' : ''}
                ${isLeadEnrollmentSettled(lead) && !isLeadShadowedByLegacyAcompanhamento(lead) ? '<span class="badge green">Matrícula confirmada</span>' : ''}
              </div>
              ${normalizePhone(lead.phone) ? `<a class="whatsapp-inline lead-whatsapp" href="${whatsAppUrl(lead)}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp Web</a>` : ''}
              ${
                lead.stage === 'Matriculado' && leadRequiresVisibleEnrollmentBoleto(lead) && normalizePhone(lead.phone)
                  ? `<a class="whatsapp-inline lead-whatsapp" href="${boletoWhatsAppUrl(lead)}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp boleto</a>`
                  : ''
              }
              <label class="lead-payment-field lead-fee-field">
                Taxa combinada
                <input type="number" min="0" step="0.01" data-lead-fee="${escapeHtml(lead.id)}" value="${Number(lead.enrollmentFee || state.store.settings.enrollmentFee || 99)}" />
              </label>
              ${leadStageActions(lead)}
              <small>${escapeHtml(lead.origin)} · ${escapeHtml(lead.consultant || 'Atendimento')}</small>
              ${
                stage === 'Matriculado'
                  ? isLeadShadowedByLegacyAcompanhamento(lead)
                    ? `
                    <div class="lead-finance-box">
                      <span class="badge green">Sem boleto retroativo</span>
                      <small>Aluno ja consta na base Acompanhamento regularizada ate ${ACOMPANHAMENTO_MATRICULA_CUTOFF_LABEL}.</small>
                    </div>
                  `
                    : canSeeFinancial()
                    ? `
                    <div class="lead-finance-box">
                      <label class="check-line">
                        <input type="checkbox" data-lead-boleto="${escapeHtml(lead.id)}" ${lead.boletoSent ? 'checked' : ''} />
                        Boleto enviado
                      </label>
                      <label class="lead-payment-field">
                        Pagamento da taxa
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
                      <small>Confirmação de pagamento disponível apenas para Administrador e Financeiro.</small>
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

function leadStageActions(lead) {
  const index = LEAD_STAGES.indexOf(lead.stage);
  const previous = LEAD_STAGES[index - 1];
  const next = LEAD_STAGES[index + 1];
  return `
    <div class="lead-stage-actions">
      ${previous ? `<button type="button" data-lead-move="${escapeHtml(lead.id)}" data-stage="${escapeHtml(previous)}">Voltar</button>` : ''}
      ${next ? `<button type="button" data-lead-move="${escapeHtml(lead.id)}" data-stage="${escapeHtml(next)}">Avançar</button>` : ''}
    </div>
  `;
}

function leadToneClass(stage) {
  const index = LEAD_STAGES.indexOf(stage);
  return `tone-stage-${Math.max(index, 0) + 1}`;
}

function formatLeadDate(value) {
  if (!value) return '-';
  return new Date(value).toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: '2-digit' });
}

function timeInColumn(value) {
  const date = parseBrazilianDate(value);
  if (!date) return 'sem registro';
  const days = Math.max(0, Math.floor((Date.now() - date.getTime()) / 86400000));
  if (days === 0) return 'hoje';
  if (days === 1) return '1 dia';
  return `${days} dias`;
}

function paymentOptions(selected = 'Pendente') {
  return ['Pendente', 'Pago', 'Isento']
    .map((status) => `<option ${status === selected ? 'selected' : ''}>${status}</option>`)
    .join('');
}

function scheduleTemplate(item) {
  const students = scheduleStudents(item);
  const teacher = teacherByName(item.teacher) || {
    name: item.teacher,
    phone: item.teacherPhone,
    email: item.teacherEmail,
    education: item.teacherEducation,
  };
  return `
    <div class="timeline-item">
      <div>
        <strong>${escapeHtml(item.subject)}</strong>
        <span>${formatDateTime(item.start)} - ${formatTime(item.end)} · ${escapeHtml(item.room)}</span>
        <small>${escapeHtml([item.course, item.cohortYear ? `${item.cohortYear}.${item.semester || 1}A-F` : '', `${students.length || item.studentCount || 0} aluno(s)`].filter(Boolean).join(' - '))}</small>
        ${
          normalizePhone(teacher.phone)
            ? `<a class="mini-link" href="${teacherClassWhatsAppUrl(teacher, item, students.length || item.studentCount || 0)}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp professor</a>`
            : '<small>Professor sem WhatsApp cadastrado.</small>'
        }
        ${students.length ? scheduleRosterTemplate(students, item) : '<small>Turma sem alunos vinculados nesta base.</small>'}
      </div>
      <span class="badge ${item.status === 'cancelada' ? 'red' : 'cyan'}">${escapeHtml(item.teacher)}</span>
    </div>
  `;
}

function scheduleRosterTemplate(students, item) {
  return `
    <details class="schedule-roster">
      <summary>Ver alunos e enviar WhatsApp da aula</summary>
      <div class="schedule-student-list">
        ${students
          .map(
            (row) => `
              <div>
                <span><strong>${escapeHtml(row.name)}</strong><small>RA ${escapeHtml(row.ra || '-')} · ${escapeHtml(row.course || '-')}</small></span>
                ${
                  normalizePhone(row.phone)
                    ? `<a class="mini-link" href="${classWhatsAppUrl(row, item)}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp aula</a>`
                    : '<em>Sem WhatsApp</em>'
                }
              </div>
            `,
          )
          .join('')}
      </div>
    </details>
  `;
}

function scheduleStudents(item) {
  if (item.course && item.cohortYear && item.semester) {
    return lessonStudents(item.course, item.cohortYear, item.semester);
  }
  if (Array.isArray(item.studentKeys) && item.studentKeys.length) {
    const keys = new Set(item.studentKeys);
    return state.allRows.filter((row) => keys.has(row.key));
  }
  return lessonStudents(item.course, item.cohortYear, item.semester);
}

function examTemplate(item) {
  return `
    <div class="timeline-item">
      <div>
        <strong>${escapeHtml(item.student)}</strong>
        <span>${escapeHtml(item.discipline)} · ${formatDateTime(item.start)} · ${item.duration} min</span>
        <small>${escapeHtml([item.ra ? `RA ${item.ra}` : '', item.cpf ? `CPF ${item.cpf}` : '', item.course].filter(Boolean).join(' - '))}</small>
      </div>
      <span class="badge yellow">${item.machines} estação</span>
    </div>
  `;
}

function studentLookupDatalist() {
  return `
    <datalist id="studentLookupOptions">
      ${state.allRows
        .slice(0, 1500)
        .map((row) => `<option value="${escapeHtml(studentOptionLabel(row))}"></option>`)
        .join('')}
    </datalist>
  `;
}

function studentOptionLabel(row) {
  return `${row.name} | CPF ${row.cpf || '-'} | RA ${row.ra || '-'} | ${row.course || '-'}`;
}

function resolveStudentSelection(value) {
  const query = normalize(value);
  if (!query) return null;
  const exact = state.allRows.find((row) => normalize(studentOptionLabel(row)) === query);
  if (exact) return exact;
  const digits = cleanText(value).replace(/\D/g, '');
  const matches = state.allRows.filter((row) => {
    const cpf = cleanText(row.cpf).replace(/\D/g, '');
    return normalize(row.name).includes(query) || normalize(row.ra) === query || Boolean(digits && cpf === digits);
  });
  return matches.length === 1 ? matches[0] : null;
}

function lessonStudents(course, cohortYear, semester) {
  const year = Number(cohortYear || 0);
  const term = Number(semester || 0);
  const modality = courseModalityCode(course);
  const rows = state.allRows
    .filter(isRetentionEligible)
    .filter((row) => courseMatchesLesson(row.course, course))
    .filter((row) => {
      const cohort = parseCohort(row.startPeriod);
      if (!cohort.year || Number(cohort.year) !== year || Number(cohort.term) !== term) return false;
      if (modality && cohort.modality && cohort.modality !== modality) return false;
      return ['A', 'B', 'C', 'D', 'E', 'F', ''].includes(cleanText(cohort.wave).toUpperCase());
    })
    .sort((a, b) => collator.compare(a.name, b.name));
  return uniqueLessonStudents(rows);
}

function courseMatchesLesson(studentCourse, selectedCourse) {
  const student = normalize(studentCourse);
  const selected = normalize(selectedCourse);
  if (!selected) return false;
  if (student === selected) return true;
  const studentKey = courseLessonKey(studentCourse);
  const selectedKey = courseLessonKey(selectedCourse);
  return Boolean(studentKey && selectedKey && studentKey === selectedKey);
}

function courseLessonKey(value) {
  return normalize(value)
    .replace(/\b(ead|semipresencial|presencial|bacharelado|licenciatura|tecnologo|tecnologia|pos|graduacao)\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function uniqueLessonStudents(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const cpf = cleanText(row.cpf).replace(/\D/g, '');
    const key = [cpf || normalize(row.name), normalize(row.ra), normalize(row.course)].join('|');
    if (!map.has(key)) map.set(key, row);
  });
  return [...map.values()];
}

function courseModalityCode(course) {
  const normalized = normalize(course);
  if (normalized.includes('semipresencial')) return 'S';
  if (normalized.includes('ead')) return 'E';
  if (normalized.includes('presencial')) return 'P';
  return '';
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
    subtitle: `${item.teacher} · ${item.room}${item.course ? ` · ${item.course}` : ''}`,
    start: item.start,
    status: item.status,
  }));
  const exams = state.store.exams.map((item) => ({
    id: item.id,
    type: 'exam',
    title: item.student,
    course: item.course || '',
    ra: item.ra || '',
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
  if (Number(candidate.studentCount || 0) > Number(candidate.capacity || 0)) {
    return `Lotação insuficiente: a turma tem ${candidate.studentCount} aluno(s) e o ambiente comporta ${candidate.capacity}.`;
  }
  const conflicts = state.store.schedule.filter((item) => overlaps(candidate.start, candidate.end, item.start, item.end));
  if (conflicts.some((item) => normalize(item.teacher) === normalize(candidate.teacher))) return 'Conflito de docência: professor já alocado neste horário.';
  if (conflicts.some((item) => normalize(item.room) === normalize(candidate.room))) return 'Conflito de espaço: sala já ocupada neste horário.';
  if (isComputerLab(candidate.room)) {
    const examConflict = state.store.exams.some((exam) => overlaps(candidate.start, candidate.end, exam.start, addMinutes(exam.start, exam.duration)));
    if (examConflict) return 'Laboratório de Informática bloqueado: já existe prova agendada nesse horário.';
  }
  return '';
}

function examConflict(candidate) {
  const total = Number(state.store.settings.computersTotal || 0);
  const maintenance = Number(state.store.settings.computersMaintenance || 0);
  const end = addMinutes(candidate.start, candidate.duration);
  const labClass = state.store.schedule.some((item) => isComputerLab(item.room) && overlaps(candidate.start, end, item.start, item.end));
  if (labClass) return 'Prova bloqueada: o Laboratório de Informática já está reservado para aula nesse horário.';
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
  const grouped = groupBy(state.store.leads, (lead) => lead.consultant || 'Atendimento');
  const rows = Object.entries(grouped).map(([name, leads]) => {
    const converted = leads.filter(isLeadEnrollmentSettled).length;
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
      title: 'Alunos precisam de atenção',
      text: `${highRisk.toLocaleString('pt-BR')} alunos estão em atenção alta. Priorize contato acadêmico e financeiro.`,
    });
  }

  if (debtRows.length && canSeeFinancial()) {
    signals.push({
      tone: 'danger',
      title: 'Pagamentos em atraso impactam o caixa',
      text: `${debtRows.length.toLocaleString('pt-BR')} alunos somam ${formatMoney(debtTotal)} em atraso.`,
    });
  }

  if (noAccess) {
    signals.push({
      tone: 'warning',
      title: 'Acesso digital em alerta',
      text: `${noAccess.toLocaleString('pt-BR')} alunos sem uso pleno de One/AVA. Acione apoio de acesso.`,
    });
  }

  if (gap < 0) {
    signals.push({
      tone: 'danger',
      title: 'Meta abaixo do necessário',
      text: `Faltam ${Math.abs(gap).toLocaleString('pt-BR')} matrículas para atingir a meta mensal.`,
    });
  }

  if (availableComputers <= 4) {
    signals.push({
      tone: 'danger',
      title: 'Poucos computadores disponíveis',
      text: `Há apenas ${availableComputers} computadores úteis. Replaneje provas ou manutenção.`,
    });
  }

  if (queue.some((item) => item.soon)) {
    signals.push({
      tone: 'info',
      title: 'Atendimento próximo',
      text: 'Existem eventos nos próximos 30 minutos. Prepare sala, docente ou estação de prova.',
    });
  }

  if (retention >= 75 && gap >= 0 && !signals.length) {
    signals.push({
      tone: 'success',
      title: 'Rotina saudável',
      text: 'Acompanhamento, meta e recursos estão em patamar confortável nesta visão.',
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
  const grouped = groupBy(rows, (row) => {
    if (row.sourceFinance) return row.periods?.[0] ? financialPeriodLabel(row.periods[0]) : 'Sem competencia';
    return row.cohort?.year ? `${row.cohort.year}.${row.cohort.term}` : 'Sem safra';
  });
  const entries = Object.entries(grouped)
    .map(([label, items]) => ({ label, value: sumBy(items, (row) => row.debtValue) }))
    .sort((a, b) => collator.compare(a.label, b.label))
    .slice(-7);
  const max = Math.max(...entries.map((item) => item.value), 1);
  return entries.map((item) => ({ ...item, height: Math.max(8, Math.round((item.value / max) * 100)) }));
}

function monthlyDebtEvolution() {
  const billed = groupAmountsByPeriod(state.store.billing);
  const received = groupAmountsByPeriod(state.store.receipts);
  const periods = [...new Set([...Object.keys(billed), ...Object.keys(received)])]
    .filter((period) => period !== 'sem-periodo')
    .sort((a, b) => collator.compare(a, b))
    .slice(-7);
  const entries = periods.map((period) => ({
    label: financialPeriodLabel(period),
    value: Math.max(0, Number(billed[period] || 0) - Number(received[period] || 0)),
  }));
  const fallback = entries.length ? entries : [{ label: 'Sem dados', value: 0 }];
  const max = Math.max(...fallback.map((item) => item.value), 1);
  return fallback.map((item) => ({ ...item, height: Math.max(8, Math.round((item.value / max) * 100)) }));
}

function groupAmountsByPeriod(records) {
  return records.reduce((map, record) => {
    const date = financialRecordDate(record);
    const period = date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}` : normalize(record.competence || 'sem-periodo');
    map[period] = Number(map[period] || 0) + Number(record.amount || 0);
    return map;
  }, {});
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
      <strong>Acesso protegido</strong>
      <p>${escapeHtml(message)}</p>
    </section>
  `;
}

function moduleTitle(title, subtitle) {
  return `
    <header class="module-title module-window-title">
      <div class="module-title-icon" aria-hidden="true">${escapeHtml(moduleTitleIcon(title))}</div>
      <div>
        <p class="eyebrow">Janela de trabalho</p>
        <h1>${escapeHtml(title)}</h1>
        <span>${escapeHtml(subtitle)}</span>
      </div>
      <div class="module-title-actions">
        <button class="mini-button" type="button" data-quick-module="inteligencia">Início</button>
        <button class="mini-button solid-mini" type="button" data-open-mega-menu>Menu principal</button>
      </div>
    </header>
  `;
}

function moduleTitleIcon(title) {
  const normalized = normalize(title);
  if (normalized.includes('indicador')) return '▥';
  if (normalized.includes('atendimento')) return '▧';
  if (normalized.includes('aluno') || normalized.includes('acompanhamento')) return '▦';
  if (normalized.includes('financeiro') || normalized.includes('pagamento') || normalized.includes('repass')) return '▨';
  if (normalized.includes('curso')) return '▧';
  if (normalized.includes('captacao') || normalized.includes('matricula') || normalized.includes('fechamento')) return '▤';
  if (normalized.includes('meta')) return '▥';
  if (normalized.includes('agenda') || normalized.includes('avaliac') || normalized.includes('fila')) return '▧';
  if (normalized.includes('usuario') || normalized.includes('acesso')) return '▩';
  return '▦';
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

function dashboardMetricCard(label, value, subtitle, tone = '', focus = '') {
  return `
    <button class="metric-card clickable ${tone ? `tone-${tone}` : ''}" type="button" data-dashboard-focus="${escapeHtml(focus)}">
      <span>${escapeHtml(label)}</span>
      <strong>${typeof value === 'number' ? value.toLocaleString('pt-BR') : escapeHtml(value)}</strong>
      <small>${escapeHtml(subtitle)}</small>
    </button>
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

function riskReason(row) {
  const reasons = [];
  if (row.isDebt) reasons.push('pagamento em atraso');
  if (['red', 'yellow'].includes(row.avaAlert.level)) reasons.push('alerta no AVA');
  if (!row.override.lastContact && row.risk.score >= 35) reasons.push('sem contato recente');
  if (isInactiveStatus(row)) reasons.push('status inativo');
  return reasons.length ? `Motivos: ${reasons.join(', ')}` : 'Aluno em acompanhamento normal.';
}

function enrollmentChecklist(row) {
  const override = state.store.overrides[row.key] || {};
  const lead = leadForStudent(row);
  const enrollment = enrollmentSettlementForStudent(row);
  return [
    {
      label: 'Contrato assinado',
      done: enrollment.contractSigned,
      detail: enrollment.sourceBase ? 'Base Acompanhamento' : lead?.matriculatedAt ? new Date(lead.matriculatedAt).toLocaleDateString('pt-BR') : row.localStatus,
    },
    {
      label: 'Boleto enviado',
      done: enrollment.boletoOk,
      detail: enrollment.boletoLabel,
    },
    {
      label: 'Taxa confirmada',
      done: enrollment.paymentOk,
      detail: enrollment.detail,
    },
    {
      label: 'Contato validado',
      done: Boolean(override.lastContact || row.retention.contacted),
      detail: override.lastContact || row.retention.reason || 'Sem contato local recente',
    },
    {
      label: 'Primeiro acesso AVA',
      done: row.avaDaysNumber >= 0 && row.avaDaysNumber <= 4,
      detail: row.avaDaysNumber >= 0 ? `${row.avaDaysNumber} dias sem acesso` : 'Sem dado de acesso',
    },
  ];
}

function checklistTemplate(items) {
  return `
    <section class="drawer-section checklist-panel">
      <div class="drawer-section-heading">
        <strong>Checklist de matrícula</strong>
        <span>${items.filter((item) => item.done).length}/${items.length} concluídos</span>
      </div>
      <div class="checklist-list">
        ${items
          .map(
            (item) => `
              <div class="checklist-item ${item.done ? 'done' : 'pending'}">
                <span>${item.done ? '✓' : '!'}</span>
                <div>
                  <strong>${escapeHtml(item.label)}</strong>
                  <small>${escapeHtml(item.detail)}</small>
                </div>
              </div>
            `,
          )
          .join('')}
      </div>
    </section>
  `;
}

function studentTimeline(row) {
  const override = state.store.overrides[row.key] || {};
  const retention = state.store.retention[row.key] || {};
  const lead = leadForStudent(row);
  const financial = buildFinancialPosition([row]).rows.find((item) => item.matchedKey === row.key || normalize(item.name) === normalize(row.name));
  const items = [
    {
      at: row.enrollmentDate || row.startPeriod,
      title: 'Entrada na base',
      text: `${row.course} · ${humanizeStartPeriod(row.startPeriod)}`,
    },
    override.updatedAt && {
      at: override.updatedAt,
      title: 'Dados do polo',
      text: `${override.status || row.localStatus} · ${override.followStatus || 'dados locais atualizados'}`,
    },
    override.lastContact && {
      at: override.updatedAt || new Date().toISOString(),
      title: 'Contato do polo',
      text: override.lastContact,
    },
    retention.updatedAt && {
      at: retention.updatedAt,
      title: 'Acompanhamento',
      text: `${retention.channel || 'Contato'} · ${retention.reason || 'sem motivo informado'}`,
    },
    lead?.updatedAt && {
      at: lead.updatedAt,
      title: 'Captação/Matrícula',
      text: `${leadStageLabel(lead.stage)} · boleto ${lead.boletoSent ? 'enviado' : 'pendente'} · pagamento ${lead.enrollmentPaymentStatus || 'Pendente'}`,
    },
    financial && {
      at: financial.periods?.[0] || new Date().toISOString(),
      title: 'Financeiro',
      text: `Faturado ${formatMoney(financial.billed)} · recebido ${formatMoney(financial.received)} · saldo ${formatMoney(financial.debtValue)}`,
    },
  ].filter(Boolean);
  const auditItems = state.store.auditTrail
    .filter((item) => normalize(item.details).includes(normalize(row.name)) || normalize(item.details).includes(normalize(row.key)))
    .slice(-3)
    .map((item) => ({
      at: item.at,
      title: item.action,
      text: `${item.actor} · ${item.details}`,
    }));
  return [...items, ...auditItems]
    .sort((a, b) => timelineTime(b.at) - timelineTime(a.at))
    .slice(0, 8);
}

function timelineTemplate(items) {
  return `
    <section class="drawer-section timeline-panel">
      <div class="drawer-section-heading">
        <strong>Linha do tempo</strong>
        <span>${items.length} eventos</span>
      </div>
      <div class="timeline-list">
        ${items
          .map(
            (item) => `
              <div class="timeline-entry">
                <i></i>
                <div>
                  <strong>${escapeHtml(item.title)}</strong>
                  <span>${escapeHtml(formatTimelineDate(item.at))}</span>
                  <small>${escapeHtml(item.text)}</small>
                </div>
              </div>
            `,
          )
          .join('') || '<p class="muted">Sem eventos registrados ainda.</p>'}
      </div>
    </section>
  `;
}

function leadForStudent(row) {
  return state.store.leads.find((lead) => {
    return leadMatchesStudentRow(lead, row);
  });
}

function acompanhamentoRowForLead(lead) {
  if (!lead) return null;
  return state.allRows.find((row) => leadMatchesStudentRow(lead, row)) || null;
}

function leadMatchesStudentRow(lead, row) {
  if (!lead || !row) return false;
  const sameCourse = !normalize(lead.course) || !normalize(row.course) || normalize(lead.course) === normalize(row.course);
  if (!sameCourse) return false;
  if (lead.localStudentId && state.store.localStudents.some((student) => student.id === lead.localStudentId && normalize(student.name) === normalize(row.name))) return true;
  if (normalize(lead.name) && normalize(lead.name) === normalize(row.name)) return true;
  return Boolean(normalizePhone(lead.phone) && normalizePhone(lead.phone) === normalizePhone(row.phone));
}

function timelineTime(value) {
  const parsed = parseBrazilianDate(value);
  return parsed ? parsed.getTime() : 0;
}

function formatTimelineDate(value) {
  const parsed = parseBrazilianDate(value);
  return parsed ? parsed.toLocaleString('pt-BR') : cleanText(value) || '-';
}

function roleCard(title, description, active) {
  return `<div class="role-card ${active ? 'active' : ''}"><strong>${escapeHtml(title)}</strong><span>${escapeHtml(description)}</span></div>`;
}

function userRowTemplate(user) {
  const active = normalize(user.ativo || 'SIM') !== 'nao';
  const lastAccess = user.lastAccess ? new Date(user.lastAccess).toLocaleString('pt-BR') : 'Nunca acessou';
  return `
    <tr>
      <td>
        <strong>${escapeHtml(user.nome || '-')}</strong>
        <span class="subtext">${active ? 'Ativo' : 'Inativo'}</span>
      </td>
      <td>${escapeHtml(user.usuario || '-')}</td>
      <td><span class="badge cyan">${escapeHtml(roleLabel(normalizeProfile(user.perfil || 'consultor')))}</span></td>
      <td>${escapeHtml(lastAccess)}</td>
      <td><button class="mini-button" type="button" data-delete-user="${escapeHtml(user.usuario || '')}">Excluir</button></td>
    </tr>
  `;
}

function teacherDatalist() {
  return `
    <datalist id="teacherOptions">
      ${(state.store.teachers || []).map((teacher) => `<option value="${escapeHtml(teacher.name)}"></option>`).join('')}
    </datalist>
  `;
}

function teacherByName(name) {
  return (state.store.teachers || []).find((teacher) => normalize(teacher.name) === normalize(name)) || null;
}

function teacherRowTemplate(teacher) {
  return `
    <tr>
      <td>
        <strong>${escapeHtml(teacher.name)}</strong>
        <span class="subtext">${escapeHtml(teacher.email || 'E-mail nao informado')}</span>
      </td>
      <td>${escapeHtml(teacher.education || '-')}</td>
      <td>
        <span>${escapeHtml(teacher.phone || '-')}</span>
        ${normalizePhone(teacher.phone) ? `<a class="mini-link" href="${teacherWhatsAppUrl(teacher)}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp</a>` : ''}
      </td>
      <td><button class="mini-button" type="button" data-delete-teacher="${escapeHtml(teacher.name)}">Excluir</button></td>
    </tr>
  `;
}

function serviceStatusOptions(selected = 'Novo') {
  return ['Novo', 'Em atendimento', 'Direcionado para sede', 'Respondido', 'Finalizado']
    .map((status) => `<option ${status === selected ? 'selected' : ''}>${status}</option>`)
    .join('');
}

function servicePriorityItems(tickets) {
  return tickets
    .filter((ticket) => !['Finalizado', 'Respondido'].includes(ticket.status))
    .sort((a, b) => {
      const lateA = a.deadline && new Date(a.deadline) < startOfToday() ? 1 : 0;
      const lateB = b.deadline && new Date(b.deadline) < startOfToday() ? 1 : 0;
      return lateB - lateA || new Date(a.deadline || a.requestedAt) - new Date(b.deadline || b.requestedAt);
    })
    .slice(0, 8)
    .map((ticket) => {
      const late = ticket.deadline && new Date(ticket.deadline) < startOfToday();
      return `<div class="decision-signal ${late ? 'danger' : 'warning'}"><strong>${escapeHtml(ticket.studentName)}</strong><span>${escapeHtml(ticket.problem)} · ${escapeHtml(ticket.status)} · prazo ${escapeHtml(ticket.deadline || 'sem prazo')}</span></div>`;
    });
}

function serviceTicketRowTemplate(ticket) {
  const student = rowByKey(ticket.studentKey);
  const phone = student?.phone || '';
  return `
    <tr>
      <td>
        <button class="text-button" type="button" data-open-student="${escapeHtml(ticket.studentKey)}">
          ${escapeHtml(ticket.studentName)}
          <span>RA ${escapeHtml(ticket.ra || '-')} · ${escapeHtml(ticket.course || '-')}</span>
        </button>
      </td>
      <td>${escapeHtml(ticket.protocol || '-')}</td>
      <td>${escapeHtml(ticket.problem)}</td>
      <td>${escapeHtml(formatDate(ticket.requestedAt))}<span class="subtext">Prazo: ${escapeHtml(formatDate(ticket.deadline) || '-')}</span></td>
      <td>${escapeHtml(ticket.attendant || '-')}<span class="subtext">${escapeHtml(ticket.sector || '-')}</span></td>
      <td>
        <textarea class="table-input" rows="2" data-service-response="${escapeHtml(ticket.id)}">${escapeHtml(ticket.response || '')}</textarea>
      </td>
      <td>
        <select class="table-input badge-select ${serviceStatusClass(ticket.status)}" data-service-status="${escapeHtml(ticket.id)}">
          ${serviceStatusOptions(ticket.status)}
        </select>
        ${normalizePhone(phone) ? `<a class="mini-link" href="${serviceWhatsAppUrl({ ...ticket, phone })}" target="mendes_flor_whatsapp_unico" data-whatsapp-link>WhatsApp aluno</a>` : ''}
      </td>
    </tr>
  `;
}

function serviceStatusClass(status) {
  const normalized = normalize(status);
  if (normalized.includes('final') || normalized.includes('respond')) return 'green';
  if (normalized.includes('sede')) return 'yellow';
  return 'cyan';
}

function auditItems() {
  if (state.store.auditTrail.length) {
    return state.store.auditTrail
      .slice(-12)
      .reverse()
      .map(
        (item) =>
          `<div><strong>${escapeHtml(item.action)}</strong><span>${escapeHtml(new Date(item.at).toLocaleString('pt-BR'))} - ${escapeHtml(item.actor)} - ${escapeHtml(item.details)}</span></div>`,
      );
  }
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
  const text = cleanText(value);
  const structured = text.match(/^([A-Z])\.(\d{4})\.(\d)([A-Z])?/i);
  if (structured) {
    return {
      modality: structured[1].toUpperCase(),
      year: Number(structured[2]),
      term: Number(structured[3]),
      wave: (structured[4] || '').toUpperCase(),
    };
  }
  const compact = text.match(/^(\d{4})([12])$/);
  if (compact) {
    return { modality: '', year: Number(compact[1]), term: Number(compact[2]), wave: '' };
  }
  return { modality: '', year: 0, term: 0, wave: '' };
}

function humanizeStartPeriod(value) {
  const date = cohortDate(value);
  if (!date) return cleanText(value) || '-';
  return `${MONTHS_FULL[date.getMonth()]}/${date.getFullYear()}`;
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
    return { level: 'ok', label: 'Verde', className: 'green', weight: 1 };
  }
  return { level: 'unknown', label: 'Sem dado', className: 'cyan', weight: 0 };
}

function parseBrazilianDate(value) {
  const text = cleanText(value);
  if (!text || text === '-') return null;

  const isoDateOnly = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:T.*)?$/);
  if (isoDateOnly) {
    return new Date(Number(isoDateOnly[1]), Number(isoDateOnly[2]) - 1, Number(isoDateOnly[3]));
  }

  const numeric = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);
  if (numeric) {
    const year = Number(numeric[3].length === 2 ? `20${numeric[3]}` : numeric[3]);
    return new Date(year, Number(numeric[2]) - 1, Number(numeric[1]));
  }

  const iso = new Date(text);
  if (!Number.isNaN(iso.getTime())) return iso;

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

async function refreshOperationalData() {
  await loadOperationalStore(true);
  render();
  toast('Dados do banco atualizados para os cálculos.');
}

function normalizeStore(input = {}) {
  return {
    ...defaultStore(),
    ...input,
    overrides: input.overrides || {},
    retention: input.retention || {},
    courses: Array.isArray(input.courses) ? input.courses.map(normalizeCourse) : [],
    leads: Array.isArray(input.leads) ? input.leads.map(normalizeLead) : [],
    teachers: Array.isArray(input.teachers) ? input.teachers.map(normalizeTeacher) : [],
    serviceTickets: Array.isArray(input.serviceTickets) ? input.serviceTickets.map(normalizeServiceTicket) : [],
    localStudents: Array.isArray(input.localStudents) ? input.localStudents : [],
    schedule: Array.isArray(input.schedule) ? input.schedule : [],
    exams: Array.isArray(input.exams) ? input.exams : [],
    archive: Array.isArray(input.archive) ? input.archive : [],
    decisions: Array.isArray(input.decisions) ? input.decisions.map(normalizeDecision) : [],
    snapshots: Array.isArray(input.snapshots) ? input.snapshots.map(normalizeSnapshot) : [],
    taskStatus: input.taskStatus && typeof input.taskStatus === 'object' ? input.taskStatus : {},
    importHistory: Array.isArray(input.importHistory) ? input.importHistory.map(normalizeImportHistory) : [],
    billing: normalizeFinancialCollection(input.billing, 'billing'),
    receipts: normalizeFinancialCollection(input.receipts, 'receipts'),
    repasses: normalizeFinancialCollection(input.repasses, 'repasses'),
    auditTrail: Array.isArray(input.auditTrail) ? input.auditTrail : [],
    settings: input.settings || {},
  };
}

function normalizeSnapshot(snapshot = {}) {
  return {
    periodKey: cleanText(snapshot.periodKey || currentPeriodKey()),
    active: Number(snapshot.active || 0),
    totalStudents: Number(snapshot.totalStudents || 0),
    highRisk: Number(snapshot.highRisk || 0),
    noAva: Number(snapshot.noAva || 0),
    retention: Number(snapshot.retention || 0),
    leads: Number(snapshot.leads || 0),
    confirmedMatriculations: Number(snapshot.confirmedMatriculations || 0),
    pendingEnrollments: Number(snapshot.pendingEnrollments || 0),
    debtCount: Number(snapshot.debtCount || 0),
    debtTotal: Number(snapshot.debtTotal || 0),
    enrollmentRevenue: Number(snapshot.enrollmentRevenue || 0),
    recurringRevenue: Number(snapshot.recurringRevenue || 0),
    repasse: Number(snapshot.repasse || 0),
    qualityScore: Number(snapshot.qualityScore || 0),
    qualityIssues: Number(snapshot.qualityIssues || 0),
    createdAt: cleanText(snapshot.createdAt) || new Date().toISOString(),
  };
}

function normalizeImportHistory(item = {}) {
  return {
    id: cleanText(item.id) || cryptoId(),
    kind: cleanText(item.kind),
    label: cleanText(item.label),
    fileName: cleanText(item.fileName),
    importedAt: cleanText(item.importedAt),
    recordIds: Array.isArray(item.recordIds) ? item.recordIds.map(cleanText) : [],
  };
}

function normalizeDecision(item = {}) {
  const title = cleanText(item.title || item.what);
  const planType = cleanText(item.planType || item.type || 'legacy');
  return {
    ...item,
    id: cleanText(item.id) || cryptoId(),
    planType,
    area: cleanText(item.area) || 'Gestão',
    title,
    what: cleanText(item.what || title),
    why: cleanText(item.why),
    where: cleanText(item.where),
    when: cleanText(item.when || item.due),
    who: cleanText(item.who || item.owner),
    how: cleanText(item.how),
    howMuch: cleanText(item.howMuch),
    kpi: cleanText(item.kpi),
    periodKey: cleanText(item.periodKey),
    weekStart: cleanText(item.weekStart),
    status: cleanText(item.status) || 'Aberta',
    createdAt: cleanText(item.createdAt) || new Date().toISOString(),
    updatedAt: cleanText(item.updatedAt),
  };
}

function normalizeServiceTicket(ticket = {}) {
  return {
    id: cleanText(ticket.id) || cryptoId(),
    studentKey: cleanText(ticket.studentKey),
    studentName: cleanText(ticket.studentName || ticket.student) || 'Aluno não identificado',
    cpf: cleanText(ticket.cpf),
    ra: cleanText(ticket.ra),
    course: cleanText(ticket.course),
    protocol: cleanText(ticket.protocol),
    problem: cleanText(ticket.problem),
    requestedAt: cleanText(ticket.requestedAt) || new Date().toISOString().slice(0, 10),
    deadline: cleanText(ticket.deadline),
    attendant: cleanText(ticket.attendant),
    sector: cleanText(ticket.sector),
    response: cleanText(ticket.response),
    status: cleanText(ticket.status) || 'Novo',
    createdAt: cleanText(ticket.createdAt) || new Date().toISOString(),
    updatedAt: cleanText(ticket.updatedAt) || new Date().toISOString(),
  };
}

function normalizeTeacher(teacher = {}) {
  return {
    id: cleanText(teacher.id) || cryptoId(),
    name: cleanText(teacher.name || teacher.nome),
    education: cleanText(teacher.education || teacher.formacao || teacher['Formação']),
    phone: cleanText(teacher.phone || teacher.telefone || teacher.whatsapp),
    email: cleanText(teacher.email || teacher['E-mail']),
    updatedAt: cleanText(teacher.updatedAt) || new Date().toISOString(),
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
    consultant: cleanText(lead.consultant) || 'Atendimento',
    enrollmentFee: Number(lead.enrollmentFee || 99),
    boletoSent: lead.boletoSent === true || String(lead.boletoSent).toLowerCase() === 'true',
    boletoSentAt: cleanText(lead.boletoSentAt),
    enrollmentPaymentStatus: paymentStatus,
    matriculatedAt: cleanText(lead.matriculatedAt),
    localStudentId: cleanText(lead.localStudentId),
    createdAt: cleanText(lead.createdAt) || new Date().toISOString(),
    updatedAt: cleanText(lead.updatedAt),
  };
}

function normalizeStoredFinancialRecord(record = {}, kind = 'billing', index = 0) {
  const inferred = normalizeFinancialRecord(
    record,
    kind,
    index,
    cleanText(record.importFile || record.__sheetName || ''),
    cleanText(record.importedAt) || new Date().toISOString(),
  );
  const normalized = {
    id: cleanText(record.id),
    kind: cleanText(record.kind || record.sourceKind) || kind,
    importedAt: cleanText(record.importedAt) || inferred.importedAt,
    importFile: cleanText(record.importFile || record.__sheetName) || inferred.importFile,
    competence: cleanText(record.competence) || inferred.competence,
    date: cleanText(record.date) || inferred.date,
    dueDate: cleanText(record.dueDate) || inferred.dueDate,
    paymentDate: cleanText(record.paymentDate) || inferred.paymentDate,
    description: cleanText(record.description) || inferred.description,
    studentName: cleanText(record.studentName) || inferred.studentName,
    cpf: cleanText(record.cpf) || inferred.cpf,
    ra: cleanText(record.ra) || inferred.ra,
    course: cleanText(record.course) || inferred.course,
    installmentId: cleanText(record.installmentId) || inferred.installmentId,
    amount: parseMoney(record.amount) || Number(record.amount || 0) || inferred.amount,
    rawJson: cleanText(record.rawJson) || JSON.stringify(record),
  };
  normalized.id = normalized.id || financialRecordId(kind, normalized, index);
  return normalized;
}

function defaultStore() {
  return {
    overrides: {},
    retention: {},
    courses: [],
    leads: [],
    teachers: [],
    serviceTickets: [],
    localStudents: [],
    schedule: [],
    exams: [],
    archive: [],
    decisions: [],
    snapshots: [],
    taskStatus: {},
    importHistory: [],
    billing: [],
    receipts: [],
    repasses: [],
    auditTrail: [],
    settings: {},
  };
}

function defaultUsers() {
  return [
    { usuario: 'admin', senha: 'admin', nome: 'Administrador', perfil: 'admin', ativo: 'SIM', lastAccess: '' },
    {
      usuario: 'financeiro',
      senha: '123456',
      nome: 'Responsável Financeiro',
      perfil: 'financeiro',
      ativo: 'SIM',
      lastAccess: '',
    },
    {
      usuario: 'retencao',
      senha: '123456',
      nome: 'Responsável pelo Acompanhamento',
      perfil: 'consultor',
      ativo: 'SIM',
      lastAccess: '',
    },
    {
      usuario: 'consultor',
      senha: '123456',
      nome: 'Consultor Comercial',
      perfil: 'consultor',
      ativo: 'SIM',
      lastAccess: '',
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

function openWhatsAppWindow(url) {
  const windowName = 'mendes_flor_whatsapp_unico';

  try {
    if (whatsappWindowRef && !whatsappWindowRef.closed) {
      whatsappWindowRef.location.href = url;
      whatsappWindowRef.focus();
      toast('WhatsApp atualizado na janela que ja estava aberta.');
      return;
    }
  } catch (error) {
    whatsappWindowRef = null;
  }

  const whatsappWindow = window.open('about:blank', windowName);
  if (whatsappWindow) {
    whatsappWindowRef = whatsappWindow;
    whatsappWindowRef.location.href = url;
    whatsappWindowRef.focus();
    return;
  }
  toast('Permita pop-ups para reutilizar a janela única do WhatsApp Web.');
}

function showMatriculationModal(lead) {
  if (!els.modalOverlay) return;
  els.modalOverlay.hidden = false;
  els.modalOverlay.innerHTML = `
    <article class="modal-card" role="dialog" aria-modal="true" aria-label="Boleto de matrícula">
      <button class="icon-button modal-close" type="button" data-modal-close aria-label="Fechar">x</button>
      <p class="eyebrow">Depois da assinatura</p>
      <h2>Boleto de matrícula enviado?</h2>
      <p>${escapeHtml(lead.name)} já está como Matriculado. Confirme o envio do boleto e, se o polo já recebeu, marque como Pago.</p>
      <div class="modal-actions">
        <button class="ghost-button" type="button" data-modal-close>Responder depois</button>
        <button class="mini-button" type="button" data-modal-boleto="${escapeHtml(lead.id)}">Boleto enviado</button>
        <button class="solid-button" type="button" data-modal-paid="${escapeHtml(lead.id)}">Marcar como pago</button>
      </div>
    </article>
  `;
}

function showPasswordModal() {
  if (!els.modalOverlay) return;
  els.modalOverlay.hidden = false;
  els.modalOverlay.innerHTML = `
    <article class="modal-card" role="dialog" aria-modal="true" aria-label="Alterar senha">
      <button class="icon-button modal-close" type="button" data-modal-close aria-label="Fechar">x</button>
      <p class="eyebrow">Segurança da conta</p>
      <h2>Alterar minha senha</h2>
      <p>Informe sua senha atual e escolha uma nova senha para o próximo acesso.</p>
      <form class="stack-form" data-form="change-password">
        <input name="currentPassword" type="password" placeholder="Senha atual" required />
        <input name="newPassword" type="password" minlength="4" placeholder="Nova senha" required />
        <input name="confirmPassword" type="password" minlength="4" placeholder="Confirmar nova senha" required />
        <div class="modal-actions">
          <button class="ghost-button" type="button" data-modal-close>Cancelar</button>
          <button class="solid-button" type="submit">Salvar senha</button>
        </div>
      </form>
    </article>
  `;
}

function closeModal() {
  if (!els.modalOverlay) return;
  els.modalOverlay.hidden = true;
  els.modalOverlay.innerHTML = '';
}

function setLoading(active, message = 'Carregando dados do polo...') {
  if (!els.loadingOverlay) return;
  els.loadingOverlay.hidden = !active;
  const label = els.loadingOverlay.querySelector('strong');
  if (label) label.textContent = message;
}

function recordAudit(action, details) {
  state.store.auditTrail.push({
    id: cryptoId(),
    at: new Date().toISOString(),
    actor: state.currentUser?.nome || roleLabel(state.profile),
    profile: state.profile,
    action,
    details,
  });
  state.store.auditTrail = state.store.auditTrail.slice(-250);
}

function indicatorContext() {
  const filter = getBiFilter();
  const rows = filterRowsByPeriod(state.allRows, filter);
  const censusRows = rows.length ? rows : state.allRows;
  const census = academicCensus(censusRows);
  const commercial = commercialKpis(filter);
  const finance = financeKpis(filter);
  const monthlyTarget = Number(state.store.settings.monthlyTarget || 65);
  const annualTarget = Number(state.store.settings.annualTarget || monthlyTarget * 12);
  return {
    rows: censusRows,
    census,
    commercial,
    finance,
    monthlyProgress: percent(commercial.monthlyMatriculations, monthlyTarget),
    annualProgress: percent(commercial.annualMatriculations, annualTarget),
  };
}

function executiveSummaryRows() {
  const rows = state.filteredRows.length ? state.filteredRows : state.allRows;
  const active = rows.filter(isActive).length;
  const quality = dataQualityIssues(rows);
  const tasks = automaticTaskItems(rows);
  const forecasts = forecastSignals(rows);
  const debtRows = hasBillingReceiptSource() ? buildFinancialPosition(rows).debtRows : rows.filter((row) => row.isDebt);
  return [
    { Indicador: 'Alunos ativos', Valor: `${active}/${rows.length}`, Status: active ? 'Ok' : 'Atenção', Recomendacao: 'Acompanhar retenção por status e safra.' },
    { Indicador: 'Tarefas abertas', Valor: tasks.filter((task) => task.status !== 'Concluída').length, Status: tasks.length ? 'Acompanhar' : 'Ok', Recomendacao: 'Cobrar donos e prazos na reunião semanal.' },
    { Indicador: 'Qualidade dos dados', Valor: `${dataQualityScore(quality, rows.length)}%`, Status: quality.length ? 'Atenção' : 'Ok', Recomendacao: 'Corrigir alertas críticos antes de decidir com a base.' },
    { Indicador: 'Inadimplência', Valor: formatMoney(sumBy(debtRows, (row) => row.debtValue)), Status: debtRows.length ? 'Crítico' : 'Ok', Recomendacao: 'Priorizar cobrança por valor e mês em aberto.' },
    { Indicador: 'Matrículas a regularizar', Valor: pendingEnrollmentActionCount(rows), Status: pendingEnrollmentActionCount(rows) ? 'Atenção' : 'Ok', Recomendacao: 'Confirmar boleto, baixa ou isenção.' },
    ...forecasts.map((item) => ({ Indicador: item.title, Valor: item.text, Status: item.tone === 'danger' ? 'Crítico' : item.tone === 'warning' ? 'Atenção' : 'Ok', Recomendacao: 'Usar a previsão para definir prioridades da semana.' })),
  ];
}

function openExecutiveReport(kind = 'executivo') {
  const title = kind === 'vendedores' ? 'Relatório de Vendedores' : 'Relatório Executivo do Polo';
  const rows = kind === 'vendedores' ? sellerPerformanceRows() : executiveSummaryRows();
  const content = kind === 'vendedores' ? sellerPerformanceTable(rows) : simpleReportTable(rows);
  const win = window.open('', 'relatorio_polo');
  if (!win) {
    toast('Permita pop-ups para abrir o relatório imprimível.');
    return;
  }
  const generatedAt = new Date().toLocaleString('pt-BR');
  win.document.write(`
    <!doctype html>
    <html lang="pt-BR">
      <head>
        <meta charset="utf-8" />
        <title>${escapeHtml(title)}</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 28px; color: #12231f; }
          h1 { color: #004a99; margin-bottom: 4px; }
          p { color: #52615d; }
          table { width: 100%; border-collapse: collapse; margin-top: 18px; font-size: 12px; }
          th, td { border: 1px solid #d8e1df; padding: 8px; text-align: left; vertical-align: top; }
          th { background: #e8f2ff; color: #004a99; }
          strong { color: #10231f; }
          .subtext { display:block; color:#52615d; font-size:11px; margin-top:3px; }
          .badge { display:inline-block; padding:3px 7px; border-radius:8px; background:#eef3f2; }
          @media print { button { display:none; } body { margin: 12mm; } }
        </style>
      </head>
      <body>
        <button onclick="window.print()">Salvar como PDF / Imprimir</button>
        <h1>${escapeHtml(title)}</h1>
        <p>MENDES & FLOR EDUCACIONAL - Polo UniFECAF · Gerado em ${escapeHtml(generatedAt)}</p>
        ${content}
      </body>
    </html>
  `);
  win.document.close();
  recordAudit('Relatório imprimível', title);
  persist();
}

function simpleReportTable(rows) {
  return `
    <table>
      <thead><tr><th>Indicador</th><th>Valor</th><th>Status</th><th>Recomendação</th></tr></thead>
      <tbody>${rows.map((row) => `<tr><td><strong>${escapeHtml(row.Indicador)}</strong></td><td>${escapeHtml(row.Valor)}</td><td>${escapeHtml(row.Status)}</td><td>${escapeHtml(row.Recomendacao)}</td></tr>`).join('')}</tbody>
    </table>
  `;
}

function governanceRows() {
  const users = state.users.length ? state.users : defaultUsers();
  const exports = state.store.auditTrail.filter((item) => normalize(item.action).includes('export') || normalize(item.action).includes('relatorio'));
  return [
    { Item: 'Controle de perfis', Status: 'Ativo', Evidencia: `${users.length} usuário(s) cadastrados`, Acao: 'Revisar usuários inativos mensalmente.' },
    { Item: 'Acesso financeiro', Status: canAccessAdminModules() ? 'Restrito' : 'Bloqueado', Evidencia: 'Financeiro protegido por perfil', Acao: 'Manter acesso apenas para gestão autorizada.' },
    { Item: 'Trilha de auditoria', Status: state.store.auditTrail.length ? 'Ativa' : 'Sem logs', Evidencia: `${state.store.auditTrail.length} registro(s)`, Acao: 'Consultar logs em alterações críticas.' },
    { Item: 'Exportações', Status: exports.length ? 'Rastreada' : 'Sem exportações', Evidencia: `${exports.length} exportação/relatório`, Acao: 'Exportar dados pessoais apenas para finalidade de gestão.' },
    { Item: 'Dados sensíveis', Status: 'Atenção', Evidencia: 'CPF, telefone, e-mail e financeiro', Acao: 'Não compartilhar relatórios fora da equipe autorizada.' },
    { Item: 'Persistência local', Status: state.remoteState ? 'Planilha operacional' : 'Local', Evidencia: state.remoteState ? 'Estado salvo em planilha operacional' : 'Fallback neste navegador', Acao: 'Manter backup e evitar limpar dados do navegador.' },
  ];
}

function governanceTable() {
  return `
    <table>
      <thead><tr><th>Item</th><th>Status</th><th>Evidência</th><th>Ação recomendada</th></tr></thead>
      <tbody>${governanceRows().map((row) => `<tr><td><strong>${escapeHtml(row.Item)}</strong></td><td><span class="badge cyan">${escapeHtml(row.Status)}</span></td><td>${escapeHtml(row.Evidencia)}</td><td>${escapeHtml(row.Acao)}</td></tr>`).join('')}</tbody>
    </table>
  `;
}

function exportDataset(kind) {
  const datasets = {
    leads: {
      filename: 'leads-matriculas.csv',
      rows: state.store.leads,
      headers: ['name', 'phone', 'course', 'origin', 'stage', 'enrollmentFee', 'boletoSent', 'enrollmentPaymentStatus'],
    },
    inadimplentes: {
      filename: 'inadimplentes.csv',
      rows: buildFinancialPosition(state.filteredRows).debtRows,
      headers: ['name', 'cpf', 'ra', 'course', 'localStatus', 'debtValue', 'overdueMonths', 'billed', 'received'],
      financial: true,
    },
    ativos: {
      filename: 'alunos-ativos.csv',
      rows: state.filteredRows.filter(isActive),
      headers: ['name', 'cpf', 'ra', 'course', 'startPeriod', 'localStatus', 'phone', 'email'],
    },
    retencao: {
      filename: 'retencao-ava.csv',
      rows: state.filteredRows.filter(isRetentionEligible),
      headers: ['name', 'cpf', 'ra', 'course', 'avaAccess', 'avaLastAccessDate', 'avaDaysNumber', 'phone', 'email'],
    },
    atendimentos: {
      filename: 'atendimentos-alunos.csv',
      rows: state.store.serviceTickets,
      headers: ['studentName', 'cpf', 'ra', 'course', 'protocol', 'problem', 'requestedAt', 'deadline', 'attendant', 'sector', 'response', 'status'],
    },
    cursos: {
      filename: 'catalogo-cursos.csv',
      rows: getCourseCatalog().map((course) => ({
        Curso: course.name,
        Modalidade: course.modality,
        Habilitacao: course.habilitation,
        Duracao: course.duration,
        Mensalidade: course.monthlyFee,
        Desconto10: courseDiscountValue(course, 10),
        Desconto20: courseDiscountValue(course, 20),
        Desconto30: courseDiscountValue(course, 30),
        Desconto40: courseDiscountValue(course, 40),
        Desconto50: courseDiscountValue(course, 50),
        Desconto60: courseDiscountValue(course, 60),
        Portaria: course.authorization,
        Alunos: course.students,
      })),
      headers: ['Curso', 'Modalidade', 'Habilitacao', 'Duracao', 'Mensalidade', 'Desconto10', 'Desconto20', 'Desconto30', 'Desconto40', 'Desconto50', 'Desconto60', 'Portaria', 'Alunos'],
    },
    repasse: {
      filename: 'repasse-sede.csv',
      rows: repasseRows(getBiFilter()),
      headers: ['competence', 'date', 'paymentDate', 'description', 'course', 'amount', 'importFile', 'importedAt'],
      financial: true,
    },
    executivo: {
      filename: 'relatorio-executivo-polo.csv',
      rows: executiveSummaryRows(),
      headers: ['Indicador', 'Valor', 'Status', 'Recomendacao'],
    },
    vendedores: {
      filename: 'painel-vendedores.csv',
      rows: sellerPerformanceRows().map((row) => ({
        Vendedor: row.seller,
        Leads: row.total,
        Ativos: row.active,
        Visitas: row.visits,
        Contratos: row.contracts,
        Matriculas: row.confirmed,
        Pendencias: row.pending,
        Conversao: `${row.conversion}%`,
        TicketMedio: row.avgTicket,
        Semana5W2H: `${row.weeklyDone}/${row.weeklyActions}`,
      })),
      headers: ['Vendedor', 'Leads', 'Ativos', 'Visitas', 'Contratos', 'Matriculas', 'Pendencias', 'Conversao', 'TicketMedio', 'Semana5W2H'],
    },
    qualidade: {
      filename: 'qualidade-dos-dados.csv',
      rows: dataQualityIssues(state.filteredRows.length ? state.filteredRows : state.allRows).map((item) => ({
        Severidade: item.severity === 'red' ? 'Crítico' : 'Atenção',
        Problema: item.title,
        Quantidade: item.count,
        Acao: item.recommendation,
      })),
      headers: ['Severidade', 'Problema', 'Quantidade', 'Acao'],
    },
    tarefas: {
      filename: 'tarefas-automaticas.csv',
      rows: automaticTaskItems(state.filteredRows.length ? state.filteredRows : state.allRows).map((task) => ({
        Prioridade: task.priority,
        Tarefa: task.title,
        Area: task.area,
        Dono: task.owner,
        Prazo: formatShortDate(task.due),
        Origem: task.source,
        Status: task.status,
      })),
      headers: ['Prioridade', 'Tarefa', 'Area', 'Dono', 'Prazo', 'Origem', 'Status'],
    },
    indicadores: {
      filename: 'dicionario-indicadores.csv',
      rows: indicatorDictionaryRows(indicatorContext()).map((item) => ({
        Indicador: item.name,
        Formula: item.formula,
        Fonte: item.source,
        Dono: item.owner,
        ValorAtual: item.current,
        Saude: item.health,
      })),
      headers: ['Indicador', 'Formula', 'Fonte', 'Dono', 'ValorAtual', 'Saude'],
    },
    snapshots: {
      filename: 'historico-mensal.csv',
      rows: state.store.snapshots.map((item) => ({
        Periodo: periodLabelFromKey(item.periodKey),
        Ativos: item.active,
        TotalAlunos: item.totalStudents,
        Retencao: `${item.retention}%`,
        RiscoAlto: item.highRisk,
        MatriculasPendentes: item.pendingEnrollments,
        Inadimplencia: item.debtTotal,
        Recebido: Number(item.enrollmentRevenue || 0) + Number(item.recurringRevenue || 0),
        Qualidade: `${item.qualityScore}%`,
      })),
      headers: ['Periodo', 'Ativos', 'TotalAlunos', 'Retencao', 'RiscoAlto', 'MatriculasPendentes', 'Inadimplencia', 'Recebido', 'Qualidade'],
    },
    '5w2h': {
      filename: 'planejamento-5w2h.csv',
      rows: state.store.decisions.filter((item) => ['monthly', 'weeklySeller'].includes(item.planType)).map((item) => ({
        Tipo: item.planType === 'weeklySeller' ? 'Semana vendedor' : 'Próximo mês',
        Area: item.area,
        OQue: item.what,
        Porque: item.why,
        Onde: item.where,
        Quando: item.when,
        Quem: item.who,
        Como: item.how,
        Quanto: item.howMuch,
        KPI: item.kpi,
        Status: item.status,
      })),
      headers: ['Tipo', 'Area', 'OQue', 'Porque', 'Onde', 'Quando', 'Quem', 'Como', 'Quanto', 'KPI', 'Status'],
    },
    governanca: {
      filename: 'governanca-lgpd.csv',
      rows: governanceRows(),
      headers: ['Item', 'Status', 'Evidencia', 'Acao'],
      financial: true,
    },
  };
  const dataset = datasets[kind];
  if (!dataset) return;
  if (dataset.financial && !canSeeFinancial()) {
    toast('Exportação financeira disponível apenas para Administrador e Financeiro.');
    return;
  }
  const csv = toCsv(dataset.rows, dataset.headers);
  const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = dataset.filename;
  link.click();
  URL.revokeObjectURL(url);
  recordAudit('Exportacao', `${dataset.filename} - ${dataset.rows.length} registros`);
  persist();
  toast('Arquivo CSV gerado para Excel.');
}

function toCsv(rows, headers) {
  const labels = headers.join(';');
  const body = rows.map((row) => headers.map((header) => csvCell(row[header])).join(';')).join('\n');
  return [labels, body].filter(Boolean).join('\n');
}

function csvCell(value) {
  const text = cleanText(value).replaceAll('"', '""');
  return `"${text}"`;
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
  const message = `Olá, ${firstName(row.name)}! Tudo bem? Aqui é da MENDES & FLOR EDUCACIONAL - Polo UniFECAF. Estou entrando em contato para acompanhar sua matrícula no curso ${row.course}.`;
  return `https://web.whatsapp.com/send?phone=${normalizePhone(row.phone)}&text=${encodeURIComponent(message)}`;
}

function classWhatsAppUrl(row, lesson) {
  const message = `Olá, ${firstName(row.name)}! Tudo bem? Aqui é da MENDES & FLOR EDUCACIONAL - Polo UniFECAF. Passando para lembrar da aula "${lesson.subject}" do curso ${lesson.course || row.course}, agendada para ${formatDateTime(lesson.start)} na ${lesson.room}.`;
  return `https://web.whatsapp.com/send?phone=${normalizePhone(row.phone)}&text=${encodeURIComponent(message)}`;
}

function teacherWhatsAppUrl(teacher) {
  const message = `Olá, ${firstName(teacher.name)}! Tudo bem? Aqui é da MENDES & FLOR EDUCACIONAL - Polo UniFECAF. Este é o contato cadastrado para avisos de aulas e alinhamentos acadêmicos.`;
  return `https://web.whatsapp.com/send?phone=${normalizePhone(teacher.phone)}&text=${encodeURIComponent(message)}`;
}

function teacherClassWhatsAppUrl(teacher, lesson, studentCount = 0) {
  const message = `Olá, ${firstName(teacher.name)}! Tudo bem? Confirmando a aula "${lesson.subject}" do curso ${lesson.course}, agendada para ${formatDateTime(lesson.start)} na ${lesson.room}. Turma prevista: ${studentCount} aluno(s).`;
  return `https://web.whatsapp.com/send?phone=${normalizePhone(teacher.phone)}&text=${encodeURIComponent(message)}`;
}

function boletoWhatsAppUrl(row) {
  const fee = Number(row.enrollmentFee || row.override?.enrollmentFee || state.store.settings.enrollmentFee || 99);
  const message = `Olá, ${firstName(row.name)}! Tudo bem? Aqui é da MENDES & FLOR EDUCACIONAL - Polo UniFECAF. Estou enviando o lembrete do boleto da taxa de matrícula do curso ${row.course || 'contratado'}. Valor combinado: ${formatMoney(fee)}. Qualquer dúvida, me chame por aqui.`;
  return `https://web.whatsapp.com/send?phone=${normalizePhone(row.phone)}&text=${encodeURIComponent(message)}`;
}

function overdueWhatsAppUrl(row) {
  const months = cleanText(row.overdueMonths) || 'mensalidade(s) em aberto';
  const message = `Olá, ${firstName(row.name)}! Tudo bem? Aqui é da MENDES & FLOR EDUCACIONAL - Polo UniFECAF. Identificamos pendência financeira referente a ${months}, no valor de ${formatMoney(row.debtValue || 0)}. Podemos verificar juntos a melhor forma de regularizar?`;
  return `https://web.whatsapp.com/send?phone=${normalizePhone(row.phone)}&text=${encodeURIComponent(message)}`;
}

function serviceWhatsAppUrl(ticket) {
  const message = `Olá, ${firstName(ticket.studentName)}! Tudo bem? Aqui é da MENDES & FLOR EDUCACIONAL - Polo UniFECAF. Estamos acompanhando sua solicitação${ticket.protocol ? ` protocolo ${ticket.protocol}` : ''}: ${ticket.problem}. Situação atual: ${ticket.status}.`;
  return `https://web.whatsapp.com/send?phone=${normalizePhone(ticket.phone)}&text=${encodeURIComponent(message)}`;
}

function firstName(value) {
  return cleanText(value).split(/\s+/)[0] || '';
}

function maskCpf(value) {
  const digits = cleanText(value).replace(/\D/g, '');
  if (!digits) return '-';
  if (digits.length !== 11) return cleanText(value);
  return `${digits.slice(0, 3)}.${digits.slice(3, 6)}.${digits.slice(6, 9)}-${digits.slice(9)}`;
}

function statusClass(status) {
  const normalized = normalize(status);
  if (normalized.includes('ativo') || normalized.includes('premat') || normalized.includes('matriculado')) return 'green';
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
  const cleaned = text.replace(/[^\d,.-]/g, '');
  if (!cleaned) return 0;
  if (cleaned.includes(',')) return Number.parseFloat(cleaned.replace(/\./g, '').replace(',', '.')) || 0;
  return Number.parseFloat(cleaned) || 0;
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

function formatDate(value) {
  if (!value) return '';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? cleanText(value) : date.toLocaleDateString('pt-BR');
}

function startOfToday() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
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

function isComputerLab(room) {
  return normalize(room).includes('laboratorio de informatica');
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
  return { admin: 'Administrador', financeiro: 'Financeiro', consultor: 'Atendimento' }[role] || role;
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

function hashText(value) {
  let hash = 0;
  const text = cleanText(value);
  for (let index = 0; index < text.length; index++) {
    hash = (hash << 5) - hash + text.charCodeAt(index);
    hash |= 0;
  }
  return `h${Math.abs(hash).toString(36)}`;
}
