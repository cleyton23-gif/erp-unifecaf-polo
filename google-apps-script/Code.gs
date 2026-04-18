const ERP_TABS = {
  config: {
    name: 'ERP_Config',
    headers: ['chave', 'valor', 'atualizadoEm'],
  },
  users: {
    name: 'ERP_Usuarios',
    headers: ['usuario', 'senha', 'nome', 'perfil', 'ativo'],
  },
  overrides: {
    name: 'ERP_Overrides',
    headers: [
      'key',
      'status',
      'note',
      'lastContact',
      'contactPhone',
      'contactEmail',
      'followStatus',
      'enrollmentPaid',
      'enrollmentExempt',
      'boletoSent',
      'boletoSentAt',
      'enrollmentFee',
      'updatedAt',
    ],
  },
  retention: {
    name: 'ERP_Retencao',
    headers: [
      'key',
      'contacted',
      'contactDate',
      'channel',
      'responsible',
      'reason',
      'note',
      'updatedAt',
    ],
  },
  courses: {
    name: 'ERP_Cursos',
    headers: ['name', 'modality', 'local'],
  },
  leads: {
    name: 'ERP_Leads',
    headers: [
      'id',
      'name',
      'phone',
      'course',
      'origin',
      'stage',
      'consultant',
      'boletoSent',
      'boletoSentAt',
      'enrollmentPaymentStatus',
      'matriculatedAt',
      'localStudentId',
      'createdAt',
      'updatedAt',
    ],
  },
  localStudents: {
    name: 'ERP_Alunos_Local',
    headers: ['id', 'name', 'phone', 'course', 'startPeriod', 'status', 'createdAt'],
  },
  schedule: {
    name: 'ERP_Agenda',
    headers: ['id', 'teacher', 'subject', 'start', 'end', 'room', 'capacity', 'status', 'reason'],
  },
  exams: {
    name: 'ERP_Provas',
    headers: ['id', 'student', 'discipline', 'start', 'duration', 'machines', 'status'],
  },
  archive: {
    name: 'ERP_Arquivo',
    headers: [
      'id',
      'type',
      'teacher',
      'subject',
      'student',
      'discipline',
      'start',
      'end',
      'room',
      'status',
      'finishedAt',
    ],
  },
  decisions: {
    name: 'ERP_Decisoes',
    headers: ['id', 'area', 'title', 'owner', 'due', 'status', 'createdAt'],
  },
  audit: {
    name: 'ERP_Auditoria',
    headers: ['timestamp', 'actor', 'action', 'details'],
  },
};

function doGet(event) {
  if (!authorize_(event)) {
    return json_({ ok: false, message: 'Token invalido.' }, 401);
  }

  const action = event.parameter.action || 'read';
  if (action === 'setup') {
    setupErpTabs_();
    return json_({ ok: true, message: 'Abas ERP criadas ou validadas.' });
  }

  if (action === 'users') {
    setupErpTabs_();
    return json_({ ok: true, users: readUsers_(), readAt: new Date().toISOString() });
  }

  setupErpTabs_();
  return json_({
    ok: true,
    state: readState_(),
    users: readUsers_(),
    readAt: new Date().toISOString(),
  });
}

function doPost(event) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    setupErpTabs_();
    const body = JSON.parse(event.postData.contents || '{}');
    if (!authorize_(event, body)) {
      return json_({ ok: false, message: 'Token invalido.' }, 401);
    }

    if (body.action === 'setup') {
      setupErpTabs_();
      appendAudit_('system', 'setup', 'Abas ERP validadas.');
      return json_({ ok: true, message: 'Setup concluido.' });
    }

    if (!body.state) {
      return json_({ ok: false, message: 'Payload sem state.' }, 400);
    }

    writeState_(body.state);
    appendAudit_(body.actor || 'netlify', 'write_state', 'Estado operacional sincronizado.');
    return json_({ ok: true, savedAt: new Date().toISOString() });
  } catch (error) {
    return json_({ ok: false, message: error.message }, 500);
  } finally {
    lock.releaseLock();
  }
}

function setupErpTabs() {
  setupErpTabs_();
  appendAudit_('manual', 'setup', 'Abas ERP validadas manualmente.');
}

function authorize_(event, body) {
  const expected = PropertiesService.getScriptProperties().getProperty('ERP_STATE_TOKEN');
  if (!expected) return true;
  const received = (body && body.token) || (event && event.parameter && event.parameter.token) || '';
  return received === expected;
}

function setupErpTabs_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(ERP_TABS).forEach((definition) => {
    let sheet = spreadsheet.getSheetByName(definition.name);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(definition.name);
    }

    let headerRange = sheet.getRange(1, 1, 1, definition.headers.length);
    let existing = headerRange.getValues()[0];
    if (definition.name === ERP_TABS.overrides.name) {
      migrateOverrideColumns_(sheet, existing);
      headerRange = sheet.getRange(1, 1, 1, definition.headers.length);
      existing = headerRange.getValues()[0];
    }
    const mustWrite = definition.headers.some((header, index) => existing[index] !== header);
    if (mustWrite) {
      headerRange.setValues([definition.headers]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#0b57b7');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  });

  seedConfig_();
  seedUsers_();
}

function migrateOverrideColumns_(sheet, existingHeaders) {
  const hasContactColumns = existingHeaders.includes('contactPhone');
  const hasOldFinancialColumns = existingHeaders.includes('enrollmentPaid');
  if (hasContactColumns || !hasOldFinancialColumns) return;

  const lastContactIndex = existingHeaders.indexOf('lastContact') + 1;
  if (lastContactIndex < 1) return;
  sheet.insertColumnsAfter(lastContactIndex, 3);
}

function seedConfig_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ERP_TABS.config.name);
  const values = rowsToObjects_(sheet.getDataRange().getValues());
  const keys = values.map((row) => row.chave);
  const defaults = [
    ['enrollmentFee', '99'],
    ['monthlyTarget', '65'],
    ['annualTarget', '780'],
    ['monthlyTicket', '299'],
    ['computersTotal', '24'],
    ['computersMaintenance', '2'],
  ];
  const rowsToAppend = defaults
    .filter(([key]) => !keys.includes(key))
    .map(([key, value]) => [key, value, new Date().toISOString()]);
  if (rowsToAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 3).setValues(rowsToAppend);
  }
}

function seedUsers_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ERP_TABS.users.name);
  const values = rowsToObjects_(sheet.getDataRange().getValues());
  if (values.length) return;

  sheet.getRange(2, 1, 4, 5).setValues([
    ['admin', 'admin', 'Administrador', 'admin', 'SIM'],
    ['financeiro', '123456', 'Responsavel Financeiro', 'financeiro', 'SIM'],
    ['retencao', '123456', 'Responsavel Retencao', 'consultor', 'SIM'],
    ['consultor', '123456', 'Consultor Comercial', 'consultor', 'SIM'],
  ]);
}

function readState_() {
  return {
    overrides: readOverrides_(),
    retention: readRetention_(),
    courses: readArray_(ERP_TABS.courses),
    leads: readArray_(ERP_TABS.leads),
    localStudents: readArray_(ERP_TABS.localStudents),
    schedule: readArray_(ERP_TABS.schedule),
    exams: readArray_(ERP_TABS.exams),
    archive: readArray_(ERP_TABS.archive),
    decisions: readArray_(ERP_TABS.decisions),
    settings: readSettings_(),
  };
}

function readUsers_() {
  return readArray_(ERP_TABS.users).map((user) => ({
    usuario: user.usuario || '',
    senha: user.senha || '',
    nome: user.nome || '',
    perfil: user.perfil || 'consultor',
    ativo: user.ativo || 'SIM',
  }));
}

function writeState_(state) {
  writeOverrides_(state.overrides || {});
  writeRetention_(state.retention || {});
  writeArray_(ERP_TABS.courses, state.courses || []);
  writeArray_(ERP_TABS.leads, state.leads || []);
  writeArray_(ERP_TABS.localStudents, state.localStudents || []);
  writeArray_(ERP_TABS.schedule, state.schedule || []);
  writeArray_(ERP_TABS.exams, state.exams || []);
  writeArray_(ERP_TABS.archive, state.archive || []);
  writeArray_(ERP_TABS.decisions, state.decisions || []);
  writeSettings_(state.settings || {});
}

function readOverrides_() {
  const rows = readArray_(ERP_TABS.overrides);
  return rows.reduce((map, row) => {
    if (!row.key) return map;
    map[row.key] = {
      status: row.status || '',
      note: row.note || '',
      lastContact: row.lastContact || '',
      contactPhone: row.contactPhone || '',
      contactEmail: row.contactEmail || '',
      followStatus: row.followStatus || '',
      enrollmentPaid: String(row.enrollmentPaid).toLowerCase() === 'true',
      enrollmentExempt: String(row.enrollmentExempt).toLowerCase() === 'true',
      boletoSent: String(row.boletoSent).toLowerCase() === 'true',
      boletoSentAt: row.boletoSentAt || '',
      enrollmentFee: Number(row.enrollmentFee || 0),
      updatedAt: row.updatedAt || '',
    };
    return map;
  }, {});
}

function writeOverrides_(overrides) {
  const rows = Object.entries(overrides).map(([key, value]) => ({
    key,
    status: value.status || '',
    note: value.note || '',
    lastContact: value.lastContact || '',
    contactPhone: value.contactPhone || '',
    contactEmail: value.contactEmail || '',
    followStatus: value.followStatus || '',
    enrollmentPaid: Boolean(value.enrollmentPaid),
    enrollmentExempt: Boolean(value.enrollmentExempt),
    boletoSent: Boolean(value.boletoSent),
    boletoSentAt: value.boletoSentAt || '',
    enrollmentFee: value.enrollmentFee || '',
    updatedAt: value.updatedAt || '',
  }));
  writeArray_(ERP_TABS.overrides, rows);
}

function readRetention_() {
  const rows = readArray_(ERP_TABS.retention);
  return rows.reduce((map, row) => {
    if (!row.key) return map;
    map[row.key] = {
      contacted: String(row.contacted).toLowerCase() === 'true',
      contactDate: row.contactDate || '',
      channel: row.channel || '',
      responsible: row.responsible || '',
      reason: row.reason || '',
      note: row.note || '',
      updatedAt: row.updatedAt || '',
    };
    return map;
  }, {});
}

function writeRetention_(retention) {
  const rows = Object.entries(retention).map(([key, value]) => ({
    key,
    contacted: Boolean(value.contacted),
    contactDate: value.contactDate || '',
    channel: value.channel || '',
    responsible: value.responsible || '',
    reason: value.reason || '',
    note: value.note || '',
    updatedAt: value.updatedAt || '',
  }));
  writeArray_(ERP_TABS.retention, rows);
}

function readSettings_() {
  return readArray_(ERP_TABS.config).reduce((settings, row) => {
    if (!row.chave) return settings;
    const value = row.valor;
    settings[row.chave] = isNaN(Number(value)) || value === '' ? value : Number(value);
    return settings;
  }, {});
}

function writeSettings_(settings) {
  const rows = Object.entries(settings).map(([chave, valor]) => ({
    chave,
    valor,
    atualizadoEm: new Date().toISOString(),
  }));
  writeArray_(ERP_TABS.config, rows);
}

function readArray_(definition) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(definition.name);
  return rowsToObjects_(sheet.getDataRange().getValues());
}

function writeArray_(definition, rows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(definition.name);
  const headers = definition.headers;
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0b57b7').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  if (!rows.length) return;

  const values = rows.map((row) => headers.map((header) => valueForSheet_(row[header])));
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  sheet.autoResizeColumns(1, headers.length);
}

function rowsToObjects_(values) {
  if (!values || values.length < 2) return [];
  const headers = values[0].map((value) => String(value || '').trim());
  return values
    .slice(1)
    .filter((row) => row.some((cell) => cell !== '' && cell !== null))
    .map((row) =>
      headers.reduce((object, header, index) => {
        object[header] = row[index];
        return object;
      }, {}),
    );
}

function valueForSheet_(value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'object') return JSON.stringify(value);
  return value;
}

function appendAudit_(actor, action, details) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ERP_TABS.audit.name);
  sheet.appendRow([new Date().toISOString(), actor, action, details]);
}

function json_(payload, status) {
  return ContentService.createTextOutput(JSON.stringify({ status: status || 200, ...payload })).setMimeType(
    ContentService.MimeType.JSON,
  );
}
