const ERP_TABS = {
  config: {
    name: 'ERP_Config',
    headers: ['chave', 'valor', 'atualizadoEm'],
  },
  users: {
    name: 'ERP_Usuarios',
    headers: ['usuario', 'senha', 'nome', 'perfil', 'ativo', 'lastAccess'],
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
  serviceTickets: {
    name: 'ERP_Atendimentos',
    headers: [
      'id',
      'studentKey',
      'studentName',
      'cpf',
      'ra',
      'course',
      'protocol',
      'problem',
      'requestedAt',
      'deadline',
      'attendant',
      'sector',
      'response',
      'status',
      'createdAt',
      'updatedAt',
    ],
  },
  courses: {
    name: 'ERP_Cursos',
    headers: [
      'key',
      'name',
      'modality',
      'habilitation',
      'duration',
      'monthlyFee',
      'discount10',
      'discount20',
      'discount30',
      'discount40',
      'discount50',
      'discount60',
      'authorization',
      'local',
      'updatedAt',
    ],
  },
  teachers: {
    name: 'ERP_Professores',
    headers: ['id', 'name', 'education', 'phone', 'email', 'updatedAt'],
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
      'enrollmentFee',
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
    headers: [
      'id',
      'teacher',
      'teacherPhone',
      'teacherEmail',
      'teacherEducation',
      'subject',
      'course',
      'cohortYear',
      'semester',
      'studentKeys',
      'studentCount',
      'start',
      'end',
      'room',
      'capacity',
      'status',
      'reason',
    ],
  },
  exams: {
    name: 'ERP_Provas',
    headers: ['id', 'student', 'studentKey', 'cpf', 'ra', 'course', 'discipline', 'start', 'duration', 'machines', 'status'],
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
    headers: [
      'id',
      'planType',
      'area',
      'title',
      'what',
      'why',
      'where',
      'when',
      'who',
      'how',
      'howMuch',
      'kpi',
      'periodKey',
      'weekStart',
      'owner',
      'due',
      'status',
      'createdAt',
      'updatedAt',
    ],
  },
  snapshots: {
    name: 'ERP_Snapshots',
    headers: [
      'periodKey',
      'active',
      'totalStudents',
      'highRisk',
      'noAva',
      'retention',
      'leads',
      'confirmedMatriculations',
      'pendingEnrollments',
      'debtCount',
      'debtTotal',
      'enrollmentRevenue',
      'recurringRevenue',
      'repasse',
      'qualityScore',
      'qualityIssues',
      'createdAt',
    ],
  },
  importHistory: {
    name: 'ERP_Importacoes',
    headers: ['id', 'kind', 'label', 'fileName', 'importedAt', 'recordIds'],
  },
  billing: {
    name: 'ERP_Faturamento',
    headers: [
      'id',
      'importedAt',
      'importFile',
      'competence',
      'date',
      'dueDate',
      'studentName',
      'cpf',
      'ra',
      'course',
      'installmentId',
      'amount',
      'rawJson',
    ],
  },
  receipts: {
    name: 'ERP_Recebimento',
    headers: [
      'id',
      'importedAt',
      'importFile',
      'competence',
      'date',
      'paymentDate',
      'studentName',
      'cpf',
      'ra',
      'course',
      'installmentId',
      'amount',
      'rawJson',
    ],
  },
  repasses: {
    name: 'ERP_Repasse',
    headers: [
      'id',
      'importedAt',
      'importFile',
      'competence',
      'date',
      'paymentDate',
      'description',
      'studentName',
      'cpf',
      'ra',
      'course',
      'installmentId',
      'amount',
      'rawJson',
    ],
  },
  audit: {
    name: 'ERP_Auditoria',
    headers: ['timestamp', 'actor', 'action', 'details'],
  },
  auditTrail: {
    name: 'ERP_Trilha_Local',
    headers: ['id', 'at', 'actor', 'profile', 'action', 'details'],
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

    if (body.action === 'users') {
      writeArray_(ERP_TABS.users, body.users || []);
      appendAudit_(body.actor || 'netlify', 'write_users', 'Usuarios sincronizados.');
      return json_({ ok: true, savedAt: new Date().toISOString() });
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
    if (definition.name === ERP_TABS.leads.name) {
      migrateLeadColumns_(sheet, existing);
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

function migrateLeadColumns_(sheet, existingHeaders) {
  const hasEnrollmentFee = existingHeaders.includes('enrollmentFee');
  const hasBoletoSent = existingHeaders.includes('boletoSent');
  if (hasEnrollmentFee || !hasBoletoSent) return;

  const consultantIndex = existingHeaders.indexOf('consultant') + 1;
  if (consultantIndex < 1) return;
  sheet.insertColumnsAfter(consultantIndex, 1);
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

  sheet.getRange(2, 1, 4, 6).setValues([
    ['admin', 'admin', 'Administrador', 'admin', 'SIM', ''],
    ['financeiro', '123456', 'Responsavel Financeiro', 'financeiro', 'SIM', ''],
    ['retencao', '123456', 'Responsavel Retencao', 'consultor', 'SIM', ''],
    ['consultor', '123456', 'Consultor Comercial', 'consultor', 'SIM', ''],
  ]);
}

function readState_() {
  return {
    overrides: readOverrides_(),
    retention: readRetention_(),
    serviceTickets: readArray_(ERP_TABS.serviceTickets),
    courses: readArray_(ERP_TABS.courses),
    teachers: readArray_(ERP_TABS.teachers),
    leads: readArray_(ERP_TABS.leads),
    localStudents: readArray_(ERP_TABS.localStudents),
    schedule: readArray_(ERP_TABS.schedule),
    exams: readArray_(ERP_TABS.exams),
    archive: readArray_(ERP_TABS.archive),
    decisions: readArray_(ERP_TABS.decisions),
    snapshots: readArray_(ERP_TABS.snapshots),
    taskStatus: readSettingsObject_('taskStatus'),
    importHistory: readArray_(ERP_TABS.importHistory),
    billing: readSourceOrErpRows_(ERP_TABS.billing, ['Faturamento', 'Faturamentos']),
    receipts: readSourceOrErpRows_(ERP_TABS.receipts, ['Recebimento', 'Recebimentos', 'Recebimeto']),
    repasses: readSourceOrErpRows_(ERP_TABS.repasses, ['Repasse', 'Repasses', 'Valores Repassados']),
    auditTrail: readArray_(ERP_TABS.auditTrail),
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
    lastAccess: user.lastAccess || '',
  }));
}

function writeState_(state) {
  writeOverrides_(state.overrides || {});
  writeRetention_(state.retention || {});
  writeArray_(ERP_TABS.serviceTickets, state.serviceTickets || []);
  writeArray_(ERP_TABS.courses, state.courses || []);
  writeArray_(ERP_TABS.teachers, state.teachers || []);
  writeArray_(ERP_TABS.leads, state.leads || []);
  writeArray_(ERP_TABS.localStudents, state.localStudents || []);
  writeArray_(ERP_TABS.schedule, state.schedule || []);
  writeArray_(ERP_TABS.exams, state.exams || []);
  writeArray_(ERP_TABS.archive, state.archive || []);
  writeArray_(ERP_TABS.decisions, state.decisions || []);
  writeArray_(ERP_TABS.snapshots, state.snapshots || []);
  writeSettingsObject_('taskStatus', state.taskStatus || {});
  writeArray_(ERP_TABS.importHistory, state.importHistory || []);
  // Faturamento, Recebimento e Repasse sao abas-fonte preenchidas pelo polo
  // com dados vindos da sede. O sistema le essas abas para calcular indicadores,
  // mas nao regrava o conteudo para nao apagar historico colado manualmente.
  writeArray_(ERP_TABS.auditTrail, state.auditTrail || []);
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

function readSettingsObject_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ERP_TABS.config.name);
  const rows = rowsToObjects_(sheet.getDataRange().getValues());
  const found = rows.find((row) => row.chave === key);
  if (!found || !found.valor) return {};
  try {
    return JSON.parse(found.valor);
  } catch (error) {
    return {};
  }
}

function writeSettings_(settings) {
  const previous = readArray_(ERP_TABS.config).filter((row) => row.chave === 'taskStatus');
  const rows = [
    ...Object.entries(settings).map(([chave, valor]) => ({
      chave,
      valor,
      atualizadoEm: new Date().toISOString(),
    })),
    ...previous,
  ];
  writeArray_(ERP_TABS.config, rows);
}

function writeSettingsObject_(key, value) {
  const rows = readArray_(ERP_TABS.config).filter((row) => row.chave !== key);
  rows.push({
    chave: key,
    valor: JSON.stringify(value || {}),
    atualizadoEm: new Date().toISOString(),
  });
  writeArray_(ERP_TABS.config, rows);
}

function readArray_(definition) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(definition.name);
  return rowsToObjectsFromSheet_(sheet);
}

function readSourceOrErpRows_(definition, sourceNames) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceRows = [];
  sourceNames.forEach((name) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    rowsToObjectsFromSheet_(sheet).forEach((row) => sourceRows.push(row));
  });
  if (sourceRows.length) return sourceRows;
  return readArray_(definition);
}

function rowsToObjectsFromSheet_(sheet) {
  if (!sheet) return [];
  const sheetName = sheet.getName();
  return rowsToObjects_(sheet.getDataRange().getValues()).map((row) => ({
    ...row,
    __sheetName: sheetName,
  }));
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
  const headerIndex = detectHeaderIndex_(values);
  const headers = values[headerIndex].map((value) => String(value || '').trim());
  return values
    .slice(headerIndex + 1)
    .filter((row) => row.some((cell) => cell !== '' && cell !== null))
    .map((row) =>
      headers.reduce((object, header, index) => {
        object[header] = row[index];
        return object;
      }, {}),
    );
}

function detectHeaderIndex_(values) {
  let bestIndex = 0;
  let bestScore = -1;
  let bestFilled = 0;
  values.forEach((row, index) => {
    const cells = row.map((cell) => String(cell || '').trim()).filter(Boolean);
    if (cells.length < 2) return;
    const score = cells.reduce((total, cell) => {
      const normalized = normalizeHeader_(cell);
      const directHeaders = [
        'ra',
        'ra do aluno',
        'cpf',
        'cpf do aluno',
        'nome',
        'nome do aluno',
        'curso',
        'parcela',
        'id da parcela',
        'valor',
        'valor pago',
        'valor faturado',
        'data pagamento',
        'data faturado',
        'repasse',
        'repasse final',
        'total recebido',
        'total faturado',
      ];
      if (directHeaders.indexOf(normalized) !== -1) return total + 3;
      if (
        normalized.indexOf('aluno') !== -1 ||
        normalized.indexOf('curso') !== -1 ||
        normalized.indexOf('competencia') !== -1 ||
        normalized.indexOf('periodo') !== -1 ||
        normalized.indexOf('receb') !== -1 ||
        normalized.indexOf('fatur') !== -1 ||
        normalized.indexOf('repasse') !== -1
      ) {
        return total + 1;
      }
      return total;
    }, 0);
    if (score > bestScore || (score === bestScore && cells.length > bestFilled)) {
      bestIndex = index;
      bestScore = score;
      bestFilled = cells.length;
    }
  });
  return bestScore >= 3 ? bestIndex : 0;
}

function normalizeHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
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
