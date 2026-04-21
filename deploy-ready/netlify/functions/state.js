const WEBAPP_URL =
  process.env.ERP_STATE_WEBAPP_URL ||
  process.env.GOOGLE_APPS_SCRIPT_URL ||
  process.env.GOOGLE_SHEETS_WEBAPP_URL ||
  '';
const STATE_TOKEN = process.env.ERP_STATE_TOKEN || '';

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return response(204, '');
  }

  if (!WEBAPP_URL) {
    return response(
      424,
      JSON.stringify({
        ok: false,
        mode: 'local',
        message: 'ERP_STATE_WEBAPP_URL nao configurada. Usando armazenamento local.',
      }),
      { 'Content-Type': 'application/json; charset=utf-8' },
    );
  }

  try {
    if (event.httpMethod === 'GET') {
      const action = event.queryStringParameters?.action || 'read';
      const readUrl = new URL(WEBAPP_URL);
      readUrl.searchParams.set('action', action);
      if (STATE_TOKEN) readUrl.searchParams.set('token', STATE_TOKEN);

      const remoteResponse = await fetch(readUrl, {
        headers: { Accept: 'application/json' },
      });
      return await forwardJson(remoteResponse);
    }

    if (event.httpMethod === 'POST') {
      const body = JSON.parse(event.body || '{}');
      if (STATE_TOKEN) body.token = STATE_TOKEN;

      const remoteResponse = await fetch(WEBAPP_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      return await forwardJson(remoteResponse);
    }

    return response(405, JSON.stringify({ ok: false, message: 'Metodo nao permitido.' }), {
      'Content-Type': 'application/json; charset=utf-8',
    });
  } catch (error) {
    return response(
      502,
      JSON.stringify({
        ok: false,
        message: `Erro ao conectar no Apps Script: ${error.message}`,
      }),
      { 'Content-Type': 'application/json; charset=utf-8' },
    );
  }
};

async function forwardJson(remoteResponse) {
  const text = await remoteResponse.text();
  return response(remoteResponse.status, text, {
    'Content-Type': remoteResponse.headers.get('content-type') || 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
  });
}

function response(statusCode, body, extraHeaders = {}) {
  return {
    statusCode,
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      ...extraHeaders,
    },
    body,
  };
}
