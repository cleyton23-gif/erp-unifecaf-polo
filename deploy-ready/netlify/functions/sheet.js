const DEFAULT_SHEET_ID = '1AoZ9KCNIaIzTEW17MyNFv5O7lVmcoaa8_bFKI8dBY8Q';

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return response(204, '');
  }

  const sheetId =
    process.env.GOOGLE_SHEET_ID ||
    event.queryStringParameters?.sheetId ||
    DEFAULT_SHEET_ID;
  const gid = process.env.GOOGLE_SHEET_GID || event.queryStringParameters?.gid || '';
  let url = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(
    sheetId,
  )}/export?format=csv`;

  if (gid) {
    url += `&gid=${encodeURIComponent(gid)}`;
  }

  try {
    const sheetResponse = await fetch(url, {
      headers: {
        'User-Agent': 'Netlify UniFECAF Dashboard',
      },
    });

    if (!sheetResponse.ok) {
      return response(
        sheetResponse.status,
        `Nao foi possivel carregar a planilha (${sheetResponse.status}).`,
      );
    }

    const csv = await sheetResponse.text();

    return response(200, csv, {
      'Content-Type': 'text/csv; charset=utf-8',
      'Cache-Control': 'public, s-maxage=300, stale-while-revalidate=600',
    });
  } catch (error) {
    return response(502, `Erro ao conectar na planilha: ${error.message}`);
  }
};

function response(statusCode, body, extraHeaders = {}) {
  return {
    statusCode,
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      ...extraHeaders,
    },
    body,
  };
}
