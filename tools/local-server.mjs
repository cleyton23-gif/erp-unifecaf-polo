import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const { handler: sheetHandler } = require('../netlify/functions/sheet.js');
const { handler: stateHandler } = require('../netlify/functions/state.js');

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..', 'site');
const port = Number(process.env.PORT || 5173);
const logPath = path.resolve(__dirname, '..', 'server.log');

function log(message) {
  fs.appendFileSync(logPath, `${new Date().toISOString()} ${message}\n`);
}

process.on('uncaughtException', (error) => {
  log(`uncaughtException ${error.stack || error.message}`);
  process.exit(1);
});

process.on('unhandledRejection', (error) => {
  log(`unhandledRejection ${error.stack || error.message}`);
  process.exit(1);
});

const types = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.txt': 'text/plain; charset=utf-8',
  '.csv': 'text/csv; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
};

const server = http.createServer(async (request, response) => {
  const url = new URL(request.url, `http://${request.headers.host}`);

  if (url.pathname === '/.netlify/functions/sheet') {
    const result = await sheetHandler({
      httpMethod: request.method,
      queryStringParameters: Object.fromEntries(url.searchParams.entries()),
    });
    response.writeHead(result.statusCode, result.headers);
    response.end(result.body);
    return;
  }

  if (url.pathname === '/.netlify/functions/state') {
    const chunks = [];
    for await (const chunk of request) {
      chunks.push(chunk);
    }

    const result = await stateHandler({
      httpMethod: request.method,
      queryStringParameters: Object.fromEntries(url.searchParams.entries()),
      body: Buffer.concat(chunks).toString('utf8'),
    });
    response.writeHead(result.statusCode, result.headers);
    response.end(result.body);
    return;
  }

  const requestPath =
    url.pathname === '/' ? 'index.html' : decodeURIComponent(url.pathname).replace(/^[/\\]+/, '');
  const safePath = path.normalize(requestPath).replace(/^(\.\.[/\\])+/, '');
  const filePath = path.resolve(root, safePath);

  if (!filePath.startsWith(root)) {
    response.writeHead(403);
    response.end('Forbidden');
    return;
  }

  fs.readFile(filePath, (error, content) => {
    if (error) {
      response.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      response.end('Not found');
      return;
    }

    response.writeHead(200, {
      'Content-Type': types[path.extname(filePath)] || 'application/octet-stream',
      'Cache-Control': 'no-store',
    });
    response.end(content);
  });
});

server.on('error', (error) => {
  log(`server error ${error.stack || error.message}`);
});

server.listen(port, () => {
  const message = `Painel local em http://localhost:${port}`;
  log(message);
  console.log(message);
});
