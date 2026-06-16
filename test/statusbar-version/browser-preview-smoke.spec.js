const assert = require('assert');
const fs = require('fs');
const http = require('http');
const path = require('path');
const {URL} = require('url');
const {test, expect} = require('playwright/test');

const repoRoot = path.resolve(__dirname, '../../..');

const fixtures = {
  docx: path.join(repoRoot, 'server/DocService/public/healthcheck.docx'),
  pdf: path.join(repoRoot, 'core/OfficeUtils/src/zlib-1.2.11/zlib.3.pdf'),
  xlsx: path.join(repoRoot, 'sdkjs/cell/documentation/Keyboard shortcuts.xlsx')
};

const documentServer = (process.env.DOCUMENT_SERVER_URL || '').replace(/\/+$/, '');
const requestedScenarios = (process.env.PREVIEW_SCENARIOS || 'pdf,docx,xlsx')
  .split(',')
  .map(item => item.trim())
  .filter(Boolean);
const settleMs = Number(process.env.PREVIEW_SETTLE_MS || 12000);
const requireDocumentReady = process.env.REQUIRE_DOCUMENT_READY === '1';
const fixtureBaseUrl = (process.env.FIXTURE_BASE_URL || '').replace(/\/+$/, '');
const fixtureHostname = (process.env.FIXTURE_HOSTNAME || '').trim();

if (process.env.CHROME_BIN) {
  test.use({
    launchOptions: {
      executablePath: process.env.CHROME_BIN
    }
  });
}

function contentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return {
    '.html': 'text/html; charset=utf-8',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.js': 'application/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.pdf': 'application/pdf',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  }[ext] || 'application/octet-stream';
}

function makePreviewPage(scenario, port) {
  const docsApiUrl = `${documentServer}/web-apps/apps/api/documents/api.js`;
  const localOrigin = `http://127.0.0.1:${port}`;
  const publicFixtureBaseUrl = fixtureBaseUrl || (fixtureHostname ? `http://${fixtureHostname}:${port}` : localOrigin);
  const docUrl = `${publicFixtureBaseUrl}/fixtures/${scenario}`;
  const callbackUrl = `${publicFixtureBaseUrl}/callback`;
  const documentTypes = {
    docx: 'word',
    pdf: 'pdf',
    xlsx: 'cell'
  };
  const documents = {
    docx: {
      url: docUrl,
      fileType: 'docx',
      key: `statusbar-smoke-docx-${Date.now()}`,
      title: 'healthcheck.docx',
      permissions: {edit: false}
    },
    pdf: {
      url: docUrl,
      fileType: 'pdf',
      key: `statusbar-smoke-pdf-${Date.now()}`,
      title: 'zlib.3.pdf',
      isForm: false,
      openPdfInBrowser: false,
      permissions: {edit: false}
    },
    xlsx: {
      url: docUrl,
      fileType: 'xlsx',
      key: `statusbar-smoke-xlsx-${Date.now()}`,
      title: 'Keyboard shortcuts.xlsx',
      openOfficePreviewLite: false,
      permissions: {edit: false}
    }
  };
  const doc = documents[scenario];

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>PREVIEW_PENDING</title>
  <style>html,body,#placeholder{margin:0;width:100%;height:100%;overflow:hidden;}#result{display:none;}</style>
  <script src="${docsApiUrl}"></script>
</head>
<body>
  <div id="placeholder"></div>
  <pre id="result">pending</pre>
  <script>
    (function() {
      var errors = [];
      var documentReady = false;
      var config = {
        type: 'desktop',
        width: '100%',
        height: '100%',
        isLocalFile: true,
        documentType: ${JSON.stringify(documentTypes[scenario])},
        document: ${JSON.stringify(doc)},
        editorConfig: {
          mode: 'view',
          lang: 'zh',
          callbackUrl: ${JSON.stringify(callbackUrl)},
          customization: {}
        },
        events: {
          onAppReady: function() {},
          onDocumentReady: function() {
            documentReady = true;
          },
          onError: function(event) {
            errors.push(event && event.data ? JSON.stringify(event.data) : 'onError');
          }
        }
      };

      window.__previewSmoke = {
        errors: errors,
        get documentReady() { return documentReady; }
      };

      try {
        new DocsAPI.DocEditor('placeholder', config);
      } catch (err) {
        errors.push(err && err.stack || String(err));
      }
    })();
  </script>
</body>
</html>`;
}

function startServer() {
  const server = http.createServer((req, res) => {
    const parsed = new URL(req.url, 'http://127.0.0.1');
    let filePath = null;

    if (parsed.pathname === '/callback') {
      const body = JSON.stringify({error: 0});
      res.writeHead(200, {'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body)});
      res.end(body);
      return;
    }

    if (parsed.pathname.startsWith('/preview/')) {
      const scenario = parsed.pathname.split('/').pop();
      const body = makePreviewPage(scenario, server.address().port);
      res.writeHead(200, {'Content-Type': 'text/html; charset=utf-8', 'Content-Length': Buffer.byteLength(body)});
      res.end(body);
      return;
    }

    if (parsed.pathname === '/fixtures/docx') {
      filePath = fixtures.docx;
    } else if (parsed.pathname === '/fixtures/pdf') {
      filePath = fixtures.pdf;
    } else if (parsed.pathname === '/fixtures/xlsx') {
      filePath = fixtures.xlsx;
    }

    if (!filePath || !fs.existsSync(filePath)) {
      res.writeHead(404, {'Content-Type': 'text/plain'});
      res.end('not found');
      return;
    }

    const stat = fs.statSync(filePath);
    const range = req.headers.range;
    if (range) {
      const match = /^bytes=(\d*)-(\d*)$/.exec(range);
      if (!match) {
        res.writeHead(416, {'Content-Range': `bytes */${stat.size}`});
        res.end();
        return;
      }
      const start = match[1] ? Number(match[1]) : 0;
      const end = match[2] ? Number(match[2]) : stat.size - 1;
      if (start > end || end >= stat.size) {
        res.writeHead(416, {'Content-Range': `bytes */${stat.size}`});
        res.end();
        return;
      }
      res.writeHead(206, {
        'Accept-Ranges': 'bytes',
        'Cache-Control': 'no-store',
        'Content-Length': end - start + 1,
        'Content-Range': `bytes ${start}-${end}/${stat.size}`,
        'Content-Type': contentType(filePath)
      });
      fs.createReadStream(filePath, {start, end}).pipe(res);
      return;
    }

    res.writeHead(200, {
      'Accept-Ranges': 'bytes',
      'Cache-Control': 'no-store',
      'Content-Length': stat.size,
      'Content-Type': contentType(filePath)
    });
    fs.createReadStream(filePath).pipe(res);
  });

  return new Promise(resolve => {
    server.listen(0, '0.0.0.0', () => resolve(server));
  });
}

async function collectFrameTexts(page) {
  const texts = [];
  for (const frame of page.frames()) {
    try {
      texts.push(await frame.locator('body').innerText({timeout: 1000}));
    } catch (_err) {
      // Ignore frames that are still navigating or have no body.
    }
  }
  return texts.join('\n');
}

test.describe.configure({timeout: Math.max(60000, settleMs + 30000)});

test.describe('Document preview smoke', () => {
  let server;
  let port;

  test.beforeAll(async () => {
    assert.ok(documentServer, 'DOCUMENT_SERVER_URL is required');
    for (const filePath of Object.values(fixtures)) {
      assert.ok(fs.existsSync(filePath), `fixture missing: ${filePath}`);
    }
    server = await startServer();
    port = server.address().port;
  });

  test.afterAll(async () => {
    if (server) {
      await new Promise(resolve => server.close(resolve));
    }
  });

  for (const scenario of requestedScenarios) {
    test(`${scenario} preview does not reload for editor version or crash statusbar`, async ({page}) => {
      const browserErrors = [];
      page.on('pageerror', error => {
        browserErrors.push(error && error.stack || String(error));
      });
      page.on('console', message => {
        if (message.type() === 'error') {
          browserErrors.push(message.text());
        }
      });

      await page.goto(`http://127.0.0.1:${port}/preview/${scenario}`, {
        waitUntil: 'domcontentloaded',
        timeout: 30000
      });
      await page.waitForTimeout(settleMs);

      const appErrors = await page.evaluate(() => window.__previewSmoke && window.__previewSmoke.errors || []);
      const documentReady = await page.evaluate(() => !!(window.__previewSmoke && window.__previewSmoke.documentReady));
      const frameText = await collectFrameTexts(page);
      const combined = browserErrors.concat(appErrors).join('\n') + '\n' + frameText;

      expect(combined).not.toMatch(/Cannot read properties of undefined \(reading 'on'\)/);
      expect(combined).not.toMatch(/Statusbar\.js/);
      expect(combined).not.toMatch(/编辑器已更新|Editor updated|editor version has been updated/i);
      expect(appErrors).toEqual([]);
      expect(browserErrors.filter(error => !/Failed to load resource: the server responded with a status of 404/.test(error))).toEqual([]);
      if (requireDocumentReady) {
        expect(documentReady).toBe(true);
      }
    });
  }
});
