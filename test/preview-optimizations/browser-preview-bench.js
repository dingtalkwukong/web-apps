#!/usr/bin/env node

'use strict';

const assert = require('assert');
const childProcess = require('child_process');
const fs = require('fs');
const http = require('http');
const os = require('os');
const path = require('path');
const {URL} = require('url');

const webAppsRoot = path.resolve(__dirname, '../..');
const repoRoot = path.resolve(webAppsRoot, '..');
const fixtures = {
  pdf: path.join(repoRoot, 'core/OfficeUtils/src/zlib-1.2.11/zlib.3.pdf'),
  docx: path.join(repoRoot, 'server/DocService/public/healthcheck.docx'),
  xlsx: path.join(repoRoot, 'sdkjs/cell/documentation/Keyboard shortcuts.xlsx'),
  pptx: path.join(repoRoot, 'sdkjs/slide/themes/src/01_blank.pptx')
};

const scenarios = [
  'pdf-native',
  'pdf-fillforms-mode-native',
  'pdf-fallback',
  'office-lite',
  'office-full',
  'office-lite-xlsx',
  'office-full-xlsx',
  'office-lite-pptx',
  'office-full-pptx'
];

const officeFixtures = {
  docx: {
    documentType: 'word',
    fileType: 'docx',
    fixture: 'docx',
    title: 'healthcheck.docx'
  },
  xlsx: {
    documentType: 'cell',
    fileType: 'xlsx',
    fixture: 'xlsx',
    title: 'Keyboard shortcuts.xlsx'
  },
  pptx: {
    documentType: 'slide',
    fileType: 'pptx',
    fixture: 'pptx',
    title: '01_blank.pptx'
  }
};

function parseArgs(argv) {
  const options = {
    documentServer: process.env.DOCUMENT_SERVER_URL || '',
    requireOfficeDocumentReady: false,
    ignoreEditorVersionCheck: false,
    officeTimeoutMs: 20000
  };

  for (const arg of argv) {
    if (arg.startsWith('--document-server=')) {
      options.documentServer = arg.substring('--document-server='.length);
    } else if (arg === '--require-office-document-ready') {
      options.requireOfficeDocumentReady = true;
    } else if (arg === '--ignore-editor-version-check') {
      options.ignoreEditorVersionCheck = true;
    } else if (arg.startsWith('--office-timeout-ms=')) {
      const timeoutMs = Number(arg.substring('--office-timeout-ms='.length));
      if (Number.isFinite(timeoutMs) && timeoutMs > 0) {
        options.officeTimeoutMs = timeoutMs;
      }
    }
  }

  if (options.documentServer) {
    options.documentServer = options.documentServer.replace(/\/+$/, '');
  }
  return options;
}

function findChrome() {
  const candidates = [
    process.env.CHROME_BIN,
    'chromium',
    'chromium-browser',
    'google-chrome',
    'google-chrome-stable'
  ].filter(Boolean);

  for (const candidate of candidates) {
    const result = childProcess.spawnSync(candidate, ['--version'], {encoding: 'utf8'});
    if (!result.error && result.status === 0) {
      if (path.isAbsolute(candidate)) {
        return candidate;
      }
      const which = childProcess.spawnSync('command', ['-v', candidate], {encoding: 'utf8', shell: true});
      return which.stdout.trim() || candidate;
    }
  }
  throw new Error('Chromium/Chrome is required. Set CHROME_BIN or install chromium.');
}

function requirePuppeteer() {
  try {
    return require('puppeteer-core');
  } catch (_err) {
    throw new Error('puppeteer-core is required. Run: npx -y -p puppeteer-core node test/preview-optimizations/browser-preview-bench.js');
  }
}

function contentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return {
    '.css': 'text/css; charset=utf-8',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.gif': 'image/gif',
    '.html': 'text/html; charset=utf-8',
    '.ico': 'image/x-icon',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.js': 'application/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.less': 'text/css; charset=utf-8',
    '.map': 'application/json; charset=utf-8',
    '.pdf': 'application/pdf',
    '.png': 'image/png',
    '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    '.svg': 'image/svg+xml',
    '.ttf': 'font/ttf',
    '.woff': 'font/woff',
    '.woff2': 'font/woff2',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  }[ext] || 'application/octet-stream';
}

function parseRange(header, totalLength) {
  if (!header) {
    return null;
  }

  const match = /^bytes=(\d*)-(\d*)$/.exec(header);
  if (!match || match[1] === '' && match[2] === '') {
    return {invalid: true};
  }

  let start;
  let end;
  if (match[1] === '') {
    const suffixLength = parseInt(match[2], 10);
    if (!Number.isFinite(suffixLength) || suffixLength <= 0) {
      return {invalid: true};
    }
    start = Math.max(totalLength - suffixLength, 0);
    end = totalLength - 1;
  } else {
    start = parseInt(match[1], 10);
    end = match[2] === '' ? totalLength - 1 : parseInt(match[2], 10);
  }

  if (!Number.isFinite(start) || !Number.isFinite(end) || start < 0 || end < start || start >= totalLength) {
    return {invalid: true};
  }

  end = Math.min(end, totalLength - 1);
  return {start, end, contentLength: end - start + 1};
}

function safeJoin(root, requestPath) {
  const decodedPath = decodeURIComponent(requestPath);
  const fullPath = path.resolve(root, decodedPath.replace(/^\/+/, ''));
  if (!fullPath.startsWith(root + path.sep) && fullPath !== root) {
    return null;
  }
  return fullPath;
}

function makeBenchPage(scenario, docsApiUrl, options) {
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>preview benchmark</title>
  <style>html,body,#placeholder{margin:0;width:100%;height:100%;}#result{white-space:pre-wrap;font:12px monospace;}</style>
  <script src="${docsApiUrl}"></script>
</head>
<body>
  <div id="placeholder"></div>
  <pre id="result">pending</pre>
  <script>
	  (function() {
	    var scenario = ${JSON.stringify(scenario)};
	    var requireOfficeDocumentReady = ${JSON.stringify(options.requireOfficeDocumentReady)};
	    var officeTimeoutMs = ${JSON.stringify(options.officeTimeoutMs)};
	    var started = performance.now();
    var errors = [];
    var iframeLoad = null;
    var iframe = null;
    var finished = false;
    var settleMs = scenario.indexOf('office-') === 0 ? 3500 : 900;
	    var timeoutMs = scenario.indexOf('office-') === 0 ? officeTimeoutMs : 5000;

    function makeUrl(kind) {
      return window.location.origin + '/fixtures/' + kind + '?scenario=' + encodeURIComponent(scenario) + '&cache=' + Date.now();
    }

    function attrs(frame) {
      if (!frame) return {};
      return {
        nativePdf: frame.getAttribute('data-onlyoffice-native-pdf-preview'),
        nativePdfLegacy: frame.getAttribute('data-native-pdf-preview'),
        previewLite: frame.getAttribute('data-onlyoffice-preview-lite'),
        previewLiteLegacy: frame.getAttribute('data-office-preview-lite')
      };
    }

    function finish(reason) {
      if (finished) return;
      finished = true;
      var frame = iframe || document.querySelector('iframe');
      var result = {
        scenario: scenario,
        reason: reason,
        elapsedMs: Math.round(performance.now() - started),
        iframeLoadMs: iframeLoad == null ? null : Math.round(iframeLoad - started),
        iframePresent: !!frame,
        iframeSrc: frame ? frame.src : null,
        attrs: attrs(frame),
        errors: errors
      };
      document.getElementById('result').textContent = 'BENCH_RESULT ' + JSON.stringify(result);
      document.title = 'BENCH_DONE';
    }

    function watchIframe(frame) {
      if (!frame || frame.__benchWatched) return;
      frame.__benchWatched = true;
      iframe = frame;
	      frame.addEventListener('load', function() {
	        iframeLoad = performance.now();
	        if (!(requireOfficeDocumentReady && scenario.indexOf('office-') === 0)) {
	          setTimeout(function() {
	            finish('iframe-load');
	          }, settleMs);
	        }
	      }, {once: true});
    }

    var observer = new MutationObserver(function() {
      watchIframe(document.querySelector('iframe'));
    });
    observer.observe(document.body, {childList: true, subtree: true});

    window.addEventListener('error', function(event) {
      errors.push(String(event.message || event.error || 'error'));
    });

    var isOffice = scenario.indexOf('office-') === 0;
    var isFull = /fallback|full/.test(scenario);
    var officeKind = /xlsx$/.test(scenario) ? 'xlsx' : (/pptx$/.test(scenario) ? 'pptx' : 'docx');
    var officeFixtures = ${JSON.stringify(officeFixtures)};
    var officeFixture = officeFixtures[officeKind];
    var doc = isOffice ? {
      url: makeUrl(officeFixture.fixture),
      fileType: officeFixture.fileType,
      key: 'bench-' + officeKind + '-' + scenario,
      title: officeFixture.title,
      permissions: {edit: false}
    } : {
      url: makeUrl('pdf'),
      fileType: 'pdf',
      key: 'bench-pdf-' + scenario,
      title: 'zlib.3.pdf',
      isForm: false,
      permissions: {edit: false}
    };

    if (scenario === 'pdf-fallback') {
      doc.openPdfInBrowser = false;
    }
    if (scenario === 'pdf-fillforms-mode-native') {
      doc.permissions.fillForms = true;
    }
    if (isOffice && isFull) {
      doc.openOfficePreviewLite = false;
    }

    var config = {
      type: 'desktop',
      width: '900px',
      height: '700px',
      documentType: isOffice ? officeFixture.documentType : 'pdf',
      document: doc,
      editorConfig: {
        mode: scenario === 'pdf-fillforms-mode-native' ? 'fillforms' : 'view',
        lang: 'en',
        customization: {}
      },
      events: {
        onAppReady: function() {},
        onDocumentReady: function() {
          setTimeout(function() {
            finish('document-ready');
          }, settleMs);
        },
        onError: function(event) {
          errors.push(event && event.data ? JSON.stringify(event.data) : 'onError');
        }
      }
    };

    try {
      new DocsAPI.DocEditor('placeholder', config);
      watchIframe(document.querySelector('iframe'));
    } catch (err) {
      errors.push(err && err.stack || String(err));
      finish('exception');
      return;
    }

    setTimeout(function() {
      finish('timeout');
    }, timeoutMs);
  })();
  </script>
</body>
</html>`;
}

function startServer(options) {
  const metrics = {
    requests: []
  };
  function serveInline(req, res, body, type, kind) {
    const bytes = Buffer.byteLength(body);
    metrics.requests.push({url: req.url, status: 200, bytes, kind});
    res.writeHead(200, {
      'Content-Type': type,
      'Content-Length': bytes,
      'Cache-Control': 'no-store'
    });
    res.end(body);
  }

  const server = http.createServer((req, res) => {
    const parsed = new URL(req.url, 'http://127.0.0.1');
    const pathname = parsed.pathname;

    if (pathname === '/favicon.ico') {
      res.statusCode = 204;
      res.end();
      return;
    }

    if (pathname === '/bench') {
      const body = makeBenchPage(parsed.searchParams.get('scenario'), options.docsApiUrl, options);
      metrics.requests.push({url: req.url, status: 200, bytes: Buffer.byteLength(body), kind: 'page'});
      res.writeHead(200, {'Content-Type': 'text/html; charset=utf-8', 'Content-Length': Buffer.byteLength(body)});
      res.end(body);
      return;
    }

    if (/^\/sdkjs\/develop\/sdkjs\/(?:word|cell|slide)\/scripts\.js$/.test(pathname)) {
      serveInline(req, res, '"use strict";\n', 'application/javascript; charset=utf-8', '.js');
      return;
    }

    if (/^\/apps\/(?:documenteditor|spreadsheeteditor|presentationeditor|pdfeditor)\/main\/resources\/less\/sprites\/.+\.less$/.test(pathname)) {
      serveInline(req, res, '/* generated sprite placeholder for source benchmark */\n', 'text/css; charset=utf-8', '.less');
      return;
    }

    let filePath = null;
    if (pathname === '/fixtures/pdf') filePath = fixtures.pdf;
    else if (pathname === '/fixtures/docx') filePath = fixtures.docx;
    else if (pathname === '/fixtures/xlsx') filePath = fixtures.xlsx;
    else if (pathname === '/fixtures/pptx') filePath = fixtures.pptx;
    else if (pathname === '/document_editor_service_worker.js') filePath = path.join(repoRoot, 'sdkjs/common/serviceworker/document_editor_service_worker.js');
    else if (pathname.startsWith('/apps/')) filePath = safeJoin(path.join(webAppsRoot, 'apps'), pathname.substring('/apps/'.length));
    else if (pathname.startsWith('/vendor/')) filePath = safeJoin(path.join(webAppsRoot, 'vendor'), pathname.substring('/vendor/'.length));
    else if (pathname.startsWith('/sdkjs/')) filePath = safeJoin(path.join(repoRoot, 'sdkjs'), pathname.substring('/sdkjs/'.length));
    else if (pathname.startsWith('/web-apps/')) filePath = safeJoin(webAppsRoot, pathname.substring('/web-apps/'.length));

    if (!filePath || !fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
      const body = 'not found';
      metrics.requests.push({url: req.url, status: 404, bytes: body.length, kind: 'missing'});
      res.writeHead(404, {'Content-Type': 'text/plain', 'Content-Length': body.length});
      res.end(body);
      return;
    }

    const stat = fs.statSync(filePath);
    const range = parseRange(req.headers.range, stat.size);
    if (range && range.invalid) {
      metrics.requests.push({url: req.url, status: 416, bytes: 0, kind: 'range'});
      res.writeHead(416, {
        'Accept-Ranges': 'bytes',
        'Content-Range': 'bytes */' + stat.size
      });
      res.end();
      return;
    }

    const status = range ? 206 : 200;
    const bytes = range ? range.contentLength : stat.size;
    const headers = {
      'Accept-Ranges': 'bytes',
      'Content-Type': contentType(filePath),
      'Content-Length': bytes,
      'Cache-Control': 'no-store'
    };
    if (range) {
      headers['Content-Range'] = 'bytes ' + range.start + '-' + range.end + '/' + stat.size;
    }

    metrics.requests.push({url: req.url, status, bytes, kind: path.extname(filePath).toLowerCase() || 'file'});
    res.writeHead(status, headers);
    fs.createReadStream(filePath, range ? {start: range.start, end: range.end} : undefined).pipe(res);
  });

  return new Promise(resolve => {
    server.listen(0, '127.0.0.1', () => {
      resolve({server, metrics, port: server.address().port});
    });
  });
}

function requestPath(url) {
  try {
    return new URL(url, 'http://127.0.0.1').pathname;
  } catch (_err) {
    return url;
  }
}

function classifyRequest(url, mimeType) {
  const pathname = requestPath(url);
  const ext = path.extname(pathname).toLowerCase();
  if (ext) {
    return ext;
  }
  if (/javascript/.test(mimeType)) {
    return '.js';
  }
  if (/css/.test(mimeType)) {
    return '.css';
  }
  return 'file';
}

function summarizeRequests(requests) {
  const summary = {
    requestCount: requests.length,
    totalBytes: 0,
    jsCssBytes: 0,
    jsCssRequests: 0,
    sdkJsCssBytes: 0,
    sdkJsCssRequests: 0,
    nonSdkJsCssBytes: 0,
    nonSdkJsCssRequests: 0,
    fixtureBytes: 0,
    fixtureRequests: 0,
    missingRequests: 0,
    partialContentRequests: 0
  };

  for (const request of requests) {
    const pathname = requestPath(request.url);

    summary.totalBytes += request.bytes;
    if (request.kind === '.js' || request.kind === '.css' || request.kind === '.less') {
      summary.jsCssBytes += request.bytes;
      summary.jsCssRequests += 1;
      if (pathname.startsWith('/sdkjs/')) {
        summary.sdkJsCssBytes += request.bytes;
        summary.sdkJsCssRequests += 1;
      } else {
        summary.nonSdkJsCssBytes += request.bytes;
        summary.nonSdkJsCssRequests += 1;
      }
    }
    if (pathname.startsWith('/fixtures/')) {
      summary.fixtureBytes += request.bytes;
      summary.fixtureRequests += 1;
    }
    if (request.status === 404) {
      summary.missingRequests += 1;
    }
    if (request.status === 206) {
      summary.partialContentRequests += 1;
    }
  }
  return summary;
}

function shouldPatchEditorVersionCheck(url) {
  return /\/web-apps\/apps\/(?:documenteditor|spreadsheeteditor|presentationeditor)\/main\/index\.html(?:[?#]|$)/.test(url);
}

async function installRequestPatches(page, options) {
  if (!options.ignoreEditorVersionCheck) {
    return;
  }

  await page.setRequestInterception(true);
  page.on('request', async request => {
    if (!shouldPatchEditorVersionCheck(request.url())) {
      await request.continue();
      return;
    }

    try {
      const response = await fetch(request.url());
      let body = await response.text();
      body = body.replace(/<head([^>]*)>/i, '<head$1><script>window.compareVersions=true;</script>');

      const headers = {};
      response.headers.forEach((value, key) => {
        const lower = key.toLowerCase();
        if (lower !== 'content-encoding' && lower !== 'content-length' && lower !== 'transfer-encoding') {
          headers[key] = value;
        }
      });
      headers['content-type'] = headers['content-type'] || 'text/html; charset=utf-8';
      headers['content-length'] = String(Buffer.byteLength(body));

      await request.respond({
        status: response.status,
        headers,
        body
      });
    } catch (_err) {
      await request.continue();
    }
  });
}

async function runChrome(chrome, url, waitTimeoutMs, options) {
  const puppeteer = requirePuppeteer();
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'oo-preview-bench-'));
  let browser;
  const runtimeErrors = [];

  try {
    browser = await puppeteer.launch({
      executablePath: chrome,
      headless: true,
      userDataDir: profile,
      args: [
        '--no-sandbox',
        '--disable-gpu',
        '--disable-dev-shm-usage',
        '--disable-background-networking',
        '--disable-extensions',
        '--disable-features=BlockInsecurePrivateNetworkRequests,PrivateNetworkAccessRespectPreflightResults,PrivateNetworkAccessSendPreflights'
      ]
    });
    const page = await browser.newPage();
    page.on('pageerror', error => {
      runtimeErrors.push(error && error.stack || String(error));
    });
    page.on('console', message => {
      if (message.type() === 'error') {
        runtimeErrors.push(message.text());
      }
    });
    await installRequestPatches(page, options);
    const client = await page.target().createCDPSession();
    const requests = new Map();

    await client.send('Network.enable');
    client.on('Network.responseReceived', event => {
      const headers = event.response.headers || {};
      const normalizedHeaders = Object.fromEntries(Object.keys(headers).map(key => [key.toLowerCase(), headers[key]]));
      const contentLength = Number(normalizedHeaders['content-length']);
      requests.set(event.requestId, {
        url: event.response.url,
        status: event.response.status,
        bytes: Number.isFinite(contentLength) && contentLength > 0 ? contentLength : 0,
        kind: classifyRequest(event.response.url, event.response.mimeType || '')
      });
    });
    client.on('Network.loadingFinished', event => {
      const request = requests.get(event.requestId);
      if (request && !request.bytes) {
        request.bytes = event.encodedDataLength || 0;
      }
    });

    await page.goto(url, {waitUntil: 'domcontentloaded', timeout: waitTimeoutMs});
    await page.waitForFunction(() => document.title === 'BENCH_DONE', {timeout: waitTimeoutMs});
    await new Promise(resolve => setTimeout(resolve, 100));
    return {
      dom: await page.$eval('#result', element => element.textContent),
      requests: Array.from(requests.values()),
      runtimeErrors
    };
  } finally {
    if (browser) {
      await browser.close();
    }
    fs.rmSync(profile, {recursive: true, force: true});
  }
}

function extractResult(dom) {
  const match = /BENCH_RESULT\s+({[\s\S]+})/.exec(dom);
  if (!match) {
    throw new Error('Benchmark page did not finish. DOM tail:\n' + dom.slice(-2000));
  }
  return JSON.parse(match[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&'));
}

async function runScenario(chrome, port, scenario, options) {
  const waitTimeoutMs = scenario.indexOf('office-') === 0 ?
    Math.max(20000, options.officeTimeoutMs + 5000) :
    20000;
  const {dom, requests, runtimeErrors} = await runChrome(chrome, `http://127.0.0.1:${port}/bench?scenario=${encodeURIComponent(scenario)}`, waitTimeoutMs, options);
  const result = extractResult(dom);
  const summary = summarizeRequests(requests);
  return {scenario, result, summary, requests, runtimeErrors};
}

function percentReduction(base, optimized) {
  if (!base) return 0;
  return Math.round((1 - optimized / base) * 100);
}

function isIgnorableRuntimeError(error) {
  return /WebSocket connection to .*ERR_BLOCKED_BY_LOCAL_NETWORK_ACCESS_CHECKS/.test(error) ||
    /^Failed to load resource: the server responded with a status of 404 \(Not Found\)$/.test(error) ||
    /\/sdkjs\/slide\/themes\/+themes\.js['" ]/.test(error);
}

function isIgnorableMissingRequest(url) {
  return /\/sdkjs\/slide\/themes\/+themes\.js(?:[?#]|$)/.test(url);
}

function assertBench(results, options) {
  const byName = Object.fromEntries(results.map(item => [item.scenario, item]));
  const pdfNative = byName['pdf-native'];
  const pdfFillformsModeNative = byName['pdf-fillforms-mode-native'];
  const pdfFallback = byName['pdf-fallback'];
  const officePairs = [
    ['office-lite', 'office-full'],
    ['office-lite-xlsx', 'office-full-xlsx'],
    ['office-lite-pptx', 'office-full-pptx']
  ];
  const assertRuntimeErrors = !!options.documentServer || options.requireOfficeDocumentReady;

  for (const item of results) {
    assert.deepStrictEqual(item.result.errors, [], `${item.scenario} reported page errors`);
    if (assertRuntimeErrors) {
      const criticalRuntimeErrors = item.runtimeErrors.filter(error => !isIgnorableRuntimeError(error));
      assert.deepStrictEqual(criticalRuntimeErrors, [], `${item.scenario} reported browser runtime errors`);
    }
    const missingRequests = item.requests.filter(request => request.status === 404 && !isIgnorableMissingRequest(request.url));
    assert.strictEqual(missingRequests.length, 0, `${item.scenario} has missing resource requests`);
  }

  assert.strictEqual(pdfNative.result.attrs.nativePdf, 'true', 'native PDF iframe marker missing');
  assert.ok(pdfNative.result.iframeSrc.includes('/fixtures/pdf'), 'native PDF should load the PDF URL directly');
  assert.strictEqual(pdfFillformsModeNative.result.attrs.nativePdf, 'true', 'fillforms-mode native PDF iframe marker missing');
  assert.ok(pdfFillformsModeNative.result.iframeSrc.includes('/fixtures/pdf'), 'fillforms-mode native PDF should load the PDF URL directly');
  assert.notStrictEqual(pdfFallback.result.attrs.nativePdf, 'true', 'PDF fallback must not use native iframe');
  assert.ok(/\/(?:web-apps\/)?apps\/pdfeditor\/main\/index\.html/.test(pdfFallback.result.iframeSrc), 'PDF fallback should load the PDF editor');
  for (const [liteName, fullName] of officePairs) {
    const officeLite = byName[liteName];
    const officeFull = byName[fullName];
    assert.strictEqual(officeLite.result.attrs.previewLite, 'true', `${liteName} iframe marker missing`);
    assert.ok(/[?&]previewLite=1(?:&|$)/.test(officeLite.result.iframeSrc), `${liteName} URL marker missing`);
    assert.notStrictEqual(officeFull.result.attrs.previewLite, 'true', `${fullName} must not use previewLite marker`);
    assert.ok(!/[?&]previewLite=1(?:&|$)/.test(officeFull.result.iframeSrc), `${fullName} must not include previewLite=1`);
    if (options.requireOfficeDocumentReady) {
      assert.strictEqual(officeLite.result.reason, 'document-ready', `${liteName} did not reach document-ready`);
      assert.strictEqual(officeFull.result.reason, 'document-ready', `${fullName} did not reach document-ready`);
    }
  }

  const pdfByteReduction = percentReduction(pdfFallback.summary.totalBytes, pdfNative.summary.totalBytes);
  const pdfRequestReduction = percentReduction(pdfFallback.summary.requestCount, pdfNative.summary.requestCount);
  assert.ok(
    pdfByteReduction >= 40 || pdfRequestReduction >= 40,
    `native PDF preview improvement is too small: bytes ${pdfByteReduction}%, requests ${pdfRequestReduction}%`
  );

  for (const [liteName, fullName] of officePairs) {
    const officeLite = byName[liteName];
    const officeFull = byName[fullName];
    const officeJsReduction = percentReduction(officeFull.summary.nonSdkJsCssBytes, officeLite.summary.nonSdkJsCssBytes);
    const officeRequestReduction = percentReduction(officeFull.summary.nonSdkJsCssRequests, officeLite.summary.nonSdkJsCssRequests);
    assert.ok(
      officeJsReduction >= 10 || officeRequestReduction >= 10,
      `${liteName} improvement is too small: non-SDK js/css bytes ${officeJsReduction}%, non-SDK js/css requests ${officeRequestReduction}%`
    );
  }
}

(async function main() {
  for (const filePath of Object.values(fixtures)) {
    assert.ok(fs.existsSync(filePath), `fixture missing: ${filePath}`);
  }

	  const chrome = findChrome();
  const options = parseArgs(process.argv.slice(2));
  const docsApiUrl = options.documentServer ?
    options.documentServer + '/web-apps/apps/api/documents/api.js' :
    '/apps/api/documents/api.js';
  const {server, port} = await startServer(Object.assign({}, options, {docsApiUrl}));
  try {
    const results = [];
    for (const scenario of scenarios) {
      results.push(await runScenario(chrome, port, scenario, options));
    }
    const printable = results.map(item => ({
      scenario: item.scenario,
      reason: item.result.reason,
      elapsedMs: item.result.elapsedMs,
      iframeLoadMs: item.result.iframeLoadMs,
      requestCount: item.summary.requestCount,
      totalBytes: item.summary.totalBytes,
      jsCssRequests: item.summary.jsCssRequests,
      jsCssBytes: item.summary.jsCssBytes,
      sdkJsCssRequests: item.summary.sdkJsCssRequests,
      sdkJsCssBytes: item.summary.sdkJsCssBytes,
      nonSdkJsCssRequests: item.summary.nonSdkJsCssRequests,
      nonSdkJsCssBytes: item.summary.nonSdkJsCssBytes,
      fixtureRequests: item.summary.fixtureRequests,
      fixtureBytes: item.summary.fixtureBytes,
	      partialContentRequests: item.summary.partialContentRequests,
	      missingRequests: item.requests.filter(request => request.status === 404 && !isIgnorableMissingRequest(request.url)).length,
	      missingUrls: item.requests.filter(request => request.status === 404 && !isIgnorableMissingRequest(request.url)).slice(0, 10).map(request => request.url),
	      ignorableMissingRequests: item.requests.filter(request => request.status === 404 && isIgnorableMissingRequest(request.url)).length,
	      errors: item.result.errors,
	      runtimeErrors: item.runtimeErrors,
	      criticalRuntimeErrors: item.runtimeErrors.filter(error => !isIgnorableRuntimeError(error)),
	      iframeSrc: item.result.iframeSrc
	    }));
    const byName = Object.fromEntries(results.map(item => [item.scenario, item]));
    const output = {
      chrome,
      documentServer: options.documentServer || null,
      results: printable,
      pdfTotalByteReductionPercent: percentReduction(
        byName['pdf-fallback'].summary.totalBytes,
        byName['pdf-native'].summary.totalBytes
      ),
    officeJsCssByteReductionPercent: percentReduction(byName['office-full'].summary.jsCssBytes, byName['office-lite'].summary.jsCssBytes),
    officeNonSdkJsCssByteReductionPercent: percentReduction(byName['office-full'].summary.nonSdkJsCssBytes, byName['office-lite'].summary.nonSdkJsCssBytes),
    officeReductions: {
      docx: {
        nonSdkJsCssByteReductionPercent: percentReduction(byName['office-full'].summary.nonSdkJsCssBytes, byName['office-lite'].summary.nonSdkJsCssBytes),
        nonSdkJsCssRequestReductionPercent: percentReduction(byName['office-full'].summary.nonSdkJsCssRequests, byName['office-lite'].summary.nonSdkJsCssRequests)
      },
      xlsx: {
        nonSdkJsCssByteReductionPercent: percentReduction(byName['office-full-xlsx'].summary.nonSdkJsCssBytes, byName['office-lite-xlsx'].summary.nonSdkJsCssBytes),
        nonSdkJsCssRequestReductionPercent: percentReduction(byName['office-full-xlsx'].summary.nonSdkJsCssRequests, byName['office-lite-xlsx'].summary.nonSdkJsCssRequests)
      },
      pptx: {
        nonSdkJsCssByteReductionPercent: percentReduction(byName['office-full-pptx'].summary.nonSdkJsCssBytes, byName['office-lite-pptx'].summary.nonSdkJsCssBytes),
        nonSdkJsCssRequestReductionPercent: percentReduction(byName['office-full-pptx'].summary.nonSdkJsCssRequests, byName['office-lite-pptx'].summary.nonSdkJsCssRequests)
      }
    }
  };
    console.log(JSON.stringify(output, null, 2));
	    assertBench(results, options);
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
})().catch(err => {
  console.error(err && err.stack ? err.stack : err);
  process.exit(1);
});
