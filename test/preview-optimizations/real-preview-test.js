#!/usr/bin/env node

'use strict';

const assert = require('assert');
const childProcess = require('child_process');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const {URL} = require('url');

const TEST_NAME = '真实预览测试';
const webAppsRoot = path.resolve(__dirname, '../..');
const repoRoot = path.resolve(webAppsRoot, '..');
const apiPath = path.join(webAppsRoot, 'apps/api/documents/api.js');
const officeDocumentTypes = {
  doc: 'word',
  docx: 'word',
  odt: 'word',
  rtf: 'word',
  txt: 'word',
  xls: 'cell',
  xlsx: 'cell',
  ods: 'cell',
  csv: 'cell',
  ppt: 'slide',
  pptx: 'slide',
  odp: 'slide'
};

function parseArgs(argv) {
  const options = {
    logFiles: [],
    ssh: '',
    sshPasswordEnv: 'SSHPASS',
    remoteLogs: [],
    remoteLines: 30000,
    maxCasesPerType: 8,
    json: false,
    analyzeOnly: false,
    showSensitive: false
  };

  for (const arg of argv) {
    if (arg.startsWith('--log-file=')) {
      options.logFiles.push(arg.substring('--log-file='.length));
    } else if (arg.startsWith('--ssh=')) {
      options.ssh = arg.substring('--ssh='.length);
    } else if (arg.startsWith('--ssh-password-env=')) {
      options.sshPasswordEnv = arg.substring('--ssh-password-env='.length);
    } else if (arg.startsWith('--remote-log=')) {
      options.remoteLogs.push(arg.substring('--remote-log='.length));
    } else if (arg.startsWith('--remote-lines=')) {
      const lines = Number(arg.substring('--remote-lines='.length));
      if (Number.isFinite(lines) && lines > 0) {
        options.remoteLines = Math.floor(lines);
      }
    } else if (arg.startsWith('--max-cases-per-type=')) {
      const count = Number(arg.substring('--max-cases-per-type='.length));
      if (Number.isFinite(count) && count > 0) {
        options.maxCasesPerType = Math.floor(count);
      }
    } else if (arg === '--json') {
      options.json = true;
    } else if (arg === '--analyze-only') {
      options.analyzeOnly = true;
    } else if (arg === '--show-sensitive') {
      options.showSensitive = true;
    } else if (arg === '--help' || arg === '-h') {
      usage();
      process.exit(0);
    } else {
      throw new Error('Unknown argument: ' + arg);
    }
  }

  return options;
}

function usage() {
  console.log([
    `${TEST_NAME}`,
    '',
    'Usage:',
    '  node web-apps/test/preview-optimizations/real-preview-test.js --log-file=/path/access.log',
    '  node web-apps/test/preview-optimizations/real-preview-test.js --ssh=user@host --remote-log=/path/access.log',
    '',
    'Remote mode is read-only: it tails compose-mounted access logs and does not write to the server.',
    '',
    'Options:',
    '  --ssh=user@host              Read logs over SSH.',
    '  --ssh-password-env=NAME      Password env var used by sshpass, default SSHPASS.',
    '  --remote-log=PATH            Remote log path. Can be repeated.',
    '  --remote-lines=N             Tail N lines from each remote log, default 30000.',
    '  --log-file=PATH              Local log file. Can be repeated.',
    '  --max-cases-per-type=N       Max replay cases per extension, default 8.',
    '  --analyze-only               Parse logs without local routing assertions.',
    '  --show-sensitive             Include raw source paths, file names, and URLs in output.',
    '  --json                       Print JSON only.'
  ].join('\n'));
}

function shellQuote(value) {
  return "'" + String(value).replace(/'/g, "'\\''") + "'";
}

function commandExists(command) {
  const result = childProcess.spawnSync('sh', ['-lc', 'command -v ' + shellQuote(command)], {encoding: 'utf8'});
  return result.status === 0;
}

function readRemoteLogs(options) {
  if (!options.ssh) {
    return [];
  }
  if (!options.remoteLogs.length) {
    throw new Error('--remote-log=PATH is required when --ssh is used.');
  }

  const remoteCommand = [
    'set -eu',
    'for f in ' + options.remoteLogs.map(shellQuote).join(' ') + '; do',
    '  if [ -r "$f" ]; then',
    '    printf "\\n# real-preview-log-source %s\\n" "$f"',
    '    tail -n ' + Number(options.remoteLines) + ' "$f"',
    '  fi',
    'done'
  ].join('\n');
  const sshArgs = [
    '-o', 'StrictHostKeyChecking=accept-new',
    '-o', 'ConnectTimeout=10',
    options.ssh,
    remoteCommand
  ];
  let command = 'ssh';
  let args = sshArgs;
  const env = Object.assign({}, process.env);
  const password = process.env[options.sshPasswordEnv];

  if (password) {
    if (!commandExists('sshpass')) {
      throw new Error('sshpass is required for password SSH. Install it or use SSH keys.');
    }
    command = 'sshpass';
    args = ['-e', 'ssh'].concat(sshArgs);
    env.SSHPASS = password;
  }

  const result = childProcess.spawnSync(command, args, {
    cwd: repoRoot,
    encoding: 'utf8',
    env,
    maxBuffer: 80 * 1024 * 1024
  });

  if (result.status !== 0) {
    throw new Error('Failed to read remote logs:\n' + (result.stderr || result.stdout));
  }

  return splitLines(result.stdout);
}

function readInputLogs(options) {
  const lines = [];

  for (const file of options.logFiles) {
    lines.push('# real-preview-log-source ' + file);
    lines.push(...splitLines(fs.readFileSync(file, 'utf8')));
  }

  lines.push(...readRemoteLogs(options));

  if (!process.stdin.isTTY && !options.logFiles.length && !options.ssh) {
    lines.push(...splitLines(fs.readFileSync(0, 'utf8')));
  }

  return lines;
}

function splitLines(text) {
  return String(text || '').split(/\r?\n/).filter(Boolean);
}

function parseAccessLine(line) {
  const match = /^(\S+) \S+ \S+ \[([^\]]+)] "([^"]*)" (\d{3}) ([0-9-]+)(?: (.*))?$/.exec(line);

  if (!match) {
    return null;
  }

  const request = /^([A-Z]+) ([^ ]+) HTTP\/[^"]+$/.exec(match[3]);
  if (!request) {
    return null;
  }

  const quoted = [];
  const quotedText = match[6] || '';
  const quoteRegex = /"([^"]*)"/g;
  let quotedMatch;
  while ((quotedMatch = quoteRegex.exec(quotedText)) !== null) {
    quoted.push(quotedMatch[1]);
  }

  return {
    remoteAddr: match[1],
    time: match[2],
    method: request[1],
    target: request[2],
    status: Number(match[4]),
    bytes: match[5] === '-' ? 0 : Number(match[5]),
    referer: quoted[0] || '',
    userAgent: quoted[1] || '',
    forwardedFor: quoted[2] || '',
    raw: line
  };
}

function parseRequestTarget(target) {
  try {
    return new URL(target, 'http://real-preview.local');
  } catch (_error) {
    return null;
  }
}

function redactLabel(prefix, index) {
  return prefix + '-' + (index + 1);
}

function sanitizeUrlForReport(value) {
  const url = parseRequestTarget(value);
  const keep = new Set(['_dc', 'type', 'mode', 'fileType', 'isForm', 'previewLite', 'frameEditorId', 'indexPostfix']);
  const params = [];

  if (!url) {
    return '';
  }
  if (!/\/(?:web-apps\/)?apps\//.test(url.pathname)) {
    return '/document-url';
  }

  url.searchParams.forEach((paramValue, name) => {
    if (keep.has(name)) {
      params.push(encodeURIComponent(name) + '=' + encodeURIComponent(paramValue));
    }
  });

  return url.pathname + (params.length ? '?' + params.join('&') : '');
}

function redactFileName(value, ext) {
  if (!value) {
    return '';
  }

  return (ext || extensionFromName(value) || 'file') + '-document';
}

function safeDecode(value) {
  try {
    return decodeURIComponent(String(value || ''));
  } catch (_error) {
    return String(value || '');
  }
}

function normalizeName(value) {
  return safeDecode(value).replace(/\s+/g, '');
}

function extensionFromName(value) {
  const decoded = safeDecode(value).replace(/[?#].*$/, '');
  const match = /\.([a-z0-9]+)$/i.exec(decoded);
  return match ? match[1].toLowerCase() : '';
}

function collectLogFacts(lines, maxCasesPerType) {
  const previewByName = new Map();
  const cases = [];
  const seenCases = new Set();
  const perType = {};
  const editorRoutes = [];
  const counts = {
    accessLines: 0,
    instPreview: 0,
    instFilePreview: 0,
    editorIndex: 0,
    downloadFile: 0,
    cacheFiles: 0
  };
  const sources = [];
  let currentSource = 'stdin';

  for (const line of lines) {
    if (line.startsWith('# real-preview-log-source ')) {
      currentSource = line.substring('# real-preview-log-source '.length);
      sources.push(currentSource);
      continue;
    }

    const parsed = parseAccessLine(line);
    if (!parsed) {
      continue;
    }
    counts.accessLines += 1;
    parsed.source = currentSource;

    const url = parseRequestTarget(parsed.target);
    if (!url) {
      continue;
    }

    if (url.pathname === '/inst/preview') {
      counts.instPreview += 1;
      const fileName = url.searchParams.get('fileName') || url.searchParams.get('downloadUrl') || '';
      const downloadUrl = url.searchParams.get('downloadUrl') || fileName;
      const fileSize = Number(url.searchParams.get('fileSize') || 0);
      const info = {
        fileName: safeDecode(fileName),
        downloadUrl: safeDecode(downloadUrl),
        fileSize: Number.isFinite(fileSize) ? fileSize : 0,
        ext: extensionFromName(fileName || downloadUrl),
        time: parsed.time,
        source: parsed.source
      };
      if (info.fileName) {
        previewByName.set(normalizeName(info.fileName), info);
      }
      if (info.downloadUrl) {
        previewByName.set(normalizeName(info.downloadUrl), info);
      }
      continue;
    }

    if (url.pathname === '/inst/filePreview') {
      counts.instFilePreview += 1;
      const fileName = url.searchParams.get('fileName') || '';
      const normalized = normalizeName(fileName);
      const previewInfo = previewByName.get(normalized) || {};
      const ext = extensionFromName(fileName || previewInfo.downloadUrl || previewInfo.fileName);
      if (!ext) {
        continue;
      }
      perType[ext] = perType[ext] || 0;
      if (perType[ext] >= maxCasesPerType) {
        continue;
      }
      const caseKey = [ext, normalized].join(':');
      if (seenCases.has(caseKey)) {
        continue;
      }
      seenCases.add(caseKey);
      perType[ext] += 1;
      cases.push({
        ext,
        fileName: safeDecode(fileName),
        fileSize: previewInfo.fileSize || 0,
        time: parsed.time,
        requestTarget: parsed.target,
        source: parsed.source,
        referer: parsed.referer
      });
      continue;
    }

    const editorMatch = /\/web-apps\/apps\/([^/]+)\/main\/index\.html$/.exec(url.pathname);
    if (editorMatch) {
      counts.editorIndex += 1;
      editorRoutes.push({
        app: editorMatch[1],
        fileType: (url.searchParams.get('fileType') || '').toLowerCase(),
        mode: url.searchParams.get('mode') || '',
        isForm: url.searchParams.get('isForm') || '',
        previewLite: url.searchParams.get('previewLite') || '',
        target: parsed.target,
        time: parsed.time,
        status: parsed.status,
        referer: parsed.referer
      });
      continue;
    }

    if (/\/downloadfile\//.test(url.pathname)) {
      counts.downloadFile += 1;
    }
    if (/\/cache\/files\//.test(url.pathname)) {
      counts.cacheFiles += 1;
    }
  }

  return {
    sources: Array.from(new Set(sources)),
    counts,
    cases,
    editorRoutes,
    casesByType: cases.reduce((acc, item) => {
      acc[item.ext] = (acc[item.ext] || 0) + 1;
      return acc;
    }, {})
  };
}

function createElement(tagName, harness) {
  const attributes = {};
  const element = {
    tagName: tagName.toUpperCase(),
    style: {},
    attributes,
    parentNode: null,
    children: [],
    width: undefined,
    height: undefined,
    src: '',
    contentWindow: null,
    setAttribute(name, value) {
      attributes[name] = String(value);
    },
    getAttribute(name) {
      return attributes[name];
    },
    appendChild(child) {
      this.children.push(child);
      child.parentNode = this;
      return child;
    },
    getBoundingClientRect() {
      return {left: 0, top: 0};
    }
  };

  if (tagName.toLowerCase() === 'iframe') {
    element.contentWindow = {
      messages: [],
      postMessage(payload, origin, transfer) {
        this.messages.push({payload, origin, transfer});
      }
    };
  }

  harness.created.push(element);
  return element;
}

function makeHarness() {
  const listeners = {};
  const elements = {};
  const storage = {};
  const harness = {
    created: [],
    replacements: [],
    listeners,
    consoleInfo: [],
    previewTraceRequests: [],
    window: null,
    DocsAPI: null,
    getIframe() {
      assert.ok(this.replacements.length > 0, 'expected an iframe replacement');
      return this.replacements[this.replacements.length - 1].newChild;
    }
  };
  const parentNode = {
    replaceChild(newChild, oldChild) {
      newChild.parentNode = parentNode;
      harness.replacements.push({newChild, oldChild});
      return oldChild;
    }
  };
  const target = createElement('div', harness);
  target.parentNode = parentNode;
  target.setAttribute('id', 'placeholder');
  elements.placeholder = target;
  const document = {
    body: {style: {}},
    currentScript: {
      src: 'https://ds.example/web-apps/apps/api/documents/api.js'
    },
    createElement(tagName) {
      return createElement(tagName, harness);
    },
    getElementById(id) {
      return elements[id] || null;
    },
    getElementsByTagName(tagName) {
      if (tagName.toLowerCase() === 'script') {
        return [{src: 'https://ds.example/web-apps/apps/api/documents/api.js'}];
      }
      return [];
    }
  };
  const localStorage = {
    getItem(key) {
      return storage[key] || null;
    },
    setItem(key, value) {
      storage[key] = String(value);
    },
    removeItem(key) {
      delete storage[key];
    }
  };

  function ImageMock() {}
  Object.defineProperty(ImageMock.prototype, 'src', {
    get() {
      return this._src;
    },
    set(value) {
      this._src = value;
      harness.previewTraceRequests.push(value);
    }
  });

  let now = 1000;
  const window = {
    DocsAPI: {},
    JSON,
    Object,
    Image: ImageMock,
    location: {
      origin: 'https://host.example',
      search: ''
    },
    parent: null,
    localStorage,
    performance: {
      now() {
        now += 17;
        return now;
      }
    },
    alert(message) {
      throw new Error(message);
    },
    addEventListener(type, listener) {
      listeners[type] = listeners[type] || [];
      listeners[type].push(listener);
    },
    removeEventListener(type, listener) {
      if (!listeners[type]) {
        return;
      }
      listeners[type] = listeners[type].filter(current => current !== listener);
    },
    console: {
      warn() {},
      info() {
        harness.consoleInfo.push(Array.prototype.slice.call(arguments));
      }
    }
  };
  window.parent = window;

  const context = {
    window,
    document,
    localStorage,
    console: window.console,
    Image: ImageMock,
    performance: window.performance,
    setTimeout,
    clearTimeout
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(apiPath, 'utf8'), context, {filename: apiPath});
  harness.window = window;
  harness.DocsAPI = window.DocsAPI;
  return harness;
}

function mergeConfig(base, overrides) {
  for (const key of Object.keys(overrides || {})) {
    const value = overrides[key];
    if (value && typeof value === 'object' && !Array.isArray(value)) {
      if (!base[key] || typeof base[key] !== 'object' || Array.isArray(base[key])) {
        base[key] = {};
      }
      mergeConfig(base[key], value);
    } else {
      base[key] = value;
    }
  }
  return base;
}

function makeConfig(testCase) {
  const isPdf = testCase.ext === 'pdf';
  const documentType = isPdf ? 'pdf' : officeDocumentTypes[testCase.ext];
  const title = testCase.fileName || 'real-preview.' + testCase.ext;
  const config = {
    type: 'desktop',
    width: '100%',
    height: '100%',
    documentType,
    document: {
      url: 'https://files.example/' + encodeURIComponent(title),
      fileType: testCase.ext,
      key: 'real-preview-' + testCase.ext + '-' + Math.random().toString(16).slice(2),
      title,
      permissions: {
        edit: false
      }
    },
    editorConfig: {
      mode: 'view',
      lang: 'zh-CN',
      customization: {}
    },
    previewTraceNetwork: true,
    events: {}
  };

  if (isPdf) {
    delete config.document.isForm;
  }

  return config;
}

function replayCase(testCase) {
  const harness = makeHarness();
  const config = makeConfig(testCase);
  new harness.DocsAPI.DocEditor('placeholder', config);
  const iframe = harness.getIframe();
  const result = {
    ext: testCase.ext,
    fileName: testCase.fileName,
    fileSize: testCase.fileSize,
    iframeSrc: iframe.src,
    nativePdf: iframe.getAttribute('data-native-pdf-preview') === 'true',
    previewLite: iframe.getAttribute('data-office-preview-lite') === 'true',
    traceEvents: harness.consoleInfo
      .filter(entry => entry[1] && entry[1].event)
      .map(entry => entry[1].event)
  };

  if (testCase.ext === 'pdf') {
    assert.strictEqual(result.nativePdf, true, 'PDF should use native browser preview');
    assert.ok(!/\/(?:common|pdfeditor)\/main\/index\.html/.test(result.iframeSrc), 'PDF must not open full editor path');
    assert.ok(!/[?&]mode=fillforms(?:&|$)/.test(result.iframeSrc), 'PDF must not use fillforms mode');
  } else if (officeDocumentTypes[testCase.ext]) {
    assert.strictEqual(result.previewLite, true, testCase.ext + ' should use previewLite');
    assert.ok(/[?&]previewLite=1(?:&|$)/.test(result.iframeSrc), testCase.ext + ' URL should include previewLite=1');
  } else {
    result.skipped = true;
  }

  return result;
}

function summarizeServerRoutes(editorRoutes) {
  const summary = {
    pdfEditorFillforms: 0,
    pdfNativeFastPathEvidence: 0,
    officePreviewLite: 0,
    officeFullEditor: 0,
    samples: []
  };

  for (const route of editorRoutes) {
    if (route.fileType === 'pdf' && (route.app === 'pdfeditor' || route.app === 'common') && route.mode === 'fillforms') {
      summary.pdfEditorFillforms += 1;
      if (summary.samples.length < 5) {
        summary.samples.push({
          issue: 'pdf-editor-fillforms',
          time: route.time,
          app: route.app,
          mode: route.mode,
          isForm: route.isForm,
          target: route.target
        });
      }
    }
    if (route.fileType === 'pdf' && route.app !== 'pdfeditor' && route.app !== 'common') {
      summary.pdfNativeFastPathEvidence += 1;
    }
    if (officeDocumentTypes[route.fileType]) {
      if (route.previewLite === '1') {
        summary.officePreviewLite += 1;
      } else {
        summary.officeFullEditor += 1;
      }
    }
  }

  return summary;
}

function redactReport(output) {
  return {
    name: output.name,
    logSources: output.logSources.map((_source, index) => redactLabel('source', index)),
    parsed: output.parsed,
    casesByType: output.casesByType,
    replayedCases: output.replayedCases,
    localReplay: output.localReplay.map(item => ({
      ext: item.ext,
      fileName: redactFileName(item.fileName, item.ext),
      fileSize: item.fileSize,
      iframeSrc: sanitizeUrlForReport(item.iframeSrc),
      nativePdf: item.nativePdf,
      previewLite: item.previewLite,
      traceEvents: item.traceEvents,
      skipped: item.skipped
    })),
    serverObserved: {
      pdfEditorFillforms: output.serverObserved.pdfEditorFillforms,
      pdfNativeFastPathEvidence: output.serverObserved.pdfNativeFastPathEvidence,
      officePreviewLite: output.serverObserved.officePreviewLite,
      officeFullEditor: output.serverObserved.officeFullEditor,
      samples: output.serverObserved.samples.map(item => ({
        issue: item.issue,
        time: item.time,
        app: item.app,
        mode: item.mode,
        isForm: item.isForm,
        target: sanitizeUrlForReport(item.target)
      }))
    }
  };
}

function main() {
  const options = parseArgs(process.argv.slice(2));
  const lines = readInputLogs(options);
  const facts = collectLogFacts(lines, options.maxCasesPerType);
  const supportedCases = facts.cases.filter(item => item.ext === 'pdf' || officeDocumentTypes[item.ext]);

  if (!facts.counts.accessLines) {
    throw new Error('No nginx access log lines were parsed.');
  }
  if (!supportedCases.length) {
    throw new Error('No /inst/filePreview office/pdf cases were found in the logs.');
  }

  const localReplay = options.analyzeOnly ? [] : supportedCases.map(replayCase);
  const output = {
    name: TEST_NAME,
    logSources: facts.sources,
    parsed: facts.counts,
    casesByType: facts.casesByType,
    replayedCases: localReplay.length,
    localReplay,
    serverObserved: summarizeServerRoutes(facts.editorRoutes)
  };
  const report = options.showSensitive ? output : redactReport(output);

  if (options.json) {
    console.log(JSON.stringify(report, null, 2));
    return;
  }

  console.log(TEST_NAME + ': OK');
  console.log('logSources: ' + (report.logSources.length ? report.logSources.join(', ') : 'stdin'));
  console.log('parsed: ' + JSON.stringify(report.parsed));
  console.log('casesByType: ' + JSON.stringify(report.casesByType));
  console.log('replayedCases: ' + report.replayedCases);
  console.log('serverObserved: ' + JSON.stringify(report.serverObserved));
}

try {
  main();
} catch (error) {
  console.error(TEST_NAME + ': FAILED');
  console.error(error && error.stack ? error.stack : error);
  process.exit(1);
}
