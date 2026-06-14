const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const apiPath = path.resolve(__dirname, '../../apps/api/documents/api.js');
const apiSource = fs.readFileSync(apiPath, 'utf8');

const tests = [];

function test(name, fn) {
    tests.push({name, fn});
}

function mergeConfig(base, overrides) {
    Object.keys(overrides || {}).forEach(function(key) {
        var value = overrides[key];

        if (value && typeof value === 'object' && !Array.isArray(value)) {
            if (!base[key] || typeof base[key] !== 'object' || Array.isArray(base[key]))
                base[key] = {};
            mergeConfig(base[key], value);
        } else {
            base[key] = value;
        }
    });
    return base;
}

function baseConfig(overrides) {
    return mergeConfig({
        type: 'desktop',
        width: '100%',
        height: '100%',
        documentType: 'pdf',
        document: {
            url: 'https://files.example/doc.pdf',
            fileType: 'pdf',
            key: 'doc-key',
            title: 'doc.pdf',
            isForm: false,
            permissions: {
                edit: false
            }
        },
        editorConfig: {
            mode: 'view',
            lang: 'en',
            customization: {}
        },
        events: {}
    }, overrides || {});
}

function officeConfig(fileType, documentType, overrides) {
    return baseConfig(mergeConfig({
        documentType: documentType,
        document: {
            url: 'https://files.example/doc.' + fileType,
            fileType: fileType,
            key: fileType + '-key',
            title: 'doc.' + fileType,
            permissions: {
                edit: false
            }
        }
    }, overrides || {}));
}

function createElement(tagName, harness) {
    var attributes = {};
    var element = {
        tagName: tagName.toUpperCase(),
        style: {},
        attributes: attributes,
        parentNode: null,
        children: [],
        width: undefined,
        height: undefined,
        src: '',
        contentWindow: null,
        setAttribute: function(name, value) {
            attributes[name] = String(value);
        },
        getAttribute: function(name) {
            return attributes[name];
        },
        appendChild: function(child) {
            this.children.push(child);
            child.parentNode = this;
            return child;
        },
        getBoundingClientRect: function() {
            return {left: 0, top: 0};
        }
    };

    if (tagName.toLowerCase() === 'iframe') {
        element.contentWindow = {
            messages: [],
            postMessage: function(payload, origin, transfer) {
                this.messages.push({
                    payload: payload,
                    origin: origin,
                    transfer: transfer
                });
            }
        };
    }

    harness.created.push(element);
    return element;
}

function makeHarness() {
    var listeners = {};
    var elements = {};
    var storage = {};
    var harness = {
        created: [],
        replacements: [],
        listeners: listeners,
        window: null,
        document: null,
        DocsAPI: null,
        dispatchMessage: function(message) {
            (listeners.message || []).slice().forEach(function(listener) {
                listener(message);
            });
        },
        getIframe: function() {
            assert.ok(this.replacements.length > 0, 'expected an iframe replacement');
            return this.replacements[this.replacements.length - 1].newChild;
        }
    };

    var parentNode = {
        replaceChild: function(newChild, oldChild) {
            newChild.parentNode = parentNode;
            harness.replacements.push({
                newChild: newChild,
                oldChild: oldChild
            });
            return oldChild;
        }
    };
    var target = createElement('div', harness);
    target.parentNode = parentNode;
    target.setAttribute('id', 'placeholder');
    elements.placeholder = target;

    var document = {
        body: {
            style: {}
        },
        currentScript: {
            src: 'https://ds.example/web-apps/apps/api/documents/api.js'
        },
        createElement: function(tagName) {
            return createElement(tagName, harness);
        },
        getElementById: function(id) {
            return elements[id] || null;
        },
        getElementsByTagName: function(tagName) {
            if (tagName.toLowerCase() === 'script')
                return [{src: 'https://ds.example/web-apps/apps/api/documents/api.js'}];
            return [];
        }
    };

    var localStorage = {
        getItem: function(key) {
            return storage[key] || null;
        },
        setItem: function(key, value) {
            storage[key] = String(value);
        },
        removeItem: function(key) {
            delete storage[key];
        }
    };

    var window = {
        DocsAPI: {},
        JSON: JSON,
        Object: Object,
        location: {
            origin: 'https://host.example',
            search: ''
        },
        parent: null,
        localStorage: localStorage,
        alert: function(message) {
            throw new Error(message);
        },
        addEventListener: function(type, listener) {
            if (!listeners[type])
                listeners[type] = [];
            listeners[type].push(listener);
        },
        removeEventListener: function(type, listener) {
            if (!listeners[type])
                return;
            listeners[type] = listeners[type].filter(function(current) {
                return current !== listener;
            });
        }
    };
    window.parent = window;

    var context = {
        window: window,
        document: document,
        localStorage: localStorage,
        console: console,
        JSON: JSON,
        Object: Object,
        Error: Error,
        Promise: Promise,
        setTimeout: setTimeout,
        clearTimeout: clearTimeout,
        decodeURIComponent: decodeURIComponent
    };

    harness.window = window;
    harness.document = document;
    vm.createContext(context);
    vm.runInContext(apiSource, context, {filename: apiPath});
    harness.DocsAPI = window.DocsAPI;
    return harness;
}

function createEditor(config) {
    var harness = makeHarness();
    new harness.DocsAPI.DocEditor('placeholder', config);
    return {
        harness: harness,
        iframe: harness.getIframe()
    };
}

function parseMessage(entry) {
    return typeof entry.payload === 'string' ? JSON.parse(entry.payload) : entry.payload;
}

function assertNoNativePdf(iframe, sourceUrl) {
    assert.notStrictEqual(iframe.getAttribute('data-native-pdf-preview'), 'true');
    assert.notStrictEqual(iframe.getAttribute('data-onlyoffice-native-pdf-preview'), 'true');
    assert.notStrictEqual(iframe.src, sourceUrl);
}

function assertNoPreviewLite(iframe) {
    assert.strictEqual(iframe.getAttribute('data-office-preview-lite'), undefined);
    assert.strictEqual(iframe.getAttribute('data-onlyoffice-preview-lite'), undefined);
    assert.ok(iframe.src.indexOf('previewLite=1') < 0, 'previewLite URL marker should be absent');
}

test('PDF read-only non-form files use a native iframe and fire ready events on load', function() {
    var fired = [];
    var config = baseConfig({
        events: {
            onAppReady: function(event) {
                assert.ok(event.target);
                fired.push('onAppReady');
            },
            onDocumentReady: function(event) {
                assert.ok(event.target);
                fired.push('onDocumentReady');
            }
        }
    });
    var result = createEditor(config);

    assert.strictEqual(result.iframe.src, config.document.url);
    assert.strictEqual(result.iframe.getAttribute('data-native-pdf-preview'), 'true');
    assert.strictEqual(result.iframe.getAttribute('data-onlyoffice-native-pdf-preview'), 'true');
    assert.strictEqual((result.harness.listeners.message || []).length, 0);
    assert.deepStrictEqual(result.iframe.contentWindow.messages, []);

    result.iframe.onload();
    assert.deepStrictEqual(fired, ['onAppReady', 'onDocumentReady']);
});

test('PDF native iframe can be disabled with openPdfInBrowser=false', function() {
    var config = baseConfig({openPdfInBrowser: false});
    var result = createEditor(config);

    assertNoNativePdf(result.iframe, config.document.url);
    assert.match(result.iframe.src, /\/pdfeditor\/main\/index\.html/);
    assert.match(result.iframe.src, /[?&]mode=view(?:&|$)/);
    assert.doesNotMatch(result.iframe.src, /[?&]mode=fillforms(?:&|$)/);
    assert.strictEqual((result.harness.listeners.message || []).length, 1);
});

test('PDF native iframe can be disabled through document.openPdfInBrowser=false', function() {
    var config = baseConfig({
        document: {
            openPdfInBrowser: false
        }
    });
    var result = createEditor(config);

    assertNoNativePdf(result.iframe, config.document.url);
    assert.match(result.iframe.src, /\/pdfeditor\/main\/index\.html/);
    assert.match(result.iframe.src, /[?&]mode=view(?:&|$)/);
    assert.doesNotMatch(result.iframe.src, /[?&]mode=fillforms(?:&|$)/);
});

test('PDF native iframe can be disabled through customization.openPdfInBrowser=false', function() {
    var config = baseConfig({
        editorConfig: {
            customization: {
                openPdfInBrowser: false
            }
        }
    });
    var result = createEditor(config);

    assertNoNativePdf(result.iframe, config.document.url);
    assert.match(result.iframe.src, /\/pdfeditor\/main\/index\.html/);
});

test('PDF openPdfAsBinary fetches the source and transfers it to the PDF editor', async function() {
    var buffer = new ArrayBuffer(8);
    var fetchCalls = [];
    var config = baseConfig({openPdfAsBinary: true});
    var harness = makeHarness();

    harness.window.fetch = function(url, options) {
        fetchCalls.push({url: url, options: options});
        return Promise.resolve({
            ok: true,
            arrayBuffer: function() {
                return Promise.resolve(buffer);
            }
        });
    };

    new harness.DocsAPI.DocEditor('placeholder', config);
    var iframe = harness.getIframe();

    assertNoNativePdf(iframe, config.document.url);
    assert.match(iframe.src, /\/pdfeditor\/main\/index\.html/);

    harness.dispatchMessage({
        origin: 'https://ds.example',
        data: JSON.stringify({
            frameEditorId: 'placeholder',
            event: 'onAppReady'
        })
    });
    await new Promise(function(resolve) {
        setTimeout(resolve, 0);
    });

    assert.strictEqual(fetchCalls.length, 1);
    assert.strictEqual(fetchCalls[0].url, config.document.url);
    assert.strictEqual(fetchCalls[0].options.credentials, 'include');

    var messages = iframe.contentWindow.messages.map(parseMessage);
    assert.strictEqual(messages[0].command, 'init');
    assert.strictEqual(messages[1].command, 'openDocumentFromBinary');
    assert.strictEqual(messages[1].data.doc.url, config.document.url);
    assert.strictEqual(messages[1].data.buffer, buffer);
    assert.strictEqual(iframe.contentWindow.messages[1].transfer.length, 1);
    assert.strictEqual(iframe.contentWindow.messages[1].transfer[0], buffer);
});

test('PDF openPdfAsBinary can be enabled through document.openPdfAsBinary=true', async function() {
    var buffer = new ArrayBuffer(4);
    var config = baseConfig({
        document: {
            openPdfAsBinary: true
        }
    });
    var harness = makeHarness();

    harness.window.fetch = function() {
        return Promise.resolve({
            ok: true,
            arrayBuffer: function() {
                return Promise.resolve(buffer);
            }
        });
    };

    new harness.DocsAPI.DocEditor('placeholder', config);
    var iframe = harness.getIframe();

    assertNoNativePdf(iframe, config.document.url);
    harness.dispatchMessage({
        origin: 'https://ds.example',
        data: JSON.stringify({
            frameEditorId: 'placeholder',
            event: 'onAppReady'
        })
    });
    await new Promise(function(resolve) {
        setTimeout(resolve, 0);
    });

    var messages = iframe.contentWindow.messages.map(parseMessage);
    assert.strictEqual(messages[1].command, 'openDocumentFromBinary');
    assert.strictEqual(messages[1].data.doc.url, config.document.url);
    assert.strictEqual(messages[1].data.buffer, buffer);
});

test('PDF forms, unknown form status, fill form, review, and edit modes keep the full editor path', function() {
    var scenarios = [
        baseConfig({
            document: {
                isForm: true
            }
        }),
        baseConfig({
            document: {
                isForm: undefined
            }
        }),
        baseConfig({
            document: {
                permissions: {
                    edit: false,
                    fillForms: true
                }
            }
        }),
        baseConfig({
            document: {
                permissions: {
                    edit: false,
                    review: true
                }
            }
        }),
        baseConfig({
            editorConfig: {
                mode: 'edit'
            },
            document: {
                permissions: {
                    edit: true
                }
            }
        })
    ];

    scenarios.forEach(function(config) {
        var result = createEditor(config);
        assertNoNativePdf(result.iframe, config.document.url);
    });
});

test('PDF with unknown form status still sends checkParams from the full chain', function() {
    var config = baseConfig({
        document: {
            isForm: undefined
        }
    });
    var result = createEditor(config);

    assertNoNativePdf(result.iframe, config.document.url);
    assert.match(result.iframe.src, /\/common\/index\.html/);
    assert.strictEqual(typeof result.iframe.onload, 'function');

    result.iframe.onload();
    var message = parseMessage(result.iframe.contentWindow.messages[0]);
    assert.strictEqual(message.command, 'checkParams');
    assert.deepStrictEqual(message.data, {
        url: config.document.url,
        key: config.document.key
    });
});

test('Office DOCX, XLSX, and PPTX read-only previews use previewLite', function() {
    [
        ['docx', 'word', /\/documenteditor\/main\/index\.html/],
        ['xlsx', 'cell', /\/spreadsheeteditor\/main\/index\.html/],
        ['pptx', 'slide', /\/presentationeditor\/main\/index\.html/]
    ].forEach(function(entry) {
        var result = createEditor(officeConfig(entry[0], entry[1]));

        assert.match(result.iframe.src, entry[2]);
        assert.match(result.iframe.src, /[?&]previewLite=1(?:&|$)/);
        assert.strictEqual(result.iframe.getAttribute('data-office-preview-lite'), 'true');
        assert.strictEqual(result.iframe.getAttribute('data-onlyoffice-preview-lite'), 'true');
    });
});

test('Office previewLite can be disabled with openOfficePreviewLite=false', function() {
    assertNoPreviewLite(createEditor(officeConfig('docx', 'word', {
        openOfficePreviewLite: false
    })).iframe);

    assertNoPreviewLite(createEditor(officeConfig('docx', 'word', {
        editorConfig: {
            customization: {
                openOfficePreviewLite: false
            }
        }
    })).iframe);

    assertNoPreviewLite(createEditor(officeConfig('docx', 'word', {
        document: {
            openOfficePreviewLite: false
        }
    })).iframe);
});

test('Office edit, review, and request-edit-rights modes keep the full editor path', function() {
    assertNoPreviewLite(createEditor(officeConfig('docx', 'word', {
        editorConfig: {
            mode: 'edit'
        },
        document: {
            permissions: {
                edit: true
            }
        }
    })).iframe);

    assertNoPreviewLite(createEditor(officeConfig('docx', 'word', {
        document: {
            permissions: {
                edit: false,
                review: true
            }
        }
    })).iframe);

    assertNoPreviewLite(createEditor(officeConfig('docx', 'word', {
        events: {
            onRequestEditRights: function() {}
        }
    })).iframe);
});

test('Office embedded and mobile editors do not use previewLite', function() {
    assertNoPreviewLite(createEditor(officeConfig('docx', 'word', {
        type: 'embedded'
    })).iframe);

    assertNoPreviewLite(createEditor(officeConfig('docx', 'word', {
        type: 'mobile'
    })).iframe);
});

test('warmUp preserves the default preload and supports previewLite preload explicitly', function() {
    var defaultHarness = makeHarness();
    defaultHarness.DocsAPI.DocEditor.warmUp('placeholder');
    var defaultIframe = defaultHarness.getIframe();

    assert.ok(defaultIframe.src.endsWith('/api/documents/preload.html'));
    assert.strictEqual(defaultIframe.getAttribute('data-office-preview-lite'), undefined);

    var liteHarness = makeHarness();
    liteHarness.DocsAPI.DocEditor.warmUp('placeholder', {previewLite: true});
    var liteIframe = liteHarness.getIframe();

    assert.ok(liteIframe.src.endsWith('/api/documents/preload-lite.html'));
    assert.strictEqual(liteIframe.getAttribute('data-office-preview-lite'), 'true');
    assert.strictEqual(liteIframe.getAttribute('data-onlyoffice-preview-lite'), 'true');
});

(async function run() {
    for (const current of tests) {
        await current.fn();
    }

    console.log('api preview optimization tests passed');
})().catch(function(error) {
    console.error(error && error.stack ? error.stack : error);
    process.exit(1);
});
