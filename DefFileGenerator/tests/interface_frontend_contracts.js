// Execute the real frontend against a small DOM adapter and controlled network responses.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

class Element {
    constructor() {
        this.listeners = new Map();
        this.classes = new Set();
        this.children = [];
        this.innerHTML = 'Convert';
        this.textContent = '';
        this.value = '';
        this.disabled = false;
        this.classList = {
            add: value => this.classes.add(value),
            remove: value => this.classes.delete(value),
            toggle: (value, enabled) => enabled ? this.classes.add(value) : this.classes.delete(value),
            contains: value => this.classes.has(value),
        };
    }
    addEventListener(name, callback) { this.listeners.set(name, callback); }
    fire(name, event = {}) { return this.listeners.get(name)?.(event); }
    appendChild(child) { this.children.push(child); }
    removeChild(child) { this.children = this.children.filter(item => item !== child); }
    setAttribute() {}
    click() {}
}

function application() {
    const elements = new Map();
    const element = id => {
        if (!elements.has(id)) elements.set(id, new Element());
        return elements.get(id);
    };
    const blobs = [];
    const alerts = [];
    const context = {
        AbortController, Blob,
        FormData: class { append() {} },
        document: {
            getElementById: element,
            addEventListener: (_, callback) => callback(),
            createElement: () => new Element(),
            body: new Element(),
        },
        URL: { createObjectURL: blob => { blobs.push(blob); return 'blob:test'; }, revokeObjectURL() {} },
        alert: value => alerts.push(value),
        setTimeout: callback => callback(),
        fetch: async () => { throw new Error('No response configured'); },
    };
    vm.runInNewContext(fs.readFileSync(process.argv[2], 'utf8'), context);
    return { element, context, blobs, alerts };
}

function selectFile(app) {
    app.element('fileInput').fire('change', { target: { files: [{ name: 'map.csv', size: 10 }] } });
}

function response(data, ok = true) { return { ok, json: async () => data }; }
const validResult = {
    csv_content: 'modbusRTU;Inverter;M;X;;;;;;;\n1;3;0;U16;;N;n;0;0;;4\n',
    filename: 'output.csv', register_count: 2, extracted_count: 3, is_valid: true,
    preview: [{ Name: '<script>', Tag: 'n', Address: 0, Factor: 0 }],
};

async function main() {
    const app = application();
    selectFile(app);
    app.context.fetch = async () => response(validResult);
    await app.element('convertBtn').fire('click');
    assert.equal(app.element('registerBadge').textContent, '2 registers generated (1 skipped)');
    assert.equal(app.element('validBadge').textContent, 'Valid');
    assert.equal(app.element('downloadBtn').disabled, false);
    assert.match(app.element('previewBody').children[0].innerHTML, /&lt;script&gt;/);
    assert.match(app.element('previewBody').children[0].innerHTML, /<code>0<\/code>/);
    app.element('downloadBtn').fire('click');
    assert.deepEqual(Array.from(new Uint8Array(await app.blobs[0].arrayBuffer())).slice(0, 3), [239, 187, 191]);

    app.context.fetch = async () => response({ ...validResult, is_valid: false });
    await app.element('convertBtn').fire('click');
    assert.equal(app.element('validBadge').textContent, 'Invalid');
    assert.equal(app.element('validBadge').classList.contains('badge-success'), false);
    assert.equal(app.element('downloadBtn').disabled, true);

    app.context.fetch = async () => response({ detail: 'Bad document' }, false);
    await app.element('convertBtn').fire('click');
    assert.equal(app.element('resultsSection').classList.contains('hidden'), true);
    assert.equal(app.element('downloadBtn').disabled, true);
    assert.equal(app.alerts.length, 1);

    let finishRequest;
    app.context.fetch = () => new Promise(resolve => { finishRequest = resolve; });
    const pending = app.element('convertBtn').fire('click');
    app.element('removeFileBtn').fire('click');
    finishRequest(response(validResult));
    await pending;
    assert.equal(app.element('resultsSection').classList.contains('hidden'), true);
    assert.equal(app.element('convertBtn').disabled, true);
    assert.equal(app.element('downloadBtn').disabled, true);

    selectFile(app);
    app.context.fetch = async () => response(validResult);
    await app.element('convertBtn').fire('click');
    app.element('convertForm').fire('input');
    assert.equal(app.element('resultsSection').classList.contains('hidden'), true);
    assert.equal(app.element('downloadBtn').disabled, true);
}

main().catch(error => { console.error(error); process.exitCode = 1; });
