// Execute the real frontend against a small DOM adapter and controlled network responses.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

class Element {
    constructor() {
        this.listeners = new Map();
        this.classes = new Set();
        this.children = [];
        this._innerHTML = 'Convert';
        this._textContent = '';
        this.value = '';
        this.disabled = false;
        this.checked = false;
        this.style = {};
        this.classList = {
            add: value => this.classes.add(value),
            remove: value => this.classes.delete(value),
            toggle: (value, enabled) => enabled ? this.classes.add(value) : this.classes.delete(value),
            contains: value => this.classes.has(value),
        };
    }
    get textContent() {
        return this._textContent || this._innerHTML.replace(/<[^>]+>/g, ' ');
    }
    set textContent(value) {
        this._textContent = value;
    }
    get innerHTML() { return this._innerHTML; }
    set innerHTML(value) {
        this._innerHTML = value;
        if (value === '') {
            this.children = [];
        }
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
    const copied = [];
    const timers = [];
    const context = {
        AbortController, Blob,
        FormData: class { append() {} },
        navigator: {
            clipboard: {
                writeText: async text => { copied.push(text); }
            }
        },
        document: {
            getElementById: element,
            addEventListener: (_, callback) => callback(),
            createElement: () => new Element(),
            body: new Element(),
        },
        URL: { createObjectURL: blob => { blobs.push(blob); return 'blob:test'; }, revokeObjectURL() {} },
        alert: value => alerts.push(value),
        setTimeout: (callback, delay) => {
            if (!delay) callback();
            else timers.push(callback);
        },
        fetch: async () => { throw new Error('No response configured'); },
    };
    vm.runInNewContext(fs.readFileSync(process.argv[2], 'utf8'), context);
    return { element, context, blobs, alerts, copied, timers };
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

    // 1. Convert tab verification
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

    // Clipboard and filter checks on Convert tab
    await app.element('copyBtn').fire('click');
    await new Promise(r => setImmediate(r));
    assert.equal(app.copied.length >= 1, true);
    assert.match(app.copied[0], /modbusRTU/);
    assert.equal(app.element('toastContainer').children.length >= 1, true);

    app.element('previewFilterInput').value = 'script';
    app.element('previewFilterInput').fire('input');
    assert.equal(app.element('previewBody').children[0].classList.contains('hidden'), false);

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

    // 2. Tab Navigation Contracts
    app.element('tabBtnTag').fire('click');
    assert.equal(app.element('tabBtnTag').classList.contains('active'), true);
    assert.equal(app.element('tabBtnConvert').classList.contains('active'), false);
    assert.equal(app.element('tab-tag').classList.contains('active'), true);
    assert.equal(app.element('tab-convert').classList.contains('hidden'), true);

    // Keyboard navigation: ArrowRight moves from tab-tag to tab-audit
    app.element('navTabs').fire('keydown', { key: 'ArrowRight', preventDefault: () => {} });
    assert.equal(app.element('tabBtnAudit').classList.contains('active'), true);
    assert.equal(app.element('tab-audit').classList.contains('active'), true);

    // Switch back to tag tab for tagger tests
    app.element('tabBtnTag').fire('click');

    // 3. SunPM Tagger Tab Flow
    app.element('tagFileInput').fire('change', { target: { files: [{ name: 'inverter.csv', size: 1024 }] } });
    assert.equal(app.element('tagProcessBtn').disabled, false);
    app.context.fetch = async () => response({
        success: true,
        filename: 'tagged_inverter.csv',
        tags_updated: 15,
        register_count: 50,
        preview: [{ Name: 'Active Power', Tag: 'P_AC', Address: 30001, Type: 'U16', Unit: 'W', Action: '4' }],
        csv_content: 'sample_csv'
    });
    await app.element('tagProcessBtn').fire('click');
    assert.equal(app.element('tagCountBadge').textContent, '15 tags updated');
    assert.equal(app.element('tagRegisterBadge').textContent, '50 registers');
    assert.equal(app.element('tagResultsSection').classList.contains('hidden'), false);
    await app.element('tagCopyBtn').fire('click');
    await new Promise(r => setImmediate(r));
    assert.equal(app.copied.includes('sample_csv'), true);

    // 4. Auditor Tab Flow (Errors & Fix applied)
    app.element('tabBtnAudit').fire('click');
    assert.equal(app.element('tab-audit').classList.contains('active'), true);
    app.element('auditFileInput').fire('change', { target: { files: [{ name: 'defective.csv', size: 512 }] } });
    assert.equal(app.element('auditProcessBtn').disabled, false);
    app.context.fetch = async () => response({
        success: true,
        filename: 'defective.csv',
        valid: false,
        register_count: 30,
        error_count: 2,
        warning_count: 1,
        issues: [
            { level: 'error', field: 'Info2', message: 'Address overlap at 40001' },
            { level: 'error', field: 'Tag', message: 'Duplicate tag P_AC' },
            { level: 'warning', field: 'Unit', message: 'Non-canonical unit kwh' }
        ],
        fixed_csv_content: 'repaired_csv_content'
    });
    await app.element('auditProcessBtn').fire('click');
    assert.equal(app.element('auditStatusBadge').textContent, 'Issues Found');
    assert.equal(app.element('auditRegistersBadge').textContent, '30 registers');
    assert.equal(app.element('auditErrorsBadge').textContent, '2 errors');
    assert.equal(app.element('auditWarningsBadge').textContent, '1 warnings');
    assert.equal(app.element('auditDownloadFixedBtn').classList.contains('hidden'), false);
    assert.equal(app.element('auditCopyFixedBtn').classList.contains('hidden'), false);
    assert.equal(app.element('auditIssuesList').children.length, 3);
    await app.element('auditCopyFixedBtn').fire('click');
    await new Promise(r => setImmediate(r));
    assert.equal(app.copied.includes('repaired_csv_content'), true);

    // 5. Equipment Library Flow (Tab 4)
    const sampleLibrary = [
        { filename: 'INV_Huawei_SUN2000.csv', category: 'Inverter', manufacturer: 'Huawei', model: 'SUN2000', size_kb: 12.5 },
        { filename: 'MET_ABB_B23.csv', category: 'Meter', manufacturer: 'ABB', model: 'B23', size_kb: 8.2 }
    ];
    app.context.fetch = async () => response(sampleLibrary);
    app.element('tabBtnLibrary').fire('click');
    // Allow promise tick
    await new Promise(r => setImmediate(r));
    assert.equal(app.element('libraryCountBadge').textContent, '2 models');
    assert.equal(app.element('libraryTableBody').children.length, 2);

    // Search filter
    app.element('librarySearchInput').value = 'ABB';
    app.element('librarySearchInput').fire('input');
    assert.equal(app.element('libraryTableBody').children.length, 1);
}

main().catch(error => { console.error(error); process.exitCode = 1; });
