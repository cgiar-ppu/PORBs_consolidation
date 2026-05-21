(async function() {
    // --- Configuration ---
    const SKIP = ['Summary', 'AOW00', 'AOW01', 'AOW02', 'AOW03', 'AOW04', 'AOW05', 'W3/Bilateral', 'MELIA Study', 'Anaplan'];
    const sleep = ms => new Promise(r => setTimeout(r, ms));
    const clean = t => t.replace(/check_circle_outline/g, '').trim();

    // Sanitize for filesystem-safe filenames
    const safe = s => s.replace(/[\\/:*?"<>|]/g, '').replace(/\s*\(.*?\)\s*/g, '').trim().replace(/\s+/g, '_');

    // =====================================================================
    //  1. LOAD JSZip
    // =====================================================================
    console.log('📦 Loading JSZip…');
    try {
        await new Promise((resolve, reject) => {
            if (window.JSZip) { resolve(); return; }
            const s = document.createElement('script');
            s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
            s.onload  = resolve;
            s.onerror = () => reject(new Error('Failed to load JSZip from CDN'));
            document.head.appendChild(s);
        });
        console.log('✅ JSZip ready');
    } catch (e) {
        console.error('❌ ' + e.message);
        console.error('   Cannot proceed without JSZip. Aborting.');
        return;
    }

    const zip = new JSZip();
    const capturedFiles = [];           // { filename, blob }
    let currentCenterName = '';
    let interceptActive = true;         // master switch — disables all hooks

    // =====================================================================
    //  HELPERS
    // =====================================================================
    function dataURLToBlob(dataURL) {
        const [header, b64] = dataURL.split(',');
        const mime = header.match(/:(.*?);/)[1];
        const bin  = atob(b64);
        const arr  = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
        return new Blob([arr], { type: mime });
    }

    /** Store one file in the ZIP. Applies center-name rename if needed. */
    function captureFile(filename, blob) {
        if (!interceptActive) return false;
        let name = filename || 'download';
        // Rename center<ID> → display name (safe if setter already did it)
        if (currentCenterName && /center\d+/i.test(name)) {
            name = name.replace(/center\d+/i, currentCenterName);
        }
        capturedFiles.push({ filename: name, blob });
        zip.file(name, blob);
        console.log(`   📥 Captured: ${name} (${(blob.size / 1024).toFixed(1)} KB)`);
        return true;
    }

    /** Try to capture a download from an anchor element. */
    function tryCapture(anchor) {
        const href     = anchor.getAttribute('href');
        const filename = anchor.getAttribute('download');
        if (!filename) return false;

        let blob = null;
        if (href && href.startsWith('blob:')) {
            blob = blobMap.get(href);
            if (blob) URL.revokeObjectURL(href);
        } else if (href && href.startsWith('data:')) {
            try { blob = dataURLToBlob(href); } catch (e) { /* ignore */ }
        }
        return blob ? captureFile(filename, blob) : false;
    }

    // =====================================================================
    //  INTERCEPTION LAYER 1 — URL.createObjectURL  (blob → URL mapping)
    // =====================================================================
    const blobMap = new Map();
    const _createObjectURL = URL.createObjectURL.bind(URL);
    URL.createObjectURL = function(obj) {
        const url = _createObjectURL(obj);
        if (obj instanceof Blob) blobMap.set(url, obj);
        return url;
    };

    // =====================================================================
    //  INTERCEPTION LAYER 2 — HTMLAnchorElement.prototype.click
    // =====================================================================
    const _anchorClick = HTMLAnchorElement.prototype.click;
    HTMLAnchorElement.prototype.click = function() {
        if (interceptActive && tryCapture(this)) return;
        return _anchorClick.call(this);
    };

    // =====================================================================
    //  INTERCEPTION LAYER 3 — HTMLElement.prototype.click  (higher chain)
    //  Some libraries call click() via the HTMLElement prototype directly.
    // =====================================================================
    const _htmlClick = HTMLElement.prototype.click;
    HTMLElement.prototype.click = function() {
        if (interceptActive && this instanceof HTMLAnchorElement && tryCapture(this)) return;
        return _htmlClick.call(this);
    };

    // =====================================================================
    //  INTERCEPTION LAYER 4 — EventTarget.prototype.dispatchEvent
    //  Catches  a.dispatchEvent(new MouseEvent('click'))  pattern.
    // =====================================================================
    const _dispatch = EventTarget.prototype.dispatchEvent;
    EventTarget.prototype.dispatchEvent = function(event) {
        if (interceptActive && event.type === 'click'
            && this instanceof HTMLAnchorElement && tryCapture(this)) {
            return true;
        }
        return _dispatch.call(this, event);
    };

    // =====================================================================
    //  INTERCEPTION LAYER 5 — Document capturing click listener
    //  Catches clicks that propagate through the DOM on anchor elements.
    // =====================================================================
    function onCapturingClick(e) {
        if (!interceptActive) return;
        const anchor = e.target.closest ? e.target.closest('a[download]') : null;
        if (anchor && tryCapture(anchor)) {
            e.preventDefault();
            e.stopImmediatePropagation();
        }
    }
    document.addEventListener('click', onCapturingClick, true);

    // =====================================================================
    //  INTERCEPTION LAYER 6 — navigator.msSaveBlob / msSaveOrOpenBlob
    //  Legacy IE/Edge download method, still used by some libraries.
    // =====================================================================
    const _msSave       = navigator.msSaveBlob       ? navigator.msSaveBlob.bind(navigator)       : null;
    const _msSaveOrOpen = navigator.msSaveOrOpenBlob ? navigator.msSaveOrOpenBlob.bind(navigator) : null;
    if (_msSave) {
        navigator.msSaveBlob = function(blob, name) {
            if (interceptActive && captureFile(name, blob)) return true;
            return _msSave(blob, name);
        };
    }
    if (_msSaveOrOpen) {
        navigator.msSaveOrOpenBlob = function(blob, name) {
            if (interceptActive && captureFile(name, blob)) return true;
            return _msSaveOrOpen(blob, name);
        };
    }

    // =====================================================================
    //  INTERCEPTION LAYER 7 — window.showSaveFilePicker  (File System API)
    //  This API ALWAYS shows a save dialog by design.  We return a mock
    //  FileSystemFileHandle that captures whatever the app writes to it.
    // =====================================================================
    const _showSavePicker = window.showSaveFilePicker
        ? window.showSaveFilePicker.bind(window) : null;
    if (_showSavePicker) {
        window.showSaveFilePicker = async function(opts = {}) {
            if (!interceptActive) return _showSavePicker(opts);
            let filename = opts.suggestedName || 'download';
            if (currentCenterName && /center\d+/i.test(filename)) {
                filename = filename.replace(/center\d+/i, currentCenterName);
            }
            console.log(`   🔀 showSaveFilePicker intercepted → ${filename}`);
            const chunks = [];
            return {
                kind: 'file',
                name: filename,
                createWritable: async () => ({
                    write: async (data) => {
                        if (data instanceof Blob) chunks.push(data);
                        else if (data instanceof ArrayBuffer) chunks.push(new Blob([data]));
                        else if (typeof data === 'object' && data !== null && data.type === 'write') {
                            const d = data.data;
                            chunks.push(d instanceof Blob ? d : new Blob([d]));
                        } else {
                            chunks.push(new Blob([data]));
                        }
                    },
                    close: async () => {
                        const blob = new Blob(chunks);
                        captureFile(filename, blob);
                    },
                    seek:     async () => {},
                    truncate: async () => {},
                }),
            };
        };
    }

    // =====================================================================
    //  INTERCEPTION LAYER 8 — Patch download attribute  (center-name rename)
    //  Preserved from the original script.
    // =====================================================================
    const _downloadDesc = Object.getOwnPropertyDescriptor(
        HTMLAnchorElement.prototype, 'download');
    Object.defineProperty(HTMLAnchorElement.prototype, 'download', {
        configurable: true,
        set(v) {
            let newVal = v;
            if (currentCenterName && typeof v === 'string') {
                newVal = v.replace(/center\d+/i, currentCenterName);
            }
            this.setAttribute('download', newVal);
        },
        get() { return this.getAttribute('download') || ''; }
    });

    // --- Report installed hooks ---
    const hooks = [
        'URL.createObjectURL', 'anchor.click()', 'HTMLElement.click()',
        'dispatchEvent(click)', 'document click listener (capture phase)',
    ];
    if (_msSave)       hooks.push('navigator.msSaveBlob');
    if (_msSaveOrOpen) hooks.push('navigator.msSaveOrOpenBlob');
    if (_showSavePicker) hooks.push('window.showSaveFilePicker');
    console.log(`🔧 ${hooks.length} interception hooks active:\n   ${hooks.join('\n   ')}`);

    // =====================================================================
    //  NAVIGATE CENTERS & TRIGGER EXPORTS
    // =====================================================================
    const nav = document.querySelector('nav.porb-nav--centers');
    if (!nav) { console.error('❌ Center nav not found'); return; }

    const tabs = Array.from(nav.querySelectorAll('button'))
        .filter(t => !SKIP.includes(clean(t.textContent)));

    const prog = window.location.href.match(/SP\d+/)?.[0] || 'Unknown';
    console.log(`\n🚀 Exporting ${tabs.length} centers for ${prog}…\n`);

    let exported = 0, skipped = 0;

    for (let i = 0; i < tabs.length; i++) {
        const name = clean(tabs[i].textContent);
        currentCenterName = safe(name);             // used by the download setter
        tabs[i].click();
        await sleep(1500);

        const exportBtn = document.querySelector('button.center-export-btn');
        if (exportBtn) {
            exportBtn.click();
            exported++;
        } else {
            console.warn(`   ⚠ No export button for "${name}" — skipping`);
            skipped++;
        }

        console.log(`✓ ${i + 1}/${tabs.length}: ${name}`);
        await sleep(1500);
    }

    // =====================================================================
    //  DEACTIVATE & RESTORE
    // =====================================================================
    interceptActive = false;

    URL.createObjectURL = _createObjectURL;
    delete HTMLAnchorElement.prototype.click;        // remove own prop → inherit again
    HTMLElement.prototype.click = _htmlClick;
    EventTarget.prototype.dispatchEvent = _dispatch;
    document.removeEventListener('click', onCapturingClick, true);
    if (_msSave)       navigator.msSaveBlob       = _msSave;
    if (_msSaveOrOpen) navigator.msSaveOrOpenBlob = _msSaveOrOpen;
    if (_showSavePicker) window.showSaveFilePicker = _showSavePicker;
    if (_downloadDesc) {
        Object.defineProperty(HTMLAnchorElement.prototype, 'download', _downloadDesc);
    }

    // =====================================================================
    //  GENERATE ZIP & SINGLE DOWNLOAD
    // =====================================================================
    if (capturedFiles.length === 0) {
        console.warn('⚠ No files were captured — nothing to bundle.');
        console.log(`   (${exported} export button(s) clicked, ${skipped} skipped)`);
        console.log('');
        console.log('💡 Debugging tips:');
        console.log('   1. Open the Network tab, manually click one Export button, and look');
        console.log('      for the download request — is it a server-side file or client blob?');
        console.log('   2. In the Console, run: URL.createObjectURL.toString()');
        console.log('      If it says [native code], this script\'s hooks were cleared before export.');
        console.log('   3. Check if the app uses an <iframe> for downloads.');
        return;
    }

    console.log(`\n📦 Bundling ${capturedFiles.length} file(s) into ZIP…\n`);
    capturedFiles.forEach((f, i) =>
        console.log(`   ${i + 1}. ${f.filename} (${(f.blob.size / 1024).toFixed(1)} KB)`)
    );

    try {
        const blob = await zip.generateAsync({ type: 'blob' });
        const url  = _createObjectURL(blob);
        const a    = document.createElement('a');
        a.setAttribute('href', url);
        a.setAttribute('download', `PORBs_${prog}.zip`);
        document.body.appendChild(a);
        _htmlClick.call(a);                         // one single browser download
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);

        const totalKB = capturedFiles.reduce((s, f) => s + f.blob.size, 0) / 1024;
        console.log(`\n✅ Done! Downloaded PORBs_${prog}.zip (${totalKB.toFixed(0)} KB)`);
        console.log(`   📊 ${exported} exported · ${skipped} skipped · ${capturedFiles.length} files bundled`);
    } catch (e) {
        console.error('❌ ZIP generation failed:', e.message);
    }
})();
