(async function() {
    // --- Configuration ---
    const SKIP = ['Summary', 'AOW00', 'AOW01', 'AOW02', 'AOW03', 'AOW04', 'AOW05', 'W3/Bilateral', 'MELIA Study', 'Anaplan'];
    const sleep = ms => new Promise(r => setTimeout(r, ms));
    const clean = t => t.replace(/check_circle_outline/g, '').trim();

    // Sanitize for filesystem-safe filenames
    const safe = s => s.replace(/[\\/:*?"<>|]/g, '').replace(/\s*\(.*?\)\s*/g, '').trim().replace(/\s+/g, '_');

    // =====================================================================
    //  1. LOAD JSZip  — needed to bundle all exports into a single .zip
    // =====================================================================
    console.log('📦 Loading JSZip…');
    try {
        await new Promise((resolve, reject) => {
            if (window.JSZip) { resolve(); return; }          // already loaded
            const s = document.createElement('script');
            s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
            s.onload  = resolve;
            s.onerror = () => reject(new Error('Failed to load JSZip from CDN'));
            document.head.appendChild(s);
        });
        console.log('✅ JSZip ready');
    } catch (e) {
        console.error('❌ ' + e.message);
        console.error('   Cannot bundle downloads without JSZip. Aborting.');
        return;
    }

    const zip = new JSZip();
    const capturedFiles = [];           // { filename, blob } entries
    let currentCenterName = '';

    // =====================================================================
    //  2. INTERCEPT BLOB CREATION  — map every blob URL back to its Blob
    // =====================================================================
    const blobMap = new Map();
    const _createObjectURL = URL.createObjectURL.bind(URL);
    URL.createObjectURL = function(obj) {
        const url = _createObjectURL(obj);
        if (obj instanceof Blob) blobMap.set(url, obj);
        return url;
    };

    // =====================================================================
    //  3. INTERCEPT ANCHOR CLICKS  — capture downloads instead of firing them
    //     When the app calls  a.click()  on an anchor whose href is a blob URL
    //     and whose download attribute is set, we grab the blob and suppress the
    //     browser's download dialog. Everything else passes through unchanged.
    // =====================================================================
    const _origClick = HTMLAnchorElement.prototype.click;
    HTMLAnchorElement.prototype.click = function() {
        const href     = this.getAttribute('href');
        const filename = this.getAttribute('download');

        if (filename && href && href.startsWith('blob:')) {
            const blob = blobMap.get(href);
            if (blob) {
                capturedFiles.push({ filename, blob });
                zip.file(filename, blob);
                console.log(`   📥 Captured: ${filename} (${(blob.size / 1024).toFixed(1)} KB)`);
                URL.revokeObjectURL(href);
                return;                                     // ← suppress browser download
            }
            // Blob wasn't in our map (created before the override ran).
            // Let the browser handle it normally so the user still gets the file.
            console.warn(`   ⚠ Blob not in map for "${filename}" — browser download will fire`);
        }
        return _origClick.call(this);
    };

    // =====================================================================
    //  4. PATCH download ATTRIBUTE  — rename files with center display name
    //     (preserved from the original script)
    // =====================================================================
    const proto = HTMLAnchorElement.prototype;
    Object.defineProperty(proto, 'download', {
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

    // =====================================================================
    //  5. NAVIGATE CENTERS & TRIGGER EXPORTS
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
    //  6. RESTORE ORIGINALS  — leave the page's prototypes clean
    // =====================================================================
    URL.createObjectURL = _createObjectURL;
    HTMLAnchorElement.prototype.click = _origClick;

    // =====================================================================
    //  7. GENERATE ZIP & SINGLE DOWNLOAD
    // =====================================================================
    if (capturedFiles.length === 0) {
        console.warn('⚠ No files were captured — nothing to bundle.');
        console.log(`   (${exported} export button(s) clicked, ${skipped} skipped)`);
        console.log('💡 The app may use a download method this script doesn\'t intercept.');
        console.log('   Try checking the Network tab for the export request format.');
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
        a.href     = url;
        a.download = `PORBs_${prog}.zip`;
        document.body.appendChild(a);
        _origClick.call(a);                         // one single browser download
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);

        const totalKB = capturedFiles.reduce((s, f) => s + f.blob.size, 0) / 1024;
        console.log(`\n✅ Done! Downloaded PORBs_${prog}.zip (${totalKB.toFixed(0)} KB)`);
        console.log(`   📊 ${exported} exported · ${skipped} skipped · ${capturedFiles.length} files bundled`);
    } catch (e) {
        console.error('❌ ZIP generation failed:', e.message);
        console.log('💡 Try re-running the script. If it persists, check the console for details.');
    }
})();
