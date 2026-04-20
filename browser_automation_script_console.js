(async function() {
    const SKIP = ['Summary', 'AOW00', 'AOW01', 'AOW02', 'AOW03', 'AOW04', 'AOW05', 'W3/Bilateral', 'MELIA Study', 'Anaplan'];
    const sleep = ms => new Promise(r => setTimeout(r, ms));
    const clean = t => t.replace(/check_circle_outline/g, '').trim();

    // Sanitize for filesystem-safe filenames
    const safe = s => s.replace(/[\\/:*?"<>|]/g, '').replace(/\s*\(.*?\)\s*/g, '').trim().replace(/\s+/g, '_');

    // --- Intercept the download filename ---
    // The app sets a.download = "Draft_PORB_SP01_center<ID>.xlsx".
    // We replace "center<ID>" with the current center's display name.
    let currentCenterName = '';
    const proto = HTMLAnchorElement.prototype;
    const origDesc = Object.getOwnPropertyDescriptor(proto, 'download')
        || Object.getOwnPropertyDescriptor(Element.prototype, 'download');
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

    const nav = document.querySelector('nav.porb-nav--centers');
    if (!nav) { console.error('❌ Center nav not found'); return; }

    const tabs = Array.from(nav.querySelectorAll('button'))
        .filter(t => !SKIP.includes(clean(t.textContent)));

    const prog = window.location.href.match(/SP\d+/)?.[0] || 'Unknown';
    console.log(`Exporting ${tabs.length} centers for ${prog}...`);

    for (let i = 0; i < tabs.length; i++) {
        const name = clean(tabs[i].textContent);
        currentCenterName = safe(name);          // <-- used by the patched setter
        tabs[i].click();
        await sleep(1500);

        const exportBtn = document.querySelector('button.center-export-btn');
        if (exportBtn) exportBtn.click();
        else console.warn(`⚠ No Export Center button for ${name}`);

        console.log(`✓ ${i+1}/${tabs.length}: ${name} → Draft_PORB_${prog}_${currentCenterName}.xlsx`);
        await sleep(1500);
    }
    console.log(`✅ ${prog} complete!`);
})();
