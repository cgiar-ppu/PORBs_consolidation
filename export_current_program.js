/**
 * CGIAR PORB - Export All Centers for CURRENT Program
 * ====================================================
 *
 * HOW TO USE:
 * 1. Navigate to any program's submission page (e.g., SP01, SP02, etc.)
 * 2. Open Chrome DevTools (F12 or Cmd+Option+I)
 * 3. Go to the Console tab
 * 4. Copy and paste this script
 * 5. Press Enter
 * 6. Repeat for each program (SP01 through SP13)
 *
 * The script exports all centers for the currently loaded program.
 */

(async function exportCurrentProgram() {
    const SKIP_TABS = ['Summary', 'AOW00', 'AOW01', 'AOW02', 'AOW03', 'AOW04', 'AOW05',
                       'W3/Bilateral', 'MELIA Study', 'Anaplan'];

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function getCenterTabs() {
        const tabs = document.querySelectorAll('[role="tab"]');
        return Array.from(tabs).filter(tab => {
            const text = tab.textContent.replace(/check_circle_outline/g, '').trim();
            return !SKIP_TABS.includes(text) && text.length > 0;
        });
    }

    function clickExportButton() {
        const buttons = document.querySelectorAll('button');
        for (const btn of buttons) {
            if (btn.textContent.includes('Export Excel')) {
                btn.click();
                return true;
            }
        }
        return false;
    }

    // Get current program from URL
    const urlMatch = window.location.href.match(/program\/(\d+)\/(SP\d+)/);
    const programCode = urlMatch ? urlMatch[2] : 'Unknown';

    console.log('='.repeat(50));
    console.log(`Exporting all centers for: ${programCode}`);
    console.log('='.repeat(50));

    // Wait for page to be ready
    await sleep(1000);

    const tabs = getCenterTabs();
    console.log(`Found ${tabs.length} center tabs to export`);

    let exported = 0;

    for (let i = 0; i < tabs.length; i++) {
        const tab = tabs[i];
        const centerName = tab.textContent.replace(/check_circle_outline/g, '').trim();

        console.log(`[${i + 1}/${tabs.length}] Exporting: ${centerName}...`);

        // Click the tab
        tab.click();
        await sleep(1500);

        // Wait for content to load
        let attempts = 0;
        while (attempts < 10) {
            const loading = document.body.textContent.includes('Loading');
            if (!loading) break;
            await sleep(300);
            attempts++;
        }
        await sleep(500);

        // Click Export Excel
        if (clickExportButton()) {
            console.log(`  ✓ ${programCode}-${centerName} exported`);
            exported++;
        } else {
            console.log(`  ✗ Export button not found`);
        }

        // Wait between exports
        await sleep(1500);
    }

    console.log('\n' + '='.repeat(50));
    console.log(`✅ ${programCode} COMPLETE: ${exported}/${tabs.length} exports`);
    console.log('='.repeat(50));
    console.log('\nNext: Navigate to the next program and run this script again.');
    console.log('Programs: SP01→SP02→SP03→SP04→SP05→SP06→SP07→SP08→SP09→SP10→SP11→SP12→SP13');

    return {program: programCode, exported, total: tabs.length};
})();
