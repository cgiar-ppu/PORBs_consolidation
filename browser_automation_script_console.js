(async function() {
    const SKIP = ['Summary', 'AOW00', 'AOW01', 'AOW02', 'AOW03', 'AOW04', 'AOW05', 'W3/Bilateral', 'MELIA Study', 'Anaplan'];
    const sleep = ms => new Promise(r => setTimeout(r, ms));
    const tabs = Array.from(document.querySelectorAll('[role="tab"]')).filter(t => !SKIP.includes(t.textContent.replace(/check_circle_outline/g, '').trim()));
    const prog = window.location.href.match(/SP\d+/)?.[0] || 'Unknown';
    console.log(`Exporting ${tabs.length} centers for ${prog}...`);
    for (let i = 0; i < tabs.length; i++) {
        const name = tabs[i].textContent.replace(/check_circle_outline/g, '').trim();
        tabs[i].click(); await sleep(1500);
        document.querySelectorAll('button').forEach(b => { if (b.textContent.includes('Export Excel')) b.click(); });
        console.log(`✓ ${i+1}/${tabs.length}: ${name}`);
        await sleep(1500);
    }
    console.log(`✅ ${prog} complete!`);
})();