const XLSX = require('xlsx');
const wb = XLSX.readFile('c:/Users/melanie.v/.gemini/antigravity/scratch/facturation-du-social/Exemple Facturation MATHEO 122025.xls');
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

console.log('Total rows (including header and blanks):', data.length);

const sirens = data.slice(1).map((r, i) => {
    let s = String(r[1] || '').trim();
    return { row: i + 2, siren: s, raw: r };
}).filter(x => x.siren !== '');

const uniqueSirens = new Set(sirens.map(x => x.siren));
console.log('Unique non-empty SIRENs:', uniqueSirens.size);

const counts = {};
sirens.forEach(s => {
    if (!counts[s.siren]) counts[s.siren] = [];
    counts[s.siren].push(s.row);
});

const duplicates = Object.entries(counts).filter(([k, v]) => v.length > 1);
console.log('Duplicate SIRENs (appearing on multiple rows):');
duplicates.forEach(([siren, rows]) => {
    console.log(`- SIREN ${siren} appears on rows: ${rows.join(', ')}`);
});
