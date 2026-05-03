const express = require('express');
const cors    = require('cors');
const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const app  = express();
const PORT = 3000;
app.use(cors());
app.use(express.json());
app.use(express.static('frontend'));

const DATA_DIR  = path.join(__dirname, 'data');
const EXCEL_PUT = path.join(DATA_DIR, 'kalkulator.xlsx');
const EUR_BAM   = 1.95583;

if (!fs.existsSync(DATA_DIR)) { fs.mkdirSync(DATA_DIR); }

const RACUNI_DIR = path.join(__dirname, 'racuni');
if (!fs.existsSync(RACUNI_DIR)) { fs.mkdirSync(RACUNI_DIR); console.log('📁 Kreiran folder: racuni/'); }

// ─── Generator HTML računa ────────────────────────────────────────
function generirajRacun(podaci) {
  const {
    id, ime, drzava, tablice, vozilo, osobe, djeca, pas, sator,
    dani, eur, placenoEUR, ostatakEUR, status,
    uplate, komentar, datum, timestamp: ts,
    uplataIznos, uplataNapomena, uplataVrijeme,
  } = podaci;

  const placenoBAM = parseFloat((placenoEUR * EUR_BAM).toFixed(2));
  const ostatakBAM = parseFloat((ostatakEUR * EUR_BAM).toFixed(2));
  const bamTotal   = parseFloat((eur * EUR_BAM).toFixed(2));

  const statusBoja  = status?.includes('✅') ? '#3ecf8e'
                    : status?.includes('Djelim') ? '#f5a623' : '#e8534a';
  const statusBg    = status?.includes('✅') ? 'rgba(62,207,142,0.12)'
                    : status?.includes('Djelim') ? 'rgba(245,166,35,0.12)' : 'rgba(232,83,74,0.12)';
  const statusBorder= status?.includes('✅') ? 'rgba(62,207,142,0.3)'
                    : status?.includes('Djelim') ? 'rgba(245,166,35,0.3)' : 'rgba(232,83,74,0.3)';

  const cijenaVozilo = podaci.cijenaOsoba != null
    ? (eur/dani - osobe*(podaci.cijenaOsoba||5) - djeca*(podaci.cijaneDijete||2.5)
       - pas*(podaci.cijanePas||2) - sator*(podaci.cijaneSator||5))
    : eur/dani;
  const stavke = [];
  stavke.push({ naziv: `Vozilo — ${vozilo}`, kol: dani, jed: 'dan', cij: parseFloat(cijenaVozilo.toFixed(2)) });
  if (osobe > 0) stavke.push({ naziv: `Odrasli (${osobe} os.)`, kol: dani, jed: 'dan', cij: parseFloat(((podaci.cijenaOsoba||5)*osobe).toFixed(2)) });
  if (djeca  > 0) stavke.push({ naziv: `Djeca (${djeca})`, kol: dani, jed: 'dan', cij: parseFloat(((podaci.cijaneDijete||2.5)*djeca).toFixed(2)) });
  if (pas   > 0) stavke.push({ naziv: `Psi 🐕 (${pas}×)`,   kol: dani, jed: 'dan', cij: (podaci.cijanePas||2)*pas });
  if (sator > 0) stavke.push({ naziv: `Šatori ⛺ (${sator}×)`, kol: dani, jed: 'dan', cij: (podaci.cijaneSator||5)*sator });

  const stavkeHTML = stavke.map((s,i) => `
    <tr class="${i%2===0?'row-even':'row-odd'}">
      <td class="td-naziv">${s.naziv}</td>
      <td class="td-center">${s.kol} ${s.jed}</td>
      <td class="td-right">${s.cij.toFixed(2)} €</td>
      <td class="td-right td-iznos">${(s.cij*s.kol).toFixed(2)} €</td>
    </tr>`).join('');

  const uplateHTML = (uplate && uplate !== '—')
    ? uplate.split('\n').map(u => `<div class="uplata-red">✓ ${u}</div>`).join('')
    : '<div class="uplata-red muted">—</div>';

  const novaUplataHTML = uplataIznos ? `
    <div class="nova-uplata">
      <span class="nova-znak">↳</span>
      <strong>${parseFloat(uplataIznos).toFixed(2)} €</strong>
      ${uplataNapomena ? `<span class="muted"> — ${uplataNapomena}</span>` : ''}
      <span class="muted"> (${uplataVrijeme || ts})</span>
    </div>` : '';

  const sada = new Date().toLocaleString('bs-BA', {
    day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit',
  });

  return `<!DOCTYPE html>
<html lang="bs">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Račun ${id} — ${ime}</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet"/>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg:      #0b0c0e;
    --surface: #13151a;
    --card:    #1a1d24;
    --border:  #252830;
    --accent:  #f5a623;
    --green:   #3ecf8e;
    --blue:    #5b9cf6;
    --red:     #e8534a;
    --text:    #eceef2;
    --muted:   #6b7280;
    --sub:     #9ca3af;
  }

  body {
    font-family: 'DM Sans', 'Segoe UI', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    padding-bottom: 80px;
  }

  .wrap {
    max-width: 560px;
    margin: 0 auto;
    padding: 40px 32px 48px;
  }

  /* ── Header ───────────────────────────────────── */
  .header {
    text-align: center;
    padding-bottom: 28px;
    margin-bottom: 28px;
    border-bottom: 1px solid var(--border);
    position: relative;
  }
  .header::after {
    content: '';
    position: absolute;
    bottom: -1px; left: 50%; transform: translateX(-50%);
    width: 60px; height: 2px;
    background: var(--accent);
  }
  .logo-icon { font-size: 36px; margin-bottom: 8px; display: block; }
  .header h1 {
    font-family: 'Syne', sans-serif;
    font-size: 26px; font-weight: 800;
    letter-spacing: -0.5px;
    color: var(--text);
    margin-bottom: 4px;
  }
  .header .subtitle {
    font-size: 12px; color: var(--muted);
    letter-spacing: 0.5px; text-transform: uppercase;
    margin-bottom: 12px;
  }
  .racun-id-badge {
    display: inline-flex; align-items: center; gap: 8px;
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 100px; padding: 5px 14px;
    font-size: 11px; color: var(--muted);
    font-family: 'DM Mono', monospace;
  }
  .racun-id-badge strong { color: var(--accent); font-size: 12px; }

  /* ── Gost info ────────────────────────────────── */
  .gost-card {
    background: var(--card); border: 1px solid var(--border);
    border-radius: 16px; padding: 22px 24px; margin-bottom: 20px;
    position: relative; overflow: hidden;
  }
  .gost-card::before {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg, var(--accent), transparent);
  }
  .gost-name {
    font-family: 'Syne', sans-serif;
    font-size: 20px; font-weight: 800;
    color: var(--text); margin-bottom: 16px;
    letter-spacing: -0.3px;
  }
  .info-grid {
    display: grid; grid-template-columns: 1fr 1fr;
    gap: 12px 24px;
  }
  .info-item {}
  .info-item .lbl {
    font-size: 10px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 1px;
    color: var(--muted); margin-bottom: 3px; display: block;
  }
  .info-item .val {
    font-size: 14px; font-weight: 600; color: var(--text);
  }
  .komentar-box {
    margin-top: 14px; padding: 10px 14px;
    background: var(--surface); border-radius: 8px;
    border-left: 2px solid var(--border);
    font-size: 12px; color: var(--sub); font-style: italic;
  }

  /* ── Stavke ───────────────────────────────────── */
  .section-title {
    font-family: 'Syne', sans-serif;
    font-size: 10px; font-weight: 700;
    letter-spacing: 2px; text-transform: uppercase;
    color: var(--accent); margin-bottom: 12px;
    display: flex; align-items: center; gap: 10px;
  }
  .section-title::after {
    content: ''; flex: 1; height: 1px; background: var(--border);
  }

  .stavke-card {
    background: var(--card); border: 1px solid var(--border);
    border-radius: 16px; overflow: hidden; margin-bottom: 20px;
  }
  table { width: 100%; border-collapse: collapse; }
  thead tr {
    background: var(--surface);
    border-bottom: 1px solid var(--border);
  }
  th {
    font-size: 10px; font-weight: 700;
    text-transform: uppercase; letter-spacing: 1px;
    color: var(--muted); padding: 12px 16px; text-align: left;
  }
  .row-even { background: var(--card); }
  .row-odd  { background: rgba(255,255,255,0.015); }
  td { padding: 13px 16px; border-bottom: 1px solid var(--border); }
  tr:last-child td { border-bottom: none; }
  .td-naziv  { font-size: 14px; font-weight: 500; color: var(--text); }
  .td-center { text-align: center; font-size: 13px; color: var(--sub); }
  .td-right  { text-align: right; font-size: 13px; color: var(--sub); }
  .td-iznos  { font-weight: 700; font-size: 14px; color: var(--text); }

  /* ── Totals ───────────────────────────────────── */
  .totals-card {
    background: var(--card); border: 1px solid var(--border);
    border-radius: 16px; overflow: hidden; margin-bottom: 20px;
  }
  .total-row {
    display: flex; justify-content: space-between; align-items: center;
    padding: 14px 20px; border-bottom: 1px solid var(--border);
  }
  .total-row:last-child { border-bottom: none; }
  .total-row .lbl { font-size: 13px; color: var(--sub); font-weight: 500; }
  .total-row .val { font-size: 14px; font-weight: 700; color: var(--text); }
  .total-row.big   { background: var(--surface); padding: 18px 20px; }
  .total-row.big .lbl { font-family: 'Syne', sans-serif; font-size: 14px; font-weight: 800; color: var(--text); letter-spacing: 0.5px; }
  .total-row.big .val { font-family: 'Syne', sans-serif; font-size: 22px; font-weight: 800; color: var(--accent); }
  .total-row.bam .val { color: var(--blue); font-size: 13px; }
  .total-row.placeno .val { color: var(--green); }
  .total-row.placeno .lbl { color: var(--green); }
  .total-row.ostatak .val { color: var(--red); font-size: 16px; }
  .total-row.ostatak .lbl { color: var(--red); font-weight: 600; }
  .total-row.ostatak-0 .val { color: var(--green); font-size: 16px; }
  .total-row.ostatak-0 .lbl { color: var(--green); font-weight: 600; }

  /* ── Status ───────────────────────────────────── */
  .status-wrap { text-align: center; margin: 6px 0 20px; }
  .status-badge {
    display: inline-block; padding: 8px 22px; border-radius: 100px;
    font-family: 'Syne', sans-serif; font-size: 13px; font-weight: 800;
    background: ${statusBg}; color: ${statusBoja};
    border: 1px solid ${statusBorder}; letter-spacing: 0.3px;
  }

  /* ── Uplate ───────────────────────────────────── */
  .uplate-card {
    background: var(--card); border: 1px solid var(--border);
    border-radius: 16px; padding: 18px 20px; margin-bottom: 20px;
  }
  .uplata-red {
    font-size: 13px; color: var(--sub); padding: 5px 0;
    border-bottom: 1px solid var(--border); line-height: 1.5;
  }
  .uplata-red:last-child { border-bottom: none; }
  .muted { color: var(--muted); }
  .nova-uplata {
    margin-top: 12px; padding: 12px 14px;
    background: rgba(62,207,142,0.08); border: 1px solid rgba(62,207,142,0.25);
    border-radius: 10px; font-size: 13px; color: var(--green);
    line-height: 1.5;
  }
  .nova-uplata strong { font-size: 15px; }
  .nova-znak { font-size: 16px; margin-right: 4px; }

  /* ── Footer ───────────────────────────────────── */
  .footer {
    text-align: center; padding: 24px 0 0;
    border-top: 1px solid var(--border);
  }
  .footer .hvala {
    font-family: 'Syne', sans-serif;
    font-size: 18px; font-weight: 800; color: var(--text);
    margin-bottom: 10px;
  }
  .footer .meta {
    font-size: 12px; color: var(--muted); line-height: 1.8;
  }
  .footer .stampano {
    margin-top: 10px; font-size: 11px; color: var(--border);
    font-family: monospace;
  }

  /* ── Print bar ────────────────────────────────── */
  .print-bar {
    position: fixed; bottom: 0; left: 0; right: 0;
    background: var(--surface); border-top: 1px solid var(--border);
    padding: 14px 24px; display: flex; gap: 12px;
    justify-content: center; z-index: 100;
  }
  .btn-print {
    font-family: 'Syne', sans-serif;
    background: var(--accent); color: #000;
    border: none; font-weight: 800; font-size: 14px;
    padding: 12px 32px; border-radius: 10px;
    cursor: pointer; transition: background 0.15s;
    letter-spacing: 0.3px;
  }
  .btn-print:hover { background: #fbb740; }
  .btn-close {
    font-family: 'Syne', sans-serif;
    background: var(--card); color: var(--muted);
    border: 1px solid var(--border); font-weight: 700; font-size: 13px;
    padding: 12px 24px; border-radius: 10px;
    cursor: pointer; transition: all 0.15s;
  }
  .btn-close:hover { border-color: var(--red); color: var(--red); }

  /* ── Print CSS ────────────────────────────────── */
  @media print {
    .print-bar { display: none !important; }
    body { background: #fff !important; color: #000 !important; padding-bottom: 0; }
    .wrap { padding: 10px 16px; }
    .gost-card, .stavke-card, .totals-card, .uplate-card {
      background: #fff !important; border: 1px solid #ccc !important;
      border-radius: 8px !important;
    }
    .gost-card::before { display: none; }
    .header::after { background: #000; }
    .gost-name, .header h1, .section-title { color: #000 !important; }
    .td-naziv, .td-iznos, .total-row .lbl, .total-row .val { color: #000 !important; }
    .td-center, .td-right, .uplata-red, .info-item .val { color: #333 !important; }
    .info-item .lbl, .muted, .footer .meta, .section-title { color: #666 !important; }
    .total-row.big { background: #f5f5f5 !important; }
    .total-row.big .val { color: #c47d00 !important; }
    .total-row.placeno .val, .total-row.ostatak-0 .val { color: #166534 !important; }
    .total-row.ostatak .val { color: #b91c1c !important; }
    .status-badge { background: #f5f5f5 !important; color: #333 !important; border-color: #ccc !important; }
    thead tr { background: #f5f5f5 !important; }
    .row-odd { background: #fafafa !important; }
    .nova-uplata { background: #f0fdf4 !important; border-color: #86efac !important; color: #166534 !important; }
    @page { margin: 10mm; size: A4; }
  }
</style>
</head>
<body>
<div class="wrap">

  <!-- Header -->
  <div class="header">
    <span class="logo-icon">🏕️</span>
    <h1>KAMPING</h1>
    <div class="subtitle">Potvrda o boravku / Receipt</div>
    <div class="racun-id-badge">Račun: <strong>${id}</strong> &nbsp;·&nbsp; ${datum}</div>
  </div>

  <!-- Gost -->
  <div class="section-title">Podaci gosta</div>
  <div class="gost-card">
    <div class="gost-name">${ime}</div>
    <div class="info-grid">
      <div class="info-item"><span class="lbl">Država</span><span class="val">${drzava}</span></div>
      <div class="info-item"><span class="lbl">Tablice</span><span class="val">${tablice && tablice !== '—' ? tablice : '—'}</span></div>
      <div class="info-item"><span class="lbl">Vozilo</span><span class="val">${vozilo}</span></div>
      <div class="info-item"><span class="lbl">Broj dana</span><span class="val">${dani} dan(a)</span></div>
      <div class="info-item"><span class="lbl">Datum dolaska</span><span class="val">${datum}</span></div>
      <div class="info-item"><span class="lbl">Putnici</span><span class="val">${osobe} odrasli${djeca > 0 ? `, ${djeca} djece` : ''}${pas>0?` + ${pas}🐕`:''}${sator>0?` + ${sator}⛺`:''}</span></div>
    </div>
    ${komentar && komentar !== '—' ? `<div class="komentar-box">📝 ${komentar}</div>` : ''}
  </div>

  <!-- Stavke -->
  <div class="section-title">Specifikacija</div>
  <div class="stavke-card">
    <table>
      <thead>
        <tr>
          <th>Stavka</th>
          <th style="text-align:center">Kol.</th>
          <th style="text-align:right">Cijena/dan</th>
          <th style="text-align:right">Iznos</th>
        </tr>
      </thead>
      <tbody>${stavkeHTML}</tbody>
    </table>
  </div>

  <!-- Totals -->
  <div class="section-title">Obračun</div>
  <div class="totals-card">
    <div class="total-row big">
      <span class="lbl">UKUPNO</span>
      <span class="val">${eur.toFixed(2)} €</span>
    </div>
    <div class="total-row bam">
      <span class="lbl" style="color:var(--muted)">Iznos u KM</span>
      <span class="val">${bamTotal.toFixed(2)} KM</span>
    </div>
    <div class="total-row placeno">
      <span class="lbl">Plaćeno</span>
      <span class="val">+ ${placenoEUR.toFixed(2)} € &nbsp;<span style="font-size:12px;font-weight:400">(${placenoBAM.toFixed(2)} KM)</span></span>
    </div>
    <div class="total-row ${ostatakEUR < 0.01 ? 'ostatak-0' : 'ostatak'}">
      <span class="lbl">Ostatak</span>
      <span class="val">${ostatakEUR < 0.01 ? '0.00 € ✅' : ostatakEUR.toFixed(2) + ' € <span style="font-size:12px;font-weight:400">(' + ostatakBAM.toFixed(2) + ' KM)</span>'}</span>
    </div>
  </div>

  <div class="status-wrap">
    <span class="status-badge">${status}</span>
  </div>

  <!-- Uplate -->
  <div class="section-title">Historija plaćanja</div>
  <div class="uplate-card">
    ${uplateHTML}
    ${novaUplataHTML}
  </div>

  <!-- Footer -->
  <div class="footer">
    <div class="hvala">Hvala na posjeti! 🙏</div>
    <div class="meta">
      Kurs: 1 € = ${EUR_BAM} KM (fiksni kurs)<br/>
      Broj računa: ${id}
    </div>
    <div class="stampano">Štampano: ${sada}</div>
  </div>

</div>

<!-- Print bar -->
<div class="print-bar no-print">
  <button class="btn-print" onclick="window.print()">🖨️ Štampaj / Print</button>
  <button class="btn-close" onclick="window.close()">✕ Zatvori</button>
</div>
</body>
</html>`;
}

// ─── GET /racun/:id ───────────────────────────────────────────────
// Servira HTML račun direktno u browser
app.get('/racun/:id', (req, res) => {
  const fajlPut = path.join(RACUNI_DIR, `${req.params.id}.html`);
  if (!fs.existsSync(fajlPut)) {
    return res.status(404).send('<h2>Račun nije pronađen.</h2>');
  }
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.send(fs.readFileSync(fajlPut, 'utf8'));
});

let VOZILO_CIJENA = { auto:5, motor:3, auto_sator:10, van:15, van_l:15, auto_karavan:20, motohome:20 };
const VOZILO_NAZIV  = { auto:'Auto', motor:'Motor', auto_sator:'Auto + Šator', van:'Van', van_l:'Van L', auto_karavan:'Auto + Karavan', motohome:'Motohome' };
let CIJENA        = { osoba:5, dijete:2.5, pas:2, sator:5 };

const KOLONE = [
  { header:'ID',              key:'id',         width:12 },
  { header:'Datum',           key:'datum',       width:14 },
  { header:'Ime i Prezime',   key:'ime',         width:22 },
  { header:'Država',          key:'drzava',      width:18 },
  { header:'Tablice',         key:'tablice',     width:14 },
  { header:'Vrsta Vozila',    key:'vozilo',      width:18 },
  { header:'Oprema',          key:'oprema',      width:28 },
  { header:'Osobe',           key:'osobe',       width:8  },
  { header:'Djeca',           key:'djeca',       width:8  },
  { header:'Pas',             key:'pas',         width:6  },
  { header:'Šator',           key:'sator',       width:8  },
  { header:'Dana',            key:'dani',        width:7  },
  { header:'Ukupno (€)',      key:'eur',         width:13 },
  { header:'Ukupno (KM)',     key:'bam',         width:13 },
  { header:'Plaćeno (€)',     key:'placenoEUR',  width:13 },
  { header:'Plaćeno (KM)',    key:'placenoBAM',  width:13 },
  { header:'Ostatak (€)',     key:'ostatakEUR',  width:13 },
  { header:'Ostatak (KM)',    key:'ostatakBAM',  width:13 },
  { header:'Status',          key:'status',      width:20 },
  { header:'Historija uplata',key:'uplate',      width:40 },
  { header:'Komentar',        key:'komentar',    width:40 },
  { header:'Vrijeme unosa',   key:'timestamp',   width:22 },
];

function danasniDatum() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function timestamp() {
  return new Date().toLocaleString('bs-BA', { day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit', second:'2-digit' });
}
function genId() {
  return Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2,5).toUpperCase();
}
function izracunajCijenu(data) {
  const pas   = parseInt(data.pas)   || 0;
  const sator = parseInt(data.sator) || 0;
  const perDay = (VOZILO_CIJENA[data.vozilo]??0) + data.osobe*CIJENA.osoba + data.djeca*CIJENA.dijete
               + pas*CIJENA.pas + sator*CIJENA.sator;
  return { eur: parseFloat((perDay*data.dani).toFixed(2)), bam: parseFloat((perDay*data.dani*EUR_BAM).toFixed(2)) };
}
function validiraj(data) {
  if (!data.ime    || data.ime.trim().length < 2)    return "Ime je obavezno.";
  if (!data.drzava || data.drzava.trim().length < 2) return "Država je obavezna.";
  if (!VOZILO_CIJENA.hasOwnProperty(data.vozilo))    return "Neispravna vrsta vozila.";
  if (!Number.isFinite(data.osobe) || data.osobe < 0) return "Nevažeći broj osoba.";
  if (!Number.isFinite(data.djeca) || data.djeca < 0) return "Nevažeći broj djece.";
  if (!Number.isInteger(data.dani) || data.dani  < 1) return "Broj dana mora biti >= 1.";
  return null;
}
function izracunajStatus(placeno, ukupno) {
  if (placeno <= 0)                    return 'Nije plaćeno';
  if (Math.max(0,ukupno-placeno)<0.01) return 'Plaćeno ✅';
  return `Djelimično (${Math.round((placeno/ukupno)*100)}%)`;
}
function stilujRed(row, status) {
  const isEven = row.number % 2 === 0;
  row.eachCell({ includeEmpty:true }, cell => {
    cell.fill      = { type:'pattern', pattern:'solid', fgColor:{ argb: isEven?'FF16181C':'FF1C1F26' } };
    cell.font      = { color:{ argb:'FFF0F0F0' }, size:10 };
    cell.alignment = { vertical:'middle', horizontal:'center', wrapText:true };
  });
  row.height = 22;
  const sc = (key,argb,bold=true,fmt=false) => {
    const c = row.getCell(key);
    c.font = { bold, color:{argb}, size: bold?11:10 };
    if (fmt) c.numFmt = '#,##0.00';
  };
  sc('eur','FFF5A623',true,true); sc('bam','FF5B9CF6',true,true);
  sc('placenoEUR','FF3ECF8E',true,true); sc('placenoBAM','FF3ECF8E',true,true);
  sc('ostatakEUR','FFE8534A',true,true); sc('ostatakBAM','FFE8534A',true,true);
  const statusArgb = status?.includes('✅')?'FF3ECF8E': status?.includes('Djelim')?'FFF5A623':'FFE8534A';
  row.getCell('status').font = { bold:true, color:{argb:statusArgb}, size:10 };
  row.getCell('komentar').font      = { italic:true, color:{argb:'FFB0B5C3'}, size:10 };
  row.getCell('komentar').alignment = { vertical:'middle', horizontal:'left', wrapText:true };
  row.getCell('uplate').alignment   = { vertical:'middle', horizontal:'left', wrapText:true };
  row.getCell('uplate').font        = { color:{argb:'FFAAB0C0'}, size:9 };
  row.getCell('tablice').font       = { color:{argb:'FFAAB0C0'}, size:10, italic:true };
}

async function ucitajWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const isNovi   = !fs.existsSync(EXCEL_PUT);
  if (!isNovi) await workbook.xlsx.readFile(EXCEL_PUT);
  let ws = workbook.getWorksheet('Podaci');
  if (!ws) ws = workbook.addWorksheet('Podaci');
  ws.columns = KOLONE;
  if (isNovi || ws.rowCount <= 1) {
    const hRow = ws.getRow(1);
    hRow.eachCell(cell => {
      cell.fill      = { type:'pattern', pattern:'solid', fgColor:{argb:'FF1A1D24'} };
      cell.font      = { bold:true, color:{argb:'FFF5A623'}, size:11 };
      cell.alignment = { vertical:'middle', horizontal:'center', wrapText:true };
      cell.border    = { bottom:{ style:'medium', color:{argb:'FFF5A623'} } };
    });
    hRow.height = 30;
  }
  return { workbook, worksheet: ws };
}

function citajSveRedove(ws) {
  const rows = []; let hm = {};
  ws.eachRow((row, rowNum) => {
    if (rowNum === 1) { row.eachCell((c,col) => { hm[String(c.value)] = col; }); return; }
    const get = k => { const col=hm[k]; if(!col) return ''; const v=row.getCell(col).value; return v!=null?v:''; };
    if (!get('Ime i Prezime')) return;
    rows.push({
      _rowNum:rowNum, id:String(get('ID')), datum:get('Datum'), rb:rowNum-1,
      ime:get('Ime i Prezime'), drzava:get('Država'), tablice:get('Tablice'),
      vozilo:get('Vrsta Vozila'), osobe:get('Osobe'), djeca:get('Djeca'),
      pas:get('Pas'), sator:get('Šator'), dani:get('Dana'),
      eur:parseFloat(get('Ukupno (€)'))||0, bam:parseFloat(get('Ukupno (KM)'))||0,
      placenoEUR:parseFloat(get('Plaćeno (€)'))||0, ostatakEUR:parseFloat(get('Ostatak (€)'))||0,
      status:get('Status'), komentar:get('Komentar'), timestamp:get('Vrijeme unosa'),
    });
  });
  return rows;
}
function sumiraj(rows) {
  const e = rows.reduce((s,r)=>s+r.eur,0);
  return {
    ukupnoEUR:  parseFloat(e.toFixed(2)),
    ukupnoBAM:  parseFloat((e*EUR_BAM).toFixed(2)),
    placenoEUR: parseFloat(rows.reduce((s,r)=>s+r.placenoEUR,0).toFixed(2)),
    ostatakEUR: parseFloat(rows.reduce((s,r)=>s+r.ostatakEUR,0).toFixed(2)),
    brojTurista:rows.length,
  };
}

// ─── Cijene config ────────────────────────────────────────────────
const CIJENE_PUT = path.join(DATA_DIR, 'cijene.json');

const DEFAULT_VOZILO = { auto:5, motor:3, auto_sator:10, van:15, van_l:15, auto_karavan:20, motohome:20 };
const DEFAULT_OPCIJE = { osoba:5, dijete:2.5, pas:2, sator:5 };

function ucitajCijene() {
  try {
    if (fs.existsSync(CIJENE_PUT)) {
      const raw = JSON.parse(fs.readFileSync(CIJENE_PUT, 'utf8'));
      return {
        vozilo: { ...DEFAULT_VOZILO, ...(raw.vozilo || {}) },
        opcije: { ...DEFAULT_OPCIJE, ...(raw.opcije || {}) },
      };
    }
  } catch {}
  return { vozilo: { ...DEFAULT_VOZILO }, opcije: { ...DEFAULT_OPCIJE } };
}

let aktivneCijene = ucitajCijene();

// Osvježi lokalne konstante koje koriste ostale funkcije
function syncCijene() {
  Object.assign(VOZILO_CIJENA, aktivneCijene.vozilo);
  Object.assign(CIJENA, aktivneCijene.opcije);
}

// ─── GET /api/cijene ──────────────────────────────────────────────
app.get('/api/cijene', (req, res) => {
  res.json(aktivneCijene);
});

// ─── POST /api/cijene ─────────────────────────────────────────────
app.post('/api/cijene', (req, res) => {
  const { vozilo, opcije } = req.body;
  if (!vozilo || !opcije) return res.status(400).json({ error: 'Nedostaju podaci.' });
  aktivneCijene = {
    vozilo: { ...DEFAULT_VOZILO, ...vozilo },
    opcije: { ...DEFAULT_OPCIJE, ...opcije },
  };
  syncCijene();
  try {
    fs.writeFileSync(CIJENE_PUT, JSON.stringify(aktivneCijene, null, 2), 'utf8');
    console.log('💰 Cijene ažurirane.');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: 'Greška pri čuvanju cijena.' });
  }
});

// ═══ API ══════════════════════════════════════════════════════════

app.get('/api/status', (_,res) => res.json({ status:'ok', eurBam:EUR_BAM }));

app.get('/api/datumi', async (_,res) => {
  if (!fs.existsSync(EXCEL_PUT)) return res.json({ datumi:[] });
  try {
    const {worksheet} = await ucitajWorkbook();
    const set = new Set(); let hm = {};
    worksheet.eachRow((row,rowNum) => {
      if (rowNum===1) { row.eachCell((c,col)=>{ hm[String(c.value)]=col; }); return; }
      const col=hm['Datum']; if (col) { const v=row.getCell(col).value; if(v) set.add(String(v)); }
    });
    res.json({ datumi:[...set].sort() });
  } catch { res.status(500).json({error:'Greška.'}); }
});

app.get('/api/podaci', async (req,res) => {
  if (!fs.existsSync(EXCEL_PUT)) return res.json({ rows:[], ...sumiraj([]) });
  try {
    const {worksheet} = await ucitajWorkbook();
    let rows = citajSveRedove(worksheet);
    const { datum, search, filter } = req.query;
    if (datum)  rows = rows.filter(r => r.datum === datum);
    if (search) { const q=search.toLowerCase().trim(); rows=rows.filter(r=>r.ime.toLowerCase().includes(q)||r.drzava.toLowerCase().includes(q)||(r.tablice||'').toLowerCase().includes(q)); }
    if (filter==='placeno')    rows=rows.filter(r=>r.status?.includes('✅'));
    if (filter==='neplaceno')  rows=rows.filter(r=>r.status==='Nije plaćeno');
    if (filter==='djelimicno') rows=rows.filter(r=>r.status?.includes('Djelimično'));
    res.json({ rows, ...sumiraj(rows) });
  } catch(err) { console.error(err); res.status(500).json({error:'Greška.'}); }
});

app.get('/api/red/:id', async (req,res) => {
  if (!fs.existsSync(EXCEL_PUT)) return res.status(404).json({error:'Nema fajla.'});
  try {
    const {worksheet} = await ucitajWorkbook();
    let hm={}; let found=null;
    worksheet.eachRow((row,rowNum) => {
      if (rowNum===1) { row.eachCell((c,col)=>{ hm[String(c.value)]=col; }); return; }
      const idCol=hm['ID'];
      if (idCol && String(row.getCell(idCol).value)===req.params.id) {
        const get=k=>{ const col=hm[k]; if(!col) return ''; const v=row.getCell(col).value; return v!=null?v:''; };
        found={ id:get('ID'), ime:get('Ime i Prezime'), tablice:get('Tablice'),
          ukupnoEUR:parseFloat(get('Ukupno (€)'))||0, placenoEUR:parseFloat(get('Plaćeno (€)'))||0,
          ostatakEUR:parseFloat(get('Ostatak (€)'))||0, status:get('Status'),
          dani:get('Dana'), historijaUplata:get('Historija uplata')||'—' };
      }
    });
    if (!found) return res.status(404).json({error:'Red nije pronađen.'});
    res.json(found);
  } catch { res.status(500).json({error:'Greška.'}); }
});

app.post('/api/kalkulator', async (req,res) => {
  const data=req.body;
  const greska=validiraj(data);
  if (greska) return res.status(400).json({error:greska});
  const {eur,bam}=izracunajCijenu(data);
  const placenoEUR=parseFloat((data.placenoEUR??0).toFixed(2));
  const ostatakEUR=parseFloat(Math.max(0,eur-placenoEUR).toFixed(2));
  const status=izracunajStatus(placenoEUR,eur);
  const uplateStr=(data.uplate??[]).map((u,i)=>`${i+1}. ${u.iznos.toFixed(2)}€${u.napomena?' — '+u.napomena:''} (${u.vrijeme})`).join('\n')||'—';
  try {
    const {workbook,worksheet}=await ucitajWorkbook();
    const id=genId(); const datum=danasniDatum();
    const tablice=(data.tablice||'').trim().toUpperCase()||'—';
    const brPasa   = parseInt(data.pas)   || 0;
    const brSatora = parseInt(data.sator) || 0;
    const noviRed=worksheet.addRow({
      id, datum, ime:data.ime, drzava:data.drzava, tablice,
      vozilo:VOZILO_NAZIV[data.vozilo]||data.vozilo,
      osobe:data.osobe, djeca:data.djeca,
      pas:   brPasa   > 0 ? `${brPasa} 🐕`   : 'Ne',
      sator: brSatora > 0 ? `${brSatora} ⛺` : 'Ne',
      dani:data.dani, eur, bam, placenoEUR,
      placenoBAM:parseFloat((placenoEUR*EUR_BAM).toFixed(2)),
      ostatakEUR, ostatakBAM:parseFloat((ostatakEUR*EUR_BAM).toFixed(2)),
      status, uplate:uplateStr, komentar:(data.komentar||'').trim()||'—', timestamp:timestamp(),
    });
    stilujRed(noviRed,status);
    await workbook.xlsx.writeFile(EXCEL_PUT);

    // ── Generiši HTML račun ─────────────────────────────
    const racunHTML = generirajRacun({
      id, ime: data.ime, drzava: data.drzava, tablice,
      vozilo: VOZILO_NAZIV[data.vozilo] || data.vozilo,
      osobe: data.osobe, djeca: data.djeca,
      pas:   brPasa,
      sator: brSatora,
      dani: data.dani, eur, bam, placenoEUR, ostatakEUR, status,
      uplate: uplateStr, komentar: (data.komentar || '').trim() || '—',
      datum, timestamp: timestamp(),
      cijenaOsoba: CIJENA.osoba, cijaneDijete: CIJENA.dijete,
      cijanePas: CIJENA.pas, cijaneSator: CIJENA.sator,
    });
    const racunPut = path.join(RACUNI_DIR, `${id}.html`);
    fs.writeFileSync(racunPut, racunHTML, 'utf8');
    console.log(`🧾 Račun: racuni/${id}.html`);

    console.log(`✅ [${datum}] ${data.ime} | ${eur}€ | ${status}`);
    res.status(200).json({success:true, id, ukupnoEUR:eur, ukupnoBAM:bam, placenoEUR, ostatakEUR, status, racunUrl:`/racun/${id}`});
  } catch(err) { console.error('❌',err); res.status(500).json({error:'Greška pri čuvanju.'}); }
});

app.put('/api/update', async (req,res) => {
  const {id,tipAkcije,iznosUplate,napomenaUplate,dodatniDani,tablice}=req.body;
  if (!id||!tipAkcije) return res.status(400).json({error:'Nedostaju parametri.'});
  if (!fs.existsSync(EXCEL_PUT)) return res.status(404).json({error:'Fajl nije pronađen.'});
  try {
    const {workbook,worksheet}=await ucitajWorkbook();
    let hm={}; let tr=null;
    worksheet.eachRow((row,rowNum) => {
      if (rowNum===1) { row.eachCell((c,col)=>{ hm[String(c.value)]=col; }); return; }
      const idCol=hm['ID']; if (idCol&&String(row.getCell(idCol).value)===id) tr=row;
    });
    if (!tr) return res.status(404).json({error:'Red nije pronađen.'});
    const get=k=>{ const col=hm[k]; if(!col) return ''; const v=tr.getCell(col).value; return v!=null?v:''; };
    const set=(k,v)=>{ const col=hm[k]; if(col) tr.getCell(col).value=v; };
    let poruka=''; let noviStatus=String(get('Status'));

    if (tipAkcije==='uplata'||tipAkcije==='ispravi') {
      const ukupno=parseFloat(get('Ukupno (€)'))||0;
      const stari=parseFloat(get('Plaćeno (€)'))||0;
      const novi=tipAkcije==='ispravi'?parseFloat((iznosUplate||0).toFixed(2)):parseFloat((stari+(iznosUplate||0)).toFixed(2));
      const ostatak=parseFloat(Math.max(0,ukupno-novi).toFixed(2));
      noviStatus=izracunajStatus(novi,ukupno);
      set('Plaćeno (€)',novi); set('Plaćeno (KM)',parseFloat((novi*EUR_BAM).toFixed(2)));
      set('Ostatak (€)',ostatak); set('Ostatak (KM)',parseFloat((ostatak*EUR_BAM).toFixed(2)));
      set('Status',noviStatus);
      const stare=String(get('Historija uplata')||'—');
      const redBr=stare==='—'?1:stare.split('\n').length+1;
      const oznaka=tipAkcije==='ispravi'?`ISPRAVKA → ${novi.toFixed(2)}€`:`${(iznosUplate||0).toFixed(2)}€`;
      const novaUp=`${redBr}. ${oznaka}${napomenaUplate?' — '+napomenaUplate:''} (${timestamp()})`;
      set('Historija uplata', stare==='—'?novaUp:stare+'\n'+novaUp);
      poruka=tipAkcije==='ispravi'?`Ispravka → ${novi.toFixed(2)}€`:`Uplata ${(iznosUplate||0).toFixed(2)}€`;

    } else if (tipAkcije==='produzenje') {
      const stariDani=parseInt(get('Dana'))||1;
      const noviDani=stariDani+parseInt(dodatniDani||1);
      const vNaz=String(get('Vrsta Vozila'));
      const vKey=Object.keys(VOZILO_NAZIV).find(k=>VOZILO_NAZIV[k]===vNaz)||'auto';
      const osobe=parseFloat(get('Osobe'))||0; const djeca=parseFloat(get('Djeca'))||0;
      // pas i sator se čitaju u perDay izračunu iznad
      const oStr=String(get('Pas')||''); const pasBr=parseInt(oStr)||0;
      const sStr=String(get('Šator')||''); const satorBr=parseInt(sStr)||0;
      const perDay=(VOZILO_CIJENA[vKey]??0)+osobe*CIJENA.osoba+djeca*CIJENA.dijete+pasBr*CIJENA.pas+satorBr*CIJENA.sator;
      const novaUk=parseFloat((perDay*noviDani).toFixed(2));
      const stariPl=parseFloat(get('Plaćeno (€)'))||0;
      const noviOst=parseFloat(Math.max(0,novaUk-stariPl).toFixed(2));
      noviStatus=izracunajStatus(stariPl,novaUk);
      set('Dana',noviDani); set('Ukupno (€)',novaUk); set('Ukupno (KM)',parseFloat((novaUk*EUR_BAM).toFixed(2)));
      set('Ostatak (€)',noviOst); set('Ostatak (KM)',parseFloat((noviOst*EUR_BAM).toFixed(2))); set('Status',noviStatus);
      poruka=`Produženo +${dodatniDani} dana`;

    } else if (tipAkcije==='tablice') {
      const t=(tablice||'').trim().toUpperCase();
      if (!t) return res.status(400).json({error:'Nevažeće tablice.'});
      set('Tablice',t); poruka=`Tablice: ${t}`;

    } else return res.status(400).json({error:'Nepoznata akcija.'});

    stilujRed(tr,noviStatus); tr.commit();
    await workbook.xlsx.writeFile(EXCEL_PUT);

    // ── Regeneriši račun pri uplati ─────────────────────
    if (tipAkcije === 'uplata' || tipAkcije === 'ispravi') {
      try {
        const racunPut = path.join(RACUNI_DIR, `${id}.html`);
        if (fs.existsSync(racunPut)) {
          // Učitaj ažurirane podatke iz reda
          const noviPlaceno = parseFloat(get('Plaćeno (€)')) || 0;
          const noviOstatak = parseFloat(get('Ostatak (€)')) || 0;
          const noviUplate  = String(get('Historija uplata') || '—');
          const racunHTML   = generirajRacun({
            id,
            ime:     String(get('Ime i Prezime')),
            drzava:  String(get('Država')),
            tablice: String(get('Tablice')),
            vozilo:  String(get('Vrsta Vozila')),
            osobe:   parseFloat(get('Osobe')) || 0,
            djeca:   parseFloat(get('Djeca')) || 0,
            pas:     String(get('Pas')),
            sator:   String(get('Šator')),
            dani:    parseInt(get('Dana')) || 1,
            eur:     parseFloat(get('Ukupno (€)')) || 0,
            bam:     parseFloat(get('Ukupno (KM)')) || 0,
            placenoEUR: noviPlaceno,
            ostatakEUR: noviOstatak,
            status:  noviStatus,
            uplate:  noviUplate,
            komentar: String(get('Komentar') || '—'),
            datum:   String(get('Datum')),
            timestamp: String(get('Vrijeme unosa')),
            uplataIznos:    tipAkcije === 'uplata' ? iznosUplate : null,
            uplataNapomena: napomenaUplate || '',
            uplataVrijeme:  timestamp(),
            cijenaOsoba: CIJENA.osoba, cijaneDijete: CIJENA.dijete,
            cijanePas: CIJENA.pas, cijaneSator: CIJENA.sator,
          });
          fs.writeFileSync(racunPut, racunHTML, 'utf8');
          console.log(`🧾 Račun ažuriran: racuni/${id}.html`);
        }
      } catch(e) { console.error('⚠️ Greška pri ažuriranju računa:', e.message); }
    }

    console.log(`✏️ [${id}] ${poruka}`);
    res.json({success:true, poruka, racunUrl:`/racun/${id}`});
  } catch(err) { console.error('❌',err); res.status(500).json({error:'Greška pri updateu.'}); }
});

app.delete('/api/red/:id', async (req,res) => {
  if (!fs.existsSync(EXCEL_PUT)) return res.status(404).json({error:'Fajl nije pronađen.'});
  try {
    const {workbook,worksheet}=await ucitajWorkbook();
    let hm={}; let trNum=null;
    worksheet.eachRow((row,rowNum) => {
      if (rowNum===1) { row.eachCell((c,col)=>{ hm[String(c.value)]=col; }); return; }
      const idCol=hm['ID']; if (idCol&&String(row.getCell(idCol).value)===req.params.id) trNum=rowNum;
    });
    if (!trNum) return res.status(404).json({error:'Red nije pronađen.'});
    worksheet.spliceRows(trNum,1);
    await workbook.xlsx.writeFile(EXCEL_PUT);
    res.json({success:true});
  } catch(err) { console.error('❌',err); res.status(500).json({error:'Greška.'}); }
});

app.delete('/api/datum/:datum', async (req,res) => {
  if (!fs.existsSync(EXCEL_PUT)) return res.status(404).json({error:'Fajl nije pronađen.'});
  try {
    const {workbook,worksheet}=await ucitajWorkbook();
    let hm={}; const toDelete=[];
    worksheet.eachRow((row,rowNum) => {
      if (rowNum===1) { row.eachCell((c,col)=>{ hm[String(c.value)]=col; }); return; }
      const dc=hm['Datum']; if (dc&&String(row.getCell(dc).value)===req.params.datum) toDelete.push(rowNum);
    });
    for (let i=toDelete.length-1;i>=0;i--) worksheet.spliceRows(toDelete[i],1);
    await workbook.xlsx.writeFile(EXCEL_PUT);
    res.json({success:true, obrisano:toDelete.length});
  } catch(err) { console.error('❌',err); res.status(500).json({error:'Greška.'}); }
});

// ─── GET /api/izvjestaj ───────────────────────────────────────────
// tip: dnevni | sedmicni | mjesecni | godisnji | privremeni
// za privremeni: datumOd + datumDo
app.get('/api/izvjestaj', async (req, res) => {
  if (!fs.existsSync(EXCEL_PUT)) return res.json({ rows: [], sumiraj: {}, period: {} });
  try {
    const { tip, datumOd, datumDo } = req.query;
    const { worksheet } = await ucitajWorkbook();
    let rows = citajSveRedove(worksheet);

    const danas = danasniDatum();
    const [gy, gm, gd] = danas.split('-').map(Number);

    let od, doo, periodNaziv;

    if (tip === 'dnevni') {
      od = danas; doo = danas;
      periodNaziv = `Dnevni izvještaj — ${gd}.${String(gm).padStart(2,'0')}.${gy}`;
    } else if (tip === 'sedmicni') {
      const now = new Date(danas);
      const day = now.getDay() || 7;
      const mon = new Date(now); mon.setDate(now.getDate() - day + 1);
      const sun = new Date(mon); sun.setDate(mon.getDate() + 6);
      od  = mon.toISOString().slice(0,10);
      doo = sun.toISOString().slice(0,10);
      periodNaziv = `Sedmični — ${mon.getDate()}.${mon.getMonth()+1}. – ${sun.getDate()}.${sun.getMonth()+1}.${gy}`;
    } else if (tip === 'mjesecni') {
      od  = `${gy}-${String(gm).padStart(2,'0')}-01`;
      doo = danas;
      periodNaziv = `Mjesečni — ${String(gm).padStart(2,'0')}/${gy}`;
    } else if (tip === 'godisnji') {
      od  = `${gy}-01-01`;
      doo = danas;
      periodNaziv = `Godišnji — ${gy}`;
    } else if (tip === 'privremeni') {
      od  = datumOd || danas;
      doo = datumDo || danas;
      const [oy,om,odd] = od.split('-');
      const [dy,dm,ddd] = doo.split('-');
      periodNaziv = `Privremeni — ${odd}.${om}.${oy} – ${ddd}.${dm}.${dy}`;
    } else {
      return res.status(400).json({ error: 'Nevažeći tip izvještaja.' });
    }

    rows = rows.filter(r => String(r.datum) >= od && String(r.datum) <= doo);

    // Statistike
    const ukupnoEUR    = parseFloat(rows.reduce((s,r) => s + (r.eur||0), 0).toFixed(2));
    const ukupnoBAM    = parseFloat((ukupnoEUR * EUR_BAM).toFixed(2));
    const placenoEUR   = parseFloat(rows.reduce((s,r) => s + (r.placenoEUR||0), 0).toFixed(2));
    const ostatakEUR   = parseFloat(rows.reduce((s,r) => s + (r.ostatakEUR||0), 0).toFixed(2));
    const placeno      = rows.filter(r => r.status?.includes('✅')).length;
    const neplaceno    = rows.filter(r => r.status === 'Nije plaćeno').length;
    const djelimicno   = rows.filter(r => r.status?.includes('Djelimično')).length;

    // Po državama
    const poDrzavi = {};
    rows.forEach(r => {
      const d = (r.drzava||'?').trim();
      poDrzavi[d] = (poDrzavi[d]||0) + 1;
    });

    // Po vozilima
    const poVozilu = {};
    rows.forEach(r => {
      const v = (r.vozilo||'?').trim().replace(/\s*\(.*?\)/,'').trim();
      poVozilu[v] = (poVozilu[v]||0) + 1;
    });

    res.json({
      rows,
      periodNaziv,
      od, doo,
      ukupnoEUR, ukupnoBAM, placenoEUR, ostatakEUR,
      brojTurista: rows.length,
      placanje: { placeno, neplaceno, djelimicno },
      poDrzavi: Object.entries(poDrzavi).sort((a,b)=>b[1]-a[1]).map(([d,n])=>({drzava:d,broj:n})),
      poVozilu: Object.entries(poVozilu).sort((a,b)=>b[1]-a[1]).map(([v,n])=>({vozilo:v,broj:n})),
    });
  } catch(err) {
    console.error('❌ Izvještaj greška:', err);
    res.status(500).json({ error: 'Greška.' });
  }
});

// ─── GET /api/statistika ──────────────────────────────────────────
app.get('/api/statistika', async (req, res) => {
  if (!fs.existsSync(EXCEL_PUT)) return res.json({ prazno: true });
  try {
    const { worksheet } = await ucitajWorkbook();
    const rows = citajSveRedove(worksheet);
    const { datumOd, datumDo } = req.query;

    let filtrirani = rows;
    if (datumOd) filtrirani = filtrirani.filter(r => String(r.datum) >= datumOd);
    if (datumDo) filtrirani = filtrirani.filter(r => String(r.datum) <= datumDo);

    // ── Po državi ────────────────────────────────────────
    const poDrzavi = {};
    filtrirani.forEach(r => {
      const d = String(r.drzava || 'Nepoznato').trim();
      if (!poDrzavi[d]) poDrzavi[d] = { turista: 0, prihod: 0 };
      poDrzavi[d].turista++;
      poDrzavi[d].prihod += r.eur || 0;
    });

    // ── Po vozilu ────────────────────────────────────────
    const poVozilu = {};
    filtrirani.forEach(r => {
      const v = String(r.vozilo || 'Ostalo').trim();
      // Normalizuj nazive grupe (npr. "Moto Klub (10× Motor)" → "Motor")
      const normV = v.includes('×') ? v.replace(/.*×\s*/,'').replace(/\)/,'').trim() : v;
      if (!poVozilu[normV]) poVozilu[normV] = { broj: 0, prihod: 0 };
      poVozilu[normV].broj++;
      poVozilu[normV].prihod += r.eur || 0;
    });

    // ── Po danu (trend) ──────────────────────────────────
    const poDanu = {};
    filtrirani.forEach(r => {
      const d = String(r.datum || '');
      if (!d) return;
      if (!poDanu[d]) poDanu[d] = { turista: 0, prihod: 0 };
      poDanu[d].turista++;
      poDanu[d].prihod += r.eur || 0;
    });

    // ── Placanje status ──────────────────────────────────
    let placeno = 0, neplaceno = 0, djelimicno = 0;
    filtrirani.forEach(r => {
      if (r.status?.includes('✅')) placeno++;
      else if (r.status?.includes('Djelimično')) djelimicno++;
      else neplaceno++;
    });

    // ── Top statistike ───────────────────────────────────
    const ukupnoTurista   = filtrirani.length;
    const ukupnoEUR       = parseFloat(filtrirani.reduce((s,r) => s + (r.eur||0), 0).toFixed(2));
    const ukupnoPlaceno   = parseFloat(filtrirani.reduce((s,r) => s + (r.placenoEUR||0), 0).toFixed(2));
    const ukupnoOstatak   = parseFloat(filtrirani.reduce((s,r) => s + (r.ostatakEUR||0), 0).toFixed(2));
    const prosjecniBoravak = ukupnoTurista > 0
      ? parseFloat((filtrirani.reduce((s,r) => s + (parseFloat(r.dani)||1), 0) / ukupnoTurista).toFixed(1))
      : 0;

    res.json({
      ukupnoTurista, ukupnoEUR, ukupnoPlaceno, ukupnoOstatak,
      prosjecniBoravak,
      placanje: { placeno, neplaceno, djelimicno },
      poDrzavi: Object.entries(poDrzavi)
        .sort((a,b) => b[1].turista - a[1].turista)
        .map(([drzava, v]) => ({ drzava, ...v, prihod: parseFloat(v.prihod.toFixed(2)) })),
      poVozilu: Object.entries(poVozilu)
        .sort((a,b) => b[1].broj - a[1].broj)
        .map(([vozilo, v]) => ({ vozilo, ...v, prihod: parseFloat(v.prihod.toFixed(2)) })),
      poDanu: Object.entries(poDanu)
        .sort((a,b) => a[0].localeCompare(b[0]))
        .map(([datum, v]) => ({ datum, ...v, prihod: parseFloat(v.prihod.toFixed(2)) })),
    });
  } catch(err) {
    console.error('❌ Statistika greška:', err);
    res.status(500).json({ error: 'Greška.' });
  }
});

// ─── POST /api/grupa ──────────────────────────────────────────────
app.post('/api/grupa', async (req, res) => {
  const { naziv, drzava, dani, pas, sator, komentar, placeno, placenoNap, tip, vozila, osobe, djeca } = req.body;

  if (!naziv || naziv.length < 2)          return res.status(400).json({ error: 'Naziv grupe je obavezan.' });
  if (!drzava || drzava.length < 2)        return res.status(400).json({ error: 'Država je obavezna.' });
  if (!Array.isArray(vozila) || vozila.length === 0) return res.status(400).json({ error: 'Dodajte barem jedno vozilo.' });
  if (!Number.isInteger(dani) || dani < 1) return res.status(400).json({ error: 'Broj dana mora biti >= 1.' });

  try {
    const { workbook, worksheet } = await ucitajWorkbook();
    const datum   = danasniDatum();
    const grupaId = genId();
    const ts      = timestamp();

    const brojOsoba  = parseInt(osobe)  || 0;
    const brojDjeca  = parseInt(djeca)  || 0;
    const brojPasa   = parseInt(pas)    || 0;
    const brojSatora = parseInt(sator)  || 0;
    const pasEUR     = brojPasa   * CIJENA.pas   * dani;
    const satorEUR   = brojSatora * CIJENA.sator * dani;
    const osobeEUR   = brojOsoba * CIJENA.osoba * dani;
    const djecaEUR   = brojDjeca * CIJENA.dijete * dani;

    let ukupnoVozila = 0;
    let ukupnoKom    = 0;
    vozila.forEach(v => {
      ukupnoVozila += (VOZILO_CIJENA[v.tip] ?? 0) * (v.kol || 1) * dani;
      ukupnoKom    += (v.kol || 1);
    });

    const ukupnoTotal = parseFloat((ukupnoVozila + osobeEUR + djecaEUR + pasEUR + satorEUR).toFixed(2));
    const placenoEUR  = parseFloat(Math.min(placeno || 0, ukupnoTotal).toFixed(2));
    const ostatakEUR  = parseFloat(Math.max(0, ukupnoTotal - placenoEUR).toFixed(2));
    const status      = izracunajStatus(placenoEUR, ukupnoTotal);

    const uplateStr = placenoEUR > 0
      ? `1. ${placenoEUR.toFixed(2)}€${placenoNap ? ' — ' + placenoNap : ''} (${ts})`
      : '—';

    const komentarFull = `GRUPA: ${naziv} (${tip || 'ostalo'}) | ID: ${grupaId}${brojOsoba?` | ${brojOsoba} odraslih`:''}${brojDjeca?` | ${brojDjeca} djece`:''}${brojPasa?` | ${brojPasa}🐕`:''}${brojSatora?` | ${brojSatora}⛺`:''}${komentar ? ' | ' + komentar : ''}`;

    // Jedan red po VRSTI vozila
    for (const v of vozila) {
      const id        = genId();
      const kol       = v.kol || 1;
      const cijVozilo = (VOZILO_CIJENA[v.tip] ?? 0);
      const redVozEUR = parseFloat((cijVozilo * kol * dani).toFixed(2));
      // Proporcionalno rasporedi osobeEUR+djecaEUR+opcije
      const udio      = ukupnoTotal > 0 ? redVozEUR / ukupnoVozila : 0;
      const redUkupno = parseFloat((redVozEUR + (osobeEUR + djecaEUR + pasEUR + satorEUR) * udio).toFixed(2));
      const redPlacer  = parseFloat((placenoEUR * redUkupno / ukupnoTotal).toFixed(2));
      const redOstatak = parseFloat(Math.max(0, redUkupno - redPlacer).toFixed(2));
      const redStatus  = izracunajStatus(redPlacer, redUkupno);

      const noviRed = worksheet.addRow({
        id, datum,
        ime:        `${naziv} (${kol}× ${VOZILO_NAZIV[v.tip] || v.tip})`,
        drzava,
        tablice:    '—',
        vozilo:     VOZILO_NAZIV[v.tip] || v.tip,
        osobe:      brojOsoba, djeca: brojDjeca,
        pas:        brojPasa   > 0 ? `${brojPasa} 🐕`   : 'Ne',
        sator:      brojSatora > 0 ? `${brojSatora} ⛺` : 'Ne',
        dani,
        eur:        redUkupno,
        bam:        parseFloat((redUkupno * EUR_BAM).toFixed(2)),
        placenoEUR: redPlacer,
        placenoBAM: parseFloat((redPlacer * EUR_BAM).toFixed(2)),
        ostatakEUR: redOstatak,
        ostatakBAM: parseFloat((redOstatak * EUR_BAM).toFixed(2)),
        status:     redStatus,
        uplate:     uplateStr,
        komentar:   komentarFull,
        timestamp:  ts,
      });
      stilujRed(noviRed, redStatus);
    }

    await workbook.xlsx.writeFile(EXCEL_PUT);

    const racunHTML = generirajRacunGrupe({
      grupaId, naziv, drzava, tip, dani,
      pas: brojPasa, sator: brojSatora, komentar,
      vozila, osobe: brojOsoba, djeca: brojDjeca,
      ukupnoTotal, placenoEUR, ostatakEUR, status,
      uplateStr, datum, ts, ukupnoKom,
    });
    const racunPut = path.join(RACUNI_DIR, `GRUPA-${grupaId}.html`);
    fs.writeFileSync(racunPut, racunHTML, 'utf8');
    console.log(`✅ Grupa [${grupaId}] "${naziv}" | ${ukupnoKom} vozila | ${ukupnoTotal.toFixed(2)}€`);

    res.status(200).json({
      success: true, grupaId,
      ukupnoKom, ukupnoEUR: ukupnoTotal,
      racunUrl: `/racun/GRUPA-${grupaId}`,
    });

  } catch (err) {
    console.error('❌ Greška grupa:', err);
    res.status(500).json({ error: 'Greška pri čuvanju grupe.' });
  }
});

// ─── Generator skupnog računa za grupu ────────────────────────────
function generirajRacunGrupe({ grupaId, naziv, drzava, tip, dani, pas, sator, komentar, vozila, osobe, djeca, ukupnoTotal, placenoEUR, ostatakEUR, status, uplateStr, datum, ts, ukupnoKom }) {
  const bamTotal   = parseFloat((ukupnoTotal * EUR_BAM).toFixed(2));
  const placenoBAM = parseFloat((placenoEUR  * EUR_BAM).toFixed(2));
  const ostatakBAM = parseFloat((ostatakEUR  * EUR_BAM).toFixed(2));

  const statusBoja   = status?.includes('✅') ? '#3ecf8e' : status?.includes('Djelim') ? '#f5a623' : '#e8534a';
  const statusBg     = status?.includes('✅') ? 'rgba(62,207,142,0.12)' : status?.includes('Djelim') ? 'rgba(245,166,35,0.12)' : 'rgba(232,83,74,0.12)';
  const statusBorder = status?.includes('✅') ? 'rgba(62,207,142,0.3)'  : status?.includes('Djelim') ? 'rgba(245,166,35,0.3)'  : 'rgba(232,83,74,0.3)';

  const tipIkona = { motociklisti:'🏍️', izvidjaci:'⛺', skola:'🎒', sportasi:'🏅', turisti:'🌍', ostalo:'👥' };

  const vozilaHTML = vozila.map((v, i) => {
    const naziv_v = VOZILO_NAZIV[v.tip] || v.tip;
    const kol     = v.kol || 1;
    const cijDan  = VOZILO_CIJENA[v.tip] ?? 0;
    const iznos   = cijDan * kol * dani;
    return `<tr class="${i%2===0?'row-even':'row-odd'}">
      <td class="td-naziv">${naziv_v} × ${kol} kom.</td>
      <td class="td-center">${dani} dan</td>
      <td class="td-center">${cijDan.toFixed(2)} €/dan</td>
      <td class="td-right td-iznos">${iznos.toFixed(2)} €</td>
    </tr>`;
  }).join('');

  const opcijeHTML = [];
  if (osobe > 0) opcijeHTML.push(`<tr class="row-odd"><td class="td-naziv">Odrasli 👤 (${osobe} os.)</td><td class="td-center">${dani} dan</td><td class="td-center">${CIJENA.osoba.toFixed(2)} €/dan</td><td class="td-right td-iznos">${(CIJENA.osoba*osobe*dani).toFixed(2)} €</td></tr>`);
  if (djeca > 0) opcijeHTML.push(`<tr class="row-even"><td class="td-naziv">Djeca 👦 (${djeca})</td><td class="td-center">${dani} dan</td><td class="td-center">${CIJENA.dijete.toFixed(2)} €/dan</td><td class="td-right td-iznos">${(CIJENA.dijete*djeca*dani).toFixed(2)} €</td></tr>`);
  if (pas > 0)   opcijeHTML.push(`<tr class="row-odd"><td class="td-naziv">Psi 🐕 (${pas}×)</td><td class="td-center">${dani} dan</td><td class="td-center">${CIJENA.pas.toFixed(2)} €/dan</td><td class="td-right td-iznos">${(CIJENA.pas*pas*dani).toFixed(2)} €</td></tr>`);
  if (sator > 0) opcijeHTML.push(`<tr class="row-even"><td class="td-naziv">Šatori ⛺ (${sator}×)</td><td class="td-center">${dani} dan</td><td class="td-center">${CIJENA.sator.toFixed(2)} €/dan</td><td class="td-right td-iznos">${(CIJENA.sator*sator*dani).toFixed(2)} €</td></tr>`);

  const uplateHTML = (uplateStr && uplateStr !== '—')
    ? uplateStr.split('\n').map(u => `<div class="uplata-red">✓ ${u}</div>`).join('')
    : '<div class="uplata-red muted">—</div>';

  const sada = new Date().toLocaleString('bs-BA', { day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit' });

  return `<!DOCTYPE html>
<html lang="bs">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Grupni račun — ${naziv}</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet"/>
<style>
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
  :root{--bg:#0b0c0e;--surface:#13151a;--card:#1a1d24;--border:#252830;--accent:#f5a623;--green:#3ecf8e;--blue:#5b9cf6;--red:#e8534a;--text:#eceef2;--muted:#6b7280;--sub:#9ca3af;}
  body{font-family:'DM Sans','Segoe UI',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;padding-bottom:80px;}
  .wrap{max-width:620px;margin:0 auto;padding:40px 32px 48px;}
  .header{text-align:center;padding-bottom:28px;margin-bottom:28px;border-bottom:1px solid var(--border);position:relative;}
  .header::after{content:'';position:absolute;bottom:-1px;left:50%;transform:translateX(-50%);width:60px;height:2px;background:var(--accent);}
  .logo-icon{font-size:36px;margin-bottom:8px;display:block;}
  .header h1{font-family:'Syne',sans-serif;font-size:26px;font-weight:800;letter-spacing:-0.5px;color:var(--text);margin-bottom:4px;}
  .header .subtitle{font-size:12px;color:var(--muted);letter-spacing:0.5px;text-transform:uppercase;margin-bottom:12px;}
  .racun-id-badge{display:inline-flex;align-items:center;gap:8px;background:var(--surface);border:1px solid var(--border);border-radius:100px;padding:5px 14px;font-size:11px;color:var(--muted);font-family:monospace;}
  .racun-id-badge strong{color:var(--accent);font-size:12px;}
  .grupa-badge{display:inline-flex;align-items:center;gap:6px;background:rgba(91,156,246,0.1);border:1px solid rgba(91,156,246,0.25);color:var(--blue);border-radius:100px;padding:6px 16px;font-family:'Syne',sans-serif;font-size:12px;font-weight:700;margin:12px 0;}
  .section-title{font-family:'Syne',sans-serif;font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--accent);margin-bottom:12px;display:flex;align-items:center;gap:10px;}
  .section-title::after{content:'';flex:1;height:1px;background:var(--border);}
  .gost-card{background:var(--card);border:1px solid var(--border);border-radius:16px;padding:22px 24px;margin-bottom:20px;position:relative;overflow:hidden;}
  .gost-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--accent),transparent);}
  .gost-name{font-family:'Syne',sans-serif;font-size:22px;font-weight:800;color:var(--text);margin-bottom:6px;}
  .info-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px 24px;}
  .info-item .lbl{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:3px;display:block;}
  .info-item .val{font-size:14px;font-weight:600;color:var(--text);}
  .komentar-box{margin-top:14px;padding:10px 14px;background:var(--surface);border-radius:8px;border-left:2px solid var(--border);font-size:12px;color:var(--sub);font-style:italic;}
  .stavke-card{background:var(--card);border:1px solid var(--border);border-radius:16px;overflow:hidden;margin-bottom:20px;}
  table{width:100%;border-collapse:collapse;}
  thead tr{background:var(--surface);border-bottom:1px solid var(--border);}
  th{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:var(--muted);padding:12px 16px;text-align:left;}
  .row-even{background:var(--card);}.row-odd{background:rgba(255,255,255,0.015);}
  td{padding:12px 16px;border-bottom:1px solid var(--border);}
  tr:last-child td{border-bottom:none;}
  .td-naziv{font-size:14px;font-weight:500;color:var(--text);}
  .td-center{text-align:center;font-size:13px;color:var(--sub);}
  .td-right{text-align:right;font-size:13px;color:var(--sub);}
  .td-iznos{font-weight:700;font-size:14px;color:var(--text);}
  .totals-card{background:var(--card);border:1px solid var(--border);border-radius:16px;overflow:hidden;margin-bottom:20px;}
  .total-row{display:flex;justify-content:space-between;align-items:center;padding:14px 20px;border-bottom:1px solid var(--border);}
  .total-row:last-child{border-bottom:none;}
  .total-row .lbl{font-size:13px;color:var(--sub);font-weight:500;}
  .total-row .val{font-size:14px;font-weight:700;color:var(--text);}
  .total-row.big{background:var(--surface);padding:18px 20px;}
  .total-row.big .lbl{font-family:'Syne',sans-serif;font-size:14px;font-weight:800;color:var(--text);}
  .total-row.big .val{font-family:'Syne',sans-serif;font-size:24px;font-weight:800;color:var(--accent);}
  .total-row.bam .val{color:var(--blue);font-size:13px;}
  .total-row.placeno .val,.total-row.placeno .lbl{color:var(--green);}
  .total-row.ostatak .val,.total-row.ostatak .lbl{color:var(--red);}
  .total-row.ostatak-0 .val,.total-row.ostatak-0 .lbl{color:var(--green);}
  .status-wrap{text-align:center;margin:6px 0 20px;}
  .status-badge{display:inline-block;padding:8px 22px;border-radius:100px;font-family:'Syne',sans-serif;font-size:13px;font-weight:800;background:${statusBg};color:${statusBoja};border:1px solid ${statusBorder};}
  .uplate-card{background:var(--card);border:1px solid var(--border);border-radius:16px;padding:18px 20px;margin-bottom:20px;}
  .uplata-red{font-size:13px;color:var(--sub);padding:5px 0;border-bottom:1px solid var(--border);line-height:1.5;}
  .uplata-red:last-child{border-bottom:none;}
  .muted{color:var(--muted);}
  .footer{text-align:center;padding:24px 0 0;border-top:1px solid var(--border);}
  .footer .hvala{font-family:'Syne',sans-serif;font-size:18px;font-weight:800;color:var(--text);margin-bottom:10px;}
  .footer .meta{font-size:12px;color:var(--muted);line-height:1.8;}
  .footer .stampano{margin-top:10px;font-size:11px;color:var(--border);font-family:monospace;}
  .print-bar{position:fixed;bottom:0;left:0;right:0;background:var(--surface);border-top:1px solid var(--border);padding:14px 24px;display:flex;gap:12px;justify-content:center;z-index:100;}
  .btn-print{font-family:'Syne',sans-serif;background:var(--accent);color:#000;border:none;font-weight:800;font-size:14px;padding:12px 32px;border-radius:10px;cursor:pointer;letter-spacing:0.3px;}
  .btn-print:hover{background:#fbb740;}
  .btn-close{font-family:'Syne',sans-serif;background:var(--card);color:var(--muted);border:1px solid var(--border);font-weight:700;font-size:13px;padding:12px 24px;border-radius:10px;cursor:pointer;}
  .btn-close:hover{border-color:var(--red);color:var(--red);}
  @media print{
    .print-bar{display:none!important;}
    body{background:#fff!important;color:#000!important;padding-bottom:0;}
    .wrap{padding:10px 16px;}
    .gost-card,.stavke-card,.totals-card,.uplate-card{background:#fff!important;border:1px solid #ccc!important;border-radius:8px!important;}
    .gost-card::before{display:none;}
    .header::after{background:#000;}
    .gost-name,.header h1,.section-title{color:#000!important;}
    .td-naziv,.td-iznos,.total-row .lbl,.total-row .val{color:#000!important;}
    .td-center,.td-right,.uplata-red,.info-item .val{color:#333!important;}
    .total-row.big{background:#f5f5f5!important;}
    .total-row.big .val{color:#c47d00!important;}
    .total-row.placeno .val,.total-row.ostatak-0 .val{color:#166534!important;}
    .total-row.ostatak .val{color:#b91c1c!important;}
    thead tr{background:#f5f5f5!important;}
    .status-badge{background:#f5f5f5!important;color:#333!important;border-color:#ccc!important;}
    @page{margin:10mm;size:A4;}
  }
</style>
</head>
<body>
<div class="wrap">
  <div class="header">
    <span class="logo-icon">🏕️</span>
    <h1>KAMPING</h1>
    <div class="subtitle">Grupna potvrda o boravku / Group Receipt</div>
    <div class="racun-id-badge">Račun: <strong>GRUPA-${grupaId}</strong> &nbsp;·&nbsp; ${datum}</div>
  </div>

  <div class="section-title">Podaci grupe</div>
  <div class="gost-card">
    <div class="gost-name">${naziv}</div>
    <div style="margin-bottom:14px;">
      <span class="grupa-badge">${tipIkona[tip]||'👥'} ${tip.charAt(0).toUpperCase()+tip.slice(1)}</span>
    </div>
    <div class="info-grid">
      <div class="info-item"><span class="lbl">Država</span><span class="val">${drzava}</span></div>
      <div class="info-item"><span class="lbl">Ukupno vozila</span><span class="val">${ukupnoKom} kom.</span></div>
      <div class="info-item"><span class="lbl">Broj dana</span><span class="val">${dani} dan(a)</span></div>
      <div class="info-item"><span class="lbl">Datum dolaska</span><span class="val">${datum}</span></div>
    </div>
    ${komentar ? `<div class="komentar-box">📝 ${komentar}</div>` : ''}
  </div>

  <div class="section-title">Članovi grupe</div>
  <div class="stavke-card">
    <table>
      <thead><tr><th>Vozilo</th><th style="text-align:center">Dana</th><th style="text-align:center">Cijena/dan</th><th style="text-align:right">Iznos</th></tr></thead>
      <tbody>${vozilaHTML}${opcijeHTML.join('')}</tbody>
    </table>
  </div>

  <div class="section-title">Obračun</div>
  <div class="totals-card">
    <div class="total-row big"><span class="lbl">UKUPNO GRUPA</span><span class="val">${ukupnoTotal.toFixed(2)} €</span></div>
    <div class="total-row bam"><span class="lbl">Iznos u KM</span><span class="val">${bamTotal.toFixed(2)} KM</span></div>
    <div class="total-row placeno"><span class="lbl">Plaćeno</span><span class="val">+ ${placenoEUR.toFixed(2)} € <span style="font-size:12px;font-weight:400">(${placenoBAM.toFixed(2)} KM)</span></span></div>
    <div class="total-row ${ostatakEUR < 0.01 ? 'ostatak-0' : 'ostatak'}"><span class="lbl">Ostatak</span><span class="val">${ostatakEUR < 0.01 ? '0.00 € ✅' : ostatakEUR.toFixed(2) + ' € <span style="font-size:12px;font-weight:400">(' + ostatakBAM.toFixed(2) + ' KM)</span>'}</span></div>
  </div>
  <div class="status-wrap"><span class="status-badge">${status}</span></div>

  <div class="section-title">Historija plaćanja</div>
  <div class="uplate-card">${uplateHTML}</div>

  <div class="footer">
    <div class="hvala">Hvala na posjeti! 🙏</div>
    <div class="meta">Kurs: 1 € = ${EUR_BAM} KM (fiksni kurs)<br/>Grupni račun: GRUPA-${grupaId}</div>
    <div class="stampano">Štampano: ${sada}</div>
  </div>
</div>
<div class="print-bar">
  <button class="btn-print" onclick="window.print()">🖨️ Štampaj / Print</button>
  <button class="btn-close" onclick="window.close()">✕ Zatvori</button>
</div>
</body>
</html>`;
}

app.listen(PORT, () => {
  console.log(`\n╔══════════════════════════════════════════════════════╗`);
  console.log(`║  🚗 Kalkulator aktivan → http://localhost:${PORT}        ║`);
  console.log(`║  📁 data/kalkulator.xlsx  (jedan fajl za sve)        ║`);
  console.log(`╚══════════════════════════════════════════════════════╝\n`);
});
