import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import logoShell from './logo-coquillage.png';
import './style.css';

const LYRECO = {
  name: 'Lyreco 70 × 37 mm',
  pageW: 210,
  pageH: 297,
  width: 70,
  height: 37,
  cols: 3,
  rows: 8,
  marginLeft: 0,
  marginTop: 0.5,
  gapX: 0,
  gapY: 0,
};

const REQUIRED_COLUMNS = ['Arrival Date', 'Departure Date', 'Accommodation Type', 'Customer First Name', 'Customer Last Name', 'Unit Name', 'Reservation Number'];
const normalize = (value) => String(value ?? '').trim();

const aliases = {
  arrival: ['Arrival Date', 'Date arrivée', 'Date d’arrivée', 'Arrivée', 'Check-in'],
  departure: ['Departure Date', 'Date départ', 'Date de départ', 'Départ', 'Check-out'],
  type: ['Accommodation Type', 'Type hébergement', 'Type d’hébergement', 'Hébergement', 'Accommodation'],
  firstName: ['Customer First Name', 'Prénom', 'Prenom', 'First Name'],
  lastName: ['Customer Last Name', 'Nom', 'Last Name', 'Customer LastName'],
  unit: ['Unit Name', 'Emplacement', 'Logement', 'Unit', 'Numéro emplacement'],
  reservation: ['Reservation Number', 'Réservation', 'Reservation', 'Booking', 'Numéro réservation'],
};

function getByAliases(row, list) {
  for (const key of list) if (Object.prototype.hasOwnProperty.call(row, key)) return row[key];
  return '';
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function formatDateParts(year, month, day) {
  if (!year || !month || !day) return '';
  return `${pad2(day)}/${pad2(month)}/${String(year).padStart(4, '0')}`;
}

function formatDate(value) {
  if (value === '' || value === null || value === undefined) return '';

  // Les dates Excel restent des nombres bruts. On récupère directement
  // leurs composantes au lieu de passer par un objet Date, ce qui évite
  // tout décalage d'un jour lié au fuseau horaire.
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    return parsed ? formatDateParts(parsed.y, parsed.m, parsed.d) : normalize(value);
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    // XLSX peut parfois fournir une date à minuit UTC : utiliser les champs UTC
    // garantit que le jour inscrit dans le tableau reste strictement identique.
    return formatDateParts(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate());
  }

  if (typeof value === 'string') {
    const clean = value.replace(/[!'’]/g, '').replace(/\s+/g, '').trim();
    const fr = clean.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/);
    if (fr) {
      const [, day, month, year] = fr;
      return formatDateParts(year.length === 2 ? Number(`20${year}`) : Number(year), Number(month), Number(day));
    }
    const iso = clean.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:T.*)?$/);
    if (iso) return formatDateParts(Number(iso[1]), Number(iso[2]), Number(iso[3]));
  }

  return normalize(value);
}

function splitShells(type) {
  const clean = normalize(type);
  const matches = clean.match(/\*+/g) || [];
  const shellCount = matches.reduce((total, stars) => total + stars.length, 0);
  return {
    label: clean.replace(/\s*\*+/g, '').replace(/\s+/g, ' ').trim(),
    shellCount: Math.min(shellCount, 5),
  };
}

function readReservations(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Lecture du fichier impossible.'));
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: false });
        const sheet = workbook.Sheets.Reservations || workbook.Sheets[workbook.SheetNames[0]];
        if (!sheet) throw new Error('Aucune feuille trouvée.');
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });
        const headers = rows[0] ? Object.keys(rows[0]) : [];
        const reservations = rows.map((row, index) => ({
          id: `${index}-${normalize(getByAliases(row, aliases.reservation))}`,
          arrival: getByAliases(row, aliases.arrival),
          departure: getByAliases(row, aliases.departure),
          type: getByAliases(row, aliases.type),
          firstName: getByAliases(row, aliases.firstName),
          lastName: getByAliases(row, aliases.lastName),
          unit: getByAliases(row, aliases.unit),
          reservation: getByAliases(row, aliases.reservation),
        })).filter((item) => normalize(item.firstName) || normalize(item.lastName) || normalize(item.unit));
        resolve({ reservations, headers, sheetName: sheet === workbook.Sheets.Reservations ? 'Reservations' : workbook.SheetNames[0] });
      } catch (error) {
        reject(error);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function loadDataUrl(url) {
  return fetch(url).then((response) => response.blob()).then((blob) => new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.readAsDataURL(blob);
  }));
}

function setFont(doc, size, style = 'normal') {
  doc.setFont('helvetica', style);
  doc.setFontSize(size);
}

function fitText(doc, text, x, y, maxWidth, size, style = 'normal', minSize = 5) {
  const clean = normalize(text);
  if (!clean) return '';
  let fontSize = size;
  setFont(doc, fontSize, style);
  while (fontSize > minSize && doc.getTextWidth(clean) > maxWidth) {
    fontSize -= 0.25;
    setFont(doc, fontSize, style);
  }
  if (doc.getTextWidth(clean) <= maxWidth) {
    doc.text(clean, x, y);
    return clean;
  }
  let shortened = clean;
  while (shortened.length > 2 && doc.getTextWidth(`${shortened}…`) > maxWidth) shortened = shortened.slice(0, -1);
  const result = `${shortened}…`;
  doc.text(result, x, y);
  return result;
}

function fitTextTwoLines(doc, text, x, y, maxWidth, size, style = 'normal', lineGap = 3.4) {
  const clean = normalize(text);
  if (!clean) return 0;
  setFont(doc, size, style);
  if (doc.getTextWidth(clean) <= maxWidth) {
    doc.text(clean, x, y);
    return 1;
  }
  const words = clean.split(/\s+/);
  let line1 = '';
  let line2 = '';
  for (const word of words) {
    const candidate = line1 ? `${line1} ${word}` : word;
    if (!line2 && doc.getTextWidth(candidate) <= maxWidth) line1 = candidate;
    else line2 = line2 ? `${line2} ${word}` : word;
  }
  if (!line1) {
    fitText(doc, clean, x, y, maxWidth, size, style, 5.2);
    return 1;
  }
  doc.text(line1, x, y);
  if (line2) fitText(doc, line2, x, y + lineGap, maxWidth, size, style, 5.0);
  return line2 ? 2 : 1;
}

function drawShell(doc, cx, cy, size = 1.65) {
  doc.setDrawColor(184, 139, 82);
  doc.setLineWidth(0.18);
  doc.ellipse(cx, cy, size * 0.72, size * 0.54);
  doc.line(cx, cy + size * 0.5, cx, cy - size * 0.52);
  doc.line(cx, cy + size * 0.45, cx - size * 0.43, cy - size * 0.32);
  doc.line(cx, cy + size * 0.45, cx + size * 0.43, cy - size * 0.32);
  doc.line(cx, cy + size * 0.45, cx - size * 0.68, cy);
  doc.line(cx, cy + size * 0.45, cx + size * 0.68, cy);
}

function drawIcon(doc, type, x, y) {
  doc.setDrawColor(6, 54, 81);
  doc.setFillColor(6, 54, 81);
  doc.setLineWidth(0.22);
  if (type === 'calendar') {
    doc.roundedRect(x, y - 2.4, 2.7, 2.7, 0.25, 0.25);
    doc.line(x, y - 1.55, x + 2.7, y - 1.55);
    doc.line(x + 0.62, y - 2.82, x + 0.62, y - 1.9);
    doc.line(x + 2.05, y - 2.82, x + 2.05, y - 1.9);
  }
  if (type === 'home') {
    doc.triangle(x, y - 1.0, x + 1.45, y - 2.35, x + 2.9, y - 1.0, 'S');
    doc.rect(x + 0.5, y - 1.0, 1.9, 1.72);
  }
}

async function buildPdf(reservations, { guides = true, fileLabel = 'arrivees' } = {}) {
  const logoDataUrl = await loadDataUrl(logoShell);
  const doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait', compress: true });
  const perPage = LYRECO.cols * LYRECO.rows;
  const totalPages = Math.max(1, Math.ceil(reservations.length / perPage));

  reservations.forEach((reservation, index) => {
    if (index > 0 && index % perPage === 0) doc.addPage('a4', 'portrait');
    const slot = index % perPage;
    const col = slot % LYRECO.cols;
    const row = Math.floor(slot / LYRECO.cols);
    const x = LYRECO.marginLeft + col * (LYRECO.width + LYRECO.gapX);
    const y = LYRECO.marginTop + row * (LYRECO.height + LYRECO.gapY);

    if (guides) {
      doc.setDrawColor(214, 214, 214);
      doc.setLineWidth(0.08);
      doc.rect(x, y, LYRECO.width, LYRECO.height);
    }

    const left = x + 4.7;
    // Zone de sécurité interne : le gabarit reste strictement en 70 × 37 mm,
    // mais le contenu reste éloigné des zones non imprimables en haut et en bas.
    const top = y + 2.7;
    const right = x + LYRECO.width - 4.7;
    const width = right - left;
    const fullName = `${normalize(reservation.lastName).toUpperCase()} ${normalize(reservation.firstName)}`.trim();
    const unit = normalize(reservation.unit) || '—';
    const { label: type, shellCount } = splitShells(reservation.type);
    const arrival = formatDate(reservation.arrival);
    const departure = formatDate(reservation.departure);
    const ref = normalize(reservation.reservation);

    // Design V11 : rework complet, plus lisible et plus premium sur 70 x 37 mm.
    const navy = [150, 132, 86];
    const teal = [177, 158, 107];
    const gold = [177, 158, 107];
    const text = [49, 45, 37];
    const muted = [112, 101, 78];

    // Marque discrète + nom client
    doc.addImage(logoDataUrl, 'PNG', left + 0.1, top + 0.2, 7.2, 4.35, undefined, 'FAST');
    doc.setTextColor(...navy);
    fitText(doc, fullName, left + 7.8, top + 4.95, width - 8.2, 10.6, 'bold', 6.8);

    // Bloc logement pleine largeur : stable, aligné, sans décalage visuel.
    const bandY = top + 7.85;
    const bandH = 7.15;
    doc.setFillColor(...navy);
    doc.roundedRect(left + 0.2, bandY, width - 0.4, bandH, 1.15, 1.15, 'F');
    doc.setTextColor(255, 255, 255);
    setFont(doc, 5.9, 'bold');
    doc.text('LOGEMENT', left + 3.0, bandY + 4.65);
    doc.setDrawColor(255, 255, 255);
    doc.setLineWidth(0.15);
    const dividerX = left + 20.7;
    doc.line(dividerX, bandY + 1.7, dividerX, bandY + bandH - 1.7);
    // V12 : on ne touche pas au mot LOGEMENT, mais le numéro est centré
    // dans toute la zone bleue restante à droite du séparateur.
    setFont(doc, 11.9, 'bold');
    const valueZoneLeft = dividerX + 1.2;
    const valueZoneRight = left + width - 0.8;
    const valueCenterX = valueZoneLeft + (valueZoneRight - valueZoneLeft) / 2;
    const fittedUnit = normalize(unit);
    setFont(doc, 11.9, 'bold');
    let unitSize = 11.9;
    while (unitSize > 8.4 && doc.getTextWidth(fittedUnit) > valueZoneRight - valueZoneLeft - 1.0) {
      unitSize -= 0.25;
      setFont(doc, unitSize, 'bold');
    }
    doc.text(fittedUnit, valueCenterX, bandY + 5.2, { align: 'center' });

    // Dates : deux zones équilibrées, plus grandes, rapides à lire.
    const dateY = top + 19.15;
    const boxW = (width - 2.2) / 2;
    const boxH = 5.35;
    function drawDateBox(label, value, bx) {
      doc.setDrawColor(217, 226, 229);
      doc.setFillColor(248, 251, 251);
      doc.roundedRect(bx, dateY - 3.7, boxW, boxH, 0.75, 0.75, 'FD');
      drawIcon(doc, 'calendar', bx + 1.3, dateY - 0.2);
      doc.setTextColor(0, 0, 0);
      setFont(doc, 6.85, 'bold');
      doc.text(label, bx + 5.1, dateY - 0.05);
      fitText(doc, value, bx + 11.0, dateY - 0.05, boxW - 12.4, 7.35, 'normal', 5.65);
    }
    drawDateBox('Arr.', arrival, left + 0.2);
    drawDateBox('Dép.', departure, left + 0.2 + boxW + 2.2);

    // Hébergement + coquillages : une ligne principale, très lisible.
    const typeY = top + 25.95;
    doc.setTextColor(...text);
    drawIcon(doc, 'home', left + 0.4, typeY);
    const shellArea = shellCount ? 15.0 : 0;
    const lines = fitTextTwoLines(doc, type, left + 5.0, typeY, width - 5.5 - shellArea, 7.85, 'normal', 3.35);
    const shellY = lines > 1 ? top + 28.55 : top + 25.25;
    const shellSpacing = 2.95;
    const shellsWidth = Math.max(0, shellCount - 1) * shellSpacing;
    const startShellX = right - 2.8 - shellsWidth;
    for (let i = 0; i < shellCount; i += 1) drawShell(doc, startShellX + i * shellSpacing, shellY, 2.02);

    // Bas de carte : trait signature + référence discrète.
    doc.setDrawColor(...teal);
    doc.setLineWidth(0.55);
    doc.line(left + 0.2, top + 29.95, right - 0.2, top + 29.95);
    doc.setTextColor(...muted);
    fitText(doc, ref ? `Réservation ${ref}` : '', left + 0.2, top + 32.85, width - 0.4, 6.75, 'normal', 5.35);
  });

  doc.setProperties({
    title: `L’étiquettix 3000 · ${fileLabel}`,
    subject: `${reservations.length} étiquette(s), ${totalPages} page(s), Lyreco 70x37 mm`,
    creator: 'L’étiquettix 3000',
  });
  return doc;
}

async function openPdf(reservations, options = {}) {
  const doc = await buildPdf(reservations, options);
  const blob = doc.output('blob');
  const url = URL.createObjectURL(blob);
  window.open(url, '_blank', 'noopener,noreferrer');
  window.setTimeout(() => URL.revokeObjectURL(url), 90_000);
}

function Toast({ message, onClose }) {
  if (!message) return null;
  return <button className="toast" onClick={onClose}>{message}</button>;
}

function App() {
  const [reservations, setReservations] = useState([]);
  const [fileName, setFileName] = useState('');
  const [sheetName, setSheetName] = useState('');
  const [headers, setHeaders] = useState([]);
  const [error, setError] = useState('');
  const guides = true;
  const [autoOpen, setAutoOpen] = useState(true);
  const [isReading, setIsReading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [toast, setToast] = useState('');

  const stats = useMemo(() => {
    const pages = Math.ceil(reservations.length / 24);
    const units = new Set(reservations.map((r) => normalize(r.unit)).filter(Boolean));
    return { pages, units: units.size };
  }, [reservations]);

  async function handleFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    setIsReading(true);
    setError('');
    setReservations([]);
    setFileName(file.name);
    try {
      const result = await readReservations(file);
      setReservations(result.reservations);
      setHeaders(result.headers);
      setSheetName(result.sheetName);
      if (!result.reservations.length) {
        setError('Aucune arrivée détectée.');
        return;
      }
      if (autoOpen) {
        const cleanName = file.name.replace(/\.[^.]+$/, '').replace(/[^a-z0-9-_]+/gi, '-').toLowerCase();
        setIsGenerating(true);
        window.setTimeout(async () => {
          await openPdf(result.reservations, { guides: true, fileLabel: cleanName || 'arrivees' });
          setIsGenerating(false);
        }, 120);
      }
    } catch (err) {
      console.error(err);
      setError("Impossible de lire ce fichier Excel.");
    } finally {
      setIsReading(false);
      event.target.value = '';
    }
  }

  async function handleGenerate() {
    if (!reservations.length) return;
    setIsGenerating(true);
    try {
      await openPdf(reservations, { guides: true, fileLabel: 'arrivees-camping' });
    } finally {
      setIsGenerating(false);
    }
  }

  const missingExpectedColumns = REQUIRED_COLUMNS.filter((column) => !headers.includes(column));

  function triggerEasterEgg() {
    setToast("Pourquoi le Tix ? j'sais pas moi, demandez à Sandroutix.");
    window.setTimeout(() => setToast(''), 3500);
  }

  return (
    <main className="app">
      <div className="ambient ambientOne" />
      <div className="ambient ambientTwo" />
      <div className="grain" />

      <section className="shell">
        <header className="topbar">
          <button className="brand" onClick={triggerEasterEgg} title="Easter egg">
            <img src={logoShell} alt="" />
            <span>L’étiquettix 3000</span>
          </button>

        </header>

        <section className="workspace">
          <div className="commandCard">
            <div className="titleBlock">
              <h1>L’étiquettix 3000</h1>
              
            </div>

            <label className="dropZone">
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileChange} />
              <span className="dropIcon">＋</span>
              <strong>{isReading ? 'Lecture du fichier…' : isGenerating ? 'Préparation du PDF…' : 'Importer Excel'}</strong>
            </label>

            <div className="actions actionsClean">
              <button className="openButton" onClick={handleGenerate} disabled={!reservations.length || isGenerating}>
                Ouvrir PDF
              </button>
            </div>
          </div>

          <aside className="readout">
            <div><small>Fichier</small><strong>{fileName || '—'}</strong></div>
            <div><small>Feuille</small><strong>{sheetName || '—'}</strong></div>
            <div><small>Arrivées</small><strong>{reservations.length || '—'}</strong></div>
            <div><small>Pages</small><strong>{stats.pages || '—'}</strong></div>
            <div><small>Logements</small><strong>{stats.units || '—'}</strong></div>
          </aside>
        </section>
      </section>

      {error && <div className="notice error">{error}</div>}
      {missingExpectedColumns.length > 0 && headers.length > 0 && <div className="notice warning">Colonnes adaptées automatiquement.</div>}

      {reservations.length > 0 && (
        <section className="previewShell">
          <header>
            <h2>Aperçu</h2>
            <button onClick={handleGenerate} disabled={isGenerating}>PDF</button>
          </header>
          <div className="previewGrid">
            {reservations.slice(0, 12).map((reservation) => {
              const fullName = `${normalize(reservation.lastName).toUpperCase()} ${normalize(reservation.firstName)}`.trim();
              const { label, shellCount } = splitShells(reservation.type);
              return (
                <article className="previewLabel" key={reservation.id}>
                  <img src={logoShell} alt="" />
                  <div className="previewBody">
                    <strong>{fullName}</strong>
                    <b>LOGEMENT <span>{normalize(reservation.unit) || '—'}</span></b>
                    <p><span>Arr. {formatDate(reservation.arrival)}</span><span>Dép. {formatDate(reservation.departure)}</span></p>
                    <small>{label} <i>{'◚'.repeat(shellCount)}</i></small>
                  </div>
                </article>
              );
            })}
          </div>
        </section>
      )}

      <Toast message={toast} onClose={() => setToast('')} />
    </main>
  );
}

createRoot(document.getElementById('root')).render(<App />);
