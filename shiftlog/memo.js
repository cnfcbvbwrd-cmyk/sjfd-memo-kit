// SJFD Monthly CA Memo engine. Runs in the browser (globals docx, PDFLib, LETTERHEAD_B64)
// and in Node for tests. Every output is a pure function of the stored data.
var MemoLib = (function () {
  const _docx = typeof docx !== 'undefined' ? docx : require('docx');
  const _pdf = typeof PDFLib !== 'undefined' ? PDFLib : require('pdf-lib');

  // ---- Category table. Order here is the memo table order and the narrative tie-break order.
  const CATS = [
    { key: 'ems',        code: 'ems',   label: 'EMS Calls',            one: 'EMS call',            many: 'EMS calls' },
    { key: 'structure',  code: 'str',   label: 'Structure Fires',      one: 'structure fire',      many: 'structure fires' },
    { key: 'vehicle',    code: 'veh',   label: 'Vehicle Fires',        one: 'vehicle fire',        many: 'vehicle fires' },
    { key: 'mvc',        code: 'mvc',   label: 'MVCs',                 one: 'MVC',                 many: 'MVCs' },
    { key: 'rescue',     code: 'resc',  label: 'Rescue Responses',     one: 'rescue response',     many: 'rescue responses' },
    { key: 'alarm',      code: 'alarm', label: 'Fire Alarms',          one: 'fire alarm',          many: 'fire alarms' },
    { key: 'outside',    code: 'out',   label: 'Outside Fires',        one: 'outside fire',        many: 'outside fires' },
    { key: 'trash',      code: 'trash', label: 'Trash Fires',          one: 'trash fire',          many: 'trash fires' },
    { key: 'highrise',   code: 'hr',    label: 'High Rise Fires',      one: 'high rise fire',      many: 'high rise fires' },
    { key: 'vegetation', code: 'veg',   label: 'Vegetation Fires',     one: 'vegetation fire',     many: 'vegetation fires' },
    { key: 'gas',        code: 'gas',   label: 'Gas Investigations',   one: 'gas investigation',   many: 'gas investigations' },
    { key: 'smoke',      code: 'smoke', label: 'Smoke Investigations', one: 'smoke investigation', many: 'smoke investigations' },
    { key: 'public',     code: 'pub',   label: 'Public Service',       one: 'public service call', many: 'public service calls' },
    { key: 'lift',       code: 'lift',  label: 'Lift Assists',         one: 'lift assist',         many: 'lift assists' },
    { key: 'wires',      code: 'wires', label: 'Wires Down',           one: 'wires down incident', many: 'wires down incidents' },
    { key: 'fuel',       code: 'fuel',  label: 'Fuel Spills',          one: 'fuel spill',          many: 'fuel spills' },
  ];
  // Log entry order (matches the Drive doc code list)
  const CODE_ORDER = ['ems','mvc','str','hr','veg','veh','out','trash','gas','smoke','resc','pub','lift','wires','alarm','fuel'];
  const byCode = Object.fromEntries(CATS.map(c => [c.code, c]));
  const byKey = Object.fromEntries(CATS.map(c => [c.key, c]));
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

  const pad2 = n => String(n).padStart(2, '0');
  const fmtHr = h => (Number.isInteger(h) ? String(h) : String(h).replace(/\.?0+$/, ''));
  // 'YYYY-MM-DD' -> 'MM/DD'
  const mmdd = iso => iso.slice(5, 7) + '/' + iso.slice(8, 10);
  const mmddyyyy = iso => iso.slice(5, 7) + '/' + iso.slice(8, 10) + '/' + iso.slice(0, 4);
  const monthName = ym => MONTHS[+ym.slice(5, 7) - 1] + ' ' + ym.slice(0, 4);
  function listJoin(items) {
    if (items.length <= 1) return items.join('');
    if (items.length === 2) return items[0] + ' and ' + items[1];
    return items.slice(0, -1).join(', ') + ', and ' + items[items.length - 1];
  }
  function sentence(s) { s = s.trim(); return /[.!?]$/.test(s) ? s : s + '.'; }
  const sortByDate = arr => arr.slice().sort((a, b) => (a.date + (a.time || '')).localeCompare(b.date + (b.time || '')));

  // ---- Text metrics. Helvetica widths per 1000 units for ASCII 32..126; Arial is metrically
  // compatible, so these predict how Word will wrap the same string.
  const W_REG = [278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,334,260,334,584];
  const W_BOLD = [278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,333,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500,389,280,389,584];
  function textW(s, size, bold) {
    const t = bold ? W_BOLD : W_REG; let w = 0;
    for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i) - 32; w += (c >= 0 && c < t.length ? t[c] : 556); }
    return w * size / 1000;
  }
  // Number of lines `text` occupies when wrapped to `width` points.
  function lineCount(text, size, width, bold, indentFirst) {
    const words = String(text || '').split(/\s+/).filter(Boolean);
    if (!words.length) return 1;
    let lines = 1, cur = indentFirst || 0;
    const space = textW(' ', size, bold);
    for (const word of words) {
      const ww = textW(word, size, bold);
      if (cur === 0) { cur = ww; continue; }
      if (cur + space + ww <= width) cur += space + ww; else { lines++; cur = ww; }
    }
    return lines;
  }

  // ---- Month metadata carried forward. A new month inherits the previous month's standing
  // fields so nothing starts blank; the Task Book count is the one the user bumps each month.
  const META_FIELDS = ['taskBook', 'targetSolutions', 'policyStudy'];
  function monthsPresent(store) {
    const s = new Set();
    [...(store.shifts || []), ...(store.trainings || []), ...(store.events || [])].forEach(e => s.add(e.date.slice(0, 7)));
    Object.keys(store.months || {}).forEach(m => s.add(m));
    return [...s].sort();
  }
  // The most recent month before `ym` that has any standing field filled in.
  function priorMeta(store, ym) {
    const earlier = monthsPresent(store).filter(m => m < ym).reverse();
    for (const m of earlier) {
      const meta = (store.months || {})[m];
      if (meta && META_FIELDS.some(f => (meta[f] || '').trim())) return { month: m, meta };
    }
    return null;
  }
  // Fills blank standing fields on `ym` from the previous month. Returns the field names filled.
  function inheritMeta(store, ym) {
    const prior = priorMeta(store, ym); if (!prior) return [];
    store.months = store.months || {};
    const cur = store.months[ym] = store.months[ym] || {};
    const filled = [];
    for (const f of META_FIELDS) if (!(cur[f] || '').trim() && (prior.meta[f] || '').trim()) { cur[f] = prior.meta[f]; filled.push(f); }
    return filled;
  }

  // ---- Duplicate shifts: same date, time and apparatus. Returns the ids of the later copies.
  function duplicateIds(shifts) {
    const seen = new Set(), dup = new Set();
    for (const s of sortByDate(shifts)) {
      const k = s.date + '|' + (s.time || '') + '|' + (s.apparatus || '');
      if (seen.has(k)) dup.add(s.id); else seen.add(k);
    }
    return dup;
  }

  // ---- Training sentence. `compact` groups repeats of the same name so a month of QAB
  // sessions reads as one entry instead of seven.
  function trainingText(trainings, compact) {
    if (!trainings.length) return '';
    const totalHr = trainings.reduce((a, t) => a + (Number(t.hours) || 0), 0);
    let parts;
    if (!compact) {
      parts = trainings.map(t => t.name + ' (' + fmtHr(t.hours) + ' hr)');
    } else {
      const order = [], group = {};
      for (const t of trainings) {
        if (!group[t.name]) { group[t.name] = []; order.push(t.name); }
        group[t.name].push(Number(t.hours) || 0);
      }
      parts = order.map(name => {
        const hrs = group[name], sum = hrs.reduce((a, b) => a + b, 0);
        if (hrs.length === 1) return name + ' (' + fmtHr(sum) + ' hr)';
        const same = hrs.every(h => h === hrs[0]);
        return name + ' (' + hrs.length + ' sessions, ' + (same ? fmtHr(hrs[0]) + ' hr each, ' : '') + fmtHr(sum) + ' hr)';
      });
    }
    return parts.join('; ') + '. Total: ' + fmtHr(totalHr) + ' hours.';
  }

  // ---- Aggregation
  function aggregate(shifts) {
    const counts = {}; CATS.forEach(c => counts[c.key] = 0);
    for (const s of shifts) for (const k in (s.counts || {})) {
      if (!byKey[k]) throw new Error('Unknown category "' + k + '" on shift ' + s.date);
      counts[k] += Number(s.counts[k]) || 0;
    }
    counts.total = CATS.reduce((a, c) => a + counts[c.key], 0);
    const check = Object.keys(counts).filter(k => k !== 'total').reduce((a, k) => a + counts[k], 0);
    if (check !== counts.total) throw new Error('Category totals do not sum to the grand total');
    return counts;
  }

  // ---- Narrative
  function narrative(ym, shifts, counts, settings, meta, events, opts) {
    opts = opts || {};
    const who = settings.rank + ' ' + settings.name;
    const last = settings.rank + ' ' + settings.name.trim().split(/\s+/).pop();
    const apps = []; shifts.forEach(s => { if (s.apparatus && !apps.includes(s.apparatus)) apps.push(s.apparatus); });
    const nonzero = CATS.filter(c => counts[c.key] > 0).sort((a, b) => counts[b.key] - counts[a.key]);
    const phrases = nonzero.map(c => counts[c.key] + ' ' + (counts[c.key] === 1 ? c.one : c.many));
    let s = 'This evaluation covers ' + who + ' for ' + monthName(ym);
    if (apps.length) s += ', while assigned to ' + listJoin(apps);
    s += '. During this rating period, ' + last + ' responded to a total of ' + counts.total + ' call' + (counts.total === 1 ? '' : 's') + ' for service';
    if (phrases.length) s += ', including ' + listJoin(phrases);
    s += '.';
    // Standing-field sentence. Each clause appears only if its field is filled, so a blank
    // Task Book drops the clause instead of leaving a gap in the sentence.
    const clauses = [];
    if ((meta.taskBook || '').trim()) clauses.push(last + ' has completed ' + meta.taskBook.trim() + ' ' + settings.rank + ' Task Book items');
    if ((meta.targetSolutions || '').trim()) clauses.push('Target Solutions compliance is ' + meta.targetSolutions.trim().toLowerCase());
    if (clauses.length) s += ' ' + clauses.join(' and ').replace(/^./, c => c.toUpperCase()) + '.';
    // Pointer sentence only. The events themselves are printed in their own section, so this
    // is the first thing dropped when the memo needs to lose a line.
    if (events.length && !opts.omitEventPointer) s += ' Significant events during this period are detailed above.';
    if (meta.narrativeNote && meta.narrativeNote.trim()) s += ' ' + sentence(meta.narrativeNote);
    // Collapse any gap left by an empty field so the memo never shows a hole in a sentence.
    return s.replace(/\u2014/g, '-').replace(/\s{2,}/g, ' ').replace(/\s+([.,])/g, '$1').trim();
  }

  // ---- Assemble everything the generators need
  function buildData(ym, store, today, opts) {
    opts = opts || {};
    const settings = store.settings;
    const meta = (store.months && store.months[ym]) || {};
    const shifts = sortByDate((store.shifts || []).filter(s => s.date.startsWith(ym)));
    const trainings = sortByDate((store.trainings || []).filter(t => t.date.startsWith(ym)));
    const events = sortByDate((store.events || []).filter(e => e.date.startsWith(ym)));
    const counts = aggregate(shifts);
    const training = trainingText(trainings, !!opts.compactTraining);
    const significantEvents = events.length ? events.map(e => sentence(e.text)).join(' ') : 'No significant events to report.';
    const d = today || new Date();
    const data = {
      to: settings.to, from: settings.rank + ' ' + settings.name,
      subject: monthName(ym) + ' Monthly Summary',
      date: (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear(),
      shifts: shifts.map(s => ({ date: mmddyyyy(s.date), time: s.time || '', apparatus: s.apparatus || '', bc: s.bc || '', notes: s.notes || '' })),
      counts, shiftsWorked: shifts.length,
      targetSolutions: meta.targetSolutions || '', taskBook: meta.taskBook || '',
      significantEvents, training,
      policyNotes: meta.policyStudy || '',
      narrative: narrative(ym, shifts, counts, settings, meta, events, opts),
      signature: settings.rank + ' ' + settings.name + '   |   ' + settings.unit,
      rankLabel: settings.rank,
    };
    for (const k of ['narrative', 'training', 'significantEvents', 'policyNotes'])
      if (/\u2014/.test(data[k])) throw new Error('Em dash found in ' + k);
    return data;
  }

  // ---- Page fit. Word decides the real page breaks, so this estimates the rendered height
  // of the assembled memo and is calibrated against LibreOffice output. Used only to choose
  // between layout options; it never changes the numbers in the memo.
  const FIT = {
    page: 792, marginTop: 39, marginBottom: 36,
    letterhead: 80, headerRow: 14, rule: 6,
    headingPlain: 21.5, heading: 24,
    shiftHeadRow: 15.5, shiftRowLine: 11, shiftRowPad: 4, shiftNoteWidth: 210.5,
    incidentRow: 13.5, totalLine: 19, metaRow: 15,
    bodyLine: 11.3, bodyAfter: 3, bodyWidth: 504,
    signatureNormal: 36.5, signatureTight: 20.5, signatureLine: 11.3,
    // Calibrated against 360 rendered variants: one-page layouts measured up to 714.7 and
    // two-page layouts began at 715.0. The margin absorbs Word/LibreOffice differences.
    safety: 7,
  };
  function estimateHeight(data, opts) {
    opts = opts || {};
    const F = FIT;
    let h = F.letterhead + F.headerRow * 2 + F.rule + F.headingPlain;
    h += F.shiftHeadRow;
    for (const s of data.shifts) {
      const lines = Math.max(1, lineCount(s.notes, 9, F.shiftNoteWidth));
      h += lines * F.shiftRowLine + F.shiftRowPad;
    }
    // Tighter headings claim back 3.5pt of lead-in on each of the four ruled headings.
    const head = F.heading - (opts.tightHeadings ? 3.5 : 0);
    h += head + F.incidentRow * 9 + F.totalLine + F.metaRow * 2;
    h += head + lineCount(data.significantEvents, 9, F.bodyWidth) * F.bodyLine + F.bodyAfter;
    h += head + lineCount(data.training, 9, F.bodyWidth) * F.bodyLine + F.bodyAfter;
    h += head + lineCount(data.narrative, 9, F.bodyWidth) * F.bodyLine + 4;
    const polIndent = textW('Policy Study Notes:  ', 9, true);
    h += lineCount(data.policyNotes, 9, F.bodyWidth, false, polIndent) * F.bodyLine + F.bodyAfter;
    h += (opts.tightSignature ? F.signatureTight : F.signatureNormal) + F.signatureLine;
    return h;
  }
  const fitsOnePage = (data, opts) => estimateHeight(data, opts) <= FIT.page - FIT.marginTop - FIT.marginBottom - FIT.safety;

  // Assembles the memo, trying layout options in order of preference until one fits a single
  // page: as entered, then grouped training, then tightened signature spacing. Returns the
  // data plus the docx/pdf options and a note of what was adjusted.
  // Levers are applied cumulatively, least intrusive first. None of them change a number or
  // drop a logged entry; the harshest only removes a sentence that points at a section the
  // reader can already see, and tightens whitespace.
  const FIT_LEVERS = [
    { key: 'compactTraining',  note: 'grouped repeated training entries' },
    { key: 'omitEventPointer', note: 'dropped the summary line pointing at Significant Events' },
    { key: 'tightSignature',   note: 'tightened signature spacing' },
    { key: 'tightHeadings',    note: 'tightened heading spacing' },
  ];
  function buildFitted(ym, store, today) {
    for (let n = 0; n <= FIT_LEVERS.length; n++) {
      const opts = {}, notes = [];
      for (let i = 0; i < n; i++) { opts[FIT_LEVERS[i].key] = true; notes.push(FIT_LEVERS[i].note); }
      const data = buildData(ym, store, today, opts);
      if (fitsOnePage(data, opts)) return { data, opts, adjustments: notes, onePage: true };
    }
    // Still over a page. Deliver the full layout rather than a squeezed one that spills anyway,
    // and report why so the user can shorten something if they want one page.
    const opts = {};
    const data = buildData(ym, store, today, opts);
    return { data, opts, adjustments: [], onePage: false, overBy: Math.ceil(estimateHeight(data, opts) - (FIT.page - FIT.marginTop - FIT.marginBottom - FIT.safety)) };
  }

  // ---- DOCX (port of build_memo.js, hdrW fix applied)
  function buildDocx(data, letterheadBytes, opts) {
    opts = opts || {};
    const { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
      WidthType, AlignmentType, BorderStyle, ShadingType, VerticalAlign } = _docx;
    const NAVY = '2C3E5A', GRAY_HDR = 'E7E8E8', LT_BLUE = 'D9E1F2', BORDER_GRAY = 'BFBFBF', FONT = 'Arial', CONTENT_W = 10080;
    const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideHorizontal: noBorder, insideVertical: noBorder };
    const thin = { style: BorderStyle.SINGLE, size: 4, color: BORDER_GRAY };
    const run = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 18, bold: !!o.bold, color: o.color || '000000', italics: !!o.italics });
    const para = (children, o = {}) => new Paragraph({ children, alignment: o.align,
      spacing: { before: o.before ?? 0, after: o.after ?? 0, line: o.line, lineRule: 'auto' }, border: o.border });
    const headBefore = opts.tightHeadings ? 90 : 160;
    const heading = (text, o = {}) => new Paragraph({ children: [run(text, { bold: true, size: 21, color: NAVY })],
      spacing: { before: o.before ?? headBefore, after: o.after ?? 50 },
      border: o.noRule ? undefined : { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 2 } } });
    const cell = (children, o = {}) => new TableCell({ children, width: { size: o.w, type: WidthType.DXA },
      borders: o.borders || noBorders, shading: o.shade ? { type: ShadingType.CLEAR, fill: o.shade } : undefined,
      verticalAlign: VerticalAlign.CENTER, margins: { top: o.pad ?? 30, bottom: o.pad ?? 30, left: 60, right: 60 }, columnSpan: o.span });
    const children = [];
    children.push(new Paragraph({ children: [new ImageRun({ type: 'png', data: letterheadBytes, transformation: { width: 622, height: 100 } })], spacing: { after: 100 } }));
    const hdrW = [1150, 3650, 1000, 4280];
    const hdrRow = (l1, v1, l2, v2) => new TableRow({ children: [
      cell([para([run(l1, { bold: true })], { align: AlignmentType.RIGHT })], { w: hdrW[0] }),
      cell([para([run(v1)])], { w: hdrW[1] }),
      cell([para([run(l2, { bold: true })], { align: AlignmentType.RIGHT })], { w: hdrW[2] }),
      cell([para([run(v2)])], { w: hdrW[3] })] });
    children.push(new Table({ columnWidths: hdrW, width: { size: CONTENT_W, type: WidthType.DXA }, borders: noBorders,
      rows: [hdrRow('TO:', data.to, 'FROM:', data.from), hdrRow('SUBJECT:', data.subject, 'DATE:', data.date)] }));
    children.push(para([run('')], { before: 80, after: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 14, color: NAVY, space: 1 } } }));
    children.push(heading('SHIFT ASSIGNMENTS', { noRule: true, before: 100 }));
    const sw = [1250, 1350, 1400, 1750, 4330];
    const box = { top: thin, bottom: thin, left: thin, right: thin };
    const shiftHdr = new TableRow({ children: ['Date', 'Time', 'Apparatus', 'Battalion Chief', 'Shift Notes / Significant Events'].map((t, i) =>
      cell([para([run(t, { bold: true, color: 'FFFFFF' })], { align: AlignmentType.CENTER })], { w: sw[i], shade: NAVY, borders: box, pad: 45 })) });
    const shiftRows = data.shifts.map(s => new TableRow({ children: [s.date, s.time, s.apparatus, s.bc, s.notes].map((t, i) =>
      cell([para([run(t || '')], { align: i === 4 ? AlignmentType.LEFT : AlignmentType.CENTER })], { w: sw[i], borders: box, pad: 40 })) }));
    children.push(new Table({ columnWidths: sw, width: { size: CONTENT_W, type: WidthType.DXA }, rows: [shiftHdr, ...shiftRows] }));
    children.push(heading('INCIDENT TOTALS'));
    const iw = [3250, 1150, 780, 3550, 1350];
    const hInside = { top: thin, bottom: thin, left: noBorder, right: noBorder };
    const hc = t => cell([para([run(t, { bold: true, size: 16, color: '595959' })], { align: AlignmentType.CENTER })], { w: 0, shade: GRAY_HDR, borders: hInside });
    const incHdr = new TableRow({ children: [
      cell([para([run('Category', { bold: true, size: 16, color: '595959' })], { align: AlignmentType.CENTER })], { w: iw[0], shade: GRAY_HDR, borders: hInside }),
      cell([para([run('Count', { bold: true, size: 16, color: '595959' })], { align: AlignmentType.CENTER })], { w: iw[1], shade: GRAY_HDR, borders: hInside }),
      cell([para([run('')])], { w: iw[2], borders: noBorders }),
      cell([para([run('Category', { bold: true, size: 16, color: '595959' })], { align: AlignmentType.CENTER })], { w: iw[3], shade: GRAY_HDR, borders: hInside }),
      cell([para([run('Count', { bold: true, size: 16, color: '595959' })], { align: AlignmentType.CENTER })], { w: iw[4], shade: GRAY_HDR, borders: hInside })] });
    const left = CATS.slice(0, 8), right = CATS.slice(8);
    const incRows = left.map((lc, i) => { const rc = right[i]; return new TableRow({ children: [
      cell([para([run(lc.label)])], { w: iw[0], borders: hInside }),
      cell([para([run(String(data.counts[lc.key]), { bold: true, color: NAVY })], { align: AlignmentType.CENTER })], { w: iw[1], borders: hInside }),
      cell([para([run('')])], { w: iw[2], borders: noBorders }),
      cell([para([run(rc.label)])], { w: iw[3], borders: hInside }),
      cell([para([run(String(data.counts[rc.key]), { bold: true, color: NAVY })], { align: AlignmentType.CENTER })], { w: iw[4], borders: hInside })] }); });
    children.push(new Table({ columnWidths: iw, width: { size: CONTENT_W, type: WidthType.DXA }, rows: [incHdr, ...incRows] }));
    children.push(para([run('TOTAL INCIDENTS:   ', { bold: true }), run(String(data.counts.total), { bold: true, size: 22, color: NAVY })],
      { align: AlignmentType.RIGHT, before: 40, after: 60, border: { top: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 2 } } }));
    const swW = [5040, 5040];
    children.push(new Table({ columnWidths: swW, width: { size: CONTENT_W, type: WidthType.DXA }, borders: noBorders, rows: [new TableRow({ children: [
      cell([para([run('Shifts Worked This Month:', { bold: true })])], { w: swW[0] }),
      cell([para([run(String(data.shiftsWorked), { bold: true })], { align: AlignmentType.CENTER })], { w: swW[1], shade: LT_BLUE })] })] }));
    const tsW = [2000, 3040, 2000, 3040];
    children.push(new Table({ columnWidths: tsW, width: { size: CONTENT_W, type: WidthType.DXA }, borders: noBorders, rows: [new TableRow({ children: [
      cell([para([run('Target Solutions:', { bold: true })])], { w: tsW[0] }),
      cell([para([run(data.targetSolutions)])], { w: tsW[1] }),
      cell([para([run(data.rankLabel + ' Task Book:', { bold: true })])], { w: tsW[2] }),
      cell([para([run(data.taskBook)])], { w: tsW[3] })] })] }));
    children.push(heading('SIGNIFICANT EVENTS'));
    children.push(para([run(data.significantEvents)], { after: 60, line: 252 }));
    children.push(heading('TRAINING & ACTIVITIES'));
    children.push(para([run(data.training)], { after: 60, line: 252 }));
    children.push(heading('EVALUATION SUMMARY'));
    children.push(para([run(data.narrative)], { after: 80, line: 252 }));
    children.push(para([run('Policy Study Notes:  ', { bold: true }), run(data.policyNotes)], { after: 60 }));
    children.push(para([run('Respectfully submitted,')],
      opts.tightSignature ? { before: 60, after: 120 } : { before: 200, after: 300 }));
    children.push(para([run(data.signature, { bold: true })]));
    const doc = new Document({ styles: { default: { document: { run: { font: FONT, size: 18 } } } },
      sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 780, bottom: 720, left: 1080, right: 1080 } } }, children }] });
    return typeof Blob !== 'undefined' && typeof window !== 'undefined' ? Packer.toBlob(doc) : Packer.toBuffer(doc);
  }

  // ---- PDF (drawn to the same measurements; twips / 20 = points)
  async function buildPdf(data, letterheadBytes, opts) {
    opts = opts || {};
    const { PDFDocument, StandardFonts, rgb } = _pdf;
    const pdf = await PDFDocument.create();
    const reg = await pdf.embedFont(StandardFonts.Helvetica);
    const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
    const img = await pdf.embedPng(letterheadBytes);
    const NAVY = rgb(0x2c / 255, 0x3e / 255, 0x5a / 255), GRAY = rgb(0xbf / 255, 0xbf / 255, 0xbf / 255);
    const GRAY_HDR = rgb(0xe7 / 255, 0xe8 / 255, 0xe8 / 255), LT_BLUE = rgb(0xd9 / 255, 0xe1 / 255, 0xf2 / 255);
    const DK = rgb(0x59 / 255, 0x59 / 255, 0x59 / 255), BLACK = rgb(0, 0, 0), WHITE = rgb(1, 1, 1);
    const W = 612, H = 792, L = 54, R = 558, CW = 504;
    let page = pdf.addPage([W, H]);
    let y = H - 39;
    const wrap = (text, font, size, width) => {
      const words = String(text || '').split(/\s+/).filter(Boolean); const lines = []; let cur = '';
      for (const w of words) { const t = cur ? cur + ' ' + w : w;
        if (font.widthOfTextAtSize(t, size) <= width || !cur) cur = t; else { lines.push(cur); cur = w; } }
      if (cur) lines.push(cur); return lines.length ? lines : [''];
    };
    const text = (s, x, yy, o = {}) => {
      const f = o.bold ? bold : reg, size = o.size || 9; let tx = x;
      const tw = f.widthOfTextAtSize(s, size);
      if (o.align === 'center') tx = x + (o.w - tw) / 2; else if (o.align === 'right') tx = x + o.w - tw;
      page.drawText(s, { x: tx, y: yy, size, font: f, color: o.color || BLACK });
    };
    const hline = (x1, x2, yy, w, c) => page.drawLine({ start: { x: x1, y: yy }, end: { x: x2, y: yy }, thickness: w, color: c });
    const ensure = need => { if (y - need < 36) { page = pdf.addPage([W, H]); y = H - 39; } };
    const paragraph = (s, o = {}) => { const lines = wrap(s, o.bold ? bold : reg, 9, CW); const lh = 11.3;
      ensure(lines.length * lh); for (const ln of lines) { text(ln, L, y - 8, o); y -= lh; } y -= (o.after || 0); };
    const headGap = opts.tightHeadings ? 4.5 : 8;
    const headingRule = (t, rule = true) => { ensure(24); y -= headGap; text(t, L, y - 9, { bold: true, size: 10.5, color: NAVY }); y -= 13.5;
      if (rule) { hline(L, R, y, 1.5, NAVY); } y -= 2.5; };

    // Letterhead 622x100 px at 96 dpi -> 466.5 x 75 pt
    page.drawImage(img, { x: L, y: y - 75, width: 466.5, height: 75 }); y -= 80;
    // Header table
    const hw = [57.5, 182.5, 50, 214];
    const hdr = (l1, v1, l2, v2) => { let x = L; const yy = y - 9;
      text(l1, x, yy, { bold: true, align: 'right', w: hw[0] - 3 }); x += hw[0];
      text(v1, x + 3, yy); x += hw[1];
      text(l2, x, yy, { bold: true, align: 'right', w: hw[2] - 3 }); x += hw[2];
      text(v2, x + 3, yy); y -= 14; };
    hdr('TO:', data.to, 'FROM:', data.from); hdr('SUBJECT:', data.subject, 'DATE:', data.date);
    y -= 6; hline(L, R, y, 1.75, NAVY); y -= 2;
    headingRule('SHIFT ASSIGNMENTS', false);
    // Shift table
    const sw = [62.5, 67.5, 70, 87.5, 216.5];
    const drawRow = (cells, o) => { const lines = cells.map((c, i) => wrap(c, o.bold ? bold : reg, 9, sw[i] - 6));
      const n = Math.max(...lines.map(l => l.length)); const rh = n * 10.5 + (o.pad || 2) * 2; ensure(rh);
      if (o.fill) page.drawRectangle({ x: L, y: y - rh, width: CW, height: rh, color: o.fill });
      let x = L; cells.forEach((c, i) => { page.drawRectangle({ x, y: y - rh, width: sw[i], height: rh, borderColor: GRAY, borderWidth: 0.5 });
        const align = o.align || (i === 4 ? 'left' : 'center'); let ty = y - (o.pad || 2) - 7.8 - ((n - lines[i].length) * 10.5) / 2;
        for (const ln of lines[i]) { text(ln, x + 3, ty, { bold: o.bold, color: o.color, align, w: sw[i] - 6 }); ty -= 10.5; } x += sw[i]; });
      y -= rh; };
    drawRow(['Date', 'Time', 'Apparatus', 'Battalion Chief', 'Shift Notes / Significant Events'], { bold: true, color: WHITE, fill: NAVY, pad: 2.5, align: 'center' });
    for (const s of data.shifts) drawRow([s.date, s.time, s.apparatus, s.bc, s.notes], { pad: 2 });
    headingRule('INCIDENT TOTALS');
    // Incident totals
    const iw = [162.5, 57.5, 39, 177.5, 67.5];
    const incRow = (l, lv, r, rv, o = {}) => { const rh = 13.5; ensure(rh); let x = L;
      const seg = (xx, ww, fill) => { if (fill) page.drawRectangle({ x: xx, y: y - rh, width: ww, height: rh, color: fill });
        hline(xx, xx + ww, y, 0.5, GRAY); hline(xx, xx + ww, y - rh, 0.5, GRAY); };
      seg(x, iw[0] + iw[1], o.fill); seg(x + iw[0] + iw[1] + iw[2], iw[3] + iw[4], o.fill);
      const ty = y - 9.5;
      text(l, x + 3, ty, { bold: o.bold, size: o.size, color: o.color, align: o.hdr ? 'center' : 'left', w: iw[0] - 6 }); x += iw[0];
      text(lv, x, ty, { bold: true, size: o.size, color: o.hdr ? o.color : NAVY, align: 'center', w: iw[1] }); x += iw[1] + iw[2];
      text(r, x + 3, ty, { bold: o.bold, size: o.size, color: o.color, align: o.hdr ? 'center' : 'left', w: iw[3] - 6 }); x += iw[3];
      text(rv, x, ty, { bold: true, size: o.size, color: o.hdr ? o.color : NAVY, align: 'center', w: iw[4] });
      y -= rh; };
    incRow('Category', 'Count', 'Category', 'Count', { hdr: true, bold: true, size: 8, color: DK, fill: GRAY_HDR });
    for (let i = 0; i < 8; i++) { const lc = CATS[i], rc = CATS[i + 8]; incRow(lc.label, String(data.counts[lc.key]), rc.label, String(data.counts[rc.key])); }
    // Total incidents
    y -= 2; hline(L, R, y, 1.5, NAVY); y -= 3; ensure(16);
    { const lab = 'TOTAL INCIDENTS:   ', num = String(data.counts.total);
      const nw = bold.widthOfTextAtSize(num, 11), lw = bold.widthOfTextAtSize(lab, 9);
      text(num, R - nw, y - 10, { bold: true, size: 11, color: NAVY }); text(lab, R - nw - lw, y - 10, { bold: true }); y -= 16; }
    // Shifts worked
    ensure(15); page.drawRectangle({ x: L + 252, y: y - 13, width: 252, height: 13, color: LT_BLUE });
    text('Shifts Worked This Month:', L + 3, y - 9.5, { bold: true }); text(String(data.shiftsWorked), L + 252, y - 9.5, { bold: true, align: 'center', w: 252 }); y -= 15;
    ensure(15); { let x = L; text('Target Solutions:', x + 3, y - 9.5, { bold: true }); x += 100; text(data.targetSolutions, x + 3, y - 9.5); x += 152;
      text(data.rankLabel + ' Task Book:', x + 3, y - 9.5, { bold: true }); x += 100; text(data.taskBook, x + 3, y - 9.5); y -= 15; }
    headingRule('SIGNIFICANT EVENTS'); paragraph(data.significantEvents, { after: 3 });
    headingRule('TRAINING & ACTIVITIES'); paragraph(data.training, { after: 3 });
    headingRule('EVALUATION SUMMARY'); paragraph(data.narrative, { after: 4 });
    ensure(14); { const lab = 'Policy Study Notes:  '; text(lab, L, y - 8, { bold: true });
      const lw = bold.widthOfTextAtSize(lab, 9); const lines = wrap(data.policyNotes, reg, 9, CW - lw);
      text(lines[0], L + lw, y - 8); y -= 11.3; for (const ln of lines.slice(1)) { text(ln, L, y - 8); y -= 11.3; } }
    y -= opts.tightSignature ? 3 : 10; ensure(40); text('Respectfully submitted,', L, y - 8);
    y -= opts.tightSignature ? 16 : 26; text(data.signature, L, y - 8, { bold: true });
    return pdf.save();
  }

  // ---- Log text export / import (same pipe format as the Drive doc)
  function exportText(store) {
    const s = store.settings; const out = [];
    out.push('SJFD Shift Log - ' + s.rank + ' ' + s.name + ', Badge ' + s.badge + ', ' + s.unit, '',
      'Shift line format: MM/DD | time on-off | apparatus | BC | incident counts (code and number, comma separated) | optional shift note after //', '');
    const months = new Set([...(store.shifts || []), ...(store.trainings || []), ...(store.events || [])].map(e => e.date.slice(0, 7)));
    Object.keys(store.months || {}).forEach(m => months.add(m));
    for (const ym of [...months].sort()) {
      const m = (store.months && store.months[ym]) || {};
      out.push(monthName(ym), '', 'Task Book: ' + (m.taskBook || ''), 'Target Solutions: ' + (m.targetSolutions || ''), 'Policy Study: ' + (m.policyStudy || ''), '');
      for (const sh of sortByDate((store.shifts || []).filter(e => e.date.startsWith(ym)))) out.push(shiftLine(sh));
      out.push('', 'Training:');
      for (const t of sortByDate((store.trainings || []).filter(e => e.date.startsWith(ym)))) out.push(mmdd(t.date) + ' | ' + t.name + ' | ' + fmtHr(t.hours) + ' hr');
      out.push('', 'Significant Events:');
      const ev = sortByDate((store.events || []).filter(e => e.date.startsWith(ym)));
      if (ev.length) ev.forEach(e => out.push(mmdd(e.date) + ' | ' + e.text)); else out.push('(none)');
      if (m.narrativeNote) out.push('', 'Narrative note: ' + m.narrativeNote);
      out.push('');
    }
    return out.join('\n');
  }
  function shiftLine(s) {
    const c = CODE_ORDER.filter(code => (s.counts || {})[byCode[code].key] > 0).map(code => code + ' ' + s.counts[byCode[code].key]).join(', ');
    let line = mmdd(s.date) + ' | ' + (s.time || '') + ' | ' + (s.apparatus || '') + ' | ' + (s.bc || '') + ' | ' + c;
    if (s.notes) line += ' // ' + s.notes; return line;
  }
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  // Parse pipe-format text. Returns {shifts, trainings, events, months, errors}. Year comes from month headings.
  function importText(txt) {
    const res = { shifts: [], trainings: [], events: [], months: {}, errors: [] };
    let ym = null, section = 'shifts';
    const mre = new RegExp('^(' + MONTHS.join('|') + ')\\s+(\\d{4})\\s*$', 'i');
    txt.split(/\r?\n/).forEach((raw, i) => {
      const line = raw.trim(); if (!line) return;
      const mm = line.match(mre);
      if (mm) { ym = mm[2] + '-' + pad2(MONTHS.findIndex(m => m.toLowerCase() === mm[1].toLowerCase()) + 1); res.months[ym] = res.months[ym] || {}; section = 'shifts'; return; }
      if (!ym) return; // header text before the first month
      let kv;
      if ((kv = line.match(/^Task Book:\s*(.*)$/i))) { res.months[ym].taskBook = kv[1].trim(); return; }
      if ((kv = line.match(/^Target Solutions:\s*(.*)$/i))) { res.months[ym].targetSolutions = kv[1].trim(); return; }
      if ((kv = line.match(/^Policy Study:\s*(.*)$/i))) { res.months[ym].policyStudy = kv[1].trim(); return; }
      if ((kv = line.match(/^Narrative note:\s*(.*)$/i))) { res.months[ym].narrativeNote = kv[1].trim(); return; }
      if (/^Training:?$/i.test(line)) { section = 'training'; return; }
      if (/^Significant Events:?$/i.test(line)) { section = 'events'; return; }
      if (/^\(none\)$/i.test(line)) return;
      const dm = line.match(/^(\d{2})\/(\d{2})\s*\|/);
      if (!dm) { res.errors.push('Line ' + (i + 1) + ' not recognized: ' + line); return; }
      const date = ym.slice(0, 4) + '-' + dm[1] + '-' + dm[2];
      const [main, note] = line.split('//').map(s => s.trim());
      const parts = main.split('|').map(s => s.trim());
      if (section === 'events') { res.events.push({ id: uid(), date, text: parts.slice(1).join(' | ') }); return; }
      if (section === 'training' || (parts.length === 3 && /^[\d.]+\s*hr/i.test(parts[2]))) {
        const h = parseFloat(parts[2]); if (parts.length < 3 || isNaN(h)) { res.errors.push('Line ' + (i + 1) + ' bad training line: ' + line); return; }
        res.trainings.push({ id: uid(), date, name: parts[1], hours: h }); return;
      }
      if (parts.length < 4) { res.errors.push('Line ' + (i + 1) + ' bad shift line: ' + line); return; }
      const counts = {};
      for (const tok of (parts[4] || '').split(',').map(s => s.trim()).filter(Boolean)) {
        const m = tok.match(/^([a-z]+)\s+(\d+)$/i);
        if (!m || !byCode[m[1].toLowerCase()]) { res.errors.push('Line ' + (i + 1) + ' bad count "' + tok + '": ' + line); return; }
        counts[byCode[m[1].toLowerCase()].key] = (counts[byCode[m[1].toLowerCase()].key] || 0) + parseInt(m[2], 10);
      }
      res.shifts.push({ id: uid(), date, time: parts[1], apparatus: parts[2], bc: parts[3], counts, notes: note || '' });
    });
    return res;
  }
  // Additive merge with duplicate detection
  function mergeImport(store, imp) {
    const key = s => s.date + '|' + s.time + '|' + s.apparatus;
    const have = new Set((store.shifts || []).map(key)); let added = 0;
    for (const s of imp.shifts) if (!have.has(key(s))) { store.shifts.push(s); have.add(key(s)); added++; }
    const tk = t => t.date + '|' + t.name + '|' + t.hours; const th = new Set((store.trainings || []).map(tk));
    for (const t of imp.trainings) if (!th.has(tk(t))) { store.trainings.push(t); th.add(tk(t)); added++; }
    const ek = e => e.date + '|' + e.text; const eh = new Set((store.events || []).map(ek));
    for (const e of imp.events) if (!eh.has(ek(e))) { store.events.push(e); eh.add(ek(e)); added++; }
    store.months = store.months || {};
    for (const ym in imp.months) store.months[ym] = Object.assign({}, store.months[ym] || {}, imp.months[ym]);
    return added;
  }

  return { CATS, CODE_ORDER, byCode, byKey, MONTHS, monthName, aggregate, narrative, buildData, buildDocx, buildPdf,
    exportText, importText, mergeImport, shiftLine, uid, mmdd, fmtHr,
    META_FIELDS, monthsPresent, priorMeta, inheritMeta, duplicateIds, trainingText,
    estimateHeight, fitsOnePage, buildFitted, FIT, lineCount };
})();
if (typeof module !== 'undefined') module.exports = MemoLib;
