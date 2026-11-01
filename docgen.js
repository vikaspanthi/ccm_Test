/* Dependency-free .docx writer for the CALTech Course Manual (CCM).
   Builds WordprocessingML by hand and packs it in a ZIP (deflate via CompressionStream when available).
   Browser: window.CCMDoc.generate(cfg) -> Promise<Uint8Array>.  Node 18+: require('./docgen.js'). */
(function (root) {
  'use strict';

  // ---------------------------------------------------------------- ZIP ----
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return buf => { let c = 0xffffffff; for (let i = 0; i < buf.length; i++) c = t[(c ^ buf[i]) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  })();

  async function deflateRaw(bytes) {
    if (typeof CompressionStream === 'undefined') return null;
    try {
      const cs = new CompressionStream('deflate-raw');
      const w = cs.writable.getWriter(); w.write(bytes); w.close();
      const chunks = []; const rd = cs.readable.getReader();
      for (;;) { const { done, value } = await rd.read(); if (done) break; chunks.push(value); }
      let len = 0; chunks.forEach(c => (len += c.length));
      const out = new Uint8Array(len); let o = 0; chunks.forEach(c => { out.set(c, o); o += c.length; });
      return out;
    } catch (e) { return null; }
  }

  async function zip(files) { // files: [{name, data:Uint8Array}]
    const enc = new TextEncoder();
    const parts = [], central = []; let offset = 0;
    const u16 = n => [n & 255, (n >> 8) & 255];
    const u32 = n => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
    for (const f of files) {
      const name = enc.encode(f.name);
      const crc = CRC(f.data);
      let comp = await deflateRaw(f.data), method = 8;
      if (!comp || comp.length >= f.data.length) { comp = f.data; method = 0; }
      const head = new Uint8Array([0x50, 0x4b, 3, 4, ...u16(20), ...u16(0x0800), ...u16(method), ...u16(0), ...u16(0x21),
        ...u32(crc), ...u32(comp.length), ...u32(f.data.length), ...u16(name.length), ...u16(0)]);
      parts.push(head, name, comp);
      central.push(new Uint8Array([0x50, 0x4b, 1, 2, ...u16(20), ...u16(20), ...u16(0x0800), ...u16(method), ...u16(0), ...u16(0x21),
        ...u32(crc), ...u32(comp.length), ...u32(f.data.length), ...u16(name.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset)]), name);
      offset += head.length + name.length + comp.length;
    }
    let csize = 0; central.forEach(c => (csize += c.length));
    const end = new Uint8Array([0x50, 0x4b, 5, 6, 0, 0, 0, 0, ...u16(files.length), ...u16(files.length), ...u32(csize), ...u32(offset), 0, 0]);
    const all = [...parts, ...central, end]; let total = 0; all.forEach(a => (total += a.length));
    const out = new Uint8Array(total); let o = 0; all.forEach(a => { out.set(a, o); o += a.length; });
    return out;
  }

  // ------------------------------------------------------- XML helpers ----
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const FULL = 9746; // text width in dxa (A4, 0.75" margins)

  function run(text, o = {}) {
    const rpr = (o.b ? '<w:b/>' : '') + (o.i ? '<w:i/>' : '') + (o.u ? '<w:u w:val="single"/>' : '') +
      (o.color ? `<w:color w:val="${o.color}"/>` : '') + (o.sz ? `<w:sz w:val="${Math.round(o.sz * 2)}"/><w:szCs w:val="${Math.round(o.sz * 2)}"/>` : '') +
      (o.sup ? '<w:vertAlign w:val="superscript"/>' : '');
    return `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
  }
  function para(content, o = {}) {
    const runs = Array.isArray(content) ? content.join('') : (typeof content === 'string' && content.startsWith('<w:r') ? content : run(content, o));
    const ppr = (o.keepNext ? '<w:keepNext/>' : '') + (o.pb ? '<w:pageBreakBefore/>' : '') +
      `<w:spacing w:before="${o.before || 0}" w:after="${o.after == null ? 60 : o.after}"${o.line ? ` w:line="${o.line}" w:lineRule="auto"` : ''}/>` +
      (o.ind ? `<w:ind w:left="${o.ind}" w:hanging="${o.hang || 0}"/>` : '') +
      (o.align ? `<w:jc w:val="${o.align}"/>` : '');
    return `<w:p><w:pPr>${ppr}</w:pPr>${runs}</w:p>`;
  }
  const blank = n => para('', { after: 0 }).repeat(n || 1);
  const pageBreak = () => '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

  // cell: {t: string|xml, w, span, shade, b, align, sz, v:'center'|'top', i}
  function cellXml(c, colW) {
    if (typeof c === 'string') c = { t: c };
    const w = colW;
    const tcpr = `<w:tcW w:w="${w}" w:type="dxa"/>` + (c.span > 1 ? `<w:gridSpan w:val="${c.span}"/>` : '') +
      (c.shade ? `<w:shd w:val="clear" w:color="auto" w:fill="${c.shade}"/>` : '') + `<w:vAlign w:val="${c.v || 'center'}"/>`;
    let inner;
    if (typeof c.t === 'string' && c.t.startsWith('<w:p')) inner = c.t;
    else {
      const lines = String(c.t == null ? '' : c.t).split('\n');
      inner = lines.map(l => para(l, { b: c.b, i: c.i, sz: c.sz || 10, align: c.align, after: 30 })).join('');
    }
    return `<w:tc><w:tcPr>${tcpr}</w:tcPr>${inner}</w:tc>`;
  }
  function table(rows, widths, o = {}) {
    const grid = widths.map(w => `<w:gridCol w:w="${w}"/>`).join('');
    const bd = o.noBorder ? '' : '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
    const total = widths.reduce((a, b) => a + b, 0);
    let x = `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/>${o.center ? '<w:jc w:val="center"/>' : ''}${bd}<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${grid}</w:tblGrid>`;
    rows.forEach((row, ri) => {
      let col = 0;
      const cells = row.map(c => {
        const span = (typeof c === 'object' && c.span) || 1;
        const w = widths.slice(col, col + span).reduce((a, b) => a + b, 0); col += span;
        return cellXml(c, w);
      }).join('');
      x += `<w:tr><w:trPr><w:cantSplit/>${o.header && ri === 0 ? '<w:tblHeader/>' : ''}</w:trPr>${cells}</w:tr>`;
    });
    return x + '</w:tbl>';
  }
  const spacer = () => para('', { after: 80 });

  function logoXml(cx, cy) {
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="Logo"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  }

  const G = '00B050', Y = 'FFC000', R = 'FF0000', HEAD = 'D9D9D9';
  const LV = { 3: G, 2: Y, 1: R, 0: null };
  const title = (t, o = {}) => para(t, Object.assign({ b: true, sz: 14, align: 'center', after: 120, before: 120, keepNext: true }, o));

  // --------------------------------------------------------- sections ----
  function cover(c, hasLogo) {
    const k = c.course;
    let x = '';
    if (hasLogo) x += para([logoXml(2926080, 1152144)], { after: 240 });
    else x += blank(3);
    x += blank(2);
    x += para([run(k.manualTitle || 'CALTech', { b: true, sz: 36 }), run('*', { b: true, sz: 36, sup: true }), run(' Course Manual', { b: true, sz: 36 })], { align: 'center', after: 400 });
    x += blank(1);
    x += para(`${k.code} – ${k.title}`, { b: true, sz: 16, align: 'center', after: 300 });
    x += para(`Slot: ${k.slot}, Class Number: ${k.classNo || '________'}`, { b: true, sz: 14, align: 'center', after: 40 });
    x += para(k.semester, { b: true, sz: 14, align: 'center', after: 500 });
    x += blank(3);
    x += para(k.faculty, { b: true, sz: 16, align: 'center', after: 40 });
    if (k.designation) x += para(k.designation, { b: true, sz: 16, align: 'center', after: 40 });
    if (k.school) x += para(k.school, { b: true, sz: 16, align: 'center', after: 40 });
    x += blank(4);
    x += para('*Collaborative and Active Learning through Technology', { sz: 9 });
    return x + pageBreak();
  }

  function contents() {
    const items = ['Students Outcomes & Program Outcome', 'Syllabus', 'CO- PO* / SO** Mapping', 'Course Plan',
      'Assessment Tools and Weightage', 'CO Attainment', 'Sessions Plan', 'Annexure -1 (Model Question Paper)'];
    let x = para('Contents', { b: true, sz: 20, align: 'center', after: 240 });
    const rows = [[{ t: 'Sl. No', b: true, align: 'center', sz: 12, shade: HEAD }, { t: 'Description', b: true, sz: 12, shade: HEAD }]];
    items.forEach((t, i) => rows.push([{ t: String(i + 1), align: 'center', sz: 12 }, { t, sz: 12 }]));
    x += table(rows, [1200, 6400], { center: true });
    x += spacer() + para('•  PO – Program Outcome as defined by NBA & NAAC', { sz: 9 }) + para('•  SO – Student Outcome as defined by ABET', { sz: 9 });
    return x + pageBreak();
  }

  function outcomes(c) {
    const D = root.CCM_DATA || {};
    let x = para('ABET Student Outcomes', { b: true, sz: 13, align: 'center', after: 120 });
    x += para([run('Student Outcomes', { b: true, sz: 11 }), run(' are outcomes (1) through (7), plus any additional outcomes that may be articulated by the program.', { sz: 11 })], { align: 'both' });
    (D.so || []).forEach((t, i) => (x += para(`SO${i + 1}.  ${t}`, { sz: 10.5, ind: 720, hang: 720, align: 'both' })));
    x += para('Reference: https://www.abet.org/accreditation/accreditation-criteria/criteria-for-accrediting-engineering-programs-2020-2021/', { sz: 8, before: 80 });
    x += para('NBA Programme Outcome', { b: true, sz: 13, align: 'center', before: 240, after: 120 });
    x += para('Engineering Graduates will be able to:', { sz: 11 });
    (D.po || []).forEach((t, i) => (x += para(`PO${i + 1}.  ${t}`, { sz: 10.5, ind: 720, hang: 720, align: 'both' })));
    if ((c.psos || []).length) {
      x += para('Program Specific Outcomes (PSOs)', { b: true, sz: 11, before: 120 });
      c.psos.forEach((t, i) => (x += para(`PSO${i + 1}.  ${t}`, { sz: 10.5, ind: 720, hang: 720, align: 'both' })));
    } else x += para('(B) PROGRAM SPECIFIC OUTCOMES (PSOs) Program should specify 2-4 program specific outcomes.', { sz: 10.5, before: 120 });
    x += para('Reference: https://www.nbaind.org/Files/ug-sar-tier-i-final-ver-1.0.pdf', { sz: 8 });
    return x + pageBreak();
  }

  function syllabus(c) {
    const k = c.course;
    let x = title('SYLLABUS');
    x += table([
      [{ t: k.title.toUpperCase(), b: true, sz: 12, align: 'center', span: 1 }, { t: 'Course Type', b: true, sz: 10 }, { t: k.type, b: true, sz: 10, align: 'center' }],
      [{ t: `Course Code: ${k.code}`, b: true, sz: 10 }, { t: 'Credits', b: true, sz: 10 }, { t: String(k.credits), b: true, sz: 10, align: 'center' }],
      [{ t: `Pre-requisite: ${k.prereq || 'NIL'}`, span: 3, sz: 10 }],
    ], [6400, 1900, 1446]);
    // objectives & outcomes
    const objs = (c.objectives || []).map((t, i) => `${i + 1}.  ${t}`).join('\n');
    const cos = (c.cos || []).map(o => `${o.id}.  ${o.text}`).join('\n');
    x += table([
      [{ t: 'Course Objectives:', b: true, sz: 10, v: 'top' }],
      [{ t: objs, sz: 10, v: 'top' }],
      [{ t: 'Course Outcomes (CO):', b: true, sz: 10 }],
      [{ t: 'At the end of the course, students should be able to\n' + cos, sz: 10, v: 'top' }],
    ], [FULL]);
    // CO-PO correlation mini table
    const hdr = ['CO \\ PO'].concat(Array.from({ length: 12 }, (_, i) => 'PO' + (i + 1)), ['PSO1', 'PSO2']);
    const w0 = 800, wr = Math.floor((FULL - w0) / 14);
    const rows = [hdr.map(h => ({ t: h, b: true, align: 'center', sz: 8, shade: HEAD }))];
    (c.cos || []).forEach(o => rows.push([{ t: o.id, b: true, sz: 8, align: 'center' }].concat((c.mapPO[o.id] || []).map(v => ({ t: v ? String(v) : '-', align: 'center', sz: 8 })))));
    x += table([[{ t: 'Correlation of COs with POs (3 – Strong, 2 – Moderate, 1 – Weak)', b: true, sz: 10, span: 15 }]].concat(rows), [w0].concat(Array(14).fill(wr)));
    // units
    const ur = [[{ t: 'Unit', b: true, align: 'center', sz: 10, shade: HEAD }, { t: 'Topics to be discussed', b: true, align: 'center', sz: 10, shade: HEAD }, { t: 'Hrs.', b: true, align: 'center', sz: 10, shade: HEAD }]];
    let tot = 0;
    (c.units || []).forEach((u, i) => {
      const h = parseFloat(u.hours) || 0; tot += h;
      const tag = [u.sos ? 'SOs: ' + u.sos : '', u.cos ? 'COs: ' + u.cos : ''].filter(Boolean).join('; ');
      ur.push([{ t: 'Unit ' + (i + 1), b: true, align: 'center', sz: 10 }, { t: (u.title ? u.title + ' – ' : '') + String(u.topics).replace(/\s*\n\s*/g, ' ') + (tag ? ` [${tag}]` : ''), sz: 10, v: 'top' }, { t: String(h).padStart(2, '0'), b: true, align: 'center', sz: 10 }]);
    });
    const gh = parseFloat(c.guestHours) || 0; tot += gh;
    if (gh) ur.push(['', { t: 'Guest Lecture on Contemporary Topics', b: true, sz: 10 }, { t: String(gh).padStart(2, '0'), b: true, align: 'center', sz: 10 }]);
    if (c.tools) ur.push(['', { t: 'Suggested Tools: ' + c.tools, b: true, sz: 10 }, '']);
    ur.push(['', { t: 'Total Hours:', b: true, align: 'right', sz: 10 }, { t: String(tot), b: true, align: 'center', sz: 10 }]);
    x += table(ur, [800, FULL - 800 - 800, 800], { header: true });
    // info
    const num = a => (a || []).map((t, i) => `${i + 1}.  ${t}`).join('\n');
    x += table([
      [{ t: 'Mode of Teaching and Learning:\n' + 'Flipped Classroom, Activity-Based Teaching/Learning, Digital/Computer-based models, wherever possible to augment lecture for practice/tutorial and minimum 2 hours lectures by industry experts on contemporary topics.', sz: 10, v: 'top' }],
      [{ t: 'Mode of Evaluation and Assessment:\n' + 'The assessment and evaluation components may consist of unannounced open book examinations, quizzes, student portfolio generation and assessment, and any other innovative assessment practices followed by faculty, in addition to the Continuous Assessment Tests and Term End Examinations.', sz: 10, v: 'top' }],
      [{ t: 'Text Books:\n' + num(c.textBooks), sz: 10, v: 'top' }],
      [{ t: 'Reference books:\n' + num(c.refBooks), sz: 10, v: 'top' }],
    ], [FULL]);
    x += table([
      [{ t: 'Recommendation by the Board of Studies on', i: true, sz: 10 }, { t: c.bos || '', b: true, sz: 10 }],
      [{ t: 'Approval by Academic council on:', i: true, sz: 10 }, { t: c.approval || '', b: true, sz: 10 }],
      [{ t: 'Compiled by:', i: true, sz: 10 }, { t: c.compiledBy || '', b: true, sz: 10 }],
    ], [6000, FULL - 6000]);
    return x + pageBreak();
  }

  function legend() {
    return table([[{ t: '', shade: G }, 'Strong Correlation'], [{ t: '', shade: Y }, 'Moderate Correlation'], [{ t: '', shade: R }, 'Weak / No Correlation']], [700, 2600], { noBorder: true });
  }
  function mappings(c) {
    let x = title('CO-SO Mapping');
    const hs = ['CO \\ SO'].concat(Array.from({ length: 7 }, (_, i) => 'SO' + (i + 1)));
    const ws = Math.floor((FULL - 1100) / 7);
    let rows = [hs.map(h => ({ t: h, b: true, align: 'center', sz: 10, shade: HEAD }))];
    (c.cos || []).forEach(o => rows.push([{ t: o.id, b: true, align: 'center' }].concat((c.mapSO[o.id] || []).map(v => ({ t: v ? String(v) : '-', align: 'center', shade: LV[v] || R })))));
    x += table(rows, [1100].concat(Array(7).fill(ws)), { center: true }) + spacer() + legend() + spacer();
    x += title('CO-PO & PSO Mapping');
    const hp = ['CO \\ PO'].concat(Array.from({ length: 12 }, (_, i) => 'PO' + (i + 1)), ['PSO1', 'PSO2']);
    const wp = Math.floor((FULL - 900) / 14);
    rows = [hp.map(h => ({ t: h, b: true, align: 'center', sz: 8, shade: HEAD }))];
    (c.cos || []).forEach(o => rows.push([{ t: o.id, b: true, align: 'center', sz: 9 }].concat((c.mapPO[o.id] || []).map(v => ({ t: v ? String(v) : '', align: 'center', sz: 9, shade: LV[v] || undefined })))));
    x += table(rows, [900].concat(Array(14).fill(wp)), { center: true }) + spacer() + legend();
    x += para('Blank cell: no correlation.', { sz: 9, before: 60 });
    return x + pageBreak();
  }

  function coursePlan(c) {
    let x = title('Course Plan with pedagogy/Activity and Weightage of various assessment tools');
    const W = [650, 950, 800, 3650, 1000, 650, 2046];
    const rows = [['Sess. #', 'Date', 'CO', 'Topic', 'Pedagogy*', 'L / P / T', 'Weightage for various assessment tools'].map(h => ({ t: h, b: true, align: 'center', sz: 9, shade: HEAD }))];
    c.sessions.forEach(s => rows.push([{ t: String(s.no), align: 'center', sz: 9 }, { t: s.date, align: 'center', sz: 9 }, { t: s.coTag, align: 'center', sz: 9 }, { t: s.topic, sz: 9 }, { t: 'Cal+PPT', align: 'center', sz: 9 }, { t: s.lpt, align: 'center', sz: 9 }, { t: s.weight, sz: 9 }]));
    x += table(rows, W, { header: true });
    x += para('*Experiential and Participative learning through CALTech', { sz: 10, before: 120, after: 240 });
    x += title('Assessment Tools and their Weightages');
    const a = c.assessment;
    x += table([[{ t: 'Sl. No', b: true, align: 'center', shade: HEAD, sz: 10 }, { t: 'Assessment Tools', b: true, align: 'center', shade: HEAD, sz: 10 }, { t: 'Weightage %', b: true, align: 'center', shade: HEAD, sz: 10 }]]
      .concat(a.map((r, i) => [{ t: String(i + 1), align: 'center', sz: 10 }, { t: r.tool, sz: 10 }, { t: String(r.weight), b: true, align: 'center', sz: 10 }])), [900, 5200, 1700], { center: true });
    x += spacer() + title('Course Outcome Attainment');
    x += table([[{ t: 'COs', b: true, align: 'center', shade: HEAD }, { t: 'Target in %', b: true, align: 'center', shade: HEAD }, { t: 'Expected Level of Attainment in %', b: true, align: 'center', shade: HEAD }, { t: 'Actual Level of Attainment in %', b: true, align: 'center', shade: HEAD }]]
      .concat((c.cos || []).map(o => [{ t: o.id, align: 'center' }, { t: String(c.target), align: 'center' }, { t: String(c.expected), align: 'center' }, ''])), [1200, 1700, 2400, 2400], { center: true });
    return x;
  }

  function sessionTables(c) {
    let x = '';
    const k = c.course;
    const W = [2100, 3300, 4346];
    c.sessions.forEach((s, i) => {
      x += i === 0 ? title('Session Plan', { pb: true }) : '';
      const lines = (s.coTag ? ['[' + s.coTag + ']'] : []).concat((s.outcomes || []).filter(Boolean));
      const lab = t => ({ t, b: true, sz: 10.5 });
      const val = t => ({ t, sz: 10.5, span: 2, v: 'center' });
      x += table([
        [{ t: 'Session Number: ' + String(s.no).padStart(2, '0'), b: true, sz: 11 }, { t: 'Date: ' + (s.date || ''), b: true, sz: 10.5 }, { t: `Class no:   ${k.classNo || '________'}\nSlot: ${k.slot}`, b: true, sz: 10.5 }],
        [lab('Topic'), val(s.topic)],
        [lab('Venue'), val(s.venue)],
        [lab('Course Outcome'), val(lines.join('\n'))],
        [lab('Before Class Activity'), val(s.before)],
        [lab('Entry Ticket / Assessment'), val(s.entry)],
        [lab('Pedagogy adopted'), val(s.pedagogy)],
        [lab('Material'), val(s.material)],
        [lab('Exit Ticket / Assessment'), val(s.exit)],
        [lab('After Class Activity'), val(s.after)],
      ], W);
      x += i < c.sessions.length - 1 ? pageBreak() : '';
    });
    return x;
  }

  function questionPaper(c) {
    const k = c.course;
    let x = para('Annexure -1 (Model Question Paper)', { b: true, sz: 12, pb: true, after: 120 });
    x += table([[{ t: '', span: 1 }, { t: 'Reg. No.:', b: true, align: 'right' }, ''], [{ t: '' }, { t: 'Name     :', b: true, align: 'right' }, '']], [5000, 1500, 3246]);
    x += para('TERM END EXAMINATIONS (TEE) – ' + k.semester, { b: true, sz: 12, align: 'center', before: 160, after: 100 });
    x += table([
      [{ t: 'Programme', b: true }, ': ' + (k.programme || 'B.Tech.'), { t: 'Semester', b: true }, ': ' + k.semester],
      [{ t: 'Course Title', b: true }, ': ' + k.title, { t: 'Course Code', b: true }, ': ' + k.code],
      [{ t: 'Faculty Name', b: true }, ': ' + k.faculty, { t: 'Slot', b: true }, ': ' + k.slot],
      [{ t: 'Time', b: true }, ': ' + (c.qpTime || '1½ hours'), { t: 'Max. Marks', b: true }, ': ' + (c.qpMarks || 50)],
    ], [1500, 3373, 1500, 3373]);
    x += para('Answer ALL the Questions', { b: true, align: 'center', before: 120, after: 100 });
    const W = [600, 600, 4900, 700, 1250, 850, 846];
    const H = ['Q. No.', '', 'Question Description', 'Marks', 'Module No.', 'RBT Level', 'CO'].map(h => ({ t: h, b: true, align: 'center', sz: 9.5, shade: HEAD }));
    const rows = [H];
    const q = c.qp;
    rows.push([{ t: 'PART - A ( 30 Marks)', b: true, align: 'center', span: 7 }]);
    for (let i = 0; i < 6; i += 2) {
      [q[i], q[i + 1]].forEach((it, j) => {
        if (!it) return;
        rows.push([{ t: it.no, align: 'center' }, { t: it.part, align: 'center' }, { t: it.text, sz: 10 }, { t: String(it.marks), align: 'center' }, { t: it.module, align: 'center', sz: 9.5 }, { t: it.kl, align: 'center' }, { t: it.co, align: 'center' }]);
        if (j === 0) rows.push([{ t: it.no, align: 'center' }, { t: 'OR', b: true, align: 'center', span: 6 }]);
      });
    }
    rows.push([{ t: 'PART - B (20 Marks)', b: true, align: 'center', span: 7 }]);
    q.slice(6).forEach(it => rows.push([{ t: it.no, align: 'center' }, { t: it.text, sz: 10, span: 2 }, { t: String(it.marks), align: 'center' }, { t: it.module, align: 'center', sz: 9.5 }, { t: it.kl, align: 'center' }, { t: it.co, align: 'center' }]));
    x += table(rows, W, { header: true });
    x += para('⇦⇦⇦', { align: 'center', before: 120 });
    return x;
  }

  // ---------------------------------------------------------- package ----
  function b64ToBytes(b64) { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; }

  async function generate(cfg) {
    const enc = new TextEncoder();
    const logo = cfg.course.useLogo !== false && root.CCM_LOGO ? b64ToBytes(root.CCM_LOGO) : null;
    const body = cover(cfg, !!logo) + contents() + outcomes(cfg) + syllabus(cfg) + mappings(cfg) + coursePlan(cfg) + pageBreak() + sessionTables(cfg) + questionPaper(cfg);
    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
    const sect = '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="500" w:footer="500" w:gutter="0"/></w:sectPr>';
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${sect}</w:body></w:document>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-IN"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="60"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;
    const fld = t => `<w:r><w:fldChar w:fldCharType="${t}"/></w:r>`;
    const footer = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NS}><w:p><w:pPr><w:jc w:val="right"/></w:pPr>${run(cfg.course.code + ' – ' + cfg.course.title + '   |   Page ', { sz: 9, color: '595959' })}${fld('begin')}<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>${fld('separate')}${run('1', { sz: 9, color: '595959' })}${fld('end')}</w:p></w:ftr>`;
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>${logo ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>' : ''}</Relationships>`;
    const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`;
    const root_rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
    const files = [
      { name: '[Content_Types].xml', data: enc.encode(ct) },
      { name: '_rels/.rels', data: enc.encode(root_rels) },
      { name: 'word/document.xml', data: enc.encode(doc) },
      { name: 'word/styles.xml', data: enc.encode(styles) },
      { name: 'word/footer1.xml', data: enc.encode(footer) },
      { name: 'word/_rels/document.xml.rels', data: enc.encode(rels) },
    ];
    if (logo) files.push({ name: 'word/media/logo.png', data: logo });
    return zip(files);
  }

  const api = { generate, zip };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.CCMDoc = api;
})(typeof window !== 'undefined' ? window : globalThis);
