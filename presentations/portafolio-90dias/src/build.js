const pptxgen = require('pptxgenjs');
const fs = require('fs');

const INK    = 'E8EFF9';
const MUTED  = '9FB4D2';
const CARD   = '1B3157';
const BORDER = '34507C';
const DARK   = '020814';
const AMBER  = 'F3B65F';
const GREEN  = '7CE7AE';
const CYAN   = '39C6F5';
const CORAL  = 'FF8C7C';
const FONT   = 'Calibri';

const W = 13.333, H = 7.5;
const M = 0.55;                 // margen lateral
const CW = W - 2 * M;           // ancho util = 12.233

const pres = new pptxgen();
pres.defineLayout({ name: 'P90', width: W, height: H });
pres.layout = 'P90';
pres.author = 'C-OTC Order Management';
pres.title = '90 days plan - Portafolio';

const BG = 'image/png;base64,' + fs.readFileSync(__dirname + '/bg.png').toString('base64');

const FOOTER = 'P&G · C-OTC Order Management — San Jose';

function bg(slide) {
  slide.addImage({ data: BG, x: 0, y: 0, w: W, h: H });
}

function footer(slide, right) {
  slide.addText(FOOTER, {
    x: M, y: 7.06, w: 6.5, h: 0.28, fontFace: FONT, fontSize: 12, color: MUTED, margin: 0, valign: 'middle'
  });
  if (right) {
    slide.addText(right, {
      x: W - M - 6.0, y: 7.06, w: 6.0, h: 0.28, fontFace: FONT, fontSize: 12, color: MUTED,
      align: 'right', margin: 0, valign: 'middle'
    });
  }
}

// tarjeta contenedora
function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.09,
    fill: { color: fill || CARD },
    line: { color: BORDER, width: 1 }
  });
}

// etiqueta de seccion dentro de una tarjeta
function cardLabel(slide, x, y, w, text, color) {
  slide.addText(text, {
    x, y, w, h: 0.3, fontFace: FONT, fontSize: 14, bold: true,
    color: color || MUTED, charSpacing: 1.2, margin: 0, valign: 'middle'
  });
}

// pastilla de estado: relleno solido + texto casi negro = maxima legibilidad de lejos
function pill(slide, x, y, w, h, text, fill, size) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.14, fill: { color: fill }, line: { color: fill, width: 1 }
  });
  slide.addText(text, {
    x, y, w, h, fontFace: FONT, fontSize: size || 16, bold: true, color: DARK,
    align: 'center', valign: 'middle', margin: 0
  });
}

function dot(slide, x, y, color, d) {
  const s = d || 0.17;
  slide.addShape(pres.ShapeType.ellipse, {
    x, y, w: s, h: s, fill: { color }, line: { color, width: 1 }
  });
}

// ---------------------------------------------------------------- datos
const PROJECTS = [
  {
    n: 1,
    name: 'ZE Category Emails & Display Auto-Release',
    status: 'In Progress', statusColor: AMBER,
    benefits: [
      'Item-category tag auto-routes emails',
      'Auto-releases displays when P&G < customer'
    ],
    otsr: ['Order Readiness', 'Total Order Flow'],
    owners: ['OMA (Mariela)', 'K. Aguilar · B. Miller'],
    partners: ['PIM (Item Cat.)'],
    help: null,
    doneBy: ['Part I ✔  ·  AMJ 2026', 'Part II  ·  TBD'],
    notes: 'Parte I ya entregada. La parte II sigue sin fecha: pedir definicion en la revision.'
  },
  {
    n: 2,
    name: 'AI Customer Material # Inclusion',
    status: 'In Progress', statusColor: AMBER,
    benefits: [
      'RPA updates Customer Material # on existing SAP lines',
      'Reduces returns risk & SAP rework'
    ],
    otsr: ['Total Order Flow', 'Customer Service on Time'],
    owners: ['Sofia Arroyo'],
    partners: ['RPA / Blue Prism'],
    help: null,
    doneBy: ['Jul 2026'],
    notes: 'Cierre previsto Jul 2026. Sin bloqueos abiertos.'
  },
  {
    n: 3,
    name: 'Smart Inbox',
    status: 'In Progress', statusColor: AMBER,
    benefits: [
      'AI classifies & routes automated emails (ZE, ZQ, YU)',
      '~2 hrs/week saved'
    ],
    otsr: ['Order Readiness', 'Productivity'],
    owners: ['M. Campos', 'D. Espinoza'],
    partners: ['OSSgenAI / IT'],
    help: null,
    doneBy: ['Jul 2026'],
    notes: 'Beneficio principal: ~2 horas por semana liberadas.'
  },
  {
    n: 4,
    name: 'EDI Orders Validation',
    status: 'Discovery', statusColor: CYAN,
    benefits: [
      'AI + ActOM validate the EDI transmission',
      'Replaces manual PDF review · avoids error'
    ],
    otsr: ['Order Readiness', 'Order Quality'],
    owners: ['Cristhofer Marin'],
    partners: ['DOOM / Azure'],
    help: 'Help needed',
    doneBy: ['Jul 2026'],
    notes: 'Unico proyecto todavia en Discovery y con ayuda pendiente de DOOM / Azure.'
  },
  {
    n: 5,
    name: 'DSD KNIME Lead-Time Exception (LTE)',
    status: 'In Progress', statusColor: AMBER,
    benefits: [
      'Unified KNIME flags LTE orders & auto-responds to RPA',
      '100% accuracy · no wrong-RDD pushes'
    ],
    otsr: ['Order Readiness', 'Total Order Flow'],
    owners: ['Cristhofer Marin'],
    partners: ['COTC'],
    help: 'Help needed to copy KNIME',
    doneBy: ['Aug 2026'],
    notes: 'Se necesita apoyo de COTC para copiar el flujo KNIME.'
  }
];

const ALSO = ['AWG Order Consolidation', 'Automatic Cancellation Request', 'YV Block Autoclean'];
const DONE = ['Block 11 Autoclean', 'RDD Autopush', 'VMI Appointment Automation', 'AWG Ship-With Automation'];

// ============================================================ SLIDE 1
{
  const s = pres.addSlide();
  bg(s);

  s.addText(
    [
      { text: '90 days plan · ', options: { color: INK } },
      { text: 'PORTAFOLIO', options: { color: CYAN } }
    ],
    { x: M, y: 0.34, w: 9.6, h: 0.78, fontFace: FONT, fontSize: 40, bold: true, margin: 0, valign: 'middle' }
  );

  s.addShape(pres.ShapeType.roundRect, {
    x: W - M - 1.9, y: 0.42, w: 1.9, h: 0.6, rectRadius: 0.14,
    fill: { color: '2A5896' }, line: { color: '3E72AE', width: 1 }
  });
  s.addText('Jul 2026', {
    x: W - M - 1.9, y: 0.42, w: 1.9, h: 0.6, fontFace: FONT, fontSize: 20, bold: true,
    color: INK, align: 'center', valign: 'middle', margin: 0
  });

  s.addText(
    [
      { text: 'IWS ', options: { bold: true, color: CYAN } },
      { text: 'methodology implementation to drive improvement and drive value.', options: { color: MUTED, italic: true } }
    ],
    { x: M, y: 1.14, w: 11.0, h: 0.34, fontFace: FONT, fontSize: 18, margin: 0, valign: 'middle' }
  );

  // ---- fila de indicadores
  const KPI = [
    { v: '12', l: 'Automations', c: INK },
    { v: '7',  l: 'In Progress', c: AMBER },
    { v: '1',  l: 'Discovery',   c: CYAN },
    { v: '4',  l: 'Completed',   c: GREEN }
  ];
  const kw = (CW - 3 * 0.3) / 4;
  KPI.forEach((k, i) => {
    const x = M + i * (kw + 0.3);
    card(s, x, 1.62, kw, 1.35);
    s.addText(k.v, {
      x, y: 1.68, w: kw, h: 0.82, fontFace: FONT, fontSize: 48, bold: true, color: k.c,
      align: 'center', valign: 'middle', margin: 0
    });
    s.addText(k.l, {
      x, y: 2.48, w: kw, h: 0.4, fontFace: FONT, fontSize: 17, color: MUTED,
      align: 'center', valign: 'middle', margin: 0, charSpacing: 0.6
    });
  });

  s.addText('TOP 5 IN PROGRESS', {
    x: M, y: 3.16, w: 5.0, h: 0.36, fontFace: FONT, fontSize: 19, bold: true, color: AMBER,
    charSpacing: 1.2, margin: 0, valign: 'middle'
  });
  s.addText('Detail on the following slides', {
    x: W - M - 6.0, y: 3.16, w: 6.0, h: 0.36, fontFace: FONT, fontSize: 15, color: MUTED,
    align: 'right', margin: 0, valign: 'middle'
  });

  PROJECTS.forEach((p, i) => {
    const y = 3.58 + i * 0.66;
    card(s, M, y, CW, 0.56);

    slideBadge(s, M + 0.16, y + 0.1, p.n);

    s.addText(p.name, {
      x: M + 0.82, y, w: 6.95, h: 0.56, fontFace: FONT, fontSize: 19, bold: true, color: INK,
      valign: 'middle', margin: 0
    });

    pill(s, M + 7.9, y + 0.09, 1.72, 0.38, p.status, p.statusColor, 14);

    if (p.help) {
      s.addText('⚑ Help', {
        x: M + 9.72, y, w: 0.95, h: 0.56, fontFace: FONT, fontSize: 14, bold: true, color: CORAL,
        valign: 'middle', margin: 0
      });
    }

    s.addText(p.doneBy[0].replace('Part I ✔  ·  ', ''), {
      x: M + 10.73, y, w: 1.3, h: 0.56, fontFace: FONT, fontSize: 15, bold: true, color: INK,
      align: 'right', valign: 'middle', margin: 0
    });
  });

  footer(s, 'Full portfolio table in the appendix');
  s.addNotes('Vista general: 12 automatizaciones, 4 completadas, 7 en progreso, 1 en discovery. El detalle de cada uno de los 5 principales va en las diapositivas siguientes.');
}

function slideBadge(s, x, y, n) {
  s.addShape(pres.ShapeType.ellipse, {
    x, y, w: 0.36, h: 0.36, fill: { color: CYAN }, line: { color: CYAN, width: 1 }
  });
  s.addText(String(n), {
    x, y, w: 0.36, h: 0.36, fontFace: FONT, fontSize: 15, bold: true, color: DARK,
    align: 'center', valign: 'middle', margin: 0
  });
}

// ==================================================== SLIDES 2-6 (detalle)
PROJECTS.forEach((p) => {
  const s = pres.addSlide();
  bg(s);

  s.addText(`TOP 5 IN PROGRESS  ·  ${p.n} OF 5`, {
    x: M, y: 0.3, w: 6.0, h: 0.3, fontFace: FONT, fontSize: 14, bold: true, color: MUTED,
    charSpacing: 1.4, margin: 0, valign: 'middle'
  });

  s.addText(p.name, {
    x: M, y: 0.64, w: 9.3, h: 1.0, fontFace: FONT, fontSize: 34, bold: true, color: INK,
    margin: 0, valign: 'middle'
  });

  pill(s, W - M - 2.5, 0.72, 2.5, 0.62, p.status, p.statusColor, 20);

  // --- izquierda: beneficios
  card(s, M, 1.9, 7.55, 2.6);
  cardLabel(s, M + 0.28, 2.02, 7.0, 'BENEFITS');
  s.addText(
    p.benefits.map((b, i) => ({
      text: b,
      options: { bullet: { indent: 20 }, breakLine: i < p.benefits.length - 1, paraSpaceAfter: 16 }
    })),
    { x: M + 0.28, y: 2.36, w: 6.99, h: 1.98, fontFace: FONT, fontSize: 23, color: INK, margin: 0, valign: 'middle' }
  );

  // --- izquierda: impactos OTSR
  card(s, M, 4.8, 7.55, 2.0);
  cardLabel(s, M + 0.28, 4.92, 7.0, 'OTSR IMPACTS');
  p.otsr.forEach((o, i) => {
    const x = M + 0.28 + i * 3.6;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 5.5, w: 3.4, h: 0.92, rectRadius: 0.1,
      fill: { color: '2A5896' }, line: { color: '3E72AE', width: 1 }
    });
    s.addText(o, {
      x: x + 0.18, y: 5.5, w: 3.04, h: 0.92, fontFace: FONT, fontSize: 19, bold: true, color: INK,
      valign: 'middle', margin: 0
    });
  });

  // --- derecha: owners
  const RX = M + 7.85, RW = CW - 7.85;
  card(s, RX, 1.9, RW, 1.75);
  cardLabel(s, RX + 0.28, 2.02, RW - 0.56, 'OWNERS');
  s.addText(
    p.owners.map((o, i) => ({ text: o, options: { breakLine: i < p.owners.length - 1, paraSpaceAfter: 8 } })),
    { x: RX + 0.28, y: 2.3, w: RW - 0.56, h: 1.25, fontFace: FONT, fontSize: 20, color: INK, margin: 0, valign: 'middle' }
  );

  // --- derecha: partners
  card(s, RX, 3.95, RW, 1.35);
  cardLabel(s, RX + 0.28, 4.05, RW - 0.56, 'PARTNERS');
  s.addText(p.partners.join(' · '), {
    x: RX + 0.28, y: 4.38, w: RW - 0.56, h: p.help ? 0.42 : 0.82,
    fontFace: FONT, fontSize: 20, color: INK, margin: 0, valign: 'middle'
  });
  if (p.help) {
    s.addText('⚑ ' + p.help, {
      x: RX + 0.28, y: 4.8, w: RW - 0.56, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: CORAL,
      margin: 0, valign: 'middle'
    });
  }

  // --- derecha: fecha
  card(s, RX, 5.6, RW, 1.2);
  cardLabel(s, RX + 0.28, 5.68, RW - 0.56, 'DONE BY');
  if (p.doneBy.length === 1) {
    s.addText(p.doneBy[0], {
      x: RX + 0.28, y: 5.96, w: RW - 0.56, h: 0.72, fontFace: FONT, fontSize: 26, bold: true, color: INK,
      margin: 0, valign: 'middle'
    });
  } else {
    s.addText(
      p.doneBy.map((d, i) => ({ text: d, options: { breakLine: i < p.doneBy.length - 1, paraSpaceAfter: 2 } })),
      { x: RX + 0.28, y: 5.96, w: RW - 0.56, h: 0.76, fontFace: FONT, fontSize: 18, bold: true, color: INK, margin: 0, valign: 'middle' }
    );
  }

  footer(s, `${p.n} of 5`);
  s.addNotes(p.notes);
});

// ============================================================ SLIDE 7
{
  const s = pres.addSlide();
  bg(s);

  s.addText('The other 7 automations', {
    x: M, y: 0.34, w: 9.6, h: 0.78, fontFace: FONT, fontSize: 40, bold: true, color: INK,
    margin: 0, valign: 'middle'
  });
  s.addText('Outside the Top 5 · same portfolio', {
    x: W - M - 5.5, y: 0.42, w: 5.5, h: 0.6, fontFace: FONT, fontSize: 16, color: MUTED,
    align: 'right', valign: 'middle', margin: 0
  });

  // Also in progress
  dot(s, M, 1.63, AMBER, 0.2);
  s.addText('ALSO IN PROGRESS (3)', {
    x: M + 0.32, y: 1.52, w: 6.0, h: 0.4, fontFace: FONT, fontSize: 19, bold: true, color: AMBER,
    charSpacing: 1.2, margin: 0, valign: 'middle'
  });
  const aw = (CW - 2 * 0.3) / 3;
  ALSO.forEach((t, i) => {
    const x = M + i * (aw + 0.3);
    card(s, x, 2.05, aw, 1.7);
    dot(s, x + 0.3, 2.33, AMBER);
    s.addText(t, {
      x: x + 0.3, y: 2.6, w: aw - 0.6, h: 0.95, fontFace: FONT, fontSize: 21, bold: true, color: INK,
      margin: 0, valign: 'top'
    });
  });

  // Completed
  dot(s, M, 4.33, GREEN, 0.2);
  s.addText('COMPLETED (4)', {
    x: M + 0.32, y: 4.22, w: 6.0, h: 0.4, fontFace: FONT, fontSize: 19, bold: true, color: GREEN,
    charSpacing: 1.2, margin: 0, valign: 'middle'
  });
  const dw = (CW - 3 * 0.3) / 4;
  DONE.forEach((t, i) => {
    const x = M + i * (dw + 0.3);
    card(s, x, 4.75, dw, 1.7);
    dot(s, x + 0.26, 5.03, GREEN);
    s.addText(t, {
      x: x + 0.26, y: 5.3, w: dw - 0.52, h: 0.95, fontFace: FONT, fontSize: 20, bold: true, color: INK,
      margin: 0, valign: 'top'
    });
  });

  footer(s, 'Full portfolio table in the appendix');
  s.addNotes('Las 7 automatizaciones restantes del portafolio: 3 en progreso fuera del Top 5 y 4 ya completadas.');
}

pres.writeFile({ fileName: __dirname + '/new.pptx' }).then(() => console.log('ok'));
