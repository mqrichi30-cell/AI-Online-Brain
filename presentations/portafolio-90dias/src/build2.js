const pptxgen = require('pptxgenjs');
const fs = require('fs');

const INK    = 'E8EFF9';
const MUTED  = '9FB4D2';
const CARD   = '1B3157';
const HERO   = '1F3A66';
const BORDER = '34507C';
const DARK   = '020814';
const AMBER  = 'F3B65F';
const GREEN  = '7CE7AE';
const CYAN   = '39C6F5';
const CORAL  = 'FF8C7C';
const FONT   = 'Calibri';

const W = 13.333, H = 7.5;
const M = 0.55;
const CW = W - 2 * M;          // 12.233

const pres = new pptxgen();
pres.defineLayout({ name: 'P90', width: W, height: H });
pres.layout = 'P90';
pres.author = 'C-OTC Order Management';
pres.title = '90 days plan - Portafolio';

const BG = 'image/png;base64,' + fs.readFileSync(__dirname + '/bg.png').toString('base64');
const FOOTER = 'P&G · C-OTC Order Management — San Jose';

function bg(s) {
  s.addImage({ data: BG, x: 0, y: 0, w: W, h: H });
}

function footer(s, right, rightColor) {
  s.addText(FOOTER, {
    x: M, y: 7.08, w: 5.2, h: 0.28, fontFace: FONT, fontSize: 12, color: MUTED, margin: 0, valign: 'middle'
  });
  if (right) {
    s.addText(right, {
      x: W - M - 7.0, y: 7.08, w: 7.0, h: 0.28, fontFace: FONT, fontSize: 13, color: rightColor || MUTED,
      align: 'right', margin: 0, valign: 'middle'
    });
  }
}

function card(s, x, y, w, h, opt) {
  const o = opt || {};
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.09,
    fill: { color: o.fill || CARD },
    line: { color: o.line || BORDER, width: o.lineWidth || 1 }
  });
}

function dot(s, x, y, color, d) {
  const sz = d || 0.17;
  s.addShape(pres.ShapeType.ellipse, { x, y, w: sz, h: sz, fill: { color }, line: { color, width: 1 } });
}

function header(s, title, meta) {
  s.addText(title, {
    x: M, y: 0.24, w: 6.4, h: 0.62, fontFace: FONT, fontSize: 34, bold: true, color: INK,
    margin: 0, valign: 'middle'
  });
  s.addText(meta, {
    x: W - M - 6.0, y: 0.24, w: 6.0, h: 0.62, fontFace: FONT, fontSize: 17, color: MUTED,
    align: 'right', margin: 0, valign: 'middle'
  });
}

// ---------------------------------------------------------------- datos
// filas 1-2: en curso · fila 3: discovery + on hold
const FLIGHT = [
  { name: 'AI Customer Material # Inclusion', owner: 'Sofia Arroyo',            color: AMBER },
  { name: 'Smart Inbox',                      owner: 'M. Campos · D. Espinoza', color: AMBER },
  { name: 'DSD KNIME Lead-Time Exception (LTE)', owner: 'Cristhofer Marin',     color: AMBER, tag: '⚑ COTC', tagColor: CORAL },
  { name: 'Automatic Cancellation Request',   owner: 'Cristhofer Marin',        color: AMBER, tag: 'in parallel', tagColor: AMBER },
  { name: 'AWG Order Consolidation',          owner: 'Cristhofer Marin',        color: AMBER, tag: 'in parallel', tagColor: AMBER },
  { name: 'YV Block Autoclean',               owner: '',                        color: AMBER }
];

const ATTENTION = [
  {
    name: 'EDI Orders Validation', owner: 'Cristhofer Marin', color: CYAN,
    tag: '⚑ DOOM / Azure', tagColor: CORAL, note: 'Discovery'
  },
  {
    name: 'PGP Ion', owner: 'Joshua (PGP Team)', color: CORAL, onHold: true,
    tag: '⚑ ON HOLD', tagColor: CORAL, note: 'High desk volume'
  },
  {
    name: 'ZE Display Auto-Release', owner: 'OMA (Mariela)', color: CORAL, onHold: true,
    tag: '⚑ ON HOLD', tagColor: CORAL, note: 'Help needed from COTS'
  }
];

const COMPLETED = [
  { name: 'ZE Category Emails', detail: 'Cristhofer Marin · item-category tag auto-routes emails' },
  { name: 'Block 11 Autoclean' },
  { name: 'RDD Autopush' },
  { name: 'VMI Appointment Automation' },
  { name: 'AWG Ship-With Automation' }
];

const SWAT_CHIPS = ['2 weeks to scope', 'Pairs', 'Rotating weekly mentor', 'Kanban board', 'Weekly sync'];

// ============================================================ IN PROGRESS
{
  const s = pres.addSlide();
  bg(s);
  header(s, 'In Progress', '7 automations in flight · 2 on hold');

  // ---- banda destacada: SWAT Team PGP
  const hy = 0.92, hh = 0.90;
  card(s, M, hy, CW, hh, { fill: HERO, line: CYAN, lineWidth: 1.5 });

  s.addText('SWAT Team PGP', {
    x: M + 0.3, y: hy + 0.1, w: 2.95, h: 0.36, fontFace: FONT, fontSize: 22, bold: true, color: INK,
    margin: 0, valign: 'middle'
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M + 3.3, y: hy + 0.14, w: 0.72, h: 0.28, rectRadius: 0.11,
    fill: { color: CYAN }, line: { color: CYAN, width: 1 }
  });
  s.addText('NEW', {
    x: M + 3.3, y: hy + 0.14, w: 0.72, h: 0.28, fontFace: FONT, fontSize: 13, bold: true, color: DARK,
    align: 'center', valign: 'middle', margin: 0
  });
  s.addText('6 weeks embedded with another team — PGP goes first', {
    x: M + 4.25, y: hy + 0.1, w: 7.4, h: 0.36, fontFace: FONT, fontSize: 16, color: INK,
    margin: 0, valign: 'middle'
  });

  let cx = M + 0.3;
  SWAT_CHIPS.forEach((t) => {
    const w = 0.32 + t.length * 0.137;
    s.addShape(pres.ShapeType.roundRect, {
      x: cx, y: hy + 0.5, w, h: 0.3, rectRadius: 0.1,
      fill: { color: '2A5896' }, line: { color: '3E72AE', width: 1 }
    });
    s.addText(t, {
      x: cx, y: hy + 0.5, w, h: 0.3, fontFace: FONT, fontSize: 14, bold: true, color: INK,
      align: 'center', valign: 'middle', margin: 0
    });
    cx += w + 0.14;
  });

  // ---- rejilla: 2 filas en curso + 1 fila que necesita atencion
  const gw = (CW - 2 * 0.28) / 3;

  function projectCard(x, y, w, h, p) {
    card(s, x, y, w, h, p.onHold ? { line: CORAL } : {});
    dot(s, x + 0.26, y + 0.18, p.color);
    if (p.tag) {
      s.addText(p.tag, {
        x: x + 0.5, y: y + 0.14, w: w - 0.76, h: 0.26, fontFace: FONT, fontSize: 14, bold: true,
        color: p.tagColor, align: 'right', margin: 0, valign: 'middle'
      });
    }
    s.addText(p.name, {
      x: x + 0.26, y: y + 0.44, w: w - 0.52, h: 0.68, fontFace: FONT, fontSize: 18, bold: true,
      color: INK, margin: 0, valign: 'top'
    });
    let my = y + 1.14;
    if (p.note) {
      s.addText(p.note, {
        x: x + 0.26, y: my, w: w - 0.52, h: 0.28, fontFace: FONT, fontSize: 15, bold: true,
        color: p.onHold ? CORAL : CYAN, margin: 0, valign: 'middle'
      });
      my += 0.30;
    }
    if (p.owner) {
      s.addText(p.owner, {
        x: x + 0.26, y: my, w: w - 0.52, h: 0.28, fontFace: FONT, fontSize: 15, color: MUTED,
        margin: 0, valign: 'middle'
      });
    }
  }

  FLIGHT.forEach((p, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    projectCard(M + col * (gw + 0.28), 1.96 + row * 1.60, gw, 1.46, p);
  });

  ATTENTION.forEach((p, i) => {
    projectCard(M + i * (gw + 0.28), 5.16, gw, 1.76, p);
  });

  s.addText(
    [
      { text: '●  ', options: { color: AMBER } },
      { text: 'In Progress      ', options: { color: MUTED } },
      { text: '●  ', options: { color: CYAN } },
      { text: 'Discovery      ', options: { color: MUTED } },
      { text: '●  ', options: { color: CORAL } },
      { text: 'On Hold · follow-up end of Aug', options: { color: MUTED } }
    ],
    {
      x: W - M - 7.6, y: 7.08, w: 7.6, h: 0.28, fontFace: FONT, fontSize: 13,
      align: 'right', margin: 0, valign: 'middle'
    }
  );
  footer(s);
  s.addNotes(
    'SWAT Team PGP — estructura de 6 semanas. Los que ya automatizamos nos integramos con otro equipo: ' +
    'las primeras 2 semanas encontramos que proyecto trabajar y les ensenamos como lo hacemos, ' +
    'despues lo desarrollamos juntos y pasamos al siguiente equipo. PGP es nuestro primer equipo.\n\n' +
    'Mecanica: cada integrante del SWAT tiene una dupla fija con la que trabaja todo el periodo. ' +
    'El equipo tiene un mentor semanal que da feedback desde fuera; ese rol rota semanalmente entre ' +
    'los propios integrantes del SWAT. Las tareas se trackean en un kanban board para poder ayudarnos ' +
    'entre nosotros, y hay una reunion semanal del equipo para contarle a los demas como vamos.\n\n' +
    'Automatic Cancellation Request y AWG Order Consolidation se estan trabajando en paralelo, ' +
    'ambos por Cristhofer.\n\n' +
    'On hold: PGP Ion esta delegado a Joshua pero frenado por el alto volumen de su desk — esperamos ' +
    'resolverlo con el SWAT team, ya que Joshua es del PGP Team; no entra en los 5 principales. ' +
    'ZE Display Auto-Release (OMA / Mariela, con K. Aguilar y B. Miller) necesita apoyo de COTS. ' +
    'Ambos dependen de otros equipos, asi que se hara follow-up a finales de agosto a solicitud ' +
    'de ambos equipos.'
  );
}

// ============================================================ COMPLETED
{
  const s = pres.addSlide();
  bg(s);
  header(s, 'Completed', '5 automations live in the desk');

  s.addText('Delivered, running and already saving time', {
    x: M, y: 0.98, w: 9.0, h: 0.34, fontFace: FONT, fontSize: 18, color: MUTED, margin: 0, valign: 'middle'
  });

  // lista de ancho completo: nombres grandes, sin espacio muerto
  COMPLETED.forEach((item, i) => {
    const y = 1.75 + i * 1.03;
    card(s, M, y, CW, 0.88);
    dot(s, M + 0.32, y + 0.34, GREEN, 0.2);
    s.addText(item.name, {
      x: M + 0.7, y, w: 7.2, h: 0.88, fontFace: FONT, fontSize: 24, bold: true, color: INK,
      margin: 0, valign: 'middle'
    });
    if (item.detail) {
      s.addText(item.detail, {
        x: M + 8.0, y, w: CW - 8.3, h: 0.88, fontFace: FONT, fontSize: 16, color: MUTED,
        align: 'right', margin: 0, valign: 'middle'
      });
    }
  });

  footer(s, 'Jul 2026');
  s.addNotes(
    'Cinco automatizaciones ya entregadas y corriendo. ZE Category Emails es el proyecto de Cristhofer: ' +
    'el tag de item-category rutea los correos automaticamente. La parte de Display Auto-Release, ' +
    'a nombre de Mariela, sigue en on hold y aparece en la lamina de In Progress.'
  );
}

pres.writeFile({ fileName: __dirname + '/Portafolio_90dias_2laminas.pptx' }).then(() => console.log('ok'));
