/* ─────────────────────────────────────────────────────────────
   generate-docx.js  –  Netlify serverless function
   Generates a Word (.docx) SOW for TransPerfect Legal.
   Floor rates & internal pricing notes live ONLY here.
   ───────────────────────────────────────────────────────────── */
'use strict';

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, HeadingLevel,
  ShadingType, UnderlineType, PageBreak, Tab, VerticalAlign,
  ImageRun, Header, Footer, PageNumber, NumberFormat,
  convertInchesToTwip, convertMillimetersToTwip,
} = require('docx');

/* ── Server-side floor rates (never sent to the client) ─────── */
const FLOOR_RATES = {
  h7:  20,   // Brainspace
  g1:  0.45, // Cicero
  g2:  0.20, // eDiscoveryAI
  g3:  0.04, // ECI Core
  g4:  0.04, // ECI + Case Elements
  g5:  0.04, // ECI + Casebot
  g6:  450,  // GenAI Consulting
};

/* ── Colour palette ─────────────────────────────────────────── */
const NAVY  = '001F5F';
const TEAL  = '139DD8';
const LTBLUE = 'DEEAF6';
const WHITE  = 'FFFFFF';
const GREY   = 'F2F2F2';

/* ── Helpers ─────────────────────────────────────────────────── */
const pt = n => n * 2;          // half-points → pts for docx fontSize
const twip = convertInchesToTwip;
const mm = convertMillimetersToTwip;

function noBorder() {
  return { top: noBorderSide(), bottom: noBorderSide(), left: noBorderSide(), right: noBorderSide(), insideVertical: noBorderSide(), insideHorizontal: noBorderSide() };
}
function noBorderSide() { return { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }; }

function cellBorder(color = 'C5D5E8') {
  const side = { style: BorderStyle.SINGLE, size: 4, color };
  return { top: side, bottom: side, left: side, right: side };
}

function shading(fill) { return { type: ShadingType.CLEAR, color: 'auto', fill }; }

function boldRun(text, opts = {}) {
  return new TextRun({ text, bold: true, ...opts });
}

function para(children, opts = {}) {
  if (typeof children === 'string') children = [new TextRun({ text: children })];
  return new Paragraph({ children, ...opts });
}

function heading(text, color = NAVY, fontSize = 14, bold = true, spaceBefore = 200, spaceAfter = 100) {
  return new Paragraph({
    children: [new TextRun({ text, bold, color, size: pt(fontSize), font: 'Calibri' })],
    spacing: { before: spaceBefore, after: spaceAfter },
  });
}

function sectionTitle(text) {
  return new Paragraph({
    children: [new TextRun({ text: text.toUpperCase(), bold: true, color: NAVY, size: pt(13), font: 'Calibri', characterSpacing: 80 })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 160 },
  });
}

function phaseHeading(label, title) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${label} | `, color: NAVY, size: pt(12), font: 'Calibri', bold: false }),
      new TextRun({ text: title.toUpperCase(), bold: true, color: TEAL, size: pt(12), font: 'Calibri' }),
    ],
    spacing: { before: 240, after: 80 },
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    children: [new TextRun({ text, size: pt(10), font: 'Calibri' })],
    bullet: { level },
    spacing: { before: 40, after: 40 },
  });
}

function bodyPara(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: pt(10), font: 'Calibri' })],
    spacing: { before: 80, after: 80 },
    alignment: AlignmentType.JUSTIFIED,
    ...opts,
  });
}

function hrRule() {
  return new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY } },
    spacing: { before: 160, after: 160 },
  });
}

function pageBreakPara() {
  return new Paragraph({ children: [new TextRun({ break: 1 })], pageBreakBefore: true });
}

/* ── Validate rates against floors ──────────────────────────── */
function validatedRate(rowId, clientRate) {
  const floor = FLOOR_RATES[rowId];
  if (floor === undefined) return clientRate;
  return Math.max(parseFloat(clientRate) || 0, floor);
}

/* ── Rate card full list (for rate sheet output) ─────────────── */
const RATE_CARD = [
  {
    title: 'Forensic Data Collection',
    groups: [
      { label: 'Hourly Rates – Business Hours', rows: [
        { desc: 'Standard Forensic Acquisitions and Services', unit: 'Hour', id: 'f1' },
        { desc: 'Forensic Consulting & Analysis', unit: 'Hour', id: 'f2' },
        { desc: 'Senior Forensic Consulting & Analysis', unit: 'Hour', id: 'f3' },
      ]},
      { label: 'Hourly Rates – After Hours / Weekends', rows: [
        { desc: 'Standard Forensic Acquisitions', unit: 'Hour', id: 'f4' },
        { desc: 'Forensic Consulting & Analysis', unit: 'Hour', id: 'f5' },
        { desc: 'Senior Forensic Consulting & Analysis', unit: 'Hour', id: 'f6' },
      ]},
      { label: 'Technology Costs', rows: [
        { desc: 'Collection Media – Hard Drive 1TB', unit: 'Per 2 Drive Set', id: 'f7' },
        { desc: 'Remote Collection Kit – Domestic Rental', unit: 'Each', id: 'f8' },
        { desc: 'Remote Collection Kit – International Rental', unit: 'Each', id: 'f9' },
      ]},
    ],
  },
  {
    title: 'Data Processing & Early Case Assessment',
    groups: [
      { label: 'ECA Digital Reef – Ingestion', rows: [
        { desc: 'ECA Digital Reef – Ingestion', unit: 'GB', id: 'p1' },
      ]},
      { label: 'Data Export for Native Review', rows: [
        { desc: 'Data Export for Native Review', unit: 'GB', id: 'p2' },
      ]},
    ],
  },
  {
    title: 'Data Hosting, Analytics & Review Support',
    groups: [
      { label: 'Active Hosting – Relativity Server', rows: [
        { desc: 'Active Hosting – Relativity Server', unit: 'GB/Month', id: 'h1' },
        { desc: 'Monthly User Fee', unit: 'Each/Month', id: 'h2' },
        { desc: 'ECA Hosting (Digital Reef)', unit: 'GB/Month', id: 'h3' },
        { desc: 'Archive Storage', unit: 'GB/Month', id: 'h4' },
      ]},
      { label: 'Active Hosting – Reef Review', rows: [
        { desc: 'Active Hosting – Reef Review', unit: 'GB/Month', id: 'h5' },
        { desc: 'Archive Storage (Reef Review)', unit: 'GB/Month', id: 'h6' },
      ]},
      { label: 'Analytics, TAR and AI', rows: [
        { desc: 'Brainspace and Relativity', unit: 'GB', id: 'h7' },
        { desc: 'Blackout – Setup', unit: 'Setup', id: 'h8' },
        { desc: 'Blackout – Per Document', unit: 'Per Doc', id: 'h9' },
      ]},
      { label: 'Case Management & Support', rows: [
        { desc: 'Project Management', unit: 'Hour', id: 'h10' },
        { desc: 'Consulting', unit: 'Hour', id: 'h11' },
        { desc: 'Senior Consulting', unit: 'Hour', id: 'h12' },
      ]},
      { label: 'Productions', rows: [
        { desc: 'Imaging for Productions', unit: 'Page', id: 'h13' },
      ]},
      { label: 'Data Disposition', rows: [
        { desc: 'Database Disposal – Under 1TB', unit: 'Per Database', id: 'h14' },
        { desc: 'Database Disposal – Over 1TB', unit: 'Per GB', id: 'h15' },
      ]},
    ],
  },
  {
    title: 'GenAI',
    groups: [
      { label: '', rows: [
        { desc: 'GenAI (Cicero)', unit: 'Per Document', id: 'g1' },
        { desc: 'GenAI (eDiscoveryAI)', unit: 'Per Document', id: 'g2' },
        { desc: 'GenAI (ECI) – Core', unit: 'Per Document', id: 'g3' },
        { desc: 'GenAI (ECI) – + Case Elements', unit: 'Per Document', id: 'g4' },
        { desc: 'GenAI (ECI) – + Casebot Setup', unit: 'Per Document', id: 'g5', rateNote: '+ Quarterly hosting' },
        { desc: 'GenAI Consulting', unit: 'Hour', id: 'g6' },
      ]},
    ],
  },
  {
    title: 'Managed Review',
    groups: [
      { label: 'English Language Review – UK Based', rows: [
        { desc: 'Contract Reviewer', unit: 'Hour', id: 'r1' },
        { desc: 'Review Manager', unit: 'Hour', id: 'r2' },
        { desc: 'Review Project Management', unit: 'Hour', id: 'r3' },
      ]},
      { label: 'Offshore Review Centre – India Based', rows: [
        { desc: 'Contract Reviewer', unit: 'Hour', id: 'r4' },
        { desc: 'Review Manager', unit: 'Hour', id: 'r5' },
        { desc: 'Review Project Management', unit: 'Hour', id: 'r6' },
      ]},
    ],
  },
];

/* ── Build rate-sheet table rows ─────────────────────────────── */
function buildRateSheetSection(section, clientRates, cur) {
  const rows = [];

  rows.push(new TableRow({
    children: [new TableCell({
      children: [new Paragraph({
        children: [new TextRun({ text: section.title.toUpperCase(), bold: true, color: WHITE, size: pt(9), font: 'Calibri' })],
      })],
      columnSpan: 3,
      shading: shading(NAVY),
      margins: { top: mm(1.5), bottom: mm(1.5), left: mm(2), right: mm(2) },
      borders: noBorder(),
    })],
  }));

  section.groups.forEach(grp => {
    if (grp.label) {
      rows.push(new TableRow({
        children: [new TableCell({
          children: [new Paragraph({
            children: [new TextRun({ text: grp.label.toUpperCase(), bold: true, color: NAVY, size: pt(8), font: 'Calibri' })],
          })],
          columnSpan: 3,
          shading: shading(LTBLUE),
          margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
          borders: noBorder(),
        })],
      }));
    }

    grp.rows.forEach(row => {
      const clientRate = clientRates[row.id] !== undefined ? clientRates[row.id] : null;
      const rate = clientRate !== null ? validatedRate(row.id, clientRate) : clientRate;
      const rateStr = rate !== null ? `${cur}${rate < 1 ? rate.toFixed(2) : Number(rate).toLocaleString('en-GB')}` : 'On request';
      const noteStr = row.rateNote ? ` (${row.rateNote})` : '';

      rows.push(new TableRow({
        children: [
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: row.desc, size: pt(9), font: 'Calibri' })] })],
            width: { size: 60, type: WidthType.PERCENTAGE },
            margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
            borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C5D5E8' }, top: noBorderSide(), left: noBorderSide(), right: noBorderSide() },
          }),
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: row.unit, size: pt(9), font: 'Calibri', color: '666666' })] })],
            width: { size: 20, type: WidthType.PERCENTAGE },
            margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
            borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C5D5E8' }, top: noBorderSide(), left: noBorderSide(), right: noBorderSide() },
          }),
          new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: rateStr + noteStr, size: pt(9), font: 'Calibri', bold: true })],
              alignment: AlignmentType.RIGHT,
            })],
            width: { size: 20, type: WidthType.PERCENTAGE },
            margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
            borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C5D5E8' }, top: noBorderSide(), left: noBorderSide(), right: noBorderSide() },
          }),
        ],
      }));
    });
  });

  return rows;
}

/* ── Build specific-cost table rows ──────────────────────────── */
function buildCostTable(section, cur) {
  const rows = [];

  rows.push(new TableRow({
    children: [new TableCell({
      children: [new Paragraph({ children: [new TextRun({ text: section.title.toUpperCase(), bold: true, color: WHITE, size: pt(9), font: 'Calibri' })] })],
      columnSpan: 5,
      shading: shading(NAVY),
      margins: { top: mm(1.5), bottom: mm(1.5), left: mm(2), right: mm(2) },
      borders: noBorder(),
    })],
  }));

  const header = new TableRow({
    children: ['Service', 'Unit', 'Rate', 'Est. Units', 'Est. Total'].map((h, i) => new TableCell({
      children: [new Paragraph({
        children: [new TextRun({ text: h, bold: true, color: NAVY, size: pt(8), font: 'Calibri' })],
        alignment: i >= 2 ? AlignmentType.RIGHT : AlignmentType.LEFT,
      })],
      shading: shading(LTBLUE),
      margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
      borders: noBorder(),
    })),
  });
  rows.push(header);

  let subtotal = 0;
  section.groups.forEach(grp => {
    grp.rows.forEach(row => {
      const rateVal = validatedRate(row.id, row.rate);
      const qty = row.qty || 0;
      const total = qty * rateVal;
      if (!row.complimentary) subtotal += total;

      const totalStr = row.complimentary
        ? `${cur}${total.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (Complimentary)`
        : (qty > 0 ? `${cur}${total.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : '—');

      rows.push(new TableRow({
        children: [
          new TableCell({
            children: [
              new Paragraph({ children: [new TextRun({ text: row.desc, size: pt(9), font: 'Calibri' })] }),
              ...(row.descNote ? [new Paragraph({ children: [new TextRun({ text: row.descNote, size: pt(8), font: 'Calibri', italics: true, color: '666666' })] })] : []),
            ],
            width: { size: 35, type: WidthType.PERCENTAGE },
            margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
            borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C5D5E8' }, top: noBorderSide(), left: noBorderSide(), right: noBorderSide() },
          }),
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: row.unit, size: pt(9), font: 'Calibri', color: '666666' })] })],
            width: { size: 15, type: WidthType.PERCENTAGE },
            margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
            borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C5D5E8' }, top: noBorderSide(), left: noBorderSide(), right: noBorderSide() },
          }),
          new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: `${cur}${rateVal < 1 ? rateVal.toFixed(2) : Number(rateVal).toLocaleString('en-GB')}`, size: pt(9), font: 'Calibri', bold: true })],
              alignment: AlignmentType.RIGHT,
            })],
            width: { size: 15, type: WidthType.PERCENTAGE },
            margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
            borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C5D5E8' }, top: noBorderSide(), left: noBorderSide(), right: noBorderSide() },
          }),
          new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: qty > 0 ? String(qty) : '—', size: pt(9), font: 'Calibri' })],
              alignment: AlignmentType.RIGHT,
            })],
            width: { size: 15, type: WidthType.PERCENTAGE },
            margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
            borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C5D5E8' }, top: noBorderSide(), left: noBorderSide(), right: noBorderSide() },
          }),
          new TableCell({
            children: [new Paragraph({
              children: [new TextRun({
                text: totalStr,
                size: pt(9), font: 'Calibri', bold: !row.complimentary,
                strike: !!row.complimentary,
              })],
              alignment: AlignmentType.RIGHT,
            })],
            width: { size: 20, type: WidthType.PERCENTAGE },
            margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
            borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C5D5E8' }, top: noBorderSide(), left: noBorderSide(), right: noBorderSide() },
          }),
        ],
      }));
    });
  });

  rows.push(new TableRow({
    children: [
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: `Subtotal – ${section.title}`, bold: true, size: pt(9), font: 'Calibri', color: NAVY })] })],
        columnSpan: 4,
        shading: shading(LTBLUE),
        margins: { top: mm(1.5), bottom: mm(1.5), left: mm(2), right: mm(2) },
        borders: noBorder(),
      }),
      new TableCell({
        children: [new Paragraph({
          children: [new TextRun({ text: `${cur}${subtotal.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`, bold: true, size: pt(9), font: 'Calibri', color: NAVY })],
          alignment: AlignmentType.RIGHT,
        })],
        shading: shading(LTBLUE),
        margins: { top: mm(1.5), bottom: mm(1.5), left: mm(2), right: mm(2) },
        borders: noBorder(),
      }),
    ],
  }));

  return { tableRows: rows, subtotal };
}

/* ── Responsibility outlines ─────────────────────────────────── */
function buildResponsibilityBlock(key) {
  const OUTLINES = {
    forensic: {
      title: 'Forensic Technology & Consulting Responsibility Outline',
      left: [
        { label: 'Forensic Tech Services', items: ['Social Media Collections', 'Basic iPhone Mobile Device Acquisition (in-lab only)', 'Basic Windows Computer Imaging (in-lab only)', 'M365 and Google Workspace collections', 'Administrative Logistics (COC documentation, Case Shutdown)'] },
        { label: 'Standard Forensic Acquisitions', items: ['Cloud-Hosted Email Collections (Gmail, iCloud)', 'Cloud Storage Repository Collections (Dropbox, Box.com)', 'Email Archive Collections (Global Relay, Smarsh)', 'M365 and Google Workspace Corporate Email Collection', 'Mobile Device Collections (advanced iPhone in-lab)', 'Remote Collections', 'Onsite Collections'] },
      ],
      right: [
        { label: 'Forensic Consulting', items: ['Forensic Analysis', 'Project Oversight, QC, and Troubleshooting', 'Client Consultation Calls and Emails', 'Custom Collection Workflow Design', 'Advanced Mobile Device Forensics'] },
        { label: 'Senior Forensic Consulting', items: ['All Remediation Work', 'Expert Affidavit, Declaration, and Report Preparation', 'Expert Testimony', 'Advanced Forensic Consulting and Analytics'] },
      ],
    },
    hosting: {
      title: 'Project Management & Consulting Responsibility Outline',
      left: [
        { label: 'Project Management', items: ['Generation of ECA ingest and exceptions reporting', 'Running search terms in Digital Reef', 'Export of data from Digital Reef', 'Loading data to Relativity', 'Relativity environment setup and customisation', 'Relativity user training', 'Review validations set up', 'Disclosure preparation', 'Any client communications and instructions'] },
        { label: 'Senior Project Management', items: ['Perform advanced ECA analysis/complex filtering', 'Review and advise on ESI stipulation specifications', 'Execute Relativity Structured Analytics + update TAR rankings', 'Implementing advanced or complex review workflows', 'Perform complex or custom reporting'] },
      ],
      right: [
        { label: 'Consulting', items: ['Relativity Analytics + TAR predictive coding', 'Brainspace TAR workflow set-up and reporting', 'Mobile data review workflow', 'Blackout Consultation and Auto-redactions', 'Custom Relativity Scripts'] },
        { label: 'Senior Consulting', items: ['Consultation on preservation and disclosure obligations', 'Early Case Assessment (DR) consultation', 'Search Term Consultation and Translations', 'Consult on data privacy and cross-border transfer', 'Design and negotiate TAR and translation protocols', 'Provide affidavits and expert testimony'] },
      ],
    },
    review: {
      title: 'Managed Review Responsibility Outline',
      prose: [
        { heading: 'Review Managers', text: "Review Managers are experienced review lawyers who have completed TransPerfect Legal's Managed Review Oversight Training. They play a critical role in escalating questions, managing the Q/A Log, and implementing Quality Control (QC) and Quality Assurance (QA) measures. Their primary responsibility is to support the case team by ensuring that quality standards are met with the utmost accuracy and efficiency. Review Managers are supervised by the Project Manager and are fully dedicated to a single matter for the duration of the review." },
        { heading: 'Project Managers', text: "The Project Manager serves as a key member of the case team, working in collaboration with outside counsel from the planning stages to the conclusion of the review. They oversee the review team, including the Review Managers, and ensure all tasks are completed on time. All communications with counsel should be led by the Review Manager, with the Project Manager copied on all correspondence to maintain oversight and provide guidance." },
      ],
    },
  };

  const outline = OUTLINES[key];
  if (!outline) return [];
  const result = [];

  result.push(new Paragraph({
    children: [new TextRun({ text: outline.title.toUpperCase(), bold: true, color: TEAL, size: pt(10), font: 'Calibri' })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: TEAL } },
    spacing: { before: 240, after: 120 },
  }));

  if (outline.prose) {
    outline.prose.forEach(p => {
      result.push(new Paragraph({ children: [new TextRun({ text: p.heading, bold: true, size: pt(10), font: 'Calibri', color: NAVY })], spacing: { before: 120, after: 60 } }));
      result.push(bodyPara(p.text));
    });
    return result;
  }

  const maxRows = Math.max(
    outline.left.reduce((a, s) => a + 1 + s.items.length, 0),
    outline.right.reduce((a, s) => a + 1 + s.items.length, 0),
  );

  function colContent(sections) {
    const paras = [];
    sections.forEach(sec => {
      paras.push(new Paragraph({ children: [new TextRun({ text: sec.label, bold: true, size: pt(9), font: 'Calibri', color: NAVY })], spacing: { before: 100, after: 60 } }));
      sec.items.forEach(item => paras.push(new Paragraph({ children: [new TextRun({ text: `• ${item}`, size: pt(9), font: 'Calibri' })], spacing: { before: 30, after: 30 } })));
    });
    return paras;
  }

  const leftContent = colContent(outline.left);
  const rightContent = colContent(outline.right);

  result.push(new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorder(),
    rows: [new TableRow({
      children: [
        new TableCell({
          children: leftContent,
          width: { size: 49, type: WidthType.PERCENTAGE },
          borders: { right: { style: BorderStyle.SINGLE, size: 2, color: LTBLUE }, top: noBorderSide(), bottom: noBorderSide(), left: noBorderSide() },
          margins: { right: mm(4) },
        }),
        new TableCell({
          children: rightContent,
          width: { size: 49, type: WidthType.PERCENTAGE },
          borders: { left: { style: BorderStyle.SINGLE, size: 2, color: LTBLUE }, top: noBorderSide(), bottom: noBorderSide(), right: noBorderSide() },
          margins: { left: mm(4) },
        }),
      ],
    })],
  }));

  return result;
}

/* ── Scope text per phase ────────────────────────────────────── */
function buildScopeSection(selectedServices) {
  const paras = [];

  paras.push(sectionTitle('Scope & Assumptions'));

  paras.push(new Paragraph({
    children: [new TextRun({ text: `Pre-Work | `, color: NAVY, size: pt(12), font: 'Calibri' }), new TextRun({ text: 'LAUNCH', bold: true, color: TEAL, size: pt(12), font: 'Calibri' })],
    spacing: { before: 160, after: 60 },
  }));
  paras.push(bullet('Our consultants will sit with the Client case team to understand the issues, goals, timeframes, strategy and deliverables of the project.'));

  let phaseNum = 0;

  if (selectedServices.includes('forensic')) {
    phaseNum++;
    paras.push(phaseHeading(`Phase ${phaseNum}`, 'Forensic Data Collection'));
    ['Digital Forensic services are included as a resource for ad-hoc collections and consulting, ensuring support is available when needed.',
      'Services include: social media collections; mobile device acquisitions; computer imaging (standard and advanced); M365 and Google Workspace collections; cloud-hosted email and storage repository collections; email archive collections; remote and onsite collections; forensic analysis, consulting, and project oversight.',
      'Senior forensic consulting available for remediation work, expert affidavit and declaration preparation, expert testimony, and advanced forensic analytics.',
    ].forEach(t => paras.push(bullet(t)));
  }
  if (selectedServices.includes('processing')) {
    phaseNum++;
    paras.push(phaseHeading(`Phase ${phaseNum}`, 'Data Processing & Early Case Assessment'));
    ['The TransPerfect Legal Project Management team will process the collected data into Digital Reef, our proprietary processing and Early Case Assessment ("ECA") tool. We will use standard culling techniques like de-duplication, deNIST and date range filters to reduce the volume of data to be promoted for review.',
      'We would use the reporting functionality of Digital Reef\'s pre-review analytics to give insight to the data, for example custodian analysis, domain names, language, clustering, threading, sentiment analysis and "find more like" analysis based on existing relevant documents.',
    ].forEach(t => paras.push(bullet(t)));
  }
  if (selectedServices.includes('hosting')) {
    phaseNum++;
    paras.push(phaseHeading(`Phase ${phaseNum}`, 'Data Hosting, Analytics & Review Support'));
    ['Once data has been processed and filtered, it will be promoted to Relativity for review by the case team. TransPerfect Legal Project Managers will support the case team through the review process.',
      'TransPerfect Legal Consultants can support the use of Technology Assisted Review ("TAR") to accelerate document review through the use of predictive coding.',
      'Opposing disclosures will be made available in Relativity for case team review. TransPerfect Legal Project Managers and Consultants can support the use of de-duplication, conceptual similarity analytics, TAR and AI to identify relevant documents.',
    ].forEach(t => paras.push(bullet(t)));
  }
  if (selectedServices.includes('genai')) {
    phaseNum++;
    paras.push(phaseHeading(`Phase ${phaseNum}`, 'GenAI'));
    ['TransPerfect Legal offers a suite of GenAI-powered tools to accelerate review and provide deeper case intelligence.',
      'Cicero: AI-powered document analysis providing summaries, key issue identification and narrative building.',
      'eDiscoveryAI: Automated relevance and privilege review acceleration.',
      'Early Case Intelligence (ECI): Core analytics, Case Elements, and Casebot options for real-time case analysis.',
    ].forEach(t => paras.push(bullet(t)));
  }
  if (selectedServices.includes('review')) {
    phaseNum++;
    paras.push(phaseHeading(`Phase ${phaseNum}`, 'Managed Review'));
    ['Review Managers are experienced review lawyers who have completed TransPerfect Legal\'s Managed Review Oversight Training. They play a critical role in escalating questions, managing the Q/A Log, and implementing Quality Control (QC) and Quality Assurance (QA) measures.',
      'Project Managers oversee the review team, including the Review Managers, and ensure all tasks are completed on time. All communications with counsel should be led by the Review Manager.',
    ].forEach(t => paras.push(bullet(t)));
  }

  return paras;
}

/* ── Payment Terms ───────────────────────────────────────────── */
function buildPaymentTerms(cur) {
  const paras = [];
  paras.push(heading('Payment Terms', NAVY, 11, true, 200, 80));
  [
    `All fees are exclusive of VAT. Subject to a ${cur}250 minimum per technology services project.`,
    'TransPerfect Legal will invoice Client upon completion of services, or on a monthly basis for ongoing matters.',
    'Client shall pay all invoices within thirty (30) days of receipt.',
    'Any late payments shall accrue interest at the rate of 1.5% per month or the maximum rate permitted by applicable law, whichever is less.',
  ].forEach(t => paras.push(bodyPara(t, { spacing: { before: 60, after: 60 } })));
  return paras;
}

/* ── T&C content (two columns — real SOW content) ──────────────── */
function buildTandC() {

  function pt(n) { return n * 2; }

  function tcHeading(text) {
    return new Paragraph({
      spacing: { before: 100, after: 40 },
      children: [new TextRun({ text, bold: true, size: pt(7.5), color: '001F5F', allCaps: true, font: 'Calibri' })],
    });
  }

  function tcPara(text) {
    return new Paragraph({
      spacing: { before: 0, after: 60 },
      alignment: AlignmentType.JUSTIFIED,
      children: [new TextRun({ text, size: pt(7.5), font: 'Calibri' })],
    });
  }

  function tcParaRuns(runs) {
    return new Paragraph({
      spacing: { before: 0, after: 60 },
      alignment: AlignmentType.JUSTIFIED,
      children: runs.map(r => new TextRun({ ...r, size: pt(7.5), font: 'Calibri' })),
    });
  }

  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const dividerBorder = { style: BorderStyle.SINGLE, size: 4, color: 'DEEAF6' };

  function makeCell(paras, isLeft) {
    return new TableCell({
      width: { size: 4680, type: WidthType.DXA },
      borders: {
        top: noBorder, bottom: noBorder, left: noBorder,
        right: isLeft ? dividerBorder : noBorder,
      },
      margins: { right: isLeft ? 200 : 0, left: isLeft ? 0 : 200 },
      children: paras,
    });
  }

  function makePage(leftParas, rightParas, pageBreakBefore = false) {
    const row = new TableRow({ children: [makeCell(leftParas, true), makeCell(rightParas, false)] });
    return new Table({
      width: { size: 9360, type: WidthType.DXA },
      borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder },
      rows: [row],
    });
  }

  // ── LEFT COLUMN CONTENT ──────────────────────────────────────────
  const p1L = [
    tcHeading('1) INVOICING AND PAYMENT'),
    tcPara('Client has the sole responsibility to make full payment of all charges and expenses relating to the services contemplated in this Agreement (the "Project") within thirty (30) days of the date of receipt of the invoice from TRANSPERFECT LEGAL (which shall be deemed to be received no later than the third business day after such invoice is sent). Applicable taxes may be added to Client\'s invoices and are payable by Client. As of the first anniversary of the date hereof, and on each anniversary thereafter, TRANSPERFECT LEGAL shall increase all fees covered by this Agreement and any related SOW by ten percent (10%). Invoices not paid within the foregoing period or disputed in accordance with Section 3 below are hereby deemed delinquent ("Delinquent Invoices"). In the event that the duration of any Project extends beyond thirty (30) days, the Client will be invoiced according to the billing schedule contained in this Agreement (if any) or for work completed to date in 30-day intervals.'),
    tcPara('Collection of related payments from any third party (including, without limitation, any client of the Client) is a private matter of the Client and shall not affect the Client\'s responsibility for payment to TRANSPERFECT LEGAL. Upon signing this Agreement, Client shall provide TRANSPERFECT LEGAL with any special address, internal coding or purchase order number to ensure that Client\'s payments to TRANSPERFECT LEGAL is remitted in accordance with this Agreement and any other applicable agreement, and Client\'s failure to provide such information shall not relieve Client of the provisions of this Paragraph. Bills to be sent to CLIENT care of CLIENT CONTACT(S) in the event any invoice hereunder needs to be pre-approved by Client or a representative or agent of Client (the "Authorising Party") before being sent to the party responsible for payment hereunder (the "Responsible Party"). The Authorising Party hereby agrees that TRANSPERFECT LEGAL shall be permitted to send the invoice directly to the Responsible Party without Authorising Party\'s prior consent in the event such consent has not been received by TRANSPERFECT LEGAL within (XX) days of the date of delivery of such invoice to the Authorising Party.'),
    tcHeading('2) ADDITIONAL INVOICING AND PAYMENT FOR STAFFING SERVICES ONLY'),
    tcPara('To the extent that a Project includes Staffing services (i.e., supplying contract review attorneys or non-attorney reviewers (a "Document Reviewer")), the terms in this paragraph apply to such Staffing services.'),
    tcPara('TRANSPERFECT LEGAL will invoice Client every month. Invoices shall be accompanied by the pertinent reviewer timesheets or applicable electronic records pertaining to time worked. Client\'s signature and/or electronic approval on TRANSPERFECT LEGAL\'s timesheets certifies that the hours shown are correct and authorises TRANSPERFECT LEGAL to bill Client for the hours worked by the named Document Reviewer.'),
    tcPara('NO PAYROLL TRANSFER — Client agrees not to directly or indirectly cause or permit any engaged Document Reviewer assigned to Client by TRANSPERFECT LEGAL to transfer to another entity\'s payroll, or to perform services for Client whilst on the payroll of any person or firm other than TRANSPERFECT LEGAL during the term of this Agreement and for a period of 12 months thereafter.'),
  ];

  const p1R = [
    tcHeading('11) NO ORAL MODIFICATION OR WAIVER'),
    tcPara('Neither this Agreement nor the terms of any Project may be orally modified. Only a modification in writing, agreed to by both parties, will be enforceable. Either party may give its authorisation to any written modification via electronic mail which shall constitute such party\'s written agreement to such modification. No waiver of any breach hereof will be effective unless in writing and signed by an authorised representative of the party against whom enforcement is sought.'),
    tcHeading('12) FORCE MAJEURE'),
    tcPara('Performance by TRANSPERFECT LEGAL of any Project is excused as long as, and to the extent that, such performance is prevented by events outside of TRANSPERFECT LEGAL\'s reasonable control (each a "Force Majeure Event"), including, but not limited to, an act of God, act of public enemy or terrorist, fire, explosion, electrical power outage, strike, lockout or other labor disturbance, pandemic, epidemic, quarantine, embargo, war (whether declared or undeclared), civil disturbance, inflation, restraint or declaration of any governments, ruler or people or any other event, whether similar or dissimilar, that affects TRANSPERFECT LEGAL or its suppliers.'),
  ];

  const p2L = [
    tcHeading('3) UNPAID INVOICES'),
    tcPara('TRANSPERFECT LEGAL reserves the right to suspend or terminate any Project and performance hereunder and under all other agreements with the Client if the Client has any Delinquent Invoices or falls into arrears. Further, Delinquent Invoices are subject to interest of one and a half percent (1.5%) per month on any outstanding balance, or the maximum permitted by law, whichever is less.'),
    tcHeading('4) CHANGES TO PROJECT'),
    tcPara('a. Should the Client change the parameters of a Project whilst it is in progress (e.g. turnaround time accelerated, source files not provided on time, project\'s scope/size expanded, etc.), TRANSPERFECT LEGAL reserves the right to bill additional charges and/or extend the deadline for performance in accordance with the change requested.'),
    tcPara('b. When undertaking rush projects, TRANSPERFECT LEGAL shall use reasonable efforts to ensure the quality and accuracy of such rush projects. However, without the payment of additional fees by Client in order to account for such rush, TRANSPERFECT LEGAL cannot warrant that any rush project will be of the same quality or accuracy in comparison with a non-rush based project.'),
    tcPara('c. Client hereby grants to TRANSPERFECT LEGAL all rights and permissions in or relating to all information, data, records and other materials that are uploaded or otherwise received, directly or indirectly, from Client ("Client Data") to process such Client Data, in compliance with its confidentiality obligations, internally for the purpose of improving the quality of its analytics and improving its artificial intelligence and machine learning algorithms.'),
  ];

  const p2R = [
    tcHeading('13) SEVERABILITY'),
    tcPara('If any term, clause or provision hereof is held invalid or unenforceable by a court of competent jurisdiction, such invalidity shall not affect the validity or operation of any other term, clause or provision and the invalid term, clause or provision shall be deemed to be severed from this Agreement.'),
    tcHeading('14) TERMINATION'),
    tcPara('Either party may terminate this Agreement without cause upon giving the other party ninety (90) days prior written notice. During the ninety (90) day termination period, Client shall maintain TRANSPERFECT LEGAL staffing levels existing at the time of the notice of termination and shall reimburse TRANSPERFECT LEGAL for any reasonable demobilisation expenses such as equipment, personnel or real estate lease terminations, etc. The effective termination date shall be the last day of such ninety (90) day termination period and Client shall be responsible for paying all fees, expenses and other monies incurred up to and including the effective termination date.'),
    tcPara('Either party may terminate this Agreement for any continued material and substantial breach of the terms and conditions of this Agreement upon giving the other party fifteen (15) days prior written notice identifying specifically the alleged breach, provided that the breaching party does not cure such breach within the fifteen (15) day notice period.'),
  ];

  const p3L = [
    tcHeading('5) NOTICE OF CONCERNS'),
    tcPara('The Client must notify TRANSPERFECT LEGAL of any concerns with TRANSPERFECT LEGAL\'s performance of a Project within thirty (30) days of receipt of deliverables or completion of services or service milestones via certified letter (return receipt requested) or via electronic mail to a TRANSPERFECT LEGAL account representative. If TRANSPERFECT LEGAL is not so notified, the Client waives all rights and claims arising out of such performance.'),
    tcHeading('6) LIMITATION OF LIABILITY'),
    tcPara('a. TRANSPERFECT LEGAL EXPRESSLY DISCLAIMS ALL REPRESENTATIONS, WARRANTIES AND CONDITIONS OF ANY KIND, EXPRESS OR IMPLIED (INCLUDING ANY WARRANTY OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR USE, NONINFRINGEMENT AND THOSE ARISING BY STATUTE OR OTHERWISE IN LAW OR FROM A COURSE OF DEALING OR USE OF TRADE) TO THE FULLEST EXTENT PERMITTED BY LAW, WHETHER RELATED TO THE PROVISION OF SERVICES PROVIDED UNDER THIS AGREEMENT OR OTHERWISE.'),
    tcPara('b. IN NO EVENT SHALL TRANSPERFECT LEGAL OR ANY OF ITS PARENTS, SUBSIDIARIES OR AFFILIATES BE LIABLE UNDER ANY THEORY AT LAW OR OTHERWISE, TO CLIENT OR ANY THIRD PARTY FOR ANY INDIRECT, INCIDENTAL, SPECIAL, PUNITIVE, EXEMPLARY OR CONSEQUENTIAL DAMAGES ARISING OUT OF OR RELATING TO THIS AGREEMENT, INCLUDING, BUT NOT LIMITED TO LOSS OF PROFITS OR REVENUES, LOSS OF SAVINGS OR THE FAILURE OF CLIENT TO RECEIVE THE BENEFITS IT EXPECTS TO DERIVE FROM THE SERVICES, EVEN IF TRANSPERFECT LEGAL HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.'),
    tcPara('c. THE SERVICES PROVIDED UNDER THIS AGREEMENT MAY BE SUBJECT TO LIMITATIONS, DELAYS AND OTHER PROBLEMS INHERENT IN THE USE OF EQUIPMENT AND TECHNOLOGY. TRANSPERFECT LEGAL DOES NOT GUARANTEE THAT THE SERVICES WILL BE PERFORMED ERROR-FREE OR UNINTERRUPTED.'),
    tcPara('d. Any action against TRANSPERFECT LEGAL in connection with or arising out of services performed under this Agreement must be commenced within 1 year after the claim arose.'),
  ];

  const p3R = [
    tcHeading('15) FUTURE PROJECTS'),
    tcPara('The Client may place any future orders for copying, scanning, document management or litigation support services from TRANSPERFECT LEGAL, by electronic mail to TRANSPERFECT LEGAL, which shall constitute authorisation by the Client for any projects and services set forth in such order, and shall constitute the Client\'s acceptance that any such order shall be subject to this Agreement. Client agrees to use its best efforts to ensure that TRANSPERFECT LEGAL shall be its principal provider of translation and related software services and legal support services.'),
    tcHeading('16) DATA STORAGE/RETURN UPON PROJECT COMPLETION OR TERMINATION'),
    tcPara('Client agrees to notify TRANSPERFECT LEGAL when a Project (or the litigation or investigation with respect to which a Project relates) is closed and/or no longer active. At such time, it is TRANSPERFECT LEGAL\'s policy to securely destroy the electronically-stored information (ESI) that is in TRANSPERFECT LEGAL\'s possession by (a) securely purging it from our servers and stored media and (b) providing to Client a Certificate of Destruction. If Client prefers that TRANSPERFECT LEGAL securely return the ESI to Client prior to remediation from TRANSPERFECT LEGAL\'s systems, or that TRANSPERFECT LEGAL continue to store the ESI for Client, additional fees will apply.'),
  ];

  const p4L = [
    tcHeading('7) INDEMNIFICATION'),
    tcPara('a. The Client shall indemnify, defend and hold harmless TRANSPERFECT LEGAL, its Affiliates and each of their respective owners, principals, managers, representatives, partners, officers, directors, agents and employees against any and all claims, damages, losses, judgments, settlements, liabilities, costs and expenses (including reasonable attorneys\' fees) ("Losses") incurred by TRANSPERFECT LEGAL with respect to claims or demands by a third party to the extent arising out of or related to (i) any gross negligence or willful misconduct by the Client; (ii) infringement of a third party\'s intellectual property rights by any materials, data or information provided to TRANSPERFECT LEGAL by Client; (iii) an act or negligence by Client that constitutes a breach of privacy law; or (iv) any tangible property damage or personal injury caused by Client.'),
    tcPara('b. TRANSPERFECT LEGAL shall indemnify, defend and hold harmless Client, and its Affiliates against any and all Losses incurred by Client with respect to claims or demands by a third party to the extent arising out of or related to (i) any gross negligence or willful misconduct by TRANSPERFECT LEGAL; and (ii) any tangible property damage or personal injury caused by TRANSPERFECT LEGAL in performance of its obligations under the Agreement.'),
    tcPara('e. THE EXPRESS PROVISIONS OF THIS SECTION ARE IN LIEU OF, AND TO THE EXCLUSION OF, ALL OTHER INDEMNITY AND CONTRIBUTION OBLIGATIONS OF ANY KIND, EXPRESS OR IMPLIED, STATUTORY OR OTHERWISE, RELATING TO THE CLAIMS.'),
    tcHeading('8) REMEDIES; WAIVER'),
    tcPara('In the event of a breach, violation or default of this Agreement, nothing herein shall prevent the non-breaching party from pursuing all remedies it is now or hereafter entitled to under law or in equity. If TRANSPERFECT LEGAL must resort to collection of Delinquent Invoices by an agency or through legal action, the Client agrees to pay collection fees and reasonable attorney fees incurred by TRANSPERFECT LEGAL.'),
  ];

  const p4R = [
    tcHeading('17) GDPR COMPLIANCE'),
    tcPara('In order for the parties to comply with EU General Data Protection Regulation 2016/679 (the "GDPR"), Client acknowledges and agrees that it shall notify TRANSPERFECT LEGAL in writing at such time as Client provides any documents or information to TRANSPERFECT LEGAL containing "personal data" of European citizens (a "GDPR-Governed Project"). The parties further agree that for any GDPR-Governed Projects, the parties will execute a Data Protection Addendum Addressing Article 28 GDPR (Processor Terms) and Incorporating Standard Contractual Clauses for Controller to Processor Transfers of Personal Data from the EEA to a Third Country (the "GDPR Addendum").'),
    tcHeading('18) GOVERNING LAW; DISPUTE RESOLUTION'),
    tcPara('This Agreement and all rights and obligations of the parties relating hereto shall be governed by and construed in accordance with the laws of England and Wales without giving effect to any conflicts of law rules that would cause the application of the laws of any other jurisdiction. Any and all claims arising under or relating directly or indirectly to this Agreement, whether sounding in contract or tort, shall be exclusively brought in and subject to the exclusive jurisdiction of the courts that are located in London, England.'),
    tcPara('Without limiting the foregoing, at the sole election of Contractor, any controversy, dispute, or claim arising out of or relating to any amounts due pursuant to any invoices issued by TRANSPERFECT LEGAL pursuant to this Agreement that cannot be resolved amicably may be settled by binding arbitration before a single arbitrator, which shall take place in London, England and be administered by JAMS, Inc. ("JAMS").'),
  ];

  const p5L = [
    tcHeading('9) CONFIDENTIALITY'),
    tcPara('Each party agrees that all materials, data or information received by it from the other party, either directly or indirectly, in writing, orally or by drawings or inspection of samples, equipment or facilities, pursuant to, and whether prior to or following the execution of, this Agreement, is confidential and proprietary, whether or not expressly designated as such ("Confidential Material"), and agrees to use its best efforts to protect the confidentiality of the Confidential Material using at least the same measures it takes to protect its own confidential information of like kind.'),
    tcPara('The receiving party shall immediately give notice to the disclosing party of any unauthorised use or disclosure of disclosing party\'s Confidential Material. Confidential Material shall be returned or destroyed upon the request of the party designating such material as Confidential Material, or upon the completion or earlier termination of any work provided under this Agreement.'),
    tcHeading('10) NON-SOLICITATION; NON-HIRING'),
    tcPara('During the period in which services are being performed under this Agreement, and for a period of one (1) year thereafter, neither Client nor any of its Affiliates shall, directly or indirectly, solicit the employment of, employ or contract with, any other employee or contractor of TRANSPERFECT LEGAL or any of its Affiliates with whom Client or any of its Affiliates had contact. If Client or any of its parents, affiliates or subsidiaries breaches this paragraph, Client shall pay as liquidated damages, and not as a penalty, the sum of $75,000 per individual.'),
  ];

  const p5R = [
    tcHeading('19) NO LEGAL ADVICE'),
    tcPara('Client acknowledges and understands that neither TRANSPERFECT LEGAL nor any of its employees will be providing legal advice to Client and that an attorney-client relationship does not exist between TRANSPERFECT LEGAL and Client. Client agrees not to request that TRANSPERFECT LEGAL provide legal advice in connection with this Agreement or the Project. Notwithstanding the foregoing, the parties acknowledge that TRANSPERFECT LEGAL is being engaged to assist Client\'s counsel\'s representation of Client.'),
    tcHeading('20) ENTIRE AGREEMENT'),
    tcPara('This Agreement represents the full and complete agreement with respect to the subject matter set forth herein, and supersedes any other agreements, promises, representations, whether written or oral.'),
    tcHeading('21) SURVIVAL'),
    tcPara('The terms of Paragraphs 1, 2, 3, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 18, 20 and 21, and other paragraphs which by their nature are intended to extend beyond termination, shall survive termination of this Agreement for any reason.'),
    tcPara('The contents of this Agreement are proprietary to TRANSPERFECT LEGAL. Client shall not, without the prior written consent of TRANSPERFECT LEGAL, disclose this document or the contents herein to any third party.'),
  ];

  const spacer = new Paragraph({ children: [], spacing: { before: 0, after: 0 } });

  return [
    makePage(p1L, p1R, false),
    spacer,
    makePage(p2L, p2R, false),
    spacer,
    makePage(p3L, p3R, false),
    spacer,
    makePage(p4L, p4R, false),
    spacer,
    makePage(p5L, p5R, false),
  ];
}


/* ── Billing requirements table ──────────────────────────────── */
function buildBillingTable(d) {
  function row(label, value, isHeader = false) {
    if (isHeader) {
      return new TableRow({ children: [new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, color: WHITE, size: pt(9), font: 'Calibri', allCaps: true })] })],
        columnSpan: 2,
        shading: shading(NAVY),
        margins: { top: mm(1.5), bottom: mm(1.5), left: mm(2), right: mm(2) },
        borders: noBorder(),
      })] });
    }
    return new TableRow({ children: [
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: label, color: TEAL, bold: true, size: pt(9), font: 'Calibri' })] })],
        width: { size: 45, type: WidthType.PERCENTAGE },
        shading: shading('F8FBFF'),
        margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
        borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C5D5E8' }, top: noBorderSide(), left: noBorderSide(), right: noBorderSide() },
      }),
      new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: value || '', size: pt(9), font: 'Calibri' })] })],
        width: { size: 55, type: WidthType.PERCENTAGE },
        margins: { top: mm(1), bottom: mm(1), left: mm(2), right: mm(2) },
        borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C5D5E8' }, top: noBorderSide(), left: noBorderSide(), right: noBorderSide() },
      }),
    ] });
  }

  const cur = d.cur || '£';
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorder(),
    rows: [
      row('Client Billing Requirements', '', true),
      row('Submitter\'s Name', ''),
      row('Department / Team', ''),
      row('PO Required? If so, please provide', ''),
      row('Case / Reference No.', d.matter || ''),
      row('Currency', cur === '£' ? 'GBP (£)' : cur === '$' ? 'USD ($)' : 'EUR (€)'),
      row('Payment Terms', '30 days'),
      row('Requestor Information', '', true),
      row('Name', d.clientContactName || ''),
      row('Company', d.clientFirm || ''),
      row('Address', [d.clientAddr1, d.clientAddr2].filter(Boolean).join(', ')),
      row('Email', d.clientEmail || ''),
      row('Contact Information for Billing Enquiries', '', true),
      row('Name', d.tplAM || ''),
      row('Email', d.tplEmail || ''),
    ],
  });
}

/* ── Signature block ─────────────────────────────────────────── */
function buildSignatureBlock(clientFirm) {
  function sigCol(name) {
    return new TableCell({
      children: [
        new Paragraph({ children: [new TextRun({ text: name.toUpperCase(), bold: true, size: pt(9), font: 'Calibri', color: NAVY })], spacing: { after: 200 } }),
        ...[['Signature', 240], ['Printed Name', 240], ['Title', 240], ['Date', 240]].flatMap(([label, sp]) => [
          new Paragraph({ children: [], border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY } }, spacing: { before: sp, after: 40 } }),
          new Paragraph({ children: [new TextRun({ text: label, size: pt(8), font: 'Calibri', color: '666666' })], spacing: { after: 80 } }),
        ]),
      ],
      width: { size: 49, type: WidthType.PERCENTAGE },
      borders: noBorder(),
      margins: { right: mm(6) },
    });
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorder(),
    rows: [new TableRow({ children: [sigCol('TransPerfect Document Management, Inc.'), sigCol(clientFirm)] })],
  });
}

/* ── Details & Definitions exhibit ──────────────────────────── */
function buildExhibit() {
  const sections = [];

  sections.push(pageBreakPara());
  sections.push(sectionTitle('Details & Definitions'));

  sections.push(heading('A. Forensic Service Case Parameters', NAVY, 11, true, 200, 80));
  [
    'All forensic collections and analysis services are provided by qualified forensic practitioners using industry-standard tools and methods.',
    'Standard acquisitions include logical and physical device acquisitions, cloud collections, and email archive collections.',
    'All collections are documented with appropriate chain-of-custody records.',
    'Forensic consulting includes project oversight, client consultation, and review of forensic methodology.',
    'Senior forensic consulting is required for remediation work, expert testimony, affidavit preparation, and advanced analytics.',
    'After-hours and weekend rates apply to services scheduled or required outside of standard business hours (09:00–18:00 Monday–Friday, excluding UK public holidays).',
  ].forEach(t => sections.push(bullet(t)));

  sections.push(heading('B. eDiscovery Definitions', NAVY, 11, true, 200, 80));
  const definitions = [
    ['Native Review', 'Review of documents in their original file format, as opposed to image/TIFF format.'],
    ['ECA (Early Case Assessment)', 'The process of evaluating a collection of documents to understand its scope, content, and relevance before full review.'],
    ['Processing', 'The conversion of raw data files into a format suitable for review, including de-duplication, deNIST filtering, and indexing.'],
    ['Hosting', 'The provision of a secure, web-based environment for document review, typically using Relativity or similar platforms.'],
    ['TAR (Technology Assisted Review)', 'The use of machine learning algorithms to assist in identifying relevant documents during the review process.'],
    ['Brainspace', 'A conceptual analytics and TAR platform integrated with Relativity, used for predictive coding and document clustering.'],
    ['Productions', 'The disclosure of relevant documents to opposing parties in the agreed format, typically TIFF images with extracted text and load files.'],
    ['Data Disposition', 'The secure deletion or return of client data at the conclusion of a matter, in accordance with agreed procedures.'],
  ];
  definitions.forEach(([term, def]) => {
    sections.push(new Paragraph({
      children: [new TextRun({ text: `${term}: `, bold: true, size: pt(9), font: 'Calibri', color: NAVY }), new TextRun({ text: def, size: pt(9), font: 'Calibri' })],
      spacing: { before: 60, after: 60 },
    }));
  });

  sections.push(heading('C. Managed Review', NAVY, 11, true, 200, 80));
  [
    'All managed reviewers are legally qualified or have equivalent experience and are trained on TransPerfect Legal\'s quality management protocols.',
    'Review Managers are responsible for day-to-day review management, issue escalation, and quality control.',
    'Project Managers provide strategic oversight of the review, coordinating with counsel and ensuring timelines are met.',
    'All review is conducted in accordance with an agreed review protocol, prepared in consultation with the instructing law firm.',
    'TransPerfect Legal maintains a continuous quality assurance programme, with a minimum 10% QC rate applied to all reviewer decisions.',
    'Daily reporting on review progress, including metrics on volumes reviewed, decisions made, and any issues identified, is provided as standard.',
  ].forEach(t => sections.push(bullet(t)));

  sections.push(heading('D. Case Closure Options', NAVY, 11, true, 200, 80));
  [
    'At the conclusion of a matter, TransPerfect Legal will work with the Client to agree appropriate data disposition procedures.',
    'Options include: secure deletion of all hosted data; return of data to Client on encrypted media; transfer to a third-party platform nominated by Client.',
    'All data disposition is documented and a certificate of destruction/transfer provided to Client upon request.',
    'Archive storage is available at reduced rates for matters that have concluded review but where data must be retained for a defined period.',
    'Database disposal fees apply for the deletion of Relativity databases. Databases under 1TB are charged a flat fee; databases over 1TB are charged per GB.',
  ].forEach(t => sections.push(bullet(t)));

  return sections;
}

/* ── Main handler ────────────────────────────────────────────── */
exports.handler = async (event) => {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
  };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 200, headers: corsHeaders, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers: corsHeaders, body: 'Method Not Allowed' };

  try {
    const payload = JSON.parse(event.body || '{}');
    const {
      clientContactName, clientFirm, clientEntity, clientAddr1, clientAddr2, clientEmail,
      tplAM, tplTitle, tplEmail, tplAddr1, tplAddr2,
      endClient, matter, agreementDate, currency,
      includeRateSheet, includeSpecificCosts, includeExhibit,
      selectedServices, specificCostData, rateSheetRates,
      filename,
    } = payload;

    const cur = currency || '£';

    /* ── Cover: parties table ───────────────────────────────── */
    const dateStr = agreementDate
      ? new Date(agreementDate + 'T12:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })
      : 'Date TBC';

    const partiesTable = new Table({
      width: { size: 70, type: WidthType.PERCENTAGE },
      alignment: AlignmentType.CENTER,
      borders: noBorder(),
      rows: [
        new TableRow({ children: [
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: 'PREPARED FOR', bold: true, size: pt(8), font: 'Calibri', color: NAVY, allCaps: true })], alignment: AlignmentType.RIGHT })],
            width: { size: 50, type: WidthType.PERCENTAGE }, borders: noBorder(), margins: { right: mm(4) },
          }),
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: 'PREPARED BY', bold: true, size: pt(8), font: 'Calibri', color: NAVY, allCaps: true })] })],
            width: { size: 50, type: WidthType.PERCENTAGE }, borders: noBorder(), margins: { left: mm(4) },
          }),
        ]}),
        ...[
          [clientFirm || '', tplAM || 'TransPerfect Legal'],
          [clientEntity || '', tplTitle || ''],
          [clientAddr1 || '', tplAddr1 || 'Aldgate House, 33 Aldgate High Street'],
          [clientAddr2 || '', tplAddr2 || 'London, EC3N 1AH'],
          [clientEmail || '', tplEmail || ''],
        ].map(([left, right]) => new TableRow({ children: [
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: left, size: pt(9), font: 'Calibri' })], alignment: AlignmentType.RIGHT })],
            borders: noBorder(), margins: { right: mm(4) },
          }),
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: right, size: pt(9), font: 'Calibri' })] })],
            borders: noBorder(), margins: { left: mm(4) },
          }),
        ]})),
      ],
    });

    /* ── Intro paragraph ────────────────────────────────────── */
    const introText = `This Case Agreement Statement of Work ("Agreement") is entered into as of ${dateStr} by and between TransPerfect Document Management, Inc. d/b/a TransPerfect Legal Solutions ("TRANSPERFECT LEGAL"), and ${clientEntity || clientFirm || '[Client Legal Entity]'} ("Client"), with respect to the matter known as ${matter || '[Matter Name]'}${endClient ? ` for ${endClient}` : ''}.`;

    /* ── Scope section ──────────────────────────────────────── */
    const scopeParas = buildScopeSection(selectedServices || []);

    /* ── Pricing section ────────────────────────────────────── */
    const pricingParas = [];
    pricingParas.push(sectionTitle('Pricing'));

    if (includeRateSheet) {
      pricingParas.push(heading('Rate Schedule', NAVY, 12, true, 200, 100));
      const clientRates = rateSheetRates || {};
      const rateTableRows = [];
      RATE_CARD.forEach(section => {
        rateTableRows.push(...buildRateSheetSection(section, clientRates, cur));
      });
      pricingParas.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorder(),
        rows: rateTableRows,
      }));
    }

    if (includeSpecificCosts && specificCostData) {
      if (includeRateSheet) {
        pricingParas.push(new Paragraph({ children: [], spacing: { before: 240 } }));
        pricingParas.push(heading('Estimated Cost Breakdown', NAVY, 12, true, 200, 100));
      }

      let grandTotal = 0;
      specificCostData.forEach(section => {
        const { tableRows, subtotal } = buildCostTable(section, cur);
        pricingParas.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders: noBorder(), rows: tableRows }));
        pricingParas.push(new Paragraph({ children: [], spacing: { before: 120 } }));
        grandTotal += subtotal;

        const respParas = buildResponsibilityBlock(section.key);
        if (respParas.length > 0) pricingParas.push(...respParas);
      });

      pricingParas.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorder(),
        rows: [new TableRow({ children: [
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: 'TOTAL ESTIMATED FEES', bold: true, color: WHITE, size: pt(11), font: 'Calibri' })] })],
            width: { size: 70, type: WidthType.PERCENTAGE },
            shading: shading(NAVY),
            margins: { top: mm(3), bottom: mm(3), left: mm(4), right: mm(4) },
            borders: noBorder(),
          }),
          new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: `${cur}${grandTotal.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`, bold: true, color: TEAL, size: pt(14), font: 'Calibri' })],
              alignment: AlignmentType.RIGHT,
            })],
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: shading(NAVY),
            margins: { top: mm(3), bottom: mm(3), left: mm(4), right: mm(4) },
            borders: noBorder(),
          }),
        ]})],
      }));
    }

    pricingParas.push(...buildPaymentTerms(cur));

    /* ── Build document ─────────────────────────────────────── */
    const doc = new Document({
      creator: 'TransPerfect Legal SOW Generator',
      title: `Statement of Work – ${clientFirm || 'Client'} – ${matter || 'Matter'}`,
      description: 'Auto-generated by TPL SOW Generator',
      styles: {
        default: {
          document: {
            run: { font: 'Calibri', size: pt(10) },
          },
        },
      },
      sections: [
        /* ── SECTION 1: Cover ─────────────────────────────── */
        {
          properties: {
            page: {
              margin: { top: mm(20), bottom: mm(20), left: mm(20), right: mm(20) },
            },
          },
          children: [
            /* Logo placeholder / company name */
            new Paragraph({
              children: [new TextRun({ text: 'TRANSPERFECT', bold: true, color: NAVY, size: pt(28), font: 'Calibri' }), new TextRun({ text: 'LEGAL', bold: true, color: TEAL, size: pt(28), font: 'Calibri' })],
              alignment: AlignmentType.CENTER,
              spacing: { before: 480, after: 120 },
            }),
            new Paragraph({
              children: [new TextRun({ text: 'CASE AGREEMENT STATEMENT OF WORK', bold: true, color: TEAL, size: pt(13), font: 'Calibri', characterSpacing: 80 })],
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 480 },
            }),
            partiesTable,
            new Paragraph({
              children: [new TextRun({ text: `Date: ${dateStr}`, bold: true, color: NAVY, size: pt(9), font: 'Calibri' })],
              alignment: AlignmentType.CENTER,
              spacing: { before: 240, after: 0 },
            }),
          ],
        },

        /* ── SECTION 2: Body ──────────────────────────────── */
        {
          properties: {
            page: {
              margin: { top: mm(25), bottom: mm(25), left: mm(25), right: mm(25) },
            },
          },
          footers: {
            default: new Footer({
              children: [new Paragraph({
                children: [
                  new TextRun({ text: 'TransPerfect Legal Solutions  |  Aldgate House, 33 Aldgate High Street, London, EC3N 1AH  |  ', size: pt(7.5), font: 'Calibri', color: '666666' }),
                  new TextRun({ text: 'www.transperfect.com', size: pt(7.5), font: 'Calibri', color: TEAL }),
                ],
                alignment: AlignmentType.CENTER,
              })],
            }),
          },
          children: [
            bodyPara(introText),
            new Paragraph({ children: [], spacing: { before: 120 } }),
            ...scopeParas,
            hrRule(),
            ...pricingParas,
          ],
        },

        /* ── SECTION 3: T&Cs ──────────────────────────────── */
        {
          properties: {
            page: {
              margin: { top: mm(20), bottom: mm(20), left: mm(18), right: mm(18) },
            },
          },
          children: [
            sectionTitle('Terms & Conditions'),
            new Paragraph({ children: [], spacing: { before: 160 } }),
            buildTandC(clientFirm || 'Client'),
            pageBreakPara(),
            sectionTitle('Client Billing Requirements'),
            new Paragraph({ children: [], spacing: { before: 120 } }),
            buildBillingTable({ clientContactName, clientFirm, clientAddr1, clientAddr2, clientEmail, tplAM, tplEmail, matter, cur }),
            hrRule(),
            new Paragraph({
              children: [new TextRun({ text: 'In Witness, whereof, TransPerfect and the Client have caused this Agreement to be executed by their duly authorised officers or agents as of the day and year written below.', size: pt(9), font: 'Calibri' })],
              spacing: { before: 160, after: 160 },
              alignment: AlignmentType.JUSTIFIED,
            }),
            buildSignatureBlock(clientFirm || 'Client'),
          ],
        },

        /* ── SECTION 4: Exhibit (conditional) ────────────── */
        ...(includeExhibit ? [{
          properties: {
            page: {
              margin: { top: mm(25), bottom: mm(25), left: mm(25), right: mm(25) },
            },
          },
          children: buildExhibit(),
        }] : []),
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    return {
      statusCode: 200,
      headers: {
        ...corsHeaders,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${filename || 'TPL_SOW.docx'}"`,
      },
      body: buffer.toString('base64'),
      isBase64Encoded: true,
    };

  } catch (err) {
    console.error('DOCX generation error:', err);
    return {
      statusCode: 500,
      headers: corsHeaders,
      body: JSON.stringify({ error: err.message, stack: err.stack }),
    };
  }
};
