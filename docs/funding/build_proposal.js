const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  ImageRun, Header, Footer, PageNumber, LevelFormat, convertInchesToTwip,
} = require('docx');

const GREEN = '2E7D32';
const DARK = '1F2A24';
const GREY = '5A6660';
const LIGHT = 'EEF5EE';
const RULE = 'C8D8C8';

const CONTENT_W = 9746; // A4 (11906) minus 2 x 1080 margins

const logo = fs.readFileSync(__dirname + '/../media/word/media/image1.jpeg');

// ---------- helpers ----------
const t = (text, o = {}) => new TextRun({ text, font: 'Calibri', size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || DARK, highlight: o.highlight });

const body = (text, o = {}) =>
  new Paragraph({
    spacing: { after: o.after === undefined ? 140 : o.after, line: 264 },
    indent: o.indent,
    alignment: o.alignment,
    children: Array.isArray(text) ? text : [t(text, o)],
  });

const h1 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 140 },
    children: [new TextRun({ text, font: 'Calibri', size: 26, bold: true, color: GREEN, allCaps: true })],
  });

const h2 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, font: 'Calibri', size: 22, bold: true, color: DARK })],
  });

const bullet = (runs, o = {}) =>
  new Paragraph({
    numbering: { reference: 'dash', level: 0 },
    spacing: { after: o.after === undefined ? 80 : o.after, line: 264 },
    children: Array.isArray(runs) ? runs : [t(runs)],
  });

const numbered = (runs, ref) =>
  new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 80, line: 264 },
    children: Array.isArray(runs) ? runs : [t(runs)],
  });

const fill = (label) => t(label, { highlight: 'yellow', bold: true });

const noBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};

const hairline = {
  top: { style: BorderStyle.SINGLE, size: 2, color: RULE },
  bottom: { style: BorderStyle.SINGLE, size: 2, color: RULE },
  left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: RULE },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};

const cell = (children, o = {}) =>
  new TableCell({
    width: { size: o.width, type: WidthType.DXA },
    shading: o.shade ? { type: ShadingType.CLEAR, fill: o.shade, color: 'auto' } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    columnSpan: o.span,
    children,
  });

const txtCell = (text, o = {}) =>
  cell([new Paragraph({
    alignment: o.alignment,
    spacing: { after: 0, line: 252 },
    children: [t(text, { bold: o.bold, size: o.size || 20, color: o.color })],
  })], o);

// ---------- letterhead ----------
const letterhead = new Header({
  children: [
    new Table({
      columnWidths: [900, 8846],
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: noBorders,
      rows: [
        new TableRow({
          children: [
            cell([new Paragraph({
              spacing: { after: 0 },
              children: [new ImageRun({ data: logo, type: 'jpg', transformation: { width: 46, height: 42 } })],
            })], { width: 900 }),
            cell([
              new Paragraph({
                spacing: { after: 0 },
                children: [new TextRun({ text: 'TLHAGO AGRI PRO (PTY) LTD', font: 'Calibri', size: 24, bold: true, color: GREEN })],
              }),
              new Paragraph({
                spacing: { after: 0 },
                children: [new TextRun({
                  text: 'Reg No 2025/973670/07  ·  Stand 159, Matankane Village, Mphahlele 0739  ·  P.O. Box 51, Lefalane 0741',
                  font: 'Calibri', size: 15, color: GREY,
                })],
              }),
              new Paragraph({
                spacing: { after: 0 },
                children: [new TextRun({
                  text: 'shakukabelo2@gmail.com  ·  069 188 0424  ·  tlhagoagricultural.net',
                  font: 'Calibri', size: 15, color: GREY,
                })],
              }),
            ], { width: 8846 }),
          ],
        }),
      ],
    }),
    new Paragraph({
      spacing: { before: 60, after: 0 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GREEN } },
      children: [new TextRun({ text: '', size: 2 })],
    }),
  ],
});

const foot = new Footer({
  children: [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        children: ['Tlhago Agri Pro (Pty) Ltd  ·  Funding Proposal  ·  Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES],
        font: 'Calibri', size: 16, color: GREY,
      })],
    }),
  ],
});

// ---------- ask box ----------
const askBox = new Table({
  columnWidths: [CONTENT_W],
  width: { size: CONTENT_W, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
    left: { style: BorderStyle.SINGLE, size: 18, color: GREEN },
    right: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
    insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  },
  rows: [
    new TableRow({
      children: [
        new TableCell({
          width: { size: CONTENT_W, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: LIGHT, color: 'auto' },
          margins: { top: 160, bottom: 160, left: 200, right: 200 },
          children: [
            new Paragraph({
              spacing: { after: 60 },
              children: [new TextRun({ text: 'WE ARE ASKING FOR', font: 'Calibri', size: 16, bold: true, color: GREEN, allCaps: true })],
            }),
            new Paragraph({
              spacing: { after: 80 },
              children: [new TextRun({ text: 'R332 700', font: 'Calibri', size: 52, bold: true, color: GREEN })],
            }),
            new Paragraph({
              spacing: { after: 0, line: 264 },
              children: [t('To plant, grow and harvest our first commercial crop — red onions on 5 hectares at Matankane Village, Lepelle-Nkumpi Municipality, Limpopo.')],
            }),
          ],
        }),
      ],
    }),
  ],
});

// ---------- budget ----------
const budgetRow = (item, amount, why, o = {}) =>
  new TableRow({
    children: [
      txtCell(item, { width: 4200, bold: o.bold }),
      txtCell(amount, { width: 1700, alignment: AlignmentType.RIGHT, bold: o.bold }),
      txtCell(why, { width: 3846, color: o.bold ? DARK : GREY }),
    ],
  });

const groupRow = (label) =>
  new TableRow({
    children: [
      cell([new Paragraph({ spacing: { after: 0 }, children: [t(label, { bold: true, size: 18, color: GREEN })] })],
        { width: 4200, shade: LIGHT }),
      cell([new Paragraph({ spacing: { after: 0 }, children: [t('', { size: 18 })] })], { width: 1700, shade: LIGHT }),
      cell([new Paragraph({ spacing: { after: 0 }, children: [t('', { size: 18 })] })], { width: 3846, shade: LIGHT }),
    ],
  });

const budgetTable = new Table({
  columnWidths: [4200, 1700, 3846],
  width: { size: CONTENT_W, type: WidthType.DXA },
  borders: hairline,
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        txtCell('ITEM', { width: 4200, bold: true, size: 18, color: GREEN }),
        txtCell('AMOUNT', { width: 1700, bold: true, size: 18, color: GREEN, alignment: AlignmentType.RIGHT }),
        txtCell('WHAT IT IS FOR', { width: 3846, bold: true, size: 18, color: GREEN }),
      ],
    }),
    groupRow('EQUIPMENT — ONCE-OFF'),
    budgetRow('Tractor', 'R250 000', 'Works the land every season from here on'),
    budgetRow('Ploughing attachment', 'R17 250', 'Prepares and ploughs the fields'),
    budgetRow('Water system and electricity', 'R10 000', 'Borehole and water reserve tank'),
    budgetRow('Fencing material', 'R3 000', 'Secures the project site'),
    budgetRow('Subtotal', 'R280 250', 'Assets we own after this season', { bold: true }),
    groupRow('GROWING COSTS — ONE SEASON'),
    budgetRow('Fertiliser', 'R20 000', 'Full season across 4.5 hectares'),
    budgetRow('Pesticide', 'R10 000', 'Pest control on the fields'),
    budgetRow('Petrol for water pump', 'R11 200', 'R200 a day for 56 days of furrow irrigation'),
    budgetRow('Red onion seed (Creole)', 'R1 250', 'The crop itself'),
    budgetRow('Subtotal', 'R42 450', 'Repeats each season', { bold: true }),
    groupRow('PEOPLE'),
    budgetRow('Harvest wages', 'R10 000', '50 community members at R200 each'),
    budgetRow('Subtotal', 'R10 000', 'Paid straight into local households', { bold: true }),
    new TableRow({
      children: [
        cell([new Paragraph({ spacing: { after: 0 }, children: [t('GRAND TOTAL', { bold: true, size: 22, color: GREEN })] })], { width: 4200, shade: LIGHT }),
        cell([new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 }, children: [t('R332 700', { bold: true, size: 22, color: GREEN })] })], { width: 1700, shade: LIGHT }),
        cell([new Paragraph({ spacing: { after: 0 }, children: [t('', { size: 20 })] })], { width: 3846, shade: LIGHT }),
      ],
    }),
  ],
});

// ---------- risk table ----------
const riskRow = (risk, fix, header) =>
  new TableRow({
    tableHeader: header,
    children: [
      cell([new Paragraph({ spacing: { after: 0, line: 252 }, children: [t(risk, { bold: header, size: header ? 18 : 20, color: header ? GREEN : DARK })] })], { width: 3900 }),
      cell([new Paragraph({ spacing: { after: 0, line: 252 }, children: [t(fix, { bold: header, size: header ? 18 : 20, color: header ? GREEN : DARK })] })], { width: 5846 }),
    ],
  });

const riskTable = new Table({
  columnWidths: [3900, 5846],
  width: { size: CONTENT_W, type: WidthType.DXA },
  borders: hairline,
  rows: [
    riskRow('THE RISK', 'WHAT WE DO ABOUT IT', true),
    riskRow('The land belongs to the traditional authority, not to us.',
      'We sign a written lease over the 5 hectares with the Induna and the Chief before a single rand is spent. No funder money goes into land we have no written right to use.'),
    riskRow('There is no access road or bridge over the Olifants River to the main road.',
      'The first season is planned around the existing route. Once there is real commercial traffic, we motivate to the municipality for the crossing with the harvest volumes as evidence.'),
    riskRow('We are a price-taker at the fresh produce markets.',
      'We split sales between national market agents and local retailers, and rotate crops by season so one bad price does not take the whole year.'),
    riskRow('Heavy rain or a dry spell can damage a crop.',
      'The river gives us irrigation through dry spells. Staggered planting means a single storm cannot reach the whole field at once.'),
    riskRow('We have no equipment of our own yet.',
      'That is what this proposal fixes. Until it is funded we keep a season moving by renting equipment rather than standing still.'),
  ],
});

// ---------- document ----------
const doc = new Document({
  creator: 'Tlhago Agri Pro (Pty) Ltd',
  title: 'Tlhago Agri Pro — Funding Proposal',
  description: 'Funding proposal for the first commercial crop at Matankane Village, Limpopo',
  numbering: {
    config: [
      {
        reference: 'dash',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 340, hanging: 200 } }, run: { color: GREEN } },
        }],
      },
      {
        reference: 'short',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 400, hanging: 260 } }, run: { color: GREEN, bold: true } },
        }],
      },
      {
        reference: 'long',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 400, hanging: 260 } }, run: { color: GREEN, bold: true } },
        }],
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        margin: { top: 1900, right: 1080, bottom: 1080, left: 1080, header: 480, footer: 480 },
      },
    },
    headers: { default: letterhead },
    footers: { default: foot },
    children: [
      new Paragraph({
        spacing: { after: 40 },
        children: [new TextRun({ text: 'FUNDING PROPOSAL', font: 'Calibri', size: 36, bold: true, color: DARK })],
      }),
      new Paragraph({
        spacing: { after: 260 },
        children: [new TextRun({ text: 'First commercial crop  ·  Matankane Village, Limpopo  ·  ', font: 'Calibri', size: 20, color: GREY }),
                   new TextRun({ text: '[Month, Year]', font: 'Calibri', size: 20, color: GREY, bold: true, highlight: 'yellow' })],
      }),

      askBox,

      new Paragraph({ spacing: { before: 200, after: 140, line: 264 }, children: [
        t('The money buys three things: '),
        t('the equipment we do not have', { bold: true }),
        t(' — a tractor, a plough, fencing and a water system; '),
        t('the inputs for one full season', { bold: true }),
        t(' — seed, fertiliser, pesticide and fuel; and '),
        t('wages for 50 local people at harvest', { bold: true }),
        t('.'),
      ]}),
      body('Eighty-four percent of it is once-off. After this season the tractor is ours, the field is fenced and watered, and every season that follows costs R52 450 to run.'),

      h1('Why we exist'),
      body([t('We exist so that people can eat. That is the whole of it.', { bold: true })]),
      body('Charity hands out food for a day. We would rather build the thing that puts food on the table every season: a working farm that employs the people around it, sells into the formal market, and puts its surplus back into the village it came from.'),
      body('Tlhago Agri Pro is led by a young farmer from Matankane and staffed by the community it serves. We are not asking for money to give away. We are asking for money to start something that pays wages.'),

      h1('The problem'),
      body('Matankane is one of the smallest villages in Lepelle-Nkumpi Municipality, in Limpopo’s Capricorn District — a district of roughly 1.4 million people spread across 663 villages, most of them rural and poor.'),
      bullet('Most households here have no wage income and survive on social grants, hunting and backyard gardening.'),
      bullet('Only about 13% of the community holds a tertiary qualification.'),
      bullet('There is no health facility nearby, and no primary school close to the children.'),
      bullet('Unemployment among the youth is near total, and the village is led by its elderly.'),
      bullet('Meanwhile hundreds of hectares of fertile land lie unused beside a permanent river.', { after: 140 }),
      body([t('The land is here. The water is here. The people are here. Only the work is missing — and that is the part we build.', { italics: true, color: GREEN, bold: true })]),

      h1('What we do'),
      bullet([t('Grow. ', { bold: true }), t('Seasonal vegetables and field crops, rotated by season, starting with red onions.')]),
      bullet([t('Pack and sell. ', { bold: true }), t('Into the national fresh produce markets — Tshwane Market and City Deep — through agents already willing to take our produce, and to local retailers such as Spar and U-Save.')]),
      bullet([t('Employ. ', { bold: true }), t('Community members on a rotational basis, with 50 seasonal workers at harvest in year one.')]),
      bullet([t('Give back. ', { bold: true }), t('Surplus is paid into a community trust for food support to the poorest households, care for orphans and vulnerable children, and skills training.')]),

      h1('Why it works here'),
      bullet([t('Land. ', { bold: true }), t('The traditional authority has offered unused village land — 5 hectares to start, with roughly 50 more that can be cleared, and further extension available for developments that benefit the community.')]),
      bullet([t('Water. ', { bold: true }), t('The Olifants River runs beside the site. Irrigation water is unlimited and costs us fuel, not purchase.')]),
      bullet([t('Soil. ', { bold: true }), t('Fertile arable land already used for field crops by households in the village.')]),
      bullet([t('Buyers. ', { bold: true }), t('Agents in place for the national markets, and retailers within delivery reach of the farm.')]),
      bullet([t('People. ', { bold: true }), t('A young workforce with the will to work and nothing else to do.')]),

      h1('Our objectives'),
      h2('First 12 months'),
      numbered('Sign a written lease over the 5-hectare site with the traditional authority.', 'short'),
      numbered('Fence the site, fit the water system, and bring the 5 hectares into production.', 'short'),
      numbered('Plant, harvest and sell one full red onion crop.', 'short'),
      numbered('Put 50 community members on the payroll through the harvest.', 'short'),
      numbered('Open the community trust account and pay in the first surplus.', 'short'),
      h2('Three to five years'),
      numbered('Bring the remaining ~50 hectares into production and rotate crops year-round.', 'long'),
      numbered('Add agro-processing: tomato sauce, mango atchar, and packed butternut.', 'long'),
      numbered('Add livestock: goats, sheep and pigs, sold live and as packed meat.', 'long'),
      numbered('Fund skills training, a local school and better health access from the trust.', 'long'),

      new Paragraph({ children: [new TextRun({ text: '', break: 1 })] }),

      h1('What the money buys'),
      budgetTable,
      new Paragraph({ spacing: { before: 140, after: 140, line: 264 }, children: [
        t('R280 250 of this — 84% — is equipment we will still own in ten years. Strip that out and a full season of growing and harvesting costs R52 450. '),
        t('This is the expensive season. Every one after it is cheap.', { bold: true }),
      ]}),

      h1('What comes back'),
      bullet('50 households earning wages at harvest in the first season alone.'),
      bullet('A fenced, irrigated, tractor-equipped farm that can plant every season after this for R52 450.'),
      bullet('Produce into the formal market, and produce donated to the poorest households in the village.'),
      bullet('Every rand of surplus paid into the community trust, for skills, schooling and infrastructure.', { after: 140 }),

      h2('Season one revenue'),
      new Paragraph({ spacing: { after: 140, line: 264 }, children: [
        t('Onion revenue is yield multiplied by price: '),
        fill('[  ] tonnes per hectare'),
        t(' across 5 hectares, sold at '),
        fill('R[  ] per 10 kg bag'),
        t(' at Tshwane Market, giving '),
        fill('R[  ]'),
        t(' for the season against R332 700 invested. Confirm these three figures with your market agent and fill them in before this proposal goes out.'),
      ]}),

      h1('The risks, honestly'),
      riskTable,

      h1('Where this goes'),
      bullet([t('Phase 1 — funded by this proposal. ', { bold: true }), t('Five hectares, red onions, 50 jobs at harvest, one full season proven end to end.')]),
      bullet([t('Phase 2. ', { bold: true }), t('The full ~50 hectares under rotation, and a second crop line: tomatoes, butternut, mango.')]),
      bullet([t('Phase 3. ', { bold: true }), t('An agro-processing plant and livestock, and a community trust large enough to build roads, a skills school and health facilities.')]),
      body('We are asking you to fund Phase 1. The rest is what Phase 1 makes possible.', { after: 60 }),

      h1('Who we are'),
      body([
        t('Tlhago Agri Pro (Pty) Ltd', { bold: true }),
        t(' is a South African company, registration number 2025/973670/07, based at Stand 159, Matankane Village, Mphahlele, Limpopo.'),
      ]),
      body([
        t('Project manager: ', { bold: true }), t('Kabelo Shaku — '), fill('[one line: your farming background and years of experience]'),
        t('. The project manager runs day-to-day operations, draws the seasonal plan, and reports to the board.'),
      ]),
      body('Two coordinators supervise field teams and report to the project manager. A board of advisors — the project management, the landowner and experienced farmers — oversees the project and mediates any agreement signed on behalf of the community.'),
      body('Field workers are hired from the community on a rotational basis, season by season.'),

      h1('Next step'),
      body('We would welcome a site visit. The land, the river and the people are all within a few minutes of each other, and they make the case better than paper does.'),
      new Paragraph({ spacing: { before: 40, after: 140, line: 264 }, children: [
        t('Kabelo Shaku  ·  shakukabelo2@gmail.com  ·  069 188 0424', { bold: true }),
      ]}),
      body([t('Available on request: ', { bold: true }), t('company registration certificate, signed land lease, site photographs, market agent letters of intent, and the full seasonal project plan.', { color: GREY })]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(__dirname + '/Tlhago_Agri_Pro_Funding_Proposal.docx', buf);
  console.log('written', buf.length, 'bytes');
});
