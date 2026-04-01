const pptxgen = require('pptxgenjs');
const {
  imageSizingCrop,
  imageSizingContain,
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
  safeOuterShadow,
} = require('/home/oai/skills/slides/pptxgenjs_helpers');
const path = require('path');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'OpenAI';
pptx.company = 'OpenAI';
pptx.subject = 'AIonOS pitch for Mitsui';
pptx.title = 'Transforming Mitsui business outcomes with AI';
pptx.lang = 'en-US';
pptx.theme = {
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'en-US'
};
pptx.defineSlideMaster({
  title: 'MASTER',
  background: { color: 'F6F8FB' },
  objects: [
    { rect: { x: 0, y: 0, w: 13.333, h: 0.2, fill: { color: '0F172A' }, line: { color: '0F172A' } } },
    { rect: { x: 0.5, y: 7.06, w: 12.33, h: 0.01, fill: { color: 'D7DFEA' }, line: { color: 'D7DFEA' } } },
    { text: { text: 'AIonOS x Mitsui', options: { x: 0.55, y: 7.1, w: 2.5, h: 0.18, fontSize: 8, color: '64748B', margin: 0 } } },
  ]
});

const C = {
  navy: '0F172A',
  blue: '0B5FFF',
  cyan: '11B5E4',
  teal: '0F766E',
  green: '16A34A',
  red: 'DC2626',
  orange: 'EA580C',
  slate: '475569',
  light: 'F6F8FB',
  lighter: 'E2E8F0',
  white: 'FFFFFF'
};
const A = p => path.join('/mnt/data/mitsui_repo/assets', p);

function addTitle(slide, title, subtitle='') {
  slide.addText(title, { x: 0.65, y: 0.42, w: 8.8, h: 0.45, fontFace: 'Aptos Display', fontSize: 24, bold: true, color: C.navy, margin: 0 });
  if (subtitle) slide.addText(subtitle, { x: 0.67, y: 0.9, w: 8.8, h: 0.26, fontSize: 10.5, color: C.slate, margin: 0 });
}

function addNote(slide, lines) { slide.addNotes(lines.join('\n')); }

function pill(slide, x,y,w,h,text,fill,color='FFFFFF') {
  slide.addText(text, { x,y,w,h, fill: { color: fill }, line: { color: fill }, color, fontSize: 12, bold: true, align: 'center', valign: 'mid', radius: 0.14, margin: 0.03 });
}

function metricCard(slide, x, y, w, h, number, label, sub, color) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: 'FFFFFF' }, line: { color: 'DCE5F2', pt: 1 } });
  slide.addShape(pptx.ShapeType.rect, { x: x, y: y, w: 0.08, h, fill: { color }, line: { color } });
  slide.addText(number, { x: x+0.18, y: y+0.16, w: w-0.25, h: 0.34, fontSize: 22, bold: true, color: C.navy, margin: 0 });
  slide.addText(label, { x: x+0.18, y: y+0.55, w: w-0.25, h: 0.26, fontSize: 11.5, bold: true, color: C.slate, margin: 0 });
  if (sub) slide.addText(sub, { x: x+0.18, y: y+0.89, w: w-0.25, h: 0.24, fontSize: 9, color: '64748B', margin: 0 });
}

function outcomeBand(slide, x, y, w, text, color) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h: 0.42, rectRadius: 0.06, fill: { color }, line: { color } });
  slide.addText(text, { x: x+0.04, y: y+0.06, w: w-0.08, h: 0.24, align: 'center', fontSize: 11, color: 'FFFFFF', bold: true, margin: 0 });
}

// Slide 1
{
  const s = pptx.addSlide('MASTER');
  s.background = { color: '0B1220' };
  s.addImage({ path: A('energy_control_room.png'), x: 5.7, y: 0.45, w: 7.15, h: 5.15 });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: '09111F', transparency: 36 }, line: { color: '09111F', transparency: 100 } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 6.2, h: 7.5, fill: { color: '09111F', transparency: 0 }, line: { color: '09111F', transparency: 100 } });
  s.addImage({ path: A('aionos_logo.png'), ...imageSizingContain(A('aionos_logo.png'), 0.65, 0.55, 1.95, 0.52) });
  s.addText('Transforming Mitsui\'s next growth engines with AI', { x: 0.7, y: 1.4, w: 4.85, h: 1.25, fontFace: 'Aptos Display', fontSize: 26, bold: true, color: 'FFFFFF', margin: 0, valign: 'mid' });
  s.addText('Executive pitch deck scaffold\nTop-consulting style · 5-second attention-span design', { x: 0.73, y: 2.9, w: 4.6, h: 0.7, fontSize: 12, color: 'D5E1F0', margin: 0, breakLine: false });
  pill(s, 0.72, 4.1, 1.5, 0.34, 'ENERGY', '0B5FFF');
  pill(s, 2.35, 4.1, 1.7, 0.34, 'HEALTHCARE', '16A34A');
  pill(s, 4.2, 4.1, 1.05, 0.34, 'DATA', 'EA580C');
  s.addText('Built around Mitsui\'s digital transformation agenda, with emphasis on energy transition, wellness ecosystem creation, and cross-business productivity.', { x: 0.73, y: 4.7, w: 4.75, h: 0.8, fontSize: 11.5, color: 'D5E1F0', margin: 0 });
  s.addText('MITSUI × AIONOS', { x: 0.75, y: 6.6, w: 2.4, h: 0.28, fontSize: 10, color: 'A9BCD1', bold: true, margin: 0 });
  addNote(s, [
    '[Sources]',
    '- Internal account plan: Mitsui Account Plan.pptx (uploaded by user)',
    '- Mitsui medium-term management plan: https://www.mitsui.com/jp/en/company/outline/management/index.html',
    '- Mitsui restructuring release: https://www.mitsui.com/jp/en/release/2026/1252736_15351.html',
    '- AIonOS website: https://aionos.io/',
    '[/Sources]'
  ]);
}

// Slide 2
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'Why now: Mitsui is reorganizing around the next value pools', 'One message: AI should accelerate Mitsui\'s strongest strategic motions, not add another tool layer.');
  s.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: 1.35, w: 3.95, h: 4.9, rectRadius: 0.06, fill: { color: 'FFFFFF' }, line: { color: 'DCE5F2' } });
  s.addShape(pptx.ShapeType.roundRect, { x: 4.7, y: 1.35, w: 3.95, h: 4.9, rectRadius: 0.06, fill: { color: 'FFFFFF' }, line: { color: 'DCE5F2' } });
  s.addShape(pptx.ShapeType.roundRect, { x: 8.7, y: 1.35, w: 3.95, h: 4.9, rectRadius: 0.06, fill: { color: 'FFFFFF' }, line: { color: 'DCE5F2' } });
  pill(s, 0.95, 1.62, 1.45, 0.34, '1', C.blue);
  pill(s, 4.95, 1.62, 1.45, 0.34, '2', C.green);
  pill(s, 8.95, 1.62, 1.45, 0.34, '3', C.orange);
  s.addText('Global energy transition', { x: 0.95, y: 2.06, w: 3.1, h: 0.35, fontSize: 20, bold: true, color: C.navy, margin: 0 });
  s.addText('Wellness ecosystem creation', { x: 4.95, y: 2.06, w: 3.2, h: 0.35, fontSize: 20, bold: true, color: C.navy, margin: 0 });
  s.addText('Industrial business solutions', { x: 8.95, y: 2.06, w: 3.25, h: 0.35, fontSize: 20, bold: true, color: C.navy, margin: 0 });
  s.addImage({ path: A('mitsui_building.jpg'), ...imageSizingCrop(A('mitsui_building.jpg'), 0.95, 2.58, 3.25, 1.5) });
  s.addImage({ path: A('healthcare_ai.png'), ...imageSizingCrop(A('healthcare_ai.png'), 4.95, 2.58, 3.25, 1.5) });
  s.addImage({ path: A('enterprise_command.png'), ...imageSizingCrop(A('enterprise_command.png'), 8.95, 2.58, 3.25, 1.5) });
  s.addText('Mitsui says the current MTMP is driven by three key initiatives and explicitly highlights Global Energy Transition and Wellness Ecosystem Creation.', { x: 0.95, y: 4.28, w: 3.12, h: 0.66, fontSize: 11.5, color: C.slate, margin: 0 });
  s.addText('Your account-plan intelligence points to healthcare as a board-level digital-transformation wedge, especially for multilingual engagement and operating consistency.', { x: 4.95, y: 4.28, w: 3.12, h: 0.86, fontSize: 11.5, color: C.slate, margin: 0 });
  s.addText('The opportunity is cross-BU orchestration: data, risk, and decision support moving faster across energy, mobility, healthcare, and trading.', { x: 8.95, y: 4.28, w: 3.12, h: 0.8, fontSize: 11.5, color: C.slate, margin: 0 });
  s.addText('AI thesis', { x: 0.95, y: 6.42, w: 0.75, h: 0.2, fontSize: 10, bold: true, color: C.blue, margin: 0 });
  s.addText('Win by deploying one governed AI operating layer across the businesses Mitsui already prioritizes.', { x: 1.85, y: 6.35, w: 10.0, h: 0.34, fontSize: 12.5, bold: true, color: C.navy, margin: 0 });
  addNote(s, [
    '[Sources]',
    '- Internal account plan: Mitsui Account Plan.pptx (uploaded by user)',
    '- Mitsui MTMP 2026 page: https://www.mitsui.com/jp/en/company/outline/management/index.html',
    '- Mitsui restructuring release noting three key strategic initiatives: https://www.mitsui.com/jp/en/release/2026/1252736_15351.html',
    '[/Sources]'
  ]);
}

// Slide 3
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'Fastest executive storyline: 3 value pools, 1 AI operating system', 'Keep the slide glanceable: three business outcomes and one unifying control layer.');
  metricCard(s, 0.8, 1.55, 3.95, 1.5, 'Energy operations', 'Predict outages, optimize throughput, reduce incident exposure', 'Control towers, plant monitoring, trading adjacencies', C.blue);
  metricCard(s, 0.8, 3.25, 3.95, 1.5, 'Healthcare service', 'Automate patient / customer journeys across languages and channels', 'Contact centers, hospitals, digital service layers', C.green);
  metricCard(s, 0.8, 4.95, 3.95, 1.5, 'Enterprise decisioning', 'Connect data, workflows, and agent actions across business units', 'Portfolio, procurement, risk, compliance, PMO', C.orange);
  s.addImage({ path: A('enterprise_command.png'), ...imageSizingCrop(A('enterprise_command.png'), 5.2, 1.55, 7.45, 4.75) });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.15, y: 5.55, w: 3.55, h: 0.56, rectRadius: 0.06, fill: { color: '0F172A' }, line: { color: '0F172A' } });
  s.addText('UNI STACK = governance + orchestration + observability', { x: 7.3, y: 5.72, w: 3.25, h: 0.18, fontSize: 11, color: 'FFFFFF', bold: true, margin: 0, align: 'center' });
  s.addText('The executive promise is simple: Mitsui does not need separate AI experiments per business. It needs one operating layer that routes models, data, workflows, and controls to the highest-value outcomes.', { x: 5.2, y: 6.3, w: 7.3, h: 0.46, fontSize: 12, color: C.slate, margin: 0 });
  addNote(s, [
    '[Sources]',
    '- Internal account plan: Mitsui Account Plan.pptx (uploaded by user)',
    '- AIonOS homepage: https://aionos.io/',
    '- UniStack page: https://aionos.io/intellimate',
    '[/Sources]'
  ]);
}

// Slide 4
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'Use case 1: AI control tower for energy and infrastructure assets', 'For executives: safer operations, faster decisions, better utilization.');
  s.addImage({ path: A('energy_control_room.png'), ...imageSizingCrop(A('energy_control_room.png'), 0.75, 1.35, 7.1, 5.3) });
  outcomeBand(s, 8.2, 1.55, 3.95, 'PREDICTIVE MONITORING', C.blue);
  s.addText('Stream plant, trading, maintenance and cybersecurity signals into one operating cockpit.', { x: 8.25, y: 2.08, w: 4.0, h: 0.6, fontSize: 14, color: C.navy, margin: 0 });
  outcomeBand(s, 8.2, 2.95, 3.95, 'MANAGEMENT BY EXCEPTION', C.green);
  s.addText('AI flags anomalies, prioritizes response playbooks, and routes only material events to humans.', { x: 8.25, y: 3.48, w: 4.0, h: 0.6, fontSize: 14, color: C.navy, margin: 0 });
  outcomeBand(s, 8.2, 4.35, 3.95, 'BOARD-LEVEL OUTCOME', C.orange);
  s.addText('Lower incident risk, less downtime, and better asset performance across energy transition investments.', { x: 8.25, y: 4.88, w: 4.0, h: 0.6, fontSize: 14, color: C.navy, margin: 0 });
  s.addText('Illustrative pilot wedge', { x: 8.23, y: 5.92, w: 1.9, h: 0.18, fontSize: 10, bold: true, color: C.blue, margin: 0 });
  s.addText('Start with a single LNG / power / industrial asset cluster, then scale the same playbook across similar portfolios.', { x: 8.25, y: 6.12, w: 4.1, h: 0.38, fontSize: 11.2, color: C.slate, margin: 0 });
  addNote(s, [
    '[Sources]',
    '- Internal account plan: Mitsui Account Plan.pptx (uploaded by user)',
    '- Mitsui MTMP and restructuring materials emphasizing energy transition: https://www.mitsui.com/jp/en/company/outline/management/index.html ; https://www.mitsui.com/jp/en/release/2026/1252736_15351.html',
    '- AIonOS homepage / UniStack page: https://aionos.io/ ; https://aionos.io/intellimate',
    '[/Sources]'
  ]);
}

// Slide 5
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'Use case 2: multilingual AI service layer for healthcare and wellness', 'For executives: better experience, lower service cost, more scalable growth.');
  s.addImage({ path: A('healthcare_ai.png'), ...imageSizingCrop(A('healthcare_ai.png'), 6.1, 1.4, 6.6, 5.15) });
  metricCard(s, 0.8, 1.65, 4.55, 1.28, 'Patient / customer journeys', 'Automate triage, scheduling, FAQs, follow-up and digital front door', '', C.green);
  metricCard(s, 0.8, 3.15, 4.55, 1.28, 'Multilingual consistency', 'Standardize service quality across countries, channels and subsidiaries', '', C.blue);
  metricCard(s, 0.8, 4.65, 4.55, 1.28, 'Growth without linear headcount', 'Support healthcare expansion with AI-first service delivery', '', C.orange);
  s.addText('Why this matters for Mitsui', { x: 6.2, y: 5.9, w: 2.2, h: 0.22, fontSize: 10.5, bold: true, color: C.green, margin: 0 });
  s.addText('Healthcare is already a strategic growth domain in the account plan; AI makes the operating model easier to replicate across regions.', { x: 6.2, y: 6.12, w: 5.9, h: 0.38, fontSize: 11.2, color: C.slate, margin: 0 });
  addNote(s, [
    '[Sources]',
    '- Internal account plan: Mitsui Account Plan.pptx (uploaded by user)',
    '- AIonOS homepage (enterprise AI operating system positioning): https://aionos.io/',
    '[/Sources]'
  ]);
}

// Slide 6
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'AIonOS platform map for Mitsui', 'Position products as an outcome stack, not a software catalog.');
  s.addShape(pptx.ShapeType.roundRect, { x: 0.9, y: 1.55, w: 2.65, h: 4.65, rectRadius: 0.06, fill: { color: 'EAF2FF' }, line: { color: 'C7DAFF' } });
  s.addShape(pptx.ShapeType.roundRect, { x: 3.63, y: 1.55, w: 2.65, h: 4.65, rectRadius: 0.06, fill: { color: 'ECFDF5' }, line: { color: 'C3F0D6' } });
  s.addShape(pptx.ShapeType.roundRect, { x: 6.36, y: 1.55, w: 2.65, h: 4.65, rectRadius: 0.06, fill: { color: 'FFF4EA' }, line: { color: 'FFD9BF' } });
  s.addShape(pptx.ShapeType.roundRect, { x: 9.09, y: 1.55, w: 2.65, h: 4.65, rectRadius: 0.06, fill: { color: 'EEF2FF' }, line: { color: 'D4DCFF' } });
  pill(s, 1.18, 1.9, 1.1, 0.34, 'UniStack', C.blue);
  pill(s, 3.91, 1.9, 1.1, 0.34, 'UniWeave', C.green);
  pill(s, 6.64, 1.9, 1.1, 0.34, 'UniScale', C.orange);
  pill(s, 9.37, 1.9, 1.1, 0.34, 'UniProtect', '5B5BD6');
  s.addText('Control layer', { x: 1.18, y: 2.38, w: 1.9, h: 0.24, fontSize: 18, bold: true, color: C.navy, margin: 0 });
  s.addText('Service layer', { x: 3.91, y: 2.38, w: 1.9, h: 0.24, fontSize: 18, bold: true, color: C.navy, margin: 0 });
  s.addText('Data layer', { x: 6.64, y: 2.38, w: 1.9, h: 0.24, fontSize: 18, bold: true, color: C.navy, margin: 0 });
  s.addText('Resilience layer', { x: 9.37, y: 2.38, w: 1.9, h: 0.24, fontSize: 18, bold: true, color: C.navy, margin: 0 });
  s.addText('Orchestrate models, agents, workflows, and cloud/data operations across Mitsui\'s businesses.', { x: 1.18, y: 2.85, w: 2.05, h: 1.0, fontSize: 13.5, color: C.slate, margin: 0 });
  s.addText('Transform customer, patient, and partner interactions across channels and geographies.', { x: 3.91, y: 2.85, w: 2.05, h: 1.0, fontSize: 13.5, color: C.slate, margin: 0 });
  s.addText('Monetize data and improve forecasting, pricing, and portfolio-level decision support.', { x: 6.64, y: 2.85, w: 2.05, h: 1.0, fontSize: 13.5, color: C.slate, margin: 0 });
  s.addText('Reduce cyber and operational risk across critical assets and digital infrastructure.', { x: 9.37, y: 2.85, w: 2.05, h: 1.0, fontSize: 13.5, color: C.slate, margin: 0 });
  outcomeBand(s, 1.12, 5.15, 2.2, 'Cross-BU governance', C.blue);
  outcomeBand(s, 3.85, 5.15, 2.2, 'Healthcare growth', C.green);
  outcomeBand(s, 6.58, 5.15, 2.2, 'Trading intelligence', C.orange);
  outcomeBand(s, 9.31, 5.15, 2.2, 'Asset protection', '5B5BD6');
  addNote(s, [
    '[Sources]',
    '- AIonOS homepage and product pages: https://aionos.io/ ; https://aionos.io/intellimate',
    '- Internal account plan: Mitsui Account Plan.pptx (uploaded by user)',
    '[/Sources]'
  ]);
}

// Slide 7
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, '90-day executive ask', 'Keep the close commercial: one pilot, one sponsor set, one measurable outcome.');
  metricCard(s, 0.9, 1.6, 3.7, 1.2, 'Step 1', 'Select 1 energy and 1 healthcare pilot wedge', 'Board-level thesis, local execution', C.blue);
  metricCard(s, 0.9, 3.05, 3.7, 1.2, 'Step 2', 'Stand up the common AI operating layer', 'Data + workflow + governance', C.green);
  metricCard(s, 0.9, 4.5, 3.7, 1.2, 'Step 3', 'Track 3 metrics only', 'Speed, cost, risk / service quality', C.orange);
  s.addShape(pptx.ShapeType.roundRect, { x: 5.1, y: 1.55, w: 7.55, h: 4.95, rectRadius: 0.06, fill: { color: '0F172A' }, line: { color: '0F172A' } });
  s.addText('Pilot scoreboard', { x: 5.45, y: 1.95, w: 2.0, h: 0.28, fontSize: 18, bold: true, color: 'FFFFFF', margin: 0 });
  metricCard(s, 5.45, 2.55, 2.0, 1.15, '<12 weeks', 'to show usable pilot output', '', C.blue);
  metricCard(s, 7.85, 2.55, 2.0, 1.15, '2 sponsors', 'business + digital owner', '', C.green);
  metricCard(s, 10.25, 2.55, 2.0, 1.15, '1 stack', 'shared controls and observability', '', C.orange);
  s.addText('Recommended narrative in the room', { x: 5.45, y: 4.15, w: 3.3, h: 0.22, fontSize: 10.5, color: 'A9BCD1', bold: true, margin: 0 });
  s.addText('“We are not proposing another AI pilot factory. We are proposing one operating layer that improves Mitsui\'s priority businesses first, then scales enterprise-wide.”', { x: 5.45, y: 4.45, w: 6.55, h: 1.05, fontSize: 17, bold: true, color: 'FFFFFF', margin: 0 });
  addNote(s, [
    '[Sources]',
    '- Internal account plan: Mitsui Account Plan.pptx (uploaded by user)',
    '- AIonOS homepage and product pages: https://aionos.io/ ; https://aionos.io/intellimate',
    '[/Sources]'
  ]);
}

for (const sl of pptx._slides) {
  warnIfSlideHasOverlaps(sl, pptx);
  warnIfSlideElementsOutOfBounds(sl, pptx);
}

pptx.writeFile({ fileName: '/mnt/data/mitsui_repo/output/Mitsui_AIonOS_Exec_Pitch.pptx' });
