// FINAL_v5 — one slide, one idea. Visual system:
// grid: 0.75" side margins, 8.5" content width, 3 columns of 2.5" with 0.5" gutters
// vertical rhythm: kicker 0.50 · title 0.78 · subtitle 1.32 · content 2.0–4.6 · footer 5.0
// type: title 24 bold · hero 60 · stat 40 · body 13 · caption 10 · kicker 9 mono
// color: one coral emphasis per slide; everything else white / grey. No boxes unless they carry meaning.
const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = 'How I would build Doodle Labs Europe';
pres.author = 'Sergio Benavidez';

const BG = '0B0D10', WHITE = 'F2F4F6', BODY = 'B4BBC3', MUTED = '6F7882', CORAL = 'F05A55', FAINT = '1A1E24', LINE = '2A3139';
const F = 'Segoe UI', MONO = 'Consolas';
const X = 0.75, W = 8.5, COL = 2.5, GUT = 0.5;
const colX = i => X + i * (COL + GUT);
let n = 0;

function slide({ kicker, title, sub, footer, number = true }) {
  const s = pres.addSlide();
  s.background = { color: BG };
  n += 1;
  if (kicker) s.addText(kicker, { x: X, y: 0.5, w: 6.5, h: 0.22, fontFace: MONO, fontSize: 9, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true });
  if (title) s.addText(title, { x: X, y: 0.78, w: W, h: 0.5, fontFace: F, fontSize: 24, bold: true, color: WHITE, margin: 0, valign: 'top', isTextBox: true });
  if (sub) s.addText(sub, { x: X, y: 1.32, w: W, h: 0.35, fontFace: F, fontSize: 13, color: BODY, margin: 0, valign: 'top', isTextBox: true });
  if (footer) s.addText(footer, { x: X, y: 5.0, w: 7.6, h: 0.25, fontFace: F, fontSize: 9.5, color: MUTED, margin: 0, isTextBox: true });
  if (number) s.addText(String(n), { x: 8.75, y: 5.0, w: 0.5, h: 0.25, fontFace: MONO, fontSize: 9, color: MUTED, align: 'right', margin: 0, isTextBox: true });
  return s;
}
const text = (s, t, o) => s.addText(t, Object.assign({ fontFace: F, fontSize: 13, color: BODY, margin: 0, valign: 'top', isTextBox: true }, o));
const label = (s, t, o) => s.addText(t, Object.assign({ fontFace: MONO, fontSize: 9, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true }, o));
function stat(s, x, y, w, big, cap, { color = WHITE, size = 40 } = {}) {
  text(s, big, { x, y, w, h: size / 60, fontSize: size, bold: true, color, valign: 'bottom' });
  text(s, cap, { x, y: y + size / 60 + 0.12, w, h: 0.6, fontSize: 12 });
}

// 1 · Title
{
  const s = slide({ kicker: 'DOODLE LABS · DIRECTOR OF SALES, EUROPE · SERGIO BENAVIDEZ', number: false });
  text(s, 'How I would build\nDoodle Labs Europe', { x: X, y: 1.45, w: W, h: 1.4, fontSize: 38, bold: true, color: WHITE });
  text(s, 'Production is scaling. The radio is chosen once, at design freeze.\nI\'d win Europe account by account.', { x: X, y: 3.0, w: 7, h: 0.7, fontSize: 14 });
  label(s, 'MARKET  ·  NEW TERRITORY  ·  LEADS & RESELLERS  ·  30-60-90', { x: X, y: 4.95, w: W, h: 0.25 });
  s.addNotes(`~30 s. First: "Before I start — what's the real end time? I'll fit this to it."
"You asked for four things; I'll take them in that order, in about eight minutes."
Everything comes from public sources — it's a hypothesis until it meets your CRM.`);
}

// 2 · Market
{
  const s = slide({ kicker: '1 · MARKET', title: 'Europe is buying unmanned systems at scale', sub: 'And the money lands with the OEM — that\'s who we sell to.' });
  const d = [['€540M', 'Loitering munitions\nBundeswehr · Feb 2026'], ['6,000', 'Soldier drones\nFrance · 2026'], ['5', 'EU drone & defence projects\n28 Sep 2026']];
  d.forEach((v, i) => stat(s, colX(i), 2.3, COL, v[0], v[1], { size: 44, color: i === 0 ? CORAL : WHITE }));
  s.addNotes(`~40 s. €540M: Bundestag, 25-Feb-2026, Helsing HX-2 and STARK Virtus, with options beyond.
6,000: DGA, 1,000 + 5,000 DELCO (Harmattan), Jun-2026.
5 EDPCI approved by EU defence ministers 28-Sep-2026, incl. drones and counter-drone.
Trend: autonomy raises the comms bar — more range, more nodes, more jamming. Air, ground and sea.
Fibre, if raised: it keeps the disposable short-range tier; RF mesh keeps reusable, multi-node, ISR, C2 and UGV. "How do you want the team to answer fibre?"`);
}

// 3 · Production
{
  const s = slide({ kicker: '1 · MARKET', title: 'Top-20 OEM output could grow ~5× by 2030', footer: 'External model, base case, 20 accounts with published volumes. Not a Doodle Labs forecast.' });
  stat(s, X, 1.95, 2.3, '~50k', 'platforms a year by 2030,\nfrom ~10k in 2025', { size: 54, color: CORAL });
  s.addChart(pres.charts.BAR, [
    { name: '2025', labels: ['Strike', 'ISR', 'UGV', 'Interceptors'], values: [3000, 3550, 3300, 0] },
    { name: '2030', labels: ['Strike', 'ISR', 'UGV', 'Interceptors'], values: [21500, 12000, 8450, 8000] },
  ], {
    x: 3.55, y: 1.55, w: 5.7, h: 3.2, barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 70,
    chartColors: [LINE, CORAL], showLegend: true, legendPos: 't', legendColor: MUTED, legendFontFace: F, legendFontSize: 9,
    showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: MUTED, dataLabelFontSize: 8, dataLabelFontFace: F, dataLabelFormatCode: '#,##0;;;',
    catAxisLabelColor: BODY, catAxisLabelFontFace: F, catAxisLabelFontSize: 10, catAxisLineShow: false,
    valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  });
  s.addNotes(`~45 s. Units, not money. 2025: 9,850; 2030 base: 49,950; upside ~112k.
Only ~13k of 2026 units rest on visible contracts; after that it's options and announced ramps — the conservative case falls after 2027.
Model: 21 platforms, 53 evidence items (39 from contracts or delivery rates). Disclosure bias: German start-ups publish volumes, Southern primes don't.
If "is 5× real?": "The signal isn't the multiple, it's where growth comes from: strike, interceptors, UGVs. Tell me which account looks wrong and I'll show you the source."`);
}

// 4 · Addressable vs forecastable
{
  const s = slide({ kicker: '1 · WHERE DOODLE FITS', title: 'Addressable is not forecastable', footer: 'Platforms × radios per platform × planning price · 2026–30 · range $18M–$648M' });
  const rows = [['~$215M', 'Addressable', 1.0, WHITE], ['~$156M', 'Open sockets', 0.73, WHITE], ['~$15M', 'On contracts visible today', 0.2, CORAL], ['?', 'Your forecast: needs your CRM and price book', 0, MUTED]];
  rows.forEach((r, i) => {
    const y = 1.6 + i * 0.78;
    text(s, r[0], { x: X, y, w: 1.9, h: 0.55, fontSize: 28, bold: true, color: r[3], valign: 'middle' });
    if (r[2] > 0) s.addShape(pres.shapes.RECTANGLE, { x: 2.8, y: y + 0.17, w: 3.2 * r[2], h: 0.2, fill: { color: i === 2 ? CORAL : LINE }, line: { type: 'none' } });
    text(s, r[1], { x: 6.25, y, w: 3.0, h: 0.55, fontSize: 12.5, color: i === 3 ? MUTED : BODY, valign: 'middle' });
  });
  s.addNotes(`~45 s. Four layers, never mixed.
~256k radio nodes; planning prices $450–1,500 by class, anchored on a public Mini OEM price — replaced by your price book on day one. The ranking of accounts barely moves with price.
Open sockets = remove platforms locked by a single incumbent (Milrem, Quantum Vector, Parrot's own radio, Tencore).
"Can we do $215M?" → "No. That's the pool. The number that matters is the bottom one, and I can't build it without your installed base and pricing."`);
}

// 5 · Fit
{
  const s = slide({ kicker: '1 · WHERE DOODLE FITS', title: 'Win as the independent, embedded link' });
  label(s, 'PLAY', { x: X, y: 1.75, w: 3.8, h: 0.25, color: CORAL });
  label(s, 'DON\'T PLAY', { x: X + 4.5, y: 1.75, w: 3.8, h: 0.25 });
  const play = ['Inside the OEM\'s platform', 'Strike, interceptors, UGVs', 'Private-label partners'];
  const dont = ['Handheld tactical radios', 'Primes with captive radios', 'A sovereignty pitch'];
  play.forEach((t, i) => text(s, t, { x: X, y: 2.15 + i * 0.62, w: 3.9, h: 0.5, fontSize: 16, color: WHITE }));
  dont.forEach((t, i) => text(s, t, { x: X + 4.5, y: 2.15 + i * 0.62, w: 3.9, h: 0.5, fontSize: 16, color: MUTED }));
  s.addNotes(`~35 s. Positioning, not a product pitch.
Independence: Silvus has been part of Motorola since Aug-2025. For a European OEM that's a supply-chain question — "is there a second source qualified?" Say it as the customer's question. Never FUD (David sold Silvus).
Captive: WB Group (Poland), Skyfall build their own radios. Sovereignty: Doodle is a US company — sell independence, form factor, supply resilience.`);
}

// 6 · Doors
{
  const s = slide({ kicker: '2 · NEW TERRITORY', title: 'Sell to a socket, cheapest door first', footer: 'Adding a fourth option to a shortlist is a smaller ask than displacing a sole source.' });
  const d = [['5', 'name Doodle on their site', 'Expand', CORAL], ['177', 'platforms publish no radio', 'Discover', WHITE], ['15', 'name Silvus publicly', 'Second source — last', WHITE]];
  d.forEach((v, i) => {
    label(s, `DOOR ${i + 1}`, { x: colX(i), y: 2.0, w: COL, h: 0.22 });
    text(s, v[0], { x: colX(i), y: 2.25, w: COL, h: 0.95, fontSize: 54, bold: true, color: v[3], valign: 'bottom' });
    text(s, v[1], { x: colX(i), y: 3.33, w: COL, h: 0.3, fontSize: 12 });
    text(s, v[2], { x: colX(i), y: 3.75, w: COL, h: 0.35, fontSize: 14, bold: true, color: WHITE });
  });
  s.addNotes(`~50 s. The buying unit is a radio socket inside a platform, and who signs it.
Door 1: Airvolute, Auterion, Evolve Dynamics, LUXUAV, UXV name Doodle on their own websites — NOT confirmed customers. Ask what they design next.
Door 2: no published radio ≠ open socket (chosen, NDA, incomplete site). Verify, don't assume.
Door 3: 15 companies name Silvus/StreamCaster publicly — never say they "run" it.
Proof = an eval kit and a number their engineer measured. Credibility = a local reference who holds the relationship — that's how the Portuguese grid deal happened.
Counts from my database, 840 companies, verified 28-Sep-2026.`);
}

// 7 · Coverage
{
  const s = slide({ kicker: '2 · NEW TERRITORY', title: 'Cover Europe by account value, not border', footer: 'One design authority, one owner.' });
  const tiers = [['CORE', 'France · Italy · Spain · Portugal', CORAL], ['STRATEGIC', 'Helsing · STARK · Rheinmetall · TYTAN', MUTED], ['PARTNER-LED', 'UXV · LUXUAV · Airvolute', MUTED], ['MONITOR', 'Poland · Nordics', MUTED]];
  tiers.forEach((t, i) => {
    const y = 1.75 + i * 0.62;
    label(s, t[0], { x: X, y: y + 0.04, w: 1.4, h: 0.3, color: t[2] });
    text(s, t[1], { x: X + 1.5, y, w: 4.3, h: 0.4, fontSize: 14, color: i === 0 ? WHITE : BODY });
  });
  stat(s, 6.75, 1.8, 2.5, '~68%', 'of evidenced value sits in German programmes', { size: 48, color: CORAL });
  s.addNotes(`~45 s. How Europe could be covered — not a claim on territory.
The radio is chosen where the platform is designed: STARK designs in Germany and builds in several countries; Tekever designs in Lisbon and builds in Swindon.
~68% German / ~17% French is where contracted volume is TODAY, partly disclosure bias — never "Germany is the best market".
France anchors the core: contracted volume (DELCO), Harmattan, Delair. Spain and Italy are prime-led build markets.
Poland: SAFE €43.7B, 89% to Polish industry, WB Group builds its own radios.
Andy: "so you want Germany?" → "No. The account should follow the design authority. Who owns Germany is your call; I'd want the rule clear."`);
}

// 8 · Qualify leads
{
  const s = slide({ kicker: '3 · LEADS', title: 'Qualify the socket, not the company', footer: 'No customer document, no stage change.' });
  const g = [['Platform', 'something\nthat moves'], ['Socket', 'open, taken\nor ours'], ['Band', 'overlaps\na SKU'], ['Authority', 'evaluator and\nsigner named'], ['Timing', 'a dated\ntrigger']];
  const cw = 1.5, gap = 0.25;
  g.forEach((v, i) => {
    const x = X + i * (cw + gap);
    text(s, String(i + 1), { x, y: 1.9, w: cw, h: 0.7, fontSize: 40, bold: true, color: i === 4 ? CORAL : LINE, valign: 'bottom' });
    text(s, v[0], { x, y: 2.75, w: cw, h: 0.35, fontSize: 15, bold: true, color: WHITE });
    text(s, v[1], { x, y: 3.15, w: cw, h: 0.6, fontSize: 12 });
  });
  s.addNotes(`~40 s. Five gates, in order — a lead fails at the first one it can't pass.
Timing is the hardest: I can date a buying trigger for only 10 of 840 companies (17 signals). My first scoring engine counted a named competitor as urgency twice; I rebuilt it. "The why-now comes from the call, not the crawl."
Stage evidence = a customer document: a BOM line, an NRE PO, or a signed evaluation agreement. Deal framing: MEDDPICC.`);
}

// 9 · Resellers
{
  const s = slide({ kicker: '3 · RESELLERS', title: 'In Europe the channel is private-label', sub: 'The deal is the radio inside their product, not a price list.' });
  stat(s, X, 2.2, 3.9, 'Solace', 'sells MultiLink', { size: 24 });
  stat(s, X + 4.5, 2.2, 4.0, 'Broadcast Solutions', 'sells meshLINK', { size: 24 });
  label(s, 'PARTNER TEST', { x: X, y: 3.95, w: W, h: 0.22, color: CORAL });
  text(s, 'Owns the customer   ·   Own RF engineering   ·   Named-account plan   ·   No export conflict', { x: X, y: 4.25, w: W, h: 0.35, fontSize: 13, color: WHITE });
  s.addNotes(`~40 s. "Leads and resellers are two different problems."
Partner types: stocking distributor, demand-generating reseller, engineering integrator. Resell vs white-label decides margin vs design-in.
Start with a bounded pilot: named accounts, deal registration, milestones. Exclusivity only if earned.
Export: one candidate distributor lists a Russia office — settled before the first channel call (only if asked).`);
}

// 10–12 · 30-60-90
function phase(idx, title, existing, fresh, gate, measure, notes, footer) {
  const s = slide({ kicker: '4 · 30-60-90', title, footer });
  ['30', '60', '90'].forEach((d, i) => {
    s.addShape(pres.shapes.OVAL, { x: 7.95 + i * 0.45, y: 0.44, w: 0.34, h: 0.34, fill: { color: i === idx ? CORAL : FAINT }, line: { type: 'none' } });
    text(s, d, { x: 7.95 + i * 0.45, y: 0.44, w: 0.34, h: 0.34, fontSize: 8.5, bold: true, color: i === idx ? BG : MUTED, align: 'center', valign: 'middle' });
  });
  [['EXISTING BUSINESS', existing, X], ['NEW BUSINESS', fresh, X + 4.5]].forEach(([l, items, x]) => {
    label(s, l, { x, y: 1.7, w: 3.8, h: 0.22 });
    items.forEach((t, i) => text(s, t, { x, y: 2.05 + i * 0.5, w: 3.9, h: 0.42, fontSize: 14, color: WHITE }));
  });
  label(s, 'GATE', { x: X, y: 3.85, w: 1, h: 0.22, color: CORAL });
  text(s, gate, { x: X, y: 4.12, w: 4.2, h: 0.4, fontSize: 16, bold: true, color: WHITE });
  label(s, 'MEASURED BY', { x: X + 4.5, y: 3.85, w: 2, h: 0.22 });
  text(s, measure, { x: X + 4.5, y: 4.15, w: 4.0, h: 0.4, fontSize: 12.5 });
  s.addNotes(notes);
}
phase(0, 'Days 1–30: know what we have',
  ['Meet every current customer', 'Ask what they design next', 'Confirm account owners'],
  ['Reconcile my map with your CRM', 'Agree 20 target accounts', 'First OEM meetings from week 2'],
  'Agreed baseline', 'All customers met · 20 targets agreed',
  `~35 s. The base is secured first, not instead — new business runs in parallel from week two.
"What are you designing next?" is where expansion and the next design-in come from.
Ownership: multi-site and cross-border accounts get a written rule before the first conflict.
Also agree stage definitions and SE capacity — "I'd ask you for the number rather than invent one." You're hiring a Sales Engineer in Germany.`);
phase(1, 'Days 31–60: evaluations in engineers\' hands',
  ['Expansion plan per top account', 'Fix what\'s at risk', 'Customers as references'],
  ['4–6 evaluations, criteria and dates', 'Sized to Sales Engineering capacity', '1–2 partner pilots'],
  'Evaluations live', 'Each with criteria, an owner and a date',
  `~35 s. A kit shipped is not progress; an evaluation with agreed criteria, owner and decision date is.
Why 4–6: one Sales Engineer can't run more strategic evaluations well; at the cap, the newest account waits in a qualified queue.
Existing users: Dual Radio (Sep-2026) is a concrete reason to talk — a hypothesis to test, not a pitch.`);
phase(2, 'Days 61–90: evidence becomes forecast',
  ['Expansion orders where timing allows', 'Renewals in the forecast', 'Top-account plans with Andy'],
  ['Design-in decisions where ready', 'Advance, nurture or stop: all 20', 'Partners: expand or end'],
  'Customer-backed forecast', '20 account decisions · every input labelled',
  `~40 s. Revenue in 90 days comes from existing customers; new logos produce decisions, not invoices.
Forecast = design-ins × platform volume × radios per platform × your price, each input tagged known or assumed, weighted by your stage benchmarks.
Design-in clock (my assumption, to calibrate with David): decision ~month 9, first production PO ~month 24.
Year-one number: "With your installed base and price book I'd commit to one in week four."
Line: "Thirty days buy information, sixty buy proof, ninety buy commitment."`,
  'No production wins promised in 90 days: the design-in cycle is longer.');

// 13 · Discussion
{
  const s = slide({ kicker: 'DISCUSSION', title: 'Three questions for you' });
  const q = [['GROWTH', 'New logos, expansion inside current platforms, or a step-up in regional revenue?'], ['OWNERSHIP', 'When an OEM sits across a border, who leads the design-in?'], ['SUCCESS', 'In twelve months, what would "clearly exceeded expectations" look like?']];
  q.forEach((v, i) => {
    const y = 1.8 + i * 0.9;
    label(s, v[0], { x: X, y: y + 0.06, w: 1.5, h: 0.25, color: CORAL });
    text(s, v[1], { x: X + 1.6, y, w: 6.9, h: 0.7, fontSize: 16, color: WHITE });
  });
  s.addNotes(`~30 s, then stop and listen. Growth → Amol. Ownership → Andy. Success → Andy.
If turned back: "From outside I'd guess expansion inside platforms you're already on, plus new logos where production is ramping — but you see the baseline, I don't."
Southern Europe, if asked: "I'd be very comfortable owning it and building it properly — France already has contracted volume. What I'd agree early is the rule for strategic accounts."
Keep in reserve: SE capacity in Europe; evaluation-to-LRIP path in the US (to David); customers and evaluations this role would inherit.`);
}

// Backup
{
  const s = slide({ kicker: 'BACKUP', title: 'How the numbers were built', footer: 'Public sources or labelled ASSUMED. Replaced by your data in month one.' });
  const r = [['ACCOUNTS', '840 companies · 2,781 matches · 153 band-verified fits'], ['SIGNALS', '5 name Doodle · 15 name Silvus · 177 publish no radio · 17 dated triggers'], ['MODEL', '20 accounts · 21 platforms · 53 evidence items, 39 from contracts'], ['PRICE', '$450–1,500 per radio by class — ASSUMED']];
  r.forEach((v, i) => {
    const y = 1.8 + i * 0.68;
    label(s, v[0], { x: X, y: y + 0.04, w: 1.5, h: 0.25 });
    text(s, v[1], { x: X + 1.6, y, w: 6.9, h: 0.45, fontSize: 13 });
  });
  s.addNotes(`Only if asked. The one-hour question (same answer as September, more precise):
"The first pass took about an hour — pulling European companies and matching platforms against your catalogue. That's discovery, not verification. Verifying bands, contacts and who names which radio took weeks, and I rebuilt the scoring once when it was double-counting."
Mention ≠ customer · no published radio ≠ open socket · band overlap ≠ validated integration.`);
}

pres.writeFile({ fileName: 'Doodle_Labs_Europe_Panel_FINAL_v5.pptx' }).then(f => console.log('wrote', f));
