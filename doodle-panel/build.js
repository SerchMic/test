const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625
pres.title = 'How I would build Doodle Labs Europe';
pres.author = 'Sergio Benavidez';

const BG = '0B0D10', CARD = '171B20', LINE = '2A3139';
const WHITE = 'F2F4F6', BODY = 'C9CED4', MUTED = '8E98A3', CORAL = 'F05A55';
const F = 'Segoe UI', MONO = 'Consolas';
const X = 0.7, W = 8.6;

function base(kicker, title, footer) {
  const s = pres.addSlide();
  s.background = { color: BG };
  if (kicker) s.addText(kicker, { x: X, y: 0.42, w: W, h: 0.25, fontFace: MONO, fontSize: 9.5, color: MUTED, charSpacing: 1, margin: 0, isTextBox: true });
  if (title) s.addText(title, { x: X, y: 0.72, w: W, h: 0.62, fontFace: F, fontSize: 22, bold: true, color: WHITE, margin: 0, valign: 'top', isTextBox: true });
  if (footer) s.addText(footer, { x: X, y: 5.02, w: W, h: 0.28, fontFace: F, fontSize: 9.5, color: CORAL, margin: 0, isTextBox: true });
  return s;
}
function card(s, x, y, w, h) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: CARD }, line: { color: LINE, width: 0.75 } });
}
function stat(s, x, y, w, big, label, color, size) {
  s.addText(big, { x, y, w, h: 0.75, fontFace: F, fontSize: size || 40, bold: true, color: color || WHITE, margin: 0, isTextBox: true });
  s.addText(label, { x, y: y + 0.8, w, h: 0.9, fontFace: F, fontSize: 12, color: BODY, margin: 0, valign: 'top', isTextBox: true });
}

// ---------- 1 · Title ----------
{
  const s = base('DOODLE LABS · DIRECTOR OF SALES, EUROPE · SERGIO BENAVIDEZ', null, null);
  s.addText('How I would build\nDoodle Labs Europe', { x: X, y: 1.0, w: W, h: 1.5, fontFace: F, fontSize: 36, bold: true, color: WHITE, margin: 0, valign: 'top', isTextBox: true });
  s.addText("Europe's unmanned production is scaling. The radio is chosen once, at design freeze. I'd win it account by account.", { x: X, y: 2.6, w: 7.2, h: 0.75, fontFace: F, fontSize: 14, color: BODY, margin: 0, valign: 'top', isTextBox: true });
  const topics = ['UxV market & fit', 'New territory', 'Leads & resellers', '30-60-90'];
  topics.forEach((t, i) => {
    const x = X + i * 2.15;
    s.addText(String(i + 1), { x, y: 3.85, w: 0.35, h: 0.35, fontFace: F, fontSize: 16, bold: true, color: CORAL, margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.35, y: 3.85, w: 1.75, h: 0.35, fontFace: F, fontSize: 12, color: WHITE, margin: 0, isTextBox: true });
  });
  s.addText('Outside-in view from public sources — a hypothesis to test against your CRM, not your forecast.', { x: X, y: 5.02, w: W, h: 0.28, fontFace: F, fontSize: 9.5, color: MUTED, margin: 0, isTextBox: true });
  s.addNotes(`~35 s. FIRST: "Before I start — what's the real end time? I'll cut this to fit." Then: "You asked for four things; the deck follows exactly those four, in about eight minutes."
One-line thesis: production is scaling, the radio is chosen once at design freeze, and I'd win it account by account.
Everything comes from public sources. It's a hypothesis until it meets your CRM.`);
}

// ---------- 2 · The seat ----------
{
  const s = base('WHO IS PRESENTING', "I've sat on the OEM side of this sale", 'The gap, stated upfront: I have not sold a radio as the component vendor. I know the seat your customers sit in.');
  const cols = [
    ['€3.5M', "Multi-year programme with Portugal's grid operator. Opened from zero, through a partner who held the relationship."],
    ['2', 'Regions opened from zero — Iberia and LATAM — as the sole commercial resource, including a BVLOS permit over critical infrastructure.'],
    ['90+', 'Distribution and integration partners built at an industrial UAV OEM, where I integrated third-party subsystems.'],
  ];
  cols.forEach((c, i) => stat(s, X + i * 2.95, 1.75, 2.6, c[0], c[1], i === 0 ? CORAL : WHITE));
  s.addNotes(`~45 s. They've read the CV — short and factual.
€3.5M (Airpelago): an Estonian data partner was leaving data collection with operations to finish. I built the numbers that made handing us that volume rational, took over their kilometres, rebuilt the client's inspection procedures, got the BVLOS permit, won the client's tender. [SAY THE SAME MONTH AS YOUR CV — CV says "within six months".] Contract value over [N] years; [who] signed.
90+ partners (Beyond Vision, 2023-25): integrated payloads, tethered power and third-party subsystems into our own UAVs and sold the combined system. NOT datalink selections — never say you chose radios.
If asked "you've never sold a radio?": "True. I've been the OEM a component vendor had to win — I know that integration costs the platform builder, that the engineer specifies and someone else signs, and that the socket closes at design freeze. Product depth is what the Sales Engineer partnership is for."`);
}

// ---------- 3 · Market ----------
{
  const s = base('1 · UXV MARKET', 'Europe is buying unmanned systems at scale — and the money lands with the OEM', 'Budgets become platforms; every platform needs a link. I call the OEM, not the ministry.');
  const cols = [
    ['€540M', 'Bundeswehr loitering munitions, Helsing and STARK. Approved Feb 2026.'],
    ['6,000', 'Soldier drones ordered by France (DGA, DELCO programme), 2026.'],
    ['5', 'EU defence projects approved 28 Sep 2026, including drones and counter-drone.'],
  ];
  cols.forEach((c, i) => stat(s, X + i * 2.95, 1.75, 2.6, c[0], c[1]));
  s.addText('Air, ground and sea: every scaled platform adds range, EW exposure and nodes to a network.', { x: X, y: 3.95, w: W, h: 0.4, fontFace: F, fontSize: 12, italic: true, color: MUTED, margin: 0, isTextBox: true });
  s.addNotes(`~45 s. Three facts, one takeaway.
€540M: Bundestag approval 25-Feb-2026, Helsing HX-2 and STARK Virtus; options beyond that.
6,000: DGA, 1,000 + 5,000 DELCO (Harmattan), Jun-2026.
5 EDPCI: EU defence ministers, 28-Sep-2026, incl. drones and counter-drones (Eastern Flank Watch context).
Trend line: autonomy raises the comms bar, not lowers it — more range, more nodes, more jamming.
If fibre comes up: fibre keeps the disposable short-range tier; RF mesh keeps reusable, multi-node, ISR, C2 and UGV. "How do you want the team to answer fibre? I'd rather use your line."`);
}

// ---------- 4 · Production chart ----------
{
  const s = base('1 · UXV MARKET', 'Output at the top 20 European OEMs could grow ~5× by 2030', 'External model, base case — not a Doodle Labs forecast. ~13k of 2026 units rest on visible contracts; the rest is ramp.');
  s.addChart(pres.charts.BAR, [
    { name: '2025', labels: ['Strike / loitering', 'ISR', 'UGV', 'Interceptors'], values: [3000, 3550, 3300, 0] },
    { name: '2030 (base)', labels: ['Strike / loitering', 'ISR', 'UGV', 'Interceptors'], values: [21500, 12000, 8450, 8000] },
  ], {
    x: X, y: 1.45, w: 5.6, h: 3.4, barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 60,
    chartColors: [MUTED, CORAL], showLegend: true, legendPos: 't', legendColor: BODY, legendFontFace: F, legendFontSize: 10,
    showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: BODY, dataLabelFontSize: 9, dataLabelFontFace: F, dataLabelFormatCode: '#,##0',
    catAxisLabelColor: BODY, catAxisLabelFontFace: F, catAxisLabelFontSize: 10, catAxisLineShow: false,
    valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  });
  stat(s, 6.75, 1.75, 2.55, '~10k → 50k', 'platforms a year, 2025 → 2030, across 20 accounts with published volumes.', WHITE, 26);
  s.addText('Growth comes from new volume — strike, interceptors, UGVs — where the radio socket is least likely to be locked.', { x: 6.75, y: 3.35, w: 2.55, h: 1.3, fontFace: F, fontSize: 12, color: CORAL, margin: 0, valign: 'top', isTextBox: true });
  s.addNotes(`~50 s. Units, not money.
2025: 9,850 platforms; 2030 base: 49,950; upside ~112k. The conservative line falls after 2027 on purpose: only ~13k of 2026 units are contracted, the rest is options and announced ramps.
Model: 20 accounts, 21 platforms, 53 evidence items (39 grade A — contracts, delivery rates).
Honest limit: German start-ups publish volumes and Southern primes don't — disclosure bias.
If "is 5× real?": "The signal isn't the multiple, it's where growth comes from. Tell me which account looks wrong and I'll show you the source."`);
}

// ---------- 5 · Addressable vs forecastable ----------
{
  const s = base('1 · WHERE DOODLE FITS', 'Addressable is not forecastable', 'Platforms × radios per platform × planning price. 2026–30, base case. Range $18M–$648M.');
  const rows = [
    ['~$215M', 'Addressable', 'Radio nodes those platforms need, where a Mesh Rider SKU fits', WHITE, 8.6],
    ['~$156M', 'Open sockets', 'Not visibly locked by a single incumbent', WHITE, 6.6],
    ['~$15M', 'Contracted today', 'Rests on contracts I can see — the design window is now', CORAL, 4.6],
    ['?', 'Your forecast', 'Needs your CRM, installed base and price book', MUTED, 2.6],
  ];
  rows.forEach((r, i) => {
    const y = 1.55 + i * 0.83;
    s.addShape(pres.shapes.RECTANGLE, { x: X, y, w: r[4] * 0.42, h: 0.66, fill: { color: i === 2 ? '3A1F21' : CARD }, line: { color: i === 2 ? CORAL : LINE, width: 0.75 } });
    s.addText(r[0], { x: X + 0.15, y, w: 1.6, h: 0.66, fontFace: F, fontSize: 20, bold: true, color: r[3], margin: 0, valign: 'middle', isTextBox: true });
    s.addText([{ text: r[1], options: { bold: true, color: WHITE, breakLine: true } }, { text: r[2], options: { color: BODY } }], { x: 4.6, y, w: 4.7, h: 0.66, fontFace: F, fontSize: 11, margin: 0, valign: 'middle', isTextBox: true });
  });
  s.addNotes(`~50 s. The slide that matters: four layers, never mixed.
~256k radio nodes over 2026–30; planning prices $450–1,500 by class (anchored on a public Mini OEM price) — to be replaced by your price book on day one. Ranking of accounts barely moves with price; it's driven by volume and socket status.
Open sockets = remove platforms locked by a single incumbent (e.g. Milrem, Quantum Vector, Parrot own radio, Tencore).
If "can we do $215M?": "No. That's the pool. The number that matters is the bottom one, and I can't build it without your installed base and pricing."
If "how do you go from addressable to forecast?": socket open + customer evidence at a stage + funded platform, then weighted by your stage benchmarks.`);
}

// ---------- 6 · Where Doodle fits ----------
{
  const s = base('1 · WHERE DOODLE FITS', 'Doodle wins as the independent embedded link', '"We are not European. We are independent." — the pitch Doodle can make and defend.');
  card(s, X, 1.55, 4.15, 2.9); card(s, X + 4.45, 1.55, 4.15, 2.9);
  s.addText('PLAY', { x: X + 0.25, y: 1.75, w: 3.7, h: 0.3, fontFace: MONO, fontSize: 10, color: CORAL, margin: 0, isTextBox: true });
  s.addText('DON\'T PLAY', { x: X + 4.7, y: 1.75, w: 3.7, h: 0.3, fontFace: MONO, fontSize: 10, color: MUTED, margin: 0, isTextBox: true });
  const play = ['Inside someone else\'s platform, chosen at design freeze', 'New-volume sockets: strike, interceptors, UGVs', 'Private-label partners who put the radio in their product'];
  const dont = ['The tactical radio in a soldier\'s hand', 'Captive architectures — primes that build their own radios', 'European sovereignty — we can\'t sell a flag'];
  const list = (arr, x, color) => s.addText(arr.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < arr.length - 1 } })), { x, y: 2.15, w: 3.7, h: 2.1, fontFace: F, fontSize: 12.5, color, paraSpaceAfter: 8, margin: 0, valign: 'top', isTextBox: true });
  list(play, X + 0.25, WHITE); list(dont, X + 4.7, BODY);
  s.addNotes(`~40 s. Positioning, not a product pitch — they know the product.
Independence: Silvus has been part of Motorola since Aug-2025. For a European OEM that's a supply-chain question: "is there a second source qualified?" Say it as the customer's question. Never FUD.
Captive: WB Group (Poland) and Skyfall build their own radios.
Supply resilience: US and European production sites under evaluation (Defense News, Jul-2026) — only if you have the source to hand.`);
}

// ---------- 7 · New territory ----------
{
  const s = base('2 · NEW TERRITORY', 'Sell to a socket, through the cheapest door first', 'Adding a fourth option to a shortlist is a smaller ask than displacing a sole source. Displacement comes last.');
  const doors = [
    ['5', 'companies name Doodle on their own site', 'Expand: ask what they design next', CORAL],
    ['177', 'platforms publish no radio', 'Discover: is the decision still open?', WHITE],
    ['15', 'companies name Silvus publicly', 'Second source: once we have proof', WHITE],
  ];
  doors.forEach((d, i) => {
    const x = X + i * 2.95;
    card(s, x, 1.55, 2.7, 2.55);
    s.addText(`DOOR ${i + 1}`, { x: x + 0.2, y: 1.7, w: 2.3, h: 0.25, fontFace: MONO, fontSize: 9.5, color: MUTED, margin: 0, isTextBox: true });
    s.addText(d[0], { x: x + 0.2, y: 1.95, w: 2.3, h: 0.7, fontFace: F, fontSize: 36, bold: true, color: d[3], margin: 0, isTextBox: true });
    s.addText(d[1], { x: x + 0.2, y: 2.7, w: 2.3, h: 0.55, fontFace: F, fontSize: 11.5, color: BODY, margin: 0, valign: 'top', isTextBox: true });
    s.addText(d[2], { x: x + 0.2, y: 3.35, w: 2.3, h: 0.6, fontFace: F, fontSize: 11.5, bold: true, color: WHITE, margin: 0, valign: 'top', isTextBox: true });
  });
  s.addText('Proof = an eval kit in an engineer\'s hands and a number he measured. Credibility = a local reference who already holds the relationship.', { x: X, y: 4.3, w: W, h: 0.5, fontFace: F, fontSize: 11.5, italic: true, color: MUTED, margin: 0, isTextBox: true });
  s.addNotes(`~55 s. Method first — method says repeatable, a story alone says lucky.
Buying unit = a radio socket inside a platform, and who signs it.
Door 1: Airvolute (SK), Auterion (DE/CH), Evolve Dynamics (UK), LUXUAV (LU), UXV Technologies (DK) name Doodle on their own websites — NOT confirmed customers. Say "companies that name Doodle", never "your customers". UXV's public partnership proves the reference model.
Door 2: no published radio ≠ open socket. Could be chosen, NDA, or an incomplete website. Verify, don't assume.
Door 3: 15 companies (17 rows) name Silvus/StreamCaster publicly. Never say they "run" it.
Local reference = how the Portuguese grid deal happened.
Counts from my own database, 840 companies, verified 28-Sep-2026.`);
}

// ---------- 8 · Coverage ----------
{
  const s = base('2 · NEW TERRITORY', 'Cover Europe by account value, not by border', 'One design authority, one owner. Factories and regional support share the credit.');
  const tiers = [
    ['CORE TERRITORY', 'France · Italy · Spain · Portugal', 'France anchors it: contracted volume today'],
    ['STRATEGIC ACCOUNTS', 'Helsing · STARK · Rheinmetall · TYTAN · Quantum', 'Owned where the radio is specified'],
    ['PARTNER-LED', 'UXV · LUXUAV · Airvolute · private-label', 'Win the radio inside their product'],
    ['MONITOR', 'Poland · Nordics', 'Open when a non-captive OEM appears'],
  ];
  tiers.forEach((t, i) => {
    const y = 1.5 + i * 0.8;
    card(s, X, y, 5.7, 0.68);
    s.addText(t[0], { x: X + 0.2, y, w: 1.9, h: 0.68, fontFace: MONO, fontSize: 9, color: i === 0 ? CORAL : MUTED, margin: 0, valign: 'middle', isTextBox: true });
    s.addText([{ text: t[1], options: { bold: true, color: WHITE, breakLine: true } }, { text: t[2], options: { color: BODY } }], { x: X + 2.1, y, w: 3.5, h: 0.68, fontFace: F, fontSize: 10.5, margin: 0, valign: 'middle', isTextBox: true });
  });
  stat(s, 6.75, 1.55, 2.55, '~68%', 'of evidenced value sits in German programmes; ~17% in French.', CORAL);
  s.addText('The radio decision sits with the design authority — which doesn\'t always match where a platform is built.', { x: 6.75, y: 3.2, w: 2.55, h: 1.3, fontFace: F, fontSize: 11.5, color: BODY, margin: 0, valign: 'top', isTextBox: true });
  s.addNotes(`~45 s. How Europe could be covered — not a claim on territory.
Design authority examples: STARK designs in Germany and builds in several countries; Tekever designs in Lisbon and builds in Swindon. Geography alone splits one account between several owners.
~68% German is where the contracted volume is TODAY, partly disclosure bias — never "Germany is the best market".
Poland: SAFE €43.7B, 89% to Polish industry, and WB Group builds its own radios.
If Andy: "so you want Germany?" → "No. I'm saying the account should follow the design authority. Who owns Germany is your call; I'd want the rule to be clear."
If "the role is only the South": "Then my job is to make the South worth more than it looks — France is the anchor; Iberia and Italy are build markets through the primes."`);
}

// ---------- 9 · Qualify ----------
{
  const s = base('3 · LEADS & RESELLERS', 'Qualify the socket, then qualify the partner', 'No customer document, no stage change: a BOM line, an NRE PO, or a signed evaluation agreement.');
  card(s, X, 1.5, 4.15, 3.25); card(s, X + 4.45, 1.5, 4.15, 3.25);
  s.addText('LEADS — FIVE GATES, IN ORDER', { x: X + 0.25, y: 1.68, w: 3.7, h: 0.3, fontFace: MONO, fontSize: 9.5, color: CORAL, margin: 0, isTextBox: true });
  const gates = [['Platform', 'builds something that moves'], ['Socket', 'open, taken, or ours'], ['Band', 'overlaps a Mesh Rider SKU'], ['Authority', 'evaluator and signer named'], ['Timing', 'a dated trigger, not a guess']];
  gates.forEach((g, i) => {
    const y = 2.1 + i * 0.5;
    s.addText(String(i + 1), { x: X + 0.25, y, w: 0.3, h: 0.4, fontFace: F, fontSize: 14, bold: true, color: CORAL, margin: 0, valign: 'middle', isTextBox: true });
    s.addText([{ text: g[0] + '  ', options: { bold: true, color: WHITE } }, { text: g[1], options: { color: BODY } }], { x: X + 0.6, y, w: 3.4, h: 0.4, fontFace: F, fontSize: 12, margin: 0, valign: 'middle', isTextBox: true });
  });
  s.addText('PARTNERS — THE CHANNEL IS PRIVATE-LABEL', { x: X + 4.7, y: 1.68, w: 3.7, h: 0.3, fontFace: MONO, fontSize: 9.5, color: CORAL, margin: 0, isTextBox: true });
  s.addText('Solace sells MultiLink. Broadcast Solutions sells meshLINK. The deal is the radio inside their product, not a price list.', { x: X + 4.7, y: 2.1, w: 3.7, h: 0.9, fontFace: F, fontSize: 12, color: WHITE, margin: 0, valign: 'top', isTextBox: true });
  const test = ['Owns the customer', 'Has its own RF engineering', 'Commits a plan with named accounts', 'No export-control conflict'];
  s.addText(test.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < test.length - 1 } })), { x: X + 4.7, y: 3.1, w: 3.7, h: 1.5, fontFace: F, fontSize: 12, color: BODY, paraSpaceAfter: 4, margin: 0, valign: 'top', isTextBox: true });
  s.addNotes(`~60 s. "Those are two different problems, so let me take them separately."
Data behind gate 5: 153 of 2,781 platform matches publish a band; I can date a buying trigger for 17 signals across only 10 of 840 companies. My first scoring engine counted a named competitor as urgency twice — I rebuilt it. "The why-now has to come from the call, not the crawl."
Partner types: stocking distributor vs demand-generating reseller vs engineering integrator. Resell vs white-label decides margin vs design-in.
Start with a bounded pilot: named accounts, deal registration, milestones. Exclusivity only if earned.
Only if asked: one candidate distributor lists a Russia office — settled before the first channel call.
Deal framing: MEDDPICC.`);
}

// ---------- 10 · 30-60-90 ----------
{
  const s = base('4 · 30-60-90 · EXISTING + NEW BUSINESS', '30 days information · 60 proof · 90 commitment', "I won't promise production wins in 90 days. I'll promise you know exactly where every account stands.");
  const cols = [
    ['DAYS 1–30 · INFORMATION', ['Reconcile my map with your CRM, installed base and account ownership', 'Meet every existing customer: what are you designing next?', 'Agree a written qualification standard'], 'Agreed baseline'],
    ['DAYS 31–60 · PROOF', ['Run 4–6 customer evaluations, sized to SE capacity', 'Open 1–2 partners in bounded pilots', 'Measure conversion by account and motion'], 'Evaluations live, with criteria and dates'],
    ['DAYS 61–90 · COMMITMENT', ['Every priority account: advance, nurture or stop', 'Design-in decisions where the customer\'s clock allows', 'Bottom-up forecast, every input labelled'], 'Customer-backed forecast'],
  ];
  cols.forEach((c, i) => {
    const x = X + i * 2.95;
    card(s, x, 1.5, 2.7, 3.3);
    s.addText(c[0], { x: x + 0.2, y: 1.65, w: 2.35, h: 0.3, fontFace: MONO, fontSize: 9, color: CORAL, margin: 0, isTextBox: true });
    s.addText(c[1].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < c[1].length - 1 } })), { x: x + 0.2, y: 2.05, w: 2.35, h: 2.0, fontFace: F, fontSize: 11, color: BODY, paraSpaceAfter: 6, margin: 0, valign: 'top', isTextBox: true });
    s.addText([{ text: 'GATE  ', options: { fontFace: MONO, color: MUTED, fontSize: 9 } }, { text: c[2], options: { bold: true, color: WHITE } }], { x: x + 0.2, y: 4.15, w: 2.35, h: 0.5, fontFace: F, fontSize: 11, margin: 0, valign: 'middle', isTextBox: true });
  });
  s.addNotes(`~60 s. One slide, three gates.
Existing business: secure the base first — every current account gets a meeting and one question: what are you designing next? New business runs in parallel from week two.
4–6 evaluations "sized to your SE capacity, which I'd ask you for rather than invent" (you're hiring a Sales Engineer in Germany — ask how capacity is shared until then).
Design-in clock (my assumption, to calibrate with David): decision ~month 9, certification behind it, first production PO ~month 24. That's why day 90 is about decisions, not revenue.
Forecast = design-ins × platform volume × attach rate × your price, each input tagged known / assumed.
Year-one number: "With your installed base and price book I'd commit to a number in week four."`);
}

// ---------- 11 · Discussion ----------
{
  const s = base('DISCUSSION', 'Three questions for you', null);
  const qs = [
    ['GROWTH', 'Is Europe mainly a new-logo build, expansion inside platforms you\'re already on, or a real step-up in regional revenue?'],
    ['OWNERSHIP', 'When a high-value OEM sits across a border, does ownership follow geography, relationships, or whoever is best placed to win the design-in?'],
    ['SUCCESS', 'Twelve months from now, what would make you say this role clearly exceeded expectations?'],
  ];
  qs.forEach((q, i) => {
    const y = 1.5 + i * 0.95;
    s.addText(q[0], { x: X, y, w: 1.6, h: 0.75, fontFace: MONO, fontSize: 10, color: CORAL, margin: 0, valign: 'top', isTextBox: true });
    s.addText(q[1], { x: X + 1.7, y, w: 6.9, h: 0.75, fontFace: F, fontSize: 13, color: WHITE, margin: 0, valign: 'top', isTextBox: true });
  });
  s.addText('Give me the account list and 90 days — you\'ll have a forecast where every number has a source.', { x: X, y: 4.55, w: W, h: 0.5, fontFace: F, fontSize: 14, bold: true, color: WHITE, margin: 0, isTextBox: true });
  s.addNotes(`~30 s, then stop and listen.
Growth → Amol. Ownership → Andy. Success → Andy.
Have your own answer if they turn it back: "My guess from outside is expansion inside platforms you're already on, plus a few new logos where production is ramping — but you see the baseline, I don't."
Southern Europe, if asked: "I'd be very comfortable owning Southern Europe and building it properly — France already has contracted volume. What I'd like to agree early is the rule for strategic accounts."
Follow-ups to keep: SE capacity in Europe; the typical evaluation-to-LRIP path in the US (to David); existing European customers and evaluations this role would inherit.`);
}

// ---------- 12 · Backup ----------
{
  const s = base('BACKUP · OPEN ONLY IF ASKED', 'How the numbers were built', 'Every input is public or labelled ASSUMED. Any number here gets replaced by your internal data in month one.');
  const rows = [
    ['Account map', '840 companies read at source → 2,781 platform–SKU matches → 153 with a published band that overlaps a SKU (37 companies)'],
    ['Signals', '5 name Doodle · 15 name Silvus · 177 platforms publish no radio · 17 dated triggers in 10 companies'],
    ['Production model', '20 accounts, 21 platforms, 53 evidence items (39 from contracts or delivery rates). Base, conservative and upside cases'],
    ['Revenue bridge', 'Units × radios per platform (1.0–2.0) × planning price $450–1,500 by class — ASSUMED, to be replaced by your price book'],
    ['Market facts', 'Bundestag €540M (Feb 2026) · DGA 6,000 DELCO (Jun 2026) · EU EDPCI approvals (28 Sep 2026) · SAFE Poland €43.7B, 89% domestic'],
  ];
  rows.forEach((r, i) => {
    const y = 1.5 + i * 0.66;
    s.addText(r[0], { x: X, y, w: 1.9, h: 0.58, fontFace: MONO, fontSize: 9.5, color: CORAL, margin: 0, valign: 'top', isTextBox: true });
    s.addText(r[1], { x: X + 2.0, y, w: 6.6, h: 0.58, fontFace: F, fontSize: 11, color: BODY, margin: 0, valign: 'top', isTextBox: true });
  });
  s.addNotes(`Only if asked "where do these numbers come from?" or "how did you do it in an hour?".
One-hour answer (same as September, more precise): "The first pass took about an hour — pulling European companies and matching platforms against your catalogue. That's discovery, not verification. Verifying bands, contacts and who names which radio took weeks, and I rebuilt the scoring once when I found it double-counting."
Mention ≠ customer. No published radio ≠ open socket. Band overlap ≠ validated integration.`);
}

pres.writeFile({ fileName: 'Doodle_Labs_Europe_Panel_FINAL_v4.pptx' }).then(f => console.log('wrote', f));
