const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Matelligence';
pptx.company = 'Matelligence';
pptx.subject = 'Matelligence Pitch Deck';
pptx.title = 'Matelligence';

const theme = {
  brand: '7C5CFF',
  cyan: '22D3EE',
  ok: '34D399',
  bg: '0B1020',
  panel: '0F1730',
  text: 'E9EEFC',
  muted: 'B7C2E3',
};

function addTopBar(slide, title) {
  slide.background = { color: theme.bg };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 0.9, fill: { color: theme.panel, transparency: 15 } });
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 0.23, w: 0.5, h: 0.5, fill: { color: theme.brand } });
  slide.addText('M', { x: 0.5, y: 0.23, w: 0.5, h: 0.5, color: 'FFFFFF', fontFace: 'Inter', fontSize: 18, bold: true, align: 'center', valign: 'mid' });
  slide.addText('Matelligence', { x: 1.1, y: 0.28, w: 6, h: 0.4, color: theme.text, fontFace: 'Inter', fontSize: 18, bold: true });
  slide.addText(title, { x: 7.3, y: 0.3, w: 5.5, h: 0.4, color: theme.muted, fontFace: 'Inter', fontSize: 14, align: 'right' });
}

function addFooter(slide, rightText) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.2, w: 13.33, h: 0.3, fill: { color: theme.panel, transparency: 35 } });
  slide.addText('© 2026 Matelligence', { x: 0.5, y: 7.2, w: 6, h: 0.3, color: theme.muted, fontFace: 'Inter', fontSize: 10 });
  slide.addText(rightText, { x: 7, y: 7.2, w: 5.83, h: 0.3, color: theme.muted, fontFace: 'Inter', fontSize: 10, align: 'right' });
}

function addTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.9, y: 2.0, w: 11.6, h: 1.0,
    color: theme.text, fontFace: 'Inter', fontSize: 44, bold: true,
  });
  slide.addText(subtitle, {
    x: 0.92, y: 3.1, w: 11.2, h: 0.9,
    color: theme.muted, fontFace: 'Inter', fontSize: 18,
  });
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.9, y: 4.2, w: 5.2, h: 0.6, fill: { color: theme.brand } });
  slide.addText('Material discovery • Evaluation • Sourcing confidence', { x: 0.9, y: 4.28, w: 5.2, h: 0.45, color: 'FFFFFF', fontFace: 'Inter', fontSize: 12, align: 'center', valign: 'mid' });
}

function addSectionTitle(slide, title, subtitle) {
  slide.addText(title, { x: 0.9, y: 1.4, w: 11.6, h: 0.8, color: theme.text, fontFace: 'Inter', fontSize: 34, bold: true });
  slide.addText(subtitle, { x: 0.9, y: 2.25, w: 11.2, h: 0.6, color: theme.muted, fontFace: 'Inter', fontSize: 16 });
}

function addCard(slide, x, y, w, h, title, body) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: '101A38', transparency: 10 }, line: { color: '2A345B', transparency: 30 } });
  slide.addText(title, { x: x + 0.3, y: y + 0.25, w: w - 0.6, h: 0.4, color: theme.text, fontFace: 'Inter', fontSize: 16, bold: true });
  slide.addText(body, { x: x + 0.3, y: y + 0.72, w: w - 0.6, h: h - 0.9, color: theme.muted, fontFace: 'Inter', fontSize: 13 });
}

// Slide 1: Title
{
  const slide = pptx.addSlide();
  slide.background = { color: theme.bg };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 7.5, fill: { color: theme.bg } });
  slide.addShape(pptx.ShapeType.ellipse, { x: -1.0, y: -1.2, w: 7.0, h: 4.5, fill: { color: theme.brand, transparency: 72 } });
  slide.addShape(pptx.ShapeType.ellipse, { x: 8.2, y: -0.6, w: 6.2, h: 4.0, fill: { color: theme.cyan, transparency: 78 } });
  slide.addShape(pptx.ShapeType.ellipse, { x: 4.0, y: 4.4, w: 7.0, h: 4.2, fill: { color: theme.ok, transparency: 86 } });
  addTitle(slide, 'Matelligence', 'Material intelligence that removes guesswork from sourcing and product development.');
  slide.addText('hello@matelligence.online', { x: 0.92, y: 5.0, w: 6.5, h: 0.4, color: theme.muted, fontFace: 'Inter', fontSize: 13 });
  slide.addText('May 2026', { x: 10.0, y: 6.9, w: 3.0, h: 0.4, color: theme.muted, fontFace: 'Inter', fontSize: 12, align: 'right' });
}

// Slide 2: Problem
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Problem');
  addSectionTitle(slide, 'Material decisions are slow, risky, and noisy.', 'Teams lose weeks in spreadsheets and supplier back-and-forth—then still test too many candidates.');

  addCard(slide, 0.9, 3.2, 6.1, 3.3, 'Too much guesswork', 'Specs vary by source. Assumptions get lost. Decisions become opinion-driven.');
  addCard(slide, 7.2, 3.2, 5.2, 3.3, 'Costly validation loops', 'Labs test broad shortlists. Procurement chases quotes late. Time-to-spec slips.');
  addFooter(slide, 'matelligence');
}

// Slide 3: Solution
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Solution');
  addSectionTitle(slide, 'Matelligence: discover, evaluate, validate.', 'A single workspace to shortlist materials with comparable specs, confidence notes, and test planning.');

  addCard(slide, 0.9, 3.1, 3.95, 3.4, 'Discover', 'Property-based search across classes, constraints, suppliers, and regions.');
  addCard(slide, 4.95, 3.1, 3.95, 3.4, 'Evaluate', 'Normalized comparisons, trade-offs, and assumption tracking—auditable by design.');
  addCard(slide, 9.0, 3.1, 3.4, 3.4, 'Validate', 'Test plan builder + RFQ-ready outputs so procurement and engineering stay aligned.');
  addFooter(slide, 'matelligence');
}

// Slide 4: Product
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Product');
  addSectionTitle(slide, 'Key capabilities', 'Built for product, R&D, and sourcing teams.');

  addCard(slide, 0.9, 3.05, 5.9, 1.75, 'Comparable material cards', 'A normalized view of properties with source links, ranges, and confidence notes.');
  addCard(slide, 0.9, 4.95, 5.9, 1.75, 'Decision log', 'Keep requirements, assumptions, evidence, and rationale together—shareable and auditable.');

  slide.addShape(pptx.ShapeType.roundRect, { x: 7.2, y: 3.05, w: 5.2, h: 3.65, fill: { color: '101A38', transparency: 10 }, line: { color: '2A345B', transparency: 30 } });
  slide.addText('Example evaluation output', { x: 7.5, y: 3.25, w: 4.6, h: 0.35, color: theme.text, fontFace: 'Inter', fontSize: 14, bold: true });
  slide.addText('• Strength-to-weight: A−\n• Heat resistance: B+\n• Supplier risk: Low\n• CO₂ footprint: −18%\n• Confidence score: 0.74', { x: 7.5, y: 3.72, w: 4.6, h: 2.6, color: theme.muted, fontFace: 'Inter', fontSize: 13 });

  addFooter(slide, 'matelligence');
}

// Slide 5: Market
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Market');
  addSectionTitle(slide, 'A large and growing need', 'Every physical product depends on better materials decisions.');

  addCard(slide, 0.9, 3.1, 6.1, 3.4, 'Target users', 'Product development, materials engineering, R&D, procurement, sustainability and compliance teams.');
  addCard(slide, 7.2, 3.1, 5.2, 3.4, 'Initial focus', 'SMEs and mid-market manufacturers and product companies seeking faster iteration and lower sourcing risk.');
  addFooter(slide, 'matelligence');
}

// Slide 6: Business model
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Business Model');
  addSectionTitle(slide, 'SaaS pricing that scales with value', 'Start with team workflows; expand into enterprise integrations.');

  addCard(slide, 0.9, 3.1, 3.95, 3.4, 'Starter', '$0 for exploration: basic discovery + limited comparisons.');
  addCard(slide, 4.95, 3.1, 3.95, 3.4, 'Team', '$299/mo: unlimited comparisons, test planning, RFQ export, shared decision log.');
  addCard(slide, 9.0, 3.1, 3.4, 3.4, 'Enterprise', 'Custom: SSO, roles, compliance workflows, supplier integrations, dedicated support.');
  addFooter(slide, 'matelligence');
}

// Slide 7: Go-to-market
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Go-to-Market');
  addSectionTitle(slide, 'Land with a workflow, expand to a platform', 'Services-led onboarding → productized templates → enterprise integrations.');

  addCard(slide, 0.9, 3.1, 6.1, 3.4, 'Acquisition channels', '• Industry partnerships\n• Design/manufacturing communities\n• Targeted outbound to product + sourcing leaders\n• Content: materials playbooks and templates');
  addCard(slide, 7.2, 3.1, 5.2, 3.4, 'Expansion', 'Add internal test result ingestion, supplier catalogs, compliance workflows, and team analytics.');
  addFooter(slide, 'matelligence');
}

// Slide 8: Traction / next steps
{
  const slide = pptx.addSlide();
  addTopBar(slide, 'Next Steps');
  addSectionTitle(slide, 'What we’re building next', 'We’re looking for design partners to shape the product.');

  addCard(slide, 0.9, 3.1, 6.1, 3.4, 'Design partners', 'Teams with active material sourcing challenges willing to share requirements and pilot workflows.');
  addCard(slide, 7.2, 3.1, 5.2, 3.4, 'Call to action', 'Request a demo and bring a real product requirement.\n\nhello@matelligence.online');
  addFooter(slide, 'matelligence');
}

pptx.writeFile({ fileName: 'Matelligence_Pitch_Deck.pptx' });
console.log('Wrote Matelligence_Pitch_Deck.pptx');

