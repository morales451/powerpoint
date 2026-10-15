const fs = require("fs");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/a9f08bbb-9e1f-48e1-b895-cf1232952f3d_c30e392a-4b21-4f80-b4de-330c3c8bc1fd/pptx/scripts/apply_theme.js");

// Henry brand palette (henry.com)
const THEME = {
  name: "Henry Tier 1 Retention",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "333132", lt1: "FFFFFF", dk2: "00306C", lt2: "F1F1F2",
    accent1: "0054A6", accent2: "0090D2", accent3: "C0392B",
    accent4: "2E7D4F", accent5: "E0A526", accent6: "D9682A",
    hlink: "0054A6", folHlink: "00306C",
  },
};
const GRAY = "BCBEC0", PALE = "C9DCEB", MUTED = "5F5D5E";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Tier 1 Roofer Retention";
pres.company = "Henry Company";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

const S = {
  G: { label: "Strength", fill: C.accent4, text: C.background1 },
  Y: { label: "Watch", fill: C.accent5, text: C.text1 },
  O: { label: "At Risk", fill: C.accent6, text: C.background1 },
  R: { label: "Critical", fill: C.accent3, text: C.background1 },
};
const png = (buf) => "image/png;base64," + buf.toString("base64");
const LOGO_W = png(fs.readFileSync("henry-logo.png"));
const LOGO_B = png(fs.readFileSync("henry-blue-logo.png"));

// ---------- Layouts ----------
const FOOT = "Confidential – Internal Use Only";
pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.text2 },
  objects: [
    { image: { x: 0.6, y: 0.55, w: 1.32, h: 0.5, data: LOGO_W } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.9, w: 8.8, h: 0.9, fontSize: 40, bold: true, color: C.background1, margin: 0, valign: "top", align: "left" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 2.8, w: 8.8, h: 0.5, fontSize: 20, color: PALE, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "date", type: "body", x: 0.6, y: 4.6, w: 6, h: 0.35, fontSize: 14, color: PALE, margin: 0, valign: "top" }, text: "" } },
  ],
});
const contentObjects = (titleW) => [
  { image: { x: 0.6, y: 5.12, w: 0.6, h: 0.23, data: LOGO_B } },
  { text: { text: FOOT, options: { x: 3.5, y: 5.12, w: 3, h: 0.23, fontSize: 9, color: MUTED, align: "center", valign: "middle", margin: 0 } } },
  { line: { x: 0.6, y: 5.0, w: 8.8, h: 0, line: { color: GRAY, width: 0.75 } } },
  { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: titleW, h: 0.6, fontSize: 26, bold: true, color: C.text2, margin: 0, valign: "top", align: "left" }, text: "" } },
];
pres.defineSlideMaster({
  title: "CONTENT", background: { color: C.background1 },
  slideNumber: { x: 9.0, y: 5.12, w: 0.4, h: 0.23, fontSize: 10, color: MUTED, align: "right" },
  objects: contentObjects(8.8),
});
pres.defineSlideMaster({
  title: "ISSUE", background: { color: C.background1 },
  slideNumber: { x: 9.0, y: 5.12, w: 0.4, h: 0.23, fontSize: 10, color: MUTED, align: "right" },
  objects: contentObjects(6.9),
});

// ---------- Helpers ----------
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, valign: "top" }, o));
}
function chip(slide, key, x, y, w = 0.85) {
  const s = S[key];
  slide.addText(s.label, { shape: pres.shapes.RECTANGLE, x, y, w, h: 0.26, fill: { color: s.fill }, line: { color: s.fill }, color: s.text, fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0, objectName: "Status " + s.label });
}
function dot(slide, key, x, y, d = 0.12) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: S[key].fill }, line: { color: S[key].fill } });
}
function bullets(slide, items, o) {
  slide.addText(items.map((t, i) => ({ text: t, options: { bullet: { indent: 16 }, breakLine: i < items.length - 1, paraSpaceAfter: 9 } })),
    Object.assign({ isTextBox: true, margin: 0, color: C.text1, valign: "top", fontSize: 15 }, o));
}
function subhead(slide, t, x, y, w, color = C.accent1) {
  txt(slide, t, { x, y, w, h: 0.3, fontSize: 14, bold: true, color });
  slide.addShape(pres.shapes.LINE, { x, y: y + 0.34, w, h: 0, line: { color, width: 1 } });
}
function add(master, section, title) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  return s;
}
// Simple table built from text cells, so every row stays editable
function table(slide, cols, rows, y0, o = {}) {
  const rowH = o.rowH || 0.5, head = o.headColor || C.text2, fs = o.fontSize || 13;
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y0, w: 8.8, h: 0.34, fill: { color: head }, line: { color: head } });
  cols.forEach(([t, x, w]) => txt(slide, t, { x: x + 0.1, y: y0, w: w - 0.2, h: 0.34, fontSize: 11, bold: true, color: C.background1, valign: "middle" }));
  rows.forEach((r, i) => {
    const y = y0 + 0.34 + i * rowH;
    if (i % 2 === 1) slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 8.8, h: rowH, fill: { color: C.background2 }, line: { color: C.background2 } });
    slide.addShape(pres.shapes.LINE, { x: 0.6, y: y + rowH, w: 8.8, h: 0, line: { color: GRAY, width: 0.5 } });
    r.forEach((cell, j) => {
      const [, x, w] = cols[j];
      if (cell && cell.chip) chip(slide, cell.chip, x + 0.1, y + (rowH - 0.26) / 2);
      else if (cell && cell.blank) slide.addShape(pres.shapes.LINE, { x: x + 0.1, y: y + rowH - 0.14, w: w - 0.2, h: 0, line: { color: C.text1, width: 0.75 } });
      else txt(slide, cell.text !== undefined ? cell.text : cell, Object.assign({ x: x + 0.1, y, w: w - 0.2, h: rowH, fontSize: fs, valign: "middle" }, cell.o || {}));
    });
  });
  return y0 + 0.34 + rows.length * rowH;
}

// ---------- 1. Cover ----------
pres.addSection({ title: "Overview" });
let s = pres.addSlide({ masterName: "COVER", sectionTitle: "Overview" });
s.addText("Tier 1 Roofer Retention", { placeholder: "title" });
s.addText("Closing the service gaps that cost us customers", { placeholder: "body" });
s.addText("Leadership Review  |  October 2026", { placeholder: "date" });
s.addNotes("Purpose of this meeting: review why Tier 1 roofing contractors drift away from Henry, agree on corrective actions, and assign an owner to each.");

// ---------- 2. Executive summary ----------
s = add("CONTENT", "Overview", "Executive Summary");
const scr = [
  ["Situation", "Henry wins Tier 1 roofers on product quality, Henry-provided leads and competitive pricing."],
  ["Problem", "We lose their business through five recurring service issues. Every one of them is within our control."],
  ["Recommendation", "Adopt four service standards, approve the corrective actions in this deck, and assign an owner to each issue."],
];
scr.forEach(([label, body], i) => {
  const y = 1.3 + i * 1.18;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 2.2, h: 0.98, fill: { color: i === 2 ? C.accent1 : C.text2 }, line: { color: i === 2 ? C.accent1 : C.text2 } });
  txt(s, label, { x: 0.8, y, w: 1.9, h: 0.98, fontSize: 16, bold: true, color: C.background1, valign: "middle" });
  s.addShape(pres.shapes.RECTANGLE, { x: 2.8, y, w: 6.6, h: 0.98, fill: { color: C.background2 }, line: { color: C.background2 } });
  txt(s, body, { x: 3.05, y, w: 6.15, h: 0.98, fontSize: 16, valign: "middle" });
});
s.addNotes("The short version. Our product, leads and price bring Tier 1 roofers to Henry. Our service lets them go. The issues are all ones we control, which means they are all ones we can fix. Today we are asking for approval of the actions and an owner for each issue.");

// ---------- 3. Issues at a glance ----------
s = add("CONTENT", "Overview", "Five Issues at a Glance");
const cols3 = [["#", 0.6, 0.45], ["Issue", 1.05, 2.55], ["What customers report", 3.6, 4.6], ["Status", 8.2, 1.2]];
table(s, cols3, [
  [{ text: "1", o: { bold: true } }, { text: "Customer Service", o: { bold: true, color: C.text2 } }, "Difficult to deal with: cumbersome ordering, slow answers", { chip: "O" }],
  [{ text: "2", o: { bold: true } }, { text: "Warranty", o: { bold: true, color: C.text2 } }, "Sign-offs and claim responses take too long", { chip: "O" }],
  [{ text: "3", o: { bold: true } }, { text: "Logistics", o: { bold: true, color: C.text2 } }, "Deliveries are unreliable and arrive without notice", { chip: "R" }],
  [{ text: "4", o: { bold: true } }, { text: "Material Availability", o: { bold: true, color: C.text2 } }, "Accessories are not ready to ship within 48 hours", { chip: "O" }],
  [{ text: "5", o: { bold: true } }, { text: "National Account Leads", o: { bold: true, color: C.text2 } }, "Our leads close well, but there are not enough of them", { chip: "Y" }],
], 1.2, { rowH: 0.58, fontSize: 14 });
txt(s, "Status reflects the team’s own assessment of each area.", { x: 0.6, y: 4.6, w: 8.8, h: 0.25, fontSize: 11, italic: true, color: MUTED });
s.addNotes("Logistics is the only area rated critical. Customer service, warranty and material availability are at risk. Lead volume is on watch. The next slide shows how these issues combine on a single order.");

// ---------- 4. Example order ----------
s = add("CONTENT", "Overview", "Example: How a Single Order Goes Wrong");
txt(s, "A composite of actual customer feedback", { x: 0.6, y: 0.98, w: 6, h: 0.25, fontSize: 12, italic: true, color: MUTED });
table(s, [["Step", 0.6, 1.6], ["What the customer does", 2.2, 3.0], ["What happens at Henry", 5.2, 3.2], ["Issue", 8.4, 1.0]], [
  [{ text: "1. Order", o: { bold: true, color: C.text2 } }, "Orders “24 drums of white silicone.”", "We send a 200-SKU list and ask for the SAP number.", "1"],
  [{ text: "2. Inquiry", o: { bold: true, color: C.text2 } }, "Asks a question about the order.", "The answer takes more than 24 hours.", "1"],
  [{ text: "3. Stock-out", o: { bold: true, color: C.text2 } }, "Needs PG988; Kingman is out.", "Carson has it, but no one offers to ship from there.", "1"],
  [{ text: "4. Shipment", o: { bold: true, color: C.text2 } }, "Expects delivery on the agreed date.", "We ship early without notice. No one is on site.", "3"],
  [{ text: "5. Delivery", o: { bold: true, color: C.text2 } }, "Receives the coating.", "The primer is missing, so the job cannot start.", "4"],
], 1.35, { rowH: 0.55, fontSize: 13 });
txt(s, "Result: the roofer looks bad in front of the building owner, and so does Henry.", { x: 0.6, y: 4.57, w: 8.8, h: 0.3, fontSize: 15, bold: true, color: C.text2 });
s.addNotes("Each step is real feedback from the detractor list. None is dramatic on its own; together they lose a customer.");

// ---------- Issue slides ----------
pres.addSection({ title: "Issues and Actions" });
function issue(n, title, status, report, actions, note) {
  const sl = add("ISSUE", "Issues and Actions", title);
  txt(sl, `Issue ${n} of 5`, { x: 7.6, y: 0.42, w: 1.8, h: 0.25, fontSize: 12, bold: true, color: MUTED, align: "right" });
  chip(sl, status, 8.55, 0.72);
  subhead(sl, "What Customers Report", 0.6, 1.3, 4.2, C.text2);
  bullets(sl, report, { x: 0.6, y: 1.85, w: 4.2, h: 3.0 });
  sl.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.2, h: 3.5, fill: { color: C.background2 }, line: { color: C.background2 } });
  subhead(sl, "Recommended Actions", 5.4, 1.45, 3.8, C.accent1);
  bullets(sl, actions, { x: 5.4, y: 2.0, w: 3.8, h: 2.7 });
  sl.addNotes(note);
}
issue(1, "Customer Service Is Hard to Work With", "O", [
  "Customers order “24 drums of white silicone” and we ask for the SAP number.",
  "Inquiries can wait more than 24 hours for a response.",
  "We will not substitute products, ship from another plant, or expedite.",
], [
  "Take orders as written; call the customer if unclear.",
  "Respond to every email within 4 hours.",
  "Give CS visibility into inventory at all plants.",
  "Train staff on approved product substitutions.",
  "Hold people accountable for service.",
], "Consistent feedback is that we are difficult to deal with. CS sends an Excel file with over 200 SKUs when the customer has already said what they want. Customers should be held to the order acknowledgement, and we should rely less on sales to sort out orders. If Kingman is out of PG988 and Carson has it, CS should get the product to the customer as quickly as possible.");
issue(2, "Warranty Response Is Too Slow", "O", [
  "Building owners expect action on a claim within 24 hours.",
  "Our claim responses do not meet that expectation.",
  "Some manufacturers sign off on warranties without walking the site.",
], [
  "Acknowledge every claim by email within 4 hours.",
  "Complete a site visit within 4 days.",
  "Sign off promptly, and walk the roof within the 2-year window while the roofer is responsible for leaks.",
  "Batch sign-offs 3+ hours from base to reduce travel cost.",
], "Sign-offs are rated at risk and claim responses are on watch. Walking the roof inside the two-year window protects the building owner and the roofer, and grouping distant sign-offs reduces travel cost.");
issue(3, "Deliveries Are Unreliable", "R", [
  "Material arrives days or a week early with no one there to accept it.",
  "No delivery notice or tracking is provided.",
  "If we can ship early, we do so without asking the customer.",
  "It makes the roofer look bad and makes Henry and Carlisle look unprofessional.",
], [
  "Notify the customer and get approval before shipping early.",
  "Place trackers on pallets and share tracking.",
  "CS provides daily updates without being asked.",
  "Move roofers to distribution and determine how to honor their pricing.",
], "This is the only area rated critical. If no delivery date is noted, we set one on the order acknowledgement, but we ship early without notice when we can. CS should communicate throughout the process.");
issue(4, "Accessories Are Not Ready to Ship", "O", [
  "We can have trucks of 988, but without primer the job cannot start.",
  "Accessories (primer, bleed blockers, mastics) are not reliably on hand. This is a long-standing issue.",
  "System gaps: no sprayable seam sealer for silicone; no cementitious coating for Arizona.",
], [
  "Raise minimum stock levels for accessories.",
  "Increase production ahead of known storms and weather.",
  "Have all job-start material ready to ship within 48 hours.",
  "Assign an owner to close the product gaps.",
], "Being ready to ship within 48 hours is not a recent problem. Everything needed to start a job must be available, not just the main coating. No owner has been assigned yet to the product gaps.");
issue(5, "Lead Quality Is Strong; Volume Is Low", "Y", [
  "National account leads close at a high rate; Facebook and website leads often do not.",
  "Sherwin-Williams provides more leads through its paint contracts.",
  "Some manufacturers offer contractors discounted thermal moisture surveys.",
], [
  "Collaborate with other Henry divisions on leads.",
  "Free the national accounts team from firefighting and CS work so it can generate leads.",
  "Moisture surveys: John T Hull.",
], "Lead quality is not the problem; volume is. Resolving the customer service issues gives the national accounts team its time back for lead generation.");

// ---------- Strengths ----------
pres.addSection({ title: "Recommendation" });
s = add("CONTENT", "Recommendation", "Strengths to Protect");
txt(s, "These are the reasons roofers choose Henry. The corrective actions must not put them at risk.", { x: 0.6, y: 1.05, w: 8.8, h: 0.3, fontSize: 14, color: MUTED });
table(s, [["Strength", 0.6, 2.6], ["Why it matters", 3.2, 6.2]], [
  [{ text: "Core products", o: { bold: true, color: C.text2 } }, "Silicone and SPF lead on quality and price."],
  [{ text: "Leads that close", o: { bold: true, color: C.text2 } }, "National account leads close at a high rate."],
  [{ text: "Flexibility on specifications", o: { bold: true, color: C.text2 } }, "Where other manufacturers require less, we remain willing to flex on scope and specifications."],
  [{ text: "Few failures", o: { bold: true, color: C.text2 } }, "A single failure ends future work with a building owner. We must keep innovating."],
], 1.55, { rowH: 0.65, fontSize: 14, headColor: C.accent4 });
s.addNotes("Fixing service should not cost us the reasons customers choose Henry.");

// ---------- Standards ----------
s = add("CONTENT", "Recommendation", "Proposed Service Standards");
table(s, [["Standard", 0.6, 2.0], ["Commitment", 2.6, 5.0], ["Addresses", 7.6, 1.8]], [
  [{ text: "4 hours", o: { bold: true, fontSize: 22, color: C.accent1 } }, "Respond to every email and acknowledge every warranty claim", "Issues 1, 2"],
  [{ text: "4 days", o: { bold: true, fontSize: 22, color: C.accent1 } }, "Complete a site visit on every warranty claim", "Issue 2"],
  [{ text: "48 hours", o: { bold: true, fontSize: 22, color: C.accent1 } }, "Have material ready to ship, accessories included", "Issue 4"],
  [{ text: "Daily", o: { bold: true, fontSize: 22, color: C.accent1 } }, "Proactive order updates from CS until delivery", "Issue 3"],
], 1.25, { rowH: 0.75, fontSize: 15 });
s.addNotes("Four simple, measurable standards drawn directly from the corrective actions. They are easy for every CS representative to remember and for leadership to track.");

// ---------- Decisions ----------
s = add("CONTENT", "Recommendation", "Decisions Requested Today");
txt(s, "Approve the corrective actions, assign an owner to each issue, and review progress monthly.", { x: 0.6, y: 1.05, w: 8.8, h: 0.3, fontSize: 14, color: C.text1 });
table(s, [["Issue", 0.6, 2.3], ["Key action", 2.9, 3.4], ["Measure", 6.3, 1.7], ["Owner", 8.0, 1.4]], [
  [{ text: "1. Customer Service", o: { bold: true, color: C.text2 } }, "4-hour response; inventory visibility", "Response time", { blank: true }],
  [{ text: "2. Warranty", o: { bold: true, color: C.text2 } }, "4-hour acknowledgement; 4-day visit", "Claim response", { blank: true }],
  [{ text: "3. Logistics", o: { bold: true, color: C.text2 } }, "Advance notice; tracking; daily updates", "On-time deliveries", { blank: true }],
  [{ text: "4. Material", o: { bold: true, color: C.text2 } }, "Accessories ready in 48 hours", "48-hour readiness", { blank: true }],
  [{ text: "5. Leads", o: { bold: true, color: C.text2 } }, "Division partnerships; team time", "Leads delivered", { blank: true }],
], 1.55, { rowH: 0.6, fontSize: 13 });
s.addNotes("This is the ask: approve the actions and name an owner for each issue before the meeting ends. The measures and the monthly review are proposals for leadership to confirm.");

// ---------- Appendix ----------
pres.addSection({ title: "Appendix" });
s = add("CONTENT", "Appendix", "Appendix: Detailed Assessment by Area");
const cards = [
  { name: "Product Quality", st: "G", items: [["G", "Silicone"], ["G", "SPF"], ["Y", "Acrylic"], ["R", "Acrylic clumping"], ["O", "Seam treatments"], ["R", "Urethane system"]] },
  { name: "Availability", st: "Y", items: [["G", "Silicone"], ["Y", "Accessories"], ["O", "Acrylic"], ["O", "Primers and bleed blockers"]] },
  { name: "Leads and Pricing", st: "G", items: [["G", "Henry-provided leads"], ["G", "Silicone price"], ["G", "Acrylic price"], ["Y", "SPF price"], ["Y", "Emulsion price"]] },
  { name: "Logistics", st: "O", items: [["O", "Tracking information"], ["R", "On-time delivery"], ["Y", "Delivered undamaged"]] },
  { name: "Customer Service", st: "O", items: [["R", "We ask for SKU numbers"], ["Y", "We ask for our form"], ["R", "Tracking without asking"], ["R", "Keeping jobs on schedule"], ["Y", "Attitude"], ["R", "Reps reachable by email"]] },
  { name: "Warranty", st: "Y", items: [["G", "Means something to owners"], ["G", "Claim process"], ["O", "Ease of processing"], ["Y", "Flexibility on standard spec"]] },
];
const cw = 2.8, ch = 1.68;
cards.forEach((c, i) => {
  const x = 0.6 + (i % 3) * (cw + 0.2), y = 1.15 + Math.floor(i / 3) * (ch + 0.1);
  s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: ch, fill: { color: C.background1 }, line: { color: GRAY, width: 0.75 }, objectName: c.name });
  txt(s, c.name, { x: x + 0.15, y: y + 0.12, w: 1.7, h: 0.26, fontSize: 13, bold: true, color: C.text2 });
  chip(s, c.st, x + cw - 1.0, y + 0.12);
  c.items.forEach(([k, t], j) => {
    const iy = y + 0.47 + j * 0.195;
    dot(s, k, x + 0.17, iy + 0.04);
    txt(s, t, { x: x + 0.38, y: iy, w: cw - 0.5, h: 0.19, fontSize: 11 });
  });
});
["G", "Y", "O", "R"].forEach((k, i) => {
  dot(s, k, 0.6 + i * 1.1, 4.69);
  txt(s, S[k].label, { x: 0.77 + i * 1.1, y: 4.63, w: 0.85, h: 0.25, fontSize: 10, color: MUTED, valign: "middle" });
});
s.addNotes("Full assessment from the team's Tier 1 analysis. Green = strength, yellow = watch, orange = at risk, red = critical.");

(async () => {
  await pres.writeFile({ fileName: "Leaky_Bucket.pptx" });
  await applyTheme("Leaky_Bucket.pptx", THEME);
  console.log("done");
})();
