const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/a9f08bbb-9e1f-48e1-b895-cf1232952f3d_c30e392a-4b21-4f80-b4de-330c3c8bc1fd/pptx/scripts/apply_theme.js");

// Henry brand palette (henry.com)
const THEME = {
  name: "Henry Leaky Bucket",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "333132", lt1: "FFFFFF", dk2: "00306C", lt2: "F1F1F2",
    accent1: "0054A6", accent2: "0090D2", accent3: "D93A2B",
    accent4: "2E8B57", accent5: "F2B134", accent6: "E8772E",
    hlink: "0054A6", folHlink: "00306C",
  },
};
const H = THEME.colors;
const GRAY = "BCBEC0", MIST = "9AD1D4", MUTED = "5F5D5E";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "We Stop Leaks for a Living";
pres.company = "Henry Company";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

const S = {
  G: { label: "Strength", fill: C.accent4, text: C.background1 },
  Y: { label: "Watch", fill: C.accent5, text: C.text1 },
  O: { label: "At risk", fill: C.accent6, text: C.background1 },
  R: { label: "Broken", fill: C.accent3, text: C.background1 },
};

const png = (buf) => "image/png;base64," + buf.toString("base64");
async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: String(size) }));
  return png(await sharp(Buffer.from(svg)).png().toBuffer());
}
const svgPng = async (svg) => png(await sharp(Buffer.from(svg), { density: 200 }).png().toBuffer());

// Top-down roof plan, drawn like an inspection report
function roofSvg(stroke, fill, unit, op = 1) {
  const dash = `stroke="${stroke}" stroke-width="2" stroke-dasharray="8 7" opacity="0.7"`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="800" viewBox="0 0 600 400" opacity="${op}">
    <rect x="8" y="8" width="584" height="384" fill="${fill}" stroke="${stroke}" stroke-width="7"/>
    <rect x="26" y="26" width="548" height="348" fill="none" stroke="${stroke}" stroke-width="2"/>
    <line x1="300" y1="26" x2="300" y2="374" ${dash}/>
    ${[[180, 200], [420, 200]].map(([x, y]) => {
      const xs = x < 300 ? [26, 300] : [300, 574];
      return [[xs[0], 26], [xs[0], 374], [xs[1], 26], [xs[1], 374]].map(([a, b]) => `<line x1="${x}" y1="${y}" x2="${a}" y2="${b}" ${dash}/>`).join("") +
        `<circle cx="${x}" cy="${y}" r="11" fill="${fill}" stroke="${stroke}" stroke-width="3"/><path d="M${x - 6} ${y - 6} L${x + 6} ${y + 6} M${x + 6} ${y - 6} L${x - 6} ${y + 6}" stroke="${stroke}" stroke-width="2.5"/>`;
    }).join("")}
    ${[[70, 52, 78, 50], [440, 296, 86, 52], [330, 80, 52, 52]].map(([x, y, w, h]) =>
      `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="${unit}" stroke="${stroke}" stroke-width="3"/><line x1="${x}" y1="${y}" x2="${x + w}" y2="${y + h}" stroke="${stroke}" stroke-width="1.5" opacity="0.6"/><line x1="${x + w}" y1="${y}" x2="${x}" y2="${y + h}" stroke="${stroke}" stroke-width="1.5" opacity="0.6"/>`).join("")}
    <rect x="506" y="54" width="40" height="40" fill="none" stroke="${stroke}" stroke-width="3"/>
  </svg>`;
}
// Leak positions on the roof plan, as fractions of its width and height
const LEAKS = [[0.19, 0.42], [0.79, 0.37], [0.5, 0.74], [0.9, 0.6], [0.21, 0.8]];

function pailSvg() {
  const navy = "#" + H.dk2, blue = "#" + H.accent1, water = "#" + H.accent2, red = "#" + H.accent3;
  const drop = (x, y, s = 1) => `<path transform="translate(${x} ${y}) scale(${s})" d="M0 0 C-9 14 -11 22 -11 27 A11 11 0 0 0 11 27 C11 22 9 14 0 0Z" fill="${red}"/>`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="800" height="960" viewBox="0 0 400 480">
    <path d="M85 92 C85 -8 315 -8 315 92" fill="none" stroke="${navy}" stroke-width="9" stroke-linecap="round"/>
    <path d="M70 92 L330 92 L304 392 Q200 412 96 392 Z" fill="#F1F1F2" stroke="${navy}" stroke-width="9" stroke-linejoin="round"/>
    <ellipse cx="200" cy="92" rx="130" ry="20" fill="${water}" stroke="${navy}" stroke-width="9"/>
    <path d="M81 200 L319 200 L311 290 L89 290 Z" fill="${blue}"/>
    <circle cx="96" cy="330" r="6" fill="${red}"/><circle cx="306" cy="250" r="6" fill="${red}"/><circle cx="200" cy="402" r="6" fill="${red}"/>
    ${drop(78, 345)}${drop(70, 405, 0.8)}${drop(326, 265)}${drop(336, 330, 0.8)}${drop(200, 418)}${drop(200, 455, 0.6)}
  </svg>`;
}

// ---------- Layouts ----------
const footer = (logo) => [{ image: { x: 0.6, y: 5.1, w: 0.66, h: 0.25, data: logo } }];
let LOGO_W, LOGO_B;

function defineLayouts() {
  pres.defineSlideMaster({
    title: "COVER",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.35, w: 5.4, h: 2.6, fontSize: 34, bold: true, color: C.background1, margin: 0, valign: "top", align: "left" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 4.55, w: 5.5, h: 0.35, fontSize: 13, color: MIST, margin: 0, valign: "top" }, text: "" } },
      { image: { x: 0.6, y: 0.5, w: 1.06, h: 0.4, data: LOGO_W } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    slideNumber: { x: 9.0, y: 5.1, w: 0.4, h: 0.25, fontSize: 10, color: MUTED, align: "right" },
    objects: [
      ...footer(LOGO_B),
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.42, w: 8.8, h: 0.95, fontSize: 28, bold: true, color: C.text2, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "LEAK",
    background: { color: C.background1 },
    slideNumber: { x: 9.0, y: 5.1, w: 0.4, h: 0.25, fontSize: 10, color: MUTED, align: "right" },
    objects: [
      ...footer(LOGO_B),
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.42, w: 7.2, h: 0.95, fontSize: 28, bold: true, color: C.text2, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "DECISION",
    background: { color: C.text2 },
    slideNumber: { x: 9.0, y: 5.1, w: 0.4, h: 0.25, fontSize: 10, color: MIST, align: "right" },
    objects: [
      { image: { x: 0.6, y: 5.1, w: 0.66, h: 0.25, data: LOGO_W } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.42, w: 8.8, h: 0.95, fontSize: 28, bold: true, color: C.background1, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
  });
}

// ---------- Helpers ----------
function txt(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, valign: "top" }, o));
}
function chip(slide, key, x, y, w = 0.8) {
  const s = S[key];
  slide.addText(s.label, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.11, x, y, w, h: 0.23, fill: { color: s.fill }, color: s.text, fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0, objectName: "Status " + s.label });
}
function dot(slide, key, x, y, d = 0.13) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: S[key].fill }, line: { color: S[key].fill } });
}
function marker(slide, n, cx, cy, d = 0.34, color = C.accent3, ripple = true) {
  if (ripple) slide.addShape(pres.shapes.OVAL, { x: cx - d * 0.85, y: cy - d * 0.85, w: d * 1.7, h: d * 1.7, fill: { color, transparency: 78 }, line: { color, transparency: 100 }, objectName: "Leak ripple " + n });
  slide.addText(String(n), { shape: pres.shapes.OVAL, x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color }, line: { color: C.background1, width: 1.5 }, color: C.background1, fontSize: d > 0.3 ? 14 : 9, bold: true, align: "center", valign: "middle", margin: 0, objectName: "Leak marker " + n });
}
function roof(slide, img, x, y, w, active, size) {
  const h = w / 1.5;
  slide.addImage({ data: img, x, y, w, h, objectName: "Roof plan" });
  LEAKS.forEach(([fx, fy], i) => {
    const on = active === undefined || active === i + 1;
    marker(slide, i + 1, x + fx * w, y + fy * h, on ? size : size * 0.7, on ? C.accent3 : GRAY, on && active === undefined);
  });
}
function iconRow(slide, img, text, x, y, w, o = {}) {
  slide.addImage({ data: img, x, y: y + 0.04, w: 0.2, h: 0.2 });
  txt(slide, text, Object.assign({ x: x + 0.34, y, w: w - 0.34, h: 0.5, fontSize: 14 }, o));
}
function addSlide(master, section, title) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (title) s.addText(title, { placeholder: "title" });
  return s;
}

(async () => {
  LOGO_W = png(fs.readFileSync("henry-logo.png"));
  LOGO_B = png(fs.readFileSync("henry-blue-logo.png"));
  defineLayouts();
  const I = {
    dropR: await icon("FaTint", H.accent3),
    checkW: await icon("FaCheck", H.lt1),
    checkB: await icon("FaCheck", H.accent1),
    checkG: await icon("FaCheck", H.accent4),
    roofDark: await svgPng(roofSvg("#" + MIST, "none", "none", 0.55)),
    roofLight: await svgPng(roofSvg("#" + H.accent1, "#F1F1F2", "#FFFFFF")),
    pail: await svgPng(pailSvg()),
  };

  // ---------- 1. Cover ----------
  pres.addSection({ title: "The problem" });
  let s = addSlide("COVER", "The problem");
  s.addText([
    { text: "We stop leaks", options: { breakLine: true } },
    { text: "for a living.", options: { breakLine: true } },
    { text: "So why are Tier 1", options: { color: MIST, breakLine: true } },
    { text: "roofers leaking out?", options: { color: MIST } },
  ], { placeholder: "title" });
  s.addText("Tier 1 roofer retention  ·  Leadership review", { placeholder: "body" });
  roof(s, I.roofDark, 6.15, 1.45, 3.45, undefined, 0.3);
  s.addNotes("Open with the irony. Henry's whole business is keeping water out of buildings. Our own customer base has leaks, and they are the kind we control. Today: where the five leaks are, what each one costs us with Tier 1 roofers, and the repair for each. At the end we need a decision: approve the repairs and put an owner on every leak.");

  // ---------- 2. The bucket ----------
  s = addSlide("CONTENT", "The problem", "Product, leads and price bring roofers in. Service lets them leak out.");
  s.addImage({ data: I.pail, x: 3.85, y: 1.45, w: 2.3, h: 2.76, objectName: "Leaking pail" });
  s.addImage({ data: LOGO_W, x: 4.5, y: 2.66, w: 1.0, h: 0.38, objectName: "Pail label" });
  txt(s, "What brings them in", { x: 0.6, y: 1.55, w: 3.0, h: 0.3, fontSize: 15, bold: true, color: C.accent1 });
  ["Silicone and SPF quality", "Leads that close", "Competitive pricing", "Flexibility on specs"].forEach((t, i) => iconRow(s, I.checkB, t, 0.6, 2.0 + i * 0.48, 3.1, { fontSize: 15, h: 0.35 }));
  txt(s, "What lets them out", { x: 6.55, y: 1.55, w: 2.9, h: 0.3, fontSize: 15, bold: true, color: C.accent3 });
  ["Hard to place an order", "Slow answers", "Unreliable deliveries", "Material not ready"].forEach((t, i) => iconRow(s, I.dropR, t, 6.55, 2.0 + i * 0.48, 2.9, { fontSize: 15, h: 0.35 }));
  txt(s, "Every leak on the right is ours. So is the fix.", { x: 0.6, y: 4.45, w: 8.8, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
  s.addNotes("This is the leaky bucket. Roofers do not leave us because of the product: silicone and SPF are strong, our leads close, and our price competes. They leave because we are hard to work with. Pouring in more leads or new products does not help while the bucket leaks.");

  // ---------- 3. Story ----------
  s = addSlide("CONTENT", "The problem", "One order, five ways to lose a roofer");
  txt(s, "A composite of real customer feedback", { x: 0.6, y: 1.0, w: 6, h: 0.3, fontSize: 12, italic: true, color: MUTED });
  const steps = [
    ["They order", "“24 drums of white silicone.”", "We send a 200-SKU list and ask for the SAP number.", 1],
    ["They ask", "A question about the order.", "The answer takes more than 24 hours.", 1],
    ["They need it", "Kingman is out of PG988.", "Carson has it. Nobody offers to ship it.", 1],
    ["We ship", "Earlier than planned.", "No call first. No one on site to sign for it.", 3],
    ["It lands", "The drums arrive.", "The primer doesn't. The job can't start.", 4],
  ];
  const colW = 1.76;
  s.addShape(pres.shapes.LINE, { x: 0.6, y: 2.75, w: 8.8, h: 0, line: { color: GRAY, width: 1.5 } });
  steps.forEach(([h1, a, b, leak], i) => {
    const x = 0.6 + i * colW, w = colW - 0.2;
    txt(s, h1, { x, y: 1.55, w, h: 0.3, fontSize: 15, bold: true, color: C.text2 });
    txt(s, a, { x, y: 1.88, w, h: 0.65, fontSize: 12, italic: true, color: MUTED });
    s.addShape(pres.shapes.OVAL, { x, y: 2.62, w: 0.26, h: 0.26, fill: { color: C.accent3 }, line: { color: C.background1, width: 2 } });
    txt(s, b, { x, y: 3.05, w, h: 0.85, fontSize: 13 });
    txt(s, "Leak " + leak, { x, y: 3.95, w, h: 0.25, fontSize: 11, bold: true, color: C.accent3 });
  });
  txt(s, "The roofer looks bad in front of the building owner. So does Henry.", { x: 0.6, y: 4.45, w: 8.8, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
  s.addNotes("Tell this as one roofer's week. Each step is real feedback from the detractor list. A roofer asks for 24 drums of white silicone and we send back a 200-SKU spreadsheet asking for the SAP number. A question sits for more than a day. Kingman is out of PG988; Carson has it, but CS doesn't offer it. We ship early without warning and no one is there to receive it. The coating arrives without the primer, so the job can't start. Small misses, one lost customer.");

  // ---------- 4. Leak map ----------
  s = addSlide("CONTENT", "The problem", "We found five leaks. Every one is fixable.");
  roof(s, I.roofLight, 0.6, 1.45, 4.95, undefined, 0.36);
  const legend = [
    ["Service", "Roofers say we're hard to deal with", "O"],
    ["Warranty", "Sign-offs and claims move too slowly", "O"],
    ["Logistics", "Deliveries arrive late or unannounced", "R"],
    ["Material", "Accessories aren't ready to ship", "O"],
    ["Leads", "Leads close, but there are too few", "Y"],
  ];
  legend.forEach(([t, d, k], i) => {
    const y = 1.45 + i * 0.68;
    marker(s, i + 1, 6.0 + 0.16, y + 0.17, 0.32, C.accent3, false);
    txt(s, t, { x: 6.5, y, w: 1.6, h: 0.3, fontSize: 15, bold: true, color: C.text2 });
    chip(s, k, 8.6, y + 0.04);
    txt(s, d, { x: 6.5, y: y + 0.3, w: 2.9, h: 0.3, fontSize: 12, color: MUTED });
  });
  s.addNotes("Think of this as the inspection report. Five leaks, each marked on the roof. Status colors come from the team's own assessment: orange is at risk, red is broken, yellow is watch. Logistics is the only one marked broken. The next five slides take each leak in turn: what roofers live with, and the repair.");

  // ---------- Leak slides ----------
  pres.addSection({ title: "The five leaks" });
  function leakSlide(n, title) {
    const sl = addSlide("LEAK", "The five leaks", title);
    roof(sl, I.roofLight, 8.2, 0.42, 1.2, n, 0.2);
    return sl;
  }
  const nLines = (t, cpl) => t.split(" ").reduce((a, w) => { const l = a[a.length - 1]; if ((l + " " + w).trim().length > cpl) a.push(w); else a[a.length - 1] = (l + " " + w).trim(); return a; }, [""]).length;
  function repairPanel(sl, items, h = 3.3) {
    sl.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.6, w: 4.0, h, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "Repair panel" });
    txt(sl, "The repair", { x: 5.65, y: 1.8, w: 3.5, h: 0.3, fontSize: 16, bold: true, color: C.background1 });
    let y = 2.25;
    items.forEach(([t]) => {
      const lines = nLines(t, 37);
      iconRow(sl, I.checkW, t, 5.65, y, 3.55, { fontSize: 14, color: C.background1, h: lines * 0.25 });
      y += lines * 0.25 + 0.2;
    });
    if (y - 0.2 > 1.6 + h) console.warn("Repair panel overflow on", sl._slideNum, y);
  }
  function liveWith(sl, items, startY = 1.6) {
    txt(sl, "What roofers live with", { x: 0.6, y: startY, w: 4.6, h: 0.3, fontSize: 13, bold: true, color: MUTED });
    let y = startY + 0.45;
    items.forEach(([t, lines]) => {
      iconRow(sl, I.dropR, t, 0.6, y, 4.4, { fontSize: 15, h: lines * 0.27 });
      y += lines * 0.27 + 0.25;
    });
  }

  // Leak 1
  s = leakSlide(1, "Leak 1: We’re hard to deal with");
  liveWith(s, [
    ["They ask for “24 drums of white silicone.” We ask for the SAP number.", 2],
    ["Questions, not even orders, wait more than 24 hours.", 2],
    ["Kingman is out of PG988. Carson has it. The roofer still waits.", 2],
  ]);
  repairPanel(s, [
    ["Take orders as written. If unclear, call.", 2],
    ["Answer every email within 4 hours.", 1],
    ["Show CS inventory at every plant, and train on substitutions.", 2],
    ["Hold people accountable for service.", 1],
  ]);
  s.addNotes("Consistent feedback: we are difficult to deal with. CS sends an Excel file with over 200 SKUs when the customer has already said what they want. Hold customers to the order acknowledgement and rely less on sales to sort out orders. Inquiries can take over 24 hours. When one plant is out and another has stock, CS should get the product to the customer, especially when they need it by a date. Underneath every fix: people are held accountable.");

  // Leak 2
  s = leakSlide(2, "Leak 2: Warranties move too slowly");
  txt(s, "Building owners expect action within", { x: 0.6, y: 1.6, w: 4.6, h: 0.3, fontSize: 15, color: C.text1 });
  txt(s, "24 hours", { x: 0.6, y: 1.95, w: 4.6, h: 0.9, fontSize: 60, bold: true, color: C.accent3 });
  iconRow(s, I.dropR, "Our claim responses don’t keep up.", 0.6, 3.1, 4.4, { fontSize: 15, h: 0.3 });
  iconRow(s, I.dropR, "Some manufacturers sign off without even walking the site.", 0.6, 3.6, 4.4, { fontSize: 15, h: 0.55 });
  repairPanel(s, [
    ["Acknowledge every claim by email within 4 hours.", 2],
    ["Visit the site within 4 days.", 1],
    ["Sign off fast. Walk the roof inside the 2-year window.", 2],
    ["Batch sign-offs 3+ hours from base to cut travel.", 2],
  ]);
  s.addNotes("Two parts. Sign-offs: competitors move fast, some without walking the site. We should sign off as soon as possible and walk the roof within the two-year window when the roofer is responsible for leaks. Bunching sign-offs in areas 3+ hours from home base cuts travel cost. Claims: owners expect action within 24 hours. Commit to a 4-hour email acknowledgement and a site visit within 4 days.");

  // Leak 3
  s = leakSlide(3, "Leak 3: Roofers say we’re unreliable");
  txt(s, "Material shows up a week early. No one is there to sign for it. The roofer looks bad, and so do we.", { x: 0.6, y: 1.6, w: 4.4, h: 1.7, fontSize: 22, bold: true, color: C.text2 });
  iconRow(s, I.dropR, "No notice and no tracking.", 0.6, 3.55, 4.4, { fontSize: 15, h: 0.3 });
  iconRow(s, I.dropR, "If we can ship early, we do, without asking.", 0.6, 4.0, 4.4, { fontSize: 15, h: 0.3 });
  repairPanel(s, [
    ["Ask before shipping early.", 1],
    ["Put trackers on every pallet and share tracking.", 2],
    ["Daily CS updates, without being asked.", 2],
    ["Move roofers to distribution, and work out how to honor their price.", 3],
  ]);
  s.addNotes("Feedback is that we are unreliable. If no delivery date is noted we set one on the order acknowledgement, but if material can go sooner we ship without telling the customer. Material arrives days or a week early with no one to accept it. It makes the roofer look bad and makes Henry and Carlisle look unprofessional. CS should communicate throughout: ask before an early shipment, provide tracking, and check in daily without being asked.");

  // Leak 4
  s = leakSlide(4, "Leak 4: No primer, no job");
  txt(s, "We can have trucks of 988. Without the primer, the job can’t start.", { x: 0.6, y: 1.6, w: 4.4, h: 1.3, fontSize: 22, bold: true, color: C.text2 });
  iconRow(s, I.dropR, "Material isn’t ready to ship within 48 hours, and this isn’t new.", 0.6, 3.1, 4.4, { fontSize: 15, h: 0.55 });
  iconRow(s, I.dropR, "No sprayable seam sealer for silicone. No cementitious coating for Arizona.", 0.6, 3.85, 4.4, { fontSize: 15, h: 0.55 });
  repairPanel(s, [
    ["Raise minimums on primer, bleed blockers and mastics.", 2],
    ["Ramp up production ahead of known storms and weather.", 2],
    ["Have accessories ready to ship within 48 hours.", 2],
    ["Name an owner for the system gaps.", 1],
  ]);
  s.addNotes("Not having material ready within 48 hours is not a recent problem. Everything needed to start a job, primer, bleed blockers, mastics, has to be on hand, not just the main coating. Our systems also have gaps: a sprayable seam sealer for silicone systems and a cementitious coating for the Arizona market. No fix has been assigned to those gaps yet.");

  // Leak 5
  s = leakSlide(5, "Leak 5: Good leads, too few of them");
  liveWith(s, [
    ["Sherwin-Williams sends more leads, thanks to its paint contracts.", 2],
    ["Some manufacturers offer contractors discounted thermal moisture surveys.", 2],
    ["Our national accounts team spends its time firefighting and covering CS.", 2],
  ]);
  repairPanel(s, [
    ["Work with other Henry divisions on leads.", 2],
    ["Give the national accounts team its time back for lead generation.", 3],
    ["Moisture surveys: John T Hull.", 1],
  ]);
  s.addNotes("Lead quality is not the problem: national account leads close at a high rate, though Facebook and website leads often don't. Volume is. We're one of the better lead providers, but Sherwin-Williams provides more through paint contracts. Some manufacturers have national agreements with moisture survey companies that give contractors discounted thermal surveys. Fixing Leak 1 frees the national accounts team to generate leads instead of putting out fires.");

  // ---------- What works ----------
  pres.addSection({ title: "The decision" });
  s = addSlide("CONTENT", "The decision", "While we fix the leaks, protect what already works");
  const works = [
    ["Core products", "Silicone and SPF lead on quality and price."],
    ["Leads that close", "National account leads close at a high rate."],
    ["Flexible on specs", "Others ask for less. We stay willing to flex on scope and specs."],
    ["Few failures", "One failure ends future work with a building owner. Keep innovating."],
  ];
  works.forEach(([t, d], i) => {
    const x = 0.6 + i * 2.25, w = 2.0;
    s.addShape(pres.shapes.OVAL, { x, y: 1.7, w: 0.55, h: 0.55, fill: { color: C.accent4 }, line: { color: C.accent4 } });
    s.addImage({ data: I.checkW, x: x + 0.15, y: 1.85, w: 0.25, h: 0.25 });
    txt(s, t, { x, y: 2.5, w, h: 0.35, fontSize: 17, bold: true, color: C.text2 });
    txt(s, d, { x, y: 2.95, w, h: 1.2, fontSize: 14 });
  });
  s.addNotes("The fixes should not cost us the reasons roofers choose Henry. Core products are strong, national account leads close, we flex on scope and specs where many manufacturers won't, and failures are rare. One failure kills future work with a building owner, so we keep innovating.");

  // ---------- Four numbers ----------
  s = addSlide("CONTENT", "The decision", "Four promises every roofer can count on");
  const hdrY = 1.5;
  [["Promise", 0.6, 2.3], ["What it means", 3.0, 4.6], ["Plugs", 8.1, 1.3]].forEach(([t, x, w]) =>
    txt(s, t, { x, y: hdrY, w, h: 0.25, fontSize: 11, bold: true, color: MUTED }));
  const promises = [
    ["4 hours", "To answer every email and acknowledge every warranty claim", [1, 2]],
    ["4 days", "To visit the site on a warranty claim", [2]],
    ["48 hours", "To have material ready to ship, accessories included", [4]],
    ["Daily", "Order updates from CS, without the roofer asking", [3]],
  ];
  promises.forEach(([big, d, leaks], i) => {
    const y = 1.85 + i * 0.75;
    s.addShape(pres.shapes.LINE, { x: 0.6, y: y - 0.06, w: 8.8, h: 0, line: { color: GRAY, width: 1 } });
    txt(s, big, { x: 0.6, y: y + 0.02, w: 2.3, h: 0.6, fontSize: 30, bold: true, color: C.accent1, valign: "middle" });
    txt(s, d, { x: 3.0, y: y + 0.02, w: 4.8, h: 0.6, fontSize: 15, valign: "middle" });
    leaks.forEach((n, j) => marker(s, n, 8.27 + j * 0.42, y + 0.32, 0.3, C.accent3, false));
  });
  s.addShape(pres.shapes.LINE, { x: 0.6, y: 1.85 + 4 * 0.75 - 0.06, w: 8.8, h: 0, line: { color: GRAY, width: 1 } });
  s.addNotes("Four numbers that turn the repairs into something a roofer can feel and we can measure. All four come straight from the fixes in the detractor list. They are simple enough for every CS rep to remember.");

  // ---------- Decision ----------
  s = addSlide("DECISION", "The decision", "Decision today: approve the repairs and name an owner for each leak");
  [["Leak", 0.6, 2.1], ["Repair", 2.75, 3.3], ["We measure", 6.15, 1.9], ["Owner", 8.1, 1.3]].forEach(([t, x, w]) =>
    txt(s, t, { x, y: 1.55, w, h: 0.25, fontSize: 11, bold: true, color: MIST }));
  const decisions = [
    ["Service", "4-hour replies; inventory visible at every plant", "Email response time"],
    ["Warranty", "4-hour acknowledgement; 4-day site visit", "Claim response time"],
    ["Logistics", "Notice before shipping; trackers; daily updates", "Notified, on-time deliveries"],
    ["Material", "Accessories ready to ship in 48 hours", "Orders ready in 48 hours"],
    ["Leads", "Division partnerships; team time for leads", "Leads delivered"],
  ];
  decisions.forEach(([l, r, m], i) => {
    const y = 1.9 + i * 0.6;
    s.addShape(pres.shapes.LINE, { x: 0.6, y: y - 0.06, w: 8.8, h: 0, line: { color: MIST, width: 0.75, transparency: 50 } });
    marker(s, i + 1, 0.74, y + 0.22, 0.28, C.accent3, false);
    txt(s, l, { x: 1.05, y, w: 1.6, h: 0.45, fontSize: 15, bold: true, color: C.background1, valign: "middle" });
    txt(s, r, { x: 2.75, y, w: 3.25, h: 0.45, fontSize: 13, color: C.background1, valign: "middle" });
    txt(s, m, { x: 6.15, y, w: 1.85, h: 0.45, fontSize: 13, color: MIST, valign: "middle" });
    s.addShape(pres.shapes.LINE, { x: 8.1, y: y + 0.36, w: 1.3, h: 0, line: { color: C.background1, width: 1.25 } });
  });
  s.addNotes("This is the ask. Approve the five repairs and put one name on each leak before we leave the room. The measures are a proposal: one number per leak, reviewed weekly. Fill in the owner column live.");

  // ---------- Appendix ----------
  pres.addSection({ title: "Appendix" });
  s = addSlide("CONTENT", "Appendix", "Appendix: full inspection results");
  const cards = [
    { name: "Product quality", st: "G", items: [["G", "Silicone"], ["G", "SPF"], ["Y", "Acrylic"], ["R", "Acrylic clumping"], ["O", "Seam treatments"], ["R", "Urethane system"]] },
    { name: "Availability", st: "Y", items: [["G", "Silicone"], ["Y", "Accessories"], ["O", "Acrylic"], ["O", "Primers and bleed blockers"]] },
    { name: "Leads and pricing", st: "G", items: [["G", "Henry-provided leads"], ["G", "Silicone price"], ["G", "Acrylic price"], ["Y", "SPF price"], ["Y", "Emulsion price"]] },
    { name: "Logistics", st: "O", items: [["O", "Tracking info"], ["R", "On-time delivery"], ["Y", "Delivered undamaged"]] },
    { name: "Customer service", st: "O", items: [["R", "We ask for SKU numbers"], ["Y", "We ask for our form"], ["R", "Tracking without asking"], ["R", "Keeping jobs on schedule"], ["Y", "Attitude"], ["R", "Reps reachable by email"]] },
    { name: "Warranty", st: "Y", items: [["G", "Means something to owners"], ["G", "Claim process"], ["O", "Easy to process"], ["Y", "Flex on standard spec"]] },
  ];
  const cw = 2.8, ch = 1.68;
  cards.forEach((c, i) => {
    const x = 0.6 + (i % 3) * (cw + 0.2), y = 1.2 + Math.floor(i / 3) * (ch + 0.12);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: ch, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: c.name });
    txt(s, c.name, { x: x + 0.15, y: y + 0.12, w: 1.7, h: 0.25, fontSize: 13, bold: true, color: C.text2 });
    chip(s, c.st, x + cw - 0.95, y + 0.13);
    c.items.forEach(([k, t], j) => {
      const iy = y + 0.47 + j * 0.195;
      dot(s, k, x + 0.17, iy + 0.04, 0.12);
      txt(s, t, { x: x + 0.38, y: iy, w: cw - 0.5, h: 0.19, fontSize: 11 });
    });
  });
  ["G", "Y", "O", "R"].forEach((k, i) => {
    dot(s, k, 4.3 + i * 1.05, 5.165, 0.12);
    txt(s, S[k].label, { x: 4.47 + i * 1.05, y: 5.1, w: 0.8, h: 0.25, fontSize: 10, color: MUTED, valign: "middle" });
  });
  s.addNotes("The full Tier 1 tree from the team's assessment, regrouped into six areas. Colors follow the original: green strength, yellow watch, orange at risk, red broken.");

  await pres.writeFile({ fileName: "Leaky_Bucket.pptx" });
  await applyTheme("Leaky_Bucket.pptx", THEME);
  console.log("done");
})();
