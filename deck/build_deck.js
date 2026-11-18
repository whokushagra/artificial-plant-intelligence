// Builds KushagraKaushik_05-Final.pptx: the investor pitch deck for Artificial Plant Intelligence.
// Backgrounds in deck/bg are clean stills rendered from index.html.
//   NODE_PATH=<folder with pptxgenjs> node deck/build_deck.js <path to pptx skill apply_theme.js>
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.argv[2]);
const BG = (f) => path.join(__dirname, "bg", f);
const OUT = path.join(__dirname, "..", "KushagraKaushik_05-Final.pptx");

const THEME = {
  name: "Artificial Plant Intelligence",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "F2F4F0", lt1: "050806", dk2: "A7B2AA", lt2: "0D1410",
    accent1: "A8E07A", accent2: "FF6F7D", accent3: "6FE0D8", accent4: "C07BFF",
    accent5: "FFD24A", accent6: "FF9A2E", hlink: "A8E07A", folHlink: "6FE0D8",
  },
};
const SERIF = "Georgia", MONO = "Courier New";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Artificial Plant Intelligence";
pres.author = "Kushagra Kaushik";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

const footer = () => [
  { text: { text: "PROJECT VENUS  ·  ARTIFICIAL PLANT INTELLIGENCE", options: { x: 0.6, y: 6.98, w: 7, h: 0.3, fontFace: MONO, fontSize: 10, color: C.text2, charSpacing: 2, margin: 0 } } },
];
const num = { x: 12.25, y: 6.98, w: 0.5, h: 0.3, fontFace: MONO, fontSize: 10, color: C.text2, align: "right" };
const ph = (name, type, o) => ({ placeholder: { options: { name, type, color: C.text1, margin: 0, ...o }, text: "" } });

pres.defineSlideMaster({ title: "COVER", background: { color: C.background1 }, slideNumber: num, objects: [
  ...footer(),
  ph("title", "title", { x: 0.8, y: 2.6, w: 11.73, h: 1.5, align: "center", valign: "middle", fontSize: 58 }),
  ph("body", "body", { x: 2, y: 4.2, w: 9.33, h: 0.7, align: "center", valign: "top", fontSize: 20, color: C.text2 }),
] });
pres.defineSlideMaster({ title: "LEFT", background: { color: C.background1 }, slideNumber: num, objects: [
  ...footer(),
  ph("title", "title", { x: 0.7, y: 1.0, w: 6.2, h: 1.7, align: "left", valign: "top", fontSize: 38 }),
] });
pres.defineSlideMaster({ title: "TOP", background: { color: C.background1 }, slideNumber: num, objects: [
  ...footer(),
  ph("title", "title", { x: 0.8, y: 0.85, w: 11.73, h: 1.1, align: "center", valign: "top", fontSize: 38 }),
] });
pres.defineSlideMaster({ title: "BOTTOM", background: { color: C.background1 }, slideNumber: num, objects: [
  ...footer(),
  ph("title", "title", { x: 0.8, y: 4.75, w: 11.73, h: 1.15, align: "center", valign: "bottom", fontSize: 38 }),
] });

// ---------- helpers ----------
let N = 0;
const T = (...parts) => parts.map((p) => (typeof p === "string" ? { text: p } : { text: p.i, options: { fontFace: SERIF, italic: true } }));
const eyebrow = (s, text, x, y, w = 6, align = "left") =>
  s.addText(text, { x, y, w, h: 0.3, fontFace: MONO, fontSize: 11, color: C.accent1, charSpacing: 3, margin: 0, align, isTextBox: true, objectName: "eyebrow" });
const base = (master, bg, overlays, section) => {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  N++;
  s.background = { path: BG(bg) };
  overlays.forEach((o) => s.addImage({ path: BG(o), x: 0, y: 0, w: 13.333, h: 7.5, objectName: "overlay" }));
  return s;
};
const body = (s, text, o) => s.addText(text, { fontSize: 16, color: C.text2, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 6, ...o });
const card = (s, x, y, w, h, name) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C.background2, transparency: 22 }, line: { color: C.text2, transparency: 72, width: 0.75 }, objectName: name });

// ===================== 1. COVER =====================
pres.addSection({ title: "Opening" });
{
  const s = base("COVER", "hero__62.jpg", ["ov_soft.png", "ov_soft.png"], "Opening");
  eyebrow(s, "A NEW KIND OF MIND", 3.67, 2.15, 6, "center");
  s.addText(T("Artificial ", { i: "Plant" }, " Intelligence"), { placeholder: "title" });
  s.addText("A way for people and plants to understand each other.", { placeholder: "body" });
  s.addNotes("Open on the meadow. Artificial Plant Intelligence: a new way for people and plants to understand each other. Not a gadget, a new kind of mind.");
}

// ===================== 2. PROBLEM =====================
pres.addSection({ title: "The problem" });
{
  const s = base("TOP", "input__7.jpg", ["ov_top.png", "ov_bottom.png"], "The problem");
  eyebrow(s, "THE PROBLEM", 3.67, 0.5, 6, "center");
  s.addText(T("What we take in ", { i: "shapes" }, " us"), { placeholder: "title" });
  body(s, "More and more of it now arrives through a screen, and what we eat, grow and own arrives through a store.", { x: 2.4, y: 1.8, w: 8.53, h: 0.7, align: "center" });
  const cards = [
    ["Silent, still costly", "Just having your phone on the desk lowers the working memory you have, even face-down.", "WARD ET AL. 2017"],
    ["Wired early", "More screen time in early childhood is linked to weaker white matter in language pathways.", "HUTTON ET AL. 2020"],
    ["Split attention", "Constant notifications split attention and may reshape how we focus and remember.", "FIRTH ET AL. 2019"],
  ];
  cards.forEach(([h, d, src], i) => {
    const x = 0.8 + i * 3.98;
    card(s, x, 4.55, 3.75, 2.05, "stat card " + (i + 1));
    s.addText(h, { x: x + 0.25, y: 4.75, w: 3.25, h: 0.4, fontSize: 18, bold: true, color: C.text1, margin: 0, isTextBox: true });
    s.addText(d, { x: x + 0.25, y: 5.2, w: 3.25, h: 0.95, fontSize: 14, color: C.text2, margin: 0, valign: "top", isTextBox: true });
    s.addText(src, { x: x + 0.25, y: 6.2, w: 3.25, h: 0.25, fontFace: MONO, fontSize: 10, color: C.text2, margin: 0, isTextBox: true });
  });
  s.addNotes("The problem. What we consume shapes our brains, and today most of it comes through screens and stores. Three findings support this.");
}

// ===================== 3. EVIDENCE =====================
{
  const s = base("TOP", "two__7.jpg", ["ov_top.png"], "The problem");
  eyebrow(s, "SAME BRAIN · DIFFERENT INPUT", 3.67, 0.5, 6, "center");
  s.addText(T("Change the input, change the ", { i: "pattern" }), { placeholder: "title" });
  const col = (x, head, rows) => {
    s.addText(head, { x, y: 5.15, w: 4.6, h: 0.3, fontFace: MONO, fontSize: 12, color: C.text1, charSpacing: 3, margin: 0, isTextBox: true });
    rows.forEach(([t, src], i) => {
      s.addText(t, { x, y: 5.5 + i * 0.6, w: 4.6, h: 0.35, fontSize: 14, color: C.text1, margin: 0, isTextBox: true });
      s.addText(src, { x, y: 5.83 + i * 0.6, w: 4.6, h: 0.22, fontFace: MONO, fontSize: 10, color: C.text2, margin: 0, isTextBox: true });
    });
  };
  col(1.3, "A DAY OF SCREENS", [["Weaker white matter in young children", "HUTTON ET AL. 2020"], ["Attention split by constant alerts", "FIRTH ET AL. 2019"]]);
  col(7.6, "AN HOUR IN NATURE", [["Activity in the amygdala (stress) goes down", "SUDIMAC, SALE & KÜHN 2022"], ["Rumination quiets; attention up ~20%", "BRATMAN 2015 · BERMAN 2008"]]);
  s.addText("Heat-map illustration based on published findings, not real brain scans", { x: 3.17, y: 6.72, w: 7, h: 0.25, fontFace: MONO, fontSize: 10, color: C.text2, align: "center", margin: 0, isTextBox: true });
  s.addNotes("Two brains, same structure, different input. The colours illustrate what the studies found; they are not scans.");
}

// ===================== 4. THE MIND =====================
pres.addSection({ title: "How a mind works" });
{
  const s = base("LEFT", "mind__45.jpg", ["ov_left.png"], "How a mind works");
  eyebrow(s, "THE HUMAN MIND", 0.7, 0.6);
  s.addText(T("Everything you feel begins as a ", { i: "signal" }), { placeholder: "title" });
  const steps = [["01", "Take in", "Sensory cortex", "Your senses catch the world"], ["02", "Sort", "Thalamus", "Most signals are filtered out"], ["03", "Remember", "Hippocampus", "What is left meets your past"], ["04", "Decide", "Prefrontal cortex", "It weighs it all and decides"], ["05", "Express", "Broca's area", "The feeling becomes words"]];
  steps.forEach(([n, a, b, d], i) => {
    const y = 3.05 + i * 0.7;
    s.addText(n, { x: 0.7, y, w: 0.5, h: 0.3, fontFace: MONO, fontSize: 12, color: C.accent1, margin: 0, isTextBox: true });
    s.addText([{ text: a + "  —  " }, { text: b, options: { bold: true } }], { x: 1.25, y: y - 0.02, w: 5, h: 0.32, fontSize: 16, color: C.text1, margin: 0, isTextBox: true });
    s.addText(d, { x: 1.25, y: y + 0.3, w: 5, h: 0.28, fontSize: 14, color: C.text2, margin: 0, isTextBox: true });
  });
  s.addNotes("A light, general picture of how a brain turns a signal into words. We will see the same path again, rebuilt for a plant.");
}

// ===================== 5. OPPORTUNITY =====================
{
  const s = base("BOTTOM", "become__75.jpg", ["ov_bottom.png"], "How a mind works");
  eyebrow(s, "THE OPPORTUNITY", 3.67, 4.15, 6, "center");
  s.addText([{ text: "Change what we take in,", options: { breakLine: true } }, { text: "and we change who we " }, { text: "become", options: { fontFace: SERIF, italic: true } }], { placeholder: "title" });
  body(s, "Two hours a week in nature is linked to better health and wellbeing (White et al. 2019). A brain that notices living things grows more connected, and more caring.", { x: 2.2, y: 6.0, w: 8.93, h: 0.75, align: "center", fontSize: 15 });
  s.addNotes("The opportunity: if input shapes the brain, change the input. Living things instead of screens.");
}

// ===================== 6. SOLUTION =====================
pres.addSection({ title: "The solution" });
{
  const s = base("TOP", "reveal__8.jpg", ["ov_top.png"], "The solution");
  eyebrow(s, "THE SOLUTION · INTRODUCING", 3.67, 0.5, 6, "center");
  s.addText(T("Artificial ", { i: "Plant" }, " Intelligence"), { placeholder: "title" });
  body(s, "A plant with a mind people can understand. It senses its world, makes sense of it, decides how it feels, and tells you in its own voice.", { x: 2.4, y: 1.95, w: 8.53, h: 0.8, align: "center" });
  s.addNotes("The reveal. Our first proof: a real Venus flytrap with real sensors and a brain we built for it.");
}

// ===================== 7. HOW IT THINKS (pipeline) =====================
pres.addSection({ title: "How it works" });
{
  const s = base("LEFT", "how__08.jpg", ["ov_left.png", "ov_bottom.png"], "How it works");
  eyebrow(s, "HOW IT WORKS", 0.7, 0.6);
  s.addText(T("How it ", { i: "thinks" }), { placeholder: "title" });
  body(s, "The same path as yours, rebuilt for a plant. Every sense becomes a line of light, and they all flow into its brain: a seed.", { x: 0.7, y: 1.85, w: 5.2, h: 1.0 });
  const flow = [["Sense", "senses"], ["Signal", "nerves"], ["Brain", "the seed"], ["Memory", "hippocampus"], ["Judge", "thalamus"], ["World", "context"], ["Decide", "prefrontal"], ["Personality", "who it is"], ["Voice", "Broca's area"]];
  const w = 11.93 / flow.length;
  flow.forEach(([a, b], i) => {
    const x = 0.7 + i * w;
    s.addShape(pres.shapes.OVAL, { x: x + w / 2 - 0.16, y: 5.35, w: 0.32, h: 0.32, fill: { color: C.background1 }, line: { color: C.accent1, width: 1 }, objectName: "step dot " + (i + 1) });
    if (i < flow.length - 1) s.addShape(pres.shapes.LINE, { x: x + w / 2 + 0.2, y: 5.51, w: w - 0.4, h: 0, line: { color: C.text2, width: 0.75, transparency: 50 }, objectName: "step link " + (i + 1) });
    s.addText(String(i + 1).padStart(2, "0"), { x: x + w / 2 - 0.16, y: 5.35, w: 0.32, h: 0.32, fontFace: MONO, fontSize: 10, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(a, { x, y: 5.8, w, h: 0.3, fontSize: 14, bold: true, color: C.text1, align: "center", margin: 0, isTextBox: true });
    s.addText("like " + b, { x, y: 6.1, w, h: 0.3, fontSize: 10, color: C.text2, align: "center", margin: 0, isTextBox: true });
  });
  s.addNotes("Nine steps, each mirroring a part of the human brain. Sensors are shown as x-ray lines of light rising from the plant.");
}

// ===================== 8. INSIDE THE BRAIN (example decision) =====================
{
  const s = base("LEFT", "how__5.jpg", ["ov_left.png"], "How it works");
  eyebrow(s, "INSIDE THE BRAIN · EXAMPLE", 0.7, 0.6);
  s.addText(T("From data to a ", { i: "feeling" }), { placeholder: "title" });
  const rows = [["MEMORY", "Soil moisture drops through the day: 63% → 31%"], ["JUDGE", "Light: plenty of sun · Soil: a bit dry · Air: a little warm"], ["WORLD", "Afternoon · cloudy, cool · late autumn"]];
  rows.forEach(([k, v], i) => {
    s.addText(k, { x: 0.7, y: 2.35 + i * 0.5, w: 1.3, h: 0.3, fontFace: MONO, fontSize: 11, color: C.accent1, margin: 0, isTextBox: true });
    s.addText(v, { x: 2.0, y: 2.33 + i * 0.5, w: 4.6, h: 0.32, fontSize: 14, color: C.text1, margin: 0, isTextBox: true });
  });
  s.addText("DECIDE", { x: 0.7, y: 4.0, w: 1.3, h: 0.3, fontFace: MONO, fontSize: 11, color: C.accent1, margin: 0, isTextBox: true });
  const bar = (y, label, v) => {
    s.addText(label, { x: 2.0, y, w: 1.3, h: 0.3, fontSize: 14, color: C.text1, margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.3, y: y + 0.08, w: 2.6, h: 0.14, fill: { color: C.text2, transparency: 80 }, line: { type: "none" }, objectName: label + " track" });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.3, y: y + 0.08, w: 2.6 * v, h: 0.14, fill: { color: C.accent1 }, line: { type: "none" }, objectName: label + " bar" });
    s.addText(v.toFixed(2), { x: 6.0, y, w: 0.6, h: 0.3, fontFace: MONO, fontSize: 12, color: C.text1, margin: 0, isTextBox: true });
  };
  bar(4.0, "Water me", 0.86); bar(4.45, "Rest", 0.14);
  s.addShape(pres.shapes.LINE, { x: 3.3 + 2.6 * 0.7, y: 3.95, w: 0, h: 0.85, line: { color: C.accent5, width: 1 }, objectName: "confidence threshold" });
  s.addText("sure at 0.70", { x: 4.6, y: 4.82, w: 1.6, h: 0.25, fontFace: MONO, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true });
  s.addText("VOICE", { x: 0.7, y: 5.4, w: 1.3, h: 0.3, fontFace: MONO, fontSize: 11, color: C.accent1, margin: 0, isTextBox: true });
  s.addText("“I feel so thirsty. Give me some water.”", { x: 2.0, y: 5.3, w: 4.8, h: 0.5, fontFace: SERIF, italic: true, fontSize: 22, color: C.text1, margin: 0, isTextBox: true });
  s.addText("Example decision", { x: 2.0, y: 5.9, w: 4, h: 0.25, fontFace: MONO, fontSize: 10, color: C.text2, margin: 0, isTextBox: true });
  s.addNotes("An example run: memory sees the soil drying, the soil judge reads 'a bit dry', the decision cortex is 86% sure, and the plant asks for water. Labelled as an example, not a recorded moment.");
}

// ===================== 9. TECHNOLOGY =====================
{
  const s = base("TOP", "how__68.jpg", ["ov_full.png"], "How it works");
  eyebrow(s, "UNDER THE HOOD · FULLY ON-DEVICE", 3.67, 0.5, 6, "center");
  s.addText(T("The ", { i: "technology" }), { placeholder: "title" });
  const stack = [
    ["HARDWARE", "Sensors + ESP32", "Light, soil moisture, air and pot temperature, humidity, read by an ESP32 microcontroller."],
    ["DATA", "Python logger", "Streams every reading over USB serial every 2 s and keeps a living log."],
    ["CLASSIFICATION", "Wekinator", "Neural-network classifiers over OSC: one trained judge per sense, taught on plant-science ranges."],
    ["CONTEXT", "World engine", "Python adds time of day, season and live weather to every reading."],
    ["DECISION", "Jev", "Decision cortex: weighs everything into one feeling and admits doubt below 0.70."],
    ["LANGUAGE", "Gemma 4 12B", "Runs on-device in LM Studio, carrying the plant's persona and Venus flytrap knowledge."],
    ["LISTENING", "Whisper (mlx-whisper)", "Speech-to-text on Apple Silicon: the plant's ears."],
    ["VOICE + FACE", "macOS voice + web", "macOS text-to-speech gives it a voice; a FastAPI web app draws its living line face."],
  ];
  stack.forEach(([tag, name, d], i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = 0.8 + col * 2.98, y = 2.15 + row * 2.3;
    card(s, x, y, 2.78, 2.1, "tech card " + (i + 1));
    s.addText(tag, { x: x + 0.2, y: y + 0.18, w: 2.4, h: 0.25, fontFace: MONO, fontSize: 10, color: C.accent1, charSpacing: 2, margin: 0, isTextBox: true });
    s.addText(name, { x: x + 0.2, y: y + 0.45, w: 2.4, h: 0.4, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true });
    s.addText(d, { x: x + 0.2, y: y + 0.88, w: 2.4, h: 1.1, fontSize: 12, color: C.text2, margin: 0, valign: "top", isTextBox: true });
  });
  s.addText("No cloud. Nothing leaves the room. Every feeling starts from a real sensor.", { x: 2.17, y: 6.72, w: 9, h: 0.25, fontFace: MONO, fontSize: 10, color: C.text2, align: "center", margin: 0, isTextBox: true });
  s.addNotes("The full stack, in order. Sensors on an ESP32, a Python logger, Wekinator neural-net classifiers, a Python context engine, Jev as the decision step (in this prototype it runs as a stand-in on the local model), Gemma 4 12B in LM Studio with the persona, mlx-whisper for listening, macOS text-to-speech for the voice, and a FastAPI web face. Everything runs locally.");
}

// ===================== 10. IT TALKS BACK =====================
{
  const s = base("LEFT", "how__99.jpg", ["ov_left.png"], "How it works");
  eyebrow(s, "THE INTERFACE", 0.7, 0.6);
  s.addText(T("And it ", { i: "talks back" }), { placeholder: "title" });
  body(s, "Press once and speak. It listens, thinks, and answers in its own voice, as one green line with a living face: the eye breathes, the lips move.", { x: 0.7, y: 2.15, w: 4.6, h: 1.5 });
  s.addText("LIVE DEMO ON THE WEBSITE", { x: 0.7, y: 3.85, w: 4.6, h: 0.3, fontFace: MONO, fontSize: 11, color: C.accent1, charSpacing: 3, margin: 0, isTextBox: true });
  s.addNotes("The real interface: one continuous line, a loop eye, toothed lips, and a single talk button.");
}

// ===================== 11. WHAT'S NEW =====================
pres.addSection({ title: "Why it matters" });
{
  const s = base("LEFT", "bond__75.jpg", ["ov_left.png"], "Why it matters");
  eyebrow(s, "WHAT'S NEW", 0.7, 0.6);
  s.addText(T("The first step to ", { i: "understand" }, " the language of plants"), { placeholder: "title" });
  const pts = [["Honest feelings", "Every word starts from a real sensor. Nothing is faked."], ["A brain like ours", "Built to mirror how a human mind senses, decides and speaks."], ["Its own point of view", "Curious about you; it gets to know you as you care for it."], ["Fully local", "Runs on one laptop. Private by design."]];
  pts.forEach(([a, b], i) => {
    const y = 3.0 + i * 0.9;
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: y + 0.08, w: 0.14, h: 0.14, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "point " + (i + 1) });
    s.addText(a, { x: 1.05, y, w: 5, h: 0.32, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true });
    s.addText(b, { x: 1.05, y: y + 0.32, w: 5.6, h: 0.3, fontSize: 14, color: C.text2, margin: 0, isTextBox: true });
  });
  s.addNotes("What makes this new: not a chatbot that read about plants, but a plant that answers from what it feels.");
}

// ===================== 12. DAILY LIFE =====================
{
  const s = base("TOP", "reveal__8.jpg", ["ov_full.png"], "Why it matters");
  eyebrow(s, "IN DAILY LIFE", 3.67, 0.5, 6, "center");
  s.addText(T("A plant in your ", { i: "day" }), { placeholder: "title" });
  const day = [["MORNING", "Check in", "Ask how it slept. It tells you about the light and the cold night."], ["CARE", "Know what it needs", "It asks for water when the soil is dry, in words, not numbers."], ["EVENING", "Talk", "Share your day. It is curious and remembers what you said."], ["WITH KIDS", "Learn by asking", "Children learn how living things work by talking to one."]];
  day.forEach(([k, h, d], i) => {
    const x = 0.8 + i * 2.98;
    card(s, x, 2.6, 2.78, 3.3, "day card " + (i + 1));
    s.addText(k, { x: x + 0.22, y: 2.85, w: 2.4, h: 0.25, fontFace: MONO, fontSize: 11, color: C.accent1, charSpacing: 2, margin: 0, isTextBox: true });
    s.addText(h, { x: x + 0.22, y: 3.25, w: 2.4, h: 0.7, fontSize: 18, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    s.addText(d, { x: x + 0.22, y: 4.05, w: 2.4, h: 1.6, fontSize: 14, color: C.text2, margin: 0, valign: "top", isTextBox: true });
  });
  s.addNotes("How someone would live with it, from morning to evening, and with children.");
}

// ===================== 13. VISION =====================
pres.addSection({ title: "Vision" });
{
  const s = base("TOP", "future__45.jpg", ["ov_top.png", "ov_bottom.png"], "Vision");
  eyebrow(s, "VISION", 3.67, 0.5, 6, "center");
  s.addText(T("What if the world around you could ", { i: "answer?" }), { placeholder: "title" });
  body(s, "We live through the small screens in our hands. We imagine a world where the living things around us become the interface.", { x: 2.4, y: 1.85, w: 8.53, h: 0.7, align: "center" });
  const v = [["Grow, don't buy", "The next generation grows the apple tree instead of buying the apple."], ["Talking surroundings", "Trees, gardens and parks you can ask and learn from."], ["Nature that helps", "Living spaces that sense the air, keep watch and power what we need."], ["A new perspective", "Less consuming, more connecting: nature seen as alive."]];
  v.forEach(([h, d], i) => {
    const x = 0.8 + i * 2.98;
    s.addText(h, { x, y: 5.3, w: 2.75, h: 0.35, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true });
    s.addText(d, { x, y: 5.68, w: 2.75, h: 0.9, fontSize: 14, color: C.text2, margin: 0, valign: "top", isTextBox: true });
  });
  s.addNotes("The bigger vision. The plant is a proof of concept for living surroundings that become the interface.");
}

// ===================== 14. AGAINST CONSUMPTION =====================
{
  const s = base("LEFT", "future__58.jpg", ["ov_left.png"], "Vision");
  eyebrow(s, "WHAT WE ARE CHALLENGING", 0.7, 0.6);
  s.addText(T("Less ", { i: "consumption" }), { placeholder: "title" });
  s.addText(T("Beyond ", { i: "consumerism" }), { x: 0.7, y: 1.95, w: 6.2, h: 0.8, fontSize: 38, color: C.text1, margin: 0, isTextBox: true });
  s.addText(T("Past ", { i: "capitalism" }, " as usual"), { x: 0.7, y: 2.9, w: 6.2, h: 0.8, fontSize: 38, color: C.text1, margin: 0, isTextBox: true });
  body(s, "People grow, care for and talk with the living things around them, instead of buying everything and looking at screens.", { x: 0.7, y: 4.1, w: 5.2, h: 1.0 });
  s.addNotes("What we are pushing against: consumption, consumerism and capitalism as usual.");
}

// ===================== 15. CLOSING =====================
{
  const s = base("TOP", "future__99.jpg", [], "Vision");
  s.addText(T("A better future awaits, ", { i: "with you, with us!" }), { placeholder: "title" });
  s.addText([{ text: "This plant is a proof of concept. The real project is a new way of seeing the world." }], { x: 1.5, y: 5.75, w: 10.33, h: 0.45, fontFace: SERIF, italic: true, fontSize: 20, color: C.text1, align: "center", margin: 0, isTextBox: true });
  s.addText(T("A better place to live, for ", { i: "all of us" }), { x: 2.5, y: 6.25, w: 8.33, h: 0.45, fontSize: 18, color: C.text2, align: "center", margin: 0, isTextBox: true });
  s.addNotes("Close on the Earth. A better future, with you, with us.");
}

// ===================== 16. SOURCES =====================
pres.addSection({ title: "Sources" });
{
  const s = pres.addSlide({ masterName: "LEFT", sectionTitle: "Sources" });
  s.addText("Sources", { placeholder: "title" });
  const refs = [
    "Bratman et al. (2015). Nature experience reduces rumination and subgenual prefrontal cortex activation. PNAS. doi.org/10.1073/pnas.1510459112",
    "Sudimac, Sale & Kühn (2022). How nature nurtures: amygdala activity decreases as the result of a one-hour walk in nature. Molecular Psychiatry. doi.org/10.1038/s41380-022-01720-6",
    "Berman, Jonides & Kaplan (2008). The cognitive benefits of interacting with nature. Psychological Science. doi.org/10.1111/j.1467-9280.2008.02225.x",
    "Hutton et al. (2020). Screen-based media use and brain white matter integrity in preschool-aged children. JAMA Pediatrics. doi.org/10.1001/jamapediatrics.2019.3869",
    "Ward et al. (2017). Brain drain: the mere presence of one's own smartphone reduces available cognitive capacity. JACR. doi.org/10.1086/691462",
    "Firth et al. (2019). The 'online brain': how the Internet may be changing our cognition. World Psychiatry. doi.org/10.1002/wps.20617",
    "White et al. (2019). Spending at least 120 minutes a week in nature is associated with good health and wellbeing. Scientific Reports. doi.org/10.1038/s41598-019-44097-3",
  ];
  s.addText(refs.map((r, i) => ({ text: r, options: { bullet: { type: "number" }, breakLine: i < refs.length - 1 } })), { x: 0.7, y: 1.9, w: 11.9, h: 4.8, fontSize: 13, color: C.text2, margin: 0, valign: "top", paraSpaceAfter: 9, isTextBox: true });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT, N + 1, "slides");
})();
