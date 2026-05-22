const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Cross-Chain Bridge Attacks — Ronin & Wormhole";
pres.author = "Security Technologies Course";

// ── PALETTE ──────────────────────────────────────────────────
const C = {
  bg:      "07090D",
  surface: "0D1117",
  card:    "111822",
  border:  "1C2A3A",
  ronin:   "FF3C3C",
  ronin2:  "FF6E6E",
  worm:    "7C6FFF",
  worm2:   "A89EFF",
  green:   "00D68F",
  amber:   "F5A623",
  white:   "FFFFFF",
  text:    "C4D4E4",
  text2:   "607A94",
  text3:   "3A5270",
};

const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.35 });

// ── HELPERS ──────────────────────────────────────────────────
function addBg(slide, color) {
  slide.background = { color: color || C.bg };
}

function card(slide, x, y, w, h, fill, opts) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: fill || C.card },
    line: { color: C.border, width: 0.5 },
    shadow: makeShadow(),
    ...opts
  });
}

function label(slide, text, x, y, w, h, color, size, opts) {
  slide.addText(text, {
    x, y, w, h,
    fontSize: size || 10,
    fontFace: "Consolas",
    color: color || C.text2,
    align: "left",
    valign: "top",
    margin: 0,
    ...opts
  });
}

function title(slide, text, x, y, w, h, size, color) {
  slide.addText(text, {
    x, y, w, h,
    fontSize: size || 40,
    fontFace: "Arial Black",
    color: color || C.white,
    align: "left",
    valign: "top",
    bold: true,
    margin: 0,
  });
}

function body(slide, text, x, y, w, h, color, size, opts) {
  slide.addText(text, {
    x, y, w, h,
    fontSize: size || 13,
    fontFace: "Calibri",
    color: color || C.text,
    align: "left",
    valign: "top",
    margin: 0,
    ...opts
  });
}

function mono(slide, text, x, y, w, h, color, size, opts) {
  slide.addText(text, {
    x, y, w, h,
    fontSize: size || 11,
    fontFace: "Consolas",
    color: color || C.text2,
    align: "left",
    valign: "top",
    margin: 0,
    ...opts
  });
}

function dot(slide, x, y, color) {
  slide.addShape(pres.shapes.OVAL, {
    x, y, w: 0.12, h: 0.12,
    fill: { color: color || C.ronin },
    line: { color: color || C.ronin, width: 0 }
  });
}

function accentBar(slide, x, y, h, color) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.06, h,
    fill: { color: color || C.ronin },
    line: { color: color || C.ronin, width: 0 }
  });
}

function statBox(slide, x, y, w, h, val, valColor, lbl) {
  card(slide, x, y, w, h, C.card);
  slide.addText(val, {
    x: x+0.15, y: y+0.12, w: w-0.3, h: 0.45,
    fontSize: 22, fontFace: "Consolas", color: valColor || C.ronin,
    bold: true, align: "center", margin: 0
  });
  slide.addText(lbl, {
    x: x+0.1, y: y+0.58, w: w-0.2, h: 0.22,
    fontSize: 9, fontFace: "Consolas", color: C.text3,
    align: "center", margin: 0, charSpacing: 0.5
  });
}

function stepNode(slide, x, y, num, text, color, subtext) {
  slide.addShape(pres.shapes.OVAL, {
    x, y: y-0.01, w: 0.36, h: 0.36,
    fill: { color: color },
    line: { color: color, width: 0 }
  });
  slide.addText(String(num), {
    x, y: y-0.01, w: 0.36, h: 0.36,
    fontSize: 13, fontFace: "Arial Black", color: C.white,
    align: "center", valign: "middle", bold: true, margin: 0
  });
  slide.addText(text, {
    x: x + 0.44, y: y, w: 2.8, h: 0.22,
    fontSize: 12, fontFace: "Calibri", color: C.white,
    bold: true, align: "left", valign: "top", margin: 0
  });
  if (subtext) {
    slide.addText(subtext, {
      x: x + 0.44, y: y + 0.22, w: 2.8, h: 0.18,
      fontSize: 10, fontFace: "Calibri", color: C.text2,
      align: "left", valign: "top", margin: 0
    });
  }
}

function flowArrow(slide, x, y, color) {
  slide.addShape(pres.shapes.LINE, {
    x, y, w: 0.55, h: 0,
    line: { color: color, width: 1.5, dashType: "solid" }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: x+0.47, y: y-0.07, w: 0.12, h: 0.14,
    fill: { color: color }, line: { color: color, width: 0 }, rotate: 45
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 1 — TITLE
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  // Grid bg effect (subtle lines as shapes)
  for (let i = 0; i < 8; i++) {
    s.addShape(pres.shapes.LINE, {
      x: i * 1.28, y: 0, w: 0, h: 5.625,
      line: { color: "1C2A3A", width: 0.3 }
    });
  }
  for (let i = 0; i < 5; i++) {
    s.addShape(pres.shapes.LINE, {
      x: 0, y: i * 1.12, w: 10, h: 0,
      line: { color: "1C2A3A", width: 0.3 }
    });
  }

  // Glow ellipse
  s.addShape(pres.shapes.OVAL, {
    x: 6.0, y: -1.2, w: 5.5, h: 5.5,
    fill: { color: "FF3C3C", transparency: 92 },
    line: { color: "FF3C3C", width: 0, transparency: 100 }
  });
  s.addShape(pres.shapes.OVAL, {
    x: 6.5, y: 1.8, w: 4, h: 4,
    fill: { color: "7C6FFF", transparency: 94 },
    line: { color: "7C6FFF", width: 0, transparency: 100 }
  });

  mono(s, "// SECURITY TECHNOLOGIES — SESSION 10 — FINAL EXAM PRESENTATION", 0.5, 0.38, 9, 0.22, C.text3, 9);
  mono(s, "ATTACK_CLASS: CROSS-CHAIN BRIDGE EXPLOITS", 0.5, 0.65, 9, 0.22, C.text3, 9, { charSpacing: 1 });

  title(s, "Cross-Chain\nBridge Attacks", 0.5, 1.0, 6.5, 2.1, 46);

  s.addText("Ronin ($625M) + Wormhole ($320M)", {
    x: 0.5, y: 3.15, w: 6.5, h: 0.42,
    fontSize: 20, fontFace: "Consolas", color: C.ronin,
    align: "left", valign: "top", margin: 0
  });

  body(s, "How $945M was stolen from two bridge protocols in 2022 —\nand why one had zero lines of code exploited.", 0.5, 3.65, 6, 0.7, C.text2, 13);

  // Stats column right
  statBox(s, 7.1, 1.0, 2.6, 1.0, "$625M", C.ronin, "RONIN BRIDGE");
  statBox(s, 7.1, 2.1, 2.6, 1.0, "$320M", C.worm, "WORMHOLE BRIDGE");
  statBox(s, 7.1, 3.2, 2.6, 1.0, "$945M+", C.ronin2, "TOTAL STOLEN");
  statBox(s, 7.1, 4.3, 2.6, 1.0, "2022", C.amber, "YEAR");

  mono(s, "SUBMIT BY: May 22, 2025  //  SESSION 10  //  COURSE: SECURITY TECHNOLOGIES", 0.5, 5.22, 9, 0.22, C.text3, 8);
}

// ══════════════════════════════════════════════════════════════
// SLIDE 2 — AGENDA / COURSE CONNECTIONS
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  mono(s, "// 00 — OVERVIEW", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  title(s, "Presentation Overview", 0.5, 0.48, 7, 0.55, 28);

  const agenda = [
    { num: "01", title: "What are cross-chain bridges?", sub: "Architecture, trust models, lock-and-mint" },
    { num: "02", title: "Ronin Bridge — $625M (March 2022)", sub: "Validator key theft, spear-phishing, stale permissions" },
    { num: "03", title: "Wormhole Bridge — $320M (Feb 2022)", sub: "Signature verification bypass, 1-line code bug" },
    { num: "04", title: "Attack comparison & taxonomy", sub: "Root causes, vulnerability classes, $2B+ landscape" },
    { num: "05", title: "Defences & mitigations", sub: "HSMs, ZK bridges, monitoring, patch ops" },
    { num: "06", title: "Course connections & takeaways", sub: "Sessions: Crypto, Network Security, Smart Contracts, OpSec" },
  ];

  agenda.forEach((a, i) => {
    const col = i < 3 ? 0 : 1;
    const row = i % 3;
    const x = col === 0 ? 0.5 : 5.3;
    const y = 1.3 + row * 1.1;
    card(s, x, y, 4.5, 0.9, C.card);
    s.addText(a.num, {
      x: x+0.18, y: y+0.12, w: 0.5, h: 0.35,
      fontSize: 16, fontFace: "Consolas", color: i < 2 ? C.ronin : i === 2 ? C.worm : C.text3,
      bold: true, align: "left", margin: 0
    });
    s.addText(a.title, {
      x: x+0.75, y: y+0.1, w: 3.5, h: 0.3,
      fontSize: 13, fontFace: "Calibri", color: C.white,
      bold: true, align: "left", margin: 0
    });
    s.addText(a.sub, {
      x: x+0.75, y: y+0.45, w: 3.5, h: 0.3,
      fontSize: 10, fontFace: "Calibri", color: C.text2,
      align: "left", margin: 0
    });
  });

  // Course connection callout
  card(s, 0.5, 4.78, 9, 0.62, "1A2535");
  mono(s, "COURSE CONNECTIONS:", 0.7, 4.9, 1.6, 0.2, C.text3, 8, { charSpacing: 0.5 });
  const connections = ["Session 3: Cryptography (signature schemes)", "Session 5: Network Security (P2P, mempool)", "Session 7: Smart Contract Security (Solidity, EVM)", "Session 9: Incident Response & Attribution"];
  connections.forEach((c, i) => {
    s.addText("■  " + c, {
      x: 0.7 + i * 2.2, y: 5.12, w: 2.1, h: 0.2,
      fontSize: 9, fontFace: "Consolas", color: C.green,
      align: "left", margin: 0
    });
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 3 — WHAT IS A CROSS-CHAIN BRIDGE
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  mono(s, "// 01 — BRIDGE ARCHITECTURE", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  title(s, "What is a Cross-Chain Bridge?", 0.5, 0.48, 9, 0.55, 26);

  body(s, "Blockchains are isolated. Ethereum cannot natively read Solana's state. A bridge locks assets on Chain A and mints a wrapped representation on Chain B. The bridge contract decides when to mint — and that authorization mechanism is the entire attack surface.", 0.5, 1.12, 9, 0.6, C.text2, 12);

  // Architecture diagram
  // Chain A box
  card(s, 0.35, 1.9, 2.2, 2.6, "0D1520");
  s.addShape(pres.shapes.LINE, { x: 0.35, y: 2.2, w: 2.2, h: 0, line: { color: C.border, width: 0.5 } });
  mono(s, "Chain A (Ethereum)", 0.55, 1.97, 1.8, 0.18, C.text3, 9);
  card(s, 0.55, 2.3, 1.8, 0.7, "1A2535");
  mono(s, "Bridge Contract", 0.65, 2.42, 1.6, 0.16, C.amber, 9, { bold: true });
  mono(s, "Locks user tokens", 0.65, 2.62, 1.6, 0.16, C.text2, 9);
  card(s, 0.55, 3.12, 1.8, 0.7, "1A2535");
  mono(s, "Token Vault", 0.65, 3.24, 1.6, 0.16, C.text2, 9, { bold: true });
  mono(s, "Holds locked assets", 0.65, 3.44, 1.6, 0.16, C.text2, 9);

  // Arrow to validator
  flowArrow(s, 2.55, 2.62, C.amber);

  // Validator layer
  card(s, 3.1, 1.9, 2.45, 2.6, "1A1500");
  s.addShape(pres.shapes.LINE, { x: 3.1, y: 2.2, w: 2.45, h: 0, line: { color: C.border, width: 0.5 } });
  mono(s, "Validator Layer", 3.3, 1.97, 2.0, 0.18, C.text3, 9);
  card(s, 3.3, 2.3, 2.05, 0.7, "261E00");
  mono(s, "Validator Set / Relayer", 3.42, 2.42, 1.8, 0.16, C.amber, 9, { bold: true });
  mono(s, "Watch + sign messages", 3.42, 2.62, 1.8, 0.16, C.text2, 9);
  card(s, 3.3, 3.12, 2.05, 0.7, "261E00");
  mono(s, "Threshold Signature", 3.42, 3.24, 1.8, 0.16, C.amber, 9, { bold: true });
  mono(s, "e.g. 5-of-9 multisig", 3.42, 3.44, 1.8, 0.16, C.text2, 9);
  mono(s, "⚠ ATTACK SURFACE", 3.3, 4.58, 2.1, 0.2, C.ronin, 9, { bold: true });

  // Arrow to Chain B
  flowArrow(s, 5.55, 2.62, C.ronin2);

  // Chain B box
  card(s, 6.1, 1.9, 2.2, 2.6, "0D1520");
  s.addShape(pres.shapes.LINE, { x: 6.1, y: 2.2, w: 2.2, h: 0, line: { color: C.border, width: 0.5 } });
  mono(s, "Chain B (Ronin/Solana)", 6.3, 1.97, 1.8, 0.18, C.text3, 9);
  card(s, 6.3, 2.3, 1.8, 0.7, "1A2535");
  mono(s, "Bridge Contract", 6.42, 2.42, 1.6, 0.16, C.text3, 9, { bold: true });
  mono(s, "Mints wrapped tokens", 6.42, 2.62, 1.6, 0.16, C.text2, 9);
  card(s, 6.3, 3.12, 1.8, 0.7, "1A2535");
  mono(s, "Wrapped Token", 6.42, 3.24, 1.6, 0.16, C.text3, 9, { bold: true });
  mono(s, "wETH, wBTC...", 6.42, 3.44, 1.6, 0.16, C.text2, 9);

  // Mega stat right
  card(s, 8.5, 1.9, 1.35, 2.6, "1A0808");
  s.addText("$2B+", { x: 8.55, y: 2.3, w: 1.25, h: 0.5, fontSize: 22, fontFace: "Consolas", color: C.ronin, bold: true, align: "center", margin: 0 });
  mono(s, "stolen from bridges in 2022 alone", 8.58, 2.88, 1.22, 0.6, C.text2, 9, { align: "center" });
  mono(s, "#1 DeFi\nvuln class", 8.6, 3.6, 1.2, 0.4, C.amber, 10, { align: "center", bold: true });

  // Bottom key point
  card(s, 0.35, 4.75, 7.95, 0.62, "1A1200");
  mono(s, "KEY INSIGHT: The validator authorization mechanism is the entire trust assumption of every bridge. Corrupt the validator layer = own the bridge.", 0.55, 4.86, 7.7, 0.38, C.amber, 11);
}

// ══════════════════════════════════════════════════════════════
// SLIDE 4 — RONIN OVERVIEW
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  s.addShape(pres.shapes.OVAL, { x: 5.5, y: -1.5, w: 6, h: 6, fill: { color: "FF3C3C", transparency: 94 }, line: { color: "FF3C3C", width: 0, transparency: 100 } });

  mono(s, "// 02 — ATTACK_01", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  s.addText("OPERATIONAL SECURITY FAILURE", { x: 0.5, y: 0.45, w: 9, h: 0.22, fontSize: 10, fontFace: "Consolas", color: C.ronin, charSpacing: 1.5, margin: 0 });
  title(s, "Ronin Bridge Hack", 0.5, 0.72, 7, 0.7, 36);
  s.addText("$625,000,000 stolen", { x: 0.5, y: 1.5, w: 5, h: 0.5, fontSize: 26, fontFace: "Consolas", color: C.ronin, bold: false, align: "left", margin: 0 });

  // Meta row
  const metas = [
    { l: "DATE", v: "March 23, 2022" },
    { l: "PROTOCOL", v: "Axie Infinity / Ronin" },
    { l: "ETH STOLEN", v: "173,600 ETH" },
    { l: "USDC STOLEN", v: "25.5M USDC" },
    { l: "DETECTION", v: "6 days later" },
    { l: "ATTACKER", v: "Lazarus Group (DPRK)" },
  ];
  metas.forEach((m, i) => {
    const x = 0.5 + (i % 3) * 3.15;
    const y = i < 3 ? 2.18 : 2.68;
    mono(s, m.l, x, y, 1.1, 0.2, C.text3, 8, { charSpacing: 0.5 });
    mono(s, m.v, x, y + 0.2, 3.0, 0.22, C.text, 12, { bold: true });
  });

  // Key facts
  const facts = [
    { num: "5-of-9", lbl: "Validator\nthreshold", color: C.amber },
    { num: "5", lbl: "Keys\ncompromised", color: C.ronin },
    { num: "2", lbl: "Exploit\ntransactions", color: C.text },
    { num: "0", lbl: "Code bugs\nexploited", color: C.green },
  ];
  facts.forEach((f, i) => {
    statBox(s, 0.5 + i * 2.38, 3.32, 2.1, 1.1, f.num, f.color, f.lbl.replace("\n", " "));
  });

  // Root cause callout
  card(s, 0.5, 4.6, 9.3, 0.73, "180808");
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 0.07, h: 0.73, fill: { color: C.ronin }, line: { color: C.ronin, width: 0 } });
  mono(s, "ROOT CAUSE:", 0.68, 4.7, 1.1, 0.18, C.ronin, 9, { bold: true, charSpacing: 0.5 });
  body(s, "No code was exploited. The smart contract worked exactly as designed. The attacker obtained 4 real validator keys via spear-phishing and a 5th via a stale permission that was never revoked after a November 2021 emergency. Classic operational security failure.", 0.68, 4.92, 9.0, 0.35, C.text2, 11);
}

// ══════════════════════════════════════════════════════════════
// SLIDE 5 — RONIN ATTACK CHAIN
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  mono(s, "// 02 — RONIN — ATTACK CHAIN", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  title(s, "How the Attack Happened", 0.5, 0.48, 9, 0.5, 26);

  const steps = [
    { n: 1, t: "Spear-phishing via LinkedIn", d: "Engineer received fake job offer PDF from Lazarus Group. PDF executed malware → 4 Sky Mavis validator keys exfiltrated.", c: C.ronin },
    { n: 2, t: "Stale DAO permission discovered", d: "Nov 2021: Sky Mavis granted emergency access to Axie DAO validator. Never revoked. Attacker exploits it as key #5.", c: C.amber },
    { n: 3, t: "Threshold reached: 5-of-9", d: "4 SM keys (phished) + 1 DAO key (stale permission) = 5/9. Exactly the withdrawal threshold. Attack is now possible.", c: C.ronin2 },
    { n: 4, t: "TX1: withdraw 173,600 ETH ($594M)", d: "March 23, 06:24 UTC. 5 valid signatures submitted. Contract verifies, approves, transfers.", c: C.ronin },
    { n: 5, t: "TX2: withdraw 25.5M USDC", d: "March 23, 06:26 UTC. Same 5 keys. Second withdrawal drains USDC. Vault now empty.", c: C.ronin },
    { n: 6, t: "Undetected for 6 days", d: "No monitoring on withdrawal events. Discovered March 29 when a user reported a failed withdrawal.", c: C.text2 },
  ];

  steps.forEach((st, i) => {
    const col = i < 3 ? 0 : 1;
    const row = i % 3;
    const x = col === 0 ? 0.35 : 5.2;
    const y = 1.15 + row * 1.3;

    card(s, x, y, 4.5, 1.1, C.card);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.1, fill: { color: st.c }, line: { color: st.c, width: 0 } });

    s.addText(String(st.n).padStart(2, "0"), {
      x: x + 0.15, y: y + 0.1, w: 0.45, h: 0.38,
      fontSize: 18, fontFace: "Consolas", color: st.c,
      bold: true, align: "left", margin: 0
    });
    s.addText(st.t, {
      x: x + 0.65, y: y + 0.1, w: 3.7, h: 0.3,
      fontSize: 12, fontFace: "Calibri", color: C.white,
      bold: true, align: "left", margin: 0
    });
    s.addText(st.d, {
      x: x + 0.65, y: y + 0.45, w: 3.7, h: 0.55,
      fontSize: 10, fontFace: "Calibri", color: C.text2,
      align: "left", valign: "top", margin: 0
    });
  });

  // Arrow between columns
  s.addShape(pres.shapes.LINE, { x: 4.87, y: 1.7, w: 0, h: 2.6, line: { color: C.border, width: 0.5, dashType: "dash" } });
  mono(s, "↓ then →", 4.72, 3.05, 0.5, 0.2, C.text3, 9, { align: "center" });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 6 — WORMHOLE OVERVIEW
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  s.addShape(pres.shapes.OVAL, { x: 5.5, y: -1.5, w: 6, h: 6, fill: { color: "7C6FFF", transparency: 94 }, line: { color: "7C6FFF", width: 0, transparency: 100 } });

  mono(s, "// 03 — ATTACK_02", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  s.addText("SIGNATURE VERIFICATION BYPASS", { x: 0.5, y: 0.45, w: 9, h: 0.22, fontSize: 10, fontFace: "Consolas", color: C.worm, charSpacing: 1.5, margin: 0 });
  title(s, "Wormhole Bridge Hack", 0.5, 0.72, 7, 0.7, 36);
  s.addText("$320,000,000 stolen", { x: 0.5, y: 1.5, w: 5, h: 0.5, fontSize: 26, fontFace: "Consolas", color: C.worm, bold: false, align: "left", margin: 0 });

  const metas = [
    { l: "DATE", v: "February 2, 2022" },
    { l: "CHAINS", v: "Solana ↔ Ethereum" },
    { l: "whETH MINTED", v: "120,000 (from nothing)" },
    { l: "ETH DEPOSITED", v: "$0" },
    { l: "FIX SIZE", v: "1 line of code" },
    { l: "BACKSTOP", v: "Jump Crypto $320M" },
  ];
  metas.forEach((m, i) => {
    const x = 0.5 + (i % 3) * 3.15;
    const y = i < 3 ? 2.18 : 2.68;
    mono(s, m.l, x, y, 1.1, 0.2, C.text3, 8, { charSpacing: 0.5 });
    mono(s, m.v, x, y + 0.2, 3.0, 0.22, C.text, 12, { bold: true });
  });

  const facts2 = [
    { num: "19", lbl: "Guardian validators", color: C.worm2 },
    { num: "0", lbl: "Guardians who signed", color: C.ronin },
    { num: "120K", lbl: "whETH minted", color: C.worm },
    { num: "1", lbl: "Line of code to fix", color: C.green },
  ];
  facts2.forEach((f, i) => {
    statBox(s, 0.5 + i * 2.38, 3.32, 2.1, 1.1, f.num, f.color, f.lbl);
  });

  card(s, 0.5, 4.6, 9.3, 0.73, "0C0818");
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 0.07, h: 0.73, fill: { color: C.worm }, line: { color: C.worm, width: 0 } });
  mono(s, "ROOT CAUSE:", 0.68, 4.7, 1.1, 0.18, C.worm, 9, { bold: true, charSpacing: 0.5 });
  body(s, "The Solana program's verify_signatures() function read signature data from a caller-supplied account — without verifying that account was owned by the real secp256k1 system program. Attacker passed a fake account with fabricated data. One missing line of Rust bypassed an entire 19-node Guardian network.", 0.68, 4.92, 9.0, 0.35, C.text2, 11);
}

// ══════════════════════════════════════════════════════════════
// SLIDE 7 — THE CODE BUG (DIFF)
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  mono(s, "// 03 — WORMHOLE — THE BUG", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  title(s, "The One-Line Bug That Cost $320M", 0.5, 0.48, 9, 0.5, 24);

  // Vulnerable panel
  card(s, 0.35, 1.1, 4.4, 3.75, "0F0808");
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.1, w: 4.4, h: 0.28, fill: { color: "3D0808" }, line: { color: C.ronin, width: 0.5 } });
  mono(s, "VULNERABLE — before patch", 0.55, 1.14, 3.5, 0.2, C.ronin2, 9, { charSpacing: 0.5 });

  const vulnCode = [
    { text: "fn verify_signatures(", c: C.text },
    { text: "  instruction_sysvar: &AccountInfo,", c: C.text },
    { text: "  guardian_set: &AccountInfo,", c: C.text },
    { text: ") -> Result<()> {", c: C.text },
    { text: "  // ❌ NO OWNER CHECK", c: "FF6E6E" },
    { text: "  // attacker passes any account", c: "FF6E6E" },
    { text: "  // with fake data", c: "FF6E6E" },
    { text: "", c: C.text },
    { text: "  let ix = load_instruction_at(", c: C.text },
    { text: "    0,", c: C.text },
    { text: "    instruction_sysvar  // ← unverified!", c: "FF6E6E" },
    { text: "  )?;", c: C.text },
    { text: "  verify(ix, guardian_set)?;", c: C.text },
    { text: "  Ok(())", c: C.text },
    { text: "}", c: C.text },
  ];
  vulnCode.forEach((line, i) => {
    mono(s, line.text, 0.5, 1.46 + i * 0.2, 4.1, 0.21, line.c, 10);
  });

  // Fixed panel
  card(s, 5.05, 1.1, 4.4, 3.75, "080F08");
  s.addShape(pres.shapes.RECTANGLE, { x: 5.05, y: 1.1, w: 4.4, h: 0.28, fill: { color: "0D3D0D" }, line: { color: C.green, width: 0.5 } });
  mono(s, "FIXED — after patch (1 line added)", 5.25, 1.14, 3.8, 0.2, C.green, 9, { charSpacing: 0.5 });

  const fixedCode = [
    { text: "fn verify_signatures(", c: C.text },
    { text: "  instruction_sysvar: &AccountInfo,", c: C.text },
    { text: "  guardian_set: &AccountInfo,", c: C.text },
    { text: ") -> Result<()> {", c: C.text },
    { text: "  // ✅ ONE LINE FIX:", c: "00D68F" },
    { text: "  check_account_owner(", c: "00D68F" },
    { text: "    instruction_sysvar,", c: "00D68F" },
    { text: "    &sysvar::id(),  // real owner check", c: "00D68F" },
    { text: "    WormholeError::InvalidSysvar)?;", c: "00D68F" },
    { text: "  let ix = load_instruction_at(", c: C.text },
    { text: "    0,", c: C.text },
    { text: "    instruction_sysvar  // ← now verified", c: "00D68F" },
    { text: "  )?;", c: C.text },
    { text: "  verify(ix, guardian_set)?;", c: C.text },
    { text: "  Ok(()) // ← only with real Guardian sigs", c: C.text },
  ];
  fixedCode.forEach((line, i) => {
    mono(s, line.text, 5.22, 1.46 + i * 0.2, 4.1, 0.21, line.c, 10);
  });

  // Bottom insight
  card(s, 0.35, 5.1, 9.3, 0.33, "1A1200");
  mono(s, "The attacker read the undeployed patch on GitHub and raced to exploit before it went live. Ops lesson: NEVER push security patches to public repos before deployment.", 0.55, 5.16, 9.0, 0.22, C.amber, 10, { bold: true });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 8 — SIDE-BY-SIDE COMPARISON
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  mono(s, "// 04 — COMPARISON", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  title(s, "Ronin vs Wormhole — Same Class, Different Root Cause", 0.5, 0.48, 9.3, 0.55, 22);

  const rows = [
    { k: "Attack class", r: "OpSec / key theft", w: "Signature verification bypass" },
    { k: "Vector", r: "LinkedIn spear-phishing → malware", w: "Missing Rust account owner check" },
    { k: "Code exploited?", r: "No — contract was correct", w: "Yes — one missing check", rC: C.green, wC: C.ronin },
    { k: "Crypto broken?", r: "No — signatures were real keys", w: "No — check simply bypassed", rC: C.green, wC: C.green },
    { k: "Detection time", r: "6 days", w: "Hours", rC: C.ronin2, wC: C.amber },
    { k: "Attacker", r: "Lazarus Group (DPRK)", w: "Unknown (sophisticated)" },
    { k: "Recovery", r: "Sky Mavis raised $150M", w: "Jump Crypto deposited 120K ETH" },
    { k: "Key fix", r: "HSMs + monitor + revoke stale perms", w: "1-line: check_account_owner()" },
  ];

  // Headers
  card(s, 0.35, 1.12, 3.0, 0.3, "1A0808");
  mono(s, "Ronin — $625M", 0.55, 1.18, 2.7, 0.2, C.ronin, 10, { bold: true, charSpacing: 0.5 });
  card(s, 6.7, 1.12, 3.0, 0.3, "0C0818");
  mono(s, "Wormhole — $320M", 6.9, 1.18, 2.7, 0.2, C.worm, 10, { bold: true, charSpacing: 0.5 });

  rows.forEach((row, i) => {
    const y = 1.52 + i * 0.48;
    const isEven = i % 2 === 0;

    // Key cell
    card(s, 0.35, y, 2.55, 0.42, isEven ? "0D1520" : C.card);
    mono(s, row.k, 0.5, y + 0.1, 2.3, 0.22, C.text3, 10);

    // Ronin value
    card(s, 2.95, y, 3.65, 0.42, isEven ? "160A0A" : C.card);
    mono(s, row.r, 3.08, y + 0.1, 3.4, 0.22, row.rC || C.text, 10);

    // Center equal sign
    mono(s, "≠", 6.55, y + 0.1, 0.18, 0.22, C.border, 13, { align: "center" });

    // Wormhole value
    card(s, 6.75, y, 3.0, 0.42, isEven ? "0A0A16" : C.card);
    mono(s, row.w, 6.88, y + 0.1, 2.8, 0.22, row.wC || C.text, 10);
  });

  // Bottom callout
  card(s, 0.35, 5.42, 9.3, 0.0);
  s.addShape(pres.shapes.LINE, { x: 0.35, y: 5.42, w: 9.3, h: 0, line: { color: C.border, width: 0.5 } });
  mono(s, "KEY POINT: Both attacks succeeded without breaking any cryptography. One was a human failure, one a code failure — together they cover the full bridge attack surface.", 0.5, 5.38, 9.0, 0.22, C.amber, 10);
}

// ══════════════════════════════════════════════════════════════
// SLIDE 9 — BRIDGE HACKS CHART + TAXONOMY
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  mono(s, "// 04 — BRIDGE HACK LANDSCAPE", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  title(s, "The Bigger Picture — $2B+ in 2022 Alone", 0.5, 0.48, 9, 0.5, 24);

  // Bar chart
  s.addChart(pres.charts.BAR, [{
    name: "Stolen ($M)",
    labels: ["Poly\nNetwork", "Ronin\nBridge", "BNB\nBridge", "Wormhole", "Nomad", "Harmony", "Multi-\nchain"],
    values: [611, 625, 570, 320, 190, 100, 126]
  }], {
    x: 0.35, y: 1.05, w: 5.5, h: 3.5,
    barDir: "bar",
    chartColors: ["F5A623", "FF3C3C", "F5A623", "7C6FFF", "F5A623", "FF6E6E", "607A94"],
    chartArea: { fill: { color: "0D1117" }, roundedCorners: false },
    catAxisLabelColor: "607A94",
    valAxisLabelColor: "607A94",
    catGridLine: { style: "none" },
    valGridLine: { color: "1C2A3A", size: 0.5 },
    showValue: true,
    dataLabelColor: "C4D4E4",
    dataLabelFontSize: 9,
    showLegend: false,
  });

  // Vulnerability taxonomy right side
  const vulns = [
    { cls: "Key compromise", ex: "Ronin, Harmony", loss: "$725M", c: C.ronin },
    { cls: "Sig verification bypass", ex: "Wormhole", loss: "$320M", c: C.worm },
    { cls: "Access control failure", ex: "Poly Network", loss: "$611M", c: C.amber },
    { cls: "Uninit state / default trust", ex: "Nomad", loss: "$190M", c: C.ronin2 },
    { cls: "Message replay", ex: "Various", loss: "varies", c: C.text2 },
  ];

  mono(s, "VULNERABILITY TAXONOMY", 6.1, 1.05, 3.7, 0.2, C.text3, 9, { charSpacing: 0.5 });
  vulns.forEach((v, i) => {
    card(s, 6.1, 1.32 + i * 0.62, 3.65, 0.54, C.card);
    s.addShape(pres.shapes.RECTANGLE, { x: 6.1, y: 1.32 + i * 0.62, w: 0.05, h: 0.54, fill: { color: v.c }, line: { color: v.c, width: 0 } });
    mono(s, v.cls, 6.24, 1.38 + i * 0.62, 2.7, 0.18, C.white, 10, { bold: true });
    mono(s, v.ex, 6.24, 1.57 + i * 0.62, 1.9, 0.16, C.text2, 9);
    mono(s, v.loss, 8.15, 1.38 + i * 0.62, 1.5, 0.18, v.c, 11, { align: "right", bold: true });
  });

  // Bottom total
  card(s, 0.35, 4.75, 9.4, 0.6, "180A0A");
  s.addText("$2,000,000,000+", { x: 0.55, y: 4.83, w: 3.2, h: 0.44, fontSize: 24, fontFace: "Consolas", color: C.ronin, bold: true, align: "left", margin: 0 });
  body(s, "stolen from cross-chain bridges in 2022 alone — more than any other DeFi vulnerability class that year.", 3.85, 4.88, 5.7, 0.4, C.text2, 11);
}

// ══════════════════════════════════════════════════════════════
// SLIDE 10 — DEFENCES
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  mono(s, "// 05 — DEFENCES & MITIGATIONS", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  title(s, "How to Defend Against Bridge Attacks", 0.5, 0.48, 9, 0.5, 26);

  const defences = [
    { num: "01", t: "Hardware Security Modules (HSMs)", d: "Validator private keys stored in HSMs or air-gapped machines. Keys must never touch internet-connected systems. Eliminates the phishing vector entirely.", c: C.ronin, tag: "Fixes Ronin" },
    { num: "02", t: "Account ownership checks (Solana)", d: "On Solana and similar VMs, always verify that system accounts (sysvars) are owned by the expected program — not just that they have valid-looking data.", c: C.worm, tag: "Fixes Wormhole" },
    { num: "03", t: "Time-locked permission expiry", d: "Every 'temporary' access grant must auto-expire. No manual revocation. The Axie DAO permission that cost $625M was never revoked after the emergency ended.", c: C.amber, tag: "Fixes Ronin" },
    { num: "04", t: "Real-time withdrawal monitoring", d: "Alert on large withdrawals. Circuit breakers that pause the bridge above a threshold. $625M sat undetected for 6 days because nobody watched on-chain events.", c: C.amber, tag: "Fixes Ronin" },
    { num: "05", t: "Private patch deployment", d: "Security patches must be deployed before public disclosure. Wormhole's attacker read the undeployed fix on public GitHub and exploited the vulnerability first.", c: C.worm, tag: "Fixes Wormhole" },
    { num: "06", t: "ZK light client bridges (long-term)", d: "Replace trusted validators with mathematical proofs. ZK bridges verify cross-chain state cryptographically — eliminating the human trust assumption entirely.", c: C.green, tag: "Architecture fix" },
  ];

  defences.forEach((d, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.35 : 5.2;
    const y = 1.15 + row * 1.3;

    card(s, x, y, 4.5, 1.15, C.card);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 1.15, fill: { color: d.c }, line: { color: d.c, width: 0 } });

    s.addText(d.num, { x: x+0.15, y: y+0.1, w: 0.45, h: 0.3, fontSize: 14, fontFace: "Consolas", color: d.c, bold: true, align: "left", margin: 0 });
    s.addText(d.t, { x: x+0.65, y: y+0.1, w: 3.0, h: 0.28, fontSize: 11, fontFace: "Calibri", color: C.white, bold: true, align: "left", margin: 0 });
    s.addText("[" + d.tag + "]", { x: x+3.7, y: y+0.1, w: 0.72, h: 0.28, fontSize: 8, fontFace: "Consolas", color: d.c, align: "right", margin: 0 });
    s.addText(d.d, { x: x+0.65, y: y+0.44, w: 3.75, h: 0.62, fontSize: 10, fontFace: "Calibri", color: C.text2, align: "left", valign: "top", margin: 0 });
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 11 — COURSE CONNECTIONS
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  mono(s, "// 06 — COURSE CONNECTIONS", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  title(s, "How This Connects to the Course", 0.5, 0.48, 9, 0.5, 26);

  const conns = [
    {
      session: "Session 3", t: "Cryptography",
      link: "Signature schemes (secp256k1) underlie both attacks. The Wormhole bug bypassed secp256k1 verification entirely — not by breaking the math, but by never calling it. The Ronin attack used real valid signatures from stolen keys. Understanding ECDSA, threshold signatures, and multisig is prerequisite knowledge.",
      c: C.worm
    },
    {
      session: "Session 5", t: "Network Security",
      link: "The Ronin bridge had no anomaly detection on its withdrawal events despite all transactions being public on-chain. Network monitoring, SIEM alerting, and circuit-breaker patterns apply directly. The 'dark forest' mempool problem (MEV, frontrunning bridge messages) connects to adversarial network environments.",
      c: C.amber
    },
    {
      session: "Session 7", t: "Smart Contract Security",
      link: "Access control failures (unprotected functions, stale permissions) and input validation bugs (missing owner checks) are foundational smart contract vulnerabilities. The Wormhole bug is a direct application of 'never trust caller-supplied accounts without verifying ownership' — a core Solana security principle.",
      c: C.ronin
    },
    {
      session: "Session 9", t: "Incident Response & Attribution",
      link: "Ronin's 6-day detection gap is a case study in inadequate monitoring. The FBI attribution of Lazarus Group via on-chain wallet clustering and malware code signatures demonstrates threat actor profiling. Jump Crypto's rapid $320M replenishment shows crisis response decision-making under pressure.",
      c: C.green
    },
  ];

  conns.forEach((c, i) => {
    const y = 1.1 + i * 1.06;
    card(s, 0.35, y, 9.3, 0.92, C.card);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 0.06, h: 0.92, fill: { color: c.c }, line: { color: c.c, width: 0 } });
    s.addText(c.session, { x: 0.52, y: y+0.08, w: 1.0, h: 0.22, fontSize: 10, fontFace: "Consolas", color: c.c, bold: true, align: "left", margin: 0 });
    s.addText(c.t, { x: 1.6, y: y+0.08, w: 2.5, h: 0.22, fontSize: 13, fontFace: "Calibri", color: C.white, bold: true, align: "left", margin: 0 });
    s.addText(c.link, { x: 0.52, y: y+0.38, w: 9.0, h: 0.48, fontSize: 10.5, fontFace: "Calibri", color: C.text2, align: "left", valign: "top", margin: 0 });
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 12 — KEY TAKEAWAYS
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  s.addShape(pres.shapes.OVAL, { x: 3.5, y: -0.5, w: 7, h: 7, fill: { color: "FF3C3C", transparency: 95 }, line: { color: "FF3C3C", width: 0, transparency: 100 } });

  mono(s, "// 06 — KEY TAKEAWAYS", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  title(s, "What $945M Teaches Us", 0.5, 0.48, 9, 0.5, 28);

  const tks = [
    { n: "01", t: "Keys are the weakest link", d: "$625M stolen without a single line of code being wrong. Private keys on internet-connected machines are the most exploited vector in crypto security.", c: C.ronin },
    { n: "02", t: "One line can cost $320M", d: "A single missing check_account_owner() bypassed 19 Guardian validators. The smallest code omission in a critical path can have catastrophic consequences.", c: C.worm },
    { n: "03", t: "\"Temporary\" means permanent", d: "The Axie DAO permission was the 5th key Lazarus needed. Every emergency access grant needs a hard-coded expiry date — no exceptions.", c: C.amber },
    { n: "04", t: "If nobody's watching, nobody knows", d: "$625M sat on-chain, publicly visible, for 6 days. On-chain transparency only protects you if someone is actually monitoring for anomalies.", c: C.amber },
    { n: "05", t: "Deploy patches before disclosing", d: "Wormhole's attacker read the GitHub fix before deployment and raced to exploit it first. Security patches for live protocols are not open-source code.", c: C.green },
  ];

  tks.forEach((tk, i) => {
    const x = i < 3 ? 0.35 : (i === 3 ? 0.35 : 5.2);
    const y = i < 3 ? 1.15 + i * 1.08 : (i === 3 ? 1.15 + 3 * 1.08 : 1.15 + 3 * 1.08);
    const actualX = i < 2 ? 0.35 : i === 2 ? 0.35 : i === 3 ? 0.35 : 5.2;
    const actualY = i < 3 ? 1.15 + i * 1.16 : 4.63;
    const w = i >= 3 ? 4.45 : 9.3;

    card(s, actualX, actualY, w, 1.0, C.card);
    s.addShape(pres.shapes.RECTANGLE, { x: actualX, y: actualY, w: 0.06, h: 1.0, fill: { color: tk.c }, line: { color: tk.c, width: 0 } });
    s.addText(tk.n, { x: actualX+0.15, y: actualY+0.08, w: 0.45, h: 0.3, fontSize: 16, fontFace: "Consolas", color: tk.c, bold: true, align: "left", margin: 0 });
    s.addText(tk.t, { x: actualX+0.68, y: actualY+0.08, w: w-0.85, h: 0.28, fontSize: 13, fontFace: "Calibri", color: C.white, bold: true, align: "left", margin: 0 });
    s.addText(tk.d, { x: actualX+0.68, y: actualY+0.42, w: w-0.85, h: 0.5, fontSize: 10.5, fontFace: "Calibri", color: C.text2, align: "left", valign: "top", margin: 0 });
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 13 — BIBLIOGRAPHY
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  mono(s, "// 07 — BIBLIOGRAPHY", 0.5, 0.22, 9, 0.2, C.text3, 9, { charSpacing: 1 });
  title(s, "Sources & Bibliography", 0.5, 0.48, 9, 0.5, 28);

  const sources = [
    { n: "[1]", t: "Sky Mavis — Ronin Network Bridge Exploit Post-Mortem", m: "Sky Mavis Official Blog · March 2022", tag: "PRIMARY", c: C.ronin },
    { n: "[2]", t: "FBI / US Treasury OFAC — Attribution to Lazarus Group (DPRK)", m: "US Treasury Sanctions Notice · April 14, 2022", tag: "GOVT", c: C.ronin },
    { n: "[3]", t: "Wormhole — Incident Report: 120,000 wETH Exploit", m: "Wormhole Official Blog · February 3, 2022", tag: "PRIMARY", c: C.worm },
    { n: "[4]", t: "GitHub — Wormhole fix commit a2640a3: 'add sysvar account owner check'", m: "Wormhole GitHub Repository · January 2022", tag: "CODE", c: C.worm },
    { n: "[5]", t: "Neodyme Security — Technical Analysis of the Wormhole Exploit", m: "Neodyme Security Blog · February 2022", tag: "TECHNICAL", c: C.worm },
    { n: "[6]", t: "Vitalik Buterin — Why cross-chain bridges are fundamentally limited in security", m: "vitalik.ca · January 2022", tag: "RESEARCH", c: C.amber },
    { n: "[7]", t: "Chainalysis — 2023 Crypto Crime Report (Bridge Hacks section)", m: "Chainalysis Annual Report · 2023", tag: "REPORT", c: C.amber },
    { n: "[8]", t: "UN Panel of Experts — DPRK Cyber Operations and Sanctions Evasion", m: "UN Security Council · 2023", tag: "GOVT", c: C.text2 },
    { n: "[9]", t: "Qin et al. — Quantifying Blockchain Extractable Value (IEEE S&P 2022)", m: "IEEE Security & Privacy · 2022", tag: "PAPER", c: C.green },
    { n: "[10]", t: "Jump Crypto — Statement on 120,000 ETH Replenishment", m: "Jump Crypto Official Statement · February 2022", tag: "PRIMARY", c: C.green },
  ];

  const tagColors = { PRIMARY: C.ronin, GOVT: C.amber, CODE: C.worm, TECHNICAL: C.worm, RESEARCH: C.amber, REPORT: C.amber, PAPER: C.green };

  sources.forEach((src, i) => {
    const col = i < 5 ? 0 : 1;
    const row = i % 5;
    const x = col === 0 ? 0.35 : 5.2;
    const y = 1.1 + row * 0.85;

    mono(s, src.n, x, y + 0.03, 0.38, 0.2, C.text3, 9);
    s.addText(src.t, { x: x+0.42, y: y, w: 4.0, h: 0.26, fontSize: 11, fontFace: "Calibri", color: C.white, bold: true, align: "left", margin: 0 });
    s.addText("[" + src.tag + "]", { x: x+4.4, y: y + 0.02, w: 0.55, h: 0.22, fontSize: 8, fontFace: "Consolas", color: tagColors[src.tag] || C.text2, align: "right", margin: 0 });
    s.addText(src.m, { x: x+0.42, y: y + 0.28, w: 4.5, h: 0.2, fontSize: 9, fontFace: "Calibri", color: C.text2, align: "left", margin: 0 });
    s.addShape(pres.shapes.LINE, { x: x, y: y + 0.55, w: 4.6, h: 0, line: { color: C.border, width: 0.3 } });
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 14 — Q&A / CLOSING
// ══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s, C.bg);

  for (let i = 0; i < 8; i++) {
    s.addShape(pres.shapes.LINE, { x: i * 1.28, y: 0, w: 0, h: 5.625, line: { color: "1C2A3A", width: 0.3 } });
  }
  for (let i = 0; i < 5; i++) {
    s.addShape(pres.shapes.LINE, { x: 0, y: i * 1.12, w: 10, h: 0, line: { color: "1C2A3A", width: 0.3 } });
  }

  s.addShape(pres.shapes.OVAL, { x: -1, y: -1, w: 6, h: 6, fill: { color: "FF3C3C", transparency: 92 }, line: { color: "FF3C3C", width: 0, transparency: 100 } });
  s.addShape(pres.shapes.OVAL, { x: 5, y: 1, w: 6, h: 6, fill: { color: "7C6FFF", transparency: 93 }, line: { color: "7C6FFF", width: 0, transparency: 100 } });

  s.addText("Q&A", { x: 0.5, y: 1.0, w: 9, h: 2.2, fontSize: 110, fontFace: "Arial Black", color: "FFFFFF", align: "center", valign: "middle", bold: true, transparency: 6, margin: 0 });

  mono(s, "CROSS-CHAIN BRIDGE ATTACKS: RONIN ($625M) + WORMHOLE ($320M)", 0.5, 3.5, 9, 0.2, C.text3, 9, { align: "center", charSpacing: 1 });

  const qas = [
    "Why 5-of-9 wasn't enough to stop Ronin?",
    "Why didn't Wormhole Guardians catch the fake VAA?",
    "How does ZK bridging prevent these attacks?",
    "Could monitoring have stopped either attack in time?",
  ];

  qas.forEach((q, i) => {
    s.addText("Q: " + q, {
      x: i < 2 ? 0.35 : 5.0, y: 3.9 + (i % 2) * 0.42, w: 4.5, h: 0.35,
      fontSize: 10, fontFace: "Consolas", color: C.text2, align: "left", margin: 0
    });
  });

  mono(s, "Security Technologies Course — Session 10", 0.5, 5.28, 9, 0.2, C.text3, 9, { align: "center" });
}

// ── WRITE FILE ──────────────────────────────────────────────
async function createPPT() {
  try {
    await pres.writeFile({
      fileName: "./bridge_attacks_presentation.pptx"
    });
    console.log("✅ Saved");
  } catch (e) {
    console.error("❌ Error:", e);
  }
}

createPPT();