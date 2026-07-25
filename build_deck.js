/**
 * Kingdom Restoration Church — "Ministering the Word" training deck.
 * Run: node build_deck.js
 */
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const Fi = require("react-icons/fi");
const sharp = require("sharp");

// ── Brand palette ────────────────────────────────────────────────────────────
const PURPLE = "61116A"; // logo purple
const PLUM = "330A3A"; // deep background
const PLUM_2 = "48124F"; // raised surface on dark
const GOLD = "AE9600"; // logo gold
const GOLD_LT = "C9B356"; // gold that reads on dark
const INK = "2B2430";
const BODY = "554B5C";
const MUTE = "8C8194";
const TINT = "F6F1F7"; // pale purple card fill
const WHITE = "FFFFFF";

const A = (f) => path.join(__dirname, "assets", f);
const LOGO_WHITE = A("logo_full_white.png");
const MARK_WHITE = A("logo_mark_white.png");
const MARK_COLOR = A("logo_mark_color.png");

const H = "Calibri"; // headings
const B = "Calibri"; // body
const S = "Cambria"; // scripture

// ── Icons ────────────────────────────────────────────────────────────────────
const iconCache = new Map();
async function icon(name, color) {
  const key = name + color;
  if (iconCache.has(key)) return iconCache.get(key);
  if (!Fi[name]) throw new Error("Unknown icon: " + name);
  let svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Fi[name], { size: 256, strokeWidth: 1.7 })
  );
  svg = svg.replace(/currentColor/g, "#" + color);
  const buf = await sharp(Buffer.from(svg)).resize(320, 320).png().toBuffer();
  const data = "image/png;base64," + buf.toString("base64");
  iconCache.set(key, data);
  return data;
}

/** Line icon centred inside a thin circle — the deck's repeated motif. */
async function iconBadge(slide, name, x, y, d, ring, tone, fill) {
  const opts = { x, y, w: d, h: d, line: { color: ring, width: 1.25 } };
  if (fill) opts.fill = { color: fill };
  else opts.fill = { type: "solid", color: ring, transparency: 100 };
  slide.addShape("ellipse", opts);
  const s = d * 0.44;
  slide.addImage({
    data: await icon(name, tone),
    x: x + (d - s) / 2,
    y: y + (d - s) / 2,
    w: s,
    h: s,
  });
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "Kingdom Restoration Church";
pres.title = "The Science & Craft of Ministering the Word";

const W = 13.33;
const M = 0.85; // page margin

function darkSlide() {
  const s = pres.addSlide();
  s.background = { color: PLUM };
  return s;
}
function lightSlide() {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  return s;
}

/** Small logo mark in the corner of content slides. */
function cornerMark(slide, dark) {
  slide.addImage({
    path: dark ? MARK_WHITE : MARK_COLOR,
    x: W - M - 0.42,
    y: 6.62,
    w: 0.42,
    h: 0.36,
    transparency: dark ? 45 : 25,
  });
}

function eyebrow(slide, text, color, y) {
  slide.addText(text.toUpperCase(), {
    x: M,
    y,
    w: 8,
    h: 0.28,
    margin: 0,
    fontFace: H,
    fontSize: 11,
    bold: true,
    charSpacing: 3,
    color,
  });
}

function title(slide, text, opts = {}) {
  slide.addText(text, {
    x: M,
    y: opts.y ?? 0.95,
    w: opts.w ?? 9.6,
    h: opts.h ?? 0.95,
    margin: 0,
    fontFace: H,
    fontSize: opts.fontSize ?? 38,
    bold: true,
    color: opts.color ?? INK,
    valign: "top",
    lineSpacingMultiple: 0.95,
  });
}

async function build() {
  // ─────────────────────────────────────────── 1. Title
  {
    const s = darkSlide();
    s.addImage({
      path: MARK_WHITE,
      x: 8.55,
      y: 0.55,
      w: 5.2,
      h: 4.46,
      transparency: 90,
    });
    s.addImage({ path: LOGO_WHITE, x: M, y: 0.7, w: 3.25, h: 0.9 });

    s.addText("Word Ministry Training".toUpperCase(), {
      x: M,
      y: 2.5,
      w: 8,
      h: 0.3,
      margin: 0,
      fontFace: H,
      fontSize: 12,
      bold: true,
      charSpacing: 4,
      color: GOLD_LT,
    });
    s.addText("The Science & Craft\nof Ministering the Word", {
      x: M,
      y: 2.95,
      w: 8.9,
      h: 1.85,
      margin: 0,
      fontFace: H,
      fontSize: 40,
      bold: true,
      color: WHITE,
      lineSpacingMultiple: 1.1,
    });
    s.addText(
      "Preparing and delivering the word with clarity, flow, encouragement and impact.",
      {
        x: M,
        y: 5.05,
        w: 7.4,
        h: 0.8,
        margin: 0,
        fontFace: B,
        fontSize: 15,
        color: "CFC3D4",
        lineSpacingMultiple: 1.2,
      }
    );
    s.addText("2 Timothy 2:15", {
      x: M,
      y: 6.5,
      w: 5,
      h: 0.3,
      margin: 0,
      fontFace: S,
      italic: true,
      fontSize: 13,
      color: GOLD_LT,
    });
    s.addNotes(
      "Welcome. There is a science and a craft to ministering the word of God — it can be learned, sharpened and practised."
    );
  }

  // ─────────────────────────────────────────── 2. What people listen for
  {
    const s = lightSlide();
    eyebrow(s, "The craft", GOLD, 0.62);
    title(s, "There is a science and a craft in\nministering the word of God.", {
      y: 1.0,
      h: 1.5,
      w: 10.5,
      fontSize: 34,
    });
    s.addText("When people are listening to you, they want to hear —", {
      x: M,
      y: 2.5,
      w: 9,
      h: 0.35,
      margin: 0,
      fontFace: B,
      fontSize: 15,
      color: BODY,
    });

    const cards = [
      ["FiEye", "Clarity", "Words they can follow"],
      ["FiGitMerge", "Flow", "Each idea connects"],
      ["FiKey", "Understanding", "Truth that opens up"],
      ["FiHeart", "Encouragement", "Strength, not shame"],
      ["FiZap", "Impact", "The anointing at work"],
    ];
    const cw = 2.22,
      gap = 0.2,
      x0 = M,
      cy = 3.45,
      ch = 2.55;
    for (let i = 0; i < cards.length; i++) {
      const [ic, label, sub] = cards[i];
      const x = x0 + i * (cw + gap);
      s.addShape("roundRect", {
        x,
        y: cy,
        w: cw,
        h: ch,
        rectRadius: 0.12,
        fill: { color: TINT },
        line: { color: TINT },
      });
      await iconBadge(s, ic, x + 0.38, cy + 0.4, 0.82, GOLD, PURPLE, WHITE);
      s.addText(label, {
        x: x + 0.26,
        y: cy + 1.42,
        w: cw - 0.4,
        h: 0.32,
        margin: 0,
        fontFace: H,
        fontSize: 14.5,
        bold: true,
        color: PURPLE,
      });
      s.addText(sub, {
        x: x + 0.26,
        y: cy + 1.8,
        w: cw - 0.45,
        h: 0.6,
        margin: 0,
        fontFace: B,
        fontSize: 11.5,
        color: MUTE,
        lineSpacingMultiple: 1.15,
      });
    }
    cornerMark(s, false);
    s.addNotes(
      "These five are what a listener is really reaching for. The rest of this session takes them one at a time."
    );
  }

  // ─────────────────────────────────────────── 3. Scripture — 2 Tim 2:15
  {
    const s = darkSlide();
    s.addImage({
      path: MARK_WHITE,
      x: 9.7,
      y: 4.0,
      w: 3.4,
      h: 2.92,
      transparency: 92,
    });
    await iconBadge(s, "FiBookOpen", M, 1.05, 1.0, GOLD_LT, GOLD_LT, null);
    s.addText(
      "Be diligent to present yourself approved to God, a worker who does not need to be ashamed, rightly dividing the word of truth.",
      {
        x: M,
        y: 2.5,
        w: 10.6,
        h: 2.7,
        margin: 0,
        fontFace: S,
        fontSize: 31,
        italic: true,
        color: WHITE,
        lineSpacingMultiple: 1.32,
      }
    );
    s.addText("2 Timothy 2:15 (NKJV)".toUpperCase(), {
      x: M,
      y: 5.5,
      w: 8,
      h: 0.35,
      margin: 0,
      fontFace: H,
      fontSize: 13,
      bold: true,
      charSpacing: 3,
      color: GOLD_LT,
    });
    s.addNotes("Diligence, approval before God, and rightly dividing — the standard we are working towards.");
  }

  // ─────────────────────────────────────────── 4. Clarity
  {
    const s = lightSlide();
    eyebrow(s, "01 — Clarity", GOLD, 0.62);
    title(s, "Make it easy to understand.", { y: 1.0, h: 1.45, w: 6.5, fontSize: 36 });
    s.addText(
      "Easy, uncomplicated words. If you are confused while preparing the word, that is a strong indicator your listeners will be confused too.",
      {
        x: M,
        y: 2.55,
        w: 6.35,
        h: 1.15,
        margin: 0,
        fontFace: B,
        fontSize: 15.5,
        color: BODY,
        lineSpacingMultiple: 1.3,
      }
    );

    s.addShape("roundRect", {
      x: M,
      y: 3.9,
      w: 6.35,
      h: 2.3,
      rectRadius: 0.12,
      fill: { color: TINT },
      line: { color: TINT },
    });
    s.addText("You are competing for their attention", {
      x: M + 0.45,
      y: 4.2,
      w: 5.5,
      h: 0.32,
      margin: 0,
      fontFace: H,
      fontSize: 15,
      bold: true,
      color: PURPLE,
    });
    s.addText(
      "Every listener arrives already occupied. Clear words cut through what is competing for them.",
      {
        x: M + 0.45,
        y: 4.6,
        w: 5.45,
        h: 0.7,
        margin: 0,
        fontFace: B,
        fontSize: 12.5,
        color: BODY,
        lineSpacingMultiple: 1.2,
      }
    );
    const foes = [
      ["FiMessageCircle", "Their own thoughts"],
      ["FiWind", "Distractions"],
      ["FiShield", "Demonic powers"],
    ];
    for (let i = 0; i < foes.length; i++) {
      const x = M + 0.45 + i * 1.95;
      await iconBadge(s, foes[i][0], x, 5.4, 0.5, PURPLE, PURPLE, WHITE);
      s.addText(foes[i][1], {
        x: x + 0.6,
        y: 5.4,
        w: 1.3,
        h: 0.5,
        margin: 0,
        fontFace: B,
        fontSize: 11,
        bold: true,
        color: INK,
        valign: "middle",
      });
    }

    // Right column — the test
    s.addShape("roundRect", {
      x: 7.85,
      y: 2.55,
      w: 4.63,
      h: 3.65,
      rectRadius: 0.14,
      fill: { color: PLUM },
      line: { color: PLUM },
    });
    await iconBadge(s, "FiAlertCircle", 8.35, 2.95, 0.8, GOLD_LT, GOLD_LT, null);
    s.addText("The preparation test", {
      x: 8.35,
      y: 3.95,
      w: 3.7,
      h: 0.35,
      margin: 0,
      fontFace: H,
      fontSize: 17,
      bold: true,
      color: WHITE,
    });
    s.addText(
      [
        { text: "If it confuses you", options: { bold: true, color: GOLD_LT } },
        { text: " on the desk, it will confuse them in the pew. Sit with the point until you can say it in one plain sentence — then preach it.", options: {} },
      ],
      {
        x: 8.35,
        y: 4.4,
        w: 3.7,
        h: 1.6,
        margin: 0,
        fontFace: B,
        fontSize: 13.5,
        color: "D9CFDD",
        lineSpacingMultiple: 1.3,
      }
    );
    cornerMark(s, false);
    s.addNotes("Clarity is not shallowness. It is the discipline of understanding a thing well enough to say it simply.");
  }

  // ─────────────────────────────────────────── 5. Flow
  {
    const s = lightSlide();
    eyebrow(s, "02 — Flow", GOLD, 0.62);
    title(s, "One idea flows into the next.", { y: 1.0, w: 9 });

    const steps = [
      ["FiPlay", "Introduction", "Open the subject and show why it matters today."],
      ["FiLayers", "Body", "Build the argument point by point, each one carrying the last."],
      ["FiFlag", "Conclusion", "Land it — what they must now believe, and do."],
    ];
    const cw = 3.7,
      gap = 0.42,
      cy = 2.15,
      ch = 2.5;
    for (let i = 0; i < steps.length; i++) {
      const x = M + i * (cw + gap);
      s.addShape("roundRect", {
        x,
        y: cy,
        w: cw,
        h: ch,
        rectRadius: 0.12,
        fill: { color: i === 1 ? PURPLE : TINT },
        line: { color: i === 1 ? PURPLE : TINT },
      });
      const onDark = i === 1;
      await iconBadge(
        s,
        steps[i][0],
        x + 0.45,
        cy + 0.42,
        0.72,
        onDark ? GOLD_LT : GOLD,
        onDark ? WHITE : PURPLE,
        onDark ? null : WHITE
      );
      s.addText(steps[i][1], {
        x: x + 1.35,
        y: cy + 0.55,
        w: cw - 1.75,
        h: 0.45,
        margin: 0,
        fontFace: H,
        fontSize: 18,
        bold: true,
        color: onDark ? WHITE : PURPLE,
        valign: "middle",
      });
      s.addText(steps[i][2], {
        x: x + 0.45,
        y: cy + 1.45,
        w: cw - 0.9,
        h: 0.85,
        margin: 0,
        fontFace: B,
        fontSize: 12.5,
        color: onDark ? "E3D8E6" : BODY,
        lineSpacingMultiple: 1.2,
      });
      if (i < steps.length - 1) {
        s.addImage({
          data: await icon("FiChevronRight", GOLD),
          x: x + cw + 0.09,
          y: cy + ch / 2 - 0.12,
          w: 0.24,
          h: 0.24,
        });
      }
    }

    s.addShape("roundRect", {
      x: M,
      y: 5.05,
      w: 11.63,
      h: 1.15,
      rectRadius: 0.1,
      fill: { color: TINT },
      line: { color: TINT },
    });
    await iconBadge(s, "FiLink", M + 0.4, 5.35, 0.55, PURPLE, PURPLE, WHITE);
    s.addText(
      "One verse supports what is being said, and the interpretation of that verse supports the next. Nothing stands alone.",
      {
        x: M + 1.15,
        y: 5.35,
        w: 10.2,
        h: 0.55,
        margin: 0,
        fontFace: B,
        fontSize: 14,
        color: INK,
        valign: "middle",
      }
    );
    cornerMark(s, false);
    s.addNotes("Flow is the listener's handrail. Introduction, body, conclusion — and every verse earning its place.");
  }

  // ─────────────────────────────────────────── 6. Encouragement
  {
    const s = lightSlide();
    eyebrow(s, "03 — Encouragement", GOLD, 0.62);
    title(s, "Don't leave people condemned.", { y: 1.0, h: 1.45, w: 7.4, fontSize: 36 });
    s.addText(
      "Guilt is not your tool. Allow the Holy Spirit to convict — your work is to heal, comfort, challenge and charge.",
      {
        x: M,
        y: 2.55,
        w: 7.3,
        h: 0.85,
        margin: 0,
        fontFace: B,
        fontSize: 15.5,
        color: BODY,
        lineSpacingMultiple: 1.3,
      }
    );

    const verbs = [
      ["FiPlusCircle", "Heal", "Bind up what is broken"],
      ["FiSun", "Comfort", "Steady the weary"],
      ["FiTrendingUp", "Challenge", "Call them higher"],
      ["FiFlag", "Charge", "Send them out to act"],
    ];
    const vw = 2.72,
      vgap = 0.25;
    for (let i = 0; i < verbs.length; i++) {
      const x = M + i * (vw + vgap);
      s.addShape("roundRect", {
        x,
        y: 3.6,
        w: vw,
        h: 1.5,
        rectRadius: 0.1,
        fill: { color: TINT },
        line: { color: TINT },
      });
      await iconBadge(s, verbs[i][0], x + 0.4, 3.88, 0.62, GOLD, PURPLE, WHITE);
      s.addText(verbs[i][1], {
        x: x + 1.15,
        y: 3.88,
        w: vw - 1.4,
        h: 0.62,
        margin: 0,
        fontFace: H,
        fontSize: 15,
        bold: true,
        color: PURPLE,
        valign: "middle",
      });
      s.addText(verbs[i][2], {
        x: x + 0.4,
        y: 4.6,
        w: vw - 0.7,
        h: 0.32,
        margin: 0,
        fontFace: B,
        fontSize: 11.5,
        color: MUTE,
      });
    }

    s.addShape("roundRect", {
      x: M,
      y: 5.45,
      w: 11.63,
      h: 1.35,
      rectRadius: 0.12,
      fill: { color: PLUM },
      line: { color: PLUM },
    });
    s.addText(
      "But the one who prophesies speaks to people for their strengthening, encouraging and comfort.",
      {
        x: M + 0.5,
        y: 5.68,
        w: 9.5,
        h: 0.5,
        margin: 0,
        fontFace: S,
        italic: true,
        fontSize: 16,
        color: WHITE,
        valign: "middle",
      }
    );
    s.addText("1 Corinthians 14:3 (NIV)".toUpperCase(), {
      x: M + 0.5,
      y: 6.24,
      w: 6,
      h: 0.3,
      margin: 0,
      fontFace: H,
      fontSize: 10.5,
      bold: true,
      charSpacing: 2.5,
      color: GOLD_LT,
    });
    s.addImage({
      path: MARK_WHITE,
      x: 11.32,
      y: 5.72,
      w: 0.95,
      h: 0.81,
      transparency: 72,
    });
    s.addNotes("Conviction belongs to the Holy Spirit. Condemnation belongs to no one in the room.");
  }

  // ─────────────────────────────────────────── 7. Biases
  {
    const s = lightSlide();
    eyebrow(s, "04 — Balance", GOLD, 0.62);
    title(s, "We all have our biases.", { y: 1.0, w: 8 });
    s.addText(
      "Some of us lean on the Old Testament, some on the New. Know where you lean, and work to balance it.",
      {
        x: M,
        y: 2.02,
        w: 7.6,
        h: 0.85,
        margin: 0,
        fontFace: B,
        fontSize: 15.5,
        color: BODY,
        lineSpacingMultiple: 1.3,
      }
    );

    const cols = [
      ["FiSunrise", "Old Testament", "Covenant, law, prophets and the long story of God with His people."],
      ["FiSunset", "New Testament", "Christ, the cross, grace and the life of the early church."],
    ];
    for (let i = 0; i < 2; i++) {
      const x = M + i * 6.53;
      s.addShape("roundRect", {
        x,
        y: 3.2,
        w: 5.1,
        h: 2.5,
        rectRadius: 0.12,
        fill: { color: TINT },
        line: { color: TINT },
      });
      await iconBadge(s, cols[i][0], x + 0.45, 3.55, 0.75, GOLD, PURPLE, WHITE);
      s.addText(cols[i][1], {
        x: x + 1.35,
        y: 3.66,
        w: 3.4,
        h: 0.5,
        margin: 0,
        fontFace: H,
        fontSize: 18,
        bold: true,
        color: PURPLE,
        valign: "middle",
      });
      s.addText(cols[i][2], {
        x: x + 0.45,
        y: 4.55,
        w: 4.25,
        h: 0.9,
        margin: 0,
        fontFace: B,
        fontSize: 13,
        color: BODY,
        lineSpacingMultiple: 1.2,
      });
    }
    // balance point between the columns
    await iconBadge(s, "FiSliders", 6.27, 4.07, 0.79, PURPLE, PURPLE, WHITE);
    s.addText("Preach the whole counsel of God.", {
      x: M,
      y: 6.05,
      w: 11.63,
      h: 0.4,
      margin: 0,
      align: "center",
      fontFace: S,
      italic: true,
      fontSize: 14.5,
      color: PURPLE,
    });
    cornerMark(s, false);
    s.addNotes("Balance is not neutrality — it is refusing to feed the church only the half you enjoy.");
  }

  // ─────────────────────────────────────────── 8. Anointing
  {
    const s = darkSlide();
    s.addImage({
      path: MARK_WHITE,
      x: 9.9,
      y: 0.5,
      w: 3.2,
      h: 2.75,
      transparency: 91,
    });
    s.addText("05 — Impact".toUpperCase(), {
      x: M,
      y: 0.62,
      w: 8,
      h: 0.3,
      margin: 0,
      fontFace: H,
      fontSize: 11,
      bold: true,
      charSpacing: 3,
      color: GOLD_LT,
    });
    title(s, "The anointing you carry.", { y: 1.0, w: 8.5, color: WHITE });
    s.addText(
      "God will minister to His people through the unique way He has anointed you. Do not borrow another person's delivery — carry your own.",
      {
        x: M,
        y: 2.05,
        w: 8.3,
        h: 1.0,
        margin: 0,
        fontFace: B,
        fontSize: 15.5,
        color: "CFC3D4",
        lineSpacingMultiple: 1.3,
      }
    );

    const styles = [
      ["FiVolume2", "Some shout", "Fire and volume that stir a room"],
      ["FiFeather", "Some are soft", "A quiet voice that settles deep"],
      ["FiSmile", "Some joke", "Humour that opens hearts"],
    ];
    const cw = 3.7,
      gap = 0.42;
    for (let i = 0; i < styles.length; i++) {
      const x = M + i * (cw + gap);
      s.addShape("roundRect", {
        x,
        y: 3.55,
        w: cw,
        h: 2.35,
        rectRadius: 0.12,
        fill: { color: PLUM_2 },
        line: { color: PLUM_2 },
      });
      await iconBadge(s, styles[i][0], x + 0.45, 3.9, 0.75, GOLD_LT, GOLD_LT, null);
      s.addText(styles[i][1], {
        x: x + 0.45,
        y: 4.85,
        w: cw - 0.9,
        h: 0.35,
        margin: 0,
        fontFace: H,
        fontSize: 16,
        bold: true,
        color: WHITE,
      });
      s.addText(styles[i][2], {
        x: x + 0.45,
        y: 5.22,
        w: cw - 0.9,
        h: 0.6,
        margin: 0,
        fontFace: B,
        fontSize: 12,
        color: "BEB0C4",
        lineSpacingMultiple: 1.15,
      });
    }
    s.addText("Same Spirit. Different vessels.", {
      x: M,
      y: 6.2,
      w: 8,
      h: 0.35,
      margin: 0,
      fontFace: S,
      italic: true,
      fontSize: 14.5,
      color: GOLD_LT,
    });
    s.addNotes("Impact follows the anointing, not imitation. Preach as the person God anointed.");
  }

  // ─────────────────────────────────────────── 9. Sunday sermons
  {
    const s = lightSlide();
    eyebrow(s, "In practice", GOLD, 0.62);
    title(s, "How do you get the sermons\nfor Sundays?", { y: 1.0, h: 1.5, w: 8.6, fontSize: 34 });

    s.addShape("roundRect", {
      x: M,
      y: 2.75,
      w: 11.63,
      h: 1.1,
      rectRadius: 0.1,
      fill: { color: PURPLE },
      line: { color: PURPLE },
    });
    s.addText(
      [
        { text: "A sermon should have a purpose", options: { bold: true } },
        { text: " — and the purpose informs the content.", options: {} },
      ],
      {
        x: M + 0.5,
        y: 2.75,
        w: 10.6,
        h: 1.1,
        margin: 0,
        fontFace: H,
        fontSize: 19,
        color: WHITE,
        valign: "middle",
      }
    );

    const steps = [
      ["FiTarget", "Start with purpose", "Ask what God is saying to this house this Sunday, and what must change in the hearer."],
      ["FiLayers", "Let purpose choose content", "Text, illustrations and application are selected because they serve that purpose — not the other way round."],
    ];
    for (let i = 0; i < 2; i++) {
      const x = M + i * 6.35;
      s.addShape("roundRect", {
        x,
        y: 4.2,
        w: 5.28,
        h: 2.05,
        rectRadius: 0.12,
        fill: { color: TINT },
        line: { color: TINT },
      });
      await iconBadge(s, steps[i][0], x + 0.45, 4.5, 0.7, GOLD, PURPLE, WHITE);
      s.addText(steps[i][1], {
        x: x + 1.3,
        y: 4.58,
        w: 3.7,
        h: 0.5,
        margin: 0,
        fontFace: H,
        fontSize: 15.5,
        bold: true,
        color: PURPLE,
        valign: "middle",
      });
      s.addText(steps[i][2], {
        x: x + 0.45,
        y: 5.4,
        w: 4.4,
        h: 0.75,
        margin: 0,
        fontFace: B,
        fontSize: 12.5,
        color: BODY,
        lineSpacingMultiple: 1.2,
      });
    }
    cornerMark(s, false);
    s.addNotes("Purpose first. Content second. A sermon without a purpose becomes a collection of good verses going nowhere.");
  }

  // ─────────────────────────────────────────── 10. Close
  {
    const s = darkSlide();
    s.addImage({
      path: MARK_WHITE,
      x: 4.6,
      y: 1.5,
      w: 4.6,
      h: 3.95,
      transparency: 94,
    });
    s.addText("Rightly dividing\nthe word of truth.", {
      x: 0,
      y: 2.4,
      w: W,
      h: 1.9,
      margin: 0,
      align: "center",
      fontFace: H,
      fontSize: 42,
      bold: true,
      color: WHITE,
      lineSpacingMultiple: 1.05,
    });
    s.addText("2 Timothy 2:15 (NKJV)".toUpperCase(), {
      x: 0,
      y: 4.42,
      w: W,
      h: 0.35,
      margin: 0,
      align: "center",
      fontFace: H,
      fontSize: 12,
      bold: true,
      charSpacing: 3.5,
      color: GOLD_LT,
    });
    s.addImage({ path: LOGO_WHITE, x: (W - 3.0) / 2, y: 5.55, w: 3.0, h: 0.83 });
    s.addNotes("Close in prayer — that every worker here would be approved, unashamed, and clear.");
  }

  const out = path.join(__dirname, "Ministering-the-Word.pptx");
  await pres.writeFile({ fileName: out });
  console.log("wrote", out, fs.statSync(out).size, "bytes");
}

build().catch((e) => {
  console.error(e);
  process.exit(1);
});
