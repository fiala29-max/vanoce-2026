const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/b533e471-6c8c-470b-86cd-06ab5a24936b_22d7d571-0632-44d6-815d-7951b00882cb/pptx/scripts/apply_theme.js");

const THEME = {
  name: "Vanoce",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1F2A24", lt1: "FFFFFF", dk2: "1B4332", lt2: "F1F5F2",
    accent1: "1B4332", accent2: "B3192B", accent3: "D4A017",
    accent4: "2D6A4F", accent5: "E9C46A", accent6: "7A1C26",
    hlink: "B3192B", folHlink: "7A1C26",
  },
};
const HEX = THEME.colors;

async function icon(name, color) {
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = "Program na Vánoční svátky 2026";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "Titul",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 11.7, h: 1.6, fontSize: 48, bold: true, color: C.background1, align: "center", valign: "middle" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.0, w: 11.7, h: 1.0, fontSize: 22, color: C.accent5, align: "center", valign: "top" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "Obsah",
    background: { color: C.background1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 10.5, h: 0.9, fontSize: 36, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.25, w: 10.5, h: 0.5, fontSize: 16, italic: true, color: C.accent2, valign: "top", margin: 0 }, text: "" } },
      { text: { text: "Vánoční svátky 2026", options: { x: 0.6, y: 7.0, w: 6, h: 0.3, fontSize: 10, color: "6B7A72", margin: 0 } } },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: "6B7A72", align: "right" },
  });

  // icon cache
  const ic = {};
  const need = [
    ["FaTree", HEX.lt1], ["FaGift", HEX.lt1], ["FaSnowflake", HEX.lt1], ["FaStar", HEX.accent3],
    ["FaUtensils", HEX.lt1], ["FaChurch", HEX.lt1], ["FaSkating", HEX.lt1], ["FaCookieBite", HEX.lt1],
    ["FaShoppingBasket", HEX.lt1], ["FaGlassCheers", HEX.lt1], ["FaWalking", HEX.lt1], ["FaFish", HEX.lt1],
    ["FaBell", HEX.lt1], ["FaUsers", HEX.lt1], ["FaFilm", HEX.lt1], ["FaDice", HEX.lt1], ["FaSun", HEX.lt1],
    ["FaCheck", HEX.lt1], ["FaHome", HEX.lt1], ["FaMugHot", HEX.lt1], ["FaSleigh", HEX.lt1], ["FaCandyCane", HEX.lt1],
    ["FaSnowflake_g", HEX.accent4],
  ];
  for (const [n, col] of need) ic[n] = await icon(n.replace("_g", ""), col);

  const circleIcon = (slide, name, x, y, d, fill) => {
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: "icon-bg-" + name });
    const p = d * 0.26;
    slide.addImage({ data: ic[name], x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: name, objectName: "icon-" + name });
  };
  const card = (slide, x, y, w, h, fill) =>
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" },
      shadow: { type: "outer", color: "1F2A24", opacity: 0.12, blur: 6, offset: 2, angle: 90 }, objectName: "card" });

  // decorative snowflakes for dark slides
  const flakes = (slide) => {
    [[0.7, 0.6, 0.5], [11.9, 0.9, 0.7], [1.6, 5.9, 0.4], [11.2, 5.6, 0.55], [6.4, 0.5, 0.35]].forEach(([x, y, s], i) =>
      slide.addImage({ data: ic.FaSnowflake, x, y, w: s, h: s, transparency: 70, altText: "Vločka", objectName: "flake" + i }));
  };

  // ---------- 1. Title
  pres.addSection({ title: "Úvod" });
  let s = pres.addSlide({ masterName: "Titul", sectionTitle: "Úvod" });
  flakes(s);
  s.addImage({ data: ic.FaStar, x: 6.22, y: 1.3, w: 0.9, h: 0.9, altText: "Hvězda", objectName: "star" });
  s.addText("Program na Vánoční svátky", { placeholder: "title" });
  s.addText("23. 12. 2026 – 1. 1. 2027  ·  Klid, rodina, tradice a trocha pohybu", { placeholder: "body" });
  s.addNotes("Úvodní slide. Program je navržen pro rodinu s dětmi, dá se volně upravit.");

  // ---------- 2. Overview timeline
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Úvod" });
  s.addText("Svátky v kostce", { placeholder: "title" });
  s.addText("Deset dní, sedm zastávek", { placeholder: "body" });
  const tl = [
    ["St 23. 12.", "Přípravy", "FaShoppingBasket", C.accent4],
    ["Čt 24. 12.", "Štědrý den", "FaGift", C.accent2],
    ["Pá 25. 12.", "Boží hod", "FaUtensils", C.accent1],
    ["So 26. 12.", "Štěpán", "FaBell", C.accent4],
    ["27.–30. 12.", "Mezi svátky", "FaSkating", C.accent1],
    ["Čt 31. 12.", "Silvestr", "FaGlassCheers", C.accent2],
    ["Pá 1. 1.", "Nový rok", "FaSun", C.accent4],
  ];
  const tx0 = 0.9, tStep = 1.9, ty = 3.3;
  s.addShape(pres.shapes.LINE, { x: tx0 + 0.45, y: ty + 0.45, w: tStep * 6, h: 0, line: { color: C.accent5, width: 4 }, objectName: "timeline" });
  // line color on light bg: use accent5 for visibility
  tl.forEach(([d, t, n, col], i) => {
    const x = tx0 + i * tStep;
    circleIcon(s, n, x, ty, 0.9, col);
    s.addText(d, { x: x - 0.5, y: ty - 0.75, w: 1.9, h: 0.4, fontSize: 14, bold: true, color: C.text2, align: "center", isTextBox: true });
    s.addText(t, { x: x - 0.5, y: ty + 1.05, w: 1.9, h: 0.4, fontSize: 16, color: C.text1, align: "center", isTextBox: true });
  });
  s.addText("Tip: každý den má jednu „hlavní“ věc a zbytek je volný čas — nic není povinné.", {
    x: 0.6, y: 5.7, w: 12.1, h: 0.5, fontSize: 15, italic: true, color: C.accent4, isTextBox: true, margin: 0 });

  // ---------- 3. Preparations
  pres.addSection({ title: "Program" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Program" });
  s.addText("Přípravy do 23. prosince", { placeholder: "title" });
  s.addText("Ať je na Štědrý den hotovo a nikdo nespěchá", { placeholder: "body" });
  const prep = [
    ["FaCookieBite", "Cukroví", "Do 20. 12. napéct 4–5 druhů: vanilkové rohlíčky, linecké, perníčky, vosí hnízda, pracky."],
    ["FaShoppingBasket", "Nákupy", "Dárky zabalit do 22. 12. Potraviny na svátky koupit 22.–23. 12., ať jsou čerstvé."],
    ["FaFish", "Kapr", "Kapra (nebo filé) koupit 23. 12. Bramborový salát připravit už večer 23. 12."],
    ["FaTree", "Stromeček", "Ozdobit 23. 12. odpoledne společně s dětmi, k tomu koledy a horký čaj."],
  ];
  prep.forEach(([n, h, t], i) => {
    const x = 0.6 + (i % 2) * 6.15, y = 2.05 + Math.floor(i / 2) * 2.35;
    card(s, x, y, 5.95, 2.05, C.background2);
    circleIcon(s, n, x + 0.3, y + 0.35, 0.85, i % 3 === 0 ? C.accent2 : C.accent1);
    s.addText(h, { x: x + 1.4, y: y + 0.3, w: 4.3, h: 0.45, fontSize: 20, bold: true, color: C.text2, isTextBox: true, margin: 0 });
    s.addText(t, { x: x + 1.4, y: y + 0.8, w: 4.3, h: 1.1, fontSize: 15, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
  });

  // ---------- 4. Štědrý den schedule
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Program" });
  s.addText("Štědrý den · čtvrtek 24. 12.", { placeholder: "title" });
  s.addText("Hlavní den svátků — večeře, dárky a staré zvyky", { placeholder: "body" });
  const sched = [
    ["9:00", "Snídaně a poslední úpravy stromečku"],
    ["10:30", "Procházka, krmení zvířátek v lese"],
    ["12:00", "Lehký oběd — kdo vydrží půst, uvidí zlaté prasátko"],
    ["14:00", "Pohádky v TV, balení posledních dárků"],
    ["17:00", "Štědrovečerní večeře: rybí polévka, kapr a salát"],
    ["18:30", "Zvoneček a rozbalování dárků"],
    ["23:30", "Půlnoční mše (pro zájemce)"],
  ];
  sched.forEach(([t, d], i) => {
    const y = 1.95 + i * 0.68;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 1.15, h: 0.5, rectRadius: 0.1, fill: { color: i === 4 || i === 5 ? C.accent2 : C.accent1 }, line: { type: "none" }, objectName: "time" + i });
    s.addText(t, { x: 0.6, y, w: 1.15, h: 0.5, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(d, { x: 1.95, y, w: 5.4, h: 0.5, fontSize: 16, color: C.text1, valign: "middle", isTextBox: true, margin: 0 });
  });
  card(s, 7.8, 1.95, 4.95, 4.6, C.text2);
  circleIcon(s, "FaStar", 8.15, 2.25, 0.75, C.accent6);
  s.addText("Zvyky po večeři", { x: 9.1, y: 2.35, w: 3.5, h: 0.55, fontSize: 20, bold: true, color: C.background1, isTextBox: true, margin: 0 });
  s.addText([
    { text: "Rozkrojit jablko — hvězdička nese zdraví", options: { bullet: true, breakLine: true } },
    { text: "Pouštět lodičky z ořechových skořápek", options: { bullet: true, breakLine: true } },
    { text: "Házet pantoflem přes rameno", options: { bullet: true, breakLine: true } },
    { text: "Lít olovo (nebo vosk) do vody", options: { bullet: true, breakLine: true } },
    { text: "Šupina z kapra do peněženky", options: { bullet: true } },
  ], { x: 8.15, y: 3.25, w: 4.4, h: 3.1, fontSize: 15, color: C.background1, paraSpaceAfter: 8, valign: "top", isTextBox: true, margin: 0 });

  // ---------- 5. 25. + 26.
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Program" });
  s.addText("Boží hod a Štěpán · 25.–26. 12.", { placeholder: "title" });
  s.addText("Sváteční jídlo, rodina a čerstvý vzduch", { placeholder: "body" });
  const two = [
    ["FaUtensils", C.accent1, "Pátek 25. 12. — Boží hod", [
      "Pozdní snídaně ve stylu „pyžamo do oběda“",
      "Sváteční oběd: pečená kachna, knedlík a zelí",
      "Odpoledne návštěva babičky a dědy",
      "Večer společná pohádka a zbytky cukroví",
    ]],
    ["FaBell", C.accent2, "Sobota 26. 12. — Štěpán", [
      "Dopolední výlet: rozhledna nebo zámecký park",
      "Prohlídka betlému v místním kostele",
      "Návštěva přátel, malá koleda s dětmi",
      "Večer deskové hry a horká čokoláda",
    ]],
  ];
  two.forEach(([n, col, h, items], i) => {
    const x = 0.6 + i * 6.15;
    card(s, x, 2.0, 5.95, 4.6, C.background2);
    circleIcon(s, n, x + 0.35, 2.3, 0.85, col);
    s.addText(h, { x: x + 1.45, y: 2.4, w: 4.3, h: 0.6, fontSize: 20, bold: true, color: C.text2, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < items.length - 1 } })),
      { x: x + 0.4, y: 3.5, w: 5.25, h: 2.9, fontSize: 16, color: C.text1, paraSpaceAfter: 12, valign: "top", isTextBox: true, margin: 0 });
  });

  // ---------- 6. Mezi svátky
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Program" });
  s.addText("Mezi svátky · 27.–30. 12.", { placeholder: "title" });
  s.addText("Jeden aktivní a jeden odpočinkový bod denně", { placeholder: "body" });
  const mid = [
    ["FaSkating", "Ne 27. 12.", "Bruslení na zimním stadionu nebo kluzišti, odpoledne odpočinek"],
    ["FaSleigh", "Po 28. 12.", "Výlet na hory — sáňky, nebo zimní túra, když nebude sníh"],
    ["FaFilm", "Út 29. 12.", "Kino nebo divadlo, večer domácí pizza"],
    ["FaHome", "St 30. 12.", "Úklid po svátcích, nákup na Silvestra, odpoledne wellness"],
    ["FaDice", "Každý večer", "Turnaj v deskových hrách — vítěz si vybere film"],
    ["FaMugHot", "Kdykoli", "Čtení u čaje, procházky, volno bez plánu"],
  ];
  mid.forEach(([n, d, t], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.1, y = 2.0 + row * 2.4;
    card(s, x, y, 3.9, 2.15, C.background2);
    circleIcon(s, n, x + 0.3, y + 0.3, 0.75, row === 0 ? C.accent1 : C.accent4);
    s.addText(d, { x: x + 1.25, y: y + 0.35, w: 2.5, h: 0.6, fontSize: 18, bold: true, color: C.accent2, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(t, { x: x + 0.3, y: y + 1.15, w: 3.35, h: 0.9, fontSize: 14, color: C.text1, valign: "top", isTextBox: true, margin: 0 });
  });

  // ---------- 7. Silvestr + Nový rok
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Program" });
  s.addText("Silvestr a Nový rok · 31. 12. – 1. 1.", { placeholder: "title" });
  s.addText("Rozloučení s rokem 2026 a pohodový start do roku 2027", { placeholder: "body" });
  card(s, 0.6, 2.0, 7.3, 4.6, C.text2);
  circleIcon(s, "FaGlassCheers", 0.95, 2.3, 0.8, C.accent2);
  s.addText("Čtvrtek 31. 12.", { x: 1.95, y: 2.4, w: 5.7, h: 0.6, fontSize: 22, bold: true, color: C.background1, valign: "middle", isTextBox: true, margin: 0 });
  const silv = [
    ["15:00", "Příprava chlebíčků a jednohubek"],
    ["19:00", "Příchod hostů, raclette nebo fondue"],
    ["21:00", "Silvestrovský kvíz o roce 2026"],
    ["23:55", "Odpočet, přípitek a ohňostroj"],
  ];
  silv.forEach(([t, d], i) => {
    const y = 3.4 + i * 0.75;
    s.addText(t, { x: 0.95, y, w: 1.1, h: 0.5, fontSize: 16, bold: true, color: C.accent5, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(d, { x: 2.15, y, w: 5.5, h: 0.5, fontSize: 16, color: C.background1, valign: "middle", isTextBox: true, margin: 0 });
  });
  card(s, 8.15, 2.0, 4.6, 4.6, C.background2);
  circleIcon(s, "FaSun", 8.5, 2.3, 0.8, C.accent4);
  s.addText("Pátek 1. 1. 2027", { x: 9.5, y: 2.4, w: 3.1, h: 0.6, fontSize: 20, bold: true, color: C.text2, valign: "middle", isTextBox: true, margin: 0 });
  s.addText([
    { text: "Dlouhé spaní a pozdní brunch", options: { bullet: true, breakLine: true } },
    { text: "Novoroční procházka nebo výstup na kopec", options: { bullet: true, breakLine: true } },
    { text: "Čočková polévka — ať se drží peníze", options: { bullet: true, breakLine: true } },
    { text: "Sepsat předsevzetí do rodinné obálky", options: { bullet: true } },
  ], { x: 8.5, y: 3.4, w: 4.0, h: 3.0, fontSize: 15, color: C.text1, paraSpaceAfter: 10, valign: "top", isTextBox: true, margin: 0 });

  // ---------- 8. Checklist
  pres.addSection({ title: "Závěr" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Závěr" });
  s.addText("Kontrolní seznam", { placeholder: "title" });
  s.addText("Co mít zajištěné, než svátky začnou", { placeholder: "body" });
  const chk = [
    ["Nakoupené a zabalené dárky", "do 22. 12."],
    ["Napečené cukroví", "do 20. 12."],
    ["Stromeček a ozdoby", "23. 12."],
    ["Kapr, brambory, majonéza", "23. 12."],
    ["Kachna a zelí na Boží hod", "do 23. 12."],
    ["Lístky do kina / divadla", "do 20. 12."],
    ["Nákup na Silvestra, šampaňské", "30. 12."],
    ["Čočka na Nový rok", "30. 12."],
  ];
  chk.forEach(([t, d], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.15, y = 2.0 + row * 1.15;
    card(s, x, y, 5.95, 0.9, C.background2);
    circleIcon(s, "FaCheck", x + 0.2, y + 0.15, 0.6, C.accent4);
    s.addText(t, { x: x + 1.0, y, w: 3.4, h: 0.9, fontSize: 16, color: C.text1, valign: "middle", isTextBox: true, margin: 0 });
    s.addText(d, { x: x + 4.3, y, w: 1.45, h: 0.9, fontSize: 15, bold: true, color: C.accent2, align: "right", valign: "middle", isTextBox: true, margin: 0 });
  });

  // ---------- 9. Closing
  s = pres.addSlide({ masterName: "Titul", sectionTitle: "Závěr" });
  flakes(s);
  s.addImage({ data: ic.FaStar, x: 6.22, y: 1.3, w: 0.9, h: 0.9, altText: "Hvězda", objectName: "star" });
  s.addText("Veselé Vánoce!", { placeholder: "title" });
  s.addText("a šťastný nový rok 2027", { placeholder: "body" });

  await pres.writeFile({ fileName: "Vanocni_program_2026.pptx" });
  await applyTheme("Vanocni_program_2026.pptx", THEME);
})();
