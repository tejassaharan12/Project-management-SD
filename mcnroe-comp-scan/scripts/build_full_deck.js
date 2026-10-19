const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const IMG = "/tmp/claude-0/-home-user-Project-management-SD/0fb55814-563c-5ce0-9100-d17133f99727/scratchpad/img";
const SIZES = JSON.parse(fs.readFileSync(path.join(IMG, "sizes.json")));
const LINKS = JSON.parse(fs.readFileSync(path.join(IMG, "links.json")));
const L = (k) => (LINKS.ads[k] || LINKS.ig[k]).link;
const YT = (id) => "https://www.youtube.com/watch?v=" + id;
const DOC = "https://claude.ai/code/artifact/b7307ea6-30ea-466f-9f1c-9db21c95589a";
const REPO = "https://github.com/tejassaharan12/Project-management-SD/tree/claude/nifty-tesla-i9gpgr/mcnroe-comp-scan/data";

const COL = { dark: "2A1A22", bg: "FAFAFA", ink: "1F1A1C", muted: "6B6466", plum: "7A2E4F", gold: "B8862F", card: "F1ECEE", line: "E0DADC", white: "FFFFFF" };
const HEAD = "Cambria", BODY = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Competitor Scan: Full Report, Jul to Oct 2026";
pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };

function slideBase(dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? COL.dark : COL.bg };
  return s;
}
function title(s, text) {
  s.addText(text, { x: 0.5, y: 0.3, w: 9, h: 0.95, fontFace: HEAD, fontSize: 24, bold: true, color: COL.ink, valign: "top", margin: 0, isTextBox: true });
}
function source(s, text) {
  s.addText(text, { x: 0.5, y: 5.2, w: 9, h: 0.3, fontFace: BODY, fontSize: 9, color: COL.muted, margin: 0, isTextBox: true });
}
// place image inside box, keeping aspect ratio, centered
function img(s, key, x, y, w, h, link) {
  const [iw, ih] = SIZES[key];
  const r = Math.min(w / iw, h / ih);
  const pw = iw * r, ph = ih * r;
  const px = x + (w - pw) / 2, py = y + (h - ph) / 2;
  const o = { path: path.join(IMG, key + ".jpg"), x: px, y: py, w: pw, h: ph, rounding: false };
  if (link) o.hyperlink = { url: link, tooltip: "Open source" };
  s.addImage(o);
  return { x: px, y: py, w: pw, h: ph };
}
function caption(s, x, y, w, head, sub, link, linkText) {
  const runs = [{ text: head, options: { bold: true, color: COL.ink, fontSize: 12, breakLine: true } }];
  if (sub) runs.push({ text: sub, options: { color: COL.muted, fontSize: 11, breakLine: !!link } });
  if (link) runs.push({ text: linkText || "View source", options: { color: COL.plum, fontSize: 11, underline: true, hyperlink: { url: link } } });
  s.addText(runs, { x, y, w, h: 0.75, fontFace: BODY, valign: "top", margin: 0, isTextBox: true });
}
function barChart(s, labels, values, o) {
  s.addChart(pres.charts.BAR, [{ name: o.name || "Value", labels, values }], {
    x: o.x, y: o.y, w: o.w, h: o.h, barDir: "bar", chartColors: [o.color || COL.plum],
    catAxisLabelColor: COL.ink, catAxisLabelFontSize: 11, catAxisLabelFontFace: BODY, catAxisOrientation: "maxMin",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: COL.ink, dataLabelFontSize: 11, dataLabelFontFace: BODY,
    dataLabelFormatCode: o.fmt || "#,##0", showLegend: false, barGapWidthPct: 45,
  });
}

// ---------- extra helpers ----------
const B = (b, t, last) => [{ text: b, options: { bold: true, bullet: true } }, { text: t, options: { breakLine: !last } }];
function bullets(s, items, x, y, w, h, size) {
  const runs = [];
  items.forEach(([b, t], i) => runs.push(...B(b, t, i === items.length - 1)));
  s.addText(runs, { x, y, w, h, fontFace: BODY, fontSize: size || 13, color: COL.ink, valign: "top", paraSpaceAfter: 7, margin: 0, isTextBox: true });
}
function divider(num, name, sub) {
  const s = slideBase(true);
  s.addText(num, { x: 0.6, y: 1.5, w: 2, h: 0.8, fontFace: HEAD, fontSize: 40, bold: true, color: "E9D8B4", margin: 0, isTextBox: true });
  s.addText(name, { x: 0.6, y: 2.3, w: 8.8, h: 0.8, fontFace: HEAD, fontSize: 34, bold: true, color: COL.white, margin: 0, isTextBox: true });
  if (sub) s.addText(sub, { x: 0.6, y: 3.15, w: 8.8, h: 0.8, fontFace: BODY, fontSize: 15, color: "CFC4C8", margin: 0, isTextBox: true, valign: "top" });
}
function table(s, header, rows, o) {
  const hdr = header.map((t, i) => ({ text: t, options: { bold: true, align: (o.align || [])[i] || "left" } }));
  const body = rows.map((r) => r.map((t, i) => ({ text: t, options: { bold: i === 0, align: (o.align || [])[i] || "left", valign: "top" } })));
  s.addTable([hdr, ...body], { x: o.x || 0.5, y: o.y || 1.3, w: o.w || 9, colW: o.colW, fontFace: BODY, fontSize: o.size || 11, color: COL.ink, rowH: [0.32].concat(rows.map(() => o.rowH || 0.3)), valign: "top", border: { type: "solid", pt: 0.75, color: COL.line }, margin: 0.05 });
}
function kpis(s, items, y) {
  const n = items.length, gap = 0.15, w = (9 - gap * (n - 1)) / n;
  items.forEach(([v, l], i) => {
    const x = 0.5 + i * (w + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.78, fill: { color: COL.card }, line: { color: COL.card }, rectRadius: 0.08 });
    s.addText([{ text: v, options: { fontSize: 17, bold: true, color: COL.plum, fontFace: HEAD, breakLine: true } }, { text: l, options: { fontSize: 9.5, color: COL.muted } }], { x: x + 0.1, y: y + 0.05, w: w - 0.2, h: 0.68, fontFace: BODY, valign: "middle", margin: 0, isTextBox: true });
  });
}
function brandSlide(name, ttl, k, items, pics) {
  const s = slideBase(false);
  title(s, ttl);
  kpis(s, k, 1.28);
  const hasPics = pics && pics.length;
  bullets(s, items, 0.5, 2.25, hasPics ? 5.5 : 9, 2.9, 12);
  if (hasPics) {
    const n = pics.length, w = n === 1 ? 3.2 : 1.6;
    pics.forEach(([key, link, cap, lt], i) => {
      const x = n === 1 ? 6.3 : 6.25 + i * 1.7;
      img(s, key, x, 2.25, w, 2.2, link);
      s.addText([{ text: cap, options: { color: COL.muted, breakLine: true } }, { text: lt || "View", options: { color: COL.plum, underline: true, hyperlink: { url: link } } }], { x, y: 4.5, w, h: 0.6, fontFace: BODY, fontSize: 9.5, valign: "top", margin: 0, isTextBox: true });
    });
  }
  source(s, "Meta = ads active since 15 Jul. IG = Instagram, 15 Jul–15 Oct. YT = YouTube uploads since 1 Jun. Search = Google Trends index (Bellavita = 100).");
  return s;
}

// ================= SLIDES =================
// 1 Title
{
  const s = slideBase(true);
  s.addText("Competitor Scan", { x: 0.6, y: 1.2, w: 8.8, h: 0.9, fontFace: HEAD, fontSize: 44, bold: true, color: COL.white, margin: 0, isTextBox: true });
  s.addText("Full report: fragrance and grooming in India, 15 July to 15 October 2026", { x: 0.6, y: 2.1, w: 8.8, h: 0.5, fontFace: BODY, fontSize: 19, color: "E9D8B4", margin: 0, isTextBox: true });
  s.addText("Denver · BlaBliBlu · Bellavita · indē wild · Secret Alchemist · Beardo · Park Avenue · Engage · Ajmal · EM5 · Fraganote · Plum", { x: 0.6, y: 3.0, w: 8.8, h: 0.6, fontFace: BODY, fontSize: 13, color: "CFC4C8", margin: 0, isTextBox: true });
  s.addText("Press · Meta ads · Google ads · Instagram · YouTube · Google Trends", { x: 0.6, y: 3.6, w: 8.8, h: 0.4, fontFace: BODY, fontSize: 13, color: "CFC4C8", margin: 0, isTextBox: true });
  s.addText("Prepared for Mcnroe · October 2026", { x: 0.6, y: 4.7, w: 8.8, h: 0.4, fontFace: BODY, fontSize: 13, color: "CFC4C8", margin: 0, isTextBox: true });
}
// 2 Agenda
{
  const s = slideBase(false);
  title(s, "Agenda");
  const sec = [["1", "Category themes", "What the whole category did: Diwali, bundles, oil %, scents, creators, funding"], ["2", "Platform by platform", "Meta ads, Google ads, Instagram, YouTube, Google Trends"], ["3", "Brand by brand", "One page per competitor, all 12"], ["4", "What it means for Mcnroe", "Launch calendar, implications, ongoing tracking"]];
  sec.forEach(([n, h, t], i) => {
    const y = 1.35 + i * 0.92;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y, w: 0.6, h: 0.6, fill: { color: COL.plum }, line: { color: COL.plum } });
    s.addText(n, { x: 0.5, y, w: 0.6, h: 0.6, fontFace: HEAD, fontSize: 20, bold: true, color: COL.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText([{ text: h, options: { bold: true, fontSize: 18, fontFace: HEAD, breakLine: true } }, { text: t, options: { fontSize: 13, color: COL.muted } }], { x: 1.35, y: y - 0.08, w: 8.1, h: 0.78, fontFace: BODY, color: COL.ink, valign: "middle", margin: 0, isTextBox: true });
  });
}
// 3 Exec summary
{
  const s = slideBase(false);
  title(s, "Executive summary");
  bullets(s, [
    ["Diwali selling started on 1 October. ", "BlaBliBlu, Denver, Fraganote and Bellavita already run festive offers, mostly free gifts and buy-1-get-1."],
    ["“Any 3” bundles set the price frame: ", "3 perfumes for ₹650 (BlaBliBlu) to ₹1,799 (Ajmal)."],
    ["Perfume-oil % is the quality claim. ", "Bellavita now claims 40%, against 25% (BlaBliBlu, Secret Alchemist) and 20% (Denver Intense)."],
    ["Gourmand and oud dominate new launches; ", "fresh, floral and women's have few strong voices."],
    ["Beardo is the heaviest Meta advertiser (960 ads). ", "Engage ran the biggest paid YouTube push (24M views) with no press coverage."],
    ["Contests and stunts win on Instagram, ", "earning 10 to 40 times normal engagement."],
  ], 0.5, 1.25, 6.0, 3.9, 13.5);
  [["3,387", "Meta ads"], ["663", "Instagram posts"], ["390", "YouTube videos"]].forEach(([n, l], i) => {
    const y = 1.25 + i * 1.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.8, y, w: 2.7, h: 1.1, fill: { color: COL.card }, line: { color: COL.card }, rectRadius: 0.1 });
    s.addText([{ text: n, options: { fontSize: 30, bold: true, color: COL.plum, fontFace: HEAD, breakLine: true } }, { text: l + " analysed", options: { fontSize: 12, color: COL.muted } }], { x: 6.95, y: y + 0.08, w: 2.4, h: 0.95, fontFace: BODY, valign: "middle", margin: 0, isTextBox: true });
  });
  source(s, "Sources: Meta Ad Library, Google Ads Transparency Center, Instagram Graph API, YouTube Data API, Google Trends, trade press. Pulled 15 October 2026.");
}
// 4 Scope
{
  const s = slideBase(false);
  title(s, "We tracked 12 brands in four groups across six data layers");
  table(s, ["Group", "Brands", "Why they matter to Mcnroe"], [
    ["D2C perfume challengers", "Bellavita, Fraganote, Secret Alchemist, BlaBliBlu", "Digital-first, creator-led, price-value fragrance"],
    ["Mass deo and fragrance", "Denver, Engage, Park Avenue", "Big ATL budgets, set mass category codes"],
    ["Heritage and premium", "Ajmal, EM5", "Oud, attar, gifting and premium cues"],
    ["Grooming and beauty", "Beardo, indē wild, Plum", "Stretching into fragrance and mists"],
  ], { y: 1.3, colW: [2.3, 3.4, 3.3], size: 12, rowH: 0.42 });
  s.addText("Layers: press and launches · Meta ads · Google ads · Instagram · YouTube · Google Trends. Window: 15 July to 15 October 2026. Influencer and marketplace tracking were out of scope.", { x: 0.5, y: 3.75, w: 9, h: 0.7, fontFace: BODY, fontSize: 13, color: COL.ink, margin: 0, isTextBox: true });
}

// ===== Section 1 =====
divider("1", "Category themes", "What the category as a whole did this quarter");
// 6 Themes overview
{
  const s = slideBase(false);
  title(s, "Seven moves defined the category this quarter");
  table(s, ["Theme", "What we saw", "Brands"], [
    ["Everyone moves to EDP; oil % is the claim", "“Long-lasting” is now a number: 20%, 25%, 40%", "Denver, BlaBliBlu, Bellavita, Beardo, Park Avenue"],
    ["Gen Z vibe language", "Rizz God, Vibe Collection, #OwnTheVibe, Love Drunk, Old Money, Berry Baddie", "Denver, Bellavita, BlaBliBlu"],
    ["Celebrities at the top, celebrity-investors below", "SRK and Ranveer Singh are paid faces; Samantha and Suryakumar Yadav hold equity", "Denver, Ajmal, Secret Alchemist, EM5"],
    ["Women's and unisex expansion", "Park Avenue “for Her”, unisex Sports Mode and Love Drunk", "Park Avenue, Fraganote, BlaBliBlu"],
    ["Marketplace events as launch pads", "Launches timed to Flipkart Glam Up Fest and Amazon Beautyverse", "Beardo, Engage"],
    ["Funding buys offline reach", "Kiosks and stores: 100 kiosks, 170+ stores, 111 stores planned", "Fraganote, Ajmal, EM5, BlaBliBlu"],
    ["Heritage and ritual stories", "Champi oil, sandalwood nostalgia, oud, Kashmir", "indē wild, Secret Alchemist, Ajmal, EM5"],
  ], { y: 1.25, colW: [2.6, 3.9, 2.5], size: 10.5, rowH: 0.47 });
}
// 7 Diwali
{
  const s = slideBase(false);
  title(s, "Diwali selling began on 1 October, more than a month before the festival");
  [["denver_diwali", "Denver", "Buy 1 get 1 + daily MacBook draw, from 13 Oct"], ["bbb_diwali", "BlaBliBlu", "Diwali House Party sale, free gifts, from 1 Oct"], ["fraganote_diwali", "Fraganote", "Festive 50 ml sets, Diwali Edit, from 9 Oct"], ["bellavita_gift", "Bellavita", "Any 3 perfumes for ₹1,298, gifting-ready"]].forEach(([k, b, t], i) => {
    const x = 0.5 + i * 2.3;
    img(s, k, x, 1.3, 2.05, 2.75, L(k));
    caption(s, x, 4.12, 2.05, b, t, L(k), "View ad");
  });
  source(s, "Diwali is on 8 November 2026. Rakhi (August) was the first gifting peak: six brands ran Rakhi gifting. Ads from Meta Ad Library, India.");
}
// 8 Bundles
{
  const s = slideBase(false);
  title(s, "“Any 3” bundles set the price frame: 3 perfumes for ₹650 to ₹1,799");
  barChart(s, ["BlaBliBlu", "Denver", "Bellavita", "Ajmal"], [650, 1000, 1298, 1799], { x: 0.4, y: 1.3, w: 4.9, h: 3.0, fmt: "₹#,##0", name: "Price of 3 perfumes" });
  s.addText("Price of a 3-perfume bundle in Meta ads (₹)", { x: 0.5, y: 4.3, w: 4.8, h: 0.3, fontFace: BODY, fontSize: 11, color: COL.muted, margin: 0, isTextBox: true });
  s.addText([{ text: "Build-your-own box appears in 6 of 10 advertisers. ", options: { bold: true } }, { text: "Beardo sells 4 perfumes for ₹999; Secret Alchemist 2 × 15 ml for ₹1,099." }], { x: 0.5, y: 4.6, w: 4.8, h: 0.55, fontFace: BODY, fontSize: 12, color: COL.ink, margin: 0, isTextBox: true });
  img(s, "denver_any3", 5.6, 1.3, 1.9, 1.9, L("denver_any3"));
  img(s, "bbb_any3", 7.6, 1.3, 1.9, 1.9, L("bbb_any3"));
  caption(s, 5.6, 3.3, 1.9, "Denver", "Any 3 for ₹1,000", L("denver_any3"), "View ad");
  caption(s, 7.6, 3.3, 1.9, "BlaBliBlu", "3 for ₹750", L("bbb_any3"), "View ad");
  source(s, "Offers seen in Meta ads, August to October 2026. BlaBliBlu runs ₹650 and ₹750 versions.");
}
// 9 Oil
{
  const s = slideBase(false);
  title(s, "Perfume-oil percentage is the quality claim, and Bellavita now claims 40%");
  barChart(s, ["Bellavita Beast Mode", "BlaBliBlu", "Secret Alchemist", "Denver Rizz God", "Denver Intense"], [40, 25, 25, 25, 20], { x: 0.4, y: 1.3, w: 5.6, h: 3.3, fmt: "0\"%\"", name: "Oil %" });
  s.addText("Perfume-oil concentration claimed in ads and launches", { x: 0.5, y: 4.6, w: 5.5, h: 0.3, fontFace: BODY, fontSize: 11, color: COL.muted, margin: 0, isTextBox: true });
  img(s, "bellavita_beast", 6.6, 1.3, 2.9, 3.0, L("bellavita_beast"));
  caption(s, 6.6, 4.35, 2.9, "Bellavita Beast Mode", "Collector's flask with Oud Supreme, 40% oil", L("bellavita_beast"), "View ad");
  source(s, "Brand claims, not lab-tested. Sources: Meta Ad Library; Denver launch press releases (June and September 2026).");
}
// 10 Gourmand
{
  const s = slideBase(false);
  title(s, "Gourmand and oud are this season's scent stories; fresh and floral are less crowded");
  [["ig_bbb_gourmand", "BlaBliBlu", "4 fruity gourmands, one a day, 21–24 Sep", L("ig_bbb_gourmand"), "View post"], ["fraganote_minis", "Fraganote", "Devil's Berry and dessert scents, 25 ml minis", L("fraganote_minis"), "View ad"], ["ajmal_oud", "Ajmal", "Oud trio: Oud Nirvana, Oud on the Rocks, White Oud", L("ajmal_oud"), "View ad"], ["yt_fraganote", "Fraganote", "Oud Collection launch, 17 Jul (570K views)", YT("_kFW1VDYPAA"), "Watch"]].forEach(([k, b, t, lk, lt], i) => {
    const x = 0.5 + i * 2.3;
    img(s, k, x, 1.3, 2.05, 2.6, lk);
    caption(s, x, 3.97, 2.05, b, t, lk, lt);
  });
  s.addText([{ text: "Also: ", options: { bold: true } }, { text: "Bellavita Gourmand Trend Series at ₹549 (26 Sep). Rising search: “bella vita perfume pistachio” +400%. Secret Alchemist's jasmine and tuberose are the main floral voice." }], { x: 0.5, y: 4.85, w: 9, h: 0.35, fontFace: BODY, fontSize: 11, color: COL.ink, margin: 0, isTextBox: true });
}
// 11 Creators
{
  const s = slideBase(false);
  title(s, "Creators front a large share of paid ads, and quick commerce is a common destination");
  [["sa_samantha", "Secret Alchemist", "Half its top ads run from Samantha's handle"], ["bv_creator", "Bellavita", "Regional creators post “comment for link” reactions"], ["beardo_creator", "Beardo", "Creator reviews of Godfather Mafia and Revenge"], ["indewild_creator", "indē wild", "Most launch ads are creator videos"]].forEach(([k, b, t], i) => {
    const x = 0.5 + i * 2.3;
    img(s, k, x, 1.3, 2.05, 2.6, L(k));
    caption(s, x, 3.97, 2.05, b, t, L(k), "View ad");
  });
  s.addText([{ text: "Quick commerce: ", options: { bold: true } }, { text: "Engage, Beardo (“in 10 minutes”), Secret Alchemist and Plum send ads straight to Blinkit, Zepto or Instamart." }], { x: 0.5, y: 4.85, w: 9, h: 0.35, fontFace: BODY, fontSize: 12, color: COL.ink, margin: 0, isTextBox: true });
}
// 12 Women / unisex / mists
{
  const s = slideBase(false);
  title(s, "Male-coded brands are crossing into women's and unisex, and indē wild is entering mists");
  bullets(s, [
    ["Park Avenue “for Her” (10 Jul): ", "Godrej's first women's EDPs, four variants (Poise, Grace, Verve, Allure), pitched for gifting. Barely visible in ads since."],
    ["Fraganote Sports Mode (7 Aug): ", "unisex 50 ml EDP for active lifestyles."],
    ["BlaBliBlu: ", "unisex hero scents (Love Drunk) and fruity gourmands aimed at young women."],
    ["indē wild hair and body mists (Aug): ", "After Heat and Solar Skin, “odour neutralising, mood-boosting”. A well-funded beauty brand now in the mist segment."],
    ["Women-first fragrance voices are few: ", "Secret Alchemist (Samantha, florals) and indē wild."],
  ], 0.5, 1.3, 5.6, 3.8, 13);
  img(s, "ig_indewild_mists", 6.4, 1.3, 3.1, 3.1, L("ig_indewild_mists"));
  caption(s, 6.4, 4.45, 3.1, "indē wild Solar Skin mist", "Launched August 2026", L("ig_indewild_mists"), "View post");
}
// 13 Funding
{
  const s = slideBase(false);
  title(s, "Funding is going into stores, kiosks and new categories");
  const cards = [["$3M", "Secret Alchemist", "Unilever Ventures-led round (Jan); claims 8x growth in FY26; hair perfumes next"], ["$3M", "Fraganote", "Series A, V3 Ventures (Jun); kiosks in 14 cities, 100 by year-end"], ["₹100 Cr", "BlaBliBlu", "Annual run-rate within 6 months (Mar); Delhi NCR stores, pan-India next"], ["170+", "Ajmal", "Exclusive stores targeted by end-2026, tier 2 and 3 focus; first London store (Aug)"], ["50 → 111", "EM5", "Stores planned; Suryakumar Yadav joined as investor (Feb)"]];
  cards.forEach(([n, b, t], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 1.3 + row * 1.9, w = 2.85;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.7, fill: { color: COL.card }, line: { color: COL.card }, rectRadius: 0.08 });
    s.addText([{ text: n, options: { fontSize: 24, bold: true, color: COL.plum, fontFace: HEAD, breakLine: true } }, { text: b, options: { fontSize: 13, bold: true, color: COL.ink, breakLine: true } }, { text: t, options: { fontSize: 11, color: COL.muted } }], { x: x + 0.15, y: y + 0.1, w: w - 0.3, h: 1.5, fontFace: BODY, valign: "top", margin: 0, isTextBox: true });
  });
  source(s, "Sources: The Week, Inc42, Retail4Growth, India Retailing, Business Standard (2026). Growth figures are company claims.");
}
// 14 Campaign log
{
  const s = slideBase(false);
  title(s, "Press-visible campaigns and launches, June to October 2026");
  table(s, ["Date", "Brand", "Campaign or move", "Talent / channel"], [
    ["Oct", "indē wild", "Champi Clarifying & Detox Shampoo launch (Nykaa)", "Rida Tharana"],
    ["18 Sep", "Denver", "Rizz God EDP, Gen Z “Vibe Collection”, ₹2,500, 25% oil, D2C only", "Press release"],
    ["Aug", "indē wild", "Champi hair-oil ritual campaign", "Anupama Parameswaran"],
    ["7 Aug", "Fraganote", "Sports Mode unisex EDP, 50 ml", "Launch"],
    ["21 Jul", "Secret Alchemist", "Sandalwood Nostalgia; clean, “step up from mass”", "Launch"],
    ["10 Jul", "Park Avenue", "Enters women's: “for Her” EDP, 4 variants", "Launch"],
    ["Jun", "Fraganote", "$3M Series A; 100 kiosks planned", "Funding"],
    ["19–20 Jun", "Beardo", "Godfather Mafia and Revenge (₹1,700) at Flipkart and Amazon events", "Danish Pandor"],
    ["3 Jun", "Denver", "Intense EDP, “Success Rules The World” TVC, ₹999, 20% oil", "Shah Rukh Khan"],
  ], { y: 1.25, colW: [1.0, 1.5, 4.6, 1.9], size: 10.5, rowH: 0.36 });
  source(s, "Earlier in 2026: Ajmal “Your Unseen Power” with Ranveer Singh on 575+ PVR screens (19 Mar); Bellavita #OwnTheVibe with Raghav Juyal (Jan); Plum BodyLovin' relaunch (Apr).");
}

// ===== Section 2 =====
divider("2", "Platform by platform", "Meta ads · Google ads · Instagram · YouTube · Google Trends");
// Meta volume
{
  const s = slideBase(false);
  title(s, "Beardo ran 960 Meta ads since mid-July, more than twice the next brand");
  barChart(s, ["Beardo", "BlaBliBlu", "Bellavita", "Secret Alchemist", "Ajmal", "Fraganote", "Denver", "Plum", "Engage", "indē wild"], [960, 429, 422, 373, 339, 231, 219, 215, 127, 72], { x: 0.4, y: 1.25, w: 6.0, h: 3.85, name: "Meta ads" });
  bullets(s, [["3,387 ads ", "ran across 10 brands; we read the 60 most-seen per brand."], ["Engage ", "runs only catalogue ads for roll-ons."], ["Park Avenue and EM5 ", "have no findable Meta advertiser page."]], 6.7, 1.4, 2.8, 3.5, 13);
  source(s, "Ads active in India, 15 July to 15 October 2026. Source: Meta Ad Library via Apify.");
}
// Meta table A
{
  const s = slideBase(false);
  title(s, "Meta ads by brand (1 of 2): the biggest advertisers lead with bundles and creators");
  table(s, ["Brand", "Ads", "What the top ads push", "Offer and format"], [
    ["Beardo", "960", "Perfume combos (Mariner, Whisky Smoke Bourbon ₹499), Godfather reviews; also de-tan, curls, beard, trimmers", "4 perfumes ₹999; 1 in 5 top ads from creators"],
    ["BlaBliBlu", "429", "Build-your-own 30 ml box, “25%, 12+ hours”; Diwali House Party sale with free gifts; payday hook", "3 for ₹650–750; 38 of top 60 started in Oct"],
    ["Bellavita", "422", "Beast Mode for Men (40% oil), half of top ads; Gourmand Trend Series ₹549; Vanilluxe", "Any 2 ₹899, 3 ₹1,298, 4 ₹1,599; regional creators"],
    ["Secret Alchemist", "373", "Clean Perfume (IFRA, 25%); Rare Vanilla “one sold every minute”; Rose Oud; hair perfumes; Instamart", "Free 10 ml with 100 ml; half of top ads from Samantha"],
    ["Ajmal", "339", "Oud trio, Cyan, Aristocrat with Ranveer Singh, gift sets", "Any 3 ₹1,799; Gold Series ₹899; brand page only"],
  ], { y: 1.25, colW: [1.5, 0.6, 4.2, 2.7], size: 10.5, rowH: 0.62 });
}
// Meta table B
{
  const s = slideBase(false);
  title(s, "Meta ads by brand (2 of 2): Denver discounts hardest; Engage and Plum sell other categories");
  table(s, ["Brand", "Ads", "What the top ads push", "Offer and format"], [
    ["Fraganote", "231", "Gourmands (Baked Vanilla, Drunken Cake, Devil's Berry), Oud Collection, 25 ml minis, Rakhi box, Diwali Edit", "From ₹999; personalised gift boxes"],
    ["Denver", "219", "Any 3 perfumes ₹1,000; deo packs ₹599–1,200; Autograph 3 for ₹2,500; Diwali BOGO + MacBook; SRK contest", "Deepest discounts; 45 of 60 static images"],
    ["Plum", "215", "Skincare only: CeraSense barrier care, niacinamide serum, dandruff shampoo on Blinkit", "3 for ₹599"],
    ["Engage", "127", "Roll-ons only (Amber Fresh, Radiant Glow “skin brightening”, 72-hour), to Blinkit and Zepto", "100% catalogue ads"],
    ["indē wild", "72", "Champi Clarifying Shampoo (“96 hours oil control”), Dewy Lip Treatment, Sleek Stick", "15–20% off duos; mostly creators"],
  ], { y: 1.25, colW: [1.5, 0.6, 4.2, 2.7], size: 10.5, rowH: 0.62 });
}
// Google
{
  const s = slideBase(false);
  title(s, "Denver, BlaBliBlu and Fraganote ramped up Google ads in September; Bellavita's are a year old");
  const hdr = ["Brand", "Creatives (cap 60)", "Median days live", "New in September"];
  table(s, hdr, [["Denver", "60", "12", "40"], ["BlaBliBlu", "60", "22", "45"], ["Fraganote", "60", "20", "28"], ["Plum", "60", "35", "22"], ["Ajmal", "60", "46", "14"], ["Secret Alchemist", "60", "101", "13"], ["Beardo", "60", "84", "2"], ["indē wild", "27", "16", "0"], ["EM5", "15", "227", "1"], ["Bellavita", "60", "348", "1"]], { w: 6.2, colW: [2.0, 1.4, 1.4, 1.4], size: 12, rowH: 0.31, align: ["left", "right", "right", "right"] });
  bullets(s, [["Short-lived creatives = testing. ", "Denver's median Google ad ran 12 days."], ["Bellavita is Meta-only for new work. ", "Beast Mode and Gourmand never reached Google."], ["Beardo defends search: ", "48 of 60 are text ads."]], 7.0, 1.35, 2.5, 3.6, 12.5);
  source(s, "Google Ads Transparency Center, India, 15 Jul–15 Oct 2026. Engage and Park Avenue advertise under ITC and Godrej and could not be separated.");
}
// IG table A
{
  const s = slideBase(false);
  title(s, "Instagram by brand (1 of 2): indē wild and Bellavita post the most");
  table(s, ["Brand", "Followers", "Posts (Jul/Aug/Sep/Oct)", "Reels", "Median eng.", "What drove the top posts"], [
    ["indē wild", "267K", "135 (17/54/37/27)", "84%", "0.16%", "Dino Morea “interim CEO” (49.7K likes); Sephora UK; summer mists; Bombay Sweet Shop collab"],
    ["Bellavita", "943K", "90 (15/24/29/22)", "60%", "0.04%", "Rakhi post (9.7K likes); product carousels; CEO Man bigger size"],
    ["Plum", "1.15M", "76 (15/26/25/10)", "67%", "0.03%", "BYOB barrier-care launch party; SPF giveaway"],
    ["Beardo", "808K", "70 (15/22/22/11)", "59%", "0.04%", "Rap battle series to brand anthem “Unstoppable” (10.1K likes)"],
    ["BlaBliBlu", "150K", "69 (12/22/21/14)", "68%", "0.12%", "4 fruity gourmands; Diwali “boring gifts are retiring”; Rakhi giveaway"],
    ["Secret Alchemist", "144K", "55 (7/23/18/7)", "42%", "0.11%", "Gifting Collection (5.2K likes); “four women, one scent”; brand giveaways"],
  ], { y: 1.25, colW: [1.4, 0.9, 1.6, 0.6, 0.9, 3.6], size: 10, rowH: 0.55 });
  source(s, "Instagram Graph API, every post 15 Jul–15 Oct 2026. Median engagement = (likes + comments) ÷ followers per post.");
}
// IG table B
{
  const s = slideBase(false);
  title(s, "Instagram by brand (2 of 2): Denver converts followers best; Park Avenue went quiet");
  table(s, ["Brand", "Followers", "Posts (Jul/Aug/Sep/Oct)", "Reels", "Median eng.", "What drove the top posts"], [
    ["Denver", "170K", "51 (6/14/19/12)", "43%", "0.30%", "SRK-birthday iPhone contest (17.4K likes, 21.9K comments); SRK “scent of success”"],
    ["Ajmal (India)", "273K", "47 (4/22/11/10)", "62%", "0.10%", "First UK store, London (1 Aug); Vogue Wedding Atelier; 75 years; perfumer series"],
    ["Fraganote", "90K", "37 (7/13/10/7)", "73%", "0.15%", "Hypnotic Oud; Sports Mode; Creme Castle collab; 25 ml minis"],
    ["Engage", "37K", "19 (5/6/5/3)", "74%", "0.11%", "Friendship Day post; perfume myths at Amazon Beautyverse; Maracujá range"],
    ["EM5", "117K", "8 (2/2/3/1)", "38%", "0.23%", "“Kashmir” scent; Inglorious Ginger and The Great Oud; coffee car perfume"],
    ["Park Avenue", "43K", "6 (1/5/0/0)", "33%", "0.06%", "Refreshed EDPs and new-look Euphoria (Aug); silent since 28 Aug"],
  ], { y: 1.25, colW: [1.4, 0.9, 1.6, 0.6, 0.9, 3.6], size: 10, rowH: 0.55 });
  source(s, "Instagram Graph API, every post 15 Jul–15 Oct 2026.");
}
// IG stunts
{
  const s = slideBase(false);
  title(s, "On Instagram, contests and stunts earned 10 to 40 times normal engagement");
  [["ig_denver_contest", "Denver: SRK-birthday iPhone contest", "17.4K likes, 21.9K comments (median post: 457 likes)"], ["ig_indewild_dino", "indē wild: Dino Morea as “interim CEO”", "49.7K likes (median post: 391)"], ["ig_beardo_anthem", "Beardo: brand anthem “Unstoppable”", "10.1K likes (median post: 284)"]].forEach(([k, b, t], i) => {
    const x = 0.5 + i * 3.08;
    img(s, k, x, 1.3, 2.85, 2.65, L(k));
    caption(s, x, 4.02, 2.85, b, t, L(k), "View post");
  });
  source(s, "Instagram Graph API. Denver's contest image shows its matching Meta ad creative.");
}
// IG engagement chart
{
  const s = slideBase(false);
  title(s, "Denver and EM5 get the most engagement per follower; the biggest accounts get the least");
  barChart(s, ["Denver", "EM5", "indē wild", "Fraganote", "BlaBliBlu", "Secret Alchemist", "Engage", "Ajmal", "Park Avenue", "Bellavita", "Beardo", "Plum"], [0.30, 0.23, 0.16, 0.15, 0.12, 0.11, 0.11, 0.10, 0.06, 0.04, 0.04, 0.03], { x: 0.4, y: 1.25, w: 5.8, h: 3.9, fmt: "0.00\"%\"", name: "Engagement" });
  bullets(s, [["Big accounts, low reaction: ", "Bellavita (943K), Plum (1.15M) and Beardo (808K)."], ["Brand collabs are cheap reach: ", "indē wild × Bombay Sweet Shop, Fraganote × Creme Castle, Secret Alchemist × Isharya, Winston, Instamart."]], 6.6, 1.35, 2.9, 3.6, 12.5);
  source(s, "Median (likes + comments) ÷ followers per post. Instagram Graph API, 663 posts.");
}
// YouTube chart
{
  const s = slideBase(false);
  title(s, "Engage's regional roll-on ads drew 24M YouTube views, more than all others combined");
  barChart(s, ["Engage", "Fraganote", "Beardo", "Ajmal", "BlaBliBlu", "indē wild", "Denver", "Plum", "EM5"], [24.0, 5.0, 3.6, 3.6, 1.6, 0.9, 0.8, 0.3, 0.2], { x: 0.4, y: 1.25, w: 5.2, h: 3.85, fmt: "0.0\"M\"", name: "Views" });
  img(s, "yt_engage", 5.9, 1.3, 3.6, 1.6, YT("TpHySVBnido"));
  caption(s, 5.9, 2.95, 3.6, "Engage roll-on, Bengali cut", "9.4M views, 3 likes: almost all paid", YT("TpHySVBnido"), "Watch");
  img(s, "yt_beardo", 5.9, 3.65, 1.75, 1.0, YT("AbEH-c62jwA"));
  s.addText([{ text: "Beardo Godfather", options: { bold: true, breakLine: true } }, { text: "2.9M views", options: { color: COL.muted, breakLine: true } }, { text: "Watch", options: { color: COL.plum, underline: true, hyperlink: { url: YT("AbEH-c62jwA") } } }], { x: 7.8, y: 3.65, w: 1.7, h: 1.0, fontFace: BODY, fontSize: 11, color: COL.ink, valign: "top", margin: 0, isTextBox: true });
  source(s, "YouTube Data API, uploads 1 Jun–15 Oct 2026, views to date.");
}
// YouTube table
{
  const s = slideBase(false);
  title(s, "Most big YouTube view counts are paid; indē wild is the only all-organic channel");
  table(s, ["Brand", "Uploads", "Views", "Paid share (est.)", "What they pushed"], [
    ["Engage", "39", "24.0M", "~100%", "32 uploads in Sep: roll-on ads in Hindi, Bengali, Kannada; Mini Perfume Spray TVC (Assamese)"],
    ["Fraganote", "92", "5.0M", "~73%", "Oud Collection, Hypnotic Oud; Rakhi sale (45% off + 3 free minis); Drunken Cake review (750K)"],
    ["Beardo", "43", "3.6M", "~15%", "“A Godfather perfume is sold every 15 seconds” (2.9M); brand anthem; curls content"],
    ["Ajmal", "67", "3.6M", "~42%", "Evoke Red gifting films; factory tour; “Authors of Ajmal” perfumer series"],
    ["BlaBliBlu", "9", "1.6M", "~95%", "“Get 6 perfumes free” sampling offer (1.2M); “It's parfum, not perfume”"],
    ["indē wild", "99", "0.9M", "0%", "Champi and desi-beauty community shorts"],
    ["Denver", "20", "0.8M", "~99%", "Intense EDP “Success rules the world” (830K paid); product shorts"],
    ["Plum / EM5", "19 / 2", "0.3M / 0.2M", "~68% / ~79%", "Plum: CeraSense skincare. EM5: The Salt, Myself Absolute ₹499"],
  ], { y: 1.25, colW: [1.2, 0.8, 1.0, 1.2, 4.8], size: 10, rowH: 0.42 });
  source(s, "Paid share estimated: videos with 50K+ views and under 0.2% likes plus comments. Quiet since June: Bellavita (last 17 Jan), Secret Alchemist (10 Feb), Park Avenue (Apr 2025).");
}
// Trends
{
  const s = slideBase(false);
  title(s, "Beardo and Bellavita lead search interest; Secret Alchemist's doubled");
  barChart(s, ["Beardo (+23%)", "Bellavita (+1%)", "Ajmal (+39%)", "Denver (+11%)", "Secret Alchemist (+102%)", "indē wild (−8%)", "BlaBliBlu (+45%)", "Fraganote (+15%)", "EM5 (+23%)", "Plum (−37%)"], [215, 100, 69, 44, 24, 20, 16, 12, 11, 6], { x: 0.4, y: 1.25, w: 6.0, h: 3.85, name: "Index" });
  bullets(s, [["Last Diwali was Bellavita's and Denver's peak week. ", "Expect the same spike now."], ["Week of 5 July lifted BlaBliBlu, Fraganote and Plum ", "together: a marketplace sale."], ["Rising: ", "“bella vita pistachio” +400%, “ceo man” +80%, Wild Stone Edge +150%."]], 6.7, 1.35, 2.8, 3.7, 12);
  source(s, "Google Trends, India, Jul–Oct average (Bellavita = 100); change vs previous 9 months. Engage (1.5) and Park Avenue (1.3) are under-counted.");
}

// ===== Section 3 =====
divider("3", "Brand by brand", "One page per competitor: what they launched, advertised and posted");
brandSlide("Denver", "Denver: Shah Rukh Khan at the top, deep bundle discounts underneath",
  [["219", "Meta ads"], ["40", "new Google ads in Sep"], ["0.30%", "IG engagement (best)"], ["0.8M", "YT views, ~99% paid"], ["44", "Search (+11%)"]],
  [["Intense EDP with SRK (3 Jun): ", "“Success Rules The World” TVC, ₹999, 20% oil."], ["Rizz God (18 Sep): ", "Gen Z “Vibe Collection” EDP, ₹2,500, 25% oil, D2C only."], ["Meta is all offers: ", "any 3 for ₹1,000, deo packs ₹599–1,200, Diwali buy-1-get-1 + daily MacBook draw."], ["SRK-birthday contest (Oct): ", "17.4K likes, 21.9K comments; iPhones as prizes."], ["Trade: ", "booth at Beautyworld Dubai; Hamilton Intense “SRK collection”."]],
  [["ig_denver_contest", L("ig_denver_contest"), "SRK contest", "View post"], ["denver_diwali", L("denver_diwali"), "Diwali BOGO", "View ad"]]);
brandSlide("BlaBliBlu", "BlaBliBlu: ₹100 Cr run-rate in 6 months, now pushing gourmands and Diwali gifts",
  [["429", "Meta ads"], ["45", "new Google ads in Sep"], ["150K", "IG followers, 69 posts"], ["1.6M", "YT views, ~95% paid"], ["16", "Search (+45%)"]],
  [["Scale: ", "₹100 Cr annual run-rate within 6 months (Mar); sub-₹1,000 parfums, 25% oil."], ["Price hook: ", "build-your-own box, 3 for ₹650–750; “get 6 perfumes free” sampling (1.2M YT views)."], ["Gourmand drop (21–24 Sep): ", "Berry Baddie, Watermelon Buddy, Mango So Goated, Happy Bananana."], ["Diwali (from 1 Oct): ", "House Party sale with free shot glasses, candle, card game; “boring gifts are retiring” series."], ["Tone: ", "Hinglish memes and trends; “Ab machega bhaukaal” tease got 251K views."]],
  [["ig_bbb_gourmand", L("ig_bbb_gourmand"), "Gourmand drop", "View post"], ["bbb_diwali", L("bbb_diwali"), "Diwali sale", "View ad"]]);
brandSlide("Bellavita", "Bellavita: biggest following and search share, with all new work on Meta",
  [["422", "Meta ads"], ["348 days", "median Google ad age"], ["943K", "IG followers, 90 posts"], ["0", "YT uploads since Jan"], ["100", "Search (index base)"]],
  [["Beast Mode for Men (24 Jun): ", "“40% perfume oil”, half of its top Meta ads; Oud Supreme + flask ₹799."], ["Gourmand Trend Series (26 Sep): ", "5 dessert scents at ₹549; Vanilluxe ₹599."], ["Bundles: ", "any 2 ₹899, 3 ₹1,298, 4 ₹1,599, sold as gifting."], ["Regional creators ", "(Tamil, Malayalam, Kannada) run “comment for link” reaction ads."], ["Search: ", "flat (+1%); peaked last Diwali; “pistachio” +400%, “ceo man” +80%."]],
  [["bellavita_gift", L("bellavita_gift"), "Any 3 ₹1,298", "View ad"], ["bv_creator", L("bv_creator"), "Creator ad", "View ad"]]);
brandSlide("indē wild", "indē wild: champi hair care, a global Sephora push and a move into mists",
  [["72", "Meta ads"], ["27", "Google creatives"], ["135", "IG posts (most)"], ["0.9M", "YT views, all organic"], ["20", "Search (−8%)"]],
  [["Global: ", "Sephora US (Mar) and Sephora UK Carnaby store (Jul)."], ["Champi: ", "campaign with Anupama Parameswaran (Aug); Clarifying & Detox Shampoo on Nykaa (30 Sep) with Rida Tharana."], ["Into fragrance: ", "hair and body mists After Heat and Solar Skin (Aug)."], ["Stunts and collabs: ", "Dino Morea “interim CEO” (49.7K likes); Bombay Sweet Shop Rakhi box; founder birthday sale 29% off."], ["Ads: ", "mostly creator videos; “96 hours oil control” claim."]],
  [["ig_indewild_dino", L("ig_indewild_dino"), "Dino Morea stunt", "View post"], ["ig_indewild_mists", L("ig_indewild_mists"), "Solar Skin mist", "View post"]]);
brandSlide("Secret Alchemist", "Secret Alchemist: clean, premium, Samantha-led, and search interest doubled",
  [["373", "Meta ads"], ["101 days", "median Google ad age"], ["144K", "IG followers, 55 posts"], ["0", "YT uploads since Feb"], ["24", "Search (+102%)"]],
  [["Funding: ", "$3M led by Unilever Ventures (Jan); claims 8x growth in FY26."], ["Product: ", "Clean Perfume (IFRA, 25%); Sandalwood Nostalgia (Jul); Rare Vanilla “one sold every minute”; Rose Oud; hair perfumes."], ["Gifting Collection (6 Aug): ", "pushed through Rakhi; gift boxes from ₹1,299."], ["Samantha fronts half its top ads; ", "“four women, one scent” series; now on Swiggy Instamart."], ["Giveaways: ", "with Isharya, Winston and Dyson × Instamart."]],
  [["sa_samantha", L("sa_samantha"), "Samantha ad", "View ad"], ["ig_sa_fourwomen", L("ig_sa_fourwomen"), "Four women", "View post"]]);
brandSlide("Beardo", "Beardo: the heaviest advertiser, built around the Godfather perfume franchise",
  [["960", "Meta ads (most)"], ["48 of 60", "Google ads are text"], ["808K", "IG followers, 70 posts"], ["3.6M", "YT views, ~15% paid"], ["215", "Search (+23%)"]],
  [["Godfather franchise: ", "film with Danish Pandor (Apr); Mafia and Revenge (₹1,700) at Flipkart Glam Up Fest and Amazon Beautyverse (Jun)."], ["“Sold every 15 seconds”: ", "2.9M YouTube views."], ["Brand anthem “Unstoppable” (25 Sep): ", "built from a July rap battle series; 10.1K likes."], ["Meta: ", "4 perfumes ₹999, Whisky Smoke Bourbon ₹499; de-tan, curls and beard kits; Blinkit “in 10 minutes”."], ["Search peaked 4 Oct, ", "highest of the 12 (includes grooming searches)."]],
  [["ig_beardo_anthem", L("ig_beardo_anthem"), "Brand anthem", "View post"], ["beardo_4", L("beardo_4"), "4 for ₹999", "View ad"]]);
brandSlide("Park Avenue", "Park Avenue: a women's launch and a refresh in July and August, then silence",
  [["—", "no Meta page found"], ["—", "Google under Godrej"], ["43K", "IG followers, 6 posts"], ["0", "YT uploads since 2025"], ["1.3", "Search (flat)"]],
  [["“for Her” (10 Jul): ", "Godrej's first women's EDPs, four variants (Poise, Grace, Verve, Allure), pitched for gifting."], ["Refresh (Aug): ", "new-look EDP collection and Euphoria; one boosted reel reached 597K views."], ["Since 28 Aug: ", "no Instagram posts, no YouTube uploads."], ["Ads: ", "not separable from Godrej's accounts on Meta or Google."], ["Read: ", "a big brand with a new women's range and little visible support behind it."]],
  [["ig_pa_euphoria", L("ig_pa_euphoria"), "New-look Euphoria", "View post"]]);
brandSlide("Engage", "Engage: invisible in the press, but the biggest paid YouTube push",
  [["127", "Meta ads (all roll-ons)"], ["—", "Google under ITC"], ["37K", "IG followers, 19 posts"], ["24.0M", "YT views, ~100% paid"], ["1.5", "Search (under-counted)"]],
  [["YouTube burst (Sep): ", "32 roll-on ads in Hindi, Bengali and Kannada; one cut reached 9.4M views."], ["Roll-ons: ", "Radiant Glow “instant skin brightening”, Amber Fresh “72 hours”, FreshEnCap technology."], ["Meta: ", "100% catalogue ads linking to Blinkit and Zepto."], ["Perfume: ", "Mini Perfume Spray TVC (Assamese, Jun); Maracujá range at Amazon Beautyverse; Verona."], ["Read: ", "the mass player is fighting on deo roll-ons and regional reach, not fine fragrance."]],
  [["engage_roll", L("engage_roll"), "Radiant Glow ad", "View ad"], ["yt_engage", YT("TpHySVBnido"), "9.4M-view cut", "Watch"]]);
brandSlide("Ajmal", "Ajmal: oud heritage, Ranveer Singh and a fast retail expansion",
  [["339", "Meta ads"], ["46 days", "median Google ad age"], ["273K", "IG followers, 47 posts"], ["3.6M", "YT views, ~42% paid"], ["69", "Search (+39%)"]],
  [["Ranveer Singh (Mar): ", "replaced Saif Ali Khan; “Your Unseen Power” on 575+ PVR screens."], ["Retail: ", "170+ exclusive stores by end-2026; first UK store, Westfield London (1 Aug)."], ["Experiences: ", "Vogue Wedding Atelier masterclass, Umang, GQ Heroes event scent, AYC Mumbai."], ["Story: ", "75 years of Ajmal; “Authors of Ajmal” master-perfumer series."], ["Meta: ", "oud trio, Cyan; any 3 ₹1,799; Gold Series gift set ₹899."]],
  [["ig_ajmal_london", L("ig_ajmal_london"), "London store", "View post"], ["ajmal_oud", L("ajmal_oud"), "Oud trio", "View ad"]]);
brandSlide("EM5", "EM5: a small online footprint, a cricket-star investor and story-led scents",
  [["—", "no Meta page found"], ["15", "Google creatives"], ["117K", "IG followers, 8 posts"], ["0.2M", "YT views"], ["11", "Search (+23%)"]],
  [["Suryakumar Yadav (Feb): ", "investor and partner; Kachra Seth parody video."], ["Retail: ", "about 50 stores, 111 planned; airport sampling; 70–80% of sales online."], ["Launches: ", "“Kashmir” (Jul–Aug); Inglorious Ginger and The Great Oud (25 Sep); DriveTherapy coffee car perfume (9 Sep)."], ["Price: ", "Myself Absolute EDP at ₹499."], ["Read: ", "high engagement per follower (0.23%) from very few posts."]],
  [["ig_em5_kashmir", L("ig_em5_kashmir"), "Kashmir", "View post"], ["ig_em5_launch", L("ig_em5_launch"), "New launches", "View post"]]);
brandSlide("Fraganote", "Fraganote: the most consistent publisher, built on gourmand and oud",
  [["231", "Meta ads"], ["28", "new Google ads in Sep"], ["90K", "IG followers, 37 posts"], ["5.0M", "YT views (92 uploads)"], ["12", "Search (+15%)"]],
  [["Funding: ", "$3M Series A (Jun); kiosks in 14 cities, 100 by year-end."], ["Launches: ", "Oud Collection (17 Jul), Hypnotic Oud, Sports Mode (7 Aug), 25 ml minis (18 Sep)."], ["Gourmands: ", "Baked Vanilla, Drunken Cake (750K YT views), Devil's Berry; Creme Castle bonbon collab."], ["Gifting: ", "Rakhi sale up to 45% off + 3 free minis; Diwali Edit (9 Oct)."], ["Ads: ", "short-lived Google creatives (median 20 days) = constant testing."]],
  [["fraganote_minis", L("fraganote_minis"), "25 ml minis", "View ad"], ["yt_fraganote", YT("_kFW1VDYPAA"), "Oud Collection", "Watch"]]);
brandSlide("Plum", "Plum: all skincare this quarter; fragrance sits with BodyLovin'",
  [["215", "Meta ads (skincare)"], ["60", "Google creatives"], ["1.15M", "IG followers, 76 posts"], ["0.3M", "YT views"], ["6", "Search (−37%)"]],
  [["BodyLovin' relaunch (Apr): ", "independent identity, 16 launches, offbeat lip balms; heavy OTT and vernacular YouTube plans."], ["This quarter: ", "CeraSense barrier care and the “BYOB” Delhi party; niacinamide jelly serum."], ["Meta: ", "skincare only; 3 for ₹599; dandruff shampoo on Blinkit."], ["Read: ", "not an active fragrance competitor this quarter; watch BodyLovin' mists."]],
  [["ig_plum_byob", L("ig_plum_byob"), "BYOB party", "View post"]]);

// ===== Section 4 =====
divider("4", "What it means for Mcnroe", "Launch calendar, implications and how we keep tracking");
// Calendar
{
  const s = slideBase(false);
  title(s, "Competitors launched something almost every week from July to October");
  const months = {
    July: ["19 Jul · Fraganote Hypnotic Oud", "Jul · Beardo rap battle series", "Jul · indē wild Sephora UK"],
    August: ["1 Aug · Ajmal first London store", "6 Aug · Secret Alchemist Gifting Collection", "7 Aug · Fraganote Sports Mode", "8 Aug · Park Avenue new-look EDPs", "Aug · indē wild hair and body mists"],
    September: ["18 Sep · Denver Rizz God (Gen Z)", "18 Sep · Fraganote 25 ml minis", "21–24 Sep · BlaBliBlu 4 gourmands", "25 Sep · Beardo anthem “Unstoppable”", "26 Sep · Bellavita Gourmand series"],
    October: ["1 Oct · BlaBliBlu Diwali sale", "8 Oct · Denver SRK iPhone contest", "9 Oct · Fraganote Diwali Edit", "13 Oct · Denver Diwali buy 1 get 1"],
  };
  Object.entries(months).forEach(([m, items], i) => {
    const x = 0.5 + i * 2.3, w = 2.1;
    s.addText(m, { x, y: 1.3, w, h: 0.4, fontFace: HEAD, fontSize: 16, bold: true, color: i === 3 ? COL.plum : COL.ink, margin: 0, isTextBox: true });
    items.forEach((t, j) => {
      const y = 1.78 + j * 0.66;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.56, fill: { color: i === 3 ? "F3E7EC" : COL.card }, line: { color: i === 3 ? "F3E7EC" : COL.card }, rectRadius: 0.06 });
      const [d, rest] = t.split(" · ");
      s.addText([{ text: d, options: { bold: true, color: COL.plum, breakLine: true } }, { text: rest, options: { color: COL.ink } }], { x: x + 0.1, y: y + 0.03, w: w - 0.2, h: 0.5, fontFace: BODY, fontSize: 10.5, valign: "middle", margin: 0, isTextBox: true });
    });
  });
  source(s, "Park Avenue “for Her” launched 10 July, just before the window.");
}
// Implications
{
  const s = slideBase(false);
  title(s, "What this means for Mcnroe");
  const items = [
    ["Lead Diwali with a gifting idea, not a deeper discount.", "Four brands already run buy-1-get-1 and free-gift offers."],
    ["State a proof of longevity.", "Competitors print 20–40% oil; name a concentration or tested hours."],
    ["Price singles against “any 3 for ₹1,000”.", "Shoppers see 3-perfume bundles from ₹650 every day."],
    ["Go where it's less crowded.", "Fresh and floral scents, and women's, have few strong voices."],
    ["Use creator-run ads and quick commerce.", "Both are normal for Secret Alchemist, Bellavita, Beardo and Engage."],
    ["Plan launches around marketplace events.", "Flipkart Glam Up Fest, Amazon Beautyverse and festive sales."],
  ];
  items.forEach(([b, t], i) => {
    const y = 1.25 + i * 0.64;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.04, w: 0.44, h: 0.44, fill: { color: COL.plum }, line: { color: COL.plum } });
    s.addText(String(i + 1), { x: 0.5, y: y + 0.04, w: 0.44, h: 0.44, fontFace: HEAD, fontSize: 15, bold: true, color: COL.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText([{ text: b + " ", options: { bold: true } }, { text: t, options: { color: COL.muted } }], { x: 1.15, y, w: 8.35, h: 0.52, fontFace: BODY, fontSize: 14, color: COL.ink, valign: "middle", margin: 0, isTextBox: true });
  });
}
// Ongoing tracking
{
  const s = slideBase(false);
  title(s, "Repeating this scan every 3 months costs about $19 in data fees");
  table(s, ["Layer", "Tool", "Cost per quarterly scan"], [
    ["Meta ads (all ads)", "Apify Facebook Ads Library Scraper", "~$12"],
    ["Google ads", "Apify Google Ads Transparency Scraper", "~$2–3"],
    ["Instagram (every post)", "Instagram Graph API (set up)", "Free"],
    ["YouTube", "YouTube Data API", "Free"],
    ["Google Trends", "Apify Google Trends Scraper", "~$0.50"],
    ["Analysis and report", "Claude", "Existing plan"],
  ], { w: 6.0, colW: [2.0, 2.6, 1.4], size: 12, rowH: 0.36, align: ["left", "left", "right"] });
  bullets(s, [["Total: ", "~$16 per scan; one Apify Starter month ($19) covers it."], ["Optional: ", "comment sentiment, ~$15 more per quarter."], ["Not available anywhere: ", "competitor ad spend and reach."]], 6.8, 1.35, 2.7, 3.5, 12.5);
}
// Appendix
{
  const s = slideBase(false);
  title(s, "Appendix: how this scan was made");
  table(s, ["Layer", "Source", "Coverage"], [
    ["Meta ads", "Meta Ad Library (Apify)", "Counts for 10 brands; top 60 ads each read"],
    ["Google ads", "Google Ads Transparency Center (Apify)", "Up to 60 creatives per brand"],
    ["Instagram", "Instagram Graph API", "All 663 posts, 12 brands"],
    ["YouTube", "YouTube Data API", "All 390 uploads since 1 June"],
    ["Search", "Google Trends (Apify)", "12 brands, India, 12 months"],
    ["Press", "Trade press and press releases", "Launches, funding, campaigns"],
  ], { colW: [1.6, 3.4, 4.0], size: 12, rowH: 0.34 });
  s.addText([{ text: "Full write-up: ", options: { bold: true } }, { text: "Competitor Scan doc", options: { color: COL.plum, underline: true, hyperlink: { url: DOC } } }, { text: "    Raw data: ", options: { bold: true } }, { text: "project repository", options: { color: COL.plum, underline: true, hyperlink: { url: REPO } } }], { x: 0.5, y: 4.0, w: 9, h: 0.35, fontFace: BODY, fontSize: 13, color: COL.ink, margin: 0, isTextBox: true });
  s.addText("Caveats: no source gives competitor ad spend or reach; YouTube paid views are inferred from like rates; brand claims (oil %, growth) are unverified; press-release figures are company claims.", { x: 0.5, y: 4.45, w: 9, h: 0.5, fontFace: BODY, fontSize: 11, color: COL.muted, margin: 0, isTextBox: true });
}

pres.writeFile({ fileName: path.join(__dirname, "comp-scan-full.pptx") }).then((f) => console.log("wrote", f));
