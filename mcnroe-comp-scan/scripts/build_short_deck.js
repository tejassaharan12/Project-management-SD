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
pres.title = "Competitor Scan: Fragrance and Grooming, Jul to Oct 2026";
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

// 1. Title
{
  const s = slideBase(true);
  s.addText("Competitor Scan", { x: 0.6, y: 1.3, w: 8.8, h: 0.9, fontFace: HEAD, fontSize: 44, bold: true, color: COL.white, margin: 0, isTextBox: true });
  s.addText("Fragrance and grooming in India, 15 July to 15 October 2026", { x: 0.6, y: 2.2, w: 8.8, h: 0.5, fontFace: BODY, fontSize: 20, color: "E9D8B4", margin: 0, isTextBox: true });
  s.addText("Denver · BlaBliBlu · Bellavita · indē wild · Secret Alchemist · Beardo · Park Avenue · Engage · Ajmal · EM5 · Fraganote · Plum", { x: 0.6, y: 3.1, w: 8.8, h: 0.6, fontFace: BODY, fontSize: 13, color: "CFC4C8", margin: 0, isTextBox: true });
  s.addText("Prepared for Mcnroe · October 2026", { x: 0.6, y: 4.6, w: 8.8, h: 0.4, fontFace: BODY, fontSize: 13, color: "CFC4C8", margin: 0, isTextBox: true });
}

// 2. Executive summary
{
  const s = slideBase(false);
  title(s, "Executive summary");
  const items = [
    ["Diwali selling started on 1 October. ", "BlaBliBlu, Denver, Fraganote and Bellavita already run festive offers on Meta."],
    ["“Any 3” bundles set the price frame: ", "3 perfumes for ₹650 (BlaBliBlu) to ₹1,799 (Ajmal)."],
    ["Perfume-oil % is the quality claim. ", "Bellavita now claims 40%, against the usual 25%."],
    ["Beardo is the heaviest Meta advertiser ", "(960 ads); Engage ran the biggest paid YouTube push (24M views) with no press coverage."],
    ["Contests and stunts win on Instagram, ", "earning 10 to 40 times normal engagement."],
  ];
  const runs = [];
  items.forEach(([b, t], i) => {
    runs.push({ text: b, options: { bold: true, bullet: true } });
    runs.push({ text: t, options: { breakLine: i < items.length - 1 } });
  });
  s.addText(runs, { x: 0.5, y: 1.25, w: 5.9, h: 3.8, fontFace: BODY, fontSize: 15, color: COL.ink, valign: "top", paraSpaceAfter: 10, margin: 0, isTextBox: true });
  const stats = [["3,387", "Meta ads"], ["663", "Instagram posts"], ["390", "YouTube videos"]];
  stats.forEach(([n, l], i) => {
    const y = 1.25 + i * 1.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.8, y, w: 2.7, h: 1.1, fill: { color: COL.card }, line: { color: COL.card }, rectRadius: 0.1 });
    s.addText([{ text: n, options: { fontSize: 30, bold: true, color: COL.plum, fontFace: HEAD, breakLine: true } }, { text: l + " analysed", options: { fontSize: 12, color: COL.muted } }], { x: 6.95, y: y + 0.08, w: 2.4, h: 0.95, fontFace: BODY, valign: "middle", margin: 0, isTextBox: true });
  });
  source(s, "Sources: Meta Ad Library, Google Ads Transparency Center, Instagram Graph API, YouTube Data API, Google Trends, trade press. Pulled 15 October 2026.");
}

// 3. Diwali already live
{
  const s = slideBase(false);
  title(s, "Diwali selling began on 1 October, more than a month before the festival");
  const cards = [
    ["denver_diwali", "Denver", "Buy 1 get 1 + daily MacBook draw, from 13 Oct"],
    ["bbb_diwali", "BlaBliBlu", "Diwali House Party sale, free gifts, from 1 Oct"],
    ["fraganote_diwali", "Fraganote", "Festive 50 ml sets, Diwali Edit, from 9 Oct"],
    ["bellavita_gift", "Bellavita", "Any 3 perfumes for ₹1,298, gifting-ready"],
  ];
  cards.forEach(([k, b, t], i) => {
    const x = 0.5 + i * 2.3;
    img(s, k, x, 1.3, 2.05, 2.75, L(k));
    caption(s, x, 4.12, 2.05, b, t, L(k), "View ad");
  });
  source(s, "Diwali is on 8 November 2026. Ads from Meta Ad Library, India; click an image to open the ad.");
}

// 4. Any 3 bundles
{
  const s = slideBase(false);
  title(s, "“Any 3” bundles set the price frame: 3 perfumes for ₹650 to ₹1,799");
  barChart(s, ["BlaBliBlu", "Denver", "Bellavita", "Ajmal"], [650, 1000, 1298, 1799], { x: 0.4, y: 1.3, w: 4.9, h: 3.0, fmt: "₹#,##0", name: "Price of 3 perfumes" });
  s.addText("Price of a 3-perfume bundle in Meta ads (₹)", { x: 0.5, y: 4.3, w: 4.8, h: 0.3, fontFace: BODY, fontSize: 11, color: COL.muted, margin: 0, isTextBox: true });
  s.addText([{ text: "Build-your-own box appears in 6 of 10 advertisers. ", options: { bold: true } }, { text: "Beardo sells 4 perfumes for ₹999." }], { x: 0.5, y: 4.6, w: 4.8, h: 0.55, fontFace: BODY, fontSize: 12, color: COL.ink, margin: 0, isTextBox: true });
  img(s, "denver_any3", 5.6, 1.3, 1.9, 1.9, L("denver_any3"));
  img(s, "bbb_any3", 7.6, 1.3, 1.9, 1.9, L("bbb_any3"));
  caption(s, 5.6, 3.3, 1.9, "Denver", "Any 3 for ₹1,000", L("denver_any3"), "View ad");
  caption(s, 7.6, 3.3, 1.9, "BlaBliBlu", "3 for ₹750", L("bbb_any3"), "View ad");
  source(s, "Offers seen in Meta ads, August to October 2026. BlaBliBlu runs ₹650 and ₹750 versions.");
}

// 5. Oil percentage
{
  const s = slideBase(false);
  title(s, "Perfume-oil percentage is the quality claim, and Bellavita now claims 40%");
  barChart(s, ["Bellavita Beast Mode", "BlaBliBlu", "Secret Alchemist", "Denver Rizz God", "Denver Intense"], [40, 25, 25, 25, 20], { x: 0.4, y: 1.3, w: 5.6, h: 3.3, fmt: "0\"%\"", name: "Oil %" });
  s.addText("Perfume-oil concentration claimed in ads and launches", { x: 0.5, y: 4.6, w: 5.5, h: 0.3, fontFace: BODY, fontSize: 11, color: COL.muted, margin: 0, isTextBox: true });
  img(s, "bellavita_beast", 6.6, 1.3, 2.9, 3.0, L("bellavita_beast"));
  caption(s, 6.6, 4.35, 2.9, "Bellavita Beast Mode", "Collector's flask with Oud Supreme, 40% oil", L("bellavita_beast"), "View ad");
  source(s, "Brand claims, not lab-tested. Sources: Meta Ad Library; Denver launch press releases (June and September 2026).");
}

// 6. Meta ad volume
{
  const s = slideBase(false);
  title(s, "Beardo ran 960 Meta ads since mid-July, more than twice the next brand");
  barChart(s, ["Beardo", "BlaBliBlu", "Bellavita", "Secret Alchemist", "Ajmal", "Fraganote", "Denver", "Plum", "Engage", "indē wild"], [960, 429, 422, 373, 339, 231, 219, 215, 127, 72], { x: 0.4, y: 1.25, w: 6.0, h: 3.85, name: "Meta ads" });
  const pts = [
    ["Beardo ", "spreads spend across perfume combos, de-tan, beard and curl kits."],
    ["Engage ", "runs only catalogue ads for roll-ons, linking to Blinkit and Zepto."],
    ["Park Avenue and EM5 ", "have no findable Meta advertiser page."],
  ];
  const runs = [];
  pts.forEach(([b, t], i) => { runs.push({ text: b, options: { bold: true, bullet: true } }); runs.push({ text: t, options: { breakLine: i < pts.length - 1 } }); });
  s.addText(runs, { x: 6.7, y: 1.4, w: 2.8, h: 3.5, fontFace: BODY, fontSize: 13, color: COL.ink, valign: "top", paraSpaceAfter: 10, margin: 0, isTextBox: true });
  source(s, "Ads active in India, 15 July to 15 October 2026. Source: Meta Ad Library via Apify.");
}

// 7. Creators
{
  const s = slideBase(false);
  title(s, "Creators front a large share of paid ads, and quick commerce is a common destination");
  const cards = [
    ["sa_samantha", "Secret Alchemist", "Half its top ads run from Samantha's handle"],
    ["bv_creator", "Bellavita", "Regional creators post “comment for link” reactions"],
    ["beardo_creator", "Beardo", "Creator reviews of Godfather Mafia and Revenge"],
    ["indewild_creator", "indē wild", "Most launch ads are creator videos"],
  ];
  cards.forEach(([k, b, t], i) => {
    const x = 0.5 + i * 2.3;
    img(s, k, x, 1.3, 2.05, 2.6, L(k));
    caption(s, x, 3.97, 2.05, b, t, L(k), "View ad");
  });
  s.addText([{ text: "Quick commerce: ", options: { bold: true } }, { text: "Engage, Beardo, Secret Alchemist and Plum send ads straight to Blinkit, Zepto or Instamart." }], { x: 0.5, y: 4.85, w: 9, h: 0.35, fontFace: BODY, fontSize: 12, color: COL.ink, margin: 0, isTextBox: true });
}

// 8. Gourmand and oud
{
  const s = slideBase(false);
  title(s, "Gourmand and oud are this season's scent stories; fresh and floral are less crowded");
  const cards = [
    ["ig_bbb_gourmand", "BlaBliBlu", "4 fruity gourmands, one a day, 21–24 Sep", L("ig_bbb_gourmand"), "View post"],
    ["fraganote_minis", "Fraganote", "Devil's Berry and dessert scents, new 25 ml minis", L("fraganote_minis"), "View ad"],
    ["ajmal_oud", "Ajmal", "Oud trio: Oud Nirvana, Oud on the Rocks, White Oud", L("ajmal_oud"), "View ad"],
    ["yt_fraganote", "Fraganote", "Oud Collection launch, 17 Jul (570K views)", YT("_kFW1VDYPAA"), "Watch"],
  ];
  cards.forEach(([k, b, t, lk, lt], i) => {
    const x = 0.5 + i * 2.3;
    img(s, k, x, 1.3, 2.05, 2.6, lk);
    caption(s, x, 3.97, 2.05, b, t, lk, lt);
  });
  s.addText([{ text: "Also: ", options: { bold: true } }, { text: "Bellavita's Gourmand Trend Series at ₹549 (26 Sep). Secret Alchemist's jasmine and tuberose are the main floral voice." }], { x: 0.5, y: 4.85, w: 9, h: 0.35, fontFace: BODY, fontSize: 12, color: COL.ink, margin: 0, isTextBox: true });
}

// 9. Google ads table
{
  const s = slideBase(false);
  title(s, "Denver, BlaBliBlu and Fraganote ramped up Google ads in September; Bellavita's are a year old");
  const hdr = ["Brand", "Creatives (cap 60)", "Median days live", "New in September"].map((t, i) => ({ text: t, options: { bold: true, align: i ? "right" : "left" } }));
  const rows = [
    ["Denver", "60", "12", "40"], ["BlaBliBlu", "60", "22", "45"], ["Fraganote", "60", "20", "28"], ["Plum", "60", "35", "22"],
    ["Ajmal", "60", "46", "14"], ["Secret Alchemist", "60", "101", "13"], ["Beardo", "60", "84", "2"], ["indē wild", "27", "16", "0"],
    ["EM5", "15", "227", "1"], ["Bellavita", "60", "348", "1"],
  ].map((r) => r.map((t, i) => ({ text: t, options: { bold: i === 0, align: i ? "right" : "left" } })));
  rows.forEach((r) => { if (["Denver", "BlaBliBlu", "Fraganote"].includes(r[0].text)) r.forEach((c) => (c.options.fill = { color: "F3E7EC" })); });
  s.addTable([hdr, ...rows], { x: 0.5, y: 1.3, w: 6.2, colW: [2.0, 1.4, 1.4, 1.4], fontFace: BODY, fontSize: 12, color: COL.ink, rowH: 0.31, border: { type: "solid", pt: 0.75, color: COL.line } });
  const runs = [
    { text: "Short-lived creatives = testing. ", options: { bold: true, breakLine: true } },
    { text: "Denver's median Google ad ran 12 days; it launched 40 new ones in September.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Bellavita is Meta-only for new work. ", options: { bold: true, breakLine: true } },
    { text: "Beast Mode and Gourmand never reached its Google ads." },
  ];
  s.addText(runs, { x: 7.0, y: 1.35, w: 2.5, h: 3.5, fontFace: BODY, fontSize: 13, color: COL.ink, valign: "top", margin: 0, isTextBox: true });
  source(s, "Google Ads Transparency Center, India, ads shown 15 Jul–15 Oct 2026. Engage and Park Avenue advertise under ITC and Godrej and could not be separated.");
}

// 10. Instagram stunts
{
  const s = slideBase(false);
  title(s, "On Instagram, contests and stunts earned 10 to 40 times normal engagement");
  const cards = [
    ["ig_denver_contest", "Denver: SRK-birthday iPhone contest", "17.4K likes, 21.9K comments (median post: 457 likes)"],
    ["ig_indewild_dino", "indē wild: Dino Morea as “interim CEO”", "49.7K likes (median post: 391)"],
    ["ig_beardo_anthem", "Beardo: brand anthem “Unstoppable”", "10.1K likes (median post: 284)"],
  ];
  cards.forEach(([k, b, t], i) => {
    const x = 0.5 + i * 3.08;
    img(s, k, x, 1.3, 2.85, 2.65, L(k));
    caption(s, x, 4.02, 2.85, b, t, L(k), "View post");
  });
  source(s, "Instagram Graph API (Business Discovery), all posts 15 Jul–15 Oct 2026. Denver's contest image shows its matching Meta ad creative.");
}

// 11. Engagement per follower
{
  const s = slideBase(false);
  title(s, "Denver and EM5 get the most engagement per follower; the biggest accounts get the least");
  barChart(s, ["Denver", "EM5", "indē wild", "Fraganote", "BlaBliBlu", "Secret Alchemist", "Engage", "Ajmal", "Park Avenue", "Bellavita", "Beardo", "Plum"], [0.30, 0.23, 0.16, 0.15, 0.12, 0.11, 0.11, 0.10, 0.06, 0.04, 0.04, 0.03], { x: 0.4, y: 1.25, w: 5.8, h: 3.9, fmt: "0.00\"%\"", name: "Engagement" });
  const runs = [
    { text: "Median (likes + comments) ÷ followers per post.", options: { color: COL.muted, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Bellavita (943K), Plum (1.15M) and Beardo (808K) ", options: { bold: true } },
    { text: "post often but convert few followers into reactions.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "indē wild posts most: ", options: { bold: true } },
    { text: "135 posts in 3 months, 84% reels." },
  ];
  s.addText(runs, { x: 6.6, y: 1.35, w: 2.9, h: 3.6, fontFace: BODY, fontSize: 13, color: COL.ink, valign: "top", margin: 0, isTextBox: true });
  source(s, "Instagram Graph API, 663 posts from 12 accounts, 15 Jul–15 Oct 2026.");
}

// 12. YouTube
{
  const s = slideBase(false);
  title(s, "Engage's regional roll-on ads drew 24M YouTube views, more than all others combined");
  barChart(s, ["Engage", "Fraganote", "Beardo", "Ajmal", "BlaBliBlu", "indē wild", "Denver", "Plum", "EM5"], [24.0, 5.0, 3.6, 3.6, 1.6, 0.9, 0.8, 0.3, 0.2], { x: 0.4, y: 1.25, w: 5.2, h: 3.85, fmt: "0.0\"M\"", name: "Views" });
  img(s, "yt_engage", 5.9, 1.3, 3.6, 1.6, YT("TpHySVBnido"));
  caption(s, 5.9, 2.95, 3.6, "Engage roll-on, Bengali cut", "9.4M views, 3 likes: almost all paid", YT("TpHySVBnido"), "Watch");
  img(s, "yt_beardo", 5.9, 3.65, 1.75, 1.0, YT("AbEH-c62jwA"));
  s.addText([{ text: "Beardo Godfather", options: { bold: true, breakLine: true } }, { text: "2.9M views", options: { color: COL.muted, breakLine: true } }, { text: "Watch", options: { color: COL.plum, underline: true, hyperlink: { url: YT("AbEH-c62jwA") } } }], { x: 7.8, y: 3.65, w: 1.7, h: 1.0, fontFace: BODY, fontSize: 11, color: COL.ink, valign: "top", margin: 0, isTextBox: true });
  source(s, "YouTube Data API, uploads 1 Jun–15 Oct 2026, views to date. Bellavita, Secret Alchemist and Park Avenue posted nothing since June.");
}

// 13. Trends
{
  const s = slideBase(false);
  title(s, "Beardo and Bellavita lead search interest; Secret Alchemist's doubled");
  barChart(s, ["Beardo (+23%)", "Bellavita (+1%)", "Ajmal (+39%)", "Denver (+11%)", "Secret Alchemist (+102%)", "indē wild (−8%)", "BlaBliBlu (+45%)", "Fraganote (+15%)", "EM5 (+23%)", "Plum (−37%)"], [215, 100, 69, 44, 24, 20, 16, 12, 11, 6], { x: 0.4, y: 1.25, w: 6.0, h: 3.85, name: "Index" });
  const runs = [
    { text: "Google search interest in India, Jul–Oct (Bellavita = 100); change vs the previous 9 months in brackets.", options: { color: COL.muted, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Last Diwali was Bellavita's and Denver's peak search week. ", options: { bold: true } },
    { text: "Expect the same spike now." },
  ];
  s.addText(runs, { x: 6.7, y: 1.35, w: 2.8, h: 3.5, fontFace: BODY, fontSize: 13, color: COL.ink, valign: "top", margin: 0, isTextBox: true });
  source(s, "Google Trends via Apify, India, 12 Oct 2025–14 Oct 2026. Engage and Park Avenue are under-counted (people search the product, not the brand).");
}

// 14. Launch calendar
{
  const s = slideBase(false);
  title(s, "Competitors launched something almost every week from July to October");
  const months = {
    "July": ["19 Jul · Fraganote Hypnotic Oud", "Jul · Beardo rap battle series"],
    "August": ["1 Aug · Ajmal first London store", "6 Aug · Secret Alchemist Gifting Collection", "7 Aug · Fraganote Sports Mode", "8 Aug · Park Avenue new-look EDPs", "Aug · indē wild hair and body mists"],
    "September": ["18 Sep · Denver Rizz God (Gen Z)", "18 Sep · Fraganote 25 ml minis", "21–24 Sep · BlaBliBlu 4 gourmands", "25 Sep · Beardo anthem “Unstoppable”", "26 Sep · Bellavita Gourmand series"],
    "October": ["1 Oct · BlaBliBlu Diwali sale", "8 Oct · Denver SRK iPhone contest", "9 Oct · Fraganote Diwali Edit", "13 Oct · Denver Diwali buy 1 get 1"],
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
  source(s, "Sources: Instagram, YouTube, Meta Ad Library and press releases. Park Avenue “for Her” launched 10 July, just before the window.");
}

// 15. Implications
{
  const s = slideBase(false);
  title(s, "What this means for Mcnroe");
  const items = [
    ["Lead Diwali with a gifting idea, not a deeper discount.", "Four brands already run buy-1-get-1 and free-gift offers."],
    ["State a proof of longevity.", "Competitors print 20–40% oil; name a concentration or tested hours."],
    ["Price singles against “any 3 for ₹1,000”.", "Shoppers see 3-perfume bundles from ₹650 every day."],
    ["Go where it's less crowded.", "Fresh and floral scents, and women's, have few strong voices."],
    ["Use creator-run ads and quick commerce.", "Both are now normal for Secret Alchemist, Bellavita, Beardo and Engage."],
  ];
  items.forEach(([b, t], i) => {
    const y = 1.3 + i * 0.75;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.04, w: 0.48, h: 0.48, fill: { color: COL.plum }, line: { color: COL.plum } });
    s.addText(String(i + 1), { x: 0.5, y: y + 0.04, w: 0.48, h: 0.48, fontFace: HEAD, fontSize: 16, bold: true, color: COL.white, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText([{ text: b + " ", options: { bold: true } }, { text: t, options: { color: COL.muted } }], { x: 1.2, y, w: 8.3, h: 0.6, fontFace: BODY, fontSize: 15, color: COL.ink, valign: "middle", margin: 0, isTextBox: true });
  });
}

// 16. Appendix
{
  const s = slideBase(false);
  title(s, "Appendix: how this scan was made");
  const hdr = ["Layer", "Source", "Coverage"].map((t) => ({ text: t, options: { bold: true } }));
  const rows = [
    ["Meta ads", "Meta Ad Library (Apify)", "Counts for 10 brands; top 60 ads each read"],
    ["Google ads", "Google Ads Transparency Center (Apify)", "Up to 60 creatives per brand"],
    ["Instagram", "Instagram Graph API", "All 663 posts, 12 brands"],
    ["YouTube", "YouTube Data API", "All 390 uploads since 1 June"],
    ["Search", "Google Trends (Apify)", "12 brands, India, 12 months"],
    ["Press", "Trade press and press releases", "Launches, funding, campaigns"],
  ].map((r) => r.map((t, i) => ({ text: t, options: { bold: i === 0 } })));
  s.addTable([hdr, ...rows], { x: 0.5, y: 1.25, w: 9, colW: [1.6, 3.4, 4.0], fontFace: BODY, fontSize: 12, color: COL.ink, rowH: 0.34, border: { type: "solid", pt: 0.75, color: COL.line } });
  s.addText([
    { text: "Full write-up: ", options: { bold: true } },
    { text: "Competitor Scan doc", options: { color: COL.plum, underline: true, hyperlink: { url: DOC } } },
    { text: "    Raw data: ", options: { bold: true } },
    { text: "project repository", options: { color: COL.plum, underline: true, hyperlink: { url: REPO } } },
  ], { x: 0.5, y: 4.0, w: 9, h: 0.35, fontFace: BODY, fontSize: 13, color: COL.ink, margin: 0, isTextBox: true });
  s.addText("Caveats: no source gives competitor ad spend or reach; YouTube paid views are inferred from like rates; brand claims (oil %, growth) are unverified.", { x: 0.5, y: 4.45, w: 9, h: 0.5, fontFace: BODY, fontSize: 11, color: COL.muted, margin: 0, isTextBox: true });
}

pres.writeFile({ fileName: path.join(__dirname, "comp-scan.pptx") }).then((f) => console.log("wrote", f));
