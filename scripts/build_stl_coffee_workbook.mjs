import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs/stl-coffee-shops";
const outputPath = `${outputDir}/best-st-louis-coffee-shops-for-work.xlsx`;

const shops = [
  {
    shop: "Coffeestamp",
    neighborhood: "Fox Park",
    address: "2511 South Jefferson Avenue, St. Louis",
    rating: 4.9,
    reviews: 210,
    workFriendly: "Yes",
    specialty: "Yes",
    bestFor: "Highest-rated coffee-forward stop; micro-roastery feel.",
    caveat: "Smaller local shop; check seating before planning a long stay.",
    source: "Workmode; My Coffee Explorer",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Coffeestamp%202511%20South%20Jefferson%20Avenue%20St.%20Louis",
  },
  {
    shop: "Kitchen House Coffee",
    neighborhood: "Tower Grove East / Benton Park West",
    address: "3149 Shenandoah Avenue, St. Louis",
    rating: 4.8,
    reviews: 284,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "High rating with neighborhood-cafe energy.",
    caveat: "Likely better for focused blocks than lots of calls.",
    source: "Workmode",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Kitchen%20House%20Coffee%203149%20Shenandoah%20Avenue%20St.%20Louis",
  },
  {
    shop: "Shameless Grounds",
    neighborhood: "Benton Park",
    address: "1901 Withnell Avenue, St. Louis",
    rating: 4.8,
    reviews: 215,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "Laptop-friendly cafe with very high public rating.",
    caveat: "Sex-positive branding/decor may not fit every work context.",
    source: "Workmode; Reddit local caveat",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Shameless%20Grounds%201901%20Withnell%20Avenue%20St.%20Louis",
  },
  {
    shop: "Sump Coffee",
    neighborhood: "South City",
    address: "3700 South Jefferson Avenue, St. Louis",
    rating: 4.7,
    reviews: 731,
    workFriendly: "Yes",
    specialty: "Yes",
    bestFor: "Best serious-coffee pick; strong roaster reputation.",
    caveat: "Can be busy; best when coffee quality matters more than all-day desk comfort.",
    source: "Workmode; My Coffee Explorer",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Sump%20Coffee%203700%20South%20Jefferson%20Avenue%20St.%20Louis",
  },
  {
    shop: "Comet Coffee",
    neighborhood: "Forest Park Southeast / Oakland",
    address: "5708 Oakland Avenue, St. Louis",
    rating: 4.7,
    reviews: 597,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "High-rated coffee and pastries near Forest Park.",
    caveat: "Seating can tighten up during peak pastry/coffee times.",
    source: "Workmode",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Comet%20Coffee%205708%20Oakland%20Avenue%20St.%20Louis",
  },
  {
    shop: "The Mud House",
    neighborhood: "Cherokee Street",
    address: "2101 Cherokee Street, St. Louis",
    rating: 4.6,
    reviews: 1432,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "Most review-backed option; good cafe/food setup.",
    caveat: "Popular spot; noise and table availability may vary.",
    source: "Workmode",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=The%20Mud%20House%202101%20Cherokee%20Street%20St.%20Louis",
  },
  {
    shop: "Rise Coffee",
    neighborhood: "The Grove",
    address: "4176 Manchester Avenue, St. Louis",
    rating: 4.6,
    reviews: 782,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "Work-friendly neighborhood cafe with strong review volume.",
    caveat: "Can fill up in The Grove; better outside peak brunch periods.",
    source: "Workmode",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Rise%20Coffee%204176%20Manchester%20Avenue%20St.%20Louis",
  },
  {
    shop: "Kaldi's Coffee - DeMun",
    neighborhood: "DeMun / Clayton edge",
    address: "700 De Mun Avenue, St. Louis",
    rating: 4.6,
    reviews: 710,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "Reliable local chain option with solid seating odds.",
    caveat: "Technically near the city edge; verify travel time from your exact spot.",
    source: "Workmode",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Kaldi%27s%20Coffee%20700%20De%20Mun%20Avenue%20St.%20Louis",
  },
  {
    shop: "Fiddlehead Fern Cafe",
    neighborhood: "Shaw / Botanical Garden",
    address: "4066 Russell Boulevard, St. Louis",
    rating: 4.6,
    reviews: 445,
    workFriendly: "Yes",
    specialty: "Yes",
    bestFor: "Botanical syrups and specialty espresso near the Garden.",
    caveat: "Published ratings differ by source; I used Workmode's review-count-backed figure.",
    source: "Workmode; My Coffee Explorer",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Fiddlehead%20Fern%20Cafe%204066%20Russell%20Boulevard%20St.%20Louis",
  },
  {
    shop: "Catalyst Coffee Bar",
    neighborhood: "Downtown West",
    address: "1223 Pine Street, St. Louis",
    rating: 4.6,
    reviews: 189,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "Central city work stop with a strong rating.",
    caveat: "Lower review count than the top picks.",
    source: "Workmode",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Catalyst%20Coffee%20Bar%201223%20Pine%20Street%20St.%20Louis",
  },
  {
    shop: "MoKaBe's Coffeehouse",
    neighborhood: "Tower Grove South",
    address: "3606 Arsenal Street, St. Louis",
    rating: 4.5,
    reviews: 1197,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "Proven long-running cafe with major review volume.",
    caveat: "Cafe atmosphere may be active; not the quietest option.",
    source: "Workmode",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=MoKaBe%27s%20Coffeehouse%203606%20Arsenal%20Street%20St.%20Louis",
  },
  {
    shop: "Kaldi's Coffee on Skinker",
    neighborhood: "Skinker-DeBaliviere",
    address: "270 Skinker Boulevard, St. Louis",
    rating: 4.5,
    reviews: 790,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "Reliable work session near WashU/Forest Park.",
    caveat: "Student traffic can make it busy.",
    source: "Workmode",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Kaldi%27s%20Coffee%20on%20Skinker%20270%20Skinker%20Boulevard%20St.%20Louis",
  },
  {
    shop: "Hartford Coffee Company",
    neighborhood: "Tower Grove South",
    address: "3974 Hartford Street, St. Louis",
    rating: 4.5,
    reviews: 714,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "Neighborhood workhorse for longer laptop sessions.",
    caveat: "Choose a non-peak window if you need a quieter table.",
    source: "Workmode; Reddit local mentions",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Hartford%20Coffee%20Company%203974%20Hartford%20Street%20St.%20Louis",
  },
  {
    shop: "Blueprint Coffee",
    neighborhood: "Delmar Loop",
    address: "6225 Delmar Boulevard, St. Louis",
    rating: 4.5,
    reviews: 641,
    workFriendly: "Yes",
    specialty: "Yes",
    bestFor: "Specialty roaster pick with a lighter-roast focus.",
    caveat: "Loop traffic/parking may be less convenient from some city neighborhoods.",
    source: "Workmode; My Coffee Explorer",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Blueprint%20Coffee%206225%20Delmar%20Boulevard%20St.%20Louis",
  },
  {
    shop: "Northwest Coffee Roasting",
    neighborhood: "Central West End / Cortex",
    address: "4251 Laclede Avenue, St. Louis",
    rating: 4.5,
    reviews: 516,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "CWE/Cortex-area work stop with patio-friendly reputation.",
    caveat: "Weather and seating matter if you prefer outdoor work.",
    source: "Workmode",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Northwest%20Coffee%20Roasting%204251%20Laclede%20Avenue%20St.%20Louis",
  },
  {
    shop: "Kaldi's Coffee - Euclid",
    neighborhood: "Central West End",
    address: "52 N. Euclid, St. Louis, MO 63108",
    rating: null,
    reviews: null,
    workFriendly: "Yes",
    specialty: "No",
    bestFor: "Reliable CWE fallback with directly confirmed free Wi-Fi.",
    caveat: "Rating not included in the Workmode ranking; listed as a practical honorable mention.",
    source: "Kaldi's official location page",
    mapUrl: "https://www.google.com/maps/search/?api=1&query=Kaldi%27s%20Coffee%2052%20N%20Euclid%20St.%20Louis",
  },
];

function scoreFormula(row) {
  return `=IF(F${row}="", "", ROUND((F${row}*20*0.7)+(MIN(LOG10(G${row}+1)/LOG10(1500),1)*20)+(IF(H${row}="Yes",5,0))+(IF(I${row}="Yes",5,0)),1))`;
}

function applyBaseFormatting(sheet, rows, cols) {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  const header = sheet.getRangeByIndexes(0, 0, 1, cols);
  header.format.fill.color = "#1f2937";
  header.format.font.color = "#ffffff";
  header.format.font.bold = true;
  header.format.wrapText = true;
  sheet.getRangeByIndexes(0, 0, rows, cols).format.font.name = "Aptos";
  sheet.getRangeByIndexes(0, 0, rows, cols).format.font.size = 10;
  sheet.getRangeByIndexes(1, 0, rows - 1, cols).format.wrapText = true;
  sheet.getRangeByIndexes(0, 0, rows, cols).format.verticalAlignment = "top";
}

const workbook = Workbook.create();
const rec = workbook.worksheets.add("Recommendations");
const method = workbook.worksheets.add("Sources & Method");

const headers = [
  "Rank",
  "Recommendation Score",
  "Coffee Shop",
  "Neighborhood / Area",
  "Address",
  "Google Rating",
  "Google Reviews",
  "Work-Friendly Source",
  "Specialty Coffee Callout",
  "Best For",
  "Caveat",
  "Primary Source",
  "Map URL",
];

const sorted = shops
  .map((shop) => {
    const ratingPart = shop.rating ? shop.rating * 20 * 0.7 : 0;
    const reviewPart = shop.reviews ? Math.min(Math.log10(shop.reviews + 1) / Math.log10(1500), 1) * 20 : 0;
    const workPart = shop.workFriendly === "Yes" ? 5 : 0;
    const specialtyPart = shop.specialty === "Yes" ? 5 : 0;
    return { ...shop, sortScore: ratingPart + reviewPart + workPart + specialtyPart };
  })
  .sort((a, b) => b.sortScore - a.sortScore);

const recRows = sorted.map((shop, index) => [
  index + 1,
  null,
  shop.shop,
  shop.neighborhood,
  shop.address,
  shop.rating,
  shop.reviews,
  shop.workFriendly,
  shop.specialty,
  shop.bestFor,
  shop.caveat,
  shop.source,
  shop.mapUrl,
]);

rec.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
rec.getRangeByIndexes(1, 0, recRows.length, headers.length).values = recRows;
for (let i = 0; i < recRows.length; i += 1) {
  rec.getCell(i + 1, 1).formulas = [[scoreFormula(i + 2)]];
}

applyBaseFormatting(rec, recRows.length + 1, headers.length);
rec.getRange("A:A").format.columnWidthPx = 48;
rec.getRange("B:B").format.columnWidthPx = 92;
rec.getRange("C:C").format.columnWidthPx = 185;
rec.getRange("D:D").format.columnWidthPx = 165;
rec.getRange("E:E").format.columnWidthPx = 245;
rec.getRange("F:G").format.columnWidthPx = 88;
rec.getRange("H:I").format.columnWidthPx = 105;
rec.getRange("J:K").format.columnWidthPx = 260;
rec.getRange("L:L").format.columnWidthPx = 165;
rec.getRange("M:M").format.columnWidthPx = 360;
rec.getRange("B:B").format.numberFormat = [["0.0"]];
rec.getRange("F:F").format.numberFormat = [["0.0"]];
rec.getRange("G:G").format.numberFormat = [["#,##0"]];
rec.getRange("A2:A17").format.horizontalAlignment = "center";
rec.getRange("B2:B17").format.horizontalAlignment = "center";
rec.getRange("F2:I17").format.horizontalAlignment = "center";
rec.getRange("A1:M1").format.rowHeightPx = 36;
rec.getRange("A2:M17").format.rowHeightPx = 58;

const methodRows = [
  ["Workbook purpose", "High-rated St. Louis city-area coffee shops that are practical for laptop work."],
  ["Current location assumption", "User said they are in St. Louis city, but did not give an exact address; this avoids exact distance ranking and keeps options within or near St. Louis city neighborhoods."],
  ["Ranking method", "Recommendation Score = 70% public rating strength, 20% review-volume confidence, 5% explicit work-friendly/Wi-Fi source, 5% specialty-coffee callout."],
  ["Rating source", "Workmode St. Louis cafe list, which states the cafes are laptop/Wi-Fi friendly and provides Google ratings/review counts."],
  ["Coffee-quality source", "My Coffee Explorer St. Louis guide, updated May 14, 2026, used for specialty-roaster/cafe callouts."],
  ["Direct official source", "Kaldi's Euclid page confirms Central West End address, hours, and free Wi-Fi; included as a practical honorable mention."],
  ["Important limitation", "Hours, table availability, outlet access, and Wi-Fi performance can change day to day; check Maps or the shop site before heading out."],
  ["Source URL", "https://workmode.co/st-louis/cafes"],
  ["Source URL", "https://mycoffeeexplorer.com/guides/st-louis"],
  ["Source URL", "https://kaldiscoffee.com/pages/the-euclid"],
];

method.getRange("A1:B1").values = [["Field", "Notes"]];
method.getRangeByIndexes(1, 0, methodRows.length, 2).values = methodRows;
applyBaseFormatting(method, methodRows.length + 1, 2);
method.getRange("A:A").format.columnWidthPx = 180;
method.getRange("B:B").format.columnWidthPx = 780;
method.getRange("A1:B1").format.rowHeightPx = 30;
method.getRangeByIndexes(1, 0, methodRows.length, 2).format.rowHeightPx = 45;

await fs.mkdir(outputDir, { recursive: true });

const inspect = await workbook.inspect({
  kind: "table",
  range: "Recommendations!A1:M17",
  include: "values,formulas",
  tableMaxRows: 18,
  tableMaxCols: 13,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "Recommendations",
  range: "A1:M17",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/recommendations-preview.png`, new Uint8Array(await preview.arrayBuffer()));

const methodPreview = await workbook.render({
  sheetName: "Sources & Method",
  range: "A1:B11",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/sources-method-preview.png`, new Uint8Array(await methodPreview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Saved ${outputPath}`);
