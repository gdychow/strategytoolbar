/**
 * One-off script (not a reusable pipeline step, unlike slice-catalog-
 * source.py): appends the 54 InsertChar-based character shortcuts from the
 * original VBA (Sub CharXxx() wrappers, each hardcoding a Unicode codepoint
 * and calling InsertChar(n)) onto the catalog-symbols.json that
 * slice-catalog-source.py already generated from Symbols.pptx's 7 real
 * slides. This list only ever gets authored once, from the VBA source
 * already extracted and read in full (Code/StrategyPPTv69a.pptm via
 * oletools.olevba) — codepoints below are copied directly from each Sub's
 * `n = <value>` line, not re-derived or guessed.
 *
 * CharReais is the one two-character exception: the VBA sub calls
 * InsertChar twice (82 = 'R', then 36 = '$') to build the "R$" Brazilian
 * Real symbol, since no single Unicode codepoint for it exists in the
 * BMP the rest of this list stays within. unicode_char is plain TEXT, so
 * storing the two-character string "R$" here needs no schema change.
 *
 * Usage: node scripts/append-symbol-chars.js db/seed/catalog-symbols.json
 */
const fs = require("fs");
const path = require("path");

// name, codepoint(s) — codepoint as a number, or an array for a
// multi-codepoint symbol (only CharReais).
const CHARS = [
  ["DoesNotEqual", 8800],
  ["AlmostEqual", 8776],
  ["Greaterthanorequal", 8805],
  ["Lessthanorequal", 8804],
  ["PlusMinus", 177],
  ["Division", 247],
  ["Multiplication", 215],
  ["EmptySet", 216],
  ["Infinity", 8734],
  ["Delta", 8710],
  ["Sigma", 8721],
  ["OneFourth", 188],
  ["OneHalf", 189],
  ["ThreeFourths", 190],
  ["Degree", 176],
  ["PerThousand", 8240],
  ["Dash", 8211],
  ["EmDash", 8212],
  ["VerticalLine", 8402],
  ["Dot", 8226],
  ["Copyright", 169],
  ["RegisteredTrademark", 174],
  ["Trademark", 8482],
  ["Euro", 8364],
  ["GBP", 163],
  ["Yen", 165],
  ["Cent", 162],
  ["ArrowTwoway", 8596],
  ["ArrowLeft", 8592],
  ["ArrowUp", 8593],
  ["ArrowRight", 8594],
  ["ArrowDown", 8595],
  ["Mu", 181],
  ["Paragraph", 182],
  ["Omega", 937],
  ["PI", 960],
  ["Ellipsis", 8230],
  ["LeftQuotation", 171],
  ["RightQuotation", 187],
  ["CommercialAt", 64],
  ["EncircledOne", 10112],
  ["EncircledTwo", 10113],
  ["EncircledThree", 10114],
  ["EncircledFour", 10115],
  ["EncircledFive", 10116],
  ["EncircledSix", 10117],
  ["EncircledSeven", 10118],
  ["EncircledEight", 10119],
  ["EncircledNine", 10120],
  ["EncircledTen", 10121],
  ["Cross", 10007],
  ["Checkmark", 10003],
  ["Yuan", 20803],
  ["Reais", [82, 36]],
];

// Manual overrides for names that don't split cleanly via mechanical
// PascalCase-to-spaced-words (either because the VBA name itself isn't
// properly cased, e.g. "Greaterthanorequal", or because a friendlier title
// than the literal name reads better, e.g. "PI" -> "Pi").
const TITLE_OVERRIDES = {
  Greaterthanorequal: "Greater Than or Equal",
  Lessthanorequal: "Less Than or Equal",
  GBP: "British Pound",
  PI: "Pi",
  Dash: "En Dash",
  CommercialAt: "At Sign (@)",
  Yuan: "Chinese Yuan",
  Reais: "Brazilian Real",
  EncircledOne: "Encircled 1",
  EncircledTwo: "Encircled 2",
  EncircledThree: "Encircled 3",
  EncircledFour: "Encircled 4",
  EncircledFive: "Encircled 5",
  EncircledSix: "Encircled 6",
  EncircledSeven: "Encircled 7",
  EncircledEight: "Encircled 8",
  EncircledNine: "Encircled 9",
  EncircledTen: "Encircled 10",
};

function titleFor(name) {
  if (TITLE_OVERRIDES[name]) return TITLE_OVERRIDES[name];
  // PascalCase -> spaced words: insert a space before each uppercase
  // letter that follows a lowercase letter.
  return name.replace(/([a-z])([A-Z])/g, "$1 $2");
}

function charFor(codepoints) {
  const points = Array.isArray(codepoints) ? codepoints : [codepoints];
  return String.fromCodePoint(...points);
}

function main() {
  const seedPath = process.argv[2];
  if (!seedPath) {
    console.error("Usage: node scripts/append-symbol-chars.js db/seed/catalog-symbols.json");
    process.exit(1);
  }

  const resolved = path.resolve(seedPath);
  const data = JSON.parse(fs.readFileSync(resolved, "utf8"));
  const startingSortOrder = data.items.length;

  CHARS.forEach(([name, codepoints], i) => {
    data.items.push({
      title: titleFor(name),
      insertMode: "unicode-char",
      unicodeChar: charFor(codepoints),
      thumbnail: null,
      sortOrder: startingSortOrder + i + 1,
    });
  });

  fs.writeFileSync(resolved, JSON.stringify(data, null, 2) + "\n");
  console.log(`Appended ${CHARS.length} character shortcut(s) to ${resolved} (${data.items.length} total).`);
}

main();
