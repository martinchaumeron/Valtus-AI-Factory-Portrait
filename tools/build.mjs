// Assemble le front à partir de custom-widget/{index.html,styles.css,app.js}
// vers ses deux cibles :
//
//   1. pages/Page1/widgets/DossierAnalyse.json — le widget « Custom » de la
//      page Appsmith, qui porte les trois volets HTML / CSS / JS.
//   2. dist/index.html — le même front en un seul fichier autonome, ouvrable
//      directement dans un navigateur pour relire ou tester sans Appsmith.
//
// Les trois fichiers de custom-widget/ sont la seule source de vérité : ces
// deux cibles sont générées, ne les éditez pas à la main.
//
//   node tools/build.mjs
//
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const read = (p) => fs.readFileSync(path.join(root, p), "utf8");

const html = read("custom-widget/index.html");
const css = read("custom-widget/styles.css");
const js = read("custom-widget/app.js");

// ---------------------------------------------------------------------------
// 1. Widget Appsmith
// ---------------------------------------------------------------------------

// La grille fixe d'Appsmith compte 64 colonnes et des lignes de 10 px. Le
// widget occupe toute la largeur et la hauteur initiale du canevas ; la page
// étant longue, elle défile à l'intérieur de l'iframe du widget.
const SNAP_COLUMNS = 64;
const WIDGET_ROWS = 124;
const WIDGET_NAME = "DossierAnalyse";
const WIDGET_ID = "dossieranalyse01";

// Les deux échanges du widget vers l'application passent par des événements, et
// les données voyagent dans l'objet de contexte de triggerEvent() — dont les
// clés sont directement lisibles dans la liaison de l'événement.
//
// Pourquoi pas le modèle du widget dans ce sens : `defaultModel` est lié au
// store pour le chemin inverse (application -> widget, voir plus bas). Or
// Appsmith réinitialise une propriété meta dès que sa propriété « par défaut »
// change ; tout ce que le widget écrirait avec updateModel() serait donc effacé
// au premier storeValue() de Portrait.lancer(). L'objet de contexte, lui,
// accompagne l'événement et ne peut pas être écrasé.
const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const ON_GENERATE = "{{Portrait.lancer(payload, jeton)}}";
const ON_DOWNLOAD_DOCX = `{{download(docxData, docxFilename, '${DOCX_MIME}')}}`;

// Chemin application -> widget : l'avancement et le résultat du traitement,
// déposés dans le store par l'objet JS Portrait, arrivent dans le modèle du
// widget par cette liaison.
const DEFAULT_MODEL =
  "{{ { etat: appsmith.store.etat, progression: appsmith.store.progression," +
  " resultat: appsmith.store.result, erreur: appsmith.store.erreur," +
  " jeton: appsmith.store.jeton } }}";

const srcDoc = { html, css, js };

const widget = {
  widgetName: WIDGET_NAME,
  type: "CUSTOM_WIDGET",
  widgetId: WIDGET_ID,
  parentId: "0",
  version: 1,
  renderMode: "CANVAS",
  isLoading: false,
  isVisible: true,
  animateLoading: false,
  // Le volet compilé et le volet source sont identiques : le JS est écrit en
  // JavaScript standard (pas de JSX), il n'y a donc rien à transpiler.
  srcDoc,
  uncompiledSrcDoc: srcDoc,
  defaultModel: DEFAULT_MODEL,
  events: ["onGenerate", "onDownloadDocx"],
  onGenerate: ON_GENERATE,
  onDownloadDocx: ON_DOWNLOAD_DOCX,
  theme: "{{appsmith.theme}}",
  dynamicBindingPathList: [{ key: "defaultModel" }, { key: "theme" }],
  dynamicTriggerPathList: [{ key: "onDownloadDocx" }, { key: "onGenerate" }],
  dynamicHeight: "FIXED",
  minDynamicHeight: 4,
  maxDynamicHeight: 9000,
  backgroundColor: "#FFFFFF",
  borderColor: "#E0DEDE",
  borderWidth: "0",
  borderRadius: "0px",
  boxShadow: "none",
  rows: WIDGET_ROWS,
  columns: SNAP_COLUMNS,
  topRow: 0,
  bottomRow: WIDGET_ROWS,
  leftColumn: 0,
  rightColumn: SNAP_COLUMNS,
  mobileTopRow: 0,
  mobileBottomRow: WIDGET_ROWS,
  mobileLeftColumn: 0,
  mobileRightColumn: SNAP_COLUMNS,
  parentColumnSpace: 1,
  parentRowSpace: 10,
};

// Appsmith écrit ses fichiers exportés avec les clés triées : on fait pareil,
// pour qu'une synchronisation faite depuis l'éditeur ne produise pas un
// remaniement complet du fichier.
const sortKeys = (value) => {
  if (Array.isArray(value)) return value.map(sortKeys);
  if (value && typeof value === "object") {
    return Object.keys(value)
      .sort()
      .reduce((acc, key) => {
        acc[key] = sortKeys(value[key]);
        return acc;
      }, {});
  }
  return value;
};

fs.mkdirSync(path.join(root, "pages/Page1/widgets"), { recursive: true });
fs.writeFileSync(
  path.join(root, "pages/Page1/widgets/" + WIDGET_NAME + ".json"),
  JSON.stringify(sortKeys(widget), null, 2) + "\n",
);

// ---------------------------------------------------------------------------
// 2. Fichier autonome
// ---------------------------------------------------------------------------
// Même ordre d'assemblage que le widget « Custom » d'Appsmith (le balisage,
// puis le module JS, puis la feuille de style) pour que les deux cibles se
// comportent à l'identique. Hors Appsmith, le front bascule de lui-même en
// mode démonstration : aucun appel n'est émis.
const standalone = `<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Dossier candidat — Valtus</title>
</head>
<body>
${html}
<script type="module">
${js}
</script>
<style>
${css}
</style>
</body>
</html>
`;

fs.mkdirSync(path.join(root, "dist"), { recursive: true });
fs.writeFileSync(path.join(root, "dist/index.html"), standalone);

const kb = (s) => (Buffer.byteLength(s, "utf8") / 1024).toFixed(1) + " Kio";
console.log("pages/Page1/widgets/" + WIDGET_NAME + ".json  (html " + kb(html) + ", css " + kb(css) + ", js " + kb(js) + ")");
console.log("dist/index.html                          " + kb(standalone));
