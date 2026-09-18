// Assemble le front à partir de custom-widget/{index.html,styles.css,app.js}
// vers ses deux cibles :
//
//   1. pages/Page1/Page1.json — la page Appsmith, dont l'unique widget est un
//      widget « Custom » qui porte les trois volets HTML / CSS / JS.
//   2. dist/index.html        — le même front en un seul fichier autonome,
//      ouvrable directement dans un navigateur pour relire ou tester sans
//      Appsmith.
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
// 1. Page Appsmith
// ---------------------------------------------------------------------------

// La grille fixe d'Appsmith compte 64 colonnes et des lignes de 10 px. Le
// widget occupe toute la largeur et la hauteur initiale du canevas ; la page
// étant longue, elle défile à l'intérieur de l'iframe du widget.
const SNAP_COLUMNS = 64;
const WIDGET_ROWS = 124;
const WIDGET_NAME = "DossierCandidat";
const WIDGET_ID = "dossiercandidat1";

// Le widget publie le .docx dans son modèle puis déclenche cet événement :
// l'iframe du widget « Custom » est sandboxée sans allow-downloads, elle ne
// peut donc pas déclencher le téléchargement elle-même (voir saveViaAppsmith
// dans custom-widget/app.js).
const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const ON_DOWNLOAD_DOCX =
  `{{download(${WIDGET_NAME}.model.docxData, ${WIDGET_NAME}.model.docxFilename, '${DOCX_MIME}')}}`;

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
  defaultModel: "{}",
  events: ["onDownloadDocx"],
  onDownloadDocx: ON_DOWNLOAD_DOCX,
  theme: "{{appsmith.theme}}",
  dynamicBindingPathList: [{ key: "theme" }],
  dynamicTriggerPathList: [{ key: "onDownloadDocx" }],
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

const page = {
  gitSyncId: "6aad2268898437a504d32b67_eeab6b01-094d-40e8-bfad-897837ae103e",
  unpublishedPage: {
    layouts: [
      {
        dsl: {
          backgroundColor: "none",
          bottomRow: 5000,
          canExtend: true,
          containerStyle: "none",
          detachFromLayout: true,
          dynamicBindingPathList: [],
          dynamicTriggerPathList: [],
          leftColumn: 0,
          minHeight: 1292,
          parentColumnSpace: 1,
          parentRowSpace: 1,
          rightColumn: 4896,
          snapColumns: SNAP_COLUMNS,
          snapRows: 124,
          topRow: 0,
          type: "CANVAS_WIDGET",
          version: 94,
          widgetId: "0",
          widgetName: "MainContainer",
          children: [widget],
        },
      },
    ],
    name: "Page1",
    slug: "page1",
  },
};

fs.writeFileSync(
  path.join(root, "pages/Page1/Page1.json"),
  JSON.stringify(page, null, 2) + "\n",
);

// ---------------------------------------------------------------------------
// 2. Fichier autonome
// ---------------------------------------------------------------------------
// Même ordre d'assemblage que le widget « Custom » d'Appsmith (le balisage,
// puis le module JS, puis la feuille de style) pour que les deux cibles se
// comportent à l'identique.
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
console.log(`pages/Page1/Page1.json  (html ${kb(html)}, css ${kb(css)}, js ${kb(js)})`);
console.log(`dist/index.html         ${kb(standalone)}`);
