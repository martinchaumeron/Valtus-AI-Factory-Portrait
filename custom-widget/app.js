/*
 * Générateur de dossier candidat — Valtus
 *
 * Un seul traitement : toutes les sources (retranscription, CV, analyse
 * externe, informations complémentaires, brief client) sont saisies sur le même
 * écran et envoyées ensemble au backend, qui se charge de l'intégralité de
 * l'analyse et de la rédaction. Le front ne construit aucun prompt, ne découpe
 * rien et n'enchaîne aucune étape : il collecte, il affiche, il exporte.
 *
 * Le livrable renvoyé tient en trois sections — compétences clés, mail client,
 * pitch — restituées dans un aperçu fidèle au document Word exporté.
 */
(function(){
  "use strict";

  // ===============================================================
  // Sources, champs de mission, sections du livrable
  // ===============================================================
  // Ces trois tables décrivent tout ce que l'écran manipule. Ajouter une source
  // ou un champ se fait ici et dans le balisage, nulle part ailleurs.

  var SOURCES = [
    { key: "transcription", id: "srcTranscription", label: "la retranscription d'entretien", required: true },
    { key: "cv", id: "srcCv", label: "le CV du candidat", required: true },
    { key: "analyseExterne", id: "srcAnalyse" },
    { key: "informationsComplementaires", id: "srcComplementaires" },
    { key: "briefClient", id: "srcBrief" }
  ];

  var MISSION_FIELDS = [
    { key: "candidat", id: "fieldCandidat" },
    { key: "poste", id: "fieldPoste" },
    { key: "reference", id: "fieldReference" },
    { key: "societe", id: "fieldSociete" },
    { key: "contact", id: "fieldContact" },
    { key: "lieu", id: "fieldLieu" },
    { key: "tauxJournalier", id: "fieldTaux" },
    { key: "signataire", id: "fieldSignataire" }
  ];

  // L'ordre de cette table est l'ordre du document, à l'écran comme dans le
  // .docx exporté.
  var SECTIONS = [
    { key: "competences", title: "Compétences clés" },
    { key: "mail", title: "Mail client (prêt à copier)", callout: true },
    { key: "pitch", title: "Pitch candidat" }
  ];

  var state = {
    result: null,        // { competences, mail, pitch } renvoyé par le backend
    generating: false,
    hasGenerated: false
  };

  function byId(id){ return document.getElementById(id); }

  var els = {
    statusChip: byId("statusChip"),
    statusLabel: byId("statusLabel"),
    demoBanner: byId("demoBanner"),
    toast: byId("toast"),
    prefs: byId("prefs"),
    generateBtn: byId("generateBtn"),
    generateLabel: byId("generateLabel"),
    generateHint: byId("generateHint"),
    downloadBtn: byId("downloadBtn"),
    copyBtn: byId("copyBtn"),
    sendBtn: byId("sendBtn"),
    aiWarning: byId("aiWarning"),
    docPage: byId("docPage")
  };

  SOURCES.forEach(function(source){
    source.field = byId(source.id);
    source.counter = byId(source.id + "Count");
  });
  MISSION_FIELDS.forEach(function(field){ field.input = byId(field.id); });

  function sourceText(key){
    for (var i = 0; i < SOURCES.length; i++){
      if (SOURCES[i].key === key) return SOURCES[i].field.value.trim();
    }
    return "";
  }

  function missionText(key){
    for (var i = 0; i < MISSION_FIELDS.length; i++){
      if (MISSION_FIELDS[i].key === key) return MISSION_FIELDS[i].input.value.trim();
    }
    return "";
  }

  // ===============================================================
  // Préférences de rédaction
  // ===============================================================
  // Note libre que le consultant saisit une fois (style, formule de politesse,
  // signature…) et qui accompagne chaque génération. Conservée dans
  // localStorage : c'est une commodité propre à ce navigateur, jamais partagée.
  // L'accès est enveloppé dans un try/catch car il peut lever une exception
  // (fenêtre privée, données de site bloquées, iframe sandboxée) — l'outil doit
  // fonctionner normalement dans tous les cas.
  var PREFS_KEY = "valtus_preferences_redaction_v1";
  try {
    var savedPrefs = window.localStorage.getItem(PREFS_KEY);
    if (savedPrefs) els.prefs.value = savedPrefs;
  } catch (e) { /* stockage indisponible — le champ démarre simplement vide */ }
  els.prefs.addEventListener("input", function(){
    try { window.localStorage.setItem(PREFS_KEY, els.prefs.value); }
    catch (e) { /* rien à faire — la préférence ne survivra pas au rechargement */ }
  });

  // ===============================================================
  // Utilitaires d'affichage
  // ===============================================================

  function esc(s){
    return String(s == null ? "" : s).replace(/[&<>"']/g, function(c){
      return ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"})[c];
    });
  }

  // The LLM is instructed not to use markdown emphasis, but sometimes does
  // anyway (**bold**). Rendering that literally (raw asterisks) looks
  // broken, so both the HTML preview and the DOCX export render **…** as
  // actual bold instead — used by blocksToHtml (HTML) and runBold (DOCX).
  function splitBoldSegments(text){
    var s = String(text == null ? "" : text);
    var parts = [];
    var re = /\*\*(.+?)\*\*/g;
    var last = 0, m;
    while ((m = re.exec(s))){
      if (m.index > last) parts.push({ text: s.slice(last, m.index), bold: false });
      parts.push({ text: m[1], bold: true });
      last = re.lastIndex;
    }
    if (last < s.length) parts.push({ text: s.slice(last), bold: false });
    return parts.length ? parts : [{ text: s, bold: false }];
  }

  function escBold(text){
    return splitBoldSegments(text).map(function(seg){
      var t = esc(seg.text);
      return seg.bold ? "<strong>" + t + "</strong>" : t;
    }).join("");
  }

  function setStatus(stateName, label){
    els.statusChip.setAttribute("data-state", stateName);
    els.statusLabel.textContent = label;
  }

  function showToast(message, tone){
    els.toast.textContent = message;
    els.toast.setAttribute("data-tone", tone || "info");
    els.toast.style.display = "block";
    clearTimeout(showToast._t);
    showToast._t = setTimeout(function(){ els.toast.style.display = "none"; }, 5200);
  }

  // ===============================================================
  // Saisie des sources
  // ===============================================================

  function updateCounts(){
    SOURCES.forEach(function(source){
      var length = source.field.value.length;
      source.counter.textContent = length.toLocaleString("fr-FR") + " caractère" + (length === 1 ? "" : "s");
    });
    refreshGenerateState();
  }

  SOURCES.forEach(function(source){
    source.field.addEventListener("input", updateCounts);
  });
  MISSION_FIELDS.forEach(function(field){
    field.input.addEventListener("input", renderPreview);
  });

  document.querySelectorAll(".btn-file[data-target]").forEach(function(btn){
    btn.addEventListener("click", function(){
      document.querySelector('input[data-file-for="' + btn.getAttribute("data-target") + '"]').click();
    });
  });

  document.querySelectorAll('input[type="file"]').forEach(function(input){
    input.addEventListener("change", function(){
      var file = input.files && input.files[0];
      if (!file) return;
      var target = byId(input.getAttribute("data-file-for"));
      var isDocx = /\.docx$/i.test(file.name);
      var isPdf = /\.pdf$/i.test(file.name);
      var kind = isDocx ? " (.docx)" : isPdf ? " (.pdf)" : "";
      var task = isDocx ? readDocxAsText(file) : isPdf ? readPdfAsText(file) : readPlainText(file);
      task.then(function(text){
        if (!text.trim() && isPdf){
          showToast("Aucun texte trouvé dans ce PDF (page scannée sans OCR ?). Collez le texte manuellement.", "error");
          return;
        }
        target.value = text;
        updateCounts();
      }).catch(function(){
        showToast("Impossible de lire ce fichier" + kind + " — collez son contenu manuellement.", "error");
      }).finally(function(){ input.value = ""; });
    });
  });

  // Extract plain text from a .docx file's word/document.xml by regex —
  // avoids any DOMParser / XML-namespace inconsistency across browsers.
  //
  // Single pass over the whole XML, in document order: capture each <w:t>
  // run's text, and insert separators at the right boundaries (</w:tc> = end
  // of a table cell -> " | ", </w:tr> = end of a table row -> line break,
  // </w:p> = end of a paragraph -> line break, <w:br/>/<w:tab/> -> line
  // break / tab). The previous version split the whole XML on every </w:p>
  // before extracting <w:t> from each chunk — but a table cell's own
  // <w:p>…</w:p> paragraphs sit NESTED inside <w:tbl>/<w:tr>/<w:tc>, so that
  // split cut straight through the table markup instead of around it.
  // Separately (and this was the actual bug users hit), the old <w:t[^>]*>
  // pattern had no boundary after "w:t", so it also matched unrelated tags
  // that happen to share the same prefix — <w:tbl>, <w:tblPr>, <w:tblGrid>,
  // <w:tblLook>, <w:tc>, <w:tcPr>, <w:tr> — as if each were a <w:t ...>
  // opening tag, then greedily captured everything up to the next real
  // </w:t> as "text". On any document containing a table, this leaked large
  // chunks of raw table XML (<w:tblPr>…</w:tblPr>, tcW/gridCol attributes,
  // etc.) straight into the imported text — costing a huge number of wasted
  // tokens and confusing the LLM. The (?=[ /nl>]) lookahead below requires
  // "w:t" to be followed by a space, "/" or ">" — never another letter —
  // so it can only ever match the real <w:t> tag.
  function docxXmlToText(xml){
    var out = [];
    var re = /<w:t(?=[ \/>])[^>]*\/>|<w:t(?=[ \/>])[^>]*>([\s\S]*?)<\/w:t>|<\/w:tc>|<\/w:tr>|<\/w:p>|<w:br\b[^>]*\/?>|<w:tab\b[^>]*\/?>/g;
    var m;
    while ((m = re.exec(xml))){
      if (m[1] !== undefined) out.push(m[1]);
      else if (m[0] === "</w:tc>") out.push(" | ");
      else if (m[0] === "</w:tr>") out.push("\n");
      else if (m[0] === "</w:p>") out.push("\n");
      else if (m[0].indexOf("<w:br") === 0) out.push("\n");
      else if (m[0].indexOf("<w:tab") === 0) out.push("\t");
    }
    return out.join("")
      .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'").replace(/&amp;/g, "&")
      .replace(/\n+ \| /g, " | ")
      .replace(/ \| \n+/g, "\n")
      .replace(/\n{3,}/g, "\n\n").trim();
  }

  function readDocxAsText(file){
    if (typeof JSZip === "undefined") return Promise.reject(new Error("JSZip indisponible"));
    return file.arrayBuffer()
      .then(function(buf){ return JSZip.loadAsync(buf); })
      .then(function(zip){
        var entry = zip.file("word/document.xml");
        if (!entry) throw new Error("document.xml introuvable");
        return entry.async("string");
      })
      .then(docxXmlToText);
  }

  function readPlainText(file){
    return new Promise(function(resolve, reject){
      var reader = new FileReader();
      reader.onload = function(){ resolve(String(reader.result || "")); };
      reader.onerror = function(){ reject(new Error("lecture impossible")); };
      reader.readAsText(file);
    });
  }

  function withTimeout(promise, ms, message){
    return new Promise(function(resolve, reject){
      var timer = setTimeout(function(){ reject(new Error(message)); }, ms);
      promise.then(function(v){ clearTimeout(timer); resolve(v); }, function(e){ clearTimeout(timer); reject(e); });
    });
  }

  function readPdfAsText(file){
    if (typeof pdfjsLib === "undefined") return Promise.reject(new Error("pdf.js indisponible"));
    var task = file.arrayBuffer().then(function(buf){
      return pdfjsLib.getDocument({ data: buf }).promise;
    }).then(function(pdf){
      var pageNums = [];
      for (var p = 1; p <= pdf.numPages; p++) pageNums.push(p);
      return pageNums.reduce(function(chain, num){
        return chain.then(function(acc){
          return pdf.getPage(num).then(function(page){ return page.getTextContent(); }).then(function(content){
            var text = content.items.map(function(it){ return it.str; }).join(" ");
            return acc.concat(text);
          });
        });
      }, Promise.resolve([]));
    }).then(function(pages){ return pages.join("\n\n").trim(); });
    return withTimeout(task, 30000, "extraction PDF trop longue");
  }

  // ===============================================================
  // Du texte renvoyé au modèle de blocs
  // ===============================================================
  // Un seul modèle de blocs sert à la fois à l'aperçu à l'écran et à l'export
  // Word : ce qui est affiché est exactement ce qui sera dans le document.

  // ---------------------------------------------------------------
  // Parsing LLM output + raw text into a shared block model, used
  // identically for the on-screen preview and the DOCX export.
  // ---------------------------------------------------------------
  function isTableSeparatorLine(l){
    return /^\|?[\s:|-]+\|?$/.test(l) && l.indexOf("-") !== -1;
  }

  function parsePipeRow(l){
    var row = l.trim();
    if (row.charAt(0) === "|") row = row.slice(1);
    if (row.charAt(row.length - 1) === "|") row = row.slice(0, -1);
    return row.split("|").map(function(c){ return c.trim(); });
  }

  // Shared parser: turns LLM markdown-ish output (## / ### headings, "- "
  // bullets, "1. " numbered lines, "|…|" tables, blank-line paragraphs)
  // into the same block model used for both the on-screen preview and the
  // DOCX export.
  function parseBrief(text){
    var blocks = [];
    var buf = [];
    function flush(){
      if (buf.length){
        blocks.push({ style: "p", text: buf.join(" ").replace(/\s+/g, " ").trim() });
        buf = [];
      }
    }
    var lines = String(text || "").replace(/\r\n/g, "\n").split("\n");
    var i = 0;
    while (i < lines.length){
      var l = lines[i].trim();
      if (l.charAt(0) === "|"){
        flush();
        var tableLines = [];
        while (i < lines.length && lines[i].trim().charAt(0) === "|"){
          tableLines.push(lines[i].trim());
          i++;
        }
        var rows = tableLines.filter(function(tl){ return !isTableSeparatorLine(tl); }).map(parsePipeRow);
        if (rows.length){
          blocks.push({ style: "table", header: rows[0], rows: rows.slice(1) });
        }
        continue;
      }
      if (l === ""){ flush(); i++; continue; }
      var m3 = /^###\s+(.*)$/.exec(l);
      var m2 = /^##\s+(.*)$/.exec(l);
      var mb = /^[-•]\s+(.*)$/.exec(l);
      var mn = /^\d+[.)]\s+(.*)$/.exec(l);
      if (m3){ flush(); blocks.push({ style: "h3", text: m3[1] }); }
      else if (m2){ flush(); blocks.push({ style: "h2", text: m2[1] }); }
      else if (mb){ flush(); blocks.push({ style: "bullet", text: mb[1] }); }
      else if (mn){ flush(); blocks.push({ style: "p", text: l }); }
      else { buf.push(l); }
      i++;
    }
    flush();
    return blocks;
  }

  function blocksToHtml(blocks){
    return blocks.map(function(b){
      switch (b.style){
        case "title": return '<h1 class="doc-title">' + esc(b.text) + "</h1>";
        case "meta": return '<p class="doc-meta">' + esc(b.text) + "</p>";
        case "h1": return '<h2 class="doc-h1">' + escBold(b.text) + "</h2>";
        case "h2": return '<h3 class="doc-h2">' + escBold(b.text) + "</h3>";
        case "h3": return '<h4 class="doc-h3">' + escBold(b.text) + "</h4>";
        case "bullet": return '<p class="doc-bullet">•&nbsp;&nbsp;' + escBold(b.text) + "</p>";
        case "placeholder": return '<p class="doc-placeholder">' + esc(b.text) + "</p>";
        case "warning": return '<p class="doc-warning">⚠ ' + esc(b.text) + "</p>";
        case "spacer": return '<div class="doc-spacer"></div>';
        case "table":
          return '<div class="doc-table-wrap"><table class="doc-table"><thead><tr>' +
            b.header.map(function(c){ return "<th>" + escBold(c) + "</th>"; }).join("") +
            "</tr></thead><tbody>" +
            b.rows.map(function(r){ return "<tr>" + r.map(function(c){ return "<td>" + escBold(c) + "</td>"; }).join("") + "</tr>"; }).join("") +
            "</tbody></table></div>";
        default: return b.text ? '<p class="doc-p">' + escBold(b.text) + "</p>" : '<div class="doc-spacer"></div>';
      }
    }).join("");
  }

  // Le mail est présenté dans un encadré : la section correspondante ouvre un
  // bloc que la section suivante referme.
  function blocksToHtmlWithCallouts(blocks){
    var html = "";
    var inCallout = false;
    blocks.forEach(function(b){
      if (b.calloutStart){
        if (inCallout) html += "</div>";
        html += '<div class="doc-callout">';
        inCallout = true;
        html += blocksToHtml([b]);
        return;
      }
      if (inCallout && (b.style === "h1" || b.style === "h2")){
        html += "</div>";
        inCallout = false;
      }
      html += blocksToHtml([b]);
    });
    if (inCallout) html += "</div>";
    return html;
  }

  function documentBlocks(){
    var blocks = [];
    blocks.push({ style: "title", text: "Dossier de présentation candidat" });

    var meta = [];
    var candidat = missionText("candidat");
    if (candidat) meta.push("Candidat : " + candidat);
    var poste = missionText("poste");
    if (poste) meta.push("Mission : " + poste);
    var societe = missionText("societe");
    if (societe) meta.push("Client : " + societe);
    var reference = missionText("reference");
    if (reference) meta.push("Référence : " + reference);
    meta.push("Préparé le : " + new Date().toLocaleDateString("fr-FR", { year: "numeric", month: "long", day: "numeric" }));
    blocks.push({ style: "meta", text: meta.join("   ·   ") });

    if (!state.result){
      blocks.push({ style: "placeholder", text: "Le dossier n'a pas encore été généré. Renseignez les sources ci-contre, puis lancez la génération." });
      return blocks;
    }

    // Filet de sécurité : si le backend renvoie un texte brut au lieu des trois
    // sections attendues, on l'affiche entier plutôt que rien.
    if (state.result.raw){
      return blocks.concat(parseBrief(state.result.raw));
    }

    SECTIONS.forEach(function(section){
      var text = String(state.result[section.key] || "").trim();
      var heading = { style: "h1", text: section.title };
      if (section.callout) heading.calloutStart = true;
      blocks.push(heading);
      blocks = blocks.concat(text
        ? parseBrief(text)
        : [{ style: "placeholder", text: "Section non renvoyée par le backend." }]);
    });
    return blocks;
  }

  function renderPreview(){
    els.docPage.innerHTML = blocksToHtmlWithCallouts(documentBlocks());
    refreshActionState();
  }

  // ===============================================================
  // État des actions
  // ===============================================================

  function missingRequired(){
    return SOURCES.filter(function(source){
      return source.required && !source.field.value.trim();
    });
  }

  function refreshGenerateState(){
    var missing = missingRequired();
    els.generateBtn.disabled = state.generating || missing.length > 0;
    if (state.generating){
      els.generateHint.textContent = "Analyse en cours côté backend…";
      els.generateHint.classList.remove("warn");
    } else if (missing.length){
      els.generateHint.textContent = "Source requise : " + missing.map(function(s){ return s.label; }).join(" et ") + ".";
      els.generateHint.classList.remove("warn");
    } else {
      els.generateHint.textContent = "Toutes les sources partent en une seule fois ; le backend se charge de l'analyse.";
      els.generateHint.classList.remove("warn");
    }
  }

  function mailText(){
    if (!state.result) return "";
    return String(state.result.mail || "").trim();
  }

  function refreshActionState(){
    var ready = state.hasGenerated && !state.generating;
    var mail = mailText();
    els.downloadBtn.disabled = !ready;
    els.copyBtn.disabled = !ready || !mail;
    els.sendBtn.disabled = !ready || !mail;
    // Visible dès qu'un dossier existe, même pendant une régénération : le
    // rappel doit précéder toute copie, tout envoi et tout export.
    els.aiWarning.hidden = !state.hasGenerated;
  }

  // ===============================================================
  // Backend — point de branchement unique
  // ===============================================================
  // AUCUN APPEL N'EST ÉMIS AUJOURD'HUI. Tant qu'aucun backend n'est fourni, la
  // génération est simulée localement : toute la chaîne (aperçu, export Word,
  // copie du mail) reste utilisable, et le texte produit s'annonce lui-même
  // comme un exemple.
  //
  // Pour brancher le backend, définir window.VALTUS_BACKEND avant ce script :
  //
  //   window.VALTUS_BACKEND = {
  //     generate: function (payload, options) {
  //       // payload = {
  //       //   mission: { candidat, poste, reference, societe, contact,
  //       //              lieu, tauxJournalier, signataire },
  //       //   sources: { transcription, cv, analyseExterne,
  //       //              informationsComplementaires, briefClient },
  //       //   preferencesMail: "…"
  //       // }
  //       // options.onProgress("Analyse du CV…") met à jour le libellé du
  //       // bouton pendant le traitement ; son appel est facultatif.
  //       // Résoudre avec { competences, mail, pitch } (du markdown simple :
  //       // titres "##", puces "- ", tableaux "|…|").
  //       // Rejeter avec { code, message } — les codes reconnus sont les clés
  //       // de ERROR_COPY ci-dessous.
  //       return Promise.resolve({ competences: "…", mail: "…", pitch: "…" });
  //     }
  //   };
  //
  // Voir custom-widget/README.md pour un exemple de branchement sur une query
  // Appsmith, qui garde la clé d'API côté serveur.

  var ERROR_COPY = {
    unauthorized: "Accès refusé par le backend. Vérifiez vos droits puis réessayez.",
    session_expired: "Votre session a expiré — reconnectez-vous puis réessayez.",
    rate_limited: "Trop de demandes en cours. Réessayez dans un instant.",
    payload_too_large: "Les sources sont trop volumineuses pour être traitées. Raccourcissez-les puis réessayez.",
    refused: "La demande n'a pas pu être traitée. Vérifiez le contenu des sources et réessayez.",
    empty_result: "Aucun contenu n'a été produit. Réessayez.",
    timeout: "Le traitement a dépassé le délai autorisé. Réessayez.",
    upstream_error: "Problème temporaire côté service. Réessayez dans quelques instants.",
    internal_error: "Erreur technique côté backend. Réessayez."
  };

  var STUB_DELAY_MS = 900;
  var STUB_NOTICE = "⚠ CONTENU DE DÉMONSTRATION — aucun backend n'est branché. Ce texte ne provient pas des sources saisies et ne doit en aucun cas être transmis à un client.";

  function backendEngine(){
    var engine = window.VALTUS_BACKEND;
    return (engine && typeof engine.generate === "function") ? engine : null;
  }

  function stubGenerate(payload, options){
    return new Promise(function(resolve){
      if (options && typeof options.onProgress === "function") options.onProgress("Analyse des sources…");
      setTimeout(function(){
        resolve({
          competences: "- " + STUB_NOTICE + "\n" +
            "- Compétence clé d'exemple n° 1 — remplacée par la sortie du backend une fois celui-ci branché.\n" +
            "- Compétence clé d'exemple n° 2 — remplacée par la sortie du backend une fois celui-ci branché.",
          mail: "Objet : [DÉMONSTRATION] Présentation de candidature\n\n" +
            "Bonjour,\n\n" + STUB_NOTICE + "\n\n" +
            "Le corps du mail client s'affichera ici une fois le backend branché.\n\n" +
            "Bien à vous,",
          pitch: STUB_NOTICE + "\n\n" +
            "Le pitch du candidat — quelques phrases situant son parcours et ce qu'il apporte à la mission — s'affichera ici une fois le backend branché."
        });
      }, STUB_DELAY_MS);
    });
  }

  function callBackend(payload, options){
    var engine = backendEngine();
    return engine ? engine.generate(payload, options || {}) : stubGenerate(payload, options);
  }

  function buildPayload(){
    var mission = {};
    MISSION_FIELDS.forEach(function(field){ mission[field.key] = field.input.value.trim(); });
    var sources = {};
    SOURCES.forEach(function(source){ sources[source.key] = source.field.value.trim(); });
    return {
      mission: mission,
      sources: sources,
      preferencesMail: els.prefs.value.trim()
    };
  }

  // Le livrable attendu tient en trois sections. Si le backend renvoie autre
  // chose (une chaîne, un { text }), on le conserve tel quel sous « raw » pour
  // l'afficher intégralement plutôt que de perdre sa réponse.
  function normalizeResult(result){
    if (typeof result === "string") return { raw: result };
    if (!result || typeof result !== "object") return { raw: String(result == null ? "" : result) };
    var hasSection = SECTIONS.some(function(section){
      return typeof result[section.key] === "string" && result[section.key].trim();
    });
    if (hasSection) return result;
    if (typeof result.text === "string") return { raw: result.text };
    return { raw: JSON.stringify(result, null, 2) };
  }

  // ===============================================================
  // Génération
  // ===============================================================

  function finishGenerate(label){
    state.generating = false;
    els.generateBtn.removeAttribute("data-busy");
    els.generateLabel.textContent = label;
    refreshGenerateState();
    refreshActionState();
  }

  function runGenerate(){
    if (state.generating || missingRequired().length) return;

    state.generating = true;
    els.generateBtn.setAttribute("data-busy", "true");
    els.generateBtn.disabled = true;
    els.generateLabel.textContent = state.hasGenerated ? "Régénération…" : "Génération en cours…";
    setStatus("busy", "Génération en cours");
    refreshActionState();

    callBackend(buildPayload(), {
      onProgress: function(label){ if (label) els.generateLabel.textContent = label; }
    }).then(function(result){
      state.result = normalizeResult(result);
      state.hasGenerated = true;
      finishGenerate("Régénérer le dossier");
      setStatus("done", "Dossier généré");
      renderPreview();
    }).catch(function(err){
      finishGenerate(state.hasGenerated ? "Régénérer le dossier" : "Générer le dossier");
      setStatus("error", "Échec de la génération");
      showToast(
        (err && ERROR_COPY[err.code]) ||
        (err && err.message) ||
        "Une erreur est survenue pendant la génération.",
        "error"
      );
    });
  }

  els.generateBtn.addEventListener("click", runGenerate);

  // ===============================================================
  // Export Word
  // ===============================================================

  // ---------------------------------------------------------------
  // DOCX export
  // ---------------------------------------------------------------
  var CT_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>';

  var RELS_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  var DOC_RELS_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';

  var STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="fr-FR"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="288" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:before="0" w:after="240"/><w:keepNext/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:b/><w:color w:val="222F42"/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:before="360" w:after="160"/><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="9B2841"/></w:pBdr></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:b/><w:caps/><w:color w:val="222F42"/><w:sz w:val="27"/><w:szCs w:val="27"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:before="240" w:after="120"/><w:keepNext/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:b/><w:color w:val="9B2841"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:before="200" w:after="100"/><w:keepNext/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:i/><w:color w:val="222F42"/><w:sz w:val="22"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:ind w:left="284" w:hanging="284"/></w:pPr></w:style>' +
    '</w:styles>';

  function xmlEsc(s){
    return String(s == null ? "" : s).replace(/[&<>"']/g, function(c){
      return ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&apos;"})[c];
    });
  }

  function run(text, rPr){
    return "<w:r>" + (rPr || "") + '<w:t xml:space="preserve">' + xmlEsc(text) + "</w:t></w:r>";
  }

  // Same **bold** handling as escBold(), but emitting Word runs: bold spans
  // get their own <w:r> with <w:b/> added to the paragraph's base rPr.
  function runBold(text, rPr){
    var segs = splitBoldSegments(text);
    if (segs.length === 1 && !segs[0].bold) return run(segs[0].text, rPr);
    return segs.map(function(seg){
      var segRPr = rPr || "";
      if (seg.bold){
        segRPr = segRPr ? segRPr.replace("<w:rPr>", "<w:rPr><w:b/>") : "<w:rPr><w:b/></w:rPr>";
      }
      return run(seg.text, segRPr);
    }).join("");
  }

  function paragraph(pStyle, runsXml){
    var pPr = pStyle ? "<w:pPr><w:pStyle w:val=\"" + pStyle + "\"/></w:pPr>" : "";
    return "<w:p>" + pPr + runsXml + "</w:p>";
  }

  var MUTED_ITALIC_RPR = '<w:rPr><w:i/><w:color w:val="6B7280"/></w:rPr>';

  var META_RPR = '<w:rPr><w:i/><w:color w:val="6B7280"/><w:sz w:val="19"/></w:rPr>';

  function blockToParagraphXml(b){
    switch (b.style){
      case "title": return paragraph("Title", run(b.text));
      case "meta": return paragraph(null, run(b.text, META_RPR));
      case "h1": return paragraph("Heading1", runBold(b.text));
      case "h2": return paragraph("Heading2", runBold(b.text));
      case "h3": return paragraph("Heading3", runBold(b.text));
      case "bullet": return paragraph("ListParagraph", "<w:r><w:t>•</w:t></w:r><w:r><w:tab/></w:r>" + runBold(b.text));
      case "placeholder": return paragraph(null, run(b.text, MUTED_ITALIC_RPR));
      case "spacer": return "<w:p/>";
      default: return b.text ? paragraph(null, runBold(b.text)) : "<w:p/>";
    }
  }

  var TABLE_BORDER_EDGES = ["top", "left", "bottom", "right", "insideH", "insideV"];

  function blockToTableXml(block){
    var ncols = block.header.length;
    block.rows.forEach(function(r){ ncols = Math.max(ncols, r.length); });
    ncols = Math.max(ncols, 1);
    var colWidth = Math.floor(9026 / ncols);
    var gridCols = "";
    for (var i = 0; i < ncols; i++) gridCols += '<w:gridCol w:w="' + colWidth + '"/>';

    function cell(text, isHeader){
      var tcPr = '<w:tcPr><w:tcW w:w="' + colWidth + '" w:type="dxa"/>' +
        (isHeader ? '<w:shd w:val="clear" w:color="auto" w:fill="9B2841"/>' : "") + "</w:tcPr>";
      var rPr = isHeader ? '<w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="19"/></w:rPr>' : '<w:rPr><w:sz w:val="19"/></w:rPr>';
      return "<w:tc>" + tcPr + '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' + runBold(text || "", rPr) + "</w:p></w:tc>";
    }
    function row(cells, isHeader){
      var out = "";
      for (var c = 0; c < ncols; c++) out += cell(cells[c], isHeader);
      return "<w:tr>" + out + "</w:tr>";
    }
    var body = row(block.header, true);
    block.rows.forEach(function(r){ body += row(r, false); });
    var borders = TABLE_BORDER_EDGES.map(function(edge){
      return "<w:" + edge + ' w:val="single" w:sz="4" w:space="0" w:color="D9DFE7"/>';
    }).join("");
    return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' + borders + "</w:tblBorders></w:tblPr>" +
      "<w:tblGrid>" + gridCols + "</w:tblGrid>" + body + "</w:tbl>";
  }

  function blockToXml(b){
    return b.style === "table" ? (blockToTableXml(b) + "<w:p/>") : blockToParagraphXml(b);
  }

  function buildDocumentXml(blocks){
    var body = blocks.map(blockToXml).join("") +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + "</w:body></w:document>";
  }

  function slugFilename(){
    var candidat = missionText("candidat");
    var base = "Dossier_Candidat" + (candidat ? "_" + candidat.replace(/\s+/g, "_") : "");
    base = base.replace(/[^A-Za-z0-9_\-]/g, "");
    return (base || "Dossier_Candidat") + ".docx";
  }

  function insideAppsmith(){
    return !!(window.appsmith
      && typeof window.appsmith.updateModel === "function"
      && typeof window.appsmith.triggerEvent === "function"
      && typeof window.appsmith.onModelChange === "function");
  }

  function blobToDataUrl(blob){
    return new Promise(function(resolve, reject){
      var reader = new FileReader();
      reader.onload = function(){ resolve(String(reader.result || "")); };
      reader.onerror = function(){ reject(new Error("blob illisible")); };
      reader.readAsDataURL(blob);
    });
  }

  function saveDirect(filename, blob){
    var url = URL.createObjectURL(blob);
    var link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.rel = "noopener";
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    setTimeout(function(){
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 4000);
    return Promise.resolve({ status: "saved" });
  }

  var APPSMITH_DOWNLOAD_TIMEOUT_MS = 20000;

  function saveViaAppsmith(filename, blob){
    return blobToDataUrl(blob).then(function(dataUrl){
      // Le jeton identifie CE fichier : l'événement n'est émis qu'une fois
      // l'aller-retour confirmé par Appsmith (le modèle nous revient), sinon
      // le gestionnaire d'événement lirait encore la valeur précédente.
      var token = String(Date.now()) + "-" + Math.random().toString(16).slice(2);
      return new Promise(function(resolve, reject){
        var settled = false;
        var unsubscribe = null;
        var timer = setTimeout(function(){
          if (settled) return;
          settled = true;
          if (unsubscribe) unsubscribe();
          reject(new Error("Appsmith n'a pas confirmé la réception du fichier"));
        }, APPSMITH_DOWNLOAD_TIMEOUT_MS);

        var stop = window.appsmith.onModelChange(function(model){
          // onModelChange rappelle immédiatement avec le modèle courant, qui
          // ne porte pas encore ce jeton : cet appel-là ne déclenche rien.
          if (settled || !model || model.docxToken !== token) return;
          settled = true;
          clearTimeout(timer);
          if (unsubscribe) unsubscribe();
          window.appsmith.triggerEvent("onDownloadDocx");
          resolve({ status: "sent" });
        });
        unsubscribe = stop;
        if (settled) stop();

        window.appsmith.updateModel({
          docxData: dataUrl,
          docxFilename: filename,
          docxToken: token
        });
      });
    });
  }

  function saveFile(filename, blob){
    return insideAppsmith() ? saveViaAppsmith(filename, blob) : saveDirect(filename, blob);
  }

  function runDownload(){
    if (typeof JSZip === "undefined"){
      showToast("La bibliothèque d'export Word n'a pas pu être chargée. Rechargez la page.", "error");
      return;
    }
    els.downloadBtn.disabled = true;
    var previousLabel = els.downloadBtn.textContent;
    els.downloadBtn.textContent = "Préparation du fichier…";

    var zip = new JSZip();
    zip.file("[Content_Types].xml", CT_XML);
    zip.folder("_rels").file(".rels", RELS_XML);
    var wordFolder = zip.folder("word");
    wordFolder.file("document.xml", buildDocumentXml(documentBlocks()));
    wordFolder.file("styles.xml", STYLES_XML);
    wordFolder.folder("_rels").file("document.xml.rels", DOC_RELS_XML);

    zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" })
      .then(function(blob){ return saveFile(slugFilename(), blob); })
      .then(function(res){
        showToast(res.status === "saved" ? "Document Word enregistré." : "Document transmis.", "ok");
      })
      .catch(function(){
        showToast("Impossible de générer le fichier Word. Réessayez.", "error");
      })
      .finally(function(){
        els.downloadBtn.textContent = previousLabel;
        refreshActionState();
      });
  }

  els.downloadBtn.addEventListener("click", runDownload);

  // ===============================================================
  // Mail client
  // ===============================================================

  function runCopyMail(){
    var mail = mailText();
    if (!mail) return;
    if (navigator.clipboard && navigator.clipboard.writeText){
      navigator.clipboard.writeText(mail)
        .then(function(){ showToast("Mail client copié dans le presse-papiers.", "ok"); })
        .catch(function(){ showToast("Impossible de copier automatiquement — sélectionnez le texte dans l'aperçu.", "error"); });
    } else {
      showToast("Copie automatique indisponible — sélectionnez le texte dans l'aperçu.", "error");
    }
  }

  els.copyBtn.addEventListener("click", runCopyMail);

  // Le mail commence par une ligne « Objet : … ». La copie garde cette ligne
  // (utile pour un collage dans un brouillon déjà ouvert) ; l'envoi la détache
  // pour alimenter le vrai champ objet du lien mailto:.
  function splitSubjectAndBody(fullText){
    var m = /^\s*Objet\s*:\s*(.+?)\s*\n+([\s\S]*)$/.exec(fullText);
    if (m) return { subject: m[1].trim(), body: m[2].trim() };
    return { subject: "", body: fullText };
  }

  // La longueur maximale d'un lien mailto: n'est normalisée nulle part et varie
  // selon le système et le client de messagerie ; 1 800 caractères est le
  // plancher commun le plus sûr. Au-delà, certains clients tronquent le corps
  // du message sans rien signaler — exactement le genre de perte silencieuse
  // que cet outil ne doit pas risquer. On copie alors plutôt que de parier.
  var MAILTO_SAFE_CHARS = 1800;

  function runSendMail(){
    var mail = mailText();
    if (!mail) return;
    var split = splitSubjectAndBody(mail);
    var mailtoUrl = "mailto:?subject=" + encodeURIComponent(split.subject) + "&body=" + encodeURIComponent(split.body);
    if (mailtoUrl.length > MAILTO_SAFE_CHARS){
      if (navigator.clipboard && navigator.clipboard.writeText){
        navigator.clipboard.writeText(mail).then(function(){
          showToast("Mail trop long pour une ouverture directe — il a été copié dans le presse-papiers, collez-le dans un nouveau message.", "error");
        }).catch(function(){
          showToast("Mail trop long pour une ouverture directe — copiez-le manuellement depuis l'aperçu.", "error");
        });
      } else {
        showToast("Mail trop long pour une ouverture directe — copiez-le manuellement depuis l'aperçu.", "error");
      }
      return;
    }
    window.location.href = mailtoUrl;
  }

  els.sendBtn.addEventListener("click", runSendMail);

  // ===============================================================
  // Démarrage
  // ===============================================================

  var isStubMode = !backendEngine();
  els.demoBanner.hidden = !isStubMode;

  updateCounts();
  renderPreview();
  setStatus(isStubMode ? "idle" : "ready", isStubMode ? "Mode démonstration" : "Prêt");
})();
