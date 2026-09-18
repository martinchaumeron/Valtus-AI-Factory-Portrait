(function(){
  "use strict";

  // ---------------------------------------------------------------
  // State
  // ---------------------------------------------------------------
  var state = {
    briefText: "",       // raw text returned by the LLM (transcript-only analysis)
    generating: false,
    hasGenerated: false,
    sample: null,
    downloads: null,
    t2: {
      outputText: "",
      generating: false,
      hasGenerated: false,
      syncedOnce: false
    }
  };

  var els = {
    transcript: document.getElementById("transcript"),
    skills: document.getElementById("skills"),
    complementary: document.getElementById("complementary"),
    transcriptCount: document.getElementById("transcriptCount"),
    skillsCount: document.getElementById("skillsCount"),
    complementaryCount: document.getElementById("complementaryCount"),
    transcriptWarn: document.getElementById("transcriptWarn"),
    generateBtn: document.getElementById("generateBtn"),
    generateLabel: document.getElementById("generateLabel"),
    generateHint: document.getElementById("generateHint"),
    downloadBtn: document.getElementById("downloadBtn"),
    docPage: document.getElementById("docPage"),
    statusChip: document.getElementById("statusChip"),
    statusLabel: document.getElementById("statusLabel"),
    candidate: document.getElementById("fieldCandidate"),
    role: document.getElementById("fieldRole"),
    ref: document.getElementById("fieldRef"),
    toast: document.getElementById("toast"),
    demoBanner: document.getElementById("demoBanner"),

    t2Dossier: document.getElementById("t2Dossier"),
    t2Cv: document.getElementById("t2Cv"),
    t2Brief: document.getElementById("t2Brief"),
    t2CustomPrefs: document.getElementById("t2CustomPrefs"),
    t2DossierCount: document.getElementById("t2DossierCount"),
    t2CvCount: document.getElementById("t2CvCount"),
    t2BriefCount: document.getElementById("t2BriefCount"),
    t2GenerateBtn: document.getElementById("t2GenerateBtn"),
    t2GenerateLabel: document.getElementById("t2GenerateLabel"),
    t2GenerateHint: document.getElementById("t2GenerateHint"),
    t2DownloadBtn: document.getElementById("t2DownloadBtn"),
    t2AiWarning: document.getElementById("t2AiWarning"),
    t2CopyBtn: document.getElementById("t2CopyBtn"),
    t2SendBtn: document.getElementById("t2SendBtn"),
    t2DocPage: document.getElementById("t2DocPage"),
    t2SyncBtn: document.getElementById("t2SyncBtn"),
    t2Societe: document.getElementById("t2Societe"),
    t2Contact: document.getElementById("t2Contact"),
    t2Lieu: document.getElementById("t2Lieu"),
    t2Taux: document.getElementById("t2Taux"),
    t2Signataire: document.getElementById("t2Signataire")
  };

  // Custom mail preferences (Étape 2): a free-text note the consultant fills
  // in once (style, signature format, tone tweaks…) that gets applied to
  // every mail generated afterwards, without having to retype it each time.
  // Persisted with localStorage — a per-viewer convenience local to this
  // browser (not shared between consultants, never read back by anyone but
  // this same browser) — wrapped in try/catch since the accessor can throw
  // or simply be unavailable in some views (private window, blocked site
  // data…); the tool must keep working normally either way.
  var T2_CUSTOM_PREFS_KEY = "valtus_t2_mail_custom_prefs_v1";
  try {
    var savedCustomPrefs = window.localStorage.getItem(T2_CUSTOM_PREFS_KEY);
    if (savedCustomPrefs) els.t2CustomPrefs.value = savedCustomPrefs;
  } catch (e) { /* localStorage unavailable in this view — field just starts empty */ }
  els.t2CustomPrefs.addEventListener("input", function(){
    try { window.localStorage.setItem(T2_CUSTOM_PREFS_KEY, els.t2CustomPrefs.value); }
    catch (e) { /* nothing to do — the preference just won't survive a reload here */ }
  });
  function t2GetCustomPrefsText(){ return els.t2CustomPrefs.value.trim(); }

  var MAX_PROMPT_BYTES = 64000; // sample() hard-caps total input at 64 KiB (65 536 bytes); this is the practical ceiling, not an arbitrary choice

  function byteLength(s){ return new TextEncoder().encode(s).length; }
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

  // ---------------------------------------------------------------
  // Navigation — page unique
  // ---------------------------------------------------------------
  // Les deux étapes tiennent sur une seule page : la barre d'onglets, qui
  // masquait une étape sur deux, devient un sommaire d'ancres. Il ne cache
  // plus rien — il fait défiler jusqu'à l'étape visée et souligne l'étape
  // effectivement à l'écran.
  var stepLinks = Array.prototype.slice.call(document.querySelectorAll(".step-link"));
  // Les cibles des ancres, pour que le sommaire suive l'étape à l'écran.
  var stepSections = stepLinks.map(function(link){
    return document.getElementById(link.getAttribute("href").slice(1));
  });

  stepLinks.forEach(function(link, index){
    link.addEventListener("click", function(){
      setActiveStepLink(index);
      // Première arrivée sur l'Étape 2 : on y reprend le dossier de l'Étape 1,
      // exactement ce que faisait l'ouverture de l'onglet 2 auparavant.
      if (index === 1) t2SyncOnce();
      // L'ancre du lien suffirait dans un navigateur ordinaire, mais certains
      // hôtes embarqués bloquent la navigation par fragment, et le défilement
      // fluide y est parfois purement ignoré (constaté à la vérification). Ce
      // défilement explicite, immédiat, vise exactement la même position que
      // l'ancre : les deux mécanismes ne peuvent pas se contredire, et le lien
      // fait toujours quelque chose.
      var target = stepSections[index];
      if (target) target.scrollIntoView({ block: "start" });
    });
  });

  function setActiveStepLink(index){
    stepLinks.forEach(function(link, i){
      link.classList.toggle("active", i === index);
      link.setAttribute("aria-current", i === index ? "true" : "false");
    });
  }

  // Défilement libre (sans clic) : le sommaire suit l'étape à l'écran.
  if (typeof IntersectionObserver === "function"){
    var stepObserver = new IntersectionObserver(function(entries){
      entries.forEach(function(entry){
        if (!entry.isIntersecting) return;
        var index = stepSections.indexOf(entry.target);
        if (index !== -1) setActiveStepLink(index);
      });
    }, { rootMargin: "-20% 0px -70% 0px" });
    stepSections.forEach(function(section){ if (section) stepObserver.observe(section); });
  }

  // Reprise du dossier de l'Étape 1 dans l'Étape 2, au plus une fois : au-delà,
  // ce serait écraser le travail de l'utilisateur sur le champ. La reprise
  // manuelle reste disponible à tout moment via le bouton « ↻ Sync. ».
  function t2SyncOnce(){
    if (state.t2.syncedOnce) return;
    t2SyncFromStep1();
    state.t2.syncedOnce = true;
  }

  // ---------------------------------------------------------------
  // Character counters + file import
  // ---------------------------------------------------------------
  function updateCounts(){
    els.transcriptCount.textContent = els.transcript.value.length.toLocaleString("fr-FR") + " caractère" + (els.transcript.value.length === 1 ? "" : "s");
    els.skillsCount.textContent = els.skills.value.length.toLocaleString("fr-FR") + " caractère" + (els.skills.value.length === 1 ? "" : "s");
    els.complementaryCount.textContent = els.complementary.value.length.toLocaleString("fr-FR") + " caractère" + (els.complementary.value.length === 1 ? "" : "s");

    var promptBytes = byteLength(buildPrompt(els.transcript.value));
    if (promptBytes > MAX_PROMPT_BYTES){
      els.transcriptWarn.textContent = promptBytes.toLocaleString("fr-FR") + " caractères — au-delà de " + MAX_PROMPT_BYTES.toLocaleString("fr-FR") + ", analysé automatiquement en plusieurs lots.";
      els.transcriptWarn.classList.remove("warn");
    } else {
      els.transcriptWarn.textContent = "Budget LLM : " + promptBytes.toLocaleString("fr-FR") + " / " + MAX_PROMPT_BYTES.toLocaleString("fr-FR") + " caractères";
      els.transcriptWarn.classList.remove("warn");
    }
    refreshGenerateState();
  }

  [els.transcript, els.skills, els.complementary].forEach(function(t){
    t.addEventListener("input", function(){ updateCounts(); renderPreview(); });
  });

  document.querySelectorAll(".btn-file[data-target]").forEach(function(btn){
    btn.addEventListener("click", function(){
      var target = btn.getAttribute("data-target");
      document.querySelector('input[data-file-for="' + target + '"]').click();
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

  document.querySelectorAll('input[type="file"]').forEach(function(input){
    input.addEventListener("change", function(){
      var file = input.files && input.files[0];
      if (!file) return;
      var target = input.getAttribute("data-file-for");
      var isDocx = /\.docx$/i.test(file.name);
      var isPdf = /\.pdf$/i.test(file.name);
      var kind = isDocx ? " (.docx)" : isPdf ? " (.pdf)" : "";
      var task = isDocx ? readDocxAsText(file) : isPdf ? readPdfAsText(file) : readPlainText(file);
      task.then(function(text){
        if (!text.trim() && isPdf){
          showToast("Aucun texte trouvé dans ce PDF (page scannée sans OCR ?). Collez le texte manuellement.", "error");
          return;
        }
        document.getElementById(target).value = text;
        updateCounts();
        renderPreview();
        t2UpdateCounts();
        t2RenderPreview();
      }).catch(function(){
        showToast("Impossible de lire ce fichier" + kind + " — collez son contenu manuellement.", "error");
      }).finally(function(){ input.value = ""; });
    });
  });

  [els.candidate, els.role, els.ref].forEach(function(f){
    f.addEventListener("input", renderPreview);
  });

  // ---------------------------------------------------------------
  // Prompt construction (transcript-only analysis)
  // ---------------------------------------------------------------
  function buildPrompt(transcript){
    return "Tu es un consultant senior en Executive Search et Management de Transition. Tu vas recevoir le transcript brut d'un entretien entre un recruteur et un candidat.\n\n" +
"Ta mission est d'extraire exclusivement les informations fournies par le candidat et de restituer ton travail sous la forme d'un « Dossier d'analyse consolidé ».\n\n" +
"CONSIGNES STRICTES\n" +
"1. N'extrais AUCUNE information provenant du recruteur.\n" +
"2. N'invente rien.\n" +
"3. N'interprète rien.\n" +
"4. N'extrapole rien.\n" +
"5. N'enjolive rien.\n" +
"6. Ne reformule pas de façon à renforcer artificiellement les réalisations du candidat.\n" +
"7. Lorsqu'une information est incertaine, ambiguë ou imprécise, indique explicitement qu'elle est déclarative ou approximative.\n" +
"8. N'attribue jamais une responsabilité, un titre, une compétence ou un résultat qui n'a pas été explicitement mentionné par le candidat.\n" +
"9. Le document final doit être 100% factuel et traçable aux propos du candidat.\n" +
"10. Si un élément n'est pas exprimé par le candidat, ne le mentionne pas.\n" +
"11. Lorsqu'un sujet a été explicitement testé par le recruteur pendant l'entretien (expertise comptable, gouvernance associative, management d'experts, expérience d'un secteur donné, etc.), crée systématiquement un paragraphe dédié, même si le candidat n'y consacre que quelques minutes.\n\n" +
"OBJECTIF\n" +
"Produis une extraction exhaustive de toutes les informations utiles concernant : le parcours professionnel ; les expériences ; les responsabilités ; les réalisations ; les compétences ; les savoir-faire ; les secteurs d'activité ; les expertises fonctionnelles, métiers et techniques ; les expériences internationales ; les expériences managériales ; la gestion de P&L ; les expériences de transformation ; les expériences M&A ; les expériences de levée de fonds ; les expériences de création d'entreprise ; les expériences de croissance ou d'hypercroissance ; le mode de fonctionnement ; les qualités et aptitudes revendiquées ; les disponibilités ; et toute information factuelle utile à l'évaluation d'une candidature de dirigeant ou de manager de transition.\n\n" +
"FORMAT ATTENDU\n" +
"Construis un récit professionnel structuré autour des expériences significatives du candidat. Pour chaque expérience, décris le contexte, le périmètre, les responsabilités, les réalisations, les transformations conduites, les enjeux humains et opérationnels, ainsi que les éléments transférables. Évite les listes thématiques génériques lorsque les informations peuvent être rattachées à une expérience précise. Si des éléments ne se rattachent à aucune expérience précise (mode de fonctionnement général, qualités et aptitudes revendiquées, disponibilité), regroupe-les dans une dernière section dédiée. Le document doit se lire comme une note exécutive de présentation d'un candidat à destination d'un client dirigeant ou investisseur.\n\n" +
"MODE DE RÉDACTION\n" +
"Rédige en style direct, jamais au format déclaratif. Par exemple, n'écris pas « M. Denis déclare avoir conduit une transformation », mais « M. Denis a conduit une transformation ». N'écris pas « Il indique avoir conduit des réflexions sur la segmentation des clients », mais « Il a conduit des réflexions sur la segmentation des clients ».\n\n" +
"IMPORTANT\n" +
"- Ne porte aucun jugement.\n" +
"- N'évalue pas le candidat.\n" +
"- N'écris jamais qu'il est « excellent », « solide », « expérimenté », « adapté », « pertinent » ou tout autre qualificatif subjectif.\n" +
"- Reste strictement descriptif.\n\n" +
"FORMAT DE SORTIE (technique, à respecter) :\n" +
"- Ne mets AUCUN titre de niveau 1 et n'écris pas « Dossier d'analyse consolidé » : commence directement par la première section.\n" +
"- Précède chaque section d'une ligne \"## Titre de la section\" (deux dièses, un espace, puis le titre). Utilise un titre par expérience significative (par exemple \"## Directeur Général — Société X (2018-2022)\"), dans l'ordre chronologique ou logique du parcours, puis une dernière section pour les éléments transverses (mode de fonctionnement, qualités revendiquées, disponibilité) si applicable.\n" +
"- Rédige chaque section en paragraphes de prose continue. N'utilise des puces (\"- \") que pour la section transverse finale, si elle contient une liste de qualités ou compétences revendiquées.\n" +
"- Rédige en français.\n" +
"- N'ajoute aucun commentaire méta (« Voici le dossier… »), aucune conclusion générale, aucune balise de code.\n\n" +
"RETRANSCRIPTION BRUTE DE L'ENTRETIEN (source unique et exclusive) :\n\"\"\"\n" + transcript + "\n\"\"\"";
  }

  // Batch mode (transcript too large for one call): a "map" pass extracts
  // facts from each excerpt, then a "reduce" pass organizes all of them
  // into the final brief. Nothing from the transcript is skipped — it is
  // only ever split on paragraph boundaries, never summarized away.
  function buildMapPromptT1(chunkText, idx, total){
    return "Tu es un consultant senior en Executive Search et Management de Transition. Ceci est une étape INTERMÉDIAIRE : tu reçois l'extrait " + idx + "/" + total + " d'un transcript d'entretien candidat plus long, découpé pour l'analyse. Ta seule mission ici est d'extraire, de façon exhaustive, TOUTES les informations factuelles fournies par le CANDIDAT dans CET EXTRAIT. Ce n'est pas le document final : pas encore de mise en forme narrative élaborée, juste une collecte complète et fidèle.\n\n" +
"CONSIGNES STRICTES (identiques à l'analyse complète)\n" +
"1. N'extrais AUCUNE information provenant du recruteur — uniquement les propos du candidat.\n" +
"2. N'invente rien, n'interprète rien, n'extrapole rien, n'enjolive rien.\n" +
"3. Ne reformule pas de façon à renforcer artificiellement les réalisations du candidat.\n" +
"4. Indique explicitement le caractère déclaratif ou approximatif d'une information incertaine.\n" +
"5. Si un sujet a été explicitement testé par le recruteur dans cet extrait (même brièvement), signale-le clairement et note ce qui a été dit, même en quelques mots.\n" +
"6. Ne porte aucun jugement, n'évalue pas le candidat, aucun qualificatif subjectif.\n" +
"7. Si un élément n'est pas exprimé par le candidat dans cet extrait, ne le mentionne pas.\n\n" +
"FORMAT : une liste dense de faits, organisée par expérience ou thème si identifiable dans cet extrait, en style direct (« il a conduit », jamais « il déclare avoir conduit »). Sois exhaustif — c'est une étape de collecte, pas de synthèse. N'ajoute aucun commentaire méta ni aucune conclusion.\n\n" +
"EXTRAIT DU TRANSCRIPT (partie " + idx + " sur " + total + ") :\n\"\"\"\n" + chunkText + "\n\"\"\"";
  }

  function buildReducePromptT1(mapOutputsJoined){
    return "Tu es un consultant senior en Executive Search et Management de Transition. Un entretien candidat volumineux a déjà été analysé PAR LOTS : tu reçois ci-dessous les notes factuelles déjà extraites pour chaque lot (uniquement des propos du candidat, sans invention ni interprétation). Ta mission est de les organiser en un « Dossier d'analyse consolidé » final, SANS ajouter aucune information qui n'y figure pas déjà.\n\n" +
"CONSIGNES STRICTES\n" +
"1. N'utilise que les informations présentes dans les notes ci-dessous — ne retourne jamais au transcript original, tu ne l'as pas.\n" +
"2. N'invente rien, n'interprète rien, n'extrapole rien, n'enjolive rien.\n" +
"3. Conserve les mentions de caractère déclaratif ou approximatif déjà signalées dans les notes.\n" +
"4. Conserve un paragraphe dédié pour chaque sujet signalé comme testé par le recruteur, même brièvement.\n" +
"5. Ne porte aucun jugement, n'évalue pas le candidat, aucun qualificatif subjectif.\n" +
"6. Si un même fait apparaît dans plusieurs lots, ne le restitue qu'une seule fois.\n\n" +
"FORMAT ATTENDU\n" +
"Construis un récit professionnel structuré autour des expériences significatives du candidat (regroupe les notes par expérience si elles sont dispersées entre plusieurs lots). Pour chaque expérience, décris le contexte, le périmètre, les responsabilités, les réalisations, les transformations conduites, les enjeux humains et opérationnels, ainsi que les éléments transférables. Regroupe les éléments transverses (mode de fonctionnement, qualités revendiquées, disponibilité) dans une dernière section dédiée. Le document doit se lire comme une note exécutive de présentation d'un candidat à destination d'un client dirigeant ou investisseur.\n\n" +
"MODE DE RÉDACTION\n" +
"Style direct, jamais déclaratif (« M. Denis a conduit », jamais « M. Denis déclare avoir conduit »).\n\n" +
"FORMAT DE SORTIE (technique)\n" +
"- Ne mets aucun titre de niveau 1 et n'écris pas « Dossier d'analyse consolidé » : commence directement par la première section.\n" +
"- Précède chaque section d'une ligne \"## Titre\" (un titre par expérience, puis une section transverse finale si besoin).\n" +
"- Prose continue par section ; puces (\"- \") uniquement pour la section transverse finale.\n" +
"- Rédige en français. N'ajoute aucun commentaire méta, aucune balise de code.\n\n" +
"NOTES FACTUELLES EXTRAITES PAR LOTS (source unique et exclusive, à consolider — ne retraite pas le transcript original) :\n\"\"\"\n" + mapOutputsJoined + "\n\"\"\"";
  }

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

  function blocksToPlainText(blocks){
    return blocks.map(function(b){
      switch (b.style){
        case "title": case "h1": return "# " + b.text;
        case "meta": return b.text;
        case "h2": return "## " + b.text;
        case "h3": return "### " + b.text;
        case "bullet": return "- " + b.text;
        case "placeholder": return b.text;
        case "spacer": return "";
        case "table": return tableBlockToPlainText(b);
        default: return b.text || "";
      }
    }).join("\n\n").replace(/\n{3,}/g, "\n\n").trim();
  }

  function tableBlockToPlainText(b){
    var lines = ["| " + b.header.join(" | ") + " |", "|" + b.header.map(function(){ return "---"; }).join("|") + "|"];
    b.rows.forEach(function(r){ lines.push("| " + r.join(" | ") + " |"); });
    return lines.join("\n");
  }

  // ---------------------------------------------------------------
  // Batch processing (map-reduce) — used automatically whenever a source
  // document is too large for a single sample() call (64 KiB hard cap).
  // Splitting always happens on natural boundaries (blank lines, "## "
  // headings) so no sentence is ever cut mid-way and nothing is silently
  // dropped: every chunk is sent in full, just in separate calls.
  // ---------------------------------------------------------------
  // Packs an array of text pieces (paragraphs, "## " sections…) into as few
  // chunks as possible while respecting targetChars: adjacent small pieces
  // are merged together (so e.g. many short experience sections don't each
  // trigger their own sample() call), and a single piece longer than the
  // target is hard-split as a last resort. Never drops or rewrites content.
  function mergeIntoChunks(pieces, targetChars){
    var chunks = [];
    var current = "";
    pieces.forEach(function(piece){
      while (piece.length > targetChars * 1.5){
        if (current){ chunks.push(current); current = ""; }
        chunks.push(piece.slice(0, targetChars));
        piece = piece.slice(targetChars);
      }
      var candidate = current ? current + "\n\n" + piece : piece;
      if (candidate.length > targetChars && current){
        chunks.push(current);
        current = piece;
      } else {
        current = candidate;
      }
    });
    if (current) chunks.push(current);
    return chunks.length ? chunks : [""];
  }

  function splitByParagraphs(text, targetChars){
    var paras = String(text || "").replace(/\r\n/g, "\n").split(/\n{2,}/);
    return mergeIntoChunks(paras, targetChars);
  }

  function splitByH2(text){
    var lines = String(text || "").replace(/\r\n/g, "\n").split("\n");
    var sections = [];
    var current = [];
    lines.forEach(function(line){
      if (/^##\s+/.test(line.trim()) && current.length){
        sections.push(current.join("\n"));
        current = [line];
      } else {
        current.push(line);
      }
    });
    if (current.length) sections.push(current.join("\n"));
    return sections.filter(function(s){ return s.trim(); });
  }

  // Splits a t2 "Dossier candidat" text into: the consolidated-dossier
  // narrative (chunkable by "## " experience headings), plus the "Analyse
  // externe" and "Informations complémentaires" sections kept whole (they
  // are normally short, and every map call needs them for the hierarchy
  // rules regardless of which experience it is processing).
  function splitDossierSections(text){
    var lines = String(text || "").replace(/\r\n/g, "\n").split("\n");
    var idxExterne = -1, idxCompl = -1;
    for (var i = 0; i < lines.length; i++){
      var l = lines[i].trim();
      if (idxExterne === -1 && /^#\s+Analyse externe/i.test(l)) idxExterne = i;
      if (idxCompl === -1 && /^#\s+Informations compl/i.test(l)) idxCompl = i;
    }
    var boundaries = [idxExterne, idxCompl].filter(function(x){ return x !== -1; }).sort(function(a, b){ return a - b; });
    var narrativeEnd = boundaries.length ? boundaries[0] : lines.length;
    var narrative = lines.slice(0, narrativeEnd).join("\n");
    var externe = idxExterne !== -1
      ? lines.slice(idxExterne, idxCompl !== -1 && idxCompl > idxExterne ? idxCompl : lines.length).join("\n")
      : "";
    var compl = idxCompl !== -1
      ? lines.slice(idxCompl, idxExterne !== -1 && idxExterne > idxCompl ? idxExterne : lines.length).join("\n")
      : "";
    return { narrative: narrative, externe: externe, complementaire: compl };
  }

  function t2SplitDossierForBatches(dossierText){
    var parts = splitDossierSections(dossierText);
    var rawSections = splitByH2(parts.narrative);
    // With 2+ experience headings, pack them into as few map calls as
    // possible (merging short experiences together) instead of one call per
    // heading. With 0 or 1 heading, fall back to paragraph-boundary splitting
    // of the whole narrative.
    var sections = rawSections.length >= 2
      ? mergeIntoChunks(rawSections, 40000)
      : splitByParagraphs(parts.narrative, 40000);
    return { sections: sections, externe: parts.externe, complementaire: parts.complementaire };
  }

  // Splits a CV into chunks — only invoked when the CV alone (with the fixed
  // per-call overhead: instructions, analyse externe, informations
  // complémentaires) is too large to keep whole in every map call. Most CVs
  // never hit this; long-career executive CVs can. Same non-lossy,
  // paragraph-boundary packing as the dossier splitter — nothing is dropped
  // or rewritten, only cut on blank-line boundaries (or hard-sliced as a
  // last resort for a single huge unbroken block of text).
  function t2SplitCvForBatches(cvText, targetChars){
    return splitByParagraphs(cvText, targetChars);
  }

  function rawBlocks(text){
    var t = String(text || "").replace(/\r\n/g, "\n");
    if (!t.trim()) return [{ style: "placeholder", text: "Non renseigné." }];
    var blocks = [];
    t.split("\n").forEach(function(line){
      if (line.trim() === "") blocks.push({ style: "spacer" });
      else blocks.push({ style: "p", text: line });
    });
    return blocks;
  }

  function documentBlocks(){
    var candidate = els.candidate.value.trim();
    var role = els.role.value.trim();
    var ref = els.ref.value.trim();
    var metaParts = [];
    if (candidate) metaParts.push("Candidat : " + candidate);
    if (role) metaParts.push("Mission : " + role);
    if (ref) metaParts.push("Référence : " + ref);
    metaParts.push("Préparé le : " + new Date().toLocaleDateString("fr-FR", { year: "numeric", month: "long", day: "numeric" }));

    var blocks = [];
    blocks.push({ style: "title", text: "Dossier d'analyse consolidé" });
    blocks.push({ style: "meta", text: metaParts.join("   ·   ") });

    if (state.briefText.trim()){
      blocks = blocks.concat(parseBrief(state.briefText));
    } else {
      blocks.push({ style: "placeholder", text: "Le dossier d'analyse consolidé n'a pas encore été généré à partir de la retranscription." });
    }

    blocks.push({ style: "spacer" });
    blocks.push({ style: "h1", text: "Analyse externe des compétences" });
    blocks = blocks.concat(rawBlocks(els.skills.value));

    blocks.push({ style: "spacer" });
    blocks.push({ style: "h1", text: "Informations complémentaires du candidat" });
    blocks = blocks.concat(rawBlocks(els.complementary.value));

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

  function renderPreview(){
    els.docPage.innerHTML = blocksToHtml(documentBlocks());
    refreshDownloadState();
  }

  // ---------------------------------------------------------------
  // Generate state / actions
  // ---------------------------------------------------------------
  var BATCH_CHUNK_CHARS = 42000;

  function refreshGenerateState(){
    var transcriptOk = els.transcript.value.trim().length > 0;
    var promptBytes = byteLength(buildPrompt(els.transcript.value));
    var willBatch = promptBytes > MAX_PROMPT_BYTES;
    var canGenerate = transcriptOk && !state.generating && !!state.sample;
    els.generateBtn.disabled = !canGenerate;
    if (!state.sample){
      els.generateHint.textContent = "Génération indisponible dans cette vue.";
      els.generateHint.classList.add("warn");
    } else if (!transcriptOk){
      els.generateHint.textContent = "Collez ou importez une retranscription d'entretien pour commencer.";
      els.generateHint.classList.remove("warn");
    } else if (willBatch){
      var n = splitByParagraphs(els.transcript.value, BATCH_CHUNK_CHARS).length;
      els.generateHint.textContent = "Retranscription volumineuse : analyse automatique en " + n + " lots, puis fusion en un seul dossier.";
      els.generateHint.classList.remove("warn");
    } else {
      els.generateHint.textContent = "Le LLM n'analyse que la retranscription d'entretien.";
      els.generateHint.classList.remove("warn");
    }
  }

  function refreshDownloadState(){
    els.downloadBtn.disabled = !state.downloads || !state.hasGenerated || state.generating;
  }

  var SAMPLE_ERROR_COPY = {
    not_granted: "L'accès à Claude n'a pas été autorisé pour cette page.",
    sampling_disabled: "La génération via Claude n'est pas disponible pour ce compte.",
    session_expired: "Votre session a expiré — reconnectez-vous puis réessayez.",
    rate_limited: "Trop de demandes en cours. Réessayez dans un instant.",
    prompt_too_large: "La retranscription est trop longue pour être analysée en une seule fois. Réduisez-la puis réessayez.",
    refused: "La demande n'a pas pu être traitée. Vérifiez le contenu de la retranscription et réessayez.",
    empty_completion: "Aucun contenu n'a été produit. Réessayez.",
    cancelled: "Génération annulée.",
    invalid_json: "Réponse inattendue. Réessayez.",
    upstream_error: "Problème temporaire côté service. Réessayez dans quelques instants.",
    invalid_request: "Erreur technique interne. Réessayez.",
    transform_error: "Erreur technique interne. Réessayez.",
    capability_disabled: "Fonction indisponible dans cette vue.",
    capability_removed: "Fonction indisponible dans cette vue.",
    not_declared: "Fonction indisponible dans cette vue."
  };

  function runGenerate(){
    if (!state.sample || state.generating) return;
    var transcript = els.transcript.value.trim();
    if (!transcript) return;

    var singlePrompt = buildPrompt(transcript);
    if (byteLength(singlePrompt) > MAX_PROMPT_BYTES){
      runGenerateBatched(transcript);
      return;
    }

    state.generating = true;
    els.generateBtn.setAttribute("data-busy", "true");
    els.generateBtn.disabled = true;
    els.generateLabel.textContent = state.hasGenerated ? "Régénération…" : "Génération en cours…";
    setStatus("busy", "Génération en cours");
    refreshDownloadState();

    state.sample(singlePrompt, {
      modelTier: "complex",
      cache: false,
      onText: function(update){
        state.briefText = update.text;
        renderPreview();
      }
    }).then(function(result){
      state.briefText = result.text;
      state.hasGenerated = true;
      state.generating = false;
      els.generateBtn.removeAttribute("data-busy");
      els.generateLabel.textContent = "Régénérer le dossier";
      setStatus("done", "Dossier généré");
      if (result.truncated){
        showToast("La réponse a été interrompue par la limite de longueur. Le dossier peut être incomplet.", "error");
      }
      renderPreview();
      refreshGenerateState();
      // Page unique : l'Étape 2, visible juste en dessous, part du dossier
      // qui vient d'être généré plutôt que d'un champ vide.
      t2SyncOnce();
    }).catch(function(err){
      state.generating = false;
      els.generateBtn.removeAttribute("data-busy");
      els.generateLabel.textContent = state.hasGenerated ? "Régénérer le dossier" : "Générer le dossier";
      setStatus("error", "Échec de la génération");
      if (err && err.text) { state.briefText = err.text; renderPreview(); }
      showToast((err && SAMPLE_ERROR_COPY[err.code]) || "Une erreur est survenue pendant la génération.", "error");
      refreshGenerateState();
    });
  }

  // Automatic batch (map-reduce) path for oversized transcripts: N parallel
  // "map" calls extract facts from each excerpt, then one "reduce" call
  // writes the final brief from those notes.
  function runGenerateBatched(transcript){
    var chunks = splitByParagraphs(transcript, BATCH_CHUNK_CHARS);
    var total = chunks.length;

    state.generating = true;
    els.generateBtn.setAttribute("data-busy", "true");
    els.generateBtn.disabled = true;
    els.generateLabel.textContent = "Analyse de " + total + " extraits…";
    setStatus("busy", "Analyse par lots (" + total + " extraits)");
    refreshDownloadState();

    var mapCalls = chunks.map(function(chunk, i){
      return state.sample(buildMapPromptT1(chunk, i + 1, total), { modelTier: "complex", cache: false });
    });

    Promise.all(mapCalls).then(function(results){
      var joined = results.map(function(r, i){
        return "--- Extrait " + (i + 1) + "/" + total + " ---\n" + r.text;
      }).join("\n\n");
      els.generateLabel.textContent = "Rédaction du dossier final…";
      setStatus("busy", "Rédaction du dossier final");
      var reducePrompt = buildReducePromptT1(joined);
      if (byteLength(reducePrompt) > MAX_PROMPT_BYTES){
        return Promise.reject({ code: "prompt_too_large", message: "notes too large even after batching" });
      }
      return state.sample(reducePrompt, {
        modelTier: "complex",
        cache: false,
        onText: function(update){
          state.briefText = update.text;
          renderPreview();
        }
      });
    }).then(function(result){
      state.briefText = result.text;
      state.hasGenerated = true;
      state.generating = false;
      els.generateBtn.removeAttribute("data-busy");
      els.generateLabel.textContent = "Régénérer le dossier";
      setStatus("done", "Dossier généré");
      if (result.truncated){
        showToast("La réponse a été interrompue par la limite de longueur. Le dossier peut être incomplet.", "error");
      }
      renderPreview();
      refreshGenerateState();
      // Page unique : l'Étape 2, visible juste en dessous, part du dossier
      // qui vient d'être généré plutôt que d'un champ vide.
      t2SyncOnce();
    }).catch(function(err){
      state.generating = false;
      els.generateBtn.removeAttribute("data-busy");
      els.generateLabel.textContent = state.hasGenerated ? "Régénérer le dossier" : "Générer le dossier";
      setStatus("error", "Échec de l'analyse par lots");
      var msg = (err && SAMPLE_ERROR_COPY[err.code]) || "Une erreur est survenue pendant l'analyse par lots.";
      if (err && err.code === "prompt_too_large"){
        msg = "Même après découpage, les notes extraites restent trop volumineuses pour la synthèse finale. Réessayez avec une retranscription un peu plus courte.";
      }
      showToast(msg, "error");
      refreshGenerateState();
    });
  }

  els.generateBtn.addEventListener("click", runGenerate);

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
    '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:b/><w:color w:val="1A2230"/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:before="360" w:after="160"/><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="96703C"/></w:pBdr></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:b/><w:caps/><w:color w:val="1A2230"/><w:sz w:val="27"/><w:szCs w:val="27"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:before="240" w:after="120"/><w:keepNext/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/><w:b/><w:color w:val="7A5A2E"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:before="200" w:after="100"/><w:keepNext/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:i/><w:color w:val="1A2230"/><w:sz w:val="22"/></w:rPr></w:style>' +
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

  var MUTED_ITALIC_RPR = '<w:rPr><w:i/><w:color w:val="6B6558"/></w:rPr>';
  var META_RPR = '<w:rPr><w:i/><w:color w:val="6B6558"/><w:sz w:val="19"/></w:rPr>';

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
        (isHeader ? '<w:shd w:val="clear" w:color="auto" w:fill="96703C"/>' : "") + "</w:tcPr>";
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
      return "<w:" + edge + ' w:val="single" w:sz="4" w:space="0" w:color="DCD7CB"/>';
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
    var candidate = els.candidate.value.trim();
    var base = "Dossier_Analyse_Consolide" + (candidate ? "_" + candidate.replace(/\s+/g, "_") : "");
    base = base.replace(/[^A-Za-z0-9_\-]/g, "");
    return (base || "Dossier_Analyse_Consolide") + ".docx";
  }

  function runDownload(){
    if (!state.downloads || typeof JSZip === "undefined") return;
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
      .then(function(blob){
        return state.downloads.save({ filename: slugFilename(), data: blob });
      })
      .then(function(res){
        showToast(res.status === "saved" ? "Document Word enregistré." : "Document transmis.", "ok");
      })
      .catch(function(err){
        if (err && err.code === "declined") { /* silent: viewer chose not to save */ }
        else showToast("Impossible de générer le fichier Word. Réessayez.", "error");
      })
      .finally(function(){
        els.downloadBtn.textContent = previousLabel;
        refreshDownloadState();
      });
  }

  els.downloadBtn.addEventListener("click", runDownload);

  // ---------------------------------------------------------------
  // Étape 2 — Dossier de présentation candidat
  // ---------------------------------------------------------------
  var T2_MAX_PROMPT_BYTES = 64000; // same hard 64 KiB sample() ceiling as Étape 1

  function t2SyncFromStep1(){
    els.t2Dossier.value = blocksToPlainText(documentBlocks());
    t2UpdateCounts();
    t2RenderPreview();
    showToast("Dossier de l'Étape 1 repris dans l'Étape 2.", "ok");
  }
  els.t2SyncBtn.addEventListener("click", t2SyncFromStep1);

  [els.t2Dossier, els.t2Cv, els.t2Brief].forEach(function(t){
    t.addEventListener("input", function(){ t2UpdateCounts(); t2RenderPreview(); });
  });
  [els.t2Societe, els.t2Contact, els.t2Lieu, els.t2Taux, els.t2Signataire].forEach(function(f){
    f.addEventListener("input", t2RenderPreview);
  });

  function t2UpdateCounts(){
    els.t2DossierCount.textContent = els.t2Dossier.value.length.toLocaleString("fr-FR") + " caractères";
    els.t2CvCount.textContent = els.t2Cv.value.length.toLocaleString("fr-FR") + " caractères";
    els.t2BriefCount.textContent = els.t2Brief.value.length.toLocaleString("fr-FR") + " caractères";
    t2RefreshGenerateState();
  }

  // Shared by the single-call and the map-reduce (batch) paths — Étapes 2 à
  // 6 never change: only how Étape 0/1 obtain their tables differs.
  // ---------------------------------------------------------------
  // Two-phase pipeline. A single mega-call asking the model to do Étapes 0
  // à 6 + Contrôle final in one shot produced answers that stopped partway
  // through (typically right after Étape 0/1) on real, information-rich
  // documents — the model ran out of room before reaching the parts the
  // client actually wants (Étape 4, Étape 6), silently leaving no mail. So
  // this is now always two sequential calls:
  //  1. "Analyse" — Étapes 0 à 3 (hierarchy-of-sources reasoning, sourced
  //     competency extraction, profile reconstruction, brief-based
  //     prioritization). This is the long, exploratory part.
  //  2. "Rédaction finale" — Étapes 4 à 6 + Contrôle final, built ONLY from
  //     the first call's already-validated analysis (never the raw sources
  //     again). Its expected output is short and bounded (≤240 words of
  //     competencies + one email + one sentence), so it reliably completes
  //     even when the analysis itself was substantial.
  // The same split is used whether or not the dossier needs batching.
  // ---------------------------------------------------------------
  function t2AnalysisStepsText(){
    return (
"## Étape 2 — Reconstruction du profil\n" +
"À partir des compétences validées (sans en créer de nouvelles), réponds brièvement à : 1. Quel est le métier principal de ce manager ? 2. Quel est son terrain de jeu naturel ? 3. Quels sont ses domaines de légitimité ? 4. Quelles sont les dimensions récurrentes de son parcours ? 5. Quels éléments différenciants sont apparus pendant l'entretien ? 6. Quels éléments des informations complémentaires permettent de mieux comprendre certaines expériences ?\n\n" +
"## Étape 3 — Priorisation selon le brief\n" +
"À partir des seules compétences démontrées, sélectionne les plus pertinentes au regard du brief client et réorganise-les selon ses priorités, sans en ajouter de nouvelles. Vise une dizaine de compétences (10) à titre indicatif, mais adapte ce nombre au profil réel : un profil très spécialisé peut n'en avoir que 7 ou 8 de réellement pertinentes et démontrées, un profil multi-casquettes peut légitimement en avoir 13 à 15. Le critère est la pertinence par rapport au brief et la solidité de la preuve, jamais d'atteindre un chiffre rond : il vaut toujours mieux quelques compétences réellement nécessaires et bien démontrées qu'une liste allongée avec des compétences secondaires ou faiblement démontrées. Pour chacune, explique brièvement pourquoi elle répond au besoin, sans jamais utiliser le brief comme preuve. Termine impérativement par la liste finale des compétences retenues et priorisées, avec pour chacune ses éléments de preuve (CV / entretien / complément) : c'est cette liste, et elle seule, qui servira de base à la rédaction finale dans un appel séparé.\n\n" +
"Format : structure chaque étape avec \"## \", les tableaux au format markdown avec des barres verticales \"|\", les listes avec \"- \". N'ajoute aucun commentaire méta, aucune balise de code.\n\n"
    );
  }

  // t2FinalStepsText() used to be one monolithic block covering Étape 4, 5,
  // 6 and the Contrôle final in a single call. Split in two (t2Step4Text /
  // t2Step56Text) so t2BuildFinalPrompt can fall back to TWO smaller calls
  // — each carrying the full analysis notes — when the combined prompt
  // would exceed the 64 KiB sample() ceiling (large CV/brief/analysis).
  // t2FinalStepsText() itself is kept as the concatenation, used whenever
  // a single call fits (the common case) so nothing else needs to change.
  function t2Step4Text(){
    return (
"On t'a déjà transmis, ci-dessous, l'analyse et la liste des compétences déjà démontrées et priorisées. Produis UNIQUEMENT la section suivante, avec un titre \"## \" et aucune section en dehors de celle-ci — l'analyse a déjà été faite, ne la refais pas et ne la restitue pas. Reprends EXACTEMENT le titre \"## Étape 4 — Compétences clés finales\" mot pour mot.\n\n" +
"## Étape 4 — Compétences clés finales\n" +
"Produis la liste finale des compétences clés reprises telles que priorisées dans l'analyse : une puce par compétence, avec une description courte, puis, uniquement si l'information existe dans l'analyse fournie, une référence entreprise et une référence chiffrée jamais inventée. Reprends le même nombre de compétences que la liste finale de l'analyse fournie (une dizaine à titre indicatif, mais 7-8 pour un profil très spécialisé ou 13-15 pour un profil multi-casquettes sont normaux) — ne complète jamais artificiellement pour atteindre un chiffre rond, ni ne coupe une compétence réellement pertinente et démontrée. Limite l'ensemble à 240 mots ; si le nombre de compétences retenu rend cette limite intenable, raccourcis chaque puce plutôt que d'en supprimer. Produis deux versions successives, avec un sous-titre \"### \" pour chacune : \"### Version avec sources documentaires\" (chaque puce peut préciser CV ou entretien) puis \"### Version sans sources documentaires\" obtenue en supprimant uniquement les mentions de source, sans aucune reformulation, enrichissement ou modification du contenu.\n\n" +
"Format : structure la section avec \"## \", les deux sous-titres avec \"### \", chaque compétence des deux versions en puce (\"- \"). N'ajoute aucun commentaire méta, aucune balise de code.\n\n" +
"MARQUEUR TECHNIQUE OBLIGATOIRE (très important, à respecter à la lettre) : le consultant qui lit ta réponse n'affichera à son client que le contenu de cette section. Pour lui permettre de l'extraire automatiquement, insère seule sur sa ligne, juste avant le contenu (après le titre \"## \") : @@DEBUT_COMPETENCES_CLES@@ — puis juste après le contenu (à la toute fin de ta réponse) : @@FIN_COMPETENCES_CLES@@. Recopie ces deux marqueurs exactement tels quels (mêmes arobases, mêmes majuscules), sans les traduire ni les commenter, et ne les utilise nulle part ailleurs dans ta réponse.\n\n"
    );
  }

  function t2Step56Text(){
    return (
"On t'a déjà transmis, ci-dessous, l'analyse et la liste des compétences déjà démontrées et priorisées. Produis UNIQUEMENT les sections suivantes, avec un titre \"## \" pour chacune et aucune section en dehors de celles listées ci-dessous — l'analyse a déjà été faite, ne la refais pas et ne la restitue pas. Reprends EXACTEMENT le titre \"## Étape 6 — Mail client\" mot pour mot.\n\n" +
"## Étape 5 — Sélection des expériences à mettre en avant\n" +
"Identifie les 3 expériences les plus pertinentes et complémentaires (pas de redondance) parmi celles de l'analyse fournie, en priorité celles qui répondent le mieux aux enjeux du brief de mission. Pour chacune, précise l'entreprise, les compétences démontrées, pourquoi elle est retenue, et ce qu'elle apporte à la démonstration globale. Présente un tableau markdown (colonnes : Expérience | Pourquoi elle est retenue | Compétences démontrées).\n\n" +
"## Étape 6 — Mail client\n" +
"Rédige ensuite un mail client selon exactement cette structure ci-dessous, qui est un gabarit à valeurs manquantes, pas un exemple de style à reformuler : reprends chaque phrase du gabarit telle quelle, en remplaçant uniquement les emplacements entre accolades ({Prénom}, {NOM}, {Société}, {Mission}, {fonction}, {premier employeur}, {lieu de mission}, etc.), ainsi que \"XXXX\" (taux journalier) et \"Didier\" (signataire), par la vraie valeur correspondante extraite de l'analyse, du brief de mission ou des données commerciales fournies plus bas. Ne remplace jamais un emplacement par une périphrase (\"un dirigeant\", \"ce candidat\", \"il\", \"lui\") : {Prénom} et {Prénom NOM} doivent devenir le vrai prénom (et nom) du candidat, à chaque occurrence, y compris quand \"{Prénom}\" apparaît plusieurs fois dans le même paragraphe (par exemple dans la formule de conclusion) et dès la phrase d'introduction (\"Il s'agit de {Prénom NOM} qui dispose...\" devient \"Il s'agit de [vrai prénom NOM] qui dispose...\", jamais \"Il s'agit d'un dirigeant/manager qui dispose...\"). Le mail final ne doit contenir strictement aucune accolade \"{\" ni \"}\", ni \"XXXX\", ni \"Didier\", nulle part.\n" +
"Familiarité : vouvoiement\n" +
"Début :\n" +
"\"Cher/Chère {Nom},\n\nJ'ai le plaisir de vous adresser le profil d'un manager de transition qui me paraît répondre au besoin de {Société} sur la mission {Mission}.\"\n" +
"Puis :\n" +
"\"Il s'agit de {Prénom NOM} qui dispose de plus de {X années} ans d'expérience dans le métier de {fonction}.\"\n" +
"Puis :\n" +
"\"{Prénom NOM} commence sa carrière chez {premier employeur} et évolue ensuite rapidement vers...\"\n" +
"Puis :\n" +
"- 1 paragraphe de 3 lignes sur le parcours ;\n" +
"- 1 paragraphe de 3 lignes sur les forces observées pendant l'entretien.\n" +
"Puis :\n" +
"\"Quelques réalisations notables de {prénom du candidat} :\"\n" +
"Puis :\n" +
"Trois bullet points de 2 à 3 lignes chacun décrivant une mission récente qui illustre son approche dans le contexte de la mission. Pour chaque exemple :\n" +
"- préciser l'entreprise ;\n" +
"- préciser l'enjeu ;\n" +
"- expliquer comment il ou elle a obtenu les résultats ;\n" +
"- utiliser les précisions du mail candidat uniquement lorsqu'elles détaillent une expérience déjà démontrée.\n" +
"Les trois exemples doivent être cohérents avec l'étape précédente. Ils doivent être complémentaires et éviter les répétitions.\n" +
"Puis :\n" +
"1 paragraphe de conclusion de 3 lignes avec quelques éléments complémentaires trouvés dans :\n" +
"- le CV ;\n" +
"- l'entretien ;\n" +
"- éventuellement le mail candidat lorsqu'il apporte une précision sur une expérience déjà démontrée.\n" +
"Puis conclure obligatoirement par (chaque {Prénom} de ce paragraphe de conclusion remplacé par le vrai prénom ou nom du candidat, jamais par \"lui\"/\"il\"/une périphrase) :\n" +
"\"Le Taux pour une mission avec {Prénom} comprenant la flexibilité de nos contrats et le support de l'Associé est de XXXX € HT / jour, hors frais.\n\n{Prénom} est disponible immédiatement, mobile sur {lieu de mission} et très motivé par la mission.\n\nJe vous recommande de le/la rencontrer et je suis à votre disposition pour organiser un rendez-vous avec {Prénom} dès que possible.\n\nTrès cordialement,\n\nDidier\"\n\n" +
"Format : structure chaque étape avec \"## \". Les tableaux au format markdown utilisent des barres verticales \"|\", les listes \"- \". N'ajoute aucun commentaire méta, aucune balise de code.\n\n" +
"MARQUEUR TECHNIQUE OBLIGATOIRE (très important, à respecter à la lettre) : le consultant qui lit ta réponse n'affichera à son client que le contenu de l'Étape 6. Pour lui permettre de l'extraire automatiquement, insère seule sur sa ligne, juste avant le contenu de l'Étape 6 (après son titre \"## \") : @@DEBUT_MAIL_CLIENT@@ — puis juste après le contenu de l'Étape 6 (à la toute fin de ta réponse) : @@FIN_MAIL_CLIENT@@. Recopie ces deux marqueurs exactement tels quels (mêmes arobases, mêmes majuscules), sans les traduire ni les commenter, et ne les utilise nulle part ailleurs dans ta réponse.\n\n"
    );
  }

  // Combined form used whenever a single call fits the 64 KiB budget (the
  // common case) — identical in substance to the pre-split t2FinalStepsText.
  function t2FinalStepsText(){
    return t2Step4Text() + t2Step56Text();
  }

  function t2CommercialDataText(){
    var societe = els.t2Societe.value.trim();
    var contact = els.t2Contact.value.trim();
    var lieu = els.t2Lieu.value.trim();
    var taux = els.t2Taux.value.trim();
    var signataire = els.t2Signataire.value.trim();
    // These 5 fields are optional "bonus" fields filled by the consultant.
    // The mail's structure is now free-form (see t2FinalStepsText's Étape 6):
    // only the opening salutation ("Nom du contact") and the closing
    // signature ("Signataire") map to fixed, verbatim mail parts. "Société
    // cliente", "Lieu de mission" and "Taux journalier" are just facts the
    // model may weave into the free-form body wherever it naturally fits —
    // never invented when absent, never forced into a fixed sentence.
    return "DONNÉES COMMERCIALES FOURNIES PAR LE CONSULTANT (informations BONUS — à insérer dans le mail quand elles sont fournies, sans jamais les modifier ni en inventer d'autres ; quand un champ est marqué absent ci-dessous, ne l'invente pas et n'en fais pas mention) :\n" +
"- Nom du contact client : " + (contact || "(absent)") + "\n" +
"- Société cliente : " + (societe || "(absente — utiliser celle du brief de mission si elle y figure, sinon absente)") + "\n" +
"- Lieu de mission : " + (lieu || "(absent — déduire du brief de mission si possible, sinon absent)") + "\n" +
"- Taux journalier : " + (taux ? taux + " €" : "(absent)") + "\n" +
"- Signataire du mail : " + (signataire || "(absent)") + "\n\n";
  }

  function t2BuildAnalysisPrompt(){
    var instructions =
"Tu es un consultant senior en Executive Search / Management de Transition chez Valtus. Tu prépares l'ANALYSE PRÉALABLE à un dossier de présentation candidat, à partir de trois sources : (1) un dossier d'analyse consolidé, qui contient une synthèse d'entretien, une analyse externe des compétences et des informations complémentaires rédigées par le candidat ; (2) le CV complet du candidat ; (3) le brief de mission du client. Cette étape sert uniquement à identifier, sourcer et prioriser les compétences réellement démontrées — la rédaction du dossier final (compétences clés, mail) sera faite séparément, dans un second appel, à partir de ton analyse.\n\n" +
"HIÉRARCHIE DES SOURCES — respecte-la strictement.\n" +
"- Le CV est une source primaire de preuve (responsabilités, compétences, réalisations, secteurs, expériences, fonctions).\n" +
"- Dans le dossier consolidé, la synthèse d'entretien est une source primaire de preuve au même titre que le CV (compétences, responsabilités, expériences, mode de fonctionnement, style de management, motivations, manière d'intervenir).\n" +
"- Dans le dossier consolidé, la section « Analyse externe des compétences » N'EST PAS une preuve : elle ne peut ni créer une compétence ni justifier une expérience. Elle sert uniquement à challenger l'analyse initiale, améliorer les regroupements, faire apparaître des compétences sous-valorisées, améliorer la structuration du raisonnement — traite-la comme une seconde lecture des mêmes faits, jamais comme une information nouvelle.\n" +
"- Dans le dossier consolidé, la section « Informations complémentaires du candidat » est une source d'explicitation : elle ne constitue pas une preuve autonome et ne peut créer aucune compétence nouvelle. Elle sert uniquement à préciser une mission déjà présente dans le CV, détailler un périmètre, expliquer un contexte, enrichir la compréhension d'une expérience déjà démontrée, ou expliciter des responsabilités. Une information qui apparaît UNIQUEMENT dans cette section ne doit jamais être utilisée comme preuve.\n" +
"- Le brief de mission du client n'est JAMAIS une preuve : il sert uniquement à prioriser et réordonner les compétences, sélectionner les expériences les plus pertinentes, adapter le discours aux enjeux de la mission, choisir les éléments à mettre en avant. Il ne peut jamais servir à créer une compétence.\n" +
"- Règle de convergence : un même élément retrouvé dans plusieurs sources est renforcé. Un élément qui n'apparaît QUE dans l'analyse externe ou QUE dans les informations complémentaires n'est jamais considéré comme démontré.\n\n" +
"RÈGLES ABSOLUES\n" +
"1. Toute compétence, expérience, spécialisation ou affirmation doit être démontrée exclusivement par le CV ou la synthèse d'entretien. Le brief client et l'analyse externe ne démontrent jamais une compétence ; les informations complémentaires ne le font jamais seules.\n" +
"2. N'invente jamais de résultats, chiffres, tailles d'équipe, références clients, responsabilités, périmètres, zones géographiques, compétences, réalisations, niveaux hiérarchiques, succès ou gains financiers. Si une information n'est pas explicitement démontrée, ne l'utilise pas.\n" +
"3. Style : simple, direct, factuel, sans emphase, sans jargon RH, sans métaphores, sans superlatifs, sans marketing, sans interprétation excessive. Privilégie toujours les faits observables. Rédige en français.\n\n" +
"Produis ta réponse en suivant EXACTEMENT cette structure, avec un titre \"## \" pour chaque étape et aucune section en dehors de celles listées ci-dessous :\n\n" +
"## Étape 0 — Analyse croisée des sources\n" +
"Avant toute extraction, construis un tableau markdown CONCIS (colonnes : Sujet | CV | Entretien | Analyse externe | Informations complémentaires | Conclusion) identifiant uniquement les convergences et divergences qui changent réellement l'analyse (compétences sous-valorisées, expériences nécessitant un approfondissement), en une ligne très courte par sujet, sans détailler ce qui est déjà évident. N'inclus aucune affirmation non démontrée. Limite ce tableau à 15 lignes maximum : regroupe les sujets proches plutôt que d'en multiplier les lignes.\n\n" +
"## Étape 1 — Extraction des compétences\n" +
"Identifie les compétences réellement démontrées par le CV et la synthèse d'entretien ; utilise l'analyse externe pour améliorer les regroupements et les informations complémentaires pour préciser certaines expériences déjà démontrées. Présente d'abord un tableau markdown CONCIS (colonnes : Compétence | Source CV | Source entretien | Complément | Niveau de preuve), niveau de preuve parmi Fortement démontré / Démontré / Partiellement démontré, une ligne courte par compétence (pas de phrase développée dans les cellules). Ne donne les compétences finales qu'après ce tableau. Ces deux tableaux (Étapes 0 et 1) sont un raisonnement de travail qui ne sera jamais montré au client ni relu en détail : va à l'essentiel, n'y consacre pas plus de mots que nécessaire pour justifier la priorisation qui suit.\n\n" +
t2AnalysisStepsText() +
"SOURCE A — DOSSIER D'ANALYSE CANDIDAT CONSOLIDÉ (synthèse d'entretien + analyse externe des compétences + informations complémentaires du candidat) :\n\"\"\"\n" + els.t2Dossier.value.trim() + "\n\"\"\"\n\n" +
"SOURCE B — CV COMPLET DU CANDIDAT :\n\"\"\"\n" + els.t2Cv.value.trim() + "\n\"\"\"\n\n" +
"SOURCE C — BRIEF DE MISSION DU CLIENT :\n\"\"\"\n" + els.t2Brief.value.trim() + "\n\"\"\"";

    return instructions;
  }

  // Shared "rédaction finale" call: takes an already-produced analysis (from
  // either t2BuildAnalysisPrompt or t2BuildReduceAnalysisPrompt) and asks
  // only for Étape 4/5/6 + Contrôle final. Deliberately never resends the
  // raw CV/dossier/analyse externe — the analysis text is the sole input —
  // which keeps this call small and fast regardless of how big the sources
  // were.
  // `stepsText` defaults to the combined Étape 4+5+6 (t2FinalStepsText) but
  // callers can pass just t2Step4Text() or t2Step56Text() to build one half
  // of a split final call — see t2RunFinalRedaction's fallback below.
  // `skipClientContext`: the Étape 4 call alone (compétences clés) never
  // references the brief, the taux, le lieu de mission ni le signataire —
  // seul le mail (Étape 5/6) en a besoin. Omettre le brief/les données
  // commerciales dans ce cas économise des octets utiles quand le prompt
  // combiné doit être scindé (voir t2RunFinalRedaction) sans rien retirer
  // de ce que l'Étape 4 est censée produire. `customPrefsText` (les
  // préférences personnelles du consultant pour le mail, voir
  // t2GetCustomPrefsText) suit la même logique : ça ne concerne que le mail
  // (Étape 6), donc c'est aussi omis quand skipClientContext est vrai.
  function t2BuildFinalPrompt(analysisNotes, briefText, commercialDataText, stepsText, skipClientContext, customPrefsText){
    var clientContext = skipClientContext ? "" :
(commercialDataText +
"BRIEF DE MISSION DU CLIENT (pour la mission citée dans le mail) :\n\"\"\"\n" + briefText + "\n\"\"\"\n\n");
    var customPrefsBlock = (skipClientContext || !customPrefsText) ? "" :
("PRÉFÉRENCES PERSONNELLES DU CONSULTANT POUR LE MAIL CLIENT DE L'ÉTAPE 6 (optionnelles, à respecter en plus de la structure imposée dans l'Étape 6 ci-dessus — jamais à la place : ne change ni l'ordre des blocs, ni les passages à recopier mot pour mot, ni les RÈGLES ABSOLUES ci-dessus) :\n\"\"\"\n" + customPrefsText + "\n\"\"\"\n\n");
    return "Tu es un consultant senior en Executive Search / Management de Transition chez Valtus. L'analyse et la priorisation des compétences d'un candidat ont déjà été réalisées (ci-dessous) à partir de son CV et de sa synthèse d'entretien, en tenant compte du brief de mission du client. Ta seule mission maintenant est de rédiger les livrables finaux à partir de cette analyse déjà validée, SANS retourner aux sources originales (tu ne les as pas dans cet appel) et sans ajouter aucune compétence, chiffre ou référence qui n'y figure pas déjà.\n\n" +
"RÈGLES ABSOLUES\n" +
"1. N'utilise que les compétences et éléments déjà validés dans l'analyse ci-dessous.\n" +
"2. N'invente jamais de résultat, chiffre, taille d'équipe, référence client, responsabilité, périmètre ou réalisation qui n'y figure pas.\n" +
"3. Style : simple, direct, factuel, sans emphase, sans jargon RH, sans métaphores, sans superlatifs, sans marketing. Rédige en français.\n\n" +
(stepsText || t2FinalStepsText()) +
clientContext +
customPrefsBlock +
"ANALYSE ET COMPÉTENCES PRIORISÉES DU CANDIDAT (source unique et exclusive — ne retourne jamais aux sources originales) :\n\"\"\"\n" + analysisNotes + "\n\"\"\"";
  }

  // Runs the "rédaction finale" step: tries the combined Étape 4+5+6 call
  // first (the common, faster case); if that single prompt would exceed the
  // 64 KiB sample() ceiling (rich CV/brief/analysis), falls back to TWO
  // smaller calls in parallel — one for Étape 4, one for Étape 5+6 — each
  // carrying the full analysis notes, instead of failing outright. Returns
  // a promise resolving to { text } (the two texts joined back together
  // when split, so downstream extraction sees one combined document exactly
  // as if a single call had produced it).
  function t2RunFinalRedaction(analysisNotes, briefText, commercialData, onLabel, customPrefsText){
    var combinedPrompt = t2BuildFinalPrompt(analysisNotes, briefText, commercialData, undefined, false, customPrefsText);
    if (byteLength(combinedPrompt) <= T2_MAX_PROMPT_BYTES){
      return state.sample(combinedPrompt, {
        modelTier: "complex",
        cache: false,
        onText: function(update){ state.t2.outputText = update.text; t2RenderPreview(); }
      });
    }
    // Combined prompt too large: split into two calls, each still carrying
    // the full analysis notes (never truncated/summarized) so nothing that
    // was validated in the analysis gets silently lost.
    var step4Prompt = t2BuildFinalPrompt(analysisNotes, briefText, commercialData, t2Step4Text(), true);
    var step56Prompt = t2BuildFinalPrompt(analysisNotes, briefText, commercialData, t2Step56Text(), false, customPrefsText);
    if (byteLength(step4Prompt) > T2_MAX_PROMPT_BYTES || byteLength(step56Prompt) > T2_MAX_PROMPT_BYTES){
      return Promise.reject({ code: "prompt_too_large", message: "final prompt too large even split in two" });
    }
    if (onLabel) onLabel("Rédaction du dossier final (compétences clés)…");
    var partialText = "";
    return Promise.all([
      state.sample(step4Prompt, {
        modelTier: "complex",
        cache: false,
        onText: function(update){ state.t2.outputText = partialText ? partialText + "\n\n" + update.text : update.text; t2RenderPreview(); }
      }),
      (function(){
        if (onLabel) onLabel("Rédaction du dossier final (mail client)…");
        return state.sample(step56Prompt, {
          modelTier: "complex",
          cache: false,
          onText: function(update){ t2RenderPreview(); }
        });
      })()
    ]).then(function(results){
      partialText = results[0].text;
      var combinedText = results[0].text + "\n\n" + results[1].text;
      state.t2.outputText = combinedText;
      return { text: combinedText, truncated: !!(results[0].truncated || results[1].truncated) };
    });
  }

  // Batch mode (Étape 2): the dossier (Source A) is the part that can grow
  // arbitrarily large — it can even be a full Étape 1 brief. The CV, the
  // analyse externe et les informations complémentaires are kept whole in
  // every "map" call since the hierarchy-of-sources rules need them together
  // with whichever dossier section is being analyzed; only the dossier is
  // ever split, on its own "## " experience boundaries (or by paragraphs as
  // a fallback). One "reduce" call then consolidates the partial cross-
  // analyses into Étapes 0 à 3 (see t2BuildReduceAnalysisPrompt below); the
  // final rédaction (Étape 4/5/6) is a separate call via t2BuildFinalPrompt,
  // same as the non-batched path.
  // cvIdx/cvTotal are optional (default 1/1 = whole CV, unchanged behavior).
  // They're only >1 when the CV itself was too large to keep whole alongside
  // the dossier extract — see t2PlanCvSplit — in which case every dossier
  // extract is paired with every CV extract (full cross-product) so nothing
  // in the CV ever goes unchecked against a given dossier extract.
  function t2BuildMapPrompt(sectionText, idx, total, cvText, externeText, complText, cvIdx, cvTotal){
    cvIdx = cvIdx || 1; cvTotal = cvTotal || 1;
    var cvPartial = cvTotal > 1;
    var cvLabel = cvPartial
      ? ("EXTRAIT DU CV DU CANDIDAT (partie " + cvIdx + " sur " + cvTotal + " — source primaire, mais partielle : le CV a dû être découpé car trop volumineux ; les autres parties sont analysées dans d'autres lots)")
      : "CV COMPLET DU CANDIDAT (source primaire, en entier)";
    return "Tu es un consultant senior en Executive Search / Management de Transition chez Valtus. Ceci est une étape INTERMÉDIAIRE d'une analyse multi-lots" + (cvPartial ? " (le dossier ET le CV, tous deux volumineux, ont été découpés et sont analysés par paires)" : "") + " : tu reçois l'extrait " + idx + "/" + total + " du dossier d'analyse candidat (synthèse d'entretien), découpé pour l'analyse, ainsi que " + (cvPartial ? "un extrait du CV et " : "") + "les autres sources en entier. Ta seule mission ici est de produire une analyse croisée PARTIELLE, limitée aux éléments présents dans CET EXTRAIT" + (cvPartial ? " et dans l'extrait de CV fourni" : "") + " — ce n'est pas le document final, ne rédige ni synthèse finale ni mail.\n\n" +
"HIÉRARCHIE DES SOURCES — respecte-la strictement.\n" +
"- Le CV est une source primaire de preuve.\n" +
"- L'extrait du dossier ci-dessous (synthèse d'entretien) est une source primaire de preuve au même titre que le CV.\n" +
"- L'analyse externe des compétences N'EST PAS une preuve : elle sert uniquement à challenger l'analyse, jamais à créer une compétence.\n" +
"- Les informations complémentaires sont une source d'explicitation, jamais une preuve autonome ; un élément qui n'apparaît QUE dans l'une de ces deux sources n'est jamais considéré comme démontré.\n\n" +
"RÈGLES ABSOLUES\n" +
"1. Toute compétence, expérience ou affirmation retenue doit être démontrée exclusivement par le CV" + (cvPartial ? " (ou son extrait fourni ici)" : "") + " ou par l'extrait du dossier ci-dessous. N'invente jamais de résultat, chiffre, taille d'équipe, référence client, responsabilité, périmètre, compétence ou réalisation.\n" +
"2. Style : factuel, sans emphase, en français.\n\n" +
"TRAVAIL DEMANDÉ SUR CET EXTRAIT\n" +
"1. Produis un tableau markdown (colonnes : Compétence | Source CV | Source entretien (extrait " + idx + "/" + total + ") | Analyse externe / Complément | Niveau de preuve parmi Fortement démontré / Démontré / Partiellement démontré) listant les compétences, responsabilités et réalisations démontrées par CET EXTRAIT, recoupées avec " + (cvPartial ? "l'extrait de CV" : "le CV complet") + " fourni ci-dessous.\n" +
"2. Note brièvement les convergences ou divergences avec l'analyse externe et les informations complémentaires, sans jamais les traiter comme une preuve autonome.\n" +
(cvPartial ? "3. Si un élément de cet extrait de dossier ne trouve aucune correspondance dans cet extrait de CV, n'écris pas qu'il est absent du CV : dis seulement qu'il n'est pas corroboré par CET extrait de CV — une autre partie du CV est analysée séparément et pourra le confirmer.\n" : "") +
(cvPartial ? "4" : "3") + ". N'écris aucune conclusion générale, aucune sélection finale, aucun mail : ce n'est qu'une étape de collecte structurée.\n\n" +
"FORMAT : le tableau demandé, puis si utile une courte liste de points de convergence/divergence. Aucun commentaire méta, aucune balise de code.\n\n" +
cvLabel + " :\n\"\"\"\n" + cvText + "\n\"\"\"\n\n" +
"ANALYSE EXTERNE DES COMPÉTENCES (jamais une preuve, en entier) :\n\"\"\"\n" + (externeText || "(non fournie)") + "\n\"\"\"\n\n" +
"INFORMATIONS COMPLÉMENTAIRES DU CANDIDAT (explicitation seulement, en entier) :\n\"\"\"\n" + (complText || "(non fournies)") + "\n\"\"\"\n\n" +
"EXTRAIT DU DOSSIER — SYNTHÈSE D'ENTRETIEN (partie " + idx + " sur " + total + ") :\n\"\"\"\n" + sectionText + "\n\"\"\"";
  }

  function t2BuildReduceAnalysisPrompt(mapOutputsJoined, briefClientText){
    return "Tu es un consultant senior en Executive Search / Management de Transition chez Valtus. Le dossier d'analyse candidat a déjà été analysé PAR LOTS : tu reçois ci-dessous les analyses croisées partielles déjà produites pour chaque lot (compétences, source CV/entretien, niveau de preuve, convergences/divergences avec l'analyse externe et les informations complémentaires). Ta mission est de les consolider en une ANALYSE PRÉALABLE unique et priorisée, SANS retourner aux sources originales (tu ne les as pas dans cet appel) et sans ajouter d'information qui n'y figure pas déjà. Cette étape ne produit pas le dossier final : la rédaction (compétences clés, mail) sera faite séparément, dans un second appel, à partir de ta consolidation.\n\n" +
"RÈGLES ABSOLUES\n" +
"1. N'utilise que les éléments déjà validés dans les analyses par lots ci-dessous.\n" +
"2. N'invente rien, n'ajoute aucun résultat, chiffre, référence client ou compétence qui n'y figure pas.\n" +
"3. Si un même élément apparaît dans plusieurs lots, ne le restitue qu'une seule fois — retiens le niveau de preuve le plus élevé mentionné.\n" +
"4. Style : simple, direct, factuel, sans emphase, sans jargon RH, sans superlatifs. Rédige en français.\n\n" +
"Produis ta réponse en suivant EXACTEMENT cette structure, avec un titre \"## \" pour chaque étape et aucune section en dehors de celles listées ci-dessous :\n\n" +
"## Étape 0 — Analyse croisée des sources\n" +
"Consolide les analyses par lots en un seul tableau markdown CONCIS (colonnes : Sujet | CV | Entretien | Analyse externe | Informations complémentaires | Conclusion), en fusionnant les lignes qui décrivent le même sujet et en une ligne très courte par sujet. Exclus toute affirmation non démontrée. Limite ce tableau à 15 lignes maximum : regroupe les sujets proches plutôt que d'en multiplier les lignes.\n\n" +
"## Étape 1 — Extraction des compétences\n" +
"Consolide les tableaux de compétences par lots en un seul tableau markdown CONCIS (colonnes : Compétence | Source CV | Source entretien | Complément | Niveau de preuve), en fusionnant les doublons, une ligne courte par compétence (pas de phrase développée dans les cellules). Ne donne les compétences finales qu'après ce tableau. Ces deux tableaux (Étapes 0 et 1) sont un raisonnement de travail qui ne sera jamais montré au client ni relu en détail : va à l'essentiel, n'y consacre pas plus de mots que nécessaire pour justifier la priorisation qui suit.\n\n" +
t2AnalysisStepsText() +
"BRIEF DE MISSION DU CLIENT :\n\"\"\"\n" + briefClientText + "\n\"\"\"\n\n" +
"ANALYSES PARTIELLES PAR LOTS À CONSOLIDER (source unique et exclusive pour les Étapes 0 et 1 — ne retourne jamais aux sources originales) :\n\"\"\"\n" + mapOutputsJoined + "\n\"\"\"";
  }

  function t2RefreshGenerateState(){
    var ready = els.t2Dossier.value.trim().length > 0 && els.t2Cv.value.trim().length > 0;
    var promptBytes = byteLength(t2BuildAnalysisPrompt());
    var willBatch = promptBytes > T2_MAX_PROMPT_BYTES;
    var budget = promptBytes.toLocaleString("fr-FR") + " / " + T2_MAX_PROMPT_BYTES.toLocaleString("fr-FR") + " caractères";
    var canGenerate = ready && !state.t2.generating && !!state.sample;
    els.t2GenerateBtn.disabled = !canGenerate;
    if (!state.sample){
      els.t2GenerateHint.textContent = "Génération indisponible dans cette vue.";
      els.t2GenerateHint.classList.add("warn");
    } else if (!ready){
      els.t2GenerateHint.textContent = "Le dossier candidat (Étape 1) et le CV sont requis. · Budget LLM : " + budget;
      els.t2GenerateHint.classList.remove("warn");
    } else if (willBatch){
      var dossierSplit = t2SplitDossierForBatches(els.t2Dossier.value);
      var cvPlan = t2PlanCvSplit(dossierSplit, els.t2Cv.value.trim());
      if (!cvPlan){
        els.t2GenerateHint.textContent = "L'analyse externe et les informations complémentaires sont à elles seules trop volumineuses, même en découpant le CV et le dossier. Raccourcissez-les. · Budget LLM : " + budget;
        els.t2GenerateHint.classList.add("warn");
      } else if (cvPlan.cvChunks.length > 1){
        var callCount = dossierSplit.sections.length * cvPlan.cvChunks.length;
        els.t2GenerateHint.textContent = "Sources volumineuses : CV découpé en " + cvPlan.cvChunks.length + " extraits et dossier en " + dossierSplit.sections.length + " lots (" + callCount + " analyses croisées), puis fusion en un seul dossier final. · Budget LLM : " + budget;
        els.t2GenerateHint.classList.remove("warn");
      } else {
        els.t2GenerateHint.textContent = "Sources volumineuses : analyse automatique du dossier en " + dossierSplit.sections.length + " lots, puis fusion en un seul dossier final. · Budget LLM : " + budget;
        els.t2GenerateHint.classList.remove("warn");
      }
    } else {
      els.t2GenerateHint.textContent = "Seuls le CV et la synthèse d'entretien démontrent une compétence. · Budget LLM : " + budget;
      els.t2GenerateHint.classList.remove("warn");
    }
  }

  // --- Primary extraction: sentinel markers ------------------------------
  // t2FinalStepsText() asks the model to wrap Étape 4 and Étape 6 in exact
  // literal markers (@@DEBUT_…@@ / @@FIN_…@@). Reading those markers off the
  // RAW text (before any markdown parsing) is far more robust than matching
  // on heading text/casing/accents, which real model output doesn't always
  // reproduce exactly.
  var T2_TAG_COMP_START = "@@DEBUT_COMPETENCES_CLES@@";
  var T2_TAG_COMP_END = "@@FIN_COMPETENCES_CLES@@";
  var T2_TAG_MAIL_START = "@@DEBUT_MAIL_CLIENT@@";
  var T2_TAG_MAIL_END = "@@FIN_MAIL_CLIENT@@";

  function t2ExtractBetween(text, startTag, endTag){
    var t = String(text || "");
    var startIdx = t.indexOf(startTag);
    if (startIdx === -1) return "";
    startIdx += startTag.length;
    var endIdx = t.indexOf(endTag, startIdx);
    if (endIdx === -1) endIdx = t.length;
    return t.slice(startIdx, endIdx).trim();
  }

  function t2ExtractFinalSections(rawText){
    return {
      competences: t2ExtractBetween(rawText, T2_TAG_COMP_START, T2_TAG_COMP_END),
      mail: t2ExtractBetween(rawText, T2_TAG_MAIL_START, T2_TAG_MAIL_END)
    };
  }

  // --- Fallback extraction: heading text ----------------------------------
  // Used only when the sentinel markers are missing from the model's answer
  // (it ignored the technical instruction). Accent/case-insensitive so
  // "Étape 4", "ETAPE 4" or "etape 4" all match.
  function foldAccents(s){
    return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "");
  }

  // The model sometimes fakes a heading with a fully-bold paragraph
  // ("**Livrable 1 — Mail…**") instead of a "## " markdown heading, despite
  // being told to use the latter. Recognize that too, so the fallbacks below
  // still find the mail/competencies sections in that case. Returns the
  // heading's text (real or bold-faked), or null if b isn't heading-like.
  function t2HeadingText(b){
    if (b.style === "h1" || b.style === "h2") return b.text;
    if (b.style === "p"){
      var m = /^\*\*(.+)\*\*$/.exec(String(b.text || "").trim());
      if (m) return m[1];
    }
    return null;
  }

  // Classifies a block for section-boundary purposes: "skills" or "mail"
  // when it names one of our two target sections (Étape 4/6, real or
  // bold-faked); "other-heading" for a genuine "## " heading that names
  // something else (Étape 5, Contrôle final, …) — reliable enough to always
  // end the current section; null for ordinary content, INCLUDING an
  // unrelated bold pseudo-heading like "**{Nom du candidat}**" — those are
  // too unreliable to trust as boundaries and would otherwise cut a section
  // short mid-way.
  function t2SectionKind(b){
    var headingText = t2HeadingText(b);
    if (headingText === null) return null;
    var folded = foldAccents(headingText);
    if (/etape\s*4/i.test(folded) || /competences?/i.test(folded)) return "skills";
    if (/etape\s*6/i.test(folded) || /mail/i.test(headingText)) return "mail";
    if (b.style === "h1" || b.style === "h2") return "other-heading";
    return null;
  }

  function t2ExtractEmail(blocks){
    var start = -1;
    for (var i = 0; i < blocks.length; i++){
      if (t2SectionKind(blocks[i]) === "mail"){ start = i; break; }
    }
    if (start === -1) return "";
    var lines = [];
    for (var j = start + 1; j < blocks.length; j++){
      var kind = t2SectionKind(blocks[j]);
      if (kind === "mail" || kind === "skills" || kind === "other-heading") break;
      var b2 = blocks[j];
      if (b2.style === "bullet") lines.push("- " + b2.text);
      else if (b2.style === "table") lines.push(tableBlockToPlainText(b2));
      else if (b2.text) lines.push(b2.text);
    }
    return lines.join("\n\n").trim();
  }

  // The LLM still runs the full Étapes 0 à 6 + Contrôle final reasoning (the
  // hierarchy-of-sources analysis, the cross-reference tables, the profile
  // reconstruction, the prioritization) — that scaffolding is what keeps the
  // final result traceable and grounded in the CV/entretien. But the client
  // only wants the finished deliverable: the key-competencies list (Étape 4)
  // and a ready-to-copy email (Étape 6). This keeps the internal reasoning
  // out of what's shown on screen and exported to Word.
  function t2FilterFinalBlocks(blocks){
    var kept = [];
    var mode = null;
    blocks.forEach(function(b){
      var kind = t2SectionKind(b);
      if (kind === "skills"){ mode = "skills"; kept.push({ style: "h1", text: "Compétences clés" }); return; }
      if (kind === "mail"){ mode = "mail"; kept.push({ style: "h1", text: "Mail client (prêt à copier)", calloutStart: true }); return; }
      if (kind === "other-heading"){ mode = null; return; }
      if (mode) kept.push(b);
    });
    return kept;
  }

  // Splits the heading-filtered blocks (from t2FilterFinalBlocks) back into
  // their two sections' content, WITHOUT the two h1 markers themselves —
  // used as an independent per-section fallback so that if the sentinel
  // markers are present for only ONE of the two sections, the other one
  // still gets a chance via heading-based extraction instead of being
  // silently dropped. The client always needs both: compétences clés AND
  // the mail.
  function t2SplitFilteredBySection(filteredBlocks){
    var competences = [], mail = [], mode = null;
    filteredBlocks.forEach(function(b){
      if (b.style === "h1" && b.text === "Compétences clés"){ mode = "skills"; return; }
      if (b.style === "h1" && b.text === "Mail client (prêt à copier)"){ mode = "mail"; return; }
      if (mode === "skills") competences.push(b);
      else if (mode === "mail") mail.push(b);
    });
    return { competencesBlocks: competences, mailBlocks: mail };
  }

  // Mail text used by both the "Copier le mail" button and its enabled state
  // — sentinel extraction first, falling back to heading-based extraction on
  // the full raw output so the button still works if markers are missing.
  function t2GetMailText(){
    var sentinelMail = t2ExtractBetween(state.t2.outputText, T2_TAG_MAIL_START, T2_TAG_MAIL_END);
    if (sentinelMail) return sentinelMail;
    return t2ExtractEmail(parseBrief(state.t2.outputText));
  }

  // Safety net: the Étape 6 prompt asks for a short, bullet-free, prose-only
  // mail with nothing else mixed in — but the model doesn't always comply.
  // Rather than silently showing (and letting the consultant send) a mail
  // that violates those constraints, detect the violation on the ALREADY
  // extracted mail blocks/text and surface a visible warning in the preview.
  // This never edits the model's text — it only warns, so the consultant
  // always sees the model's real output and decides what to do about it.
  var T2_MAIL_WORD_LIMIT = 300;
  var T2_MAIL_WORD_WARN_THRESHOLD = 340; // small grace margin above the target before warning, to avoid false positives on borderline-but-acceptable mails

  function t2CheckMailQuality(mailBlocks, mailText){
    var issues = [];
    var hasBullets = mailBlocks.some(function(b){ return b.style === "bullet"; });
    var hasTable = mailBlocks.some(function(b){ return b.style === "table"; });
    if (hasBullets) issues.push("contient des puces");
    if (hasTable) issues.push("contient un tableau");
    var wordCount = (mailText || "").trim().split(/\s+/).filter(Boolean).length;
    if (wordCount > T2_MAIL_WORD_WARN_THRESHOLD){
      issues.push(wordCount.toLocaleString("fr-FR") + " mots, au-delà des " + T2_MAIL_WORD_LIMIT + " mots visés");
    }
    if (!issues.length) return null;
    return "Ce mail " + issues.join(", ") + " — ce qui ne respecte pas le format demandé (mail court, sans puces). Vérifiez-le avant de l'envoyer, ou régénérez.";
  }

  function t2DocumentBlocks(){
    var blocks = [];
    blocks.push({ style: "title", text: "Dossier de présentation candidat" });
    var meta = [];
    var candidate = els.candidate.value.trim();
    if (candidate) meta.push("Candidat : " + candidate);
    var societe = els.t2Societe.value.trim();
    if (societe) meta.push("Client : " + societe);
    meta.push("Préparé le : " + new Date().toLocaleDateString("fr-FR", { year: "numeric", month: "long", day: "numeric" }));
    blocks.push({ style: "meta", text: meta.join("   ·   ") });

    if (state.t2.outputText.trim()){
      // The client always needs BOTH deliverables — compétences clés AND
      // mail. Each is extracted independently (sentinel markers first, then
      // heading-based fallback) so that if only one of the two mechanisms
      // works for a given section, the other section isn't dragged down
      // with it. Only if NEITHER section can be isolated at all do we fall
      // back to showing the model's full raw output.
      var sections = t2ExtractFinalSections(state.t2.outputText);
      var parsed = parseBrief(state.t2.outputText);
      var filtered = t2FilterFinalBlocks(parsed);
      var bySection = t2SplitFilteredBySection(filtered);

      var competencesBlocks = sections.competences ? parseBrief(sections.competences) : bySection.competencesBlocks;
      var mailBlocks = sections.mail ? parseBrief(sections.mail) : bySection.mailBlocks;

      if (competencesBlocks.length || mailBlocks.length){
        if (competencesBlocks.length){
          blocks.push({ style: "h1", text: "Compétences clés" });
          blocks = blocks.concat(competencesBlocks);
        }
        if (mailBlocks.length){
          blocks.push({ style: "h1", text: "Mail client (prêt à copier)", calloutStart: true });
          var mailWarning = t2CheckMailQuality(mailBlocks, t2GetMailText());
          if (mailWarning) blocks.push({ style: "warning", text: mailWarning });
          blocks = blocks.concat(mailBlocks);
        }
      } else {
        // Last resort: neither section could be isolated by any mechanism
        // — show everything rather than nothing.
        blocks = blocks.concat(parsed);
      }
    } else {
      blocks.push({ style: "placeholder", text: "Le dossier n'a pas encore été généré." });
    }
    return blocks;
  }

  function t2BlocksToHtml(blocks){
    var html = "";
    var inCallout = false;
    blocks.forEach(function(b){
      if (b.calloutStart){
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

  function t2RenderPreview(){
    els.t2DocPage.innerHTML = t2BlocksToHtml(t2DocumentBlocks());
    t2RefreshActionState();
  }

  function t2RefreshActionState(){
    els.t2DownloadBtn.disabled = !state.downloads || !state.t2.hasGenerated || state.t2.generating;
    els.t2CopyBtn.disabled = !state.t2.hasGenerated || state.t2.generating || !t2GetMailText();
    els.t2SendBtn.disabled = !state.t2.hasGenerated || state.t2.generating || !t2GetMailText();
    // Visible as soon as a dossier has been generated (even while a
    // regeneration is running) so the consultant sees the reminder before
    // using any of the export/copy/send actions above — never shown before
    // there is anything to review yet.
    els.t2AiWarning.hidden = !state.t2.hasGenerated;
  }

  // Shared success/failure handlers for the final "rédaction" call, used by
  // both t2RunGenerate and t2RunGenerateBatched.
  function t2OnFinalSuccess(result){
    state.t2.outputText = result.text;
    state.t2.hasGenerated = true;
    state.t2.generating = false;
    els.t2GenerateBtn.removeAttribute("data-busy");
    els.t2GenerateLabel.textContent = "Régénérer le dossier";
    if (result.truncated){
      showToast("La réponse a été interrompue par la limite de longueur. Le dossier peut être incomplet.", "error");
    }
    t2RenderPreview();
    t2RefreshGenerateState();
  }

  function t2OnFailure(failure){
    state.t2.generating = false;
    els.t2GenerateBtn.removeAttribute("data-busy");
    els.t2GenerateLabel.textContent = state.t2.hasGenerated ? "Régénérer le dossier" : "Générer le dossier";
    var err = failure && failure.err;
    var phase = failure && failure.phase;
    if (err && err.text && phase === "redaction"){ state.t2.outputText = err.text; t2RenderPreview(); }
    var msg = (err && SAMPLE_ERROR_COPY[err.code]) || "Une erreur est survenue pendant la génération.";
    if (err && err.code === "prompt_too_large"){
      msg = phase === "redaction"
        ? "L'analyse est trop volumineuse pour la rédaction finale, même en la scindant en deux appels. Réessayez avec un CV ou un brief plus courts."
        : "Même après découpage, les notes extraites restent trop volumineuses pour la consolidation. Réessayez avec un CV un peu plus court.";
    }
    if (phase === "collecte") msg = "L'analyse par lots a échoué : " + msg;
    else if (phase === "analyse") msg = "L'analyse des sources a échoué : " + msg;
    else if (phase === "redaction") msg = "La rédaction finale a échoué : " + msg;
    showToast(msg, "error");
    t2RefreshGenerateState();
  }

  function t2RunGenerate(){
    if (!state.sample || state.t2.generating) return;
    if (!els.t2Dossier.value.trim() || !els.t2Cv.value.trim()) return;

    var analysisPrompt = t2BuildAnalysisPrompt();
    if (byteLength(analysisPrompt) > T2_MAX_PROMPT_BYTES){
      t2RunGenerateBatched();
      return;
    }

    var brief = els.t2Brief.value.trim();
    var commercialData = t2CommercialDataText();

    state.t2.generating = true;
    els.t2GenerateBtn.setAttribute("data-busy", "true");
    els.t2GenerateBtn.disabled = true;
    els.t2GenerateLabel.textContent = "Analyse en cours…";
    t2RefreshActionState();

    state.sample(analysisPrompt, { modelTier: "complex", cache: false })
      .catch(function(err){ return Promise.reject({ phase: "analyse", err: err }); })
      .then(function(analysisResult){
        els.t2GenerateLabel.textContent = "Rédaction du dossier final…";
        return t2RunFinalRedaction(analysisResult.text, brief, commercialData, function(label){ els.t2GenerateLabel.textContent = label; }, t2GetCustomPrefsText())
          .catch(function(err){ return Promise.reject({ phase: "redaction", err: err }); });
      })
      .then(t2OnFinalSuccess)
      .catch(t2OnFailure);
  }

  // Decides whether the CV also needs to be split into chunks, in addition
  // to the dossier (which t2SplitDossierForBatches always splits when
  // needed). Most CVs stay whole — this only kicks in when a CV, combined
  // with the fixed per-call overhead (instructions + analyse externe +
  // informations complémentaires) and the largest dossier extract, doesn't
  // fit the 64 KiB budget on its own: long-career executive CVs especially.
  // Verified against the real built prompt (byteLength), not estimated, and
  // shrinks the CV chunk size step by step until a chunk actually fits.
  // Returns { cvChunks } (cvChunks.length === 1 means "kept whole", the
  // common case) or null if even the smallest chunking still doesn't fit
  // (meaning the fixed overhead itself — analyse externe / informations
  // complémentaires — is what's too large, not the CV).
  function t2PlanCvSplit(dossierSplit, cvText){
    var sections = dossierSplit.sections;
    var largestSection = sections.reduce(function(a, b){ return b.length > a.length ? b : a; }, sections[0] || "");
    var wholeCvPrompt = t2BuildMapPrompt(largestSection, 1, sections.length, cvText, dossierSplit.externe, dossierSplit.complementaire, 1, 1);
    if (byteLength(wholeCvPrompt) <= T2_MAX_PROMPT_BYTES){
      return { cvChunks: [cvText] };
    }
    var targets = [24000, 16000, 10000, 6000, 3500];
    for (var t = 0; t < targets.length; t++){
      var cvChunks = t2SplitCvForBatches(cvText, targets[t]);
      var largestCvChunk = cvChunks.reduce(function(a, b){ return b.length > a.length ? b : a; }, cvChunks[0] || "");
      var testPrompt = t2BuildMapPrompt(largestSection, 1, sections.length, largestCvChunk, dossierSplit.externe, dossierSplit.complementaire, 1, cvChunks.length);
      if (byteLength(testPrompt) <= T2_MAX_PROMPT_BYTES){
        return { cvChunks: cvChunks };
      }
    }
    return null;
  }

  // Automatic batch (map-reduce) path for oversized Étape 2 sources. See the
  // comment above t2BuildMapPrompt for the splitting strategy. Just like the
  // non-batched path, the actual rédaction (Étape 4/5/6) is always a
  // separate, small final call — see t2BuildFinalPrompt.
  function t2RunGenerateBatched(){
    var cv = els.t2Cv.value.trim();
    var brief = els.t2Brief.value.trim();
    var commercialData = t2CommercialDataText();
    var dossierSplit = t2SplitDossierForBatches(els.t2Dossier.value.trim());
    var sections = dossierSplit.sections;
    var total = sections.length;

    // The CV is kept whole in every map call whenever possible. Only when it
    // alone (with the fixed overhead and the largest dossier extract) is too
    // large — long-career executive CVs — is it also split, in which case
    // every dossier extract is paired with every CV extract (full cross-
    // product) so no part of either source ever goes unchecked against the
    // other. If even that doesn't fit, the fixed overhead itself (analyse
    // externe / informations complémentaires) is what's too large.
    var cvPlan = t2PlanCvSplit(dossierSplit, cv);
    if (!cvPlan){
      showToast("L'analyse externe et les informations complémentaires sont à elles seules trop volumineuses pour être analysées, même en découpant le CV et le dossier. Raccourcissez-les.", "error");
      return;
    }
    var cvChunks = cvPlan.cvChunks;
    var totalCv = cvChunks.length;
    var totalCalls = total * totalCv;

    state.t2.generating = true;
    els.t2GenerateBtn.setAttribute("data-busy", "true");
    els.t2GenerateBtn.disabled = true;
    els.t2GenerateLabel.textContent = totalCv > 1
      ? ("Analyse de " + totalCalls + " extraits (CV × dossier)…")
      : ("Analyse de " + total + " extraits…");
    t2RefreshActionState();

    var mapCalls = [];
    sections.forEach(function(section, i){
      cvChunks.forEach(function(cvChunk, k){
        var mapPrompt = t2BuildMapPrompt(section, i + 1, total, cvChunk, dossierSplit.externe, dossierSplit.complementaire, k + 1, totalCv);
        mapCalls.push(state.sample(mapPrompt, { modelTier: "complex", cache: false }));
      });
    });

    Promise.all(mapCalls)
      .catch(function(err){ return Promise.reject({ phase: "collecte", err: err }); })
      .then(function(results){
        var joined = results.map(function(r, i){
          return "--- Lot " + (i + 1) + "/" + results.length + " ---\n" + r.text;
        }).join("\n\n");
        els.t2GenerateLabel.textContent = "Consolidation de l'analyse…";
        var reduceAnalysisPrompt = t2BuildReduceAnalysisPrompt(joined, brief);
        if (byteLength(reduceAnalysisPrompt) > T2_MAX_PROMPT_BYTES){
          return Promise.reject({ phase: "analyse", err: { code: "prompt_too_large", message: "consolidated notes too large" } });
        }
        return state.sample(reduceAnalysisPrompt, { modelTier: "complex", cache: false })
          .catch(function(err){ return Promise.reject({ phase: "analyse", err: err }); });
      })
      .then(function(analysisResult){
        els.t2GenerateLabel.textContent = "Rédaction du dossier final…";
        return t2RunFinalRedaction(analysisResult.text, brief, commercialData, function(label){ els.t2GenerateLabel.textContent = label; }, t2GetCustomPrefsText())
          .catch(function(err){ return Promise.reject({ phase: "redaction", err: err }); });
      })
      .then(t2OnFinalSuccess)
      .catch(t2OnFailure);
  }
  els.t2GenerateBtn.addEventListener("click", t2RunGenerate);

  function t2SlugFilename(){
    var candidate = els.candidate.value.trim();
    var base = "Dossier_Presentation_Candidat" + (candidate ? "_" + candidate.replace(/\s+/g, "_") : "");
    base = base.replace(/[^A-Za-z0-9_\-]/g, "");
    return (base || "Dossier_Presentation_Candidat") + ".docx";
  }

  function t2RunDownload(){
    if (!state.downloads || typeof JSZip === "undefined") return;
    els.t2DownloadBtn.disabled = true;
    var previousLabel = els.t2DownloadBtn.textContent;
    els.t2DownloadBtn.textContent = "Préparation du fichier…";

    var zip = new JSZip();
    zip.file("[Content_Types].xml", CT_XML);
    zip.folder("_rels").file(".rels", RELS_XML);
    var wordFolder = zip.folder("word");
    wordFolder.file("document.xml", buildDocumentXml(t2DocumentBlocks()));
    wordFolder.file("styles.xml", STYLES_XML);
    wordFolder.folder("_rels").file("document.xml.rels", DOC_RELS_XML);

    zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" })
      .then(function(blob){
        return state.downloads.save({ filename: t2SlugFilename(), data: blob });
      })
      .then(function(res){
        showToast(res.status === "saved" ? "Document Word enregistré." : "Document transmis.", "ok");
      })
      .catch(function(err){
        if (err && err.code === "declined"){ /* silent */ }
        else showToast("Impossible de générer le fichier Word. Réessayez.", "error");
      })
      .finally(function(){
        els.t2DownloadBtn.textContent = previousLabel;
        t2RefreshActionState();
      });
  }
  els.t2DownloadBtn.addEventListener("click", t2RunDownload);

  function t2RunCopyEmail(){
    var email = t2GetMailText();
    if (!email) return;
    if (navigator.clipboard && navigator.clipboard.writeText){
      navigator.clipboard.writeText(email)
        .then(function(){ showToast("Mail client copié dans le presse-papiers.", "ok"); })
        .catch(function(){ showToast("Impossible de copier automatiquement — sélectionnez le texte dans l'aperçu.", "error"); });
    } else {
      showToast("Copie automatique indisponible — sélectionnez le texte dans l'aperçu.", "error");
    }
  }
  els.t2CopyBtn.addEventListener("click", t2RunCopyEmail);

  // The Étape 6 prompt now generates a leading "Objet : …" line as the
  // first line of the mail (see t2FinalStepsText). t2GetMailText() keeps
  // returning the full text INCLUDING that line — useful for "Copier le
  // mail" (the consultant may want it visible when pasting into an
  // already-open draft). For "Envoyer par email", though, a real mailto:
  // subject field exists, so we split the leading "Objet : …" line off
  // and use it there instead of leaving it as a redundant first line of
  // the body. If the model didn't produce that line (older/partial
  // output), the subject is simply left blank rather than guessed.
  function t2SplitSubjectAndBody(fullText){
    var m = /^\s*Objet\s*:\s*(.+?)\s*\n+([\s\S]*)$/.exec(fullText);
    if (m) return { subject: m[1].trim(), body: m[2].trim() };
    return { subject: "", body: fullText };
  }

  // mailto: links have no universally-safe length limit (practical caps
  // vary by OS/client, roughly 1,800–2,000 characters is the safest common
  // floor). Above that, some clients silently truncate the body instead of
  // erroring — which would be exactly the kind of silent information loss
  // this tool must never risk. So above the threshold we don't attempt
  // mailto at all: we copy the mail to the clipboard instead and tell the
  // user to paste it, rather than gamble on a partial mailto body.
  var MAILTO_SAFE_CHARS = 1800;

  function t2RunSendEmail(){
    var email = t2GetMailText();
    if (!email) return;
    var split = t2SplitSubjectAndBody(email);
    var mailtoUrl = "mailto:?subject=" + encodeURIComponent(split.subject) + "&body=" + encodeURIComponent(split.body);
    if (mailtoUrl.length > MAILTO_SAFE_CHARS){
      if (navigator.clipboard && navigator.clipboard.writeText){
        navigator.clipboard.writeText(email).then(function(){
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
  els.t2SendBtn.addEventListener("click", t2RunSendEmail);

  // ===============================================================
  // Adaptateurs de plateforme
  // ===============================================================
  // Les deux seuls points de contact du front avec l'extérieur : enregistrer
  // un fichier, et faire rédiger un texte. Tout le reste de ce fichier est
  // autonome. Le code appelant continue de passer par state.downloads.save()
  // et state.sample(), dont les signatures sont conservées à l'identique.

  function insideAppsmith(){
    return !!(window.appsmith
      && typeof window.appsmith.updateModel === "function"
      && typeof window.appsmith.triggerEvent === "function"
      && typeof window.appsmith.onModelChange === "function");
  }

  // ---------------------------------------------------------------
  // 1. Enregistrement du fichier Word
  // ---------------------------------------------------------------
  // Dans Appsmith, le widget « Custom » s'exécute dans une iframe sandboxée
  // sans le drapeau allow-downloads : un <a download> y est bloqué par le
  // navigateur SANS lever d'erreur — l'échec serait donc silencieux, ce qui
  // est exactement ce qu'il ne faut pas ici. On ne tente donc pas le
  // téléchargement direct « pour voir » : la voie est choisie en fonction de
  // l'hôte. Sous Appsmith, le fichier est publié dans le modèle du widget,
  // puis l'événement « onDownloadDocx » demande à l'application de le
  // télécharger via sa fonction download(). Hors Appsmith (ouverture directe
  // de dist/index.html), le lien de téléchargement classique suffit.

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

  // ---------------------------------------------------------------
  // 2. Moteur de génération (LLM) — point de branchement unique
  // ---------------------------------------------------------------
  // AUCUN APPEL N'EST ÉMIS AUJOURD'HUI. Tant qu'aucun moteur n'est fourni, la
  // génération est simulée localement : toute la chaîne (aperçu, export Word,
  // enchaînement Étape 1 -> Étape 2) reste utilisable, et le texte produit
  // s'annonce lui-même comme un exemple.
  //
  // Pour brancher un vrai moteur, définir window.VALTUS_LLM avant ce script :
  //
  //   window.VALTUS_LLM = {
  //     complete: function(prompt, options){
  //       // options.onText({ text }) peut être appelé pendant le streaming.
  //       // Rejeter avec { code, message } ; les codes reconnus sont ceux de
  //       // SAMPLE_ERROR_COPY (rate_limited, prompt_too_large, refused…).
  //       return Promise.resolve({ text: "…", truncated: false });
  //     }
  //   };
  //
  // Voir custom-widget/README.md pour un exemple de branchement sur une API
  // Appsmith. Les limites de taille (64 Kio par appel, découpage automatique
  // en lots au-delà) sont celles du moteur d'origine : à réajuster via
  // MAX_PROMPT_BYTES / T2_MAX_PROMPT_BYTES selon le moteur branché.

  var LLM_STUB_DELAY_MS = 700;
  var LLM_STUB_NOTICE = "⚠ CONTENU DE DÉMONSTRATION — aucun moteur de génération n'est branché. Ce texte ne provient pas des sources saisies et ne doit en aucun cas être transmis à un client.";

  function llmEngine(){
    var engine = window.VALTUS_LLM;
    return (engine && typeof engine.complete === "function") ? engine : null;
  }

  function llmStubText(prompt){
    // Rédaction finale de l'Étape 2 : l'aperçu n'affiche que ce qui se trouve
    // entre les marqueurs sentinelles, le texte d'exemple doit donc les poser.
    if (prompt.indexOf(T2_TAG_COMP_START) !== -1){
      return T2_TAG_COMP_START + "\n" +
        "- " + LLM_STUB_NOTICE + "\n" +
        "- Compétence clé d'exemple n° 1 — remplacée par la sortie du moteur une fois celui-ci branché.\n" +
        "- Compétence clé d'exemple n° 2 — remplacée par la sortie du moteur une fois celui-ci branché.\n" +
        T2_TAG_COMP_END + "\n\n" +
        T2_TAG_MAIL_START + "\n" +
        "Objet : [DÉMONSTRATION] Présentation de candidature\n\n" +
        "Bonjour,\n\n" +
        LLM_STUB_NOTICE + "\n\n" +
        "Le corps du mail client s'affichera ici une fois le moteur de génération branché.\n\n" +
        "Bien à vous,\n" +
        T2_TAG_MAIL_END;
    }
    // Tous les autres appels (Étape 1, collecte et consolidation par lots)
    // partagent le format « ## section » attendu par le parseur.
    return "## " + LLM_STUB_NOTICE + "\n\n" +
      "Ce bloc occupe la place du dossier qui sera rédigé à partir des sources saisies une fois le moteur de génération branché. Il permet de vérifier dès maintenant l'aperçu, la structure du document et l'export Word.\n\n" +
      "## Expérience d'exemple — Société X (20XX-20XX)\n\n" +
      "Contexte, périmètre, responsabilités, réalisations et transformations conduites apparaîtront ici, une section par expérience significative du candidat.\n\n" +
      "## Éléments transverses\n\n" +
      "- Mode de fonctionnement, qualités revendiquées et disponibilité seront regroupés dans cette dernière section.";
  }

  function llmStub(prompt, options){
    var text = llmStubText(String(prompt || ""));
    return new Promise(function(resolve){
      setTimeout(function(){
        if (options && typeof options.onText === "function") options.onText({ text: text });
        resolve({ text: text, truncated: false });
      }, LLM_STUB_DELAY_MS);
    });
  }

  function llmComplete(prompt, options){
    var engine = llmEngine();
    return engine ? engine.complete(prompt, options || {}) : llmStub(prompt, options);
  }

  // ---------------------------------------------------------------
  // Démarrage
  // ---------------------------------------------------------------
  state.downloads = { save: function(request){ return saveFile(request.filename, request.data); } };
  state.sample = function(prompt, options){ return llmComplete(prompt, options); };

  var isStubMode = !llmEngine();
  els.demoBanner.hidden = !isStubMode;

  updateCounts();
  renderPreview();
  t2UpdateCounts();
  t2RenderPreview();
  setStatus(isStubMode ? "idle" : "ready", isStubMode ? "Mode démonstration" : "Prêt");

  refreshGenerateState();
  refreshDownloadState();
  t2RefreshGenerateState();
  t2RefreshActionState();
})();
