# Front « Générateur de dossier candidat »

Un écran, **un seul traitement**. Toutes les sources (retranscription
d'entretien, CV, analyse externe, informations complémentaires, brief client)
sont saisies au même endroit et envoyées ensemble au backend, qui se charge de
l'intégralité de l'analyse et de la rédaction. Le front ne construit aucun
prompt, ne découpe rien et n'enchaîne aucune étape : il collecte, il affiche,
il exporte.

Le livrable renvoyé tient en trois sections, dans cet ordre :

1. **Compétences clés**
2. **Mail client** (encadré, prêt à copier ou à envoyer)
3. **Pitch candidat**

## Organisation

| Fichier | Rôle |
| --- | --- |
| `custom-widget/index.html` | Le balisage — volet **HTML** du widget Appsmith |
| `custom-widget/styles.css` | Les styles — volet **Style** du widget |
| `custom-widget/app.js` | Toute la logique — volet **JS** du widget |
| `tools/build.mjs` | Assemble les trois volets vers les deux cibles |
| `pages/Page1/Page1.json` | **Généré** — la page Appsmith |
| `dist/index.html` | **Généré, hors dépôt** — le même front en un fichier autonome |

Les trois fichiers de `custom-widget/` sont la seule source de vérité. Après
chaque modification :

```bash
node tools/build.mjs
```

`dist/index.html` s'ouvre directement dans un navigateur, sans Appsmith : c'est
le moyen le plus rapide de relire une modification. Il est ignoré par git (voir
`.gitignore`) puisqu'il se régénère.

## Charte graphique

La palette reprend les couleurs de la charte Valtus, relevées dans les
variables CSS de valtus.fr :

| Rôle | Couleur |
| --- | --- |
| Bleu profond — texte, bande d'en-tête, bouton principal | `#222f42` |
| Rouge — filets de titres, en-têtes de tableaux, éléments actifs | `#9b2841` |
| Rouge clair — accent sur fond sombre, alertes en thème sombre | `#e65160` |
| Bleu clair — focus clavier | `#4196b4` |
| Bleu pâle — confirmations | `#8cbed2` |

Les alertes reprennent le rouge de la charte : elles restent reconnaissables
parce qu'elles s'affichent toujours dans un encadré teinté, jamais en simple
filet. Ces mêmes couleurs sont reprises dans le document Word exporté (bleu
pour le titre, rouge pour les filets de section et les en-têtes de tableaux).

## Côté Appsmith

`Page1` contient un unique widget `CUSTOM_WIDGET` nommé `DossierCandidat`, qui
occupe toute la largeur du canevas sur 124 lignes (1 240 px). La page étant
longue, elle défile à l'intérieur de l'iframe du widget.

Pour appliquer une modification sans passer par le build : ouvrir le widget
dans Appsmith, bouton **Edit source**, et coller le contenu des trois fichiers
dans les volets HTML / Style / JS correspondants.

### Téléchargement du .docx

L'iframe d'un widget « Custom » est sandboxée **sans** `allow-downloads` : un
`<a download>` y est bloqué par le navigateur, silencieusement. Le widget ne
télécharge donc pas lui-même sous Appsmith — il publie le fichier dans son
modèle puis déclenche l'événement `onDownloadDocx`, câblé dans `Page1.json` sur :

```
{{download(DossierCandidat.model.docxData, DossierCandidat.model.docxFilename,
           'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}}
```

Si vous renommez le widget, mettez cette liaison à jour (elle est générée depuis
`WIDGET_NAME` dans `tools/build.mjs`). Hors Appsmith, `dist/index.html` retombe
sur un lien de téléchargement classique — voir `saveFile()` dans `app.js`.

## Brancher le backend

**Aucun appel n'est émis aujourd'hui.** Le bouton « Générer » produit un texte
d'exemple qui s'annonce lui-même comme tel, et un bandeau « Mode démonstration »
le rappelle en haut de page. Toute la chaîne en aval (aperçu, structure du
document, export Word, copie et envoi du mail) est en revanche complète.

Le branchement se fait en un seul point : définir `window.VALTUS_BACKEND` avant
le chargement du module (par exemple en tête du volet HTML). Dès qu'il est
défini, le bandeau disparaît et le texte d'exemple n'est plus utilisé.

```html
<script>
window.VALTUS_BACKEND = {
  generate: function (payload, options) {
    // payload = {
    //   mission: { candidat, poste, reference, societe, contact,
    //              lieu, tauxJournalier, signataire },
    //   sources: { transcription, cv, analyseExterne,
    //              informationsComplementaires, briefClient },
    //   preferencesMail: "…"
    // }
    // options.onProgress("Analyse du CV…") met à jour le libellé du bouton
    // pendant le traitement ; son appel est facultatif.
    return fetch("https://…/dossier-candidat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    })
      .then(function (r) {
        if (!r.ok) return Promise.reject({ code: "upstream_error" });
        return r.json();
      });
      // -> { competences: "…", mail: "Objet : …\n\n…", pitch: "…" }
  }
};
</script>
```

Les trois sections sont du markdown simple : titres `##`, puces `- `, tableaux
`|…|`, `**gras**`. Le mail commence par une ligne `Objet : …` — la copie la
conserve, l'envoi la détache pour alimenter le champ objet du lien `mailto:`.

En cas d'échec, rejeter avec `{ code, message }`. Les codes reconnus, qui
donnent un message adapté à l'écran, sont les clés de `ERROR_COPY` dans
`app.js` : `unauthorized`, `session_expired`, `rate_limited`,
`payload_too_large`, `refused`, `empty_result`, `timeout`, `upstream_error`,
`internal_error`.

### Variante : passer par une query Appsmith

Pour que l'appel parte du serveur Appsmith et que la clé d'API ne soit jamais
exposée au navigateur, déclarer une query `GenererDossier`, ajouter l'événement
`onGenerate` au widget, puis :

```js
window.VALTUS_BACKEND = {
  generate: function (payload) {
    return new Promise(function (resolve, reject) {
      var token = String(Date.now());
      var stop = appsmith.onModelChange(function (model) {
        if (!model || model.resultToken !== token) return;
        stop();
        if (model.resultError) reject({ code: "upstream_error", message: model.resultError });
        else resolve(model.result);
      });
      appsmith.updateModel({ request: payload, requestToken: token });
      appsmith.triggerEvent("onGenerate");
    });
  }
};
```

Côté Appsmith, `onGenerate` exécute la query avec
`{{DossierCandidat.model.request}}` puis renvoie le résultat au widget en
alimentant `result` et `resultToken` dans le modèle.

## Limites connues de l'environnement Appsmith

L'iframe est sandboxée sans `allow-same-origin`, ce qui a trois conséquences,
toutes déjà gérées dans le code :

- `localStorage` peut lever une exception — les préférences de rédaction sont
  lues et écrites dans un `try/catch` et le champ démarre simplement vide si le
  stockage est indisponible ;
- `navigator.clipboard` peut être refusé — « Copier le mail » affiche alors un
  message invitant à sélectionner le texte dans l'aperçu ;
- « Envoyer par email » ouvre un lien `mailto:`, dont le support dépend du poste
  de travail ; au-delà de 1 800 caractères le mail est copié dans le
  presse-papiers plutôt que tronqué silencieusement.

JSZip (export Word) et pdf.js (import de PDF) sont chargés depuis cdnjs par le
volet HTML. Sur une instance Appsmith sans accès Internet sortant, il faut les
héberger et adapter les deux balises `<script src>` en tête de `index.html` —
sans JSZip l'export Word est inopérant, sans pdf.js l'import de PDF se désactive
de lui-même.
