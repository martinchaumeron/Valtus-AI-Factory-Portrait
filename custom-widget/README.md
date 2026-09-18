# Front « Dossier d'analyse consolidé »

Le front complet de l'outil, sur **une seule page** : l'Étape 1 (dossier
d'analyse consolidé) et l'Étape 2 (dossier de présentation candidat) sont
empilées l'une sous l'autre, reliées par un sommaire collant en haut de page.
La version d'origine les séparait en deux onglets.

## Organisation

| Fichier | Rôle |
| --- | --- |
| `custom-widget/index.html` | Le balisage — volet **HTML** du widget Appsmith |
| `custom-widget/styles.css` | Les styles — volet **Style** du widget |
| `custom-widget/app.js` | Toute la logique — volet **JS** du widget |
| `tools/build.mjs` | Assemble les trois volets vers les deux cibles |
| `pages/Page1/Page1.json` | **Généré** — la page Appsmith |
| `dist/index.html` | **Généré** — le même front en un fichier autonome |

Les trois fichiers de `custom-widget/` sont la seule source de vérité. Après
chaque modification :

```bash
node tools/build.mjs
```

`dist/index.html` s'ouvre directement dans un navigateur, sans Appsmith : c'est
le moyen le plus rapide de relire une modification.

## Côté Appsmith

`Page1` contient un unique widget `CUSTOM_WIDGET` nommé `DossierAnalyse`, qui
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
{{download(DossierAnalyse.model.docxData, DossierAnalyse.model.docxFilename,
           'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}}
```

Si vous renommez le widget, mettez cette liaison à jour (elle est générée depuis
`WIDGET_NAME` dans `tools/build.mjs`). Hors Appsmith, `dist/index.html` retombe
sur un lien de téléchargement classique — voir `saveFile()` dans `app.js`.

## Brancher le moteur de génération

**Aucun appel n'est émis aujourd'hui.** Les boutons « Générer » produisent un
texte d'exemple qui s'annonce lui-même comme tel, et un bandeau « Mode
démonstration » le rappelle en haut de page. Toute la chaîne en aval (aperçu,
structure du document, export Word, enchaînement Étape 1 → Étape 2) est en
revanche complète et fonctionnelle.

Le branchement se fait en un seul point : définir `window.VALTUS_LLM` avant le
chargement du module (par exemple en tête du volet HTML). Dès qu'il est défini,
le bandeau disparaît et le stub n'est plus utilisé.

```html
<script>
window.VALTUS_LLM = {
  // prompt  : le texte complet à envoyer
  // options : { modelTier, cache, onText } — onText({ text }) peut être
  //           appelé pendant le streaming pour alimenter l'aperçu au fil de
  //           l'eau ; il est facultatif.
  // Résoudre avec { text, truncated } ; rejeter avec { code, message }, où
  // code est l'une des clés de SAMPLE_ERROR_COPY (rate_limited,
  // prompt_too_large, refused, upstream_error…) pour obtenir un message
  // d'erreur adapté à l'écran.
  complete: function (prompt, options) {
    return fetch("https://…/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt: prompt })
    })
      .then(function (r) {
        if (!r.ok) return Promise.reject({ code: "upstream_error" });
        return r.json();
      })
      .then(function (data) {
        return { text: data.text, truncated: !!data.truncated };
      });
  }
};
</script>
```

### Variante : passer par une query Appsmith

Pour que l'appel parte du serveur Appsmith (et que la clé d'API ne soit jamais
exposée au navigateur), déclarer une query `GenerateBrief`, ajouter l'événement
`onGenerate` au widget, puis :

```js
window.VALTUS_LLM = {
  complete: function (prompt) {
    return new Promise(function (resolve, reject) {
      var token = String(Date.now());
      var stop = appsmith.onModelChange(function (model) {
        if (!model || model.briefToken !== token) return;
        stop();
        if (model.briefError) reject({ code: "upstream_error", message: model.briefError });
        else resolve({ text: model.briefText, truncated: false });
      });
      appsmith.updateModel({ promptText: prompt, promptToken: token });
      appsmith.triggerEvent("onGenerate");
    });
  }
};
```

Côté Appsmith, `onGenerate` exécute la query avec
`{{DossierAnalyse.model.promptText}}` puis renvoie le résultat au widget via
`setModel`/`storeValue` alimentant `briefText` et `briefToken`.

### Tailles de prompt

Le découpage automatique en lots (au-delà de 64 Kio par appel) vient du moteur
d'origine. Selon le moteur branché, ajuster `MAX_PROMPT_BYTES` et
`T2_MAX_PROMPT_BYTES` dans `app.js`.

## Limites connues de l'environnement Appsmith

L'iframe est sandboxée sans `allow-same-origin`, ce qui a trois conséquences,
toutes déjà gérées dans le code :

- `localStorage` peut lever une exception — les préférences de mail de
  l'Étape 2 sont lues et écrites dans un `try/catch` et le champ démarre
  simplement vide si le stockage est indisponible ;
- `navigator.clipboard` peut être refusé — « Copier le mail client » affiche
  alors un message invitant à sélectionner le texte dans l'aperçu ;
- le bouton « Envoyer par email » ouvre un lien `mailto:`, dont le support
  dépend du poste de travail ; au-delà de 1 800 caractères le mail est copié
  dans le presse-papiers plutôt que tronqué silencieusement.

JSZip (export Word) et pdf.js (import de PDF) sont chargés depuis cdnjs par le
volet HTML. Sur une instance Appsmith sans accès Internet sortant, il faut les
héberger et adapter les deux balises `<script src>` en tête de `index.html` —
sans JSZip l'export Word est inopérant, sans pdf.js l'import de PDF se désactive
de lui-même.
