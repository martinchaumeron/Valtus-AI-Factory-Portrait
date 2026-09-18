# Front « Générateur de dossier candidat »

Un écran, **un seul traitement**. Toutes les sources (retranscription
d'entretien, CV, analyse externe, informations complémentaires, brief client)
sont saisies au même endroit et envoyées ensemble au service, qui se charge de
l'intégralité de l'analyse et de la rédaction. Le front ne construit aucun
prompt, ne découpe rien et n'enchaîne aucune étape : il collecte, il affiche,
il exporte.

Le livrable tient en trois sections, dans cet ordre :

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
| `pages/Page1/widgets/DossierAnalyse.json` | **Généré** — le widget de la page |
| `dist/index.html` | **Généré, hors dépôt** — le même front en un fichier autonome |

Les trois fichiers de `custom-widget/` sont la seule source de vérité. Après
chaque modification :

```bash
node tools/build.mjs
```

`dist/index.html` s'ouvre directement dans un navigateur, sans Appsmith : c'est
le moyen le plus rapide de relire une modification. Il est ignoré par git
puisqu'il se régénère, et il bascule de lui-même en mode démonstration — aucun
appel n'y est émis.

## Le chemin d'un dossier

```
   widget (iframe)                    application Appsmith              service
   ───────────────                    ────────────────────              ───────
   clic « Générer »
     └─ triggerEvent("onGenerate",
            { payload, jeton })  ──▶  Portrait.lancer(payload, jeton)
                                        └─ StartJob.run({ payload })  ──▶  POST /jobs
                                                                       ◀──  { job_id }
                                        └─ storeValue("jobId", …)
                                        └─ Portrait.attendre()
                                             toutes les 5 s, 60 fois :
                                             GetJob.run()             ──▶  GET /jobs/{id}
                                                                       ◀──  { status, … }
   onModelChange(modèle)         ◀──  storeValue("progression" | "result"
                                                 | "erreur" | "etat")
   promesse résolue / rejetée
```

**Widget → application : les événements.** Les données voyagent dans l'objet de
contexte de `triggerEvent()`, dont les clés sont directement lisibles dans la
liaison :

| Événement | Liaison | Contexte |
| --- | --- | --- |
| `onGenerate` | `{{Portrait.lancer(payload, jeton)}}` | `payload`, `jeton` |
| `onDownloadDocx` | `{{download(docxData, docxFilename, '…')}}` | `docxData`, `docxFilename` |

**Application → widget : `defaultModel`.** Le modèle du widget est lié au
store :

```
{{ { etat: appsmith.store.etat, progression: appsmith.store.progression,
     resultat: appsmith.store.result, erreur: appsmith.store.erreur,
     jeton: appsmith.store.jeton } }}
```

`etat` vaut `EN_COURS`, `TERMINE` ou `ECHEC` ; `progression` porte le libellé
affiché sur le bouton. `result` et `erreur` sont toujours déposés **avant**
l'état terminal, de sorte que le widget les trouve déjà présents quand il le
voit.

**Pourquoi pas `updateModel()` dans le sens widget → application.** Appsmith
réinitialise une propriété meta dès que sa propriété « par défaut » change. Le
modèle étant lié au store, tout ce que le widget y écrirait serait effacé au
premier `storeValue()` de `Portrait.lancer()`. L'objet de contexte, lui,
accompagne l'événement et ne peut pas être écrasé — d'où son usage pour les
deux sens sortants, y compris le fichier Word.

**Le jeton.** Il identifie un lancement. Le widget n'écoute que les mises à jour
portant son jeton : une génération précédente encore en vol ne peut pas résoudre
celle en cours, et le rappel immédiat d'`onModelChange` (qui rejoue le modèle
courant) est ignoré.

## Contrat d'API

**`POST /jobs`** — corps envoyé, construit par `buildPayload()` dans `app.js` et
transmis à la query `StartJob` en paramètre (`{{ this.params.payload }}`) :

```json
{
  "sources": {
    "transcript":                   { "text": "…" },
    "cv":                           { "text": "…" },
    "brief":                        { "text": "…" },
    "analyse_externe":              { "text": "…" },
    "informations_complementaires": { "text": "…" }
  },
  "commercial": {
    "contact": "…", "societe": "…", "lieu": "…",
    "taux": "…", "signataire": "…"
  },
  "dossier": {
    "candidat": "…", "poste": "…", "reference": "…", "preferences_mail": "…"
  }
}
```

`sources.transcript|cv|brief` et `commercial` sont les clés déjà attendues par
le service. **`sources.analyse_externe`, `sources.informations_complementaires`
et `dossier` sont des ajouts** : l'écran collecte ces informations, elles sont
donc transmises dans la même forme. Un service qui ne les connaît pas les
ignore ; à confirmer côté backend s'il valide strictement son entrée.

Réponse attendue : `{ "job_id": "…" }`.

**`GET /jobs/{job_id}`** — `{ "status": "DONE" | "FAILED" | …, "error": "…" }`.

À `DONE`, le front cherche les trois sections dans la réponse : par leur nom,
**accents, casse et séparateurs ignorés** (`competences`, `competences_cles`,
`skills` ; `mail`, `email`, `message` ; `pitch`, `accroche`, `resume`), à la
racine puis un niveau plus bas — `{ status, result: { … } }` fonctionne donc
aussi bien que des sections à plat. Si rien n'est reconnu, la réponse brute est
affichée intégralement plutôt que perdue. Voir `normalizeResult()` dans
`app.js` ; c'est le seul endroit à ajuster si la forme de la réponse se précise.

## Remplacer le moteur

Pour court-circuiter les queries Appsmith (test, autre service), définir
`window.VALTUS_BACKEND` avant le chargement du module — il a la priorité sur
tout le reste :

```html
<script>
window.VALTUS_BACKEND = {
  generate: function (payload, options) {
    // options.onProgress("Analyse du CV…") met à jour le libellé du bouton.
    // Résoudre avec { competences, mail, pitch } ; rejeter avec
    // { code, message } — les codes reconnus sont les clés de ERROR_COPY.
    return Promise.resolve({ competences: "…", mail: "…", pitch: "…" });
  }
};
</script>
```

Sans lui et hors Appsmith, le front produit un texte d'exemple explicitement
identifié comme tel, signalé par un bandeau « Mode démonstration ».

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
filet. Ces mêmes couleurs sont reprises dans le document Word exporté.

## Côté Appsmith

`Page1` contient un unique widget `CUSTOM_WIDGET` nommé `DossierAnalyse`, qui
occupe toute la largeur du canevas sur 124 lignes (1 240 px). La page étant
longue, elle défile à l'intérieur de l'iframe du widget.

Pour appliquer une modification sans passer par le build : ouvrir le widget
dans Appsmith, bouton **Edit source**, et coller le contenu des trois fichiers
dans les volets HTML / Style / JS correspondants.

L'iframe d'un widget « Custom » est sandboxée **sans** `allow-downloads` : un
`<a download>` y est bloqué par le navigateur, silencieusement. Sous Appsmith
le fichier Word part donc par l'événement `onDownloadDocx` ; hors Appsmith,
`dist/index.html` retombe sur un lien de téléchargement classique — voir
`saveFile()` dans `app.js`.

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
