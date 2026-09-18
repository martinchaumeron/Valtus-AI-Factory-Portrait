export default {
  /**
   * Déclenché par l'événement onGenerate du widget DossierAnalyse.
   *
   * @param payload  le corps de la requête, construit par le front (sources,
   *                 données commerciales, informations de dossier). Il est
   *                 passé en paramètre à StartJob plutôt que déposé dans le
   *                 store : la query le lit avec {{this.params.payload}}, sans
   *                 dépendre d'un cycle de réévaluation.
   * @param jeton    identifie CE lancement. Le widget n'écoute que les mises à
   *                 jour portant son jeton : un lancement précédent, encore en
   *                 vol, ne peut pas résoudre la génération en cours.
   *
   * Le widget suit l'avancement par la liaison `defaultModel`, qui recopie
   * `etat`, `progression`, `result`, `erreur` et `jeton` depuis le store.
   * `etat` vaut EN_COURS, TERMINE ou ECHEC ; `progression` porte le libellé
   * affiché. `result` et `erreur` sont toujours déposés AVANT l'état terminal,
   * de sorte que le widget les trouve déjà présents quand il le voit.
   */
  async lancer(payload, jeton) {
    storeValue("jeton", jeton);
    storeValue("result", null);
    storeValue("erreur", null);
    storeValue("progression", "Envoi des sources…");
    storeValue("etat", "EN_COURS");

    try {
      const r = await StartJob.run({ payload: JSON.stringify(payload) });

      if (!r || !r.job_id) {
        storeValue("erreur", "Le service n'a pas renvoyé d'identifiant de traitement.");
        storeValue("etat", "ECHEC");
        return;
      }

      storeValue("jobId", r.job_id);
      storeValue("progression", "Analyse en cours…");
      return this.attendre();
    } catch (e) {
      storeValue("erreur", "Impossible de lancer le traitement : " + (e.message || e));
      storeValue("etat", "ECHEC");
    }
  },

  /**
   * Interroge GetJob toutes les 5 secondes, jusqu'à 5 minutes.
   */
  async attendre() {
    for (let i = 0; i < 60; i++) {
      await new Promise(s => setTimeout(s, 5000));

      let r;
      try {
        r = await GetJob.run();
      } catch (e) {
        storeValue("erreur", "Interrogation du traitement impossible : " + (e.message || e));
        storeValue("etat", "ECHEC");
        return;
      }

      if (r.status === "DONE") {
        storeValue("result", r);
        storeValue("progression", "Terminé");
        storeValue("etat", "TERMINE");
        return;
      }

      if (r.status === "FAILED") {
        storeValue("erreur", r.error || "Le traitement a échoué (voir CloudWatch).");
        storeValue("etat", "ECHEC");
        return;
      }

      storeValue("progression", `Génération… ${(i + 1) * 5} s`);
    }

    storeValue("erreur", "Le traitement dépasse 5 minutes. Il se poursuit peut-être côté service : relancez dans un moment.");
    storeValue("etat", "ECHEC");
  }
}
