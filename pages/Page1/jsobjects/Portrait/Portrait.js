export default {
  async lancer() {
    storeValue("result", null);
    storeValue("etat", "Analyse en cours…");
    const r = await StartJob.run();
    storeValue("jobId", r.job_id);
    this.attendre();
  },

  async attendre() {
    for (let i = 0; i < 60; i++) {
      await new Promise(s => setTimeout(s, 5000));
      const r = await GetJob.run();
      if (r.status === "DONE")   { storeValue("result", r); storeValue("etat", "Terminé"); return; }
      if (r.status === "FAILED") { storeValue("etat", "Échec : " + (r.error || "voir CloudWatch")); return; }
      storeValue("etat", `Génération… ${(i + 1) * 5}s`);
    }
    storeValue("etat", "Toujours en cours après 5 min");
  }
}