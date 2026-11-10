// POST /api/crash : rapport d'erreur anonyme d'OpenWrite -> issue GitHub (creee ou completee).
// Reglages de l'application Static Web App (portail Azure > Configuration > Variables d'environnement) :
//   GITHUB_TOKEN  jeton a grain fin limite au depot, permission « Issues : lecture et ecriture » (obligatoire)
//   GITHUB_REPO   « proprietaire/depot » (defaut prolice/open-write)
'use strict';
const { validate, handle } = require('./lib');

module.exports = async function (context, req) {
  const reply = (status, body) => { context.res = { status, headers: { 'Content-Type': 'application/json' }, body }; };
  const token = process.env.GITHUB_TOKEN;
  if (!token) return reply(503, { error: 'relais non configure' });   // 5xx : l'application reessaiera plus tard

  const raw = typeof req.rawBody === 'string' ? req.rawBody : JSON.stringify(req.body ?? '');
  const r = validate(raw);
  if (typeof r === 'string') return reply(400, { error: r });

  try {
    const result = await handle(r, { token, repo: process.env.GITHUB_REPO || 'prolice/open-write' });
    context.log(`crash ${r.signature} : issue #${result.number} ${result.action}`);
    return reply(200, result);
  } catch (e) {
    context.log.error(`crash ${r.signature} : ${e.message}`);
    return reply(502, { error: 'GitHub indisponible' });
  }
};
