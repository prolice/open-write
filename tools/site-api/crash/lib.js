// Logique du relais des rapports d'erreur (sans dependance, testable avec `node --test`).
// Le depot GitHub est PUBLIC : le rapport est valide et re-anonymise ici avant d'ouvrir une issue.
'use strict';

const LABEL = 'crash';
const MAX_BODY = 32 * 1024;          // octets
const MAX_COMMENTS = 30;             // au-dela, les occurrences ne sont plus commentees (le compteur suffit)

function str(v, max) {
  return typeof v === 'string' ? v.slice(0, max) : '';
}

// Seconde ligne de defense : meme anonymisation que l'application (CrashReport.Scrub), en plus simple.
function scrub(s) {
  return s
    .replace(/[A-Za-z]:\\(Users|Documents and Settings)\\[^\\/:*?"<>|\r\n]+/gi, '%USERPROFILE%')
    .replace(/\\\\[^\\\s]+\\[^\\\s]+/g, '\\\\serveur\\partage')
    .replace(/[\w.+-]+@[\w-]+(\.[\w-]+)+/g, '<e-mail>')
    .replace(/(ghp_|github_pat_|gho_|ghs_)[A-Za-z0-9_]+/g, '<jeton>');
}

// Rapport -> objet valide, ou message d'erreur (chaine).
function validate(raw) {
  if (typeof raw !== 'string' || raw.length === 0) return 'corps vide';
  if (Buffer.byteLength(raw, 'utf8') > MAX_BODY) return 'rapport trop volumineux';
  let r;
  try { r = JSON.parse(raw); } catch { return 'JSON invalide'; }
  if (!r || typeof r !== 'object') return 'JSON invalide';
  const signature = str(r.signature, 64);
  const version = str(r.version, 20);
  const type = str(r.type, 200);
  if (!/^[0-9a-f]{12}$/.test(signature)) return 'signature invalide';
  if (!/^\d{1,3}\.\d{1,3}\.\d{1,5}$/.test(version)) return 'version invalide';
  if (!/^[A-Za-z_][\w.`+<>,\[\]]*$/.test(type)) return 'type invalide';
  return {
    signature, version, type,
    context: scrub(str(r.context, 40)),
    fatal: r.fatal === true,
    message: scrub(str(r.message, 600)),
    stack: scrub(str(r.stack, 12000)),
    os: scrub(str(r.os, 120)),
    runtime: str(r.runtime, 80),
    culture: str(r.culture, 20),
    timeUtc: str(r.timeUtc, 40),
  };
}

// Methode d'OpenWrite la plus haute de la pile, pour un titre lisible.
function topFrame(stack) {
  const frames = [...stack.matchAll(/^\s*at\s+([^(\r\n]+)\(/gm)].map(m => m[1].trim());
  const own = frames.find(f => f.startsWith('OpenWrite.')) || frames[0];
  if (!own) return '';
  const parts = own.split('.');
  return parts.slice(-2).join('.');
}

function titleOf(r) {
  const shortType = r.type.split('.').pop();
  const where = topFrame(r.stack);
  return `[crash ${r.signature}] ${shortType}${where ? ' dans ' + where : ''}`.slice(0, 250);
}

// Neutralise un eventuel ``` dans le texte (sinon le bloc de code se termine).
function fence(s) {
  return s.replace(/```/g, "'''");
}

function issueBody(r) {
  return [
    `Rapport d'erreur automatique envoyé par OpenWrite (anonymisé : ni document, ni nom de fichier, ni utilisateur).`,
    '',
    '| | |', '|---|---|',
    `| Signature | \`${r.signature}\` |`,
    `| Version | ${r.version} |`,
    `| Origine | ${r.context}${r.fatal ? ' (**fatale**)' : ''} |`,
    `| Système | ${r.os} — ${r.runtime} — ${r.culture} |`,
    `| Date (UTC) | ${r.timeUtc} |`,
    '',
    `**${r.type}**`,
    '',
    '```', fence(r.message), '```',
    '',
    '```', fence(r.stack), '```',
    '',
    `<!-- openwrite-crash:${r.signature} -->`,
  ].join('\n');
}

function commentBody(r) {
  return `Nouvelle occurrence — version ${r.version}, ${r.context}${r.fatal ? ' (fatale)' : ''}, ${r.os}, ${r.timeUtc}`;
}

// Client GitHub minimal (fetch natif de Node 18+).
function github(token, fetchImpl = fetch) {
  return async (method, path, body) => {
    const res = await fetchImpl('https://api.github.com' + path, {
      method,
      headers: {
        'Authorization': 'Bearer ' + token,
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'User-Agent': 'openwrite-crash-relay',
        ...(body ? { 'Content-Type': 'application/json' } : {}),
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    const text = await res.text();
    if (!res.ok) throw new Error(`GitHub ${method} ${path} : ${res.status} ${text.slice(0, 300)}`);
    return text ? JSON.parse(text) : null;
  };
}

// Cree l'issue, ou commente l'issue ouverte de meme signature. Une issue fermee dont l'erreur revient est
// rouverte avec le label « regression ». Le message et la pile sont dans des blocs de code (pas de @mention).
async function handle(r, { token, repo, fetchImpl }) {
  const gh = github(token, fetchImpl);
  const q = encodeURIComponent(`repo:${repo} is:issue label:${LABEL} in:title "crash ${r.signature}"`);
  const found = await gh('GET', `/search/issues?q=${q}&per_page=5`);
  const issues = (found.items || []).filter(i => i.title.includes(`[crash ${r.signature}]`));
  const open = issues.find(i => i.state === 'open');
  if (open) {
    if (open.comments < MAX_COMMENTS) await gh('POST', `/repos/${repo}/issues/${open.number}/comments`, { body: commentBody(r) });
    return { action: 'commentee', number: open.number };
  }
  const closed = issues[0];
  if (closed) {
    await gh('PATCH', `/repos/${repo}/issues/${closed.number}`, { state: 'open', labels: [LABEL, 'regression'] });
    await gh('POST', `/repos/${repo}/issues/${closed.number}/comments`, { body: 'Erreur réapparue après fermeture.\n\n' + commentBody(r) });
    return { action: 'rouverte', number: closed.number };
  }
  const created = await gh('POST', `/repos/${repo}/issues`, {
    title: titleOf(r), body: issueBody(r), labels: r.fatal ? [LABEL, 'fatale'] : [LABEL],
  });
  return { action: 'creee', number: created.number };
}

module.exports = { validate, scrub, titleOf, issueBody, handle, MAX_BODY };
