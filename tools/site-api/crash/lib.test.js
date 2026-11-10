'use strict';
const test = require('node:test');
const assert = require('node:assert');
const { validate, titleOf, issueBody, handle } = require('./lib');

const sample = {
  signature: 'a1b2c3d4e5f6', version: '2.1.3', context: 'interface', fatal: false,
  type: 'System.InvalidOperationException', message: 'Ouch dans C:\\Users\\marie\\Documents',
  stack: 'System.InvalidOperationException: Ouch\n   at OpenWrite.App.Controls.PageView.OnMouseMove(MouseEventArgs e)\n   at System.Windows.UIElement.X()',
  os: 'Microsoft Windows NT 10.0.26100.0 x64', runtime: '.NET 10.0.0', culture: 'fr-BE', timeUtc: '2026-11-04T10:00:00Z',
};

test('validation et anonymisation', () => {
  const r = validate(JSON.stringify(sample));
  assert.equal(typeof r, 'object');
  assert.ok(!r.message.includes('marie'));
  assert.equal(validate('{"signature":"zz"}'), 'signature invalide');
  assert.equal(validate('pas du json'), 'JSON invalide');
  assert.equal(validate('x'.repeat(40000)), 'rapport trop volumineux');
  assert.equal(validate(JSON.stringify({ ...sample, type: 'a b <script>' })), 'type invalide');
});

test('titre et corps', () => {
  const r = validate(JSON.stringify(sample));
  assert.equal(titleOf(r), '[crash a1b2c3d4e5f6] InvalidOperationException dans PageView.OnMouseMove');
  assert.ok(issueBody(r).includes('openwrite-crash:a1b2c3d4e5f6'));
  const evil = validate(JSON.stringify({ ...sample, stack: 'x\n```\n@everyone' }));
  assert.ok(!issueBody(evil).includes('```\n@everyone'));
});

function fakeGitHub(items) {
  const calls = [];
  const fetchImpl = async (url, opts) => {
    calls.push([opts.method, url.replace('https://api.github.com', ''), opts.body && JSON.parse(opts.body)]);
    const body = opts.method === 'GET' ? { items } : { number: 42 };
    return { ok: true, status: 200, text: async () => JSON.stringify(body) };
  };
  return { calls, fetchImpl };
}

test('nouvelle erreur : issue creee', async () => {
  const gh = fakeGitHub([]);
  const res = await handle(validate(JSON.stringify(sample)), { token: 't', repo: 'o/r', fetchImpl: gh.fetchImpl });
  assert.deepEqual(res, { action: 'creee', number: 42 });
  assert.equal(gh.calls[1][0], 'POST');
  assert.equal(gh.calls[1][1], '/repos/o/r/issues');
  assert.deepEqual(gh.calls[1][2].labels, ['crash']);
});

test('erreur connue : commentaire sur l issue ouverte', async () => {
  const gh = fakeGitHub([{ number: 7, state: 'open', comments: 2, title: '[crash a1b2c3d4e5f6] X' }]);
  const res = await handle(validate(JSON.stringify(sample)), { token: 't', repo: 'o/r', fetchImpl: gh.fetchImpl });
  assert.deepEqual(res, { action: 'commentee', number: 7 });
  assert.equal(gh.calls[1][1], '/repos/o/r/issues/7/comments');
});

test('erreur corrigee qui revient : issue rouverte', async () => {
  const gh = fakeGitHub([{ number: 9, state: 'closed', comments: 0, title: '[crash a1b2c3d4e5f6] X' }]);
  const res = await handle(validate(JSON.stringify(sample)), { token: 't', repo: 'o/r', fetchImpl: gh.fetchImpl });
  assert.deepEqual(res, { action: 'rouverte', number: 9 });
  assert.deepEqual(gh.calls[1][2], { state: 'open', labels: ['crash', 'regression'] });
});
