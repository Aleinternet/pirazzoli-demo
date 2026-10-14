// Offline browser-logic regressions: node --test tests/login-resilience.test.cjs
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const { test } = require('node:test');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
function between(start, end) {
  const first = html.indexOf(start);
  const last = html.indexOf(end, first);
  assert(first >= 0 && last > first, `Missing code block ${start}`);
  return html.slice(first, last);
}
const helpers = between('    function sleep(ms)', '    function buildHistoryState()');
const indicators = between('    async function updateExchangeRates()', '    function renderAll()');

function context(extra = {}) {
  const ctx = vm.createContext({
    setTimeout, clearTimeout, AbortController, console: { error() {} },
    loadData: async () => {}, updateExchangeRates: () => new Promise(() => {}),
    ...extra,
  });
  vm.runInContext('let initialDataPromise = null;\n' + helpers, ctx);
  return ctx;
}

test('inline JavaScript parses', () => {
  for (const match of html.matchAll(/<script([^>]*)>([\s\S]*?)<\/script>/gi)) {
    if (!/\bsrc\s*=/.test(match[1])) new vm.Script(match[2]);
  }
});

test('initial load completes even when indicators never respond', async () => {
  let loads = 0;
  const ctx = context({ loadData: async () => { loads++; } });
  let timer;
  try {
    await Promise.race([
      ctx.ensureInitialData(),
      new Promise((_, reject) => { timer = setTimeout(() => reject(new Error('Indicators blocked login')), 100); }),
    ]);
    await ctx.ensureInitialData();
    assert.equal(loads, 1);
  } finally { clearTimeout(timer); }
});

test('a failed initial load can be retried successfully', async () => {
  let fail = true;
  let calls = 0;
  const ctx = context({
    loadData: async () => { calls++; if (fail) throw new Error('offline'); },
    setTimeout: (fn) => setTimeout(fn, 0),
  });
  await assert.rejects(ctx.ensureInitialData());
  fail = false;
  await ctx.ensureInitialData();
  assert.equal(calls, 3);
});

test('a stalled fetch is aborted and its timer is cleared', async () => {
  let aborted = false;
  let cleared = false;
  const ctx = context({
    fetch: (_, { signal }) => new Promise((_, reject) => {
      signal.addEventListener('abort', () => { aborted = true; reject(new Error('aborted')); });
    }),
    clearTimeout: id => { cleared = true; clearTimeout(id); },
  });
  await assert.rejects(ctx.fetchJsonWithTimeout('/api', 5));
  assert(aborted && cleared);
});

test('server error bodies are not copied into errors', async () => {
  const ctx = context({ fetch: async () => ({ ok: false, json: async () => ({ error: 'apikey=secret' }) }) });
  await assert.rejects(ctx.fetchJsonWithTimeout('/api'), error => !error.message.includes('secret'));
});

test('one failed indicator does not erase the working indicator', async () => {
  const nodes = { fxUf: {}, fxUsd: {}, updated: {} };
  const ctx = context({
    fxUf: nodes.fxUf, fxUsd: nodes.fxUsd,
    document: { getElementById: () => nodes.updated },
    formatDateTime: date => date.toISOString(),
    fetch: async url => {
      if (url.endsWith('/dolar')) throw new Error('offline');
      return { ok: true, json: async () => ({ serie: [{ valor: 40000, fecha: '2026-10-14' }] }) };
    },
  });
  vm.runInContext(indicators, ctx);
  await ctx.updateExchangeRates();
  assert.equal(nodes.fxUsd.textContent, 'No disponible');
  assert.notEqual(nodes.fxUf.textContent, 'No disponible');
  assert.equal(nodes.fxUf.className, '');
});

test('login failure shows a safe message and restores the submit button', async () => {
  const nodes = { loginError: { style: {} }, loginSubmitBtn: {}, statProfile: {} };
  let submit;
  const ctx = context({
    ...nodes,
    loginForm: { addEventListener: (_, handler) => { submit = handler; } },
    loginUser: { value: 'test-user' }, loginPass: { value: 'test-only' },
    findAuth: () => ({ label: 'Test' }), updateProfileBadge() {}, renderLoadingHome() {},
  });
  ctx.ensureInitialData = async () => { throw new Error('apikey=secret'); };
  vm.runInContext(between('    loginForm.addEventListener("submit"', '    searchInput.addEventListener'), ctx);
  await submit({ preventDefault() {} });
  assert(!nodes.loginError.textContent.includes('secret'));
  assert(!nodes.loginError.textContent.includes('Detalle:'));
  assert.equal(nodes.loginSubmitBtn.disabled, false);
  assert.equal(nodes.loginSubmitBtn.textContent, 'Entrar');
});
