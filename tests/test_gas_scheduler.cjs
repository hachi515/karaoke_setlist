const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, '../.github/gas/hourly_update.gs'), 'utf8');
const endpoint = 'https://api.github.com/repos/hachi515/karaoke_setlist/actions/workflows/update.yml';

function harness({token = 'test-fixture-token', replies = [], triggers = [], createError = false, locked = false} = {}) {
  const state = {requests: [], deleted: [], sleeps: [], logs: [], released: 0, hours: null, created: [], properties: {GITHUB_ACTIONS_TOKEN: token}};
  const context = vm.createContext({
    console: {log: message => state.logs.push(message)},
    PropertiesService: {getScriptProperties: () => ({
      getProperty: name => state.properties[name],
      setProperty: (name, value) => {state.properties[name] = value;}
    })},
    UrlFetchApp: {fetch: (url, options) => {
      state.requests.push({url, options});
      const reply = replies.shift();
      assert(reply, 'No unexpected remote requests');
      return {getResponseCode: () => reply.status, getContentText: () => JSON.stringify(reply.body), getAllHeaders: () => reply.headers || {}};
    }},
    Utilities: {sleep: ms => state.sleeps.push(ms)},
    LockService: {getScriptLock: () => ({tryLock: () => !locked, releaseLock: () => state.released++})},
    ScriptApp: {
      getProjectTriggers: () => triggers,
      deleteTrigger: trigger => state.deleted.push(trigger.getHandlerFunction()),
      newTrigger: handler => ({timeBased: () => ({everyHours: hours => {
        state.hours = hours;
        return {create: () => {
          if (createError) throw new Error('trigger quota');
          state.created.push(handler);
        }};
      }})})
    }
  });
  vm.runInContext(source, context);
  return {context, state};
}
const trigger = name => ({getHandlerFunction: () => name});
const validWorkflow = {status: 200, body: {state: 'active', path: '.github/workflows/update.yml'}};

test('dispatch targets the existing main workflow and records acceptance, without logging credentials', () => {
  const {context, state} = harness({replies: [{status: 204}]});
  context.dispatchKaraokeUpdate();
  assert.equal(state.requests.length, 1);
  const {url, options} = state.requests[0];
  assert.equal(url, endpoint + '/dispatches');
  assert.equal(options.method, 'post');
  assert.deepEqual(JSON.parse(options.payload), {ref: 'main', inputs: {trigger_source: 'gas-hourly'}});
  assert.equal(options.headers.Authorization, 'Bearer test-fixture-token');
  assert.ok(state.properties.KARAOKE_LAST_DISPATCH_AT);
  assert.equal(state.released, 1);
  assert.ok(!state.logs.join('').includes('test-fixture-token'));
});

test('installation verifies access, replaces only its own triggers, and dispatches the first update', () => {
  const {context, state} = harness({replies: [validWorkflow, {status: 204}], triggers: [trigger('existingCsvHandler'), trigger('dispatchKaraokeUpdate')]});
  context.installKaraokeHourlyTrigger();
  assert.equal(state.requests[0].url, endpoint);
  assert.equal(state.requests[0].options.method, 'get');
  assert.equal(state.hours, 1);
  assert.deepEqual(state.created, ['dispatchKaraokeUpdate']);
  assert.deepEqual(state.deleted, ['dispatchKaraokeUpdate']);
  assert.equal(state.requests[1].options.method, 'post');
});

test('missing authentication never creates a trigger or sends a request', () => {
  const {context, state} = harness({token: ''});
  assert.throws(() => context.installKaraokeHourlyTrigger(), /GITHUB_ACTIONS_TOKEN/);
  assert.deepEqual(state.created, []);
  assert.deepEqual(state.requests, []);
});

test('disabled workflow is rejected without changing existing triggers', () => {
  const {context, state} = harness({replies: [{status: 200, body: {state: 'disabled_manually', path: '.github/workflows/update.yml'}}]});
  assert.throws(() => context.installKaraokeHourlyTrigger(), /有効ではありません/);
  assert.deepEqual(state.created, []);
  assert.deepEqual(state.deleted, []);
});

test('trigger creation failure preserves the old trigger', () => {
  const {context, state} = harness({replies: [validWorkflow], triggers: [trigger('dispatchKaraokeUpdate')], createError: true});
  assert.throws(() => context.installKaraokeHourlyTrigger(), /trigger quota/);
  assert.deepEqual(state.deleted, []);
});

test('authentication rejection releases the lock and does not record successful dispatch', () => {
  const {context, state} = harness({replies: [{status: 401}]});
  assert.throws(() => context.dispatchKaraokeUpdate(), /HTTP 401/);
  assert.equal(state.requests.length, 1);
  assert.equal(state.released, 1);
  assert.equal(state.properties.KARAOKE_LAST_DISPATCH_AT, undefined);
});

test('transient failure retries and honors a short Retry-After header', () => {
  const {context, state} = harness({replies: [{status: 503}, {status: 429, headers: {'Retry-After': '2'}}, {status: 204}]});
  context.dispatchKaraokeUpdate();
  assert.equal(state.requests.length, 3);
  assert.deepEqual(state.sleeps, [5000, 2000]);
  assert.ok(state.properties.KARAOKE_LAST_DISPATCH_AT);
});

test('long rate limit stops without retrying early', () => {
  const {context, state} = harness({replies: [{status: 429, headers: {'retry-after': '120'}}]});
  assert.throws(() => context.dispatchKaraokeUpdate(), /一時的に利用できません/);
  assert.equal(state.requests.length, 1);
  assert.deepEqual(state.sleeps, []);
  assert.equal(state.released, 1);
});

test('overlapping dispatch does not send a duplicate request', () => {
  const {context, state} = harness({locked: true});
  context.dispatchKaraokeUpdate();
  assert.deepEqual(state.requests, []);
  assert.equal(state.released, 0);
});

test('uninstall leaves unrelated GAS triggers untouched', () => {
  const {context, state} = harness({triggers: [trigger('existingCsvHandler'), trigger('dispatchKaraokeUpdate')]});
  context.removeKaraokeHourlyTrigger();
  assert.deepEqual(state.deleted, ['dispatchKaraokeUpdate']);
});
