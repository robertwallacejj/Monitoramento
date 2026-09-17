const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const root = path.resolve(__dirname, '..');
const context = {
  window: {},
  document: {},
  console,
  localStorage: {
    getItem() { return null; },
    setItem() {},
    removeItem() {}
  },
  XLSX: { SSF: { parse_date_code() { return null; } } }
};
context.window = context;

const utils = fs.readFileSync(path.join(root, 'assets/js/core/utils.js'), 'utf8');
const domain = fs.readFileSync(path.join(root, 'assets/js/core/domain.js'), 'utf8');
const metrics = fs.readFileSync(path.join(root, 'assets/js/modules/insucessos-metrics.js'), 'utf8');

vm.runInNewContext(utils, context);
vm.runInNewContext(domain, context);
vm.runInNewContext(metrics, context);

const metricsApi = context.window.CTInsucessosMetrics;
assert.ok(metricsApi && typeof metricsApi.buildGeneralSlaSummary === 'function', 'buildGeneralSlaSummary deve existir');

const sample = [
  { status: 'insucesso', reason: 'Devolução' },
  { status: 'insucesso', reason: 'Devolução' },
  { status: 'insucesso', reason: 'Recusa' },
  { status: 'entregue', reason: '' },
  { status: 'pendente', reason: '' }
];

const summary = metricsApi.buildGeneralSlaSummary(sample);
assert.strictEqual(summary.totalExpedido, 5, 'Total Expedido deve contar as linhas válidas');
assert.strictEqual(summary.totalInsucessos, 3, 'Total de insucessos deve contabilizar apenas falhas');
assert.deepStrictEqual(summary.reasonBreakdown.map((item) => item.label), ['Devolução', 'Recusa']);
assert.ok(Math.abs(summary.reasonBreakdown[0].percent - 66.66666666666666) < 1e-9, 'Devolução deve representar 66.67%');
assert.ok(Math.abs(summary.reasonBreakdown[1].percent - 33.33333333333333) < 1e-9, 'Recusa deve representar 33.33%');

const empty = metricsApi.buildGeneralSlaSummary([]);
assert.strictEqual(empty.totalExpedido, 0);
assert.strictEqual(empty.totalInsucessos, 0);
assert.deepStrictEqual(empty.reasonBreakdown, []);

console.log('OK: regra de SLA validada');
