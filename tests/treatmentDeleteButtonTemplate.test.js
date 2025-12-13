const fs = require('fs');
const path = require('path');
const assert = require('assert');

const html = fs.readFileSync(path.join(__dirname, '../src/app.html'), 'utf8');

const match = html.match(/function renderThisMonthList[\s\S]*?return `\n        <tr>[\s\S]*?<\/tr>`;/);
assert.ok(match, 'renderThisMonthList template should be present');

const template = match[0];

assert.ok(
  template.includes('data-treatment-id="${treatmentIdAttr}"'),
  'delete button should carry treatment id via data attribute'
);

assert.ok(
  template.includes('onclick="delRow(this.dataset.treatmentId)"'),
  'delete button should read treatment id from dataset'
);

assert.ok(
  template.includes("const treatmentIdAttr = escapeHtml(String(r.treatmentId || ''))"),
  'treatment id should be escaped when embedded in markup'
);

console.log('treatmentDeleteButtonTemplate test passed');
