const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const code = fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8');
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(code, sandbox);

sandbox.normId_ = v => String(v || '').trim();
sandbox.PATIENT_CACHE_KEYS = { header: pid => `header:${pid}` };
sandbox.PATIENT_CACHE_TTL_SECONDS = 60;
sandbox.cacheHit_ = () => false;
sandbox.cacheFetch_ = (_key, cb) => cb();
sandbox.ensureAuxSheets_ = () => {};
sandbox.findPatientRow_ = () => ({
  head: ['施術録番号', '名前', '病院名', '医師', 'ﾌﾘｶﾞﾅ', '生年月日', '同意年月日', '配布', '負担割合', '電話', '同意症状'],
  rowValues: ['P001', '患者 太郎', 'A病院', '医師 花子', 'カンジャ タロウ', '', '', '', '1割', '000-0000-0000', '']
});
sandbox.sh = () => ({ getName: () => '患者情報' });
sandbox.getColFlexible_ = (head, _labels, fixed) => fixed;
sandbox.LABELS = {
  name: [], hospital: [], doctor: [], furigana: [], birth: [], consent: [],
  consentHandout: [], share: [], phone: [], consentContent: []
};
sandbox.PATIENT_COLS_FIXED = {
  name: 2, hospital: 3, doctor: 4, furigana: 5, birth: 6,
  consent: 7, consentHandout: 8, share: 9, phone: 10, consentContent: 11
};
sandbox.parseDateFlexible_ = () => null;
sandbox.calcConsentExpiry_ = () => '';
sandbox.normalizeBurdenRatio_ = value => value;
sandbox.toBurdenDisp_ = value => value;
sandbox.getRecentActivity_ = () => ({ lastTreat: '', lastConsent: '', lastStaff: '' });
sandbox.getStatus_ = () => ({ status: 'active', pauseUntil: '' });
sandbox.resolveYearMonthOrCurrent_ = (year, month) => {
  if (typeof year === 'number' && typeof month === 'number') {
    const base = new Date(year, month - 1, 1);
    return { year: base.getFullYear(), month: base.getMonth() + 1 };
  }
  return { year: 2025, month: 3 };
};
sandbox.APP = { BASE_FEE_YEN: 4170 };

const monthlyCalls = [];
sandbox.listTreatmentsForMonth = (pid, year, month) => {
  monthlyCalls.push({ pid, year, month });
  if (year === 2025 && month === 3) return [{}, {}, {}];
  if (year === 2025 && month === 2) return [{}];
  return [];
};

const header = sandbox.getPatientHeader('P001');

assert.strictEqual(header.monthly.current.count, 3);
assert.strictEqual(header.monthly.previous.count, 1);
assert.strictEqual(header.monthly.current.est, Math.round(3 * 4170 * 0.1));
assert.strictEqual(header.monthly.previous.est, Math.round(1 * 4170 * 0.1));
assert.deepStrictEqual(monthlyCalls, [
  { pid: 'P001', year: 2025, month: 3 },
  { pid: 'P001', year: 2025, month: 2 }
]);

console.log('getPatientHeader monthly summary tests passed');
