/***** Logic layer: billing JSON generation (pure functions) *****/

const BILLING_TREATMENT_PRICE = 4070;
const BILLING_ELECTRO_PRICE = 100;
const BILLING_UNIT_PRICE = BILLING_TREATMENT_PRICE + BILLING_ELECTRO_PRICE;
const BILLING_TRANSPORT_UNIT_PRICE = 33;
const BILLING_TREATMENT_UNIT_PRICE_BY_BURDEN = { 1: 417, 2: 834, 3: 1251 };

const billingResolveStaffDisplayName_ = typeof resolveStaffDisplayName_ === 'function'
  ? resolveStaffDisplayName_
  : function fallbackResolveStaffDisplayName_(email) {
    const normalized = String(email || '').trim();
    if (!normalized) return '';
    const parts = normalized.split('@');
    return parts[0] || normalized;
  };

function roundToNearestTen_(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.round(num / 10) * 10;
}

function normalizeBillingSource_(source) {
  if (!source || typeof source !== 'object') {
    throw new Error('請求生成の入力が不正です');
  }
  const billingMonth = source.billingMonth || (source.month && source.month.key);
  if (!billingMonth) {
    throw new Error('請求月が指定されていません');
  }
  const patientMap = source.patients || source.patientMap || {};
  const visitCounts = source.treatmentVisitCounts || source.visitCounts || {};
  const staffByPatient = source.staffByPatient || {};
  return {
    billingMonth,
    patients: patientMap,
    treatmentVisitCounts: visitCounts,
    staffByPatient
  };
}

function normalizeVisitCount_(value) {
  const num = Number(value && value.visitCount != null ? value.visitCount : value);
  return Number.isFinite(num) && num > 0 ? num : 0;
}

function normalizeMoneyNumber_(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }
  const text = String(value || '').replace(/,/g, '').trim();
  if (!text) return 0;
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
}

function normalizeBurdenRateInt_(burdenRate) {
  if (burdenRate == null || burdenRate === '') return 0;
  const num = Number(burdenRate);
  if (Number.isFinite(num)) {
    if (num > 0 && num < 1) return Math.round(num * 10);
    if (num >= 1 && num < 10) return Math.round(num);
    if (num >= 10 && num <= 100) return Math.round(num / 10);
  }

  const normalized = String(burdenRate).normalize('NFKC').replace(/\s+/g, '').replace('％', '%');
  const withoutUnits = normalized.replace(/割|分/g, '').replace('%', '');
  const parsed = Number(withoutUnits);
  if (!Number.isFinite(parsed)) return 0;
  if (normalized.indexOf('%') >= 0) return Math.round(parsed / 10);
  if (parsed > 0 && parsed < 10) return Math.round(parsed);
  if (parsed >= 10 && parsed <= 100) return Math.round(parsed / 10);
  return 0;
}

function resolveInvoiceUnitPrice_(insuranceType, burdenRate, customUnitPrice) {
  const type = String(insuranceType || '').trim();
  if (type === '自費') {
    const custom = normalizeMoneyNumber_(customUnitPrice);
    return custom > 0 ? custom : 0;
  }
  if (type === '生保' || type === 'マッサージ') return 0;
  const normalizedRate = normalizeBurdenRateInt_(burdenRate);
  return BILLING_TREATMENT_UNIT_PRICE_BY_BURDEN[normalizedRate] || BILLING_UNIT_PRICE;
}

function calculateBillingAmounts_(params) {
  const visits = normalizeVisitCount_(params.visitCount);
  const insuranceType = String(params.insuranceType || '').trim();
  const unitPrice = resolveInvoiceUnitPrice_(insuranceType, params.burdenRate, params.unitPrice);
  const treatmentAmount = visits > 0 ? unitPrice * visits : 0;
  const transportAmount = visits > 0 ? BILLING_TRANSPORT_UNIT_PRICE * visits : 0;
  const carryOverAmount = normalizeMoneyNumber_(params.carryOverAmount);
  const grandTotal = treatmentAmount + transportAmount + carryOverAmount;

  return { visits, unitPrice, treatmentAmount, transportAmount, carryOverAmount, grandTotal };
}

function resolveBillingAddress_(patient) {
  if (!patient) return '';
  if (patient.address) return String(patient.address).trim();
  const raw = patient.raw || {};
  const candidates = ['住所', '住所1', '住所２', '住所2', 'address', 'Address'];
  for (let i = 0; i < candidates.length; i++) {
    const key = candidates[i];
    if (raw.hasOwnProperty(key) && raw[key] != null && String(raw[key]).trim()) {
      return String(raw[key]).trim();
    }
  }
  return '';
}

function generateBillingJsonFromSource(sourceData) {
  const { billingMonth, patients, treatmentVisitCounts, staffByPatient } = normalizeBillingSource_(sourceData);
  const patientIds = Object.keys(treatmentVisitCounts || {});

  return patientIds.map(pid => {
    const patient = patients[pid] || {};
    const visitCount = normalizeVisitCount_(treatmentVisitCounts[pid]);
    const responsibleEmail = staffByPatient[pid] || '';
    const responsibleName = billingResolveStaffDisplayName_(responsibleEmail);
    const amountCalc = calculateBillingAmounts_({
      visitCount,
      insuranceType: patient.insuranceType,
      burdenRate: patient.burdenRate,
      unitPrice: patient.unitPrice,
      carryOverAmount: patient.carryOverAmount
    });

    return {
      billingMonth,
      patientId: pid,
      nameKanji: patient.nameKanji || '',
      nameKana: patient.nameKana || '',
      address: resolveBillingAddress_(patient),
      insuranceType: patient.insuranceType || '',
      burdenRate: normalizeBurdenRateInt_(patient.burdenRate),
      visitCount: amountCalc.visits,
      unitPrice: amountCalc.unitPrice,
      treatmentAmount: amountCalc.treatmentAmount,
      transportAmount: amountCalc.transportAmount,
      carryOverAmount: amountCalc.carryOverAmount,
      grandTotal: amountCalc.grandTotal,
      responsibleEmail,
      responsibleName
    };
  });
}

function simulateBillingGeneration(sourceData) {
  return generateBillingJsonFromSource(sourceData);
}
