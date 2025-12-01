/***** Logic layer: billing JSON generation (pure functions) *****/

const BILLING_TREATMENT_PRICE = 4070;
const BILLING_ELECTRO_PRICE = 100;
const BILLING_UNIT_PRICE = BILLING_TREATMENT_PRICE + BILLING_ELECTRO_PRICE;
const BILLING_TRANSPORT_UNIT_PRICE_FALLBACK = 33;
const BILLING_TRANSPORT_UNIT_PRICE = (typeof globalThis !== 'undefined' && typeof globalThis.BILLING_TRANSPORT_UNIT_PRICE === 'number')
  ? globalThis.BILLING_TRANSPORT_UNIT_PRICE
  : BILLING_TRANSPORT_UNIT_PRICE_FALLBACK;
const BILLING_TREATMENT_UNIT_PRICE_BY_BURDEN = { 1: 417, 2: 834, 3: 1251 };

const billingResolveStaffDisplayName_ = typeof resolveStaffDisplayName_ === 'function'
  ? resolveStaffDisplayName_
  : function fallbackResolveStaffDisplayName_(email, directory) {
    const normalized = String(email || '').trim();
    if (!normalized) return '';
    const key = normalized.toLowerCase();
    if (directory && directory[key]) return directory[key];
    const parts = normalized.split('@');
    return parts[0] || normalized;
  };

function roundToNearestTen_(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.round(num / 10) * 10;
}

function normalizeBurdenMultiplier_(burdenRate, insuranceType) {
  const type = String(insuranceType || '').trim();
  if (type === '自費') return 1;
  const rateInt = normalizeBurdenRateInt_(burdenRate);
  return rateInt > 0 ? rateInt / 10 : 0;
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
  const staffDirectory = source.staffDirectory || {};
  const staffDisplayByPatient = source.staffDisplayByPatient || {};
  const bankStatuses = source.bankStatuses || {};
  const carryOverByPatient = source.carryOverByPatient || {};
  return {
    billingMonth,
    patients: patientMap,
    treatmentVisitCounts: visitCounts,
    staffByPatient,
    staffDirectory,
    staffDisplayByPatient,
    bankStatuses,
    carryOverByPatient
  };
}

function normalizeVisitCount_(value) {
  const source = value && value.visitCount != null ? value.visitCount : value;
  if (typeof source === 'number') {
    return Number.isFinite(source) && source > 0 ? source : 0;
  }
  const normalized = String(source || '')
    .normalize('NFKC')
    .replace(/[，,]/g, '')
    .trim();
  const num = Number(normalized);
  return Number.isFinite(num) && num > 0 ? num : 0;
}

function normalizeMoneyNumber_(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }
  const text = String(value || '')
    .normalize('NFKC')
    .replace(/[，,]/g, '')
    .trim();
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
  if (type === '自費') return 0;
  if (type === '生保' || type === 'マッサージ') return 0;
  return BILLING_UNIT_PRICE;
}

function calculateBillingAmounts_(params) {
  const visits = normalizeVisitCount_(params.visitCount);
  const insuranceType = String(params.insuranceType || '').trim();
  const unitPrice = resolveInvoiceUnitPrice_(insuranceType, params.burdenRate, params.unitPrice);
  const isMassage = insuranceType === 'マッサージ';
  const isZeroCharge = insuranceType === '生保' || insuranceType === '自費';
  const treatmentAmount = visits > 0 && !isMassage && !isZeroCharge ? unitPrice * visits : 0;
  const transportAmount = visits > 0 && !isMassage && !isZeroCharge ? BILLING_TRANSPORT_UNIT_PRICE * visits : 0;
  const burdenMultiplier = normalizeBurdenMultiplier_(params.burdenRate, insuranceType);
  const carryOverAmount = normalizeMoneyNumber_(params.carryOverAmount);
  const billingAmount = roundToNearestTen_(treatmentAmount * burdenMultiplier);
  const total = treatmentAmount + transportAmount;
  const grandTotal = billingAmount + transportAmount + carryOverAmount;

  return { visits, unitPrice, treatmentAmount, transportAmount, carryOverAmount, billingAmount, total, grandTotal };
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
  const {
    billingMonth,
    patients,
    treatmentVisitCounts,
    staffByPatient,
    staffDirectory,
    staffDisplayByPatient,
    bankStatuses,
    carryOverByPatient
  } = normalizeBillingSource_(sourceData);
  const patientIds = Object.keys(treatmentVisitCounts || {});

  return patientIds.map(pid => {
    const patient = patients[pid] || {};
    const visitCount = normalizeVisitCount_(treatmentVisitCounts[pid]);
    const staffEmails = Array.isArray(staffByPatient[pid]) ? staffByPatient[pid] : (staffByPatient[pid] ? [staffByPatient[pid]] : []);
    const resolvedStaffNames = Array.isArray(staffDisplayByPatient[pid]) ? staffDisplayByPatient[pid] : [];
    const responsibleNames = resolvedStaffNames.length
      ? resolvedStaffNames
      : staffEmails.map(email => billingResolveStaffDisplayName_(email, staffDirectory)).filter(Boolean);
    const responsibleEmail = staffEmails.length ? staffEmails[0] : '';
    const responsibleName = responsibleNames.join('・');
    const carryOverFromPatient = normalizeMoneyNumber_(patient.carryOverAmount);
    const carryOverFromHistory = normalizeMoneyNumber_(carryOverByPatient[pid]);
    const amountCalc = calculateBillingAmounts_({
      visitCount,
      insuranceType: patient.insuranceType,
      burdenRate: patient.burdenRate,
      unitPrice: patient.unitPrice,
      carryOverAmount: carryOverFromPatient + carryOverFromHistory
    });

    const bankStatusEntry = bankStatuses && bankStatuses[pid];

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
      carryOverFromHistory,
      billingAmount: amountCalc.billingAmount,
      total: amountCalc.total,
      grandTotal: amountCalc.grandTotal,
      responsibleEmail,
      responsibleNames,
      responsibleName,
      payerType: patient.payerType || '',
      bankStatus: bankStatusEntry && bankStatusEntry.bankStatus ? bankStatusEntry.bankStatus : '',
      paidStatus: bankStatusEntry && bankStatusEntry.paidStatus ? bankStatusEntry.paidStatus : ''
    };
  });
}

function simulateBillingGeneration(sourceData) {
  return generateBillingJsonFromSource(sourceData);
}
