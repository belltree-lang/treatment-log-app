/***** Logic layer: billing JSON generation (pure functions) *****/

if (typeof billingLogger_ === 'undefined') {
  const billingFallbackLog_ = typeof console !== 'undefined' && console && typeof console.log === 'function'
    ? (...args) => console.log(...args)
    : () => {};
  billingLogger_ = { log: billingFallbackLog_ }; // eslint-disable-line no-global-assign
}

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

function normalizeMedicalAssistanceFlag_(value) {
  if (value === null || value === undefined) return 0;
  const num = Number(value);
  if (Number.isFinite(num)) return num ? 1 : 0;
  if (value === true) return 1;
  if (value === false) return 0;
  const text = String(value || '').trim().toLowerCase();
  if (!text) return 0;
  return ['1', 'true', 'yes', 'y', 'on', '有', 'あり', '〇', '○', '◯'].indexOf(text) >= 0 ? 1 : 0;
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
  billingLogger_.log('[billing] patients keys=' + JSON.stringify(Object.keys(patientMap)));
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
  if (String(burdenRate).trim() === '自費') return '自費';
  const num = Number(burdenRate);
  if (Number.isFinite(num)) {
    if (num > 0 && num < 1) return Math.round(num * 10);
    if (num >= 1 && num < 10) return Math.round(num);
    if (num >= 10 && num <= 100) return Math.round(num / 10);
  }

  const normalized = String(burdenRate).normalize('NFKC').replace(/\s+/g, '').replace('％', '%');
  const hasPercent = normalized.indexOf('%') >= 0;
  const numericText = normalized.replace(/[^0-9.]/g, '');
  const parsed = Number(numericText);
  if (!Number.isFinite(parsed)) return 0;
  if (parsed === 0) return 0;
  if (hasPercent) return Math.round(parsed / 10);
  if (parsed > 0 && parsed < 10) return Math.round(parsed);
  if (parsed >= 10 && parsed <= 100) return Math.round(parsed / 10);
  return 0;
}

function resolveInvoiceUnitPrice_(insuranceType, burdenRate, customUnitPrice, medicalAssistance) {
  const type = String(insuranceType || '').trim();
  const manualUnitPrice = normalizeMoneyNumber_(customUnitPrice);
  const hasManualUnitPrice = Number.isFinite(manualUnitPrice) && manualUnitPrice !== 0;
  if (hasManualUnitPrice) return manualUnitPrice;

  const isMedicalAssistance = normalizeMedicalAssistanceFlag_(medicalAssistance) === 1;
  if (isMedicalAssistance) return 0;

  const isLifeProtection = ['生保', '生活保護', '生活扶助'].indexOf(type) >= 0;
  if (isLifeProtection) return 0;

  const normalizedBurdenRate = normalizeBurdenRateInt_(burdenRate);
  const isSelfPaid = type === '自費' || normalizedBurdenRate === '自費';
  if (isSelfPaid) return 0;

  if (type === 'マッサージ') return 0;
  return BILLING_UNIT_PRICE;
}

function calculateBillingAmounts_(params) {
  const visits = normalizeVisitCount_(params.visitCount);
  const insuranceType = String(params.insuranceType || '').trim();
  const normalizedBurdenRate = normalizeBurdenRateInt_(params.burdenRate);
  const isSelfPaid = insuranceType === '自費' || normalizedBurdenRate === '自費';
  const medicalAssistance = normalizeMedicalAssistanceFlag_(params.medicalAssistance);
  const manualUnitPrice = normalizeMoneyNumber_(params.unitPrice);
  const isMassage = insuranceType === 'マッサージ';
  let unitPrice = resolveInvoiceUnitPrice_(insuranceType, params.burdenRate, manualUnitPrice, medicalAssistance);
  if (isMassage) unitPrice = 0;
  const hasManualUnitPrice = Number.isFinite(manualUnitPrice) && manualUnitPrice !== 0;
  const isLifeProtection = ['生保', '生活保護', '生活扶助'].indexOf(insuranceType) >= 0;
  const shouldZero = (isLifeProtection || medicalAssistance) && !hasManualUnitPrice;
  const isZeroCharge = shouldZero || (isSelfPaid && !hasManualUnitPrice);
  const treatmentAmount = visits > 0 && !isMassage && !isZeroCharge ? unitPrice * visits : 0;
  const transportAmount = visits > 0 && !isMassage && !isZeroCharge ? BILLING_TRANSPORT_UNIT_PRICE * visits : 0;
  const burdenMultiplier = normalizeBurdenMultiplier_(normalizedBurdenRate, insuranceType);
  const carryOverAmount = normalizeMoneyNumber_(params.carryOverAmount);
  const billingAmount = isSelfPaid
    ? treatmentAmount
    : roundToNearestTen_(treatmentAmount * burdenMultiplier);
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
  const zeroVisitDebug = [];

  const billingJson = patientIds.map(pid => {
    const patient = patients[pid] || {};
    const rawVisitCount = treatmentVisitCounts[pid];
    const visitCount = normalizeVisitCount_(rawVisitCount);
    if (!visitCount && zeroVisitDebug.length < 20) {
      zeroVisitDebug.push({
        patientId: pid,
        rawVisitCount,
        normalizedVisitCount: visitCount,
        payerType: patient.payerType,
        carryOverAmount: patient.carryOverAmount
      });
    }
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
      medicalAssistance: patient.medicalAssistance,
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
      medicalAssistance: normalizeMedicalAssistanceFlag_(patient.medicalAssistance),
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

  billingLogger_.log('[billing] raw billingJson=' + JSON.stringify(billingJson));
  billingLogger_.log('[billing] generateBillingJsonFromSource: zero visit samples=' + JSON.stringify(zeroVisitDebug));
  if (!billingJson.length) {
    billingLogger_.log('[billing] generateBillingJsonFromSource: empty billingJson with visitCountKeys=' + JSON.stringify(patientIds));
  }
  billingLogger_.log('[billing] billingJson length=' + billingJson.length);
  return billingJson;
}

function simulateBillingGeneration(sourceData) {
  return generateBillingJsonFromSource(sourceData);
}
