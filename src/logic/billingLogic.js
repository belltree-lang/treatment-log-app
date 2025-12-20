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
  if (value === 1 || value === '1' || value === true) return 1;
  return 0;
}

function normalizeMedicalSubsidyFlag_(value) {
  return normalizeMedicalAssistanceFlag_(value);
}

function normalizeSelfPayItems_(params) {
  const items = [];
  if (Array.isArray(params && params.selfPayItems)) {
    params.selfPayItems.forEach(entry => {
      if (!entry || typeof entry !== 'object') return;
      const amount = normalizeMoneyNumber_(entry.amount);
      if (!Number.isFinite(amount) || amount === 0) return;
      items.push({ type: entry.type || '自費', amount });
    });
  }

  if (items.length === 0 && params && params.manualSelfPayAmount !== undefined) {
    const manualAmount = normalizeMoneyNumber_(params.manualSelfPayAmount);
    if (Number.isFinite(manualAmount) && manualAmount !== 0) {
      items.push({ type: '自費', amount: manualAmount });
    }
  }

  return items;
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

function normalizePatientIdSortKey_(value) {
  const num = Number(value);
  if (Number.isFinite(num)) return num;
  const trimmed = String(value || '').trim();
  if (!trimmed) return 0;
  const fallback = Number(trimmed.replace(/[^0-9]/g, ''));
  return Number.isFinite(fallback) ? fallback : 0;
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
  if (type === 'マッサージ') return 0;
  const manualUnitPrice = normalizeMoneyNumber_(customUnitPrice);
  const hasManualUnitPrice = Number.isFinite(manualUnitPrice) && manualUnitPrice !== 0;
  if (hasManualUnitPrice) return manualUnitPrice;

  const normalizedMedicalAssistance = normalizeMedicalAssistanceFlag_(medicalAssistance);
  if (normalizedMedicalAssistance === 1) return 0;

  const isLifeProtection = ['生保', '生活保護', '生活扶助'].indexOf(type) >= 0;
  if (isLifeProtection) return 0;

  const normalizedBurdenRate = normalizeBurdenRateInt_(burdenRate);
  const isSelfPaid = type === '自費' || normalizedBurdenRate === '自費';
  if (isSelfPaid) return 0;

  const patientUnitPrice = arguments.length >= 5 ? normalizeMoneyNumber_(arguments[4]) : 0;
  const hasPatientUnitPrice = Number.isFinite(patientUnitPrice) && patientUnitPrice !== 0;
  if (hasPatientUnitPrice) return patientUnitPrice;

  return BILLING_UNIT_PRICE;
}

function calculateBillingAmounts_(params) {
  const visits = normalizeVisitCount_(params.visitCount);
  const insuranceType = String(params.insuranceType || '').trim();
  const normalizedBurdenRate = normalizeBurdenRateInt_(params.burdenRate);
  const medicalAssistance = normalizeMedicalAssistanceFlag_(params.medicalAssistance);
  const manualUnitPrice = normalizeMoneyNumber_(params.manualUnitPrice != null ? params.manualUnitPrice : params.unitPrice);
  const patientUnitPrice = normalizeMoneyNumber_(params.unitPrice);
  const manualTransportInput = Object.prototype.hasOwnProperty.call(params, 'manualTransportAmount')
    ? params.manualTransportAmount
    : params.transportAmount;
  const manualTransportAmount = manualTransportInput === '' || manualTransportInput === null || manualTransportInput === undefined
    ? null
    : normalizeMoneyNumber_(manualTransportInput);
  const unitPrice = resolveInvoiceUnitPrice_(
    insuranceType,
    normalizedBurdenRate,
    manualUnitPrice,
    medicalAssistance,
    patientUnitPrice
  );
  const isSelfPaid = insuranceType === '自費' || normalizedBurdenRate === '自費';
  const hasChargeableUnitPrice = Number.isFinite(unitPrice) && unitPrice !== 0;
  const treatmentAmount = visits > 0 && hasChargeableUnitPrice ? unitPrice * visits : 0;
  const transportAmount = visits > 0 && hasChargeableUnitPrice ? BILLING_TRANSPORT_UNIT_PRICE * visits : 0;
  const resolvedTransportAmount = (manualTransportInput !== '' && manualTransportInput !== null && manualTransportInput !== undefined
    && Number.isFinite(manualTransportAmount))
    ? manualTransportAmount
    : transportAmount;
  const burdenMultiplier = normalizeBurdenMultiplier_(normalizedBurdenRate, insuranceType);
  const carryOverAmount = normalizeMoneyNumber_(params.carryOverAmount);
  const selfPayItems = normalizeSelfPayItems_(params);
  const selfPayTotal = selfPayItems.reduce((sum, entry) => sum + (Number(entry.amount) || 0), 0);
  const billingAmount = isSelfPaid
    ? treatmentAmount
    : roundToNearestTen_(treatmentAmount * burdenMultiplier);
  const total = treatmentAmount + resolvedTransportAmount;
  const grandTotal = billingAmount + resolvedTransportAmount + carryOverAmount + selfPayTotal;

  return {
    visits,
    unitPrice,
    manualUnitPrice,
    treatmentAmount,
    transportAmount: resolvedTransportAmount,
    manualTransportAmount: manualTransportInput === '' || manualTransportInput === null || manualTransportInput === undefined
      ? ''
      : manualTransportAmount,
    carryOverAmount,
    manualSelfPayAmount: params && params.manualSelfPayAmount !== undefined ? params.manualSelfPayAmount : selfPayTotal,
    selfPayItems,
    billingAmount,
    total,
    grandTotal
  };
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
  const patientIds = Object.keys(treatmentVisitCounts || {})
    .sort((a, b) => normalizePatientIdSortKey_(a) - normalizePatientIdSortKey_(b));
  const zeroVisitDebug = [];

  const billingJson = patientIds.map(pid => {
    const patient = patients[pid] || {};
    const isMedicalSubsidy = normalizeMedicalSubsidyFlag_(patient.medicalSubsidy);
    if (isMedicalSubsidy) {
      const name = patient.nameKanji || patient.nameKana || '';
      billingLogger_.log(`[exclude] 患者ID ${pid}${name ? `（${name}）` : ''}は医療助成のため請求対象外`);
      return null;
    }
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
    const responsibleName = responsibleNames.length
      ? responsibleNames[0]
      : (responsibleEmail ? billingResolveStaffDisplayName_(responsibleEmail, staffDirectory) : '');
    const carryOverFromPatient = normalizeMoneyNumber_(patient.carryOverAmount);
    const carryOverFromHistory = normalizeMoneyNumber_(carryOverByPatient[pid]);
    const manualUnitPrice = normalizeMoneyNumber_(patient.manualUnitPrice != null ? patient.manualUnitPrice : patient.unitPrice);
    const patientUnitPrice = normalizeMoneyNumber_(patient.unitPrice);
    const normalizedBurdenRate = normalizeBurdenRateInt_(patient.burdenRate);
    const normalizedMedicalAssistance = normalizeMedicalAssistanceFlag_(patient.medicalAssistance);
    const amountCalc = calculateBillingAmounts_({
      visitCount,
      insuranceType: patient.insuranceType,
      burdenRate: normalizedBurdenRate,
      manualUnitPrice,
      manualTransportAmount: patient.manualTransportAmount,
      manualSelfPayAmount: patient.manualSelfPayAmount,
      selfPayItems: Array.isArray(patient.selfPayItems) ? patient.selfPayItems : [],
      unitPrice: patientUnitPrice,
      medicalAssistance: normalizedMedicalAssistance,
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
      burdenRate: normalizedBurdenRate,
      medicalAssistance: normalizedMedicalAssistance,
      visitCount: amountCalc.visits,
      manualUnitPrice,
      unitPrice: amountCalc.unitPrice,
      treatmentAmount: amountCalc.treatmentAmount,
      manualTransportAmount: amountCalc.manualTransportAmount,
      transportAmount: amountCalc.transportAmount,
      manualSelfPayAmount: amountCalc.manualSelfPayAmount,
      selfPayItems: amountCalc.selfPayItems,
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
  }).filter(Boolean);

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
