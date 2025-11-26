/***** Logic layer: billing JSON generation (pure functions) *****/

const BILLING_TREATMENT_PRICE = 4070;
const BILLING_ELECTRO_PRICE = 100;
const BILLING_UNIT_PRICE = BILLING_TREATMENT_PRICE + BILLING_ELECTRO_PRICE;
const BILLING_COMBINE_STATUSES = ['NO_DOCUMENT', 'INSUFFICIENT', 'NOT_FOUND'];

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
  return {
    billingMonth,
    patients: patientMap,
    treatmentVisitCounts: visitCounts,
    bankStatuses: source.bankStatuses || {},
    bankInfoByName: source.bankInfoByName || {}
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

function normalizeBurdenMultiplier_(burdenRate, insuranceType) {
  if (String(insuranceType || '').trim() === '自費') return 1;
  const raw = Number(burdenRate);
  if (!Number.isFinite(raw) || raw <= 0) return 0;
  if (raw < 1) return raw;
  if (raw < 10) return raw / 10;
  return raw / 10;
}

function resolveBillingUnitPrice_(params) {
  const custom = normalizeMoneyNumber_(params.unitPrice);
  return custom > 0 ? custom : BILLING_UNIT_PRICE;
}

function calculateBillingAmounts_(params) {
  const visits = normalizeVisitCount_(params.visitCount);
  const insuranceType = String(params.insuranceType || '').trim();
  const unitPrice = insuranceType === 'マッサージ' ? 0 : resolveBillingUnitPrice_(params);
  const total = visits * unitPrice;
  const burdenMultiplier = normalizeBurdenMultiplier_(params.burdenRate, insuranceType);

  let billingAmount = 0;
  if (insuranceType === '自費') {
    billingAmount = visits * unitPrice;
  } else if (insuranceType === '生保' || insuranceType === 'マッサージ' || burdenMultiplier === 0) {
    billingAmount = 0;
  } else {
    billingAmount = Math.round(visits * unitPrice * burdenMultiplier);
  }

  const carryOverAmount = normalizeMoneyNumber_(params.carryOverAmount);
  const grandTotal = billingAmount + carryOverAmount;

  return { visits, unitPrice, total, billingAmount, carryOverAmount, grandTotal };
}

function shouldCombineBilling_(bankStatus) {
  const normalized = String(bankStatus || '').trim().toUpperCase();
  return BILLING_COMBINE_STATUSES.indexOf(normalized) >= 0;
}

function resolveBankRecordForPatient_(patient, bankInfoByName) {
  const nameKey = normalizeBillingNameKey_(patient && patient.nameKanji);
  if (!nameKey) return null;
  return bankInfoByName && bankInfoByName[nameKey];
}

function buildBankJoinResult_(patient, bankInfoByName) {
  const bankRecord = resolveBankRecordForPatient_(patient, bankInfoByName);
  const nameLabel = patient && patient.nameKanji ? patient.nameKanji : '';
  if (!bankRecord) {
    return {
      bankJoinError: true,
      bankJoinMessage: '銀行情報が不足しています（氏名：' + nameLabel + '）',
      bankCode: '',
      branchCode: '',
      accountNumber: '',
      regulationCode: '',
      isNew: '',
      bankRecord: null
    };
  }

  const missing = ['bankCode', 'branchCode', 'accountNumber'].filter(field => !bankRecord[field]);
  const bankJoinError = missing.length > 0;
  const bankJoinMessage = bankJoinError
    ? '銀行情報が不足しています（氏名：' + nameLabel + '、不足項目: ' + missing.join(', ') + '）'
    : '';

  return {
    bankJoinError,
    bankJoinMessage,
    bankCode: bankRecord.bankCode || '',
    branchCode: bankRecord.branchCode || '',
    accountNumber: bankRecord.accountNumber || '',
    regulationCode: bankRecord.regulationCode != null ? bankRecord.regulationCode : 1,
    isNew: bankRecord.isNew ? 1 : 0,
    bankRecord
  };
}

function generateBillingJsonFromSource(sourceData) {
  const { billingMonth, patients, treatmentVisitCounts, bankStatuses, bankInfoByName } = normalizeBillingSource_(sourceData);
  const patientIds = Object.keys(treatmentVisitCounts || {});

  return patientIds.map(pid => {
    const patient = patients[pid] || {};
    const bank = bankStatuses[pid] || {};
    const bankJoin = buildBankJoinResult_(patient, bankInfoByName);
    const visitCount = normalizeVisitCount_(treatmentVisitCounts[pid]);
    const amountCalc = calculateBillingAmounts_({
      visitCount,
      insuranceType: patient.insuranceType,
      burdenRate: patient.burdenRate,
      unitPrice: patient.unitPrice,
      carryOverAmount: patient.carryOverAmount
    });

    const bankStatus = bank.bankStatus || '';

    return {
      billingMonth,
      patientId: pid,
      nameKanji: patient.nameKanji || '',
      nameKana: patient.nameKana || '',
      insuranceType: patient.insuranceType || '',
      burdenRate: Number(patient.burdenRate) || 0,
      visitCount: amountCalc.visits,
      unitPrice: amountCalc.unitPrice,
      total: amountCalc.total,
      billingAmount: amountCalc.billingAmount,
      bankCode: bankJoin.bankCode,
      branchCode: bankJoin.branchCode,
      regulationCode: bankJoin.regulationCode,
      accountNumber: bankJoin.accountNumber,
      bankStatus,
      carryOverAmount: amountCalc.carryOverAmount,
      grandTotal: amountCalc.grandTotal,
      isNew: bankJoin.isNew,
      bankJoinError: bankJoin.bankJoinError,
      bankJoinMessage: bankJoin.bankJoinMessage,
      raw: bankJoin.bankRecord ? Object.assign({}, bankJoin.bankRecord.raw || {}, patient.raw || {}) : (patient.raw || {}),
      shouldCombine: shouldCombineBilling_(bankStatus)
    };
  });
}

function simulateBillingGeneration(sourceData) {
  return generateBillingJsonFromSource(sourceData);
}
