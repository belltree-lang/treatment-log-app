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
  return {
    billingMonth,
    patients: source.patients || {},
    treatmentVisitCounts: source.treatmentVisitCounts || {},
    bankStatuses: source.bankStatuses || {}
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
  if (raw <= 1) return raw;
  return raw / 10;
}

function calculateBillingAmounts_(params) {
  const visits = normalizeVisitCount_(params.visitCount);
  const unitPrice = BILLING_UNIT_PRICE;
  const total = visits * unitPrice;
  const insuranceType = String(params.insuranceType || '').trim();
  const burdenMultiplier = normalizeBurdenMultiplier_(params.burdenRate, insuranceType);

  let billingAmount = 0;
  if (insuranceType === '自費') {
    billingAmount = total;
  } else if (insuranceType === '生保' || burdenMultiplier === 0) {
    billingAmount = 0;
  } else {
    billingAmount = Math.round(total * burdenMultiplier);
  }

  const carryOverAmount = normalizeMoneyNumber_(params.carryOverAmount);
  const grandTotal = billingAmount + carryOverAmount;

  return { visits, unitPrice, total, billingAmount, carryOverAmount, grandTotal };
}

function shouldCombineBilling_(bankStatus) {
  const normalized = String(bankStatus || '').trim().toUpperCase();
  return BILLING_COMBINE_STATUSES.indexOf(normalized) >= 0;
}

function generateBillingJsonFromSource(sourceData) {
  const { billingMonth, patients, treatmentVisitCounts, bankStatuses } = normalizeBillingSource_(sourceData);
  const patientIds = Object.keys(treatmentVisitCounts || {});

  return patientIds.map(pid => {
    const patient = patients[pid] || {};
    const bank = bankStatuses[pid] || {};
    const visitCount = normalizeVisitCount_(treatmentVisitCounts[pid]);
    const amountCalc = calculateBillingAmounts_({
      visitCount,
      insuranceType: patient.insuranceType,
      burdenRate: patient.burdenRate,
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
      bankCode: patient.bankCode || '',
      branchCode: patient.branchCode || '',
      accountNumber: patient.accountNumber || '',
      bankStatus,
      carryOverAmount: amountCalc.carryOverAmount,
      grandTotal: amountCalc.grandTotal,
      isNew: patient.isNew ? 1 : 0,
      raw: patient.raw || {},
      shouldCombine: shouldCombineBilling_(bankStatus)
    };
  });
}

function simulateBillingGeneration(sourceData) {
  return generateBillingJsonFromSource(sourceData);
}
