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

function selectLatestStaffFromHistory_(patientId, staffHistoryByPatient) {
  const history = staffHistoryByPatient && staffHistoryByPatient[patientId];
  if (!history || typeof history !== 'object') return null;

  return Object.keys(history).reduce((latest, key) => {
    const entry = history[key];
    if (!entry || !(entry.timestamp instanceof Date) || isNaN(entry.timestamp.getTime())) return latest;

    if (!latest) return entry;
    const latestTime = latest.timestamp instanceof Date ? latest.timestamp.getTime() : 0;
    const entryTime = entry.timestamp.getTime();
    return entryTime > latestTime ? entry : latest;
  }, null);
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
  const staffHistoryByPatient = source.staffHistoryByPatient || {};
  const bankStatuses = source.bankStatuses || {};
  const carryOverByPatient = source.carryOverByPatient || {};
  const bankFlagsByPatient = source.bankFlagsByPatient || {};
  billingLogger_.log('[billing] patients keys=' + JSON.stringify(Object.keys(patientMap)));
  return {
    billingMonth,
    patients: patientMap,
    treatmentVisitCounts: visitCounts,
    staffByPatient,
    staffDirectory,
    staffDisplayByPatient,
    staffHistoryByPatient,
    bankStatuses,
    carryOverByPatient,
    bankFlagsByPatient
  };
}

function resolveOnlineFeeFlag_(patientId, bankFlagsByPatient) {
  const pid = String(patientId || '').trim();
  if (!pid) return false;
  const flags = bankFlagsByPatient && bankFlagsByPatient[pid];
  return !!(flags && flags.ag);
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

function normalizeSelfPayCount_(value) {
  if (!value || typeof value !== 'object') return 0;
  if (Object.prototype.hasOwnProperty.call(value, 'selfPayVisitCount')) {
    const explicitCount = Number(value.selfPayVisitCount);
    return Number.isFinite(explicitCount) && explicitCount > 0 ? explicitCount : 0;
  }
  const self30 = Number(value.self30) || 0;
  const self60 = Number(value.self60) || 0;
  const mixed = Number(value.mixed) || 0;
  const self30Only = Math.max(0, self30 - mixed);
  return self30Only + self60 + mixed;
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

function resolveInvoiceUnitPrice_(insuranceType, burdenRate, customUnitPrice, medicalAssistance, patientUnitPrice, options) {
  const type = String(insuranceType || '').trim();
  if (type === 'マッサージ') return 0;
  const manualUnitPrice = normalizeMoneyNumber_(customUnitPrice);
  const hasManualUnitPrice = Number.isFinite(manualUnitPrice) && manualUnitPrice !== 0;
  if (hasManualUnitPrice) return manualUnitPrice;

  const normalizedMedicalAssistance = normalizeMedicalAssistanceFlag_(medicalAssistance);
  if (normalizedMedicalAssistance === 1) return 0;

  const isLifeProtection = ['生保', '生活保護', '生活扶助'].indexOf(type) >= 0;
  if (isLifeProtection) return 0;

  const ignoreSelfPayInsuranceType = options && options.ignoreSelfPayInsuranceType;
  const normalizedBurdenRate = normalizeBurdenRateInt_(burdenRate);
  const isSelfPaid = (type === '自費' && !ignoreSelfPayInsuranceType) || normalizedBurdenRate === '自費';
  if (isSelfPaid) return 0;

  const normalizedPatientUnitPrice = normalizeMoneyNumber_(patientUnitPrice);
  const hasPatientUnitPrice = Number.isFinite(normalizedPatientUnitPrice) && normalizedPatientUnitPrice !== 0;
  if (hasPatientUnitPrice) return normalizedPatientUnitPrice;

  return BILLING_UNIT_PRICE;
}

function calculateBillingAmounts_(params) {
  const visitCountSource = params && params.visitCount;
  const visits = normalizeVisitCount_(visitCountSource);
  const insuranceType = String(params.insuranceType || '').trim();
  const normalizedBurdenRate = normalizeBurdenRateInt_(params.burdenRate);
  const hasMixedVisitCount = visitCountSource && typeof visitCountSource === 'object'
    && Number(visitCountSource.mixed) > 0;
  const ignoreSelfPayInsuranceType = !!(params && params.ignoreSelfPayInsuranceType) || hasMixedVisitCount;
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
    patientUnitPrice,
    { ignoreSelfPayInsuranceType }
  );
  const isSelfPaid = !ignoreSelfPayInsuranceType
    && (insuranceType === '自費' || normalizedBurdenRate === '自費');
  const hasChargeableUnitPrice = Number.isFinite(unitPrice) && unitPrice !== 0;
  const treatmentAmount = visits > 0 && hasChargeableUnitPrice ? unitPrice * visits : 0;
  const transportAmount = visits > 0 && hasChargeableUnitPrice ? BILLING_TRANSPORT_UNIT_PRICE * visits : 0;
  const resolvedTransportAmount = (manualTransportInput !== '' && manualTransportInput !== null && manualTransportInput !== undefined
    && Number.isFinite(manualTransportAmount))
    ? manualTransportAmount
    : transportAmount;
  const burdenMultiplier = normalizeBurdenMultiplier_(
    normalizedBurdenRate,
    ignoreSelfPayInsuranceType ? '' : insuranceType
  );
  const carryOverAmount = normalizeMoneyNumber_(params.carryOverAmount);
  const selfPayItems = normalizeSelfPayItems_(params);
  const selfPayTotal = selfPayItems.reduce((sum, entry) => sum + (Number(entry.amount) || 0), 0);
  const billingAmount = isSelfPaid
    ? treatmentAmount
    : treatmentAmount * burdenMultiplier;
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

function normalizeBillingEntryType_(entryType) {
  if (!entryType) return '';
  const rawType = String(entryType || '').trim();
  const lower = rawType.toLowerCase();
  if (lower === 'insurance') return 'insurance';
  if (lower === 'selfpay' || lower === 'self_pay' || lower === 'self-pay') return 'self_pay';
  if (rawType === 'selfPay') return 'self_pay';
  return rawType;
}

function shouldApplyOverrideForEntryType_(overrideEntryType, targetEntryType) {
  const normalizedOverride = normalizeBillingEntryType_(overrideEntryType);
  const normalizedTarget = normalizeBillingEntryType_(targetEntryType);
  if (!normalizedOverride || !normalizedTarget) return false;
  return normalizedOverride === normalizedTarget;
}

function resolveEntryTotalAmount_(entry) {
  if (!entry) return 0;
  const manualOverride = entry.manualOverride && entry.manualOverride.amount;
  const normalized = normalizeMoneyNumber_(manualOverride);
  if (manualOverride !== '' && manualOverride !== null && manualOverride !== undefined) {
    return normalized;
  }
  return normalizeMoneyNumber_(entry.total);
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
    staffHistoryByPatient,
    bankStatuses,
    carryOverByPatient,
    bankFlagsByPatient
  } = normalizeBillingSource_(sourceData);
  const patientIdSet = new Set(Object.keys(treatmentVisitCounts || {}));
  Object.keys(bankFlagsByPatient || {}).forEach(pid => {
    if (resolveOnlineFeeFlag_(pid, bankFlagsByPatient)) {
      patientIdSet.add(pid);
    }
  });
  const patientIds = Array.from(patientIdSet)
    .sort((a, b) => normalizePatientIdSortKey_(a) - normalizePatientIdSortKey_(b));
  const zeroVisitDebug = [];

  const billingJson = patientIds.map(pid => {
    const patient = patients[pid] || {};
    const hasOnlineFee = resolveOnlineFeeFlag_(pid, bankFlagsByPatient);
    const consentFlag = typeof normalizeZeroOneFlag_ === 'function'
      ? normalizeZeroOneFlag_(patient.onlineConsent)
      : (patient.onlineConsent === true || patient.onlineConsent === 1 || patient.onlineConsent === '1' ? 1 : 0);
    if (!consentFlag && hasOnlineFee) {
      const errorMessage = `[billing] online consent mismatch: patientId=${pid} onlineConsent=false but online_fee=true`;
      if (typeof console !== 'undefined' && console && typeof console.error === 'function') {
        console.error(errorMessage);
      } else {
        billingLogger_.log(errorMessage);
      }
    }
    const isMedicalSubsidy = normalizeMedicalSubsidyFlag_(patient.medicalSubsidy);
    const billingItems = hasOnlineFee
      ? [{ type: 'online_fee', label: 'オンライン同意サービス使用料', amount: 1000 }]
      : [];
    if (isMedicalSubsidy && !hasOnlineFee) {
      const name = patient.nameKanji || patient.nameKana || '';
      billingLogger_.log(`[exclude] 患者ID ${pid}${name ? `（${name}）` : ''}は医療助成のため請求対象外`);
      return null;
    }
    const rawVisitCount = treatmentVisitCounts[pid];
    const visitCount = normalizeVisitCount_(rawVisitCount);
    const selfPayVisitCount = normalizeSelfPayCount_(rawVisitCount);
    const hasMixedVisitCount = rawVisitCount && typeof rawVisitCount === 'object'
      && Number(rawVisitCount.mixed) > 0;
    if (!visitCount && zeroVisitDebug.length < 20) {
      zeroVisitDebug.push({
        patientId: pid,
        rawVisitCount,
        normalizedVisitCount: visitCount,
        payerType: patient.payerType,
        carryOverAmount: patient.carryOverAmount
      });
    }
    const latestStaffEntry = selectLatestStaffFromHistory_(pid, staffHistoryByPatient);
    const responsibleEmail = latestStaffEntry && (latestStaffEntry.email || latestStaffEntry.key)
      ? latestStaffEntry.email || latestStaffEntry.key || ''
      : '';
    const responsibleName = responsibleEmail
      ? billingResolveStaffDisplayName_(responsibleEmail, staffDirectory) || ''
      : '';
    const responsibleNames = responsibleName ? [responsibleName] : [];
    const carryOverFromPatient = isMedicalSubsidy && hasOnlineFee
      ? 0
      : normalizeMoneyNumber_(patient.carryOverAmount);
    const carryOverFromHistory = isMedicalSubsidy && hasOnlineFee
      ? 0
      : normalizeMoneyNumber_(carryOverByPatient[pid]);
    const manualUnitPriceEntryType = normalizeBillingEntryType_(patient.manualUnitPriceEntryType);
    const manualBillingAmountEntryType = normalizeBillingEntryType_(patient.manualBillingAmountEntryType);
    const manualSelfPayAmountEntryType = normalizeBillingEntryType_(patient.manualSelfPayAmountEntryType);
    const manualUnitPriceValue = normalizeMoneyNumber_(patient.manualUnitPrice != null ? patient.manualUnitPrice : patient.unitPrice);
    const manualUnitPriceAppliesToInsurance = shouldApplyOverrideForEntryType_(manualUnitPriceEntryType, 'insurance');
    const manualUnitPrice = manualUnitPriceAppliesToInsurance
      ? manualUnitPriceValue
      : normalizeMoneyNumber_(patient.unitPrice);
    const patientUnitPrice = normalizeMoneyNumber_(patient.unitPrice);
    const normalizedBurdenRate = normalizeBurdenRateInt_(patient.burdenRate);
    const normalizedMedicalAssistance = normalizeMedicalAssistanceFlag_(patient.medicalAssistance);
    const hasManualUnitPriceInput = patient.manualUnitPrice !== '' && patient.manualUnitPrice !== null
      && patient.manualUnitPrice !== undefined
      && shouldApplyOverrideForEntryType_(manualUnitPriceEntryType, 'self_pay');
    const selfPayUnitPriceSource = hasManualUnitPriceInput ? patient.manualUnitPrice : patient.unitPrice;
    const resolvedSelfPayUnitPrice = normalizeMoneyNumber_(selfPayUnitPriceSource);
    const selfPayChargeAmount = selfPayVisitCount > 0 && resolvedSelfPayUnitPrice > 0
      ? resolvedSelfPayUnitPrice * selfPayVisitCount
      : 0;
    // Item-only self-pay entries (e.g., online_fee) should never create visit counts.
    const itemOnlySelfPayItems = (() => {
      const items = [];
      if (!isMedicalSubsidy || !hasOnlineFee) {
        const baseItems = Array.isArray(patient.selfPayItems) ? patient.selfPayItems.slice() : [];
        baseItems.forEach(item => items.push(item));
      }
      if (billingItems.length) {
        billingItems.forEach(item => {
          items.push({
            type: item.type || item.label || '自費',
            amount: item.amount
          });
        });
      }
      return items;
    })();
    // Visit-based self-pay entries are derived from treatment time categories and carry visitCount.
    const visitBasedSelfPayItems = selfPayChargeAmount
      ? [{ type: '自費', amount: selfPayChargeAmount }]
      : [];
    const combinedSelfPayItems = itemOnlySelfPayItems.concat(visitBasedSelfPayItems);
    const amountCalc = calculateBillingAmounts_({
      visitCount,
      insuranceType: patient.insuranceType,
      burdenRate: normalizedBurdenRate,
      ignoreSelfPayInsuranceType: hasMixedVisitCount,
      manualUnitPrice,
      manualTransportAmount: isMedicalSubsidy && hasOnlineFee ? '' : patient.manualTransportAmount,
      manualSelfPayAmount: isMedicalSubsidy && hasOnlineFee
        ? 0
        : (shouldApplyOverrideForEntryType_(manualSelfPayAmountEntryType, 'self_pay')
          ? patient.manualSelfPayAmount
          : undefined),
      selfPayItems: combinedSelfPayItems,
      unitPrice: patientUnitPrice,
      medicalAssistance: normalizedMedicalAssistance,
      carryOverAmount: carryOverFromPatient + carryOverFromHistory
    });
    const manualBillingInput = shouldApplyOverrideForEntryType_(manualBillingAmountEntryType, 'insurance')
      && Object.prototype.hasOwnProperty.call(patient, 'manualBillingAmount')
      ? patient.manualBillingAmount
      : undefined;
    const hasManualBillingAmount = manualBillingInput !== '' && manualBillingInput !== null && manualBillingInput !== undefined;
    const manualBillingAmount = hasManualBillingAmount ? normalizeMoneyNumber_(manualBillingInput) : '';
    const manualSelfPayInput = shouldApplyOverrideForEntryType_(manualSelfPayAmountEntryType, 'self_pay')
      && Object.prototype.hasOwnProperty.call(patient, 'manualSelfPayAmount')
      ? patient.manualSelfPayAmount
      : undefined;
    const hasManualSelfPayAmount = manualSelfPayInput !== '' && manualSelfPayInput !== null && manualSelfPayInput !== undefined;
    const manualSelfPayAmount = hasManualSelfPayAmount ? normalizeMoneyNumber_(manualSelfPayInput) : '';
    const itemOnlySelfPayTotal = itemOnlySelfPayItems.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    const selfPayItemsTotal = amountCalc.selfPayItems.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    const insuranceEntryTotal = hasManualBillingAmount
      ? manualBillingAmount
      : amountCalc.billingAmount + amountCalc.transportAmount + amountCalc.carryOverAmount;
    const selfPayEntryTotal = hasManualSelfPayAmount ? manualSelfPayAmount : selfPayItemsTotal;
    const entries = [
      Object.assign({}, {
        type: 'insurance',
        entryType: 'insurance',
        unitPrice: amountCalc.unitPrice,
        visitCount: amountCalc.visits,
        treatmentAmount: amountCalc.treatmentAmount,
        transportAmount: amountCalc.transportAmount,
        billingAmount: amountCalc.billingAmount,
        total: insuranceEntryTotal
      }, hasManualBillingAmount ? { manualOverride: { amount: manualBillingAmount } } : {})
    ];
    const hasVisitBasedSelfPay = selfPayVisitCount > 0;
    const hasItemOnlySelfPay = itemOnlySelfPayItems.length > 0;
    const visitBasedEntryTotal = resolvedSelfPayUnitPrice * selfPayVisitCount;
    const itemOnlyEntryTotal = hasManualSelfPayAmount
      ? (hasVisitBasedSelfPay ? 0 : manualSelfPayAmount)
      : itemOnlySelfPayTotal;
    entries.push(Object.assign({}, {
      type: 'self_pay',
      entryType: 'self_pay',
      visitCount: selfPayVisitCount,
      unitPrice: resolvedSelfPayUnitPrice,
      items: visitBasedSelfPayItems,
      selfPayItems: visitBasedSelfPayItems,
      total: visitBasedEntryTotal
    }, (hasManualSelfPayAmount && hasVisitBasedSelfPay) ? { manualOverride: { amount: manualSelfPayAmount } } : {}));
    if (hasItemOnlySelfPay || (!hasVisitBasedSelfPay && (hasManualSelfPayAmount || selfPayEntryTotal))) {
      entries.push(Object.assign({}, {
        type: 'self_pay',
        entryType: 'self_pay',
        items: itemOnlySelfPayItems,
        selfPayItems: itemOnlySelfPayItems,
        total: itemOnlyEntryTotal
      }, (!hasVisitBasedSelfPay && hasManualSelfPayAmount) ? { manualOverride: { amount: manualSelfPayAmount } } : {}));
    }
    const normalizedEntries = entries.map(entryItem => {
      const normalizedType = normalizeBillingEntryType_(
        entryItem && (entryItem.type || entryItem.entryType)
      );

      return Object.assign({}, entryItem, {
        type: normalizedType,
        entryType: entryItem.entryType || normalizedType
      });
    });
    const insuranceEntry = normalizedEntries.find(item => item && item.type === 'insurance') || null;
    const selfPayEntries = normalizedEntries.filter(item => item && item.type === 'self_pay');
    const insuranceTotal = insuranceEntry ? resolveEntryTotalAmount_(insuranceEntry) : 0;
    const selfPayTotal = selfPayEntries.reduce((sum, entryItem) => sum + resolveEntryTotalAmount_(entryItem), 0);
    const resolvedGrandTotal = insuranceTotal + selfPayTotal;
    const resolvedManualBillingAmount = insuranceEntry && insuranceEntry.manualOverride
      ? insuranceEntry.manualOverride.amount
      : '';
    const resolvedManualSelfPayAmount = selfPayEntries.find(entryItem => entryItem && entryItem.manualOverride
      && entryItem.manualOverride.amount !== '' && entryItem.manualOverride.amount !== null
      && entryItem.manualOverride.amount !== undefined);
    const resolvedManualSelfPayAmountValue = resolvedManualSelfPayAmount
      ? resolvedManualSelfPayAmount.manualOverride.amount
      : '';
    const resolvedSelfPayItems = selfPayEntries.reduce((list, entryItem) => {
      const items = Array.isArray(entryItem.items)
        ? entryItem.items
        : (Array.isArray(entryItem.selfPayItems) ? entryItem.selfPayItems : []);
      return list.concat(items);
    }, []);

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
      selfPayVisitCount,
      selfPayCount: selfPayVisitCount,
      manualUnitPrice,
      unitPrice: amountCalc.unitPrice,
      treatmentAmount: amountCalc.treatmentAmount,
      manualTransportAmount: amountCalc.manualTransportAmount,
      transportAmount: amountCalc.transportAmount,
      manualSelfPayAmount: resolvedManualSelfPayAmountValue,
      selfPayItems: resolvedSelfPayItems,
      billingItems,
      carryOverAmount: amountCalc.carryOverAmount,
      carryOverFromHistory,
      billingAmount: insuranceEntry && insuranceEntry.billingAmount != null
        ? insuranceEntry.billingAmount
        : amountCalc.billingAmount,
      total: insuranceTotal,
      grandTotal: resolvedGrandTotal,
      entries: normalizedEntries,
      manualBillingAmount: resolvedManualBillingAmount,
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
