/**
 * 当月の請求書PDFリンクを患者IDと紐付けて返す。
 * @param {Object} [options]
 * @param {Object} [options.patientInfo]
 * @param {Object} [options.nameToId]
 * @param {Object} [options.rootFolder]
 * @param {Date} [options.now]
 * @return {{invoices: Object<string, string|null>, warnings: string[]}}
 */
function loadInvoices(options) {
  const opts = options || {};
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const fetchOptions = Object.assign({}, opts, { now });
  return loadInvoicesUncached_(fetchOptions);
}

function loadInvoicesUncached_(options) {
  const opts = options || {};
  const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo() : null);
  const patients = patientInfo && patientInfo.patients ? patientInfo.patients : {};
  const nameToId = opts.nameToId || (patientInfo && patientInfo.nameToId) || {};
  const warnings = patientInfo && Array.isArray(patientInfo.warnings) ? [].concat(patientInfo.warnings) : [];
  const setupIncomplete = !!(patientInfo && patientInfo.setupIncomplete);
  const invoices = {};
  const invoiceMeta = {};
  const latestMeta = {};
  const logContext = (label, details) => {
    if (typeof dashboardLogContext_ === 'function') {
      dashboardLogContext_(label, details);
    } else if (typeof dashboardWarn_ === 'function') {
      const payload = details ? ` ${details}` : '';
      dashboardWarn_(`[${label}]${payload}`);
    }
  };

  Object.keys(patients).forEach(pid => {
    const normalized = dashboardNormalizePatientId_(pid);
    if (normalized) {
      invoices[normalized] = null;
    }
  });

  const root = opts.rootFolder || dashboardGetInvoiceRootFolder_();
  if (!root || typeof root.getFolders !== 'function') {
    warnings.push('請求書フォルダが取得できませんでした');
    dashboardWarn_('[loadInvoices] invoice root folder not found');
    logContext('loadInvoices:done', `patients=${Object.keys(invoices).length} linked=0 warnings=${warnings.length} setupIncomplete=true`);
    return { invoices, warnings, setupIncomplete: true };
  }

  const tz = dashboardResolveTimeZone_();
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const includePreviousMonth = !!opts.includePreviousMonth;
  const targetMonths = buildInvoiceMonthTargets_(now, tz, includePreviousMonth);
  const currentMonthKey = dashboardFormatDate_(now, tz, 'yyyy-MM');

  const folders = root.getFolders();
  while (folders && typeof folders.hasNext === 'function' && folders.hasNext()) {
    const folder = folders.next();
    const name = folder && typeof folder.getName === 'function' ? folder.getName() : '';
    if (!isTargetInvoiceFolder_(name, targetMonths)) continue;

    const files = folder.getFiles && folder.getFiles();
    while (files && typeof files.hasNext === 'function' && files.hasNext()) {
      const file = files.next();
      const fileName = file && typeof file.getName === 'function' ? file.getName() : '';
      const parsed = parseInvoiceFileName_(fileName, targetMonths);
      if (!parsed) continue;

      const pid = dashboardResolvePatientIdFromName_(parsed.patientName, nameToId);
      if (!pid) {
        warnings.push(`患者名をIDに紐付けできません: ${parsed.patientName}`);
        continue;
      }

      const updated = file && typeof file.getLastUpdated === 'function' ? file.getLastUpdated() : null;
      const updatedDate = dashboardCoerceDate_(updated);
      const updatedTs = updatedDate ? updatedDate.getTime() : 0;
      const url = file && typeof file.getUrl === 'function' ? file.getUrl() : '';

      if (!invoiceMeta[pid]) {
        invoiceMeta[pid] = { months: {} };
      }
      const currentMonthMeta = invoiceMeta[pid].months[parsed.monthKey];
      if (!currentMonthMeta || currentMonthMeta.updatedTs < updatedTs) {
        invoiceMeta[pid].months[parsed.monthKey] = { updatedTs, url };
      }

      if (parsed.monthKey === currentMonthKey) {
        const current = latestMeta[pid];
        if (current && current.updatedTs >= updatedTs) continue;
        latestMeta[pid] = { updatedTs, url };
        invoices[pid] = url || null;
      }
    }
  }

  const linkedCount = Object.keys(invoices).reduce((count, pid) => (invoices[pid] ? count + 1 : count), 0);
  logContext('loadInvoices:done', `patients=${Object.keys(invoices).length} linked=${linkedCount} warnings=${warnings.length} setupIncomplete=${setupIncomplete}`);
  return { invoices, invoiceMeta, warnings, setupIncomplete };
}

function buildInvoiceMonthTargets_(now, tz, includePreviousMonth) {
  const base = dashboardCoerceDate_(now) || new Date();
  const currentMonth = new Date(base.getFullYear(), base.getMonth(), 1);
  const months = [currentMonth];
  if (includePreviousMonth) {
    months.push(new Date(base.getFullYear(), base.getMonth() - 1, 1));
  }
  return months.map(monthDate => {
    const digits = dashboardFormatDate_(monthDate, tz, 'yyyyMM');
    const hyphen = dashboardFormatDate_(monthDate, tz, 'yyyy-MM');
    const kanji = dashboardFormatDate_(monthDate, tz, 'yyyy年MM月');
    return {
      key: hyphen,
      digits,
      hyphen,
      kanji,
      digitPrefix: digits ? digits + '請求書_' : '',
      kanjiPrefix: kanji ? kanji + '請求書_' : ''
    };
  });
}

function isTargetInvoiceFolder_(name, targetMonths) {
  const trimmed = String(name == null ? '' : name).trim();
  if (!trimmed) return false;
  return (targetMonths || []).some(target => {
    return (target.digitPrefix && trimmed.indexOf(target.digitPrefix) === 0)
      || (target.kanjiPrefix && trimmed.indexOf(target.kanjiPrefix) === 0);
  });
}

function parseInvoiceFileName_(fileName, targetMonths) {
  const trimmed = String(fileName == null ? '' : fileName).trim();
  const match = trimmed.match(/^(.+?)_([0-9]{4}-[0-9]{2}|[0-9]{6,8})_請求書\.pdf$/i);
  if (!match) return null;
  const rawMonth = match[2];
  const monthKey = resolveInvoiceMonthKey_(rawMonth, targetMonths);
  if (!monthKey) return null;
  return { patientName: match[1], monthKey };
}

function resolveInvoiceMonthKey_(rawMonth, targetMonths) {
  const normalized = String(rawMonth || '');
  const digits = normalized.replace(/[^0-9]/g, '').slice(0, 6);
  const targets = Array.isArray(targetMonths) ? targetMonths : [];
  for (let i = 0; i < targets.length; i++) {
    const target = targets[i];
    if (target.hyphen && normalized === target.hyphen) return target.key;
    if (target.digits && digits === target.digits) return target.key;
  }
  return '';
}
