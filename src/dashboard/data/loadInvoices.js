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
  const debugFilesFound = { count: 0 };
  const debugExtractedRawIds = [];
  const debugNormalizedIds = [];
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
    logInvoiceDebugSummary_({
      filesFound: debugFilesFound.count,
      extractedRawIds: debugExtractedRawIds,
      normalizedIds: debugNormalizedIds,
      patientMasterIds: Object.keys(invoices),
      linkedInvoices: invoices
    });
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
      debugFilesFound.count += 1;
      const fileName = file && typeof file.getName === 'function' ? file.getName() : '';
      const parsed = parseInvoiceFileName_(fileName, targetMonths);
      if (!parsed) continue;

      const pid = dashboardResolvePatientIdFromName_(parsed.patientName, nameToId);
      if (!pid) {
        warnings.push(`患者名をIDに紐付けできません: ${parsed.patientName}`);
        continue;
      }
      debugExtractedRawIds.push(pid);
      const normalizedPid = dashboardNormalizePatientId_(pid);
      if (normalizedPid) debugNormalizedIds.push(normalizedPid);

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
  logInvoiceDebugSummary_({
    filesFound: debugFilesFound.count,
    extractedRawIds: debugExtractedRawIds,
    normalizedIds: debugNormalizedIds,
    patientMasterIds: Object.keys(invoices),
    linkedInvoices: invoices
  });
  return { invoices, invoiceMeta, warnings, setupIncomplete };
}

function logInvoiceDebugSummary_(payload) {
  if (typeof Logger === 'undefined' || !Logger || typeof Logger.log !== 'function') return;
  const data = payload || {};
  const filesFound = Number(data.filesFound || 0);
  const extractedRawIds = Array.isArray(data.extractedRawIds) ? data.extractedRawIds : [];
  const normalizedIds = uniqueNormalizedIds_(data.normalizedIds);
  const patientMasterIds = uniqueNormalizedIds_(data.patientMasterIds);
  const linkedInvoices = data.linkedInvoices && typeof data.linkedInvoices === 'object' ? data.linkedInvoices : {};
  const matchedIds = normalizedIds.filter(pid => Object.prototype.hasOwnProperty.call(linkedInvoices, pid));
  const invoicesBeforeLink = patientMasterIds.length;
  const invoicesAfterLink = Object.keys(linkedInvoices).filter(pid => !!linkedInvoices[pid]).length;

  Logger.log(`[invoice-debug] filesFound=${filesFound}`);
  Logger.log(`[invoice-debug] extractedRawIds=${JSON.stringify(extractedRawIds.slice(0, 50))}`);
  Logger.log(`[invoice-debug] normalizedIds=${JSON.stringify(normalizedIds.slice(0, 50))}`);
  Logger.log(`[invoice-debug] masterIdsCount=${patientMasterIds.length}`);
  Logger.log(`[invoice-debug] matchedIds=${JSON.stringify(uniqueNormalizedIds_(matchedIds).slice(0, 50))}`);
  Logger.log(`[invoice-debug] invoicesBeforeLink=${invoicesBeforeLink}`);
  Logger.log(`[invoice-debug] invoicesAfterLink=${invoicesAfterLink}`);

  let directCause = '該当なし';
  if (filesFound === 0) {
    directCause = '対象フォルダ内で請求ファイルが検出されていない';
  } else if (!normalizedIds.length) {
    directCause = 'ファイル名から patientId を解決できていない（nameToId 解決失敗）';
  } else if (!matchedIds.length) {
    directCause = '解決した patientId が患者マスタIDに一致していない';
  } else if (invoicesAfterLink === 0) {
    directCause = 'マスタ一致はあるが当月リンクURLが invoices に反映されていない';
  }

  Logger.log('■ linked=0 の直接原因');
  Logger.log(directCause);

  Logger.log('■ 想定される3つの論理パターン');
  Logger.log('- パターン1: フォルダ/ファイル走査段階で対象ファイルが0件（filesFound=0）');
  Logger.log('- パターン2: ファイルはあるが patientId 解決に失敗（normalizedIds=0）');
  Logger.log('- パターン3: patientId 解決後にマスタ不一致またはリンク未反映（matchedIds=0 または invoicesAfterLink=0）');

  Logger.log('■ 修正候補');
  Logger.log('- フォルダ名/ファイル名規約（isTargetInvoiceFolder_ / parseInvoiceFileName_）の実データ整合を確認');
  Logger.log('- dashboardResolvePatientIdFromName_ の nameToId 対応表・患者名正規化ルールを確認');
  Logger.log('- currentMonthKey と請求書月の一致条件、最新URL選定（getLastUpdated）を確認');
}

function uniqueNormalizedIds_(values) {
  const list = Array.isArray(values) ? values : [];
  const seen = Object.create(null);
  const result = [];
  for (let i = 0; i < list.length; i++) {
    const normalized = dashboardNormalizePatientId_(list[i]);
    if (!normalized || seen[normalized]) continue;
    seen[normalized] = true;
    result.push(normalized);
  }
  return result;
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
