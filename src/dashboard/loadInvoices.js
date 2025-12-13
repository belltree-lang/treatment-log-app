/**
 * 当月の請求書PDFリンクを患者IDと紐付けて返す。
 * @param {Object} [options]
 * @param {Object} [options.patientInfo] - loadPatientInfo() の戻り値を差し込む場合に利用。
 * @param {Object} [options.nameToId] - 氏名から患者IDへのマップを直接指定する場合に利用。
 * @param {Object} [options.rootFolder] - 請求書フォルダのルートを直接指定する場合に利用。
 * @param {Date} [options.now] - テスト用に現在日時を差し替える。
 * @return {{invoices: Object<string, string|null>, warnings: string[]}}
 */
function loadInvoices(options) {
  const opts = options || {};
  const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo() : null);
  const patients = patientInfo && patientInfo.patients ? patientInfo.patients : {};
  const nameToId = opts.nameToId || (patientInfo && patientInfo.nameToId) || {};
  const warnings = patientInfo && Array.isArray(patientInfo.warnings) ? [].concat(patientInfo.warnings) : [];
  const invoices = {};
  const latestMeta = {};

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
    return { invoices, warnings };
  }

  const tz = dashboardResolveTimeZone_();
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const targetYmDigits = dashboardFormatDate_(now, tz, 'yyyyMM');
  const targetYmHyphen = dashboardFormatDate_(now, tz, 'yyyy-MM');
  const targetYmKanji = dashboardFormatDate_(now, tz, 'yyyy年MM月');

  const folders = root.getFolders();
  while (folders && typeof folders.hasNext === 'function' && folders.hasNext()) {
    const folder = folders.next();
    const name = folder && typeof folder.getName === 'function' ? folder.getName() : '';
    if (!isTargetInvoiceFolder_(name, targetYmDigits, targetYmKanji)) continue;

    const files = folder.getFiles && folder.getFiles();
    while (files && typeof files.hasNext === 'function' && files.hasNext()) {
      const file = files.next();
      const fileName = file && typeof file.getName === 'function' ? file.getName() : '';
      const parsed = parseInvoiceFileName_(fileName, targetYmDigits, targetYmHyphen);
      if (!parsed) continue;

      const pid = dashboardResolvePatientIdFromName_(parsed.patientName, nameToId);
      if (!pid) {
        warnings.push(`患者名をIDに紐付けできません: ${parsed.patientName}`);
        continue;
      }

      const updated = file && typeof file.getLastUpdated === 'function' ? file.getLastUpdated() : null;
      const updatedDate = dashboardCoerceDate_(updated);
      const updatedTs = updatedDate ? updatedDate.getTime() : 0;
      const current = latestMeta[pid];
      if (current && current.updatedTs >= updatedTs) continue;

      const url = file && typeof file.getUrl === 'function' ? file.getUrl() : '';
      latestMeta[pid] = { updatedTs, url };
      invoices[pid] = url || null;
    }
  }

  return { invoices, warnings };
}

function isTargetInvoiceFolder_(name, targetYmDigits, targetYmKanji) {
  const trimmed = String(name == null ? '' : name).trim();
  if (!trimmed) return false;
  const digitPrefix = targetYmDigits ? targetYmDigits + '請求書_' : '';
  const kanjiPrefix = targetYmKanji ? targetYmKanji + '請求書_' : '';
  return (digitPrefix && trimmed.indexOf(digitPrefix) === 0)
    || (kanjiPrefix && trimmed.indexOf(kanjiPrefix) === 0);
}

function parseInvoiceFileName_(fileName, targetYmDigits, targetYmHyphen) {
  const trimmed = String(fileName == null ? '' : fileName).trim();
  const match = trimmed.match(/^(.+?)_([0-9]{4}-[0-9]{2}|[0-9]{6,8})_請求書\.pdf$/i);
  if (!match) return null;
  const rawMonth = match[2];
  const matchesDigits = targetYmDigits && rawMonth.replace(/[^0-9]/g, '').slice(0, 6) === targetYmDigits;
  const matchesHyphen = targetYmHyphen && rawMonth === targetYmHyphen;
  if (!matchesDigits && !matchesHyphen) return null;
  return { patientName: match[1] };
}

if (typeof dashboardCoerceDate_ === 'undefined') {
  function dashboardCoerceDate_(value) {
    if (value instanceof Date) return value;
    if (value && typeof value.getTime === 'function') {
      const ts = value.getTime();
      if (Number.isFinite(ts)) return new Date(ts);
    }
    if (value === null || value === undefined) return null;
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
}

if (typeof dashboardGetInvoiceRootFolder_ === 'undefined') {
  function dashboardGetInvoiceRootFolder_() {
    const folderId = typeof INVOICE_PARENT_FOLDER_ID !== 'undefined' ? INVOICE_PARENT_FOLDER_ID : '';
    if (folderId && typeof DriveApp !== 'undefined' && DriveApp && typeof DriveApp.getFolderById === 'function') {
      try { return DriveApp.getFolderById(folderId); } catch (e) { /* ignore */ }
    }
    return null;
  }
}

if (typeof dashboardResolvePatientIdFromName_ === 'undefined') {
  function dashboardResolvePatientIdFromName_(name, nameToId) {
    const key = dashboardNormalizeNameKey_(name);
    return key && nameToId ? nameToId[key] : '';
  }
}

if (typeof dashboardNormalizeNameKey_ === 'undefined') {
  function dashboardNormalizeNameKey_(name) {
    return String(name == null ? '' : name)
      .replace(/\s+/g, '')
      .toLowerCase();
  }
}

if (typeof dashboardNormalizePatientId_ === 'undefined') {
  function dashboardNormalizePatientId_(value) {
    const raw = value == null ? '' : value;
    return String(raw).trim();
  }
}

if (typeof dashboardResolveTimeZone_ === 'undefined') {
  function dashboardResolveTimeZone_() {
    if (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function') {
      const tz = Session.getScriptTimeZone();
      if (tz) return tz;
    }
    return 'Asia/Tokyo';
  }
}

if (typeof dashboardFormatDate_ === 'undefined') {
  function dashboardFormatDate_(date, tz, format) {
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
      try { return Utilities.formatDate(date, tz, format); } catch (e) { /* ignore */ }
    }
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
    const iso = date.toISOString();
    if (format === 'yyyy年MM月') {
      const y = iso.slice(0, 4);
      const m = iso.slice(5, 7);
      return `${y}年${m}月`;
    }
    if (format && format.indexOf('HH') >= 0) {
      return iso.replace('T', ' ').slice(0, 16);
    }
    if (format && format.indexOf('-') >= 0) {
      return iso.slice(0, 10);
    }
    return iso.replace(/-/g, '').slice(0, 6);
  }
}

if (typeof dashboardWarn_ === 'undefined') {
  function dashboardWarn_(message) {
    if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
      try { Logger.log(message); return; } catch (e) { /* ignore */ }
    }
    if (typeof console !== 'undefined' && console && typeof console.warn === 'function') {
      console.warn(message);
    }
  }
}
