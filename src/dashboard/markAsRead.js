/**
 * 患者カード展開時に最終既読日時を更新する。
 * 旧パス互換のため、必要なユーティリティが未定義の場合は最小限の代替を提供する。
 */

if (typeof dashboardNormalizePatientId_ === 'undefined') {
  function dashboardNormalizePatientId_(value) {
    const raw = value == null ? '' : value;
    return String(raw).trim();
  }
}

if (typeof dashboardNormalizeEmail_ === 'undefined') {
  function dashboardNormalizeEmail_(email) {
    const raw = email == null ? '' : email;
    const normalized = String(raw).trim().toLowerCase();
    return normalized || '';
  }
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

function markAsRead(patientIdOrPayload, email, readAt) {
  const payload = typeof patientIdOrPayload === 'object' && patientIdOrPayload
    ? patientIdOrPayload
    : { patientId: patientIdOrPayload, email, readAt };

  const pid = dashboardNormalizePatientId_(payload.patientId);
  const normalizedEmail = dashboardNormalizeEmail_(payload.email || getActiveUserEmail_());
  const ts = payload.readAt || new Date();

  if (!pid) {
    return { ok: false, patientId: null, readAt: null, reason: 'patientId required' };
  }

  const ok = updateHandoverLastRead(pid, ts, normalizedEmail);
  return { ok: !!ok, patientId: pid, readAt: ok ? dashboardCoerceDate_(ts).toISOString() : null };
}

function getActiveUserEmail_() {
  if (typeof Session !== 'undefined' && Session && typeof Session.getActiveUser === 'function') {
    try {
      const user = Session.getActiveUser();
      if (user && typeof user.getEmail === 'function') {
        const email = user.getEmail();
        if (email) return email;
      }
    } catch (e) { /* ignore */ }
  }
  return '';
}
