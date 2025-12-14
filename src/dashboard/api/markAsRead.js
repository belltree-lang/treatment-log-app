/**
 * 患者カード展開時に最終既読日時を更新する。
 * @param {string|Object} patientIdOrPayload - 患者ID、または { patientId, email, readAt } を含むオブジェクト。
 * @param {string} [email]
 * @param {Date|string} [readAt]
 * @return {{ ok: boolean, patientId: string|null, readAt: string|null, reason?: string }}
 */
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
