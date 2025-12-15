/**
 * 前月末時点で最後に施術した施術者を担当者として判定する。
 * @param {Object} [options]
 * @param {Object} [options.patientInfo] - loadPatientInfo() の戻り値を差し替える際に利用。
 * @param {Object} [options.treatmentLogs] - loadTreatmentLogs() の戻り値を差し替える際に利用。
 * @return {{responsible: Object<string, string|null>, warnings: string[]}}
 */
function assignResponsibleStaff(options) {
  const opts = options || {};
  const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo() : null);
  const patients = patientInfo && patientInfo.patients ? patientInfo.patients : {};
  const warnings = patientInfo && Array.isArray(patientInfo.warnings) ? [].concat(patientInfo.warnings) : [];
  const setupIncomplete = !!(patientInfo && patientInfo.setupIncomplete);

  const treatment = opts.treatmentLogs || (typeof loadTreatmentLogs === 'function' ? loadTreatmentLogs({ patientInfo }) : null);
  const lastStaffByPatient = treatment && treatment.lastStaffByPatient ? treatment.lastStaffByPatient : {};
  if (treatment && Array.isArray(treatment.warnings)) {
    warnings.push.apply(warnings, treatment.warnings);
  }

  const responsible = {};
  Object.keys(patients).forEach(pid => {
    const normalized = dashboardNormalizePatientId_(pid);
    if (!normalized) return;
    responsible[normalized] = lastStaffByPatient && lastStaffByPatient[normalized] ? lastStaffByPatient[normalized] : null;
  });

  // 施術録側に存在するが患者情報に未登録の患者も含める
  Object.keys(lastStaffByPatient || {}).forEach(pid => {
    const normalized = dashboardNormalizePatientId_(pid);
    if (!normalized || Object.prototype.hasOwnProperty.call(responsible, normalized)) return;
    responsible[normalized] = lastStaffByPatient[pid] || null;
  });

  return { responsible, warnings, setupIncomplete: setupIncomplete || !!(treatment && treatment.setupIncomplete) };
}
