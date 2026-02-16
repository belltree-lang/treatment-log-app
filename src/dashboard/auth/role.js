const ADMIN_EMAILS = [
  'belltree@belltree1102.com'
];

function getUserRole_(email) {
  const normalized = dashboardNormalizeEmail_(email);
  if (ADMIN_EMAILS.includes(normalized)) return 'admin';
  return 'staff';
}

function isAdminUser_(email) {
  return getUserRole_(email) === 'admin';
}

if (typeof module !== 'undefined' && module && module.exports) {
  module.exports = {
    getUserRole_,
    isAdminUser_
  };
}
