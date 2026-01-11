function doGet(e) {
  if (typeof handleDashboardDoGet_ === 'function') {
    return handleDashboardDoGet_(e);
  }
  if (typeof handleBillingDoGet_ === 'function') {
    return handleBillingDoGet_(e);
  }
  throw new Error('doGet handler is not configured');
}
