function doGet(e) {
  if (typeof handleDashboardDoGet_ === 'function') {
    const response = handleDashboardDoGet_(e);
    if (response) return response;
  }
  if (typeof handleBillingDoGet_ === 'function') {
    const response = handleBillingDoGet_(e);
    if (response) return response;
  }
  throw new Error('doGet handler is not configured');
}
