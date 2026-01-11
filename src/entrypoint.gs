/**
 * Ensure the web app entrypoint doGet is always available in the global scope.
 *
 * Some deployments may omit files that declare doGet directly; this shim restores
 * the global function and delegates to whichever handler is present.
 */
(function ensureGlobalDoGet() {
  if (typeof globalThis === 'undefined') return;
  if (typeof globalThis.doGet === 'function') return;

  globalThis.doGet = function(e) {
    if (typeof handleDashboardDoGet_ === 'function') {
      const response = handleDashboardDoGet_(e);
      if (response != null) return response;
    }

    if (typeof handleBillingDoGet_ === 'function') {
      return handleBillingDoGet_(e);
    }

    throw new Error('doGet handler is not configured');
  };
})();
