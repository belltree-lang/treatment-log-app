/***** Shared billing logger *****/

/**
 * Provide a single logger instance for all billing-related scripts.
 * Falls back to console when Logger is unavailable (e.g., during tests).
 */
const billingLogger_ = (() => {
  try {
    if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
      return { log: (...args) => Logger.log(...args) };
    }
  } catch (err) {
    // ignore logging setup errors and fall back to console
  }

  const fallback = typeof console !== 'undefined' && console && typeof console.log === 'function'
    ? (...args) => console.log(...args)
    : () => {};

  return { log: fallback };
})();

