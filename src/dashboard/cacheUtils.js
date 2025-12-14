/**
 * ダッシュボード用のシンプルなキャッシュユーティリティ。
 * CacheService が利用可能な場合のみ動作し、利用できない場合は透過的にバイパスする。
 */
const DASHBOARD_CACHE_TTL_SECONDS = 60 * 60 * 12; // 12h

function dashboardGetCache_() {
  if (typeof CacheService === 'undefined' || !CacheService || typeof CacheService.getScriptCache !== 'function') return null;
  try {
    return CacheService.getScriptCache();
  } catch (e) {
    dashboardWarn_('[dashboardGetCache] unavailable: ' + (e && e.message ? e.message : e));
    return null;
  }
}

function dashboardCacheFetch_(key, fetchFn, ttlSeconds) {
  const cache = dashboardGetCache_();
  const ttl = Math.max(5, ttlSeconds || DASHBOARD_CACHE_TTL_SECONDS);
  if (!cache || !key || typeof fetchFn !== 'function') {
    return typeof fetchFn === 'function' ? fetchFn() : null;
  }

  try {
    const hit = cache.get(key);
    if (hit) {
      return JSON.parse(hit);
    }
  } catch (e) {
    dashboardWarn_('[dashboardCacheFetch] failed to read: ' + (e && e.message ? e.message : e));
  }

  const fresh = fetchFn();
  try {
    cache.put(key, JSON.stringify(fresh), ttl);
  } catch (e) {
    dashboardWarn_('[dashboardCacheFetch] failed to write: ' + (e && e.message ? e.message : e));
  }
  return fresh;
}

function dashboardCacheInvalidate_(key) {
  const cache = dashboardGetCache_();
  if (!cache || !key || typeof cache.remove !== 'function') return;
  try {
    cache.remove(key);
  } catch (e) {
    dashboardWarn_('[dashboardCacheInvalidate] failed: ' + (e && e.message ? e.message : e));
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
