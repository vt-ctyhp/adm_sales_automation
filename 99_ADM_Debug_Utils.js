/*** 99_ADM_Debug_Utils.gs — Centralized ADM_DEBUG (shared) ***/

// Internal 60s memory cache (per execution)
var __ADM_DBG_CACHE = { val: null, ts: 0 };

function ADM_isDebug(forceReload) {
  try {
    var now = Date.now();
    if (!forceReload && __ADM_DBG_CACHE.val !== null && (now - __ADM_DBG_CACHE.ts) < 60000) {
      return __ADM_DBG_CACHE.val;
    }
    var sp = PropertiesService.getScriptProperties();
    var raw = String(sp.getProperty('ADM_DEBUG') || 'false').trim().toLowerCase();
    var on = (raw === 'true' || raw === '1' || raw === 'yes' || raw === 'on');
    __ADM_DBG_CACHE = { val: on, ts: now };
    return on;
  } catch (e) {
    return false;
  }
}

function ADM_dbg(tag, data) {
  if (!ADM_isDebug()) return;
  try {
    Logger.log('[ADM_DEBUG] ' + String(tag) + (arguments.length > 1 ? ' :: ' + JSON.stringify(data) : ''));
  } catch (e) {
    Logger.log('[ADM_DEBUG] ' + String(tag));
  }
}

function ADM_dbj(tag, obj) { ADM_dbg(tag, obj); }

function ADM_time(label) {
  var t0 = Date.now();
  return function done(extra) {
    ADM_dbg('⏱ ' + (label || 'timer') + ' ' + (Date.now() - t0) + 'ms', extra);
  };
}

// Export a tiny namespace if you prefer that style elsewhere.
this.ADM_DEBUG_UTIL = { isDebug: ADM_isDebug, log: ADM_dbg, time: ADM_time };

// --- Back‑compat aliases (only define if not already present) ---
if (typeof dbg !== 'function') this.dbg = ADM_dbg;
if (typeof DBG !== 'function') this.DBG = function() {
  try { ADM_dbg([].slice.call(arguments).join(' ')); }
  catch (_) { ADM_dbg(String(arguments[0] || ''), arguments[1]); }
};
if (typeof DBJ !== 'function') this.DBJ = ADM_dbj;
if (typeof T  !== 'function') this.T  = ADM_time;

// Optional: expose a boolean for templates that used `ADM_DEBUG` directly
if (typeof this.ADM_DEBUG === 'undefined') this.ADM_DEBUG = ADM_isDebug();
