// =============================================================
// Config.gs - Centralized configuration & helpers
// =============================================================

/** Header rows count used across sheets */
var CONFIG_HEADER_ROWS = 3;

/** Property accessor with fallback */
function cfgProp(key, fallback) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(key);
    return (v !== null && v !== '') ? v : fallback;
  } catch (_e) { return fallback; }
}

// Core properties (add more as needed)
var CFG = {
  parentFolderId: cfgProp('PARENT_FOLDER_ID', ''),
  templateFileId: cfgProp('TEMPLATE_FILE_ID', ''),
  n8nWebhookUrl: cfgProp('N8N_WEBHOOK_URL', ''),
  centralUrl: cfgProp('CENTRAL_URL', ''),
  geocodeKey: cfgProp('GEOCODE_API_KEY', cfgProp('API_KEY', '')), // support legacy key name
  geocodeLastRowKey: 'GEOCODE_LAST_ROW_PROCESSED_CURRENT_COMPS',
  geocodeBatchSize: parseInt(cfgProp('GEOCODE_BATCH_SIZE', '250'), 10) || 250
};

/** Simple log helper with module tag */
function log_(mod, msg) { Logger.log('[' + mod + '] ' + msg); }

/** Percent/ratio parser (shared) */
function parsePercentOrRatio_(raw) {
  if (raw == null || raw === '') return null;
  if (typeof raw === 'number') return raw > 1 ? raw / 100 : raw;
  var s = String(raw).trim();
  if (!s) return null;
  var pct = s.indexOf('%') !== -1;
  var num = parseFloat(s.replace(/[^0-9.+-]/g, ''));
  if (isNaN(num)) return null;
  return (pct || num > 1) ? num / 100 : num;
}

/** Safe folder/file name sanitizer */
function sanitizeName__(name) {
  return String(name).replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim().substring(0,150);
}
