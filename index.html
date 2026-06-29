// Netlify Function — אחסון שיתופי לדשבורד הייצור והאריזה.
// כל המשתמשים שמגיעים לאתר קוראים/כותבים לאותו "Blob Store" בצד השרת,
// כך שהנתונים (רשומות, יעדים, איכות, אירועים) משותפים לכולם, לא נשמרים בדפדפן בלבד.
//
// ה-API: בקשת POST יחידה עם גוף JSON: { action: 'get'|'set'|'delete'|'list', key, value, prefix }
// אין כאן הגנת סיסמה — כל מי שמגיע לכתובת האתר יכול לקרוא ולכתוב נתונים (לפי בקשת המשתמש).

const { getStore } = require('@netlify/blobs');

const STORE_NAME = 'production-dashboard';

exports.handler = async (event) => {
  const headers = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  let body;
  try {
    body = JSON.parse(event.body || '{}');
  } catch (e) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid JSON body' }) };
  }

  const { action, key, value, prefix } = body;
  if (!action || (action !== 'list' && !key)) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Missing action or key' }) };
  }

  let store;
  try {
    store = getStore(STORE_NAME);
  } catch (e) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: 'Blob store unavailable: ' + (e && e.message) }) };
  }

  try {
    if (action === 'get') {
      const v = await store.get(key, { type: 'text' });
      return { statusCode: 200, headers, body: JSON.stringify({ key, value: (v === null || v === undefined) ? null : v }) };
    }

    if (action === 'set') {
      await store.set(key, String(value == null ? '' : value));
      return { statusCode: 200, headers, body: JSON.stringify({ key, ok: true }) };
    }

    if (action === 'delete') {
      await store.delete(key);
      return { statusCode: 200, headers, body: JSON.stringify({ key, deleted: true }) };
    }

    if (action === 'list') {
      const { blobs } = await store.list({ prefix: prefix || '' });
      return { statusCode: 200, headers, body: JSON.stringify({ keys: (blobs || []).map((b) => b.key) }) };
    }

    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Unknown action: ' + action }) };
  } catch (e) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: String((e && e.message) || e) }) };
  }
};
