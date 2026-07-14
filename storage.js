const { getStore } = require('@netlify/blobs');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  let body;
  try { body = JSON.parse(event.body); }
  catch { return { statusCode: 400, body: 'Invalid JSON' }; }

  const { action, key, value } = body;
  if (!action || !key) return { statusCode: 400, body: 'Missing action or key' };

  const store = getStore('dashboard-storage');

  try {
    if (action === 'get') {
      const val = await store.get(key);
      return {
        statusCode: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ value: val ?? null })
      };
    }

    if (action === 'set') {
      await store.set(key, value ?? '');
      return {
        statusCode: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ok: true })
      };
    }

    if (action === 'delete') {
      await store.delete(key);
      return {
        statusCode: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ok: true })
      };
    }

    return { statusCode: 400, body: 'Unknown action' };
  } catch (err) {
    console.error('Storage error:', err);
    return { statusCode: 500, body: 'Storage error: ' + err.message };
  }
};
