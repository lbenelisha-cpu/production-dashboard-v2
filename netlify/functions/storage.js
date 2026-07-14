const { getStore } = require('@netlify/blobs');

exports.handler = async (event) => {
  const headers = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers, body: '{}' };
  }

  let body;
  try { body = JSON.parse(event.body); }
  catch { return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid JSON' }) }; }

  const { action, key, value } = body;
  if (!action || !key) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Missing params' }) };
  }

  try {
    const store = getStore('dashboard-storage');

    if (action === 'get') {
      const val = await store.get(key);
      return { statusCode: 200, headers, body: JSON.stringify({ value: val ?? null }) };
    }
    if (action === 'set') {
      await store.set(key, value ?? '');
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true }) };
    }
    if (action === 'delete') {
      await store.delete(key);
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true }) };
    }
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Unknown action' }) };
  } catch (err) {
    console.error('Storage error:', err.message);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
