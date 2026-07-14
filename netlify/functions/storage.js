const { getStore } = require('@netlify/blobs');

exports.handler = async (event, context) => {
  const headers = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method Not Allowed' }) };
  }

  let body;
  try { body = JSON.parse(event.body); }
  catch { return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid JSON' }) }; }

  const { action, key, value } = body;
  if (!action || !key) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Missing action or key' }) };
  }

  try {
    const store = getStore({ name: 'dashboard-storage', siteID: context.site.id, token: process.env.NETLIFY_BLOBS_TOKEN || process.env.TOKEN });
    
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
    console.error('Storage error:', err);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
