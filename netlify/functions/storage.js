const https = require('https');

exports.handler = async (event) => {
  const headers = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
  };

  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: '{}' };

  let body;
  try { body = JSON.parse(event.body); }
  catch { return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid JSON' }) }; }

  const { action, key, value } = body;
  if (!action || !key) return { statusCode: 400, headers, body: JSON.stringify({ error: 'Missing params' }) };

  const siteId = process.env.SITE_ID || process.env.NETLIFY_SITE_ID;
  const token = process.env.NETLIFY_TOKEN || process.env.NETLIFY_ACCESS_TOKEN;
  const store = 'production-dashboard';
  const encodedKey = encodeURIComponent(key);
  const baseUrl = `https://api.netlify.com/api/v1/sites/${siteId}/blobs/${store}/${encodedKey}`;

  const request = (method, data) => new Promise((resolve, reject) => {
    const url = new URL(baseUrl);
    const options = {
      hostname: url.hostname,
      path: url.pathname,
      method,
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/octet-stream' }
    };
    const req = https.request(options, res => {
      let buf = '';
      res.on('data', d => buf += d);
      res.on('end', () => resolve({ status: res.statusCode, body: buf }));
    });
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });

  try {
    if (action === 'get') {
      const res = await request('GET');
      const val = res.status === 200 ? res.body : null;
      return { statusCode: 200, headers, body: JSON.stringify({ value: val }) };
    }
    if (action === 'set') {
      await request('PUT', value ?? '');
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true }) };
    }
    if (action === 'delete') {
      await request('DELETE');
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true }) };
    }
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Unknown action' }) };
  } catch (err) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
