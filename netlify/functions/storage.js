const crypto = require('crypto');

const FUNCTION_VERSION = '3.7.0';
const TOKEN_TTL_MS = 8 * 60 * 60 * 1000;
const STORE_NAME = 'production-dashboard';

function secret() {
  return process.env.ADMIN_TOKEN_SECRET || process.env.ADMIN_CODE || '';
}

function sign(payload) {
  return crypto.createHmac('sha256', secret()).update(payload).digest('hex');
}

function makeToken() {
  const payload = Buffer.from(JSON.stringify({ role: 'admin', exp: Date.now() + TOKEN_TTL_MS })).toString('base64url');
  return `${payload}.${sign(payload)}`;
}

function validToken(token) {
  try {
    if (!token || !secret()) return false;
    const [payload, sig] = String(token).split('.');
    if (!payload || !sig) return false;
    const expected = sign(payload);
    const a = Buffer.from(sig);
    const b = Buffer.from(expected);
    if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return false;
    const data = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
    return data.role === 'admin' && Number(data.exp) > Date.now();
  } catch {
    return false;
  }
}

function response(statusCode, payload) {
  return {
    statusCode,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Cache-Control': 'no-store'
    },
    body: payload === null ? '' : JSON.stringify(payload)
  };
}

function extractSignedUrl(raw) {
  const text = String(raw || '').trim();
  if (!text) return null;
  if (/^https:\/\//i.test(text)) return text;
  try {
    const parsed = JSON.parse(text);
    if (typeof parsed === 'string' && /^https:\/\//i.test(parsed)) return parsed;
    for (const key of ['url', 'download_url', 'downloadUrl', 'signed_url', 'signedUrl']) {
      if (parsed && typeof parsed[key] === 'string' && /^https:\/\//i.test(parsed[key])) return parsed[key];
    }
  } catch {}
  return null;
}

async function blobRequest({ siteID, token, key, method, value }) {
  const encodedStore = encodeURIComponent(STORE_NAME);
  const encodedKey = encodeURIComponent(key);
  const url = `https://api.netlify.com/api/v1/sites/${encodeURIComponent(siteID)}/blobs/${encodedStore}/${encodedKey}`;
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(method === 'PUT' ? { 'Content-Type': 'application/octet-stream' } : {})
    },
    body: method === 'PUT' ? String(value == null ? '' : value) : undefined,
    cache: 'no-store'
  });
  const text = await res.text();
  return { ok: res.ok, status: res.status, text };
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return response(204, null);
  if (event.httpMethod !== 'POST') return response(405, { error: 'METHOD_NOT_ALLOWED', version: FUNCTION_VERSION });

  let body;
  try {
    body = JSON.parse(event.body || '{}');
  } catch {
    return response(400, { error: 'INVALID_JSON', version: FUNCTION_VERSION });
  }

  if (body.action === 'auth') {
    const configured = String(process.env.ADMIN_CODE || '').trim();
    const supplied = String(body.code || '').trim();
    if (!configured) return response(500, { error: 'ADMIN_CODE_NOT_CONFIGURED', version: FUNCTION_VERSION });
    const a = Buffer.from(supplied);
    const b = Buffer.from(configured);
    const ok = a.length === b.length && crypto.timingSafeEqual(a, b);
    return ok
      ? response(200, { token: makeToken(), expiresIn: TOKEN_TTL_MS, version: FUNCTION_VERSION })
      : response(401, { error: 'INVALID_ADMIN_CODE', version: FUNCTION_VERSION });
  }

  const { action, key, value, token } = body;
  if (!['get', 'set', 'delete', 'status'].includes(action)) {
    return response(400, { error: 'UNKNOWN_ACTION', version: FUNCTION_VERSION });
  }
  if (action !== 'status' && (!key || typeof key !== 'string')) {
    return response(400, { error: 'MISSING_KEY', version: FUNCTION_VERSION });
  }
  if ((action === 'set' || action === 'delete') && !validToken(token)) {
    return response(401, { error: 'ADMIN_AUTHORIZATION_REQUIRED', version: FUNCTION_VERSION });
  }

  const siteID = String(process.env.NETLIFY_SITE_ID || process.env.SITE_ID || '').trim();
  const netlifyToken = String(process.env.NETLIFY_TOKEN || process.env.NETLIFY_ACCESS_TOKEN || '').trim();
  if (!siteID || !netlifyToken) {
    return response(500, {
      error: 'BLOB_CREDENTIALS_NOT_CONFIGURED',
      message: 'NETLIFY_SITE_ID and NETLIFY_TOKEN must be configured for shared storage.',
      version: FUNCTION_VERSION
    });
  }

  try {
    if (action === 'status') {
      return response(200, { ok: true, backend: 'netlify-blobs-rest', store: STORE_NAME, version: FUNCTION_VERSION });
    }

    if (action === 'get') {
      const result = await blobRequest({ siteID, token: netlifyToken, key, method: 'GET' });
      if (result.status === 404) return response(200, { value: null, backend: 'netlify-blobs-rest', version: FUNCTION_VERSION });
      if (!result.ok) throw new Error(`BLOB_GET_FAILED_${result.status}: ${result.text.slice(0, 300)}`);

      const signedUrl = extractSignedUrl(result.text);
      if (signedUrl) {
        const download = await fetch(signedUrl, { cache: 'no-store' });
        const downloadedText = await download.text();

        // Netlify can return a signed download URL even when the object does not
        // exist yet. Amazon S3 then answers with 404 / NoSuchKey. This is a normal
        // first-run state and must be treated as an empty value, not as an error.
        if (download.status === 404 || /<Code>\s*NoSuchKey\s*<\/Code>/i.test(downloadedText)) {
          return response(200, { value: null, backend: 'netlify-blobs-rest', version: FUNCTION_VERSION });
        }

        if (!download.ok) throw new Error(`BLOB_DOWNLOAD_FAILED_${download.status}: ${downloadedText.slice(0, 300)}`);
        return response(200, { value: downloadedText, backend: 'netlify-blobs-rest', version: FUNCTION_VERSION });
      }

      return response(200, { value: result.text, backend: 'netlify-blobs-rest', version: FUNCTION_VERSION });
    }

    if (action === 'set') {
      const result = await blobRequest({ siteID, token: netlifyToken, key, method: 'PUT', value });
      if (!result.ok) throw new Error(`BLOB_SET_FAILED_${result.status}: ${result.text.slice(0, 300)}`);
      return response(200, { ok: true, backend: 'netlify-blobs-rest', version: FUNCTION_VERSION });
    }

    if (action === 'delete') {
      const result = await blobRequest({ siteID, token: netlifyToken, key, method: 'DELETE' });
      if (!result.ok && result.status !== 404) throw new Error(`BLOB_DELETE_FAILED_${result.status}: ${result.text.slice(0, 300)}`);
      return response(200, { ok: true, backend: 'netlify-blobs-rest', version: FUNCTION_VERSION });
    }
  } catch (error) {
    console.error('Shared storage error:', error);
    return response(500, {
      error: 'BLOB_STORAGE_ERROR',
      message: error && error.message ? error.message : 'Unknown storage error',
      version: FUNCTION_VERSION
    });
  }
};
