const crypto = require('crypto');

const FUNCTION_VERSION = '3.9.0';
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

async function openStore() {
  const { getStore } = await import('@netlify/blobs');
  const siteID = String(process.env.NETLIFY_SITE_ID || process.env.SITE_ID || '').trim();
  const token = String(process.env.NETLIFY_TOKEN || process.env.NETLIFY_ACCESS_TOKEN || '').trim();

  if (siteID && token) {
    return getStore(STORE_NAME, { siteID, token });
  }

  // In a correctly configured Netlify Function, the SDK can use runtime context.
  return getStore(STORE_NAME);
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

  try {
    const store = await openStore();

    if (action === 'status') {
      return response(200, { ok: true, backend: 'netlify-blobs-sdk', store: STORE_NAME, version: FUNCTION_VERSION });
    }

    if (action === 'get') {
      const storedValue = await store.get(key, { type: 'text', consistency: 'strong' });
      return response(200, { value: storedValue, backend: 'netlify-blobs-sdk', version: FUNCTION_VERSION });
    }

    if (action === 'set') {
      await store.set(key, String(value == null ? '' : value));
      const verification = await store.get(key, { type: 'text', consistency: 'strong' });
      if (verification === null) throw new Error('WRITE_VERIFICATION_FAILED');
      return response(200, { ok: true, verified: true, backend: 'netlify-blobs-sdk', version: FUNCTION_VERSION });
    }

    if (action === 'delete') {
      await store.delete(key);
      return response(200, { ok: true, backend: 'netlify-blobs-sdk', version: FUNCTION_VERSION });
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
