const crypto = require('crypto');
const { getStore } = require('@netlify/blobs');

const FUNCTION_VERSION = '3.3.0';
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
    // In Netlify Functions, project credentials are provided automatically.
    const store = getStore({ name: STORE_NAME, consistency: 'strong' });

    if (action === 'status') {
      return response(200, { ok: true, backend: 'netlify-blobs', store: STORE_NAME, version: FUNCTION_VERSION });
    }
    if (action === 'get') {
      const storedValue = await store.get(key, { type: 'text', consistency: 'strong' });
      return response(200, { value: storedValue, backend: 'netlify-blobs', version: FUNCTION_VERSION });
    }
    if (action === 'set') {
      await store.set(key, value == null ? '' : String(value));
      return response(200, { ok: true, backend: 'netlify-blobs', version: FUNCTION_VERSION });
    }
    if (action === 'delete') {
      await store.delete(key);
      return response(200, { ok: true, backend: 'netlify-blobs', version: FUNCTION_VERSION });
    }
  } catch (error) {
    console.error('Netlify Blobs error:', error);
    return response(500, {
      error: 'BLOB_STORAGE_ERROR',
      message: error && error.message ? error.message : 'Unknown storage error',
      version: FUNCTION_VERSION
    });
  }
};
