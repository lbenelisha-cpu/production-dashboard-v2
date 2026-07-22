const crypto = require('crypto');

const FUNCTION_VERSION = '4.6.3';
const TOKEN_TTL_MS = 8 * 60 * 60 * 1000;
const STORE_NAME = 'production-dashboard';
const MAX_GET_MANY_KEYS = 30;
const READ_CONCURRENCY = 4;

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

  // Prefer Netlify's automatically injected runtime credentials. This avoids
  // stale NETLIFY_TOKEN / SITE_ID environment variables overriding the valid
  // function runtime context and producing gateway errors after a deploy.
  try {
    return getStore(STORE_NAME);
  } catch (runtimeError) {
    const siteID = String(process.env.NETLIFY_SITE_ID || process.env.SITE_ID || '').trim();
    const token = String(process.env.NETLIFY_TOKEN || process.env.NETLIFY_ACCESS_TOKEN || '').trim();
    if (siteID && token) return getStore(STORE_NAME, { siteID, token });
    throw runtimeError;
  }
}

async function mapWithConcurrency(items, limit, worker) {
  const output = new Array(items.length);
  let cursor = 0;
  async function run() {
    while (true) {
      const index = cursor++;
      if (index >= items.length) return;
      output[index] = await worker(items[index], index);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, run));
  return output;
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
  if (!['get', 'getMany', 'set', 'delete', 'status', 'appendEvent'].includes(action)) {
    return response(400, { error: 'UNKNOWN_ACTION', version: FUNCTION_VERSION });
  }
  if (!['status', 'getMany'].includes(action) && (!key || typeof key !== 'string')) {
    return response(400, { error: 'MISSING_KEY', version: FUNCTION_VERSION });
  }
  if (action === 'getMany' && (!Array.isArray(value) || value.length === 0 || value.length > MAX_GET_MANY_KEYS || value.some(k => typeof k !== 'string' || !k))) {
    return response(400, { error: 'INVALID_KEYS', maxKeys: MAX_GET_MANY_KEYS, version: FUNCTION_VERSION });
  }
  if ((action === 'set' || action === 'delete') && !validToken(token)) {
    return response(401, { error: 'ADMIN_AUTHORIZATION_REQUIRED', version: FUNCTION_VERSION });
  }

  // A lightweight health check must not depend on Blobs being reachable.
  if (action === 'status') {
    return response(200, { ok: true, backend: 'netlify-function', store: STORE_NAME, version: FUNCTION_VERSION });
  }

  try {
    const store = await openStore();

    if (action === 'get') {
      const storedValue = await store.get(key, { type: 'text' });
      return response(200, { value: storedValue, backend: 'netlify-blobs-sdk', version: FUNCTION_VERSION });
    }

    if (action === 'getMany') {
      const keys = Array.from(new Set(value));
      const rows = await mapWithConcurrency(keys, READ_CONCURRENCY, async k => {
        const storedValue = await store.get(k, { type: 'text' });
        return [k, storedValue];
      });
      const values = Object.fromEntries(rows);
      return response(200, { values, count: keys.length, backend: 'netlify-blobs-sdk', version: FUNCTION_VERSION });
    }

    if (action === 'set') {
      await store.set(key, String(value == null ? '' : value));
      const verification = await store.get(key, { type: 'text', consistency: 'strong' });
      if (verification === null) throw new Error('WRITE_VERIFICATION_FAILED');
      return response(200, { ok: true, verified: true, backend: 'netlify-blobs-sdk', version: FUNCTION_VERSION });
    }

    if (action === 'appendEvent') {
      const EVENTS_KEY = 'production_events_quality_environment_v1';
      if (key !== EVENTS_KEY) return response(403, { error: 'PUBLIC_WRITE_NOT_ALLOWED', version: FUNCTION_VERSION });
      const input = value && typeof value === 'object' ? value : {};
      const allowedTypes = new Set(['quality', 'safety', 'environment']);
      const cleanText = (v, max) => String(v == null ? '' : v).trim().slice(0, max);
      const date = cleanText(input.date, 10);
      const description = cleanText(input.description, 3000);
      const type = cleanText(input.type, 20);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !description || !allowedTypes.has(type)) {
        return response(400, { error: 'INVALID_EVENT_DATA', version: FUNCTION_VERSION });
      }
      const savedEvent = {
        id: cleanText(input.id, 80) || `e_${Date.now()}_${crypto.randomBytes(4).toString('hex')}`,
        date,
        type,
        facility: cleanText(input.facility, 40),
        area: cleanText(input.area, 200),
        workOrder: cleanText(input.workOrder, 100),
        batch: cleanText(input.batch, 100),
        severity: cleanText(input.severity, 50),
        description,
        action: cleanText(input.action, 3000),
        source: 'manual',
        createdAt: Number.isFinite(Number(input.createdAt)) ? Number(input.createdAt) : Date.now()
      };
      const existingText = await store.get(key, { type: 'text', consistency: 'strong' });
      let rows = [];
      try { rows = existingText ? JSON.parse(existingText) : []; } catch { rows = []; }
      if (!Array.isArray(rows)) rows = [];
      if (!rows.some(row => row && row.id === savedEvent.id)) rows.push(savedEvent);
      await store.set(key, JSON.stringify(rows));
      return response(200, { ok: true, event: savedEvent, count: rows.length, backend: 'netlify-blobs-sdk', version: FUNCTION_VERSION });
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
