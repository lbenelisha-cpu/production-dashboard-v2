const https = require('https');
const crypto = require('crypto');

const FUNCTION_VERSION = '3.1.0';
const TOKEN_TTL_MS = 8 * 60 * 60 * 1000;
function secret(){ return process.env.ADMIN_TOKEN_SECRET || process.env.ADMIN_CODE || ''; }
function sign(payload){ return crypto.createHmac('sha256', secret()).update(payload).digest('hex'); }
function makeToken(){
  const payload = Buffer.from(JSON.stringify({role:'admin', exp:Date.now()+TOKEN_TTL_MS})).toString('base64url');
  return payload + '.' + sign(payload);
}
function validToken(token){
  try{
    if(!token || !secret()) return false;
    const [payload,sig]=String(token).split('.');
    const expected=sign(payload);
    if(!sig || sig.length!==expected.length || !crypto.timingSafeEqual(Buffer.from(sig),Buffer.from(expected))) return false;
    const data=JSON.parse(Buffer.from(payload,'base64url').toString('utf8'));
    return data.role==='admin' && Number(data.exp)>Date.now();
  }catch(e){ return false; }
}

exports.handler = async (event) => {
  const headers = {'Content-Type':'application/json','Access-Control-Allow-Origin':'*'};
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: '{}' };
  let body;
  try { body = JSON.parse(event.body || '{}'); }
  catch { return { statusCode: 400, headers, body: JSON.stringify({ error: 'Invalid JSON' }) }; }

  if(body.action === 'auth'){
    const configured=String(process.env.ADMIN_CODE || '').trim();
    const supplied=String(body.code || '').trim();
    const ok=configured && supplied.length===configured.length && crypto.timingSafeEqual(Buffer.from(supplied),Buffer.from(configured));
    if (!configured) {
      return {statusCode:500,headers,body:JSON.stringify({error:'ADMIN_CODE_NOT_CONFIGURED',version:FUNCTION_VERSION})};
    }
    return ok
      ? {statusCode:200,headers,body:JSON.stringify({token:makeToken(),expiresIn:TOKEN_TTL_MS,version:FUNCTION_VERSION})}
      : {statusCode:401,headers,body:JSON.stringify({error:'INVALID_ADMIN_CODE',version:FUNCTION_VERSION})};
  }

  const { action, key, value, token } = body;
  if (!action || !key) return { statusCode: 400, headers, body: JSON.stringify({ error: 'Missing params' }) };
  if ((action === 'set' || action === 'delete') && !validToken(token))
    return { statusCode: 401, headers, body: JSON.stringify({ error: 'Admin authorization required' }) };

  const siteId = process.env.SITE_ID || process.env.NETLIFY_SITE_ID;
  const accessToken = process.env.NETLIFY_TOKEN || process.env.NETLIFY_ACCESS_TOKEN;
  if(!siteId || !accessToken) return {statusCode:500,headers,body:JSON.stringify({error:'Storage environment variables missing'})};
  const store = 'production-dashboard';
  const encodedKey = encodeURIComponent(key);
  const baseUrl = `https://api.netlify.com/api/v1/sites/${siteId}/blobs/${store}/${encodedKey}`;
  const request = (method, data) => new Promise((resolve, reject) => {
    const url = new URL(baseUrl);
    const options = {hostname:url.hostname,path:url.pathname,method,headers:{'Authorization':`Bearer ${accessToken}`,'Content-Type':'application/octet-stream'}};
    const req = https.request(options, res => { let buf=''; res.on('data',d=>buf+=d); res.on('end',()=>resolve({status:res.statusCode,body:buf})); });
    req.on('error', reject); if(data !== undefined && data !== null) req.write(String(data)); req.end();
  });
  try {
    if (action === 'get') { const res=await request('GET'); return {statusCode:200,headers,body:JSON.stringify({value:res.status===200?res.body:null})}; }
    if (action === 'set') { const res=await request('PUT', value ?? ''); if(res.status<200||res.status>=300) throw new Error('Blob write failed '+res.status); return {statusCode:200,headers,body:JSON.stringify({ok:true})}; }
    if (action === 'delete') { const res=await request('DELETE'); if(res.status<200||res.status>=300) throw new Error('Blob delete failed '+res.status); return {statusCode:200,headers,body:JSON.stringify({ok:true})}; }
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Unknown action' }) };
  } catch (err) { return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) }; }
};
