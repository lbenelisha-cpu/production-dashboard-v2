const storage = require('./storage.js');
const headers = {'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'};
const reply = (statusCode, body) => ({statusCode,headers,body:JSON.stringify(body)});
exports.handler = async event => {
  if(event.httpMethod === 'GET') return reply(200,{ok:true,service:'daily-report-upload',fields:['name','base64','code']});
  if(event.httpMethod !== 'POST') return reply(405,{error:'METHOD_NOT_ALLOWED'});
  let input; try {input=JSON.parse(event.body||'{}');} catch {return reply(400,{error:'INVALID_JSON'});}
  if(!input || typeof input !== 'object') return reply(400,{error:'INVALID_FILE'});
  const auth=await storage.handler({httpMethod:'POST',body:JSON.stringify({action:'auth',code:input.code})});
  if(auth.statusCode!==200)return auth;
  const name=typeof input.name==='string'?input.name.trim():'';
  const base64=typeof input.base64==='string'?input.base64.replace(/\s/g,''):'';
  if(!/\.(xlsx|xls)$/i.test(name)||name.length>255||!/(20\d{2})\d{4}/.test(name))
    return reply(400,{error:'INVALID_FILE_NAME',message:'Excel filename must include YYYYMMDD'});
  if(!base64 || base64.length>5500000 || base64.length%4!==0 || !/^[A-Za-z0-9+/]+={0,2}$/.test(base64))
    return reply(400,{error:'INVALID_BASE64_OR_FILE_TOO_LARGE'});
  const bytes=Buffer.from(base64,'base64');
  const xlsx=bytes.subarray(0,4).equals(Buffer.from([0x50,0x4b,0x03,0x04]));
  const xls=bytes.subarray(0,8).equals(Buffer.from([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1]));
  if((/\.xlsx$/i.test(name)&&!xlsx)||(/\.xls$/i.test(name)&&!xls))return reply(400,{error:'INVALID_EXCEL_FILE'});
  const token=JSON.parse(auth.body).token;
  const saved=await storage.handler({httpMethod:'POST',body:JSON.stringify({
    action:'set',key:'daily_report_workbook_v1',token,
    value:JSON.stringify({name,base64,uploadedAt:new Date().toISOString()})
  })});
  if(saved.statusCode!==200)return saved;
  return reply(200,{ok:true,verified:true,name,url:'https://adama-pro.netlify.app/daily-report.html'});
};
