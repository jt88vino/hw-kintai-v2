// Same-origin transport: do not log or cache request bodies, credentials or attendance.
const { randomUUID } = require('node:crypto');
const GAS_URL = 'https://script.google.com/macros/s/AKfycbxpTfsziHbGsM6kOXuPvlVQdnnufZ2RS_reX5NhoE0aOpra-RD5m8RBJXwEwc_34a-C/exec';
const WRITE_ACTIONS = new Set(['add']); // Admin credentials continue to go directly to Google.
module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'private, no-store');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  if (!['GET', 'POST'].includes(req.method)) return res.status(405).json({ok:false,error:'method_not_allowed'});
  try {
    let params;
    if (req.method === 'GET') {
      const q = req.query || {};
      if (q.action && q.action !== 'read') return res.status(400).json({ok:false,error:'invalid_action'});
      const scope = q.scope || 'recent';
      if (!['recent','month'].includes(scope) || (scope === 'month' && !/^\d{4}-(0[1-9]|1[0-2])$/.test(q.month || ''))) return res.status(400).json({ok:false,error:'invalid_scope'});
      params = {action:'read',scope};
      if (scope === 'month') params.month = q.month;
      if (q.fresh === '1') params.fresh = '1';
    } else {
      if (req.headers.origin && req.headers.origin !== 'https://hw-kintai-v2.vercel.app' && req.headers.origin !== 'https://' + process.env.VERCEL_URL) return res.status(403).json({ok:false,error:'invalid_origin'});
      if (!String(req.headers['content-type'] || '').includes('application/json')) return res.status(415).json({ok:false,error:'json_required'});
      const body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
      if (!body || Buffer.byteLength(JSON.stringify(body)) > 24000 || !WRITE_ACTIONS.has(body.action)) return res.status(400).json({ok:false,error:'invalid_action'});
      if (!body.requestId || typeof body.requestId !== 'string' || body.requestId.length > 100) return res.status(400).json({ok:false,error:'request_id_required'});
      params = {action:'add',receipt:randomUUID()};
      for (const key of ['requestId','id','name','type','time','month','transport','memo','category','allocations','correction','reason']) {
        if (body[key] !== undefined) {
          if (typeof body[key] !== 'string') return res.status(400).json({ok:false,error:'invalid_field'});
          params[key] = body[key];
        }
      }
    }
    const upstream = await fetch(req.method === 'GET' ? GAS_URL + '?' + new URLSearchParams(params) : GAS_URL, {
      method:req.method, redirect:'follow', cache:'no-store', signal:AbortSignal.timeout(25000),
      ...(req.method === 'POST' ? {body:new URLSearchParams(params)} : {})
    });
    if (!upstream.ok) throw new Error('upstream');
    const data = await upstream.json();
    if (!data || typeof data.ok !== 'boolean') throw new Error('invalid_response');
    // Allow only the public attendance response; never relay a server stack trace.
    const output = {};
    for (const key of ['ok','action','id','duplicate','memo','allocations','corrected','logs','users','transportationCosts','scope','month','months','schemaVersion','revision','serverTime','error','message','actualMinutes']) if (data[key] !== undefined) output[key] = data[key];
    return res.status(200).json(output);
  } catch (error) {
    const timeout = ['TimeoutError','AbortError'].includes(error.name);
    return res.status(timeout ? 504 : 502).json({ok:false,error:timeout?'upstream_timeout':'upstream_error',message:'保存結果・最新情報を確認できませんでした。同じ操作を再試行してください。'});
  }
};
