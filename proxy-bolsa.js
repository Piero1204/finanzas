// Proxy gratuito para precios de bolsa (Cloudflare Workers). Pega este código tal cual en un Worker nuevo.
// Solo reenvía consultas a Yahoo Finance y Stooq, y agrega el permiso (CORS) que los celulares necesitan.
const PERMITIDOS = ['query1.finance.yahoo.com', 'query2.finance.yahoo.com', 'stooq.com'];
const CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': '*', 'Access-Control-Allow-Methods': 'GET,OPTIONS' };

export default {
  async fetch(req) {
    if (req.method === 'OPTIONS') return new Response(null, { headers: CORS });
    const destino = new URL(req.url).searchParams.get('u');
    let u;
    try { u = new URL(destino); } catch (e) { return new Response('Falta ?u=', { status: 400, headers: CORS }); }
    if (u.protocol !== 'https:' || PERMITIDOS.indexOf(u.hostname) === -1) return new Response('Sitio no permitido', { status: 403, headers: CORS });
    const r = await fetch(u.toString(), { headers: { 'User-Agent': 'Mozilla/5.0 (compatible; FinanzasApp/1.0)', 'Accept': '*/*' }, cf: { cacheTtl: 60, cacheEverything: true } });
    const h = new Headers(CORS);
    h.set('Content-Type', r.headers.get('Content-Type') || 'text/plain');
    return new Response(r.body, { status: r.status, headers: h });
  }
};
