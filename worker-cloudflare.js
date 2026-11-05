// Puente gratuito (Cloudflare Workers) para dos cosas:
//  1) precios de bolsa (Yahoo Finance / Stooq)
//  2) mantener tu sesión de Google abierta para siempre (renueva el acceso sin que tengas que volver a entrar)
// Variables que debes crear en el Worker (Settings > Variables and Secrets):
//   GOOGLE_CLIENT_ID      (texto)     tu ID de cliente de Google
//   GOOGLE_CLIENT_SECRET  (secreto)   la clave secreta de ese ID de cliente
//   ALLOWED_ORIGIN        (texto)     la dirección de tu app, sin barra final, ej: https://tu-usuario.github.io
const PERMITIDOS = ['query1.finance.yahoo.com', 'query2.finance.yahoo.com', 'stooq.com'];

function cors(req, env) {
  const o = req.headers.get('Origin') || '';
  const lista = String(env.ALLOWED_ORIGIN || '').split(',').map(function (x) { return x.trim().replace(/\/+$/, ''); }).filter(Boolean);
  const ok = lista.length === 0 || lista.indexOf(o) !== -1;
  return { ok: ok, h: { 'Access-Control-Allow-Origin': ok && o ? o : (lista[0] || '*'), 'Access-Control-Allow-Headers': 'Content-Type', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS', 'Vary': 'Origin' } };
}
function json(obj, status, h) { return new Response(JSON.stringify(obj), { status: status || 200, headers: Object.assign({ 'Content-Type': 'application/json' }, h) }); }

export default {
  async fetch(req, env) {
    const c = cors(req, env);
    const url = new URL(req.url);
    if (req.method === 'OPTIONS') return new Response(null, { headers: c.h });

    // --- Google: intercambio de código y renovación de sesión ---
    if (url.pathname === '/google/code' || url.pathname === '/google/refresh') {
      if (!c.ok) return json({ error: 'origen_no_permitido' }, 403, c.h);
      if (req.method !== 'POST') return json({ error: 'usa_POST' }, 405, c.h);
      if (!env.GOOGLE_CLIENT_ID || !env.GOOGLE_CLIENT_SECRET) return json({ error: 'faltan_variables_en_el_worker' }, 500, c.h);
      let b = {};
      try { b = await req.json(); } catch (e) { return json({ error: 'cuerpo_invalido' }, 400, c.h); }
      const p = new URLSearchParams({ client_id: env.GOOGLE_CLIENT_ID, client_secret: env.GOOGLE_CLIENT_SECRET });
      if (url.pathname === '/google/code') {
        if (!b.code) return json({ error: 'falta_code' }, 400, c.h);
        p.set('grant_type', 'authorization_code'); p.set('code', b.code); p.set('redirect_uri', b.redirect_uri || 'postmessage');
      } else {
        if (!b.refresh_token) return json({ error: 'falta_refresh_token' }, 400, c.h);
        p.set('grant_type', 'refresh_token'); p.set('refresh_token', b.refresh_token);
      }
      const r = await fetch('https://oauth2.googleapis.com/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: p.toString() });
      const j = await r.json();
      const salida = { access_token: j.access_token, expires_in: j.expires_in, refresh_token: j.refresh_token, error: j.error, error_description: j.error_description };
      return json(salida, r.ok ? 200 : 400, c.h);
    }

    // --- Bolsa: reenvía consultas permitidas ---
    const destino = url.searchParams.get('u');
    let u;
    try { u = new URL(destino); } catch (e) { return new Response('Falta ?u=', { status: 400, headers: c.h }); }
    if (u.protocol !== 'https:' || PERMITIDOS.indexOf(u.hostname) === -1) return new Response('Sitio no permitido', { status: 403, headers: c.h });
    const r = await fetch(u.toString(), { headers: { 'User-Agent': 'Mozilla/5.0 (compatible; FinanzasApp/1.0)', 'Accept': '*/*' }, cf: { cacheTtl: 60, cacheEverything: true } });
    const h = new Headers(c.h);
    h.set('Content-Type', r.headers.get('Content-Type') || 'text/plain');
    return new Response(r.body, { status: r.status, headers: h });
  }
};
