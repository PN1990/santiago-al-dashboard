// API gatekeeper for the PNStays / Santiago AL dashboard.
// Keeps the database credentials server-side; the browser only ever
// talks to this Worker over an authenticated session token.

const SESSION_DAYS = 30;

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, PATCH, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', ...CORS_HEADERS },
  });
}

function bytesToBase64Url(bytes) {
  let bin = '';
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function base64UrlToBytes(b64url) {
  const b64 = b64url.replace(/-/g, '+').replace(/_/g, '/').padEnd(b64url.length + (4 - (b64url.length % 4)) % 4, '=');
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

async function hmacKey(secret) {
  return crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(secret),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign', 'verify']
  );
}

async function createToken(secret) {
  const payload = { exp: Date.now() + SESSION_DAYS * 24 * 60 * 60 * 1000 };
  const payloadBytes = new TextEncoder().encode(JSON.stringify(payload));
  const payloadB64 = bytesToBase64Url(payloadBytes);
  const key = await hmacKey(secret);
  const sig = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(payloadB64));
  const sigB64 = bytesToBase64Url(new Uint8Array(sig));
  return `${payloadB64}.${sigB64}`;
}

async function verifyToken(token, secret) {
  if (!token || !token.includes('.')) return false;
  const [payloadB64, sigB64] = token.split('.');
  const key = await hmacKey(secret);
  const valid = await crypto.subtle.verify(
    'HMAC',
    key,
    base64UrlToBytes(sigB64),
    new TextEncoder().encode(payloadB64)
  );
  if (!valid) return false;
  try {
    const payload = JSON.parse(new TextDecoder().decode(base64UrlToBytes(payloadB64)));
    return typeof payload.exp === 'number' && payload.exp > Date.now();
  } catch {
    return false;
  }
}

function getBearer(request) {
  const auth = request.headers.get('Authorization') || '';
  const match = auth.match(/^Bearer\s+(.+)$/i);
  return match ? match[1] : null;
}

async function requireSession(request, env) {
  const token = getBearer(request);
  return token ? verifyToken(token, env.AUTH_SECRET) : false;
}

// Columns the manual XLS import path is allowed to merge from existing rows.
const MANUAL_FIELDS = [
  'hora_checkin', 'hora_checkout', 'hora_checkin_manual', 'hora_checkout_manual',
  'caucao_necessaria', 'caucao_cobrada', 'caucao_valor', 'pessoas_extra',
  'custo_pessoa_extra', 'notas_internas', 'dados_pessoais_ok',
];

const RESERVA_IMPORT_FIELDS = [
  'id', 'hospede', 'checkin', 'hora_checkin', 'checkout', 'hora_checkout',
  'noites', 'adultos', 'criancas', 'bebes', 'telefone', 'email', 'pais',
  'codigo_pais', 'alojamento', 'tmt', 'total', 'estado', 'estado_pagamento',
  'canal', 'comissao', 'comissao_pct', 'id_canal', 'data_criacao',
  'antecedencia', 'checkin_efetuado', 'checkout_efetuado', 'notas_canal',
  'fatura', 'estado_aima',
];

function toDbBool(v) {
  return v ? 1 : 0;
}

async function replaceReservas(db, incoming) {
  // 1. Snapshot manual fields of existing rows
  const existing = await db.prepare(`SELECT id, ${MANUAL_FIELDS.join(', ')} FROM reservas`).all();
  const manuais = {};
  for (const row of existing.results) manuais[row.id] = row;

  // 2. Build merged rows
  const merged = incoming.map((r) => {
    const m = manuais[r.id] || {};
    const novaHoraCheckin = m.hora_checkin_manual ? m.hora_checkin : (r.hora_checkin || m.hora_checkin || '');
    const novaHoraCheckout = m.hora_checkout_manual ? m.hora_checkout : (r.hora_checkout || m.hora_checkout || '');
    return {
      ...r,
      hora_checkin: novaHoraCheckin,
      hora_checkout: novaHoraCheckout,
      hora_checkin_manual: toDbBool(m.hora_checkin_manual),
      hora_checkout_manual: toDbBool(m.hora_checkout_manual),
      caucao_necessaria: m.caucao_necessaria ?? r.caucao_necessaria ?? null,
      caucao_cobrada: toDbBool(m.caucao_cobrada ?? r.caucao_cobrada),
      caucao_valor: m.caucao_valor ?? r.caucao_valor ?? 0,
      pessoas_extra: m.pessoas_extra ?? r.pessoas_extra ?? 0,
      custo_pessoa_extra: m.custo_pessoa_extra ?? r.custo_pessoa_extra ?? 0,
      notas_internas: m.notas_internas ?? r.notas_internas ?? '',
      dados_pessoais_ok: toDbBool(m.dados_pessoais_ok ?? r.dados_pessoais_ok),
    };
  });

  // 3. Replace all rows atomically
  const stmts = [db.prepare('DELETE FROM reservas')];
  const cols = [...RESERVA_IMPORT_FIELDS, 'hora_checkin_manual', 'hora_checkout_manual',
    'caucao_necessaria', 'caucao_cobrada', 'caucao_valor', 'pessoas_extra',
    'custo_pessoa_extra', 'notas_internas', 'dados_pessoais_ok'];
  const placeholders = cols.map(() => '?').join(', ');
  const insertSql = `INSERT INTO reservas (${cols.join(', ')}) VALUES (${placeholders})`;
  for (const r of merged) {
    stmts.push(db.prepare(insertSql).bind(...cols.map((c) => r[c] ?? null)));
  }
  await db.batch(stmts);
  return merged.length;
}

// Regista quem/quando fez a última importação (origem 'bot' ou 'manual').
async function registarImport(db, origem, count) {
  await db.prepare('CREATE TABLE IF NOT EXISTS meta (chave TEXT PRIMARY KEY, valor TEXT)').run();
  const upsert = (chave, valor) =>
    db.prepare(`INSERT INTO meta (chave, valor) VALUES (?, ?)
                ON CONFLICT(chave) DO UPDATE SET valor = excluded.valor`).bind(chave, valor);
  await db.batch([
    upsert('last_import_at', new Date().toISOString()),
    upsert('last_import_source', origem),
    upsert('last_import_count', String(count)),
  ]);
}

async function lerMeta(db) {
  try {
    const { results } = await db.prepare('SELECT chave, valor FROM meta').all();
    const meta = {};
    for (const row of results) meta[row.chave] = row.valor;
    return meta;
  } catch {
    return {}; // tabela ainda não existe (antes da 1ª importação após deploy)
  }
}

const ALLOWED_UPDATE_FIELDS = [
  'hora_checkin', 'hora_checkout', 'hora_checkin_manual', 'hora_checkout_manual',
  'pessoas_extra', 'custo_pessoa_extra', 'caucao_valor', 'caucao_necessaria',
  'caucao_cobrada', 'notas_internas', 'dados_pessoais_ok', 'updated_at',
];

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const { pathname } = url;

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: CORS_HEADERS });
    }

    try {
      // ---- Public: login ----
      if (pathname === '/api/login' && request.method === 'POST') {
        const body = await request.json().catch(() => ({}));
        if (typeof body.password !== 'string' || body.password !== env.APP_PASSWORD) {
          return json({ error: 'Password incorreta' }, 401);
        }
        const token = await createToken(env.AUTH_SECRET);
        return json({ token });
      }

      // ---- Bot ingest (separate static secret, no session) ----
      if (pathname === '/api/bot/import' && request.method === 'POST') {
        const token = getBearer(request);
        if (!token || token !== env.BOT_SECRET) {
          return json({ error: 'Não autorizado' }, 401);
        }
        const body = await request.json().catch(() => null);
        if (!body || !Array.isArray(body.reservas)) {
          return json({ error: 'Payload inválido' }, 400);
        }
        const count = await replaceReservas(env.DB, body.reservas);
        try { await registarImport(env.DB, 'bot', count); } catch (e) { /* não falhar o import */ }
        return json({ ok: true, count });
      }

      // ---- Everything below requires a valid session ----
      if (!(await requireSession(request, env))) {
        return json({ error: 'Sessão inválida ou expirada' }, 401);
      }

      if (pathname === '/api/reservas' && request.method === 'GET') {
        const { results } = await env.DB.prepare('SELECT * FROM reservas ORDER BY checkin').all();
        const meta = await lerMeta(env.DB);
        return json({ data: results, meta });
      }

      if (pathname === '/api/reservas/import' && request.method === 'POST') {
        const body = await request.json().catch(() => null);
        if (!body || !Array.isArray(body.reservas)) {
          return json({ error: 'Payload inválido' }, 400);
        }
        const count = await replaceReservas(env.DB, body.reservas);
        try { await registarImport(env.DB, 'manual', count); } catch (e) { /* não falhar o import */ }
        return json({ ok: true, count });
      }

      const reservaMatch = pathname.match(/^\/api\/reservas\/([^/]+)$/);
      if (reservaMatch && request.method === 'PATCH') {
        const id = decodeURIComponent(reservaMatch[1]);
        const body = await request.json().catch(() => null);
        if (!body || typeof body !== 'object') return json({ error: 'Payload inválido' }, 400);

        const fields = Object.keys(body).filter((k) => ALLOWED_UPDATE_FIELDS.includes(k));
        if (fields.length === 0) return json({ error: 'Nenhum campo válido' }, 400);

        const setClause = fields.map((f) => `${f} = ?`).join(', ');
        const values = fields.map((f) => {
          const v = body[f];
          if (typeof v === 'boolean') return toDbBool(v);
          return v;
        });
        await env.DB.prepare(`UPDATE reservas SET ${setClause} WHERE id = ?`)
          .bind(...values, id)
          .run();
        return json({ ok: true });
      }

      if (pathname === '/api/push-subscriptions' && request.method === 'POST') {
        const body = await request.json().catch(() => null);
        if (!body || !body.endpoint || !body.p256dh || !body.auth) {
          return json({ error: 'Payload inválido' }, 400);
        }
        await env.DB.prepare(
          `INSERT INTO push_subscriptions (endpoint, p256dh, auth) VALUES (?, ?, ?)
           ON CONFLICT(endpoint) DO UPDATE SET p256dh = excluded.p256dh, auth = excluded.auth`
        ).bind(body.endpoint, body.p256dh, body.auth).run();
        return json({ ok: true });
      }

      return json({ error: 'Não encontrado' }, 404);
    } catch (err) {
      return json({ error: err.message || 'Erro interno' }, 500);
    }
  },
};
