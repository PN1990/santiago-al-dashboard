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

// ── SIBA / Registo de Viajantes ──────────────────────────────────────────────
// Dados de identificação dos hóspedes, guardados SÓ o tempo necessário para
// comunicar o registo. A imagem do documento nunca é guardada — apenas os
// campos extraídos dela. Cada linha traz o seu próprio prazo de validade e é
// apagada automaticamente, independentemente do ciclo de vida das reservas.
const SIBA_COLS = ['nome_completo', 'data_nascimento', 'nacionalidade', 'pais_residencia',
  'tipo_documento', 'numero_documento', 'pais_emissor'];

const SIBA_RETENCAO_DIAS_DEFAULT = 5;

async function garantirTabelaSiba(db) {
  await db.prepare(`CREATE TABLE IF NOT EXISTS siba_hospedes (
    id TEXT PRIMARY KEY,
    reserva_id TEXT NOT NULL,
    ${SIBA_COLS.map((c) => `${c} TEXT`).join(', ')},
    criado_em TEXT,
    expira_em TEXT,
    gravado_em TEXT
  )`).run();
  // Tabelas criadas antes desta coluna existir precisam de a receber agora.
  try { await db.prepare('ALTER TABLE siba_hospedes ADD COLUMN gravado_em TEXT').run(); }
  catch (e) { /* já existe */ }
}

// Apaga tudo o que passou do prazo. Corre a cada pedido ao SIBA — é barato e
// dispensa um agendador só para isto.
async function limparSibaExpirados(db) {
  try {
    await db.prepare('DELETE FROM siba_hospedes WHERE expira_em IS NOT NULL AND expira_em < ?')
      .bind(new Date().toISOString()).run();
  } catch (e) { /* tabela ainda não criada */ }
}

const SIBA_PROMPT = `Lês documentos de identificação (passaportes e cartões de identidade) para preencher o Registo de Viajantes (SIBA) em Portugal.

Responde APENAS com um objeto JSON, sem texto à volta e sem blocos de código.

Campos:
- "nome_completo": nome completo como aparece no documento, na ordem natural (nome próprio antes do apelido). Usa capitalização normal, não tudo em maiúsculas.
- "data_nascimento": no formato DD-MM-YYYY.
- "nacionalidade": nome do país EM PORTUGUÊS, na forma curta e comum em Portugal (ex: "Alemanha", "Estados Unidos", "Reino Unido", "França", "Polónia", "Países Baixos", "Finlândia", "Hungria").
- "tipo_documento": exatamente um de "Passaporte", "Cartão de Cidadão" ou "Outro". Passaportes → "Passaporte"; cartões de identidade nacionais (de qualquer país) → "Cartão de Cidadão"; tudo o resto (carta de condução, título de residência) → "Outro".
- "numero_documento": número do documento, sem espaços.
- "pais_emissor": país emissor EM PORTUGUÊS, mesmo critério da nacionalidade.
- "confianca": "alta" se leste a zona de leitura ótica (MRZ) sem ambiguidade, "media" se leste do texto impresso, "baixa" se a imagem está pouco legível.

Se um campo não for legível, põe null nesse campo. Nunca inventes dados.`;

async function extrairDocumento(env, imagemBase64, mediaType) {
  if (!env.ANTHROPIC_API_KEY) {
    throw new Error('Falta o secret ANTHROPIC_API_KEY no Worker.');
  }
  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model: env.ANTHROPIC_MODEL || 'claude-opus-5',
      max_tokens: 1024,
      output_config: { effort: 'low' },
      system: SIBA_PROMPT,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mediaType, data: imagemBase64 } },
          { type: 'text', text: 'Extrai os dados deste documento.' },
        ],
      }],
    }),
  });

  if (!resp.ok) {
    const detalhe = await resp.text().catch(() => '');
    throw new Error(`API de leitura devolveu ${resp.status}: ${detalhe.slice(0, 300)}`);
  }
  const data = await resp.json();
  if (data.stop_reason === 'refusal') {
    throw new Error('A leitura do documento foi recusada. Tenta outra fotografia.');
  }
  const texto = (data.content || [])
    .filter((b) => b.type === 'text')
    .map((b) => b.text)
    .join('\n');
  const inicio = texto.indexOf('{');
  const fim = texto.lastIndexOf('}');
  if (inicio === -1 || fim === -1) {
    throw new Error('Não consegui ler dados nesta imagem. Tenta uma foto mais nítida.');
  }
  return JSON.parse(texto.slice(inicio, fim + 1));
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

      // Hóspedes por gravar no SIBA da Talkguest. O bot grava-os lá (sem
      // comunicar) e marca-os como gravados, para não os repetir na corrida
      // seguinte — a lista da Talkguest não é idempotente.
      if (pathname === '/api/bot/siba' && request.method === 'GET') {
        const token = getBearer(request);
        if (!token || token !== env.BOT_SECRET) return json({ error: 'Não autorizado' }, 401);
        await garantirTabelaSiba(env.DB);
        await limparSibaExpirados(env.DB);
        const { results } = await env.DB.prepare(
          `SELECT h.*, r.hospede AS reserva_hospede, r.alojamento, r.checkin, r.checkout
             FROM siba_hospedes h
             LEFT JOIN reservas r ON r.id = h.reserva_id
            WHERE h.gravado_em IS NULL
            ORDER BY h.reserva_id, h.criado_em`).all();
        return json({ data: results });
      }

      if (pathname === '/api/bot/siba/gravado' && request.method === 'POST') {
        const token = getBearer(request);
        if (!token || token !== env.BOT_SECRET) return json({ error: 'Não autorizado' }, 401);
        const body = await request.json().catch(() => null);
        if (!body || !Array.isArray(body.ids) || body.ids.length === 0) {
          return json({ error: 'Payload inválido' }, 400);
        }
        await garantirTabelaSiba(env.DB);
        const marcas = body.ids.map(() => '?').join(', ');
        await env.DB.prepare(
          `UPDATE siba_hospedes SET gravado_em = ? WHERE id IN (${marcas})`
        ).bind(new Date().toISOString(), ...body.ids).run();
        return json({ ok: true, marcados: body.ids.length });
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

      // ---- SIBA: dados de identificação (temporários) ----
      if (pathname === '/api/siba' && request.method === 'GET') {
        const reservaId = url.searchParams.get('reserva_id');
        if (!reservaId) return json({ error: 'reserva_id em falta' }, 400);
        await garantirTabelaSiba(env.DB);
        await limparSibaExpirados(env.DB);
        const { results } = await env.DB
          .prepare('SELECT * FROM siba_hospedes WHERE reserva_id = ? ORDER BY criado_em')
          .bind(reservaId).all();
        return json({ data: results });
      }

      if (pathname === '/api/siba' && request.method === 'POST') {
        const body = await request.json().catch(() => null);
        if (!body || !body.reserva_id) return json({ error: 'Payload inválido' }, 400);
        await garantirTabelaSiba(env.DB);
        await limparSibaExpirados(env.DB);

        const dias = Number(env.SIBA_RETENCAO_DIAS) || SIBA_RETENCAO_DIAS_DEFAULT;
        const expira = new Date(Date.now() + dias * 86400000).toISOString();
        const id = body.id || crypto.randomUUID();
        const valores = SIBA_COLS.map((c) => (body[c] ?? null));

        await env.DB.prepare(
          `INSERT INTO siba_hospedes (id, reserva_id, ${SIBA_COLS.join(', ')}, criado_em, expira_em)
           VALUES (?, ?, ${SIBA_COLS.map(() => '?').join(', ')}, ?, ?)
           ON CONFLICT(id) DO UPDATE SET
             ${SIBA_COLS.map((c) => `${c} = excluded.${c}`).join(', ')},
             expira_em = excluded.expira_em`
        ).bind(id, body.reserva_id, ...valores, new Date().toISOString(), expira).run();

        return json({ ok: true, id, expira_em: expira });
      }

      if (pathname === '/api/siba' && request.method === 'DELETE') {
        const id = url.searchParams.get('id');
        const reservaId = url.searchParams.get('reserva_id');
        await garantirTabelaSiba(env.DB);
        if (id) {
          await env.DB.prepare('DELETE FROM siba_hospedes WHERE id = ?').bind(id).run();
        } else if (reservaId) {
          await env.DB.prepare('DELETE FROM siba_hospedes WHERE reserva_id = ?').bind(reservaId).run();
        } else {
          return json({ error: 'Indica id ou reserva_id' }, 400);
        }
        return json({ ok: true });
      }

      if (pathname === '/api/siba/extrair' && request.method === 'POST') {
        const body = await request.json().catch(() => null);
        if (!body || !body.imagem) return json({ error: 'Imagem em falta' }, 400);
        try {
          const campos = await extrairDocumento(
            env, body.imagem, body.media_type || 'image/jpeg');
          return json({ ok: true, campos });
        } catch (e) {
          return json({ error: e.message }, 502);
        }
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
