# Arquitetura

## Fluxo dos dados

```
Talkguest (app.talkguest.com)
      │  Selenium + Chrome (GitHub Actions, 1x/dia)
      │  exporta Excel das reservas
      ▼
talkguest_import.py  ──POST /api/bot/import (BOT_SECRET)──►  Worker
                                                               │
                                                               ▼
                                                      Cloudflare D1
                                                   "santiago-al-dashboard"
                                                               ▲
                                                               │  token de sessão
index.html (GitHub Pages, reservas.pnstays.com)  ──────────────┘
```

O mesmo mapeamento de colunas do Excel existe em dois sítios: em Python
(`openpyxl`, no bot) e em JavaScript (`SheetJS`, na importação manual pelo
browser). **Se mudar um, mudar o outro.**

## Worker (`worker/src/index.js`)

Ficheiro único, zero dependências npm. Guarda as credenciais do lado do
servidor; o browser só fala com ele com um token de sessão assinado (HMAC
SHA-256, validade 30 dias).

### Segredos (postos com `wrangler secret put`, nunca no repo)

| Segredo | Para quê |
|---|---|
| `APP_PASSWORD` | password única de login da família |
| `AUTH_SECRET` | assina os tokens de sessão |
| `BOT_SECRET` | autentica o bot do GitHub Actions |
| `ANTHROPIC_API_KEY` | extração de dados de documentos (módulo SIBA) |
| `ANTHROPIC_MODEL` | opcional; por omissão `claude-opus-5` |
| `SIBA_RETENCAO_DIAS` | opcional; por omissão 5 dias |

### Endpoints

**Browser (token de sessão no header `Authorization`):**

- `POST /api/login` — devolve o token.
- `GET /api/reservas` — devolve `{ data, meta }`; o `meta` diz como e quando
  foi a última importação (bot ou manual, data/hora) e alimenta a barra de
  proveniência na app.
- `POST /api/reservas/import` — importação manual do Excel.
- `PATCH /api/reservas/:id` — grava os campos manuais de uma reserva.
- `GET|POST|DELETE /api/siba` — hóspedes do registo de viajantes.
- `POST /api/siba/extrair` — foto de documento → dados (Anthropic, visão).
- `POST /api/push-subscriptions` — subscrições de notificações push.

**Bot (header com `BOT_SECRET`):**

- `POST /api/bot/import` — a importação diária.
- `GET /api/bot/siba` — hóspedes ainda por gravar na Talkguest.
- `POST /api/bot/siba/gravado` — marca um hóspede como gravado.

### Tabelas D1

- **`reservas`** — uma linha por reserva, `id` = id da Talkguest.
- **`meta`** — chave/valor com o registo da última importação.
- **`siba_hospedes`** — dados de documentos, **temporários**: cada linha tem
  `expira_em` e é apagada ao fim de 5 dias (`limparSibaExpirados()`). É
  intencional: os dados têm de ser enviados ao SIBA nos primeiros 3 dias da
  estadia e não podem ser guardados depois disso.

### Campos manuais vs. importados

`replaceReservas()` **apaga todas as reservas e reinsere**, mas antes faz
snapshot dos campos manuais e volta a fundi-los por `id`:

```
hora_checkin, hora_checkout, hora_checkin_manual, hora_checkout_manual,
caucao_necessaria, caucao_cobrada, caucao_valor, pessoas_extra,
custo_pessoa_extra, notas_internas, dados_pessoais_ok
```

As horas têm uma flag `_manual`: se o Pedro editou a hora à mão, a
importação não a volta a pisar.

É isto que torna a importação **idempotente** — correr o bot duas vezes no
mesmo dia dá o mesmo resultado que correr uma. Daí haver um segundo
agendamento ao meio-dia como rede de segurança (ver `docs/bot-talkguest.md`).

Salvaguarda importante no bot: se o Excel vier sem reservas, ele **aborta em
vez de importar** — senão o DELETE apagava tudo.

## Normalizações de valores

Feitas na importação, para a app ter nomes consistentes:

- `Casa Santiago II` → `Casa Santiago II Funchal`
- `Booking.com` → `Booking`
- `Airbnb` → `AirBnB`
- as duas colunas de crianças do Excel são somadas numa só
- linhas `Indisponivel` e `Interno` são ignoradas (bloqueios, não reservas)
