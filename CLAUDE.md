# PNStays — Dashboard Santiago AL

Dashboard privado de gestão de 3 apartamentos de alojamento local no Funchal
(Pedro Freitas / PNStays). Está em **https://reservas.pnstays.com** e é usado
sobretudo no telemóvel, instalado como app (PWA).

**Idioma: tudo em português de Portugal** — código, comentários, commits,
interface e conversa com o utilizador.

---

## As três partes do sistema

| Parte | Onde vive | Onde corre |
|---|---|---|
| **Frontend** — `index.html` (app inteira num ficheiro) + `sw.js` | raiz do repo | GitHub Pages, domínio `reservas.pnstays.com` (DNS na Cloudflare, ver `CNAME`) |
| **API** — `worker/src/index.js` + `worker/wrangler.toml` | pasta `worker/` | Cloudflare Workers, projeto **`santiago-al-dashboard-api`**, base de dados **D1 `santiago-al-dashboard`** |
| **Bot** — `talkguest_import.py` + `.github/workflows/` | raiz e `.github/` | GitHub Actions |

Não há Supabase, nem backend próprio, nem node_modules: **os dados estão todos
no D1 da Cloudflare** e o Worker não tem dependências npm (de propósito — usa
`fetch` em cru, para o deploy não precisar de instalar nada).

O frontend fala com a API em
`https://santiago-al-dashboard-api.pnstays.workers.dev` (constante `API_BASE`
em `index.html`).

## Deploys

- **Frontend:** basta commit para `main` — o GitHub Pages serve a raiz.
  O `sw.js` é *network-first* no HTML, por isso os dispositivos apanham a
  versão nova sozinhos (ver `docs/frontend.md` antes de mexer no service
  worker — já houve um ciclo infinito de recargas aqui).
- **Worker:** push para `main` que toque em `worker/**` dispara
  `.github/workflows/deploy-worker.yml` (`npx wrangler@4 deploy`).
  Alterações aos workflows ou ao bot **não** disparam deploy.

## Onde ler mais

- `docs/arquitetura.md` — estrutura, base de dados, endpoints, segredos.
- `docs/bot-talkguest.md` — como o bot funciona, a idempotência da
  importação, e **como diagnosticar uma falha pelo log** (ler isto antes de
  investigar qualquer corrida vermelha).
- `docs/frontend.md` — a app, o service worker, o módulo SIBA.
- `docs/manutencao.md` — a regra dos 60 dias do GitHub, privacidade dos
  artifacts, o que não tem backup.

## Dois avisos que custam caro esquecer

1. **O repositório é público.** Artifacts de repositórios públicos são
   descarregáveis por qualquer pessoa, e os screenshots do bot tirados depois
   do login mostram nomes de hóspedes. Ver `docs/manutencao.md`.
2. **Workflows agendados param sozinhos.** O GitHub desativa `schedule` em
   repositórios públicos após **60 dias sem commits**. Qualquer commit
   reinicia a contagem. Ver `docs/manutencao.md`.
