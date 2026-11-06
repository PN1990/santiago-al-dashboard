# Manutenção

## 1. A regra dos 60 dias do GitHub

O GitHub **desativa automaticamente os workflows agendados (`schedule`) em
repositórios públicos após 60 dias sem atividade de commits**. Quando isso
acontece o bot simplesmente deixa de correr, sem erro e sem aviso visível —
não há corrida vermelha, não há nada.

Consequência prática: **se o repositório ficar dois meses quieto, a importação
diária para em silêncio.** Qualquer commit reinicia a contagem.

Se o Pedro disser "a app não está a atualizar os dados" e não houver corridas
recentes em Actions, verificar isto primeiro: o GitHub mostra um aviso no topo
da página de Actions com um botão para reativar os workflows.

## 2. Privacidade dos artifacts — o repositório é público

**Artifacts de repositórios públicos são descarregáveis por qualquer pessoa**,
sem sessão, só com o link. Durante um período o workflow publicava
`/tmp/tg_*.png` com `if: always()` — ou seja, em todas as corridas, incluindo
as que corriam bem — e os screenshots `tg_4*`/`tg_5*`/`tg_6*` são tirados
depois do login e mostram a lista de reservas **com nomes de hóspedes**.

Está corrigido: o passo é agora `if: failure()` e só publica `tg_0*`, `tg_1*` e
`tg_3*` (login e 2FA). O `tg_2_apos_login` ficou de fora de propósito — pode já
mostrar o dashboard.

Ao mexer nesse passo, **nunca** voltar a `/tmp/tg_*.png` nem a `if: always()`.

Duas coisas que são tratadas fora do repositório:

- **Apagar os artifacts já publicados.** A correção ao workflow só protege o
  futuro; os artifacts antigos continuam lá até serem apagados ou expirarem.

  > **Filtrar sempre por `name == "debug-screenshots"`.** Apagar *todos* os
  > artifacts do repositório parte o deploy do GitHub Pages: o workflow
  > `pages build and deployment` empacota o site num artifact temporário
  > chamado **`github-pages`**, que o passo `deploy` vai buscar poucos
  > segundos depois. Se for apagado nesse intervalo, o deploy falha com
  > `No artifacts named "github-pages" were found`. Já aconteceu, a 6 de
  > novembro de 2026. Não é grave — o site continua a servir a versão
  > anterior e basta re-correr o workflow — mas evita-se assim:
  >
  > ```bash
  > gh api -X GET repos/PN1990/santiago-al-dashboard/actions/artifacts \
  >   --paginate -q '.artifacts[] | select(.name=="debug-screenshots") | .id' \
  >   | xargs -I{} gh api -X DELETE \
  >       repos/PN1990/santiago-al-dashboard/actions/artifacts/{}
  > ```
- **Baixar a retenção de artifacts** em Settings → Actions → General, de 90
  dias para 1 dia.

## 3. O que não tem backup

Re-correr o bot recupera **as reservas**. Não recupera os **campos manuais**:

```
horas de check-in/check-out editadas à mão, caução (necessária, cobrada,
valor), pessoas extra e custo, notas internas, dados_pessoais_ok
```

Esses só existem no D1. Não há exportação nem cópia automática. Se o D1 se
perder, perdem-se.

Já foi oferecido um backup automático desses campos; o Pedro ainda não pediu
para avançar.

## 4. Secrets

Dois sítios diferentes, não confundir:

- **Secrets do repositório** (Settings → Secrets → Actions) — usados pelos
  workflows: `TALKGUEST_*`, `MAIL_2FA_*`, `API_URL`, `BOT_SECRET`,
  `CLOUDFLARE_API_TOKEN`, `CLOUDFLARE_ACCOUNT_ID`.
- **Secrets do Worker** (`wrangler secret put`) — usados em runtime pela API:
  `APP_PASSWORD`, `AUTH_SECRET`, `BOT_SECRET`, `ANTHROPIC_API_KEY`. Ver
  `docs/arquitetura.md`.

O `BOT_SECRET` tem de ser **o mesmo valor** nos dois lados.

## 5. Custo

A extração de documentos do módulo SIBA é a única coisa que custa dinheiro
por uso (API da Anthropic), na ordem de **~2 €/ano** ao volume atual.
Pode ficar mais barata pondo `ANTHROPIC_MODEL=claude-haiku-4-5` nos secrets do
Worker, com alguma perda de qualidade na leitura dos documentos.

O resto — GitHub Pages, GitHub Actions em repo público, Cloudflare Workers e
D1 — cabe nos planos gratuitos.
