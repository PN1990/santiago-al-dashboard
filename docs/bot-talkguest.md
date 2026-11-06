# O bot de importação da Talkguest

`talkguest_import.py`, corrido por `.github/workflows/talkguest-import.yml`.

A Talkguest **não tem API**. O bot faz scraping com Selenium + Chrome
headless. A app da Talkguest é feita em OutSystems (ASP.NET WebForms), por
isso os ids dos elementos são coisas como
`OutSystemsUIWeb_wt124$block$...` e há `__VIEWSTATE`/`__OSVSTATE` pelo meio —
mudam de versão para versão, e é por aí que o bot se parte quando se parte.

## O que ele faz, por ordem

1. Abre `https://app.talkguest.com/Theme_UI/Login.aspx` (configurável por
   `TALKGUEST_LOGIN_URL`) e submete email + password.
2. Se a Talkguest pedir **2FA**, escolhe "Enviar código Email", espera, e vai
   **ler o código a uma caixa de correio por IMAP** (ver abaixo).
3. Preenche os 6 dígitos. O site submete sozinho ao escrever o último — só se
   clica em "Submeter" se o formulário ainda estiver lá.
4. Na lista de reservas, põe o filtro de **check-out** de hoje até +364 dias
   (a Talkguest recusa períodos acima de 1 ano), aplica a pesquisa.
5. Ações → "Exportar Reservas" → descarrega o Excel.
6. Lê o Excel com `openpyxl`, normaliza os valores, e faz
   `POST /api/bot/import` com o `BOT_SECRET`.

## O 2FA — o ponto frágil

O código chega por email. O bot lê a caixa por IMAP
(`imap.gmail.com`), com uma **app password** do Gmail, e procura primeiro na
label configurada em `MAIL_2FA_FOLDER` e depois, como rede de segurança, na
`INBOX`. Quando está a ler uma label dedicada não filtra pelo remetente (a
label já é o filtro).

Variáveis de ambiente, todas secrets do repositório:

```
TALKGUEST_EMAIL, TALKGUEST_PASSWORD,
API_URL, BOT_SECRET,
MAIL_2FA_ADDRESS, MAIL_2FA_PASSWORD,
MAIL_2FA_IMAP_HOST (por omissão imap.gmail.com),
MAIL_2FA_FOLDER    (por omissão INBOX)
```

A espera pelo código é o `timeout` de `ler_codigo_2fa()`: **240 segundos**
(era 150 e não chegava).

## Diagnosticar uma corrida vermelha

**Regra prática: se a corrida falhou perto dos 3m40s, é quase sempre o 2FA.**
Os outros pontos de falha desistem muito antes — 15s, 25s, 90s — por isso uma
falha nesse intervalo de tempo aponta para a janela de espera do código.
Nesse caso o primeiro palpite é "o email chegou tarde", não "está mal
configurado", e **voltar a correr o workflow à mão costuma resolver**.

O log do próprio bot é suficiente para distinguir os casos. Procurar:

| No log | Significa |
|---|---|
| `Pasta 'X': N mensagens` | o IMAP ligou e abriu a pasta. Se `N` for 0 durante toda a espera, o email não chegou (ou está noutra label). |
| `Nao consegui abrir a pasta` | `MAIL_2FA_FOLDER` está errado — a label não existe com esse nome. |
| `Erro IMAP (retry)` | problema de ligação ao Gmail; se persistir em todas as rondas, suspeitar da app password. |
| nada disto, falha mais cedo | não é o 2FA — é login, filtro de datas, ou a exportação. Ver os screenshots. |

Causas permanentes a excluir (se uma re-corrida sem alterações funcionar,
**não é nenhuma destas**): app password do Gmail revogada, `MAIL_2FA_FOLDER`
com o nome errado, endereço do remetente mudado.

## Screenshots

O bot grava screenshots em `/tmp/tg_*.png` em cada passo:

- `tg_0*`, `tg_1*` — página de login
- `tg_2_apos_login` — logo depois do login (**pode já mostrar o dashboard**)
- `tg_3*` — ecrã de 2FA (escolha do método, código enviado, código inserido)
- `tg_4*`, `tg_5*`, `tg_6*` — **lista de reservas: mostram nomes de hóspedes**

Por isso o workflow só publica `tg_0*`, `tg_1*` e `tg_3*`, e só quando a
corrida falha. Ver `docs/manutencao.md`.

## Agendamento

```yaml
schedule:
  - cron: '0 7 * * *'    # corrida principal
  - cron: '0 12 * * *'   # rede de segurança
```

O segundo agendamento é seguro porque a importação é idempotente (ver
`docs/arquitetura.md`): substitui sempre a lista completa de reservas e
preserva os campos manuais, logo correr duas vezes no mesmo dia não duplica
nem perde nada.

Também se pode correr à mão: Actions → "Importar Reservas Talkguest" → Run
workflow (`workflow_dispatch`).

## Por fazer

O bot **não grava** os boletins de alojamento na Talkguest. A infraestrutura
está feita do lado da API (coluna `gravado_em`, endpoints `/api/bot/siba` e
`/api/bot/siba/gravado`, badge na app), mas a navegação Selenium do ecrã
"Registo de Viajantes" nunca foi escrita. Quando for: a primeira corrida deve
**parar antes de gravar**, para o Pedro validar o que ia ser submetido.
