# Frontend

`index.html` é a app toda num ficheiro: HTML, CSS e JS inline, mais Chart.js e
SheetJS por CDN. Não há build. Commit para `main` = publicado.

Está pensada para telemóvel e instalada como app (PWA). Usada pelo Pedro e
pela mãe.

## Service worker (`sw.js`) — ler antes de mexer

Estratégia: **network-first só na navegação** (o documento HTML). Tudo o
resto — Chart.js, SheetJS, chamadas à API noutra origem — passa direto sem ser
interceptado. Offline, serve a última versão guardada em cache.

O registo em `index.html` faz `reg.update()` ao carregar e outra vez quando a
página volta a ficar visível. **Não recarrega a página automaticamente.**

> **Não voltar a adicionar recarga automática.** Já existiu um
> `controllerchange` → `location.reload()` com uma guarda numa variável do
> escopo da página. A variável não sobrevive à recarga, por isso o iPhone da
> mãe do Pedro ficou num ciclo infinito de refresh. O network-first já garante
> que todos os dispositivos apanham a versão nova; não é preciso recarregar.

O `sw.js` também trata as notificações push (check-ins) e o clique na
notificação.

## Módulo SIBA (registo de viajantes)

Quando os hóspedes não fazem check-in online, o Pedro tem de transcrever à mão
os dados dos documentos para o boletim de alojamento da Talkguest. O módulo
SIBA encurta isso: fotografa o documento, a API extrai os campos, e a app
oferece botões de **copiar** por campo.

Funções em `index.html`: `abrirSiba`, `renderSiba`, `sibaCopiar`,
`sibaFotografar`, `sibaReduzirImagem`, `sibaGuardar`, `sibaApagarTudo`.

Notas:

- A imagem é reduzida no browser antes de subir (`sibaReduzirImagem`), para a
  extração ser mais rápida e barata.
- A submissão do boletim é **sempre manual**, feita pelo Pedro. A app só
  prepara os dados.
- Os dados ficam no D1 no máximo 5 dias e depois são apagados
  automaticamente — ver `docs/arquitetura.md`.
- Tipos de documento aceites pela Talkguest: cartão de cidadão, passaporte,
  outro.

## Importação manual do Excel

O botão "Importar" aceita o mesmo Excel que o bot descarrega, lido com
SheetJS. O mapeamento de colunas é um espelho do que está em
`processar_excel()` no `talkguest_import.py` — **mudar um implica mudar o
outro**.

A barra de proveniência (`renderImportInfo()`) mostra se a última importação
foi automática ou manual, e quando.
