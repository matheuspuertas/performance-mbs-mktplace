# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## O que é

Atualizador mensal de `P:\Meu Drive\Empresas\MBS Pro Grooming\Performance MBS MKTPLACE.xlsx` (performance dos canais de venda da MBS Pro Grooming). Todo mês ele adiciona uma linha em cada aba de canal a partir dos relatórios exportados do Mercado Livre, Shopee, Tray (site) e Meta Ads.

## Comandos

```
pip install openpyxl pandas anthropic
python update_performance.py        # execução mensal interativa (igual ao rodar.bat)
python atualizar_<mes>26.py         # script pontual de um mês específico
```

Requer Python 3.12+ e o Google Drive mapeado em `P:\`. Não há testes, linter nem etapa de build. Todo script grava na planilha real do Drive — não existe modo de simulação; os scripts imprimem um resumo do que será gravado antes de salvar.

## Organização

- Esta pasta (`projeto 1\`) é o repositório git e contém apenas código.
- A **pasta pai** (`Desenvolvimento de relatórios\`) é o `MATERIAIS_DIR`: é onde o usuário coloca os exports e as capturas de tela do mês. O README ainda fala de uma subpasta `materiais\` e de um fallback de arquivo bloqueado no script principal — as duas coisas estão desatualizadas; confie no código.
- O `MATERIAIS_DIR` deve conter arquivos de um mês só por vez: `encontrar_arquivos()` fica com a última ocorrência de cada padrão e trata todo `.jpg`/`.png` como captura de tela a processar; `detectar_mes()` usa o ano/mês mais frequente nos nomes dos arquivos.

| Padrão de nome | Origem |
|---|---|
| `Relatorio_evolucao_negocio_YYYY_MM_DD-YYYY_MM_DD.xlsx` | Mercado Livre, dados gerais |
| `sales_overview_YYYYMMDD-YYYYMMDD.xlsx` | Shopee, dados gerais |
| `Dados+Gerais+de+Anúncios+Shopee-DD_MM_YYYY-DD_MM_YYYY.csv` | Shopee, anúncios |
| `Relatório-sem-título-<período>.xlsx/.csv` | Export do Meta Ads (abas Site / Meta Marketplace) |
| `pedidos_*tray.csv` | Pedidos do site na Tray (latin1, números em formato BR) |
| capturas de tela | Publicidade ML, afiliados ML, afiliados Shopee, métricas Instagram/Meta |

### De onde o usuário tira as métricas do Mercado Livre

Painel do vendedor (exige login na conta dele):

- Publicidade (campos `pub_*`) — https://vendedores.mercadolivre.com.br/publicidade/product-ads/admin/campaigns?advertiser_id=174249&account_id=184913&navigate_to=seller_central&from=ads-manager&status=A%2CD
- Métricas de negócio ("Resumo de desempenho" e download do `Relatorio_evolucao_negocio_*.xlsx`) — https://vendedores.mercadolivre.com.br/metricas
- Afiliados (campos `af_*`) — https://vendedores.mercadolivre.com.br/seller-affiliates/dashboard#hub

### De onde o usuário tira os dados da Shopee

Seller Centre (exige login na conta dele):

- Afiliados (campos `af_*`, captura "Principais Indicadores") — https://seller.shopee.com.br/portal/web-seller-affiliate/dashboard
- Anúncios (campos `pub_*`, export `Dados+Gerais+de+Anúncios+Shopee-*.csv`) — https://seller.shopee.com.br/portal/marketing/pas/index?source_page_id=1&from=1788663600&to=1791341999&type=new_cpc_homepage&group=last_month (`from`/`to` são timestamps Unix do período; os do link são 06/09 a 06/10/2026. No painel, "Último mês" = últimos 30 dias, não o mês fechado: o CSV precisa ser exportado com período personalizado de dia 01 ao último dia do mês — confira as datas no nome do arquivo antes de usar)
- Visão geral da loja (dados gerais, export `sales_overview_*.xlsx`) — https://seller.shopee.com.br/datacenter/overview

### De onde o usuário tira os dados do Site

- Usuários ("Total de usuários") — Looker Studio, acessado logado com a conta Google matheusvpuertas@gmail.com: https://datastudio.google.com/u/0/reporting/d81525ac-8c4e-4962-9b17-5780c9addc51/page/kTbyD
- Pedidos e valor total — CSV de pedidos da Tray (`pedidos_1277450_*.csv`).

## Arquitetura

`update_performance.py` é ao mesmo tempo o ponto de entrada interativo e a biblioteca que os scripts pontuais importam.

1. **Extração de arquivos** — `processar_ml_evolucao()`, `processar_shopee_overview()` e `processar_shopee_ads_csv()` devolvem dicionários cujas chaves são os nomes de campo de `ML_COLS` / `SHOPEE_COLS`. Os índices (conversão, ROAS, ACOS, ticket médio…) são calculados aqui a partir dos totais somados, não lidos dos relatórios.
2. **Extração de capturas de tela** — `processar_imagens_ia()` envia cada imagem ao Claude (`claude-haiku-4-5-20251001`) com o `PROMPT_VISAO`, que devolve um JSON com `tipo` (`ml_ads`, `ml_af`, `shopee_af`) e os valores; `CAMPOS_IA` mapeia essas chaves para os nomes de campo da planilha. O usuário confirma ou edita cada resultado; `_fallback_manual()` cobre as falhas. A chave da API vem do `config.env` (fora do git) ou da variável de ambiente.
3. **Gravação** — `atualizar_aba(ws, col_map, formula_cols, dados, ano, mes, linha_inicio)` localiza a linha do mês (a coluna A guarda o serial Excel do dia 1 do mês) ou acrescenta uma após a última, copia os estilos da linha de cima, copia as fórmulas das colunas `*_FORMULA_COLS` via `ajustar_formula()` (desloca as referências relativas de linha) e então grava os valores. Pode ser executada de novo para o mesmo mês: a linha existente é atualizada. Campos ausentes de `dados` (ou `None`) não são tocados, então atualizações parciais funcionam.

Os dicionários `*_COLS` são a fonte única da estrutura da planilha (índices de coluna começando em 1); as colunas de `*_FORMULA_COLS` nunca são sobrescritas com dados. Primeira linha de dados: ML 3, Shopee 3, Site 4.

## Scripts pontuais mensais (`atualizar_<mes>26.py`)

Na prática, todo mês desde abr/26 foi feito com um script dedicado em vez do fluxo interativo, porque as fontes mudam com frequência. Para um mês novo, copie o mais recente e ajuste. O padrão:

- `sys.path.insert` + import das constantes/funções de `update_performance`; definir `ANO` / `MES`.
- Os valores das capturas de tela são lidos e fixados em dicionários (`ML_ADS`, `ML_AFILIADOS`, `SHOPEE_AFILIADOS`), com um comentário indicando a imagem de origem de cada um.
- Os dados vindos de arquivo continuam passando pelas funções `processar_*` compartilhadas; tudo é mesclado e enviado a `atualizar_aba()`.
- Backup em `Performance MBS MKTPLACE.bkp_YYYYMM[_sufixo].xlsx`, só se ainda não existir.
- `wb.save()` dentro de `try/except PermissionError`, com fallback para `~temp_<nome>.xlsx` e instruções de substituição manual (arquivo aberto no Excel ou Drive sincronizando). O script principal não tem esse fallback.
- As decisões confirmadas com o usuário ficam registradas na docstring do script — mantenha essa prática.
- **Gráficos**: logo após o `wb.save()`, chamar `restaurar_graficos(bkp, destino, ANO, MES)` de `graficos.py` (ver abaixo). Atualizar o mês inclui atualizar os gráficos — o usuário espera isso.

### Apresentação mensal (`apresentacao/build.js`)

Todo mês, depois de fechar Mercado Livre e Shopee na planilha, gerar a apresentação "Performance Marketplaces — <Mês> de <Ano>" — o usuário quer isso sempre, sem precisar pedir. É feita com pptxgenjs (`NODE_PATH` apontando para um `node_modules` com pptxgenjs) na identidade da Puertas Marketing (Montserrat/Manrope, azul-marinho 0A2B5C, azuis 0F5BD9/1E7CFF/4C95FF, fundo claro F4F3EF), com o logo da Puertas e o da MBS (`logo_mbs.png`, em cartão branco na capa). Copiar `build.js`, trocar os números pelos da planilha e escrever os destaques e pontos de atenção do mês. Conferir renderizando cada slide pelo PowerPoint via COM (`Slide.Export`).

Entregáveis do mês vão para `P:\Meu Drive\Empresas\MBS Pro Grooming\Relatórios\<Ano>\<MM> - <Mês>\` (ex.: `2026\09 - Setembro\`): a apresentação e uma cópia da planilha como estava ao fechar o mês (`Performance MBS MKTPLACE - <Mês> <Ano>.xlsx`). A planilha principal **não** sai de `MBS Pro Grooming\` — é ela que os scripts atualizam (`PERFORMANCE_FILE`).

### Gráficos (`graficos.py`)

As abas `Mercado Livre` e `Shopee` têm um gráfico de colunas empilhadas "Receita por canal | Mês" (publicidade + afiliados + orgânica) cujos rótulos são **texto fixo por ponto** (`R$ 401.052,00 (54%)`), mais uma série "Total" de zeros que só carrega o rótulo do total. Todo `wb.save()` do openpyxl regrava o XML dos gráficos e transforma esses rótulos em "None". `restaurar_graficos()` pega o XML original do backup (que precisa ter sido salvo pelo Excel), acrescenta o ponto do mês novo (intervalos, cache e rótulos calculados) e injeta de volta no `.xlsx`. Só funciona se o mês novo for a linha imediatamente seguinte ao fim do intervalo do gráfico. Para conferir o resultado, abra a planilha via COM do Excel (PowerShell, somente leitura) e leia `Series.Points(n).DataLabel.Text`.

Mudanças nas fontes que o script principal **não** trata (os scripts de jul/ago tratam):

- **Relatório de evolução do ML** muda de formato de um mês para o outro e sempre vaza dias do mês seguinte; `processar_ml_evolucao()` somaria tudo. Abra o arquivo antes de escolher a variante: em jul/ago veio com uma linha por nome de mês (ex.: `Agosto`) → `processar_ml_evolucao_mes(caminho, "Agosto")`; em set veio com uma linha por dia (`dd/mm/aa`) mais o dia 01 do mês seguinte → `processar_ml_evolucao_dias(caminho, ano, mes)` em `atualizar_setembro26.py`.
- **Linhas de mês pré-criadas**: as abas já trazem a coluna A preenchida com datas até dez/26, então `atualizar_aba()` cai sempre em "linha existente" e grava o serial numérico por cima da data; `atualizar_setembro26.py` restaura o valor original da célula depois.
- **Painel de publicidade do ML** agora é o "Painel ao vivo": `Vendas atribuídas` → `pub_vendas_product_ads`, `Outras vendas` → `pub_vendas_sem_products`. O `PROMPT_VISAO` ainda descreve os rótulos antigos.

Abas fora do fluxo principal:

- **Site** (`atualizar_julho26_site.py`) — `num_pedidos` / `valor_total` vêm do CSV da Tray, `meta_ads` = soma de `Valor usado (BRL)` na aba `Raw Data Report` do export do Meta, `google` e `usuarios` são informados pelo usuário; `ticket_medio`, `valor_investido`, `taxa_tacos` e `taxa_conversao` são derivados. Costuma ser feita numa segunda etapa, depois de ML/Shopee.
- **Meta Marketplace** (`atualizar_julho26_meta_marketplace.py`) — formato diferente: uma linha por *anúncio*, não por mês, em duas tabelas empilhadas ("Compras no site" e "Conversas por mensagem iniciadas"), com o mês gravado como texto (`jul/26`). Não usa `atualizar_aba()`; as posições de linha estão fixas no código, então confira as últimas linhas atuais da aba antes de gravar um mês novo.

## Convenções

- **Padrão visual das abas mensais** (Mercado Livre, Shopee, Site) — o usuário quer sempre assim: coluna A do ano passado em cinza (RGB BFBFBF) e do ano atual em azul claro (RGB 8CB5F9), borda inferior média na linha de dezembro do ano passado, e todos os meses até dezembro do ano atual já criados (data + formatação; no Site, com as fórmulas protegidas por `IFERROR`). Formatação se faz pelo Excel via COM e salva pelo Excel, nunca pelo openpyxl.

- Identificadores, comentários e saída no console em português; as strings de `print` são ASCII puro (sem acentos) por causa do console do Windows.
- Percentuais são gravados como decimais de 0 a 1; valores monetários como float simples. `limpar_numero()` interpreta formatos BR (`1.234,56`, `70mil`, `R$`).
