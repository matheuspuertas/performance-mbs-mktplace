# Performance MBS MKTPLACE — Atualizador

Automatiza a atualização mensal da planilha de performance de canais de venda da **MBS Pro Grooming**.

## O que o programa faz

Lê os relatórios exportados do Mercado Livre e Shopee, extrai as métricas do mês e atualiza a planilha `Performance MBS MKTPLACE.xlsx` no Google Drive com uma nova linha para o mês.

## Canais suportados

| Canal | Aba | Fontes |
|---|---|---|
| Mercado Livre | `Mercado Livre` | Relatório evolução (Excel) + 2 capturas de tela (publicidade e afiliados) |
| Shopee | `Shopee` | Sales overview (Excel) + CSV de anúncios + 1 captura de tela (afiliados) |
| Site | `Site` | Entrada manual (a implementar) |

## Pré-requisitos

- Python 3.12+
- Google Drive mapeado em `P:\`

```bash
pip install openpyxl pandas
```

## Como usar (todo mês)

### 1. Coloque os arquivos na pasta `materiais`

```
P:\Meu Drive\Empresas\MBS Pro Grooming\materiais\
```

| Arquivo | Canal | Tipo |
|---|---|---|
| `Relatorio_evolucao_negocio_YYYY_MM_DD-YYYY_MM_DD.xlsx` | ML | Dados gerais diários |
| `sales_overview_YYYYMMDD-YYYYMMDD.xlsx` | Shopee | Dados gerais do mês |
| `Dados+Gerais+de+Anúncios+Shopee-DD_MM_YYYY-DD_MM_YYYY.csv` | Shopee | Dados de anúncios |
| Imagem com gráfico azul | ML | Publicidade (Print da tela) |
| Imagem "Métricas / Rendimento geral" | ML | Afiliados (Print da tela) |
| Imagem "Principais Indicadores" | Shopee | Afiliados (Print da tela) |

### 2. Execute o programa

Duplo-clique em **`rodar.bat`** ou rode no terminal:

```bash
python update_performance.py
```

### 3. Siga o fluxo interativo

```
[OK] Mes detectado: mai/26  (05/2026)
     Correto? [S/n]: S

>> Lendo relatorio ML evolucao...     ← automático
>> Lendo Shopee overview...           ← automático
>> Lendo CSV anuncios Shopee...       ← automático

IMAGENS (3 encontradas) — identificacao necessaria
  [1] ML -- Publicidade (grafico azul)
  [2] ML -- Afiliados (Metricas / Rendimento geral)
  [3] Shopee -- Afiliados (Principais Indicadores)

  → Para cada imagem: o programa abre a foto e pede os números

Confirmar gravacao? [S/n]: S
[OK] Planilha salva!
```

## Estrutura da planilha

### Aba Mercado Livre (31 colunas, dados a partir da linha 3)

| Grupo | Colunas | Fonte |
|---|---|---|
| Atividade dos compradores | C2–C7 | Relatório evolução (agregado mensal) |
| Desempenho de vendas | C8–C13 | Relatório evolução (agregado mensal) |
| Publicidade | C14–C23 | Captura de tela (manual) |
| Afiliados | C24–C31 | Captura de tela (manual) |

### Aba Shopee (38 colunas, dados a partir da linha 3)

| Grupo | Colunas | Fonte |
|---|---|---|
| Dados gerais da loja | C2–C10 | Sales overview (linha de resumo mensal) |
| Publicidade | C11–C28 | CSV de anúncios (agregado de todos os grupos) |
| Afiliados | C29–C38 | Captura de tela (manual) |

> Campos com fórmula são copiados automaticamente da linha acima (ACOS, ROAS, TACOS, % Fat. Total, Receita Orgânica).

## Arquivos do projeto

```
projeto 1/
├── update_performance.py   # Script principal (uso mensal via rodar.bat)
├── rodar.bat               # Atalho para execução
├── atualizar_abril26.py    # Script pontual — atualização de abr/26
└── README.md               # Este arquivo
```

## Localização dos arquivos no Drive

```
P:\Meu Drive\Empresas\MBS Pro Grooming\
├── Performance MBS MKTPLACE.xlsx   ← planilha principal
└── materiais\                      ← coloque aqui os arquivos do mês
```

## Observações

- O programa gera um **backup automático** antes de salvar: `Performance MBS MKTPLACE.bkp_YYYYMM.xlsx`
- Se a planilha estiver aberta no Excel durante a execução, o programa salva um arquivo temporário (`~temp_Performance MBS MKTPLACE.xlsx`) e exibe instruções para substituição manual
- Valores de afiliados da Shopee podem aparecer arredondados nas capturas de tela (ex: `R$70mil`) — verifique e corrija na planilha se necessário
