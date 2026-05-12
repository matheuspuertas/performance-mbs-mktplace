# Performance MBS MKTPLACE — Atualizador

Automatiza a atualização mensal da planilha de performance de canais de venda da **MBS Pro Grooming**.

## O que o programa faz

Lê os relatórios exportados do Mercado Livre e Shopee, usa **IA (Claude Vision)** para extrair automaticamente os dados das capturas de tela, e atualiza a planilha `Performance MBS MKTPLACE.xlsx` no Google Drive com uma nova linha para o mês.

## Canais suportados

| Canal | Aba | Fontes |
|---|---|---|
| Mercado Livre | `Mercado Livre` | Relatório evolução (Excel) + 2 capturas de tela (publicidade e afiliados) |
| Shopee | `Shopee` | Sales overview (Excel) + CSV de anúncios + 1 captura de tela (afiliados) |
| Site | `Site` | Entrada manual (a implementar) |

## Pré-requisitos

- Python 3.12+
- Google Drive mapeado em `P:\`
- Chave da API Anthropic (armazenada em `config.env`)

```bash
pip install openpyxl pandas anthropic
```

## Configuração inicial

Crie o arquivo `config.env` na pasta do projeto:

```
ANTHROPIC_API_KEY=sk-ant-...
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
| Imagem com gráfico azul | ML | Publicidade (print da tela) |
| Imagem "Métricas / Rendimento geral" | ML | Afiliados (print da tela) |
| Imagem "Principais Indicadores" | Shopee | Afiliados (print da tela) |

### 2. Execute o programa

Duplo-clique em **`rodar.bat`** ou rode no terminal:

```bash
python update_performance.py
```

### 3. Fluxo de execução

```
[OK] Mes detectado: mai/26  (05/2026)
     Correto? [S/n]: S

>> Lendo relatorio ML evolucao...     <- automatico
>> Lendo Shopee overview...           <- automatico
>> Lendo CSV anuncios Shopee...       <- automatico

IMAGENS (3 encontradas) -- leitura automatica com IA

  Analisando: WhatsApp Image ...g.jpeg ...
  [OK] Tipo identificado: ML -- Publicidade
  Valores extraidos:
    Impressoes:         11.332.096
    Cliques:            22.945
    Investimento (R$):  22.314,47
    Receita (R$):       371.017,00
  Confirmar? [S/n/e(ditar)]: S        <- so confirmar, sem digitar nada

  Analisando: WhatsApp Image ...h.jpeg ...
  [OK] Tipo identificado: ML -- Afiliados
  Confirmar? [S/n/e(ditar)]: S

  Analisando: WhatsApp Image ...jpeg ...
  [OK] Tipo identificado: Shopee -- Afiliados
  Confirmar? [S/n/e(ditar)]: S

Confirmar gravacao? [S/n]: S
[OK] Planilha salva!
```

> Se um valor estiver errado, digite **e** para editar campo a campo antes de confirmar.

## Estrutura da planilha

### Aba Mercado Livre (31 colunas, dados a partir da linha 3)

| Grupo | Colunas | Fonte |
|---|---|---|
| Atividade dos compradores | C2–C7 | Relatório evolução (agregado mensal) |
| Desempenho de vendas | C8–C13 | Relatório evolução (agregado mensal) |
| Publicidade | C14–C23 | IA extrai da captura de tela |
| Afiliados | C24–C31 | IA extrai da captura de tela |

### Aba Shopee (38 colunas, dados a partir da linha 3)

| Grupo | Colunas | Fonte |
|---|---|---|
| Dados gerais da loja | C2–C10 | Sales overview (linha de resumo mensal) |
| Publicidade | C11–C28 | CSV de anúncios (agregado de todos os grupos) |
| Afiliados | C29–C38 | IA extrai da captura de tela |

> Campos com fórmula são copiados automaticamente da linha acima (ACOS, ROAS, TACOS, % Fat. Total, Receita Orgânica).

## Arquivos do projeto

```
projeto 1/
├── update_performance.py   # Script principal (uso mensal via rodar.bat)
├── rodar.bat               # Atalho para execução
├── atualizar_abril26.py    # Script pontual — atualização de abr/26
├── config.env              # Chave da API Anthropic (nao commitado)
└── README.md               # Este arquivo
```

## Localização dos arquivos no Drive

```
P:\Meu Drive\Empresas\MBS Pro Grooming\
├── Performance MBS MKTPLACE.xlsx   <- planilha principal
└── materiais\                      <- coloque aqui os arquivos do mes
```

## Observações

- O programa gera um **backup automático** antes de salvar: `Performance MBS MKTPLACE.bkp_YYYYMM.xlsx`
- Se a planilha estiver aberta no Excel durante a execução, o programa salva um arquivo temporário e exibe instruções para substituição manual
- A IA identifica automaticamente o tipo de cada imagem e extrai os valores — sem necessidade de digitar nada
- Em caso de falha da IA, o programa cai em modo manual como fallback
- O arquivo `config.env` com a chave da API **não é commitado** no git (está no `.gitignore`)
