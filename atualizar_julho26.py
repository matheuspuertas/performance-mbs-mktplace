"""
Atualiza a planilha Performance MBS com os dados de Julho/2026.
Executa sem interacao - todos os valores foram extraidos e verificados.

ATENCAO: o painel de Publicidade do ML mudou de layout ("Painel ao vivo"),
         com os rotulos "Vendas atribuidas" / "Outras vendas" no lugar de
         "Vendas por Product Ads" / "Vendas sem Products". Mapeamento usado
         (confirmado com o usuario): Vendas atribuidas -> pub_vendas_product_ads,
         Outras vendas -> pub_vendas_sem_products.

Imagens ignoradas (confirmado com o usuario):
  - Captura de tela 2026-08-04 173513.png -> tela de config da Hotmart, sem relacao.
  - Captura de tela 2026-08-04 210337.png -> painel de Anuncios do Shopee Ads
    Manager (redundante com o CSV, que ja e processado automaticamente).

ATENCAO: o relatorio "Relatorio_evolucao_negocio_2026_07_01-2026_07_31.xlsx"
         nao vem mais agregado por dia -- vem com uma linha "Julho" e uma
         linha "Agosto" (vazamento de poucos dias do mes seguinte). A funcao
         processar_ml_evolucao() padrao somaria as duas linhas, misturando
         dados de agosto no relatorio de julho. Por isso este script usa
         processar_ml_evolucao_mes() abaixo, que le apenas a linha "Julho"
         (confirmado com o usuario).
"""
import sys
sys.path.insert(0, r"P:\Meu Drive\Empresas\MBS Pro Grooming\Desenvolvimento de relatórios\projeto 1")

import openpyxl
import shutil
from pathlib import Path

from update_performance import (
    PERFORMANCE_FILE,
    MATERIAIS_DIR,
    ML_COLS, ML_FORMULA_COLS, ML_LINHA_DADOS,
    SHOPEE_COLS, SHOPEE_FORMULA_COLS, SHOPEE_LINHA_DADOS,
    encontrar_arquivos,
    processar_shopee_overview,
    processar_shopee_ads_csv,
    atualizar_aba,
    limpar_numero,
)

ANO = 2026
MES = 7


def processar_ml_evolucao_mes(caminho, nome_mes):
    """Le apenas a linha `nome_mes` (ex: 'Julho') do relatorio de evolucao do ML,
    ignorando outras linhas de mes (ex: vazamento de dias do mes seguinte)."""
    print(f"\n  >> Lendo relatorio ML: {caminho.name}  (somente linha '{nome_mes}')")
    wb = openpyxl.load_workbook(str(caminho), data_only=True)

    aba = None
    for nome in wb.sheetnames:
        if "neg" in nome.lower():
            aba = wb[nome]
            break
    if aba is None:
        aba = wb.active

    header_row = None
    for r in range(1, 20):
        val = aba.cell(r, 1).value
        if val and str(val).strip().lower() == "data":
            header_row = r
            break
    if header_row is None:
        raise ValueError("Cabecalho 'Data' nao encontrado na aba Negocio.")

    col_map = {}
    for c in range(1, aba.max_column + 1):
        val = aba.cell(header_row, c).value
        if val:
            col_map[str(val).strip()] = c

    campos = {
        "Visitas":                "visitas",
        "Compradores únicos":     "compradores_unicos",
        "Novos compradores":      "novos_compradores",
        "Compradores existentes": "compradores_existentes",
        "Quantidade de vendas":   "qtd_vendas",
        "Unidades vendidas":      "unidades_vendidas",
        "Vendas brutas":          "vendas_brutas",
    }

    linha_alvo = None
    for r in range(header_row + 1, aba.max_row + 1):
        data_val = aba.cell(r, 1).value
        if data_val and str(data_val).strip().lower() == nome_mes.lower():
            linha_alvo = r
            break
    if linha_alvo is None:
        raise ValueError(f"Linha '{nome_mes}' nao encontrada no relatorio.")

    totais = {}
    for campo_orig, campo_dest in campos.items():
        col = col_map.get(campo_orig)
        val = aba.cell(linha_alvo, col).value if col else None
        if isinstance(val, str):
            val = limpar_numero(val)
        totais[campo_dest] = float(val) if val is not None else 0.0

    cu, vb, qv, uv, vi, ce = (
        totais["compradores_unicos"], totais["vendas_brutas"], totais["qtd_vendas"],
        totais["unidades_vendidas"], totais["visitas"], totais["compradores_existentes"],
    )
    totais["media_vendas_comprador"] = vb / cu if cu else 0
    totais["taxa_recompra"]          = ce / cu if cu else 0
    totais["conversao"]              = qv / vi if vi else 0
    totais["valor_medio_venda"]      = vb / qv if qv else 0
    totais["preco_medio_unidade"]    = vb / uv if uv else 0

    print(f"  [OK] Linha '{nome_mes}': {vi:,.0f} visitas, R${vb:,.2f} vendas brutas")
    return totais

# -------------------------------------------------------------------------------
# DADOS DAS IMAGENS (verificados manualmente)
# -------------------------------------------------------------------------------

# Imagens: Captura de tela 2026-08-04 204901.png + 204912.png (ML Publicidade - Painel ao vivo)
ML_ADS = {
    "pub_impressoes":          7_481_586,
    "pub_cliques":             23_641,
    "pub_vendas_product_ads":  2_436,     # "Vendas atribuidas"
    "pub_vendas_sem_products": 1_367,     # "Outras vendas"
    "pub_investimento":        23_136.00,
    "pub_receita":             390_922.00,
}

# Imagem: Captura de tela 2026-08-04 205123.png (ML Afiliados - aba "Metricas de produtos")
ML_AFILIADOS = {
    "af_receita":    56_303.51,
    "af_unidades":   414,
    "af_qtd_vendas": 402,
    "af_custo":      1_894.53,
}

# Imagem: Captura de tela 2026-08-04 210404.png (Shopee Afiliados - Principais Indicadores, 07/2026)
SHOPEE_AFILIADOS = {
    "af_vendas":             78_100.00,
    "af_itens_brutos":       748,
    "af_pedidos":            602,
    "af_cliques":            25_000,
    "af_comissao":           4_600.00,
    "af_roi":                16.8,
    "af_compradores_totais": 586,
    "af_novos_compradores":  409,
}

# -------------------------------------------------------------------------------
# EXECUCAO
# -------------------------------------------------------------------------------

print("=" * 62)
print("   ATUALIZACAO JULHO/2026 -- PERFORMANCE MBS MKTPLACE")
print("=" * 62)

arqs = encontrar_arquivos(MATERIAIS_DIR)

print("\n[1/3] Lendo relatorio ML evolucao...")
dados_ml = processar_ml_evolucao_mes(arqs["ml_evolucao"], "Julho")
dados_ml.update(ML_ADS)
dados_ml.update(ML_AFILIADOS)

print("\n[2/3] Lendo Shopee overview + CSV...")
dados_shopee = processar_shopee_overview(arqs["shopee_overview"])
dados_shopee.update(processar_shopee_ads_csv(arqs["shopee_ads_csv"]))
dados_shopee.update(SHOPEE_AFILIADOS)

print("\n[3/3] Processamento concluido.")

print("\n" + "=" * 62)
print("  RESUMO DO QUE SERA GRAVADO -- jul/26")
print("=" * 62)

print("\nMercado Livre:")
print(f"  Visitas:             {dados_ml['visitas']:,.0f}")
print(f"  Compradores unicos:  {dados_ml['compradores_unicos']:,.0f}")
print(f"  Novos compradores:   {dados_ml['novos_compradores']:,.0f}")
print(f"  Unidades vendidas:   {dados_ml['unidades_vendidas']:,.0f}")
print(f"  Vendas brutas:       R${dados_ml['vendas_brutas']:,.2f}")
print(f"  Impressoes ads:      {dados_ml['pub_impressoes']:,.0f}")
print(f"  Cliques ads:         {dados_ml['pub_cliques']:,.0f}")
print(f"  Investimento ads:    R${dados_ml['pub_investimento']:,.2f}")
print(f"  Receita ads:         R${dados_ml['pub_receita']:,.2f}")
print(f"  Rec. afiliados:      R${dados_ml['af_receita']:,.2f}")
print(f"  Unid. afiliados:     {dados_ml['af_unidades']:,.0f}")
print(f"  Qtd vendas afil.:    {dados_ml['af_qtd_vendas']:,.0f}")
print(f"  Custo afiliados:     R${dados_ml['af_custo']:,.2f}")

print("\nShopee:")
print(f"  Visitantes:          {dados_shopee['visitantes']:,.0f}")
print(f"  Vendas pagos:        R${dados_shopee['pedidos_pagos_valor']:,.2f}")
print(f"  GMV ads:             R${dados_shopee['pub_gmv']:,.2f}")
print(f"  Despesas ads:        R${dados_shopee['pub_despesas']:,.2f}")
print(f"  Vendas afiliados:    R${dados_shopee['af_vendas']:,.2f}")
print(f"  Comissao afiliados:  R${dados_shopee['af_comissao']:,.2f}")

print("\nSite: sem dados (adicionar manualmente se necessario)")

resp = input("\nConfirmar gravacao? [S/n]: ").strip().lower()
if resp == "n":
    print("Cancelado.")
    sys.exit(0)

# Backup
bkp = Path(PERFORMANCE_FILE).with_suffix(f".bkp_{ANO}{MES:02d}.xlsx")
if not bkp.exists():
    shutil.copy2(PERFORMANCE_FILE, bkp)
    print(f"\n[OK] Backup salvo: {bkp.name}")
else:
    print(f"\n[OK] Backup ja existe: {bkp.name}")

# Carregar e gravar
wb = openpyxl.load_workbook(PERFORMANCE_FILE)

ws = wb["Mercado Livre"]
ln = atualizar_aba(ws, ML_COLS, ML_FORMULA_COLS, dados_ml, ANO, MES, ML_LINHA_DADOS)
print(f"[OK] Aba Mercado Livre -> linha {ln}")

ws = wb["Shopee"]
ln = atualizar_aba(ws, SHOPEE_COLS, SHOPEE_FORMULA_COLS, dados_shopee, ANO, MES, SHOPEE_LINHA_DADOS)
print(f"[OK] Aba Shopee -> linha {ln}")

# Salvar (com fallback para arquivo temporario caso esteja bloqueado)
destino = Path(PERFORMANCE_FILE)
temp    = destino.parent / ("~temp_" + destino.name)

try:
    wb.save(str(destino))
    print(f"\n[OK] Planilha salva!")
    print(f"     {destino}")
except PermissionError:
    wb.save(str(temp))
    print(f"\n[!] Arquivo bloqueado (aberto no Excel ou sincronizando no Google Drive).")
    print(f"    Arquivo salvo em: {temp.name}")
    print()
    print("    Para finalizar:")
    print("    1. Feche a planilha no Excel (se estiver aberta)")
    print("    2. Aguarde o Google Drive sincronizar")
    print(f"    3. Substitua o arquivo original pelo '{temp.name}'")
