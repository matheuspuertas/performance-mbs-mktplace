"""
Atualiza a planilha Performance MBS com os dados de Abril/2026.
Executa sem interacao - todos os valores ja foram extraidos.

Valores marcados com ~ sao aproximados (imagens com numeros arredondados).
Verifique esses campos na planilha apos a execucao.
"""
import sys
sys.path.insert(0, r"C:\Users\mathe\Downloads\projeto 1")

from update_performance import (
    PERFORMANCE_FILE,
    MATERIAIS_DIR,
    ML_COLS, ML_FORMULA_COLS, ML_LINHA_DADOS,
    SHOPEE_COLS, SHOPEE_FORMULA_COLS, SHOPEE_LINHA_DADOS,
    encontrar_arquivos,
    processar_ml_evolucao,
    processar_shopee_overview,
    processar_shopee_ads_csv,
    atualizar_aba,
)
import shutil
import openpyxl
from pathlib import Path

ANO = 2026
MES = 4

# -------------------------------------------------------------------------------
# DADOS EXTRAIDOS DAS IMAGENS
# -------------------------------------------------------------------------------

# Imagem: 09.16.49g.jpeg  (ML Publicidade - grafico azul, 1 abr - 30 abr 2026)
ML_ADS = {
    "pub_impressoes":          11_332_096,
    "pub_cliques":             22_945,
    "pub_vendas_product_ads":  2_055,
    "pub_vendas_sem_products": 598,
    "pub_investimento":        22_314.47,
    "pub_receita":             371_017.00,
}

# Imagem: 09.16.49h.jpeg  (ML Afiliados - Metricas > Rendimento geral)
ML_AFILIADOS = {
    "af_receita":    62_781.68,
    "af_unidades":   359,
    "af_qtd_vendas": 341,
    "af_custo":      1_717.86,
}

# Imagem: 09.16.49.jpeg  (Shopee Afiliados - Principais Indicadores 04/2026)
# ATENCAO: valores com ~ sao arredondados na imagem - confira na planilha
SHOPEE_AFILIADOS = {
    "af_vendas":             70_000,    # ~ imagem mostra "R$70mil"
    "af_itens_brutos":       875,
    "af_pedidos":            736,
    "af_cliques":            14_800,    # ~ imagem mostra "14,8mil"
    "af_comissao":           3_600,     # ~ imagem mostra "R$3,6mil"
    "af_roi":                19.7,
    "af_compradores_totais": 686,
    "af_novos_compradores":  549,
}

# -------------------------------------------------------------------------------
# EXECUCAO
# -------------------------------------------------------------------------------

print("=" * 62)
print("   ATUALIZACAO ABRIL/2026 -- PERFORMANCE MBS MKTPLACE")
print("=" * 62)

arqs = encontrar_arquivos(MATERIAIS_DIR)

print("\n[1/3] Lendo relatorio ML evolucao...")
dados_ml = processar_ml_evolucao(arqs["ml_evolucao"])
dados_ml.update(ML_ADS)
dados_ml.update(ML_AFILIADOS)

print("\n[2/3] Lendo Shopee overview...")
dados_shopee = processar_shopee_overview(arqs["shopee_overview"])
dados_shopee.update(processar_shopee_ads_csv(arqs["shopee_ads_csv"]))
dados_shopee.update(SHOPEE_AFILIADOS)

print("\n[3/3] Processamento concluido.")

print("\n" + "=" * 62)
print("  RESUMO DO QUE SERA GRAVADO")
print("=" * 62)

print("\nMercado Livre:")
print(f"  Visitas:            {dados_ml['visitas']:,.0f}")
print(f"  Compradores unicos: {dados_ml['compradores_unicos']:,.0f}")
print(f"  Vendas brutas:      R${dados_ml['vendas_brutas']:,.2f}")
print(f"  Impressoes ads:     {dados_ml['pub_impressoes']:,.0f}")
print(f"  Investimento ads:   R${dados_ml['pub_investimento']:,.2f}")
print(f"  Receita ads:        R${dados_ml['pub_receita']:,.2f}")
print(f"  Rec. afiliados:     R${dados_ml['af_receita']:,.2f}")

print("\nShopee:")
print(f"  Visitantes:         {dados_shopee['visitantes']:,.0f}")
print(f"  Vendas pagos:       R${dados_shopee['pedidos_pagos_valor']:,.2f}")
print(f"  GMV ads:            R${dados_shopee['pub_gmv']:,.2f}")
print(f"  Despesas ads:       R${dados_shopee['pub_despesas']:,.2f}")
print(f"  Vendas afiliados:   R${dados_shopee['af_vendas']:,.0f}  [~ APROXIMADO]")
print(f"  Comissao afiliados: R${dados_shopee['af_comissao']:,.0f}  [~ APROXIMADO]")
print(f"  Cliques afiliados:  {dados_shopee['af_cliques']:,.0f}  [~ APROXIMADO]")

print("\nSite: sem dados (arquivos nao encontrados)")

resp = input("\nConfirmar gravacao? [S/n]: ").strip().lower()
if resp == "n":
    print("Cancelado.")
    sys.exit(0)

# Backup
bkp = Path(PERFORMANCE_FILE).with_suffix(f".bkp_{ANO}{MES:02d}.xlsx")
if not bkp.exists():
    shutil.copy2(PERFORMANCE_FILE, bkp)
    print(f"\n[OK] Backup salvo: {bkp.name}")

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

print()
print("ATENCAO: confira os campos aproximados na planilha (aba Shopee, mes abr/26):")
print("  - Shopee Afiliados > Vendas     (imagem mostrava 'R$70mil')")
print("  - Shopee Afiliados > Cliques    (imagem mostrava '14,8mil')")
print("  - Shopee Afiliados > Comissao   (imagem mostrava 'R$3,6mil')")
