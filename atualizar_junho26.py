"""
Atualiza a planilha Performance MBS com os dados de Junho/2026.
Executa sem interacao - todos os valores foram extraidos e verificados.

ATENCAO: af_qtd_vendas (ML Afiliados) apareceu como 0 na screenshot
         (aba "Metricas de afiliados" - pode ser artefato da UI).
         Verifique na planilha apos a execucao se o valor esta correto.
"""
import sys
sys.path.insert(0, r"P:\Meu Drive\Empresas\MBS Pro Grooming\Desenvolvimento de relatórios\projeto 1")

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
MES = 6

# -------------------------------------------------------------------------------
# DADOS DAS IMAGENS (verificados manualmente)
# -------------------------------------------------------------------------------

# Imagens: Captura de tela 2026-07-01 163413.png + 163500.png  (ML Publicidade)
ML_ADS = {
    "pub_impressoes":          7_717_246,
    "pub_cliques":             21_088,
    "pub_vendas_product_ads":  2_059,
    "pub_vendas_sem_products": 1_064,
    "pub_investimento":        18_546.00,
    "pub_receita":             343_509.00,
}

# Imagem: Captura de tela 2026-07-01 163812.png  (ML Afiliados - Metricas de afiliados)
# ATENCAO: af_qtd_vendas mostrou 0 na aba "Metricas de afiliados" - verifique na planilha
ML_AFILIADOS = {
    "af_receita":    62_766.55,
    "af_unidades":   437,
    "af_qtd_vendas": 0,
    "af_custo":      1_665.95,
}

# Imagem: Captura de tela 2026-07-01 164444.png  (Shopee Afiliados - Principais Indicadores)
SHOPEE_AFILIADOS = {
    "af_vendas":             81_200.00,
    "af_itens_brutos":       721,
    "af_pedidos":            568,
    "af_cliques":            11_400,
    "af_comissao":           4_400.00,
    "af_roi":                18.3,
    "af_compradores_totais": 550,
    "af_novos_compradores":  378,
}

# -------------------------------------------------------------------------------
# EXECUCAO
# -------------------------------------------------------------------------------

print("=" * 62)
print("   ATUALIZACAO JUNHO/2026 -- PERFORMANCE MBS MKTPLACE")
print("=" * 62)

arqs = encontrar_arquivos(MATERIAIS_DIR)

print("\n[1/3] Lendo relatorio ML evolucao...")
dados_ml = processar_ml_evolucao(arqs["ml_evolucao"])
dados_ml.update(ML_ADS)
dados_ml.update(ML_AFILIADOS)

print("\n[2/3] Lendo Shopee overview + CSV...")
dados_shopee = processar_shopee_overview(arqs["shopee_overview"])
dados_shopee.update(processar_shopee_ads_csv(arqs["shopee_ads_csv"]))
dados_shopee.update(SHOPEE_AFILIADOS)

print("\n[3/3] Processamento concluido.")

print("\n" + "=" * 62)
print("  RESUMO DO QUE SERA GRAVADO -- jun/26")
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
print(f"  Qtd vendas afil.:    {dados_ml['af_qtd_vendas']}  <-- VERIFICAR (0 na screenshot)")
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

print()
print("ATENCAO: verificar na planilha:")
print("  - ML Afiliados > Quantidade de vendas  (apareceu 0 na screenshot - pode estar errado)")
