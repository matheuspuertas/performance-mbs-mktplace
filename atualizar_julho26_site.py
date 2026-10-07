"""
Atualiza a aba Site da planilha Performance MBS com os dados de Julho/2026.
Executa sem interacao - complementa atualizar_julho26.py (que ja gravou ML e Shopee).

Fontes:
  - pedidos_..._tray.csv (pasta Desenvolvimento de relatorios) -> num_pedidos, valor_total
    (62 pedidos, canal "LOJA VIRTUAL", todos com status ativo/em andamento -- nenhum cancelado)
  - Relatorio-sem-titulo-jul-1-2026-a-jul-31-2026.xlsx -> meta_ads (soma "Valor usado (BRL)")
  - Informado manualmente pelo usuario: google = R$1.002,27 ; usuarios = 6.939
"""
import sys
sys.path.insert(0, r"P:\Meu Drive\Empresas\MBS Pro Grooming\Desenvolvimento de relatórios\projeto 1")

import shutil
from pathlib import Path

import openpyxl
import pandas as pd

from update_performance import (
    PERFORMANCE_FILE,
    MATERIAIS_DIR,
    SITE_COLS, SITE_FORMULA_COLS, SITE_LINHA_DADOS,
    atualizar_aba,
)

ANO = 2026
MES = 7

TRAY_CSV = Path(MATERIAIS_DIR) / "pedidos_1277450_b5c01b17-4ba6-4b0a-b20a-4ac3180e0dca tray.csv"
META_XLSX = Path(MATERIAIS_DIR) / "Relatório-sem-título-jul-1-2026-a-jul-31-2026.xlsx"

USUARIOS = 6_939
GOOGLE   = 1_002.27

# -------------------------------------------------------------------------------
# Pedidos do site (Tray)
# -------------------------------------------------------------------------------

df = pd.read_csv(str(TRAY_CSV), encoding="latin1", sep=None, engine="python")
df["Total_num"] = (
    df["Total"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)
)
num_pedidos = len(df)
valor_total = df["Total_num"].sum()
ticket_medio = valor_total / num_pedidos if num_pedidos else 0

print(f"[OK] Pedidos (Tray): {num_pedidos} pedidos | Valor total: R${valor_total:,.2f}")

# -------------------------------------------------------------------------------
# Meta Ads
# -------------------------------------------------------------------------------

wb_meta = openpyxl.load_workbook(str(META_XLSX), data_only=True)
ws_meta = wb_meta["Raw Data Report"]
col_map = {}
for c in range(1, ws_meta.max_column + 1):
    v = ws_meta.cell(1, c).value
    if v:
        col_map[str(v).strip()] = c

col_valor = col_map["Valor usado (BRL)"]
meta_ads = 0.0
for r in range(2, ws_meta.max_row + 1):
    v = ws_meta.cell(r, col_valor).value
    if isinstance(v, (int, float)):
        meta_ads += v

print(f"[OK] Meta Ads: R${meta_ads:,.2f}")

# -------------------------------------------------------------------------------
# Consolidar
# -------------------------------------------------------------------------------

valor_investido = GOOGLE + meta_ads
taxa_tacos = valor_investido / valor_total if valor_total else 0
taxa_conversao = num_pedidos / USUARIOS if USUARIOS else 0

dados_site = {
    "usuarios":        USUARIOS,
    "num_pedidos":     num_pedidos,
    "valor_total":     valor_total,
    "ticket_medio":    ticket_medio,
    "valor_investido": valor_investido,
    "google":          GOOGLE,
    "meta_ads":        meta_ads,
    "taxa_tacos":      taxa_tacos,
    "taxa_conversao":  taxa_conversao,
}

print("\n" + "=" * 62)
print("  RESUMO DO QUE SERA GRAVADO -- SITE -- jul/26")
print("=" * 62)
print(f"  Usuarios:            {dados_site['usuarios']:,}")
print(f"  N. Pedidos:          {dados_site['num_pedidos']:,}")
print(f"  Valor Total:         R${dados_site['valor_total']:,.2f}")
print(f"  Ticket Medio:        R${dados_site['ticket_medio']:,.2f}")
print(f"  Valor Investido:     R${dados_site['valor_investido']:,.2f}")
print(f"  Google Ads:          R${dados_site['google']:,.2f}")
print(f"  Meta Ads:            R${dados_site['meta_ads']:,.2f}")
print(f"  Taxa TACOS:          {dados_site['taxa_tacos']:.4f}  ({dados_site['taxa_tacos']*100:.2f}%)")
print(f"  Taxa Conversao:      {dados_site['taxa_conversao']:.4f}  ({dados_site['taxa_conversao']*100:.2f}%)")

resp = input("\nConfirmar gravacao? [S/n]: ").strip().lower()
if resp == "n":
    print("Cancelado.")
    sys.exit(0)

# Backup
bkp = Path(PERFORMANCE_FILE).with_suffix(f".bkp_{ANO}{MES:02d}_site.xlsx")
if not bkp.exists():
    shutil.copy2(PERFORMANCE_FILE, bkp)
    print(f"\n[OK] Backup salvo: {bkp.name}")
else:
    print(f"\n[OK] Backup ja existe: {bkp.name}")

wb = openpyxl.load_workbook(PERFORMANCE_FILE)
ws = wb["Site"]
ln = atualizar_aba(ws, SITE_COLS, SITE_FORMULA_COLS, dados_site, ANO, MES, SITE_LINHA_DADOS)
print(f"[OK] Aba Site -> linha {ln}")

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
