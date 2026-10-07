"""
Atualiza a aba Meta Marketplace da planilha Performance MBS com os dados de Julho/2026.
Executa sem interacao.

Diferente das abas Mercado Livre/Shopee/Site, esta aba tem uma linha por
ANUNCIO (nao uma linha agregada por mes), dividida em duas tabelas:
  - Tabela 1 (cabecalho linha 1, dados a partir da linha 2): anuncios do tipo
    "Compras no site" -> coluna J = "Quantidade de vendas"
  - Tabela 2 (cabecalho linha 12, dados a partir da linha 13): anuncios do
    tipo "Conversas por mensagem iniciadas" -> coluna J = "Quantidade de Conversas"

Fontes:
  - Relatorio-sem-titulo-jul-1-2026-a-jul-31-2026.xlsx (Meta Ads "Raw Data Report")
    -> Valor Gasto, Impressoes, Resultados (vendas/conversas)
  - Captura de tela 2026-08-04 224738.png -> Visitas ao perfil do Instagram
  - Captura de tela 2026-08-04 230810.png -> Cliques no link, CPC

Confirmado com o usuario: "Inicio/Fim da veiculacao" usa o periodo do
relatorio (01/07/2026 - 31/07/2026), ja que o export nao traz a data real
de inicio/fim de cada anuncio.
"""
import sys
sys.path.insert(0, r"P:\Meu Drive\Empresas\MBS Pro Grooming\Desenvolvimento de relatórios\projeto 1")

import shutil
from datetime import date
from pathlib import Path

import openpyxl

from update_performance import PERFORMANCE_FILE, copiar_estilo

ANO = 2026
MES = 7
INICIO = date(2026, 7, 1)
FIM    = date(2026, 7, 31)

# Tabela 1 -- "Compras no site" (vai para as linhas 6-9, logo apos a ultima
# linha existente na linha 5)
COMPRAS_NO_SITE = [
    {"produto": "Tapete - Meli",        "valor_gasto": 500.23, "cpc": 1.79, "cliques": 279, "impressoes": 44_361, "visitas_ig": 225, "qtd_vendas": 6},
    {"produto": "AEOLIAN",              "valor_gasto": 499.21, "cpc": 1.10, "cliques": 453, "impressoes": 37_426, "visitas_ig": 273, "qtd_vendas": 16},
    {"produto": "KIT HYDRA - Mateus",   "valor_gasto": 497.32, "cpc": 1.63, "cliques": 306, "impressoes": 25_908, "visitas_ig": 200, "qtd_vendas": 16},
    {"produto": "HYDRA",                "valor_gasto": 232.82, "cpc": 0.63, "cliques": 372, "impressoes": 20_467, "visitas_ig": 206, "qtd_vendas": 8},
]

# Tabela 2 -- "Conversas por mensagem iniciadas" (vai para a linha 14, logo
# apos a ultima linha existente na linha 13)
CONVERSAS = [
    {"produto": "Post Insta Gi - Evento", "valor_gasto": 27.41, "custo_conversa": 27.41 / 15, "cliques": 26, "impressoes": 1_091, "visitas_ig": 4, "qtd_conversas": 15},
]

print("=" * 62)
print("   ATUALIZACAO JULHO/2026 -- ABA META MARKETPLACE")
print("=" * 62)

wb = openpyxl.load_workbook(PERFORMANCE_FILE)
ws = wb["Meta Marketplace"]

print("\nCompras no site:")
linha_ref = 5
linha = 6
for item in COMPRAS_NO_SITE:
    for c in range(1, 11):
        copiar_estilo(ws.cell(linha_ref, c), ws.cell(linha, c))
    ws.cell(linha, 1).value  = "jul/26"
    ws.cell(linha, 2).value  = item["produto"]
    ws.cell(linha, 3).value  = INICIO
    ws.cell(linha, 4).value  = FIM
    ws.cell(linha, 5).value  = item["valor_gasto"]
    ws.cell(linha, 6).value  = item["cpc"]
    ws.cell(linha, 7).value  = item["cliques"]
    ws.cell(linha, 8).value  = item["impressoes"]
    ws.cell(linha, 9).value  = item["visitas_ig"]
    ws.cell(linha, 10).value = item["qtd_vendas"]
    print(f"  linha {linha}: {item['produto']:<20} Valor gasto R${item['valor_gasto']:.2f} | "
          f"CPC R${item['cpc']:.2f} | Cliques {item['cliques']} | Impr {item['impressoes']:,} | "
          f"Visitas IG {item['visitas_ig']} | Vendas {item['qtd_vendas']}")
    linha += 1

print("\nConversas por mensagem iniciadas:")
linha_ref = 13
linha = 14
for item in CONVERSAS:
    for c in range(1, 11):
        copiar_estilo(ws.cell(linha_ref, c), ws.cell(linha, c))
    ws.cell(linha, 1).value  = "jul/26"
    ws.cell(linha, 2).value  = item["produto"]
    ws.cell(linha, 3).value  = INICIO
    ws.cell(linha, 4).value  = FIM
    ws.cell(linha, 5).value  = item["valor_gasto"]
    ws.cell(linha, 6).value  = item["custo_conversa"]
    ws.cell(linha, 7).value  = item["cliques"]
    ws.cell(linha, 8).value  = item["impressoes"]
    ws.cell(linha, 9).value  = item["visitas_ig"]
    ws.cell(linha, 10).value = item["qtd_conversas"]
    print(f"  linha {linha}: {item['produto']:<20} Valor gasto R${item['valor_gasto']:.2f} | "
          f"Custo/conversa R${item['custo_conversa']:.2f} | Cliques {item['cliques']} | "
          f"Impr {item['impressoes']:,} | Visitas IG {item['visitas_ig']} | Conversas {item['qtd_conversas']}")
    linha += 1

resp = input("\nConfirmar gravacao? [S/n]: ").strip().lower()
if resp == "n":
    print("Cancelado.")
    sys.exit(0)

bkp = Path(PERFORMANCE_FILE).with_suffix(f".bkp_{ANO}{MES:02d}_meta.xlsx")
if not bkp.exists():
    shutil.copy2(PERFORMANCE_FILE, bkp)
    print(f"\n[OK] Backup salvo: {bkp.name}")
else:
    print(f"\n[OK] Backup ja existe: {bkp.name}")

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
