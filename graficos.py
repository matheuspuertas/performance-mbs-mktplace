"""
Restaura e estende os graficos da planilha Performance MBS depois de um wb.save() do openpyxl.

Por que existe: as abas Mercado Livre e Shopee tem um grafico de colunas empilhadas
("Receita por canal | Mes") com rotulos de dados digitados como texto fixo, ponto a ponto
(ex: "R$ 401.052,00 (54%)"), mais uma serie "Total" de zeros que so serve para carregar o
rotulo do total. O openpyxl regrava o XML dos graficos ao salvar e perde esses rotulos
(viram "None"). Esta rotina pega o XML original dos graficos em um arquivo salvo pelo
Excel (o backup feito antes do save), acrescenta o ponto do mes novo -- intervalos,
valores em cache e rotulos -- e injeta de volta na planilha ja atualizada.

Uso (depois do wb.save do script do mes):
    from graficos import restaurar_graficos
    restaurar_graficos(bkp, PERFORMANCE_FILE, ANO, MES)
"""
import os
import re
import uuid
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string

from update_performance import MESES_PT, encontrar_linha_mes, mes_serial_excel

LINHA_DADOS = 3   # primeira linha de dados nas abas Mercado Livre e Shopee

_RE_REF = re.compile(r"^(?:'([^']+)'|([^!]+))!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$")


def _formato_br(v):
    """1234567.8 -> '1.234.567,80'"""
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _valor(ws, col, linha, _nivel=0):
    """Valor numerico da celula; formulas aritmeticas simples (=J22-S22-X22) sao calculadas."""
    v = ws.cell(linha, column_index_from_string(col)).value
    if isinstance(v, str) and v.startswith("="):
        if _nivel > 5:
            raise ValueError(f"Formula muito aninhada em {col}{linha}")
        expr = re.sub(
            r"([A-Z]+)(\d+)",
            lambda m: repr(float(_valor(ws, m.group(1), int(m.group(2)), _nivel + 1))),
            v[1:],
        )
        if not re.fullmatch(r"[\d.eE+\-*/() ]+", expr):
            raise ValueError(f"Formula nao suportada em {col}{linha}: {v}")
        return eval(expr, {"__builtins__": {}})
    if v is None:
        raise ValueError(f"Celula {ws.title}!{col}{linha} vazia")
    return float(v)


def _acrescentar_ponto(bloco, valor_cache):
    """Soma 1 ao ptCount de um bloco <c:cat>/<c:val> e acrescenta o ponto novo no fim."""
    m = re.search(r'<c:ptCount val="(\d+)"/>', bloco)
    n = int(m.group(1))
    bloco = bloco.replace(m.group(0), f'<c:ptCount val="{n + 1}"/>', 1)
    ponto = f'<c:pt idx="{n}"><c:v>{valor_cache}</c:v></c:pt>'
    fecho = re.search(r"</c:(numCache|strCache|numLit|strLit)>", bloco)
    return bloco[:fecho.start()] + ponto + bloco[fecho.start():]


def _estender_grafico(xml, wb, ano, mes):
    """Devolve o XML do grafico com o mes novo, ou None se o mes ja estiver no grafico."""
    series = re.findall(r"<c:ser>.*?</c:ser>", xml, re.S)

    # 1a passada: aba, linha do mes e valores das series que apontam para a planilha
    infos = []
    ws = linha = None
    for ser in series:
        m_val = re.search(r"<c:val><c:numRef><c:f>([^<]+)</c:f>", ser)
        if not m_val:
            infos.append(None)          # serie "Total" (valores literais)
            continue
        m = _RE_REF.match(m_val.group(1).replace("&apos;", "'"))
        if not m:
            raise ValueError(f"Intervalo de serie nao reconhecido: {m_val.group(1)}")
        aba, col, fim = m.group(1) or m.group(2), m.group(3), int(m.group(6))
        ws = wb[aba]
        linha = encontrar_linha_mes(ws, mes_serial_excel(ano, mes), LINHA_DADOS)
        if linha is None:
            raise ValueError(f"Mes {mes:02d}/{ano} nao encontrado na aba {aba}")
        if linha <= fim:
            return None
        if linha != fim + 1:
            raise ValueError(f"Grafico da aba {aba} termina na linha {fim}; o mes novo esta na {linha}")
        infos.append({"col": col, "valor": _valor(ws, col, linha)})

    total = sum(i["valor"] for i in infos if i)
    rotulo_mes = f"{MESES_PT[mes]}/{str(ano)[2:]}"

    # 2a passada: reescreve cada serie
    for ser, info in zip(series, infos):
        novo = ser
        if info:
            novo = re.sub(
                r"(<c:f>[^<]*?\$[A-Z]+\$\d+:\$[A-Z]+\$)(\d+)(</c:f>)",
                lambda m: f"{m.group(1)}{int(m.group(2)) + 1}{m.group(3)}",
                novo,
            )
            texto = f"R$ {_formato_br(info['valor'])} ({info['valor'] / total * 100:.0f}%)"
            valor_cache = repr(info["valor"])
        else:
            texto = f"R$ {_formato_br(total)}"
            valor_cache = "0"

        cat = re.search(r"<c:cat>.*?</c:cat>", novo, re.S).group(0)
        cache_cat = rotulo_mes if "<c:strLit>" in cat else mes_serial_excel(ano, mes)
        novo = novo.replace(cat, _acrescentar_ponto(cat, cache_cat), 1)
        val = re.search(r"<c:val>.*?</c:val>", novo, re.S).group(0)
        novo = novo.replace(val, _acrescentar_ponto(val, valor_cache), 1)

        # Rotulo do ponto novo: copia do ultimo rotulo, com idx, texto e id novos
        rotulos = re.findall(r"<c:dLbl>.*?</c:dLbl>", novo, re.S)
        if rotulos:
            ultimo = rotulos[-1]
            idx = int(re.search(r'<c:idx val="(\d+)"/>', ultimo).group(1))
            rot = ultimo.replace(f'<c:idx val="{idx}"/>', f'<c:idx val="{idx + 1}"/>', 1)
            rot = re.sub(r"<a:t>[^<]*</a:t>", f"<a:t>{texto}</a:t>", rot, count=1)
            rot = re.sub(
                r'(<c16:uniqueId val=")[^"]+(")',
                lambda m: f"{m.group(1)}{{{str(uuid.uuid4()).upper()}}}{m.group(2)}",
                rot,
            )
            pos = novo.rfind(ultimo) + len(ultimo)
            novo = novo[:pos] + rot + novo[pos:]

        xml = xml.replace(ser, novo, 1)

    print(f"  [OK] Grafico da aba {ws.title}: {rotulo_mes} acrescentado (total R$ {_formato_br(total)})")
    return xml


def restaurar_graficos(origem, destino, ano, mes):
    """Copia os graficos de `origem` (arquivo salvo pelo Excel, com os rotulos intactos)
    para `destino` (planilha ja atualizada pelo openpyxl), acrescentando o mes informado."""
    origem, destino = Path(origem), Path(destino)
    wb = openpyxl.load_workbook(str(destino))

    novos = {}
    with zipfile.ZipFile(origem) as z:
        for nome in z.namelist():
            if not re.fullmatch(r"xl/charts/chart\d+\.xml", nome):
                continue
            xml = z.read(nome).decode("utf-8")
            if "<c:chartSpace" not in xml:
                raise ValueError(
                    f"{origem.name}: {nome} ja foi regravado pelo openpyxl (rotulos perdidos). "
                    "Use como origem um arquivo salvo pelo Excel."
                )
            estendido = _estender_grafico(xml, wb, ano, mes)
            novos[nome] = (estendido or xml).encode("utf-8")
            if estendido is None:
                print(f"  [OK] {nome}: mes ja estava no grafico, apenas restaurado")

    temp = destino.parent / ("~temp_graficos_" + destino.name)
    with zipfile.ZipFile(destino) as zin, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
        faltando = set(novos) - set(zin.namelist())
        if faltando:
            raise ValueError(f"{destino.name} nao tem as partes {sorted(faltando)}")
        for item in zin.infolist():
            zout.writestr(item, novos.get(item.filename, zin.read(item.filename)))

    try:
        os.replace(temp, destino)
        print(f"[OK] Graficos restaurados e atualizados em {destino.name}")
    except PermissionError:
        print(f"\n[!] Arquivo bloqueado (aberto no Excel ou sincronizando no Google Drive).")
        print(f"    Planilha com os graficos corrigidos salva em: {temp.name}")
        print(f"    Feche o Excel e substitua o arquivo original pelo '{temp.name}'")
