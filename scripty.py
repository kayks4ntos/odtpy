from odf.opendocument import load
from odf.text import P, LineBreak
from odf.table import Table, TableRow, TableCell
from odf.element import Element
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict
import re

# ---------------- utilidades ----------------
def data_para_nome_br(d: datetime) -> str:
    meses = ["", "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
             "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
    return f"{d.day:02d} {meses[d.month]} {str(d.year)[-2:]}"

def achar_arquivo_escala(pasta: str, data_ref: datetime) -> Path | None:
    padrao = re.compile(rf"ADT \d+ DE {re.escape(data_para_nome_br(data_ref))}\.odt$", re.I)
    for arq in Path(pasta).iterdir():
        if padrao.match(arq.name):
            return arq
    return None

def extrair_texto(elem) -> str:
    txt = []
    for n in elem.childNodes:
        if n.nodeType == 3:  # TEXT_NODE
            txt.append(n.data)
        elif hasattr(n, "childNodes"):
            txt.append(extrair_texto(n))
    return "".join(txt)

# ------------ configuração de cargos ------------

cargos_para_placeholder = {
    # GUARNIÇÃO INTERNA
    "SGT_DE_DIA"        : "SGT DE DIA",
    "CB_DE_DIA"         : "CB DE DIA SU",
    "PLANTAO_1"         : "PLANTÕES SU",
    "PLANTAO_2"         : "PLANTÕES SU",
    "PLANTAO_3"         : "PLANTÕES SU",

    # GUARNIÇÃO EXTERNA
    "MOTORISTA"         : "MOTORISTA DE DIA",
    "PERMANENCIA_ENFER" : "PERMANÊNCIA ENFERMARIA",
    "SENTINELA_1"       : "GDA QTL 02",
    "SENTINELA_2"       : "GDA QTL 02",

    "OFICIAL_DE_DIA"    : "OFICIAL DE DIA",
    "ADJUNTO"           : "ADJUNTO",
    "CMT_GDA"           : "CMT DA GDA",
    "CB_GDA_I"          : "CB DA GDA I",
    "CB_GDA_II"         : "CB DA GDA II"
}

# ------------ coleta de nomes da escala ------------

hoje  = datetime.today()
ontem = hoje - timedelta(days=1)

arqADT = achar_arquivo_escala("escalas", ontem)
if not arqADT:
    raise FileNotFoundError("❌ Escala não encontrada na pasta 'escalas'.")

docADT = load(arqADT)

nomes_por_funcao = defaultdict(list)

for tabela in docADT.getElementsByType(Table):
    for linha in tabela.getElementsByType(TableRow):
        celulas = linha.getElementsByType(TableCell)
        if not celulas:
            continue
        cargo_txt = extrair_texto(celulas[0]).strip().upper()
        for chave, nome_cargo in cargos_para_placeholder.items():
            if cargo_txt == nome_cargo.upper():
                for cel in celulas[1:]:
                    nome = extrair_texto(cel).strip()
                    if nome:
                        nomes_por_funcao[nome_cargo].append(nome)

# ------------ monta mapa_funcoes -------------------

mapa_funcoes = {}

for chave, nome_cargo in cargos_para_placeholder.items():
    nomes = nomes_por_funcao.get(nome_cargo, [])
    chaves_mesmo_cargo = [k for k,v in cargos_para_placeholder.items() if v == nome_cargo]
    idx = chaves_mesmo_cargo.index(chave) if chave in chaves_mesmo_cargo else 0

    if idx < len(nomes):
        nome_puro = nomes[idx].strip()
        if chave == "PERMANENCIA_ENFER":
            nome_puro = re.sub(r"^PERMAN[ÊE]NCIA ENFERMARIA[:\- ]*", "", nome_puro, flags=re.I).strip()
        mapa_funcoes[chave] = nome_puro
    else:
        mapa_funcoes[chave] = ""

# ------------ datas do documento -------------------

mapa_funcoes["DATA_DE_HOJE"] = f"{hoje.day:02d}"
mapa_funcoes["MES_DE_HJ"]    = data_para_nome_br(hoje).split()[1].capitalize().upper()
mapa_funcoes["MES_DE_HJ_n"]  = data_para_nome_br(hoje).split()[1].capitalize()

# ------------ placeholders vazios como "–" -----------

for k in ["OFICIAL_DE_DIA", "ADJUNTO", "CMT_GDA", "CB_GDA_I", "CB_GDA_II"]:
    if not mapa_funcoes.get(k):
        mapa_funcoes[k] = "–"

# ------------ formatação especial para CBs da GDA ------------

cb_gda_i = mapa_funcoes["CB_GDA_I"]
cb_gda_ii = mapa_funcoes["CB_GDA_II"]

if cb_gda_i != "–" and cb_gda_ii != "–":
    mapa_funcoes["CB_GUARNICAO"] = f"- CB DA GDA I: {cb_gda_i}\n  CB DA GDA II: {cb_gda_ii}"
elif cb_gda_i != "–":
    mapa_funcoes["CB_GUARNICAO"] = f"- CB DA GDA I: {cb_gda_i}"
elif cb_gda_ii != "–":
    mapa_funcoes["CB_GUARNICAO"] = f"- CB DA GDA II: {cb_gda_ii}"
else:
    mapa_funcoes["CB_GUARNICAO"] = "–"

# ------------ Guarnição Interna ------------

soma_internos = {
    "SGT": 1 if mapa_funcoes["SGT_DE_DIA"] else 0,
    "CB":  1 if mapa_funcoes["CB_DE_DIA"] else 0,
    "SOLDADO": sum(1 for k in ["PLANTAO_1", "PLANTAO_2", "PLANTAO_3"] if mapa_funcoes.get(k))
}
mapa_funcoes["SOMA_SGT_INT"] = str(soma_internos["SGT"])
mapa_funcoes["SOMA_CB_INT"] = str(soma_internos["CB"])
mapa_funcoes["SD_INT"] = str(soma_internos["SOLDADO"])
mapa_funcoes["SOMA_TOTAL_INT"] = str(sum(soma_internos.values()))

# ------------ Guarnição Externa ------------

soma_externos = {
    "OF": int(mapa_funcoes.get("OFICIAL_DE_DIA") not in ["", "–"]),
    "SGT": sum(1 for k in ["ADJUNTO", "CMT_GDA"] if mapa_funcoes.get(k) not in ["", "–"]),
    "CB": sum(1 for k in ["CB_GDA_I", "CB_GDA_II"] if mapa_funcoes.get(k) not in ["", "–"]),
    "SOLDADO": sum(1 for k in ["MOTORISTA", "PERMANENCIA_ENFER", "SENTINELA_1", "SENTINELA_2"] if mapa_funcoes.get(k) not in ["", "–"])
}
mapa_funcoes["SOMA_OF"] = str(soma_externos["OF"])
mapa_funcoes["SOMA_SGT"] = str(soma_externos["SGT"])
mapa_funcoes["SOMA_CB"] = str(soma_externos["CB"])
mapa_funcoes["SD"] = str(soma_externos["SOLDADO"])
mapa_funcoes["SOMA_TOTAL"] = str(sum(soma_externos.values()))

# ------------ Total Geral ------------
total_geral = sum(soma_internos.values()) + sum(soma_externos.values())
mapa_funcoes["TOTAL_GERAL"] = str(total_geral)

# ------------ substituição de placeholders ------------

def substituir_placeholders(doc, dados):
    for p in doc.getElementsByType(P):
        texto = extrair_texto(p)
        if not texto:
            continue
        novo_texto = texto
        for k, v in dados.items():
            novo_texto = novo_texto.replace(f"{{{{{k}}}}}", v)
        if novo_texto != texto:
            filhos = [child for child in p.childNodes if isinstance(child, Element)]
            for child in filhos:
                p.removeChild(child)
            p.addText(novo_texto)

    for cell in doc.getElementsByType(TableCell):
        texto = extrair_texto(cell)
        if not texto:
            continue
        novo_texto = texto
        for k, v in dados.items():
            novo_texto = novo_texto.replace(f"{{{{{k}}}}}", v)
        if novo_texto != texto:
            filhos = [child for child in cell.childNodes if isinstance(child, Element)]
            for child in filhos:
                cell.removeChild(child)
            novo_p = P()
            for part in novo_texto.split("\n"):
                novo_p.addText(part)
                novo_p.addElement(LineBreak())
            cell.addElement(novo_p)

# ------------ aplica modelo e salva ------------

modelo = load("modelo_pernoite.odt")
substituir_placeholders(modelo, mapa_funcoes)

Path("pernoites").mkdir(exist_ok=True)
nome_saida = f"pernoite_{data_para_nome_br(hoje).replace(' ','_')}.odt"
modelo.save(Path("pernoites") / nome_saida)

print("✅ Pernoite gerado em:", Path("pernoites") / nome_saida)
