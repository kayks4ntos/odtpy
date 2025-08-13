import sys
from pathlib import Path

# Adiciona a pasta 'libs' ao caminho de importação
sys.path.insert(0, str(Path(__file__).parent / "libs"))

from odf.opendocument import load
from odf.text import P, LineBreak
from odf.table import Table, TableRow, TableCell
from odf.element import Element
from odf.style import Style, ParagraphProperties
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
    "SGT_DE_DIA"        : "SGT DE DIA",
    "CB_DE_DIA"         : "CB DE DIA SU",
    "PLANTAO_1"         : "PLANTÕES SU",
    "PLANTAO_2"         : "PLANTÕES SU",
    "PLANTAO_3"         : "PLANTÕES SU",
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
hoje = datetime.today()
ontem = hoje - timedelta(days=1)

pasta_adt = r"C:\Users\Sgte-CCAP\Desktop\Sargenteação 2025\01 - ADT 2025\07 JULHO 25"
arqADT = achar_arquivo_escala(pasta_adt, ontem)
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
        print(f"Cargo identificado: '{cargo_txt}'")
        for chave, nome_cargo in cargos_para_placeholder.items():
            if cargo_txt == nome_cargo.upper():
                nomes_extraidos = set()  # ← evita duplicatas
                for cel in celulas[1:]:
                    nome = extrair_texto(cel).strip()
                    if nome:
                        nomes_extraidos.add(nome)
                nomes_por_funcao[nome_cargo].extend(nomes_extraidos)

# ------------ monta mapa_funcoes -------------------
mapa_funcoes = {}

for chave, nome_cargo in cargos_para_placeholder.items():
    nomes = nomes_por_funcao.get(nome_cargo, [])
    chaves_mesmo_cargo = [k for k, v in cargos_para_placeholder.items() if v == nome_cargo]
    idx = chaves_mesmo_cargo.index(chave) if chave in chaves_mesmo_cargo else 0

    if idx < len(nomes):
        nome_puro = nomes[idx].strip()
        # Remove prefixo duplicado se estiver presente
        prefixo = re.escape(nome_cargo)
        nome_puro = re.sub(rf"^{prefixo}[:\- ]*", "", nome_puro, flags=re.I).strip()
        mapa_funcoes[chave] = nome_puro if nome_puro else ""
    else:
        mapa_funcoes[chave] = ""

# ------------ datas do documento -------------------
mapa_funcoes["DATA_DE_HOJE"] = f"{hoje.day:02d}"
mapa_funcoes["MES_DE_HJ"] = data_para_nome_br(hoje).split()[1].upper()
mapa_funcoes["MES_DE_HJ_n"] = data_para_nome_br(hoje).split()[1].capitalize()

# ------------ placeholders vazios como "–" -----------
for k in ["OFICIAL_DE_DIA", "ADJUNTO", "CMT_GDA", "CB_GDA_I", "CB_GDA_II"]:
    if not mapa_funcoes.get(k):
        mapa_funcoes[k] = "–"

# ------------ formatação especial para CBs da GDA ------------
cb_gda_i = mapa_funcoes.get("CB_GDA_I", "–")
cb_gda_ii = mapa_funcoes.get("CB_GDA_II", "–")

cb_lines = []
if cb_gda_i != "–":
    cb_lines.append(f"- CB DA GDA I: {cb_gda_i}")
if cb_gda_ii != "–":
    cb_lines.append(f"- CB DA GDA II: {cb_gda_ii}")

mapa_funcoes["CB_GUARNICAO"] = "\n  ".join(cb_lines) if cb_lines else "–"

# ------------ Guarnição Interna ------------
soma_internos = {
    "SGT": 1 if mapa_funcoes.get("SGT_DE_DIA", "–") != "–" else 0,
    "CB":  1 if mapa_funcoes.get("CB_DE_DIA", "–") != "–" else 0,
    "SOLDADO": sum(1 for k in ["PLANTAO_1", "PLANTAO_2", "PLANTAO_3"] if mapa_funcoes.get(k, "–") != "–")
}
mapa_funcoes["SOMA_SGT_INT"] = f"{soma_internos['SGT']:02d}"
mapa_funcoes["SOMA_CB_INT"] = f"{soma_internos['CB']:02d}"
mapa_funcoes["SD_INT"] = f"{soma_internos['SOLDADO']:02d}"
mapa_funcoes["SOMA_TOTAL_INT"] = f"{sum(soma_internos.values()):02d}"

# ------------ Guarnição Externa ------------
soma_externos = {
    "OF": int(mapa_funcoes.get("OFICIAL_DE_DIA", "–") != "–"),
    "SGT": sum(1 for k in ["ADJUNTO", "CMT_GDA"] if mapa_funcoes.get(k, "–") != "–"),
    "CB": sum(1 for k in ["CB_GDA_I", "CB_GDA_II"] if mapa_funcoes.get(k, "–") != "–"),
    "SOLDADO": sum(1 for k in ["MOTORISTA", "PERMANENCIA_ENFER", "SENTINELA_1", "SENTINELA_2"] if mapa_funcoes.get(k, "–") != "–")
}
mapa_funcoes["SOMA_OF"] = f"{soma_externos['OF']:02d}"
mapa_funcoes["SOMA_SGT"] = f"{soma_externos['SGT']:02d}"
mapa_funcoes["SOMA_CB"] = f"{soma_externos['CB']:02d}"
mapa_funcoes["SD"] = f"{soma_externos['SOLDADO']:02d}"
mapa_funcoes["SOMA_TOTAL"] = f"{sum(soma_externos.values()):02d}"

# ------------ Total Geral ------------
total_geral = sum(soma_internos.values()) + sum(soma_externos.values())
mapa_funcoes["TOTAL_GERAL"] = f"{total_geral:02d}"

# ------------ SARGENTOS_EXT ------------
adjunto = mapa_funcoes["ADJUNTO"]
cmt_gda = mapa_funcoes["CMT_GDA"]

# Eliminar duplicações se forem iguais
sgt_set = set()

if adjunto != "–" and adjunto != "":
    sgt_set.add(f"ADJUNTO: {adjunto}")
if cmt_gda != "–" and cmt_gda != "":
    sgt_set.add(f"CMT DA GUARDA: {cmt_gda}")

# Convertendo para lista para manter consistência
sgt_lines = list(sgt_set)

# Para manter ordem preferencial: ADJUNTO primeiro
sgt_lines.sort(key=lambda x: 0 if x.startswith("ADJUNTO") else 1)

mapa_funcoes["SARGENTOS_EXT"] = "\n  ".join(sgt_lines) if sgt_lines else "–"


# ------------ substituição de placeholders ------------
def substituir_placeholders(doc, dados):
    estilo_centralizado = Style(name="Centralizado", family="paragraph")
    estilo_centralizado.addElement(ParagraphProperties(textalign="center"))
    doc.styles.addElement(estilo_centralizado)

    chaves_para_centralizar_se_hifen = {
        "OFICIAL_DE_DIA", "SARGENTOS_EXT", "CB_GUARNICAO"
    }
    chaves_somas = {
        "SOMA_SGT_INT", "SOMA_CB_INT", "SD_INT", "SOMA_TOTAL_INT",
        "SOMA_OF", "SOMA_SGT", "SOMA_CB", "SD", "SOMA_TOTAL", "TOTAL_GERAL"
    }

    for p in doc.getElementsByType(P):
        texto = extrair_texto(p)
        if not texto:
            continue
        novo_texto = texto
        for k, v in dados.items():
            v = v.strip().rstrip('-').strip()
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
        estilo_para_usar = None

        for k, v in dados.items():
            v = v.strip().rstrip('-').strip()
            if f"{{{{{k}}}}}" in novo_texto:
                novo_texto = novo_texto.replace(f"{{{{{k}}}}}", v)
                if v == "–" and k in chaves_para_centralizar_se_hifen:
                    estilo_para_usar = estilo_centralizado
                elif k in chaves_somas:
                    estilo_para_usar = estilo_centralizado

        if novo_texto != texto:
            filhos = [child for child in cell.childNodes if isinstance(child, Element)]
            for child in filhos:
                cell.removeChild(child)

            novo_p = P(stylename=estilo_para_usar) if estilo_para_usar else P()
            for i, part in enumerate(novo_texto.split("\n")):
                novo_p.addText(part)
                if i < len(novo_texto.split("\n")) - 1:
                    novo_p.addElement(LineBreak())
            cell.addElement(novo_p)

# ------------ aplica modelo e salva ------------
modelo = load("modelo_pernoite.odt")
substituir_placeholders(modelo, mapa_funcoes)

pasta_saida = Path(r"C:\Users\Sgte-CCAP\Desktop\Sargenteação 2025\02 - PERNOITE 2025\07 JULHO 2025")
pasta_saida.mkdir(parents=True, exist_ok=True)

nome_saida = f"pernoite_{data_para_nome_br(hoje).replace(' ', '_')}.odt"
caminho_final = pasta_saida / nome_saida
modelo.save(caminho_final)

print("✅ Pernoite gerado em:", caminho_final)