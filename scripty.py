from odf.opendocument import load
from odf.text import P
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

# ------------ configuração de cargos corrigida ------------
cargos_para_placeholder = {
    "SGT_DE_DIA"        : "SGT DE DIA",
    "SGT_PERMANENCIA"   : "SGT PERMANÊNCIA",
    "CB_DE_DIA"         : "CB DE DIA SU",

    "PLANTAO_1"         : "PLANTÕES SU",
    "PLANTAO_2"         : "PLANTÕES SU",
    "PLANTAO_3"         : "PLANTÕES SU",

    "MOTORISTA"         : "MOTORISTA DE DIA",

    "PERMANENCIA_ENFER" : "PERMANÊNCIA ENFERMARIA",

    # Dois sentinelas, ambos com mesmo nome na escala "GDA QTL 02"
    "SENTINELA_1"       : "GDA QTL 02",
    "SENTINELA_2"       : "GDA QTL 02",
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
    mapa_funcoes[chave] = nomes[idx] if idx < len(nomes) else "----------"

# Datas para o modelo
mapa_funcoes["DATA_DE_HOJE"] = f"{ontem.day:02d}"
mapa_funcoes["MES_DE_HJ"]    = data_para_nome_br(ontem).split()[1].capitalize().upper()
mapa_funcoes["MES_DE_HJ_n"]    = data_para_nome_br(ontem).split()[1].capitalize()

# ------------ substituição de placeholders ------------
def substituir_placeholders(doc, dados):
    from odf.text import Span

    # Substitui em parágrafos
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

    # Substitui em células de tabela
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
            novo_p.addText(novo_texto)
            cell.addElement(novo_p)

# ------------ aplica modelo e salva ----------------
modelo = load("modelo_pernoite.odt")
substituir_placeholders(modelo, mapa_funcoes)

Path("pernoites").mkdir(exist_ok=True)
nome_saida = f"pernoite_{data_para_nome_br(hoje).replace(' ','_')}.odt"
modelo.save(Path("pernoites") / nome_saida)

print("✅ Pernoite gerado em:", Path('pernoites') / nome_saida)