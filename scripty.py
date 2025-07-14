from odf.opendocument import load
from odf.text         import P
from odf.table        import Table, TableRow, TableCell
from datetime         import datetime, timedelta
from pathlib          import Path
from collections      import defaultdict
import re

# ------------ utilidades ------------
def data_para_nome_br(d: datetime) -> str:
    meses = ["","JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO",
             "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"]
    return f"{d.day:02d} {meses[d.month]} {str(d.year)[-2:]}"

def achar_arquivo_escala(pasta: str, data_ref: datetime) -> Path|None:
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

# ------------ configuração dos cargos ------------
cargos_para_placeholder = {
    "SGT_DE_DIA"         : "SGT DE DIA",
    "SGT_PERMANENCIA"    : "SGT PERMANÊNCIA",
    "CB_DE_DIA"          : "CB DE DIA SU",
    "PLANTAO_1"          : "PLANTÕES SU",
    "PLANTAO_2"          : "PLANTÕES SU",
    "PLANTAO_3"          : "PLANTÕES SU",
    "GDA_QTL_1"          : "GDA QTL 02",
    "GDA_QTL_2"          : "GDA QTL 02",
    "MOTORISTA"          : "MOTORISTA DE DIA",
    "PERMANENCIA_PMT"    : "PERMANENCIA PMT",
    "PERMANENCIA_ENFER"  : "PERMANÊNCIA ENFERMARIA",
    "SENTINELA_1"        : "FAX PAVILHAO"
}

# ------------ coleta dos nomes da escala ------------
hoje   = datetime.today()
ontem  = hoje - timedelta(days=1)
arqADT = achar_arquivo_escala("escalas", ontem)
if not arqADT:
    raise FileNotFoundError("❌ Escala não encontrada")

docADT = load(arqADT)

# agrupa nomes por texto do cargo (ex: "PLANTÕES SU" -> [nomes...])
nomes_por_funcao = defaultdict(list)

for tabela in docADT.getElementsByType(Table):
    for linha in tabela.getElementsByType(TableRow):
        celulas = linha.getElementsByType(TableCell)
        if not celulas:
            continue
        cargo_txt = extrair_texto(celulas[0]).strip().upper()
        for chave, nome_cargo in cargos_para_placeholder.items():
            if cargo_txt == nome_cargo.upper():
                # adiciona nomes das células seguintes na lista do cargo
                for cel in celulas[1:]:
                    nome = extrair_texto(cel).strip()
                    if nome:
                        nomes_por_funcao[nome_cargo].append(nome)

# preenche mapa_funcoes distribuindo nomes para cada chave do placeholder
mapa_funcoes = {}

for chave, nome_cargo in cargos_para_placeholder.items():
    nomes = nomes_por_funcao.get(nome_cargo, [])
    # para cargos repetidos (ex: PLANTAO_1, PLANTAO_2, ...) pega índice relativo da chave
    # índice é a posição da chave entre as chaves que possuem o mesmo nome_cargo
    chaves_mesmo_cargo = [k for k, v in cargos_para_placeholder.items() if v == nome_cargo]
    index = chaves_mesmo_cargo.index(chave) if chave in chaves_mesmo_cargo else 0
    mapa_funcoes[chave] = nomes[index] if index < len(nomes) else "----------"

# ------------ substitui placeholders no modelo ------------
def substituir_placeholders(doc, dados):
    def substituir_no_nodo(nodo):
        if nodo.nodeType == 3:               # TEXT_NODE
            txt = nodo.data
            novo = txt
            for k, v in dados.items():
                novo = novo.replace(f"{{{{{k}}}}}", v)
            if novo != txt:
                nodo.data = novo
        elif hasattr(nodo, "childNodes"):
            for filho in nodo.childNodes:
                substituir_no_nodo(filho)

    for elem in doc.getElementsByType(P) + doc.getElementsByType(TableCell):
        substituir_no_nodo(elem)

modelo = load("modelo_pernoite.odt")
substituir_placeholders(modelo, mapa_funcoes)

# ------------ salva o resultado ------------
Path("pernoites").mkdir(exist_ok=True)
saida = Path("pernoites") / f"pernoite_{data_para_nome_br(ontem).replace(' ','_')}.odt"
modelo.save(saida)
print("✅ Pernoite gerado em:", saida)
