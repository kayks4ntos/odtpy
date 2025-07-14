from odf.opendocument import load, OpenDocumentText
from odf.text import H, P
from odf.table import Table, TableRow, TableCell
from datetime import datetime, timedelta
from pathlib import Path
import re

# ------------------------------
# Função para formatar data estilo "09 JULHO 25"
# ------------------------------
def data_para_nome_br(data: datetime) -> str:
    meses_pt = {
        1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO",     4: "ABRIL",
        5: "MAIO",    6: "JUNHO",      7: "JULHO",     8: "AGOSTO",
        9: "SETEMBRO",10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
    }
    dia = f"{data.day:02d}"               # Ex: "09"
    mes = meses_pt[data.month]           # Ex: "JULHO"
    ano = str(data.year)[-2:]            # Ex: "25"
    return f"{dia} {mes} {ano}"

# ------------------------------
# Função para encontrar o arquivo .odt da escala
# ------------------------------
def achar_arquivo_escala(pasta: str, data_ref: datetime) -> Path | None:
    data_str = data_para_nome_br(data_ref)  # Ex: "09 JULHO 25"
    padrao = re.compile(rf"ADT \d+ DE {re.escape(data_str)}\.odt$", re.I)

    for arquivo in Path(pasta).iterdir():
        if padrao.match(arquivo.name):
            return arquivo
    return None

# ------------------------------
# Função para extrair texto puro de um elemento ODF (recursiva)
# ------------------------------
def extrair_texto(elemento) -> str:
    textos = []
    for node in elemento.childNodes:
        if node.nodeType == 3:  # TEXT_NODE
            textos.append(node.data)
        elif hasattr(node, "childNodes"):
            textos.append(extrair_texto(node))
    return "".join(textos)

# ------------------------------
# Parte principal do script
# ------------------------------

# Define a data de ontem
hoje = datetime.today()
ontem = hoje - timedelta(days=1)

# Tenta localizar a escala
arq_escala = achar_arquivo_escala("escalas", ontem)
if arq_escala is None:
    raise FileNotFoundError(f"❌ Escala de {data_para_nome_br(ontem)} não encontrada na pasta 'escalas'.")

print("📄 Arquivo da escala encontrado:", arq_escala.name)

# Abre a escala
odt = load(arq_escala)

# Lista dos cargos desejados para filtro nas tabelas
cargos_desejados = [
    "SGT DE DIA",
    "SGT PERMANÊNCIA",
    "CB DE DIA SU",
    "PLANTÕES SU",
    "GDA QTL 02",
    "MOTORISTA DE DIA",
    "PERMANENCIA PMT",
    "PERMANÊNCIA ENFERMARIA",
    "FAX PAVILHAO",
]

nomes = []

# --- Busca por parágrafos com "SENTINELA" e "PLANTÃO" ---
for par in odt.getElementsByType(P):
    texto = extrair_texto(par).strip().upper()
    if any(palavra in texto for palavra in ("SENTINELA", "PLANTÃO")):
        nomes.append(texto)

# --- Busca por nomes nas tabelas filtrando pelo cargo ---
for tabela in odt.getElementsByType(Table):
    for linha in tabela.getElementsByType(TableRow):
        celulas = linha.getElementsByType(TableCell)
        if not celulas:
            continue
        # Obtém o texto da primeira célula da linha (cargo)
        primeiro_paragrafo = celulas[0].getElementsByType(P)
        if not primeiro_paragrafo:
            continue
        cargo = extrair_texto(primeiro_paragrafo[0]).strip().upper()
        if cargo in [c.upper() for c in cargos_desejados]:
            # Nas outras células da linha, extrai os nomes
            for celula in celulas[1:]:
                for par in celula.getElementsByType(P):
                    texto = extrair_texto(par).strip()
                    if texto:
                        nomes.append(texto)

# Remove possíveis duplicatas e ordena a lista
nomes = sorted(set(nomes))

print("👥 Nomes encontrados filtrados:")
for nome in nomes:
    print("-", nome)

# Cria o documento de pernoite
pernoite_doc = OpenDocumentText()
pernoite_doc.text.addElement(H(outlinelevel=1, text=f"Pernoite - {data_para_nome_br(ontem)}"))

for nome in nomes:
    pernoite_doc.text.addElement(P(text=nome))






# Salva o pernoite
Path("pernoites").mkdir(exist_ok=True)  # cria pasta se não existir
nome_pernoite = f"pernoite_{data_para_nome_br(ontem).replace(' ', '_')}.odt"
caminho_final = Path("pernoites") / nome_pernoite
pernoite_doc.save(caminho_final)

print(f"✅ Documento de pernoite salvo em: {caminho_final}")
