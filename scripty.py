from odf.opendocument import load
from odf.opendocument import OpenDocumentText
from odf.text import H, P
from datetime import datetime, timedelta

# Abrir arquivo .odt da escala
hoje = datetime.today()
ontem = hoje - timedelta(days=1)
data_ontem = ontem.strftime("%d-%m-%Y")

arquivo = f"escalas/escala_{data_ontem}.odt"
odt = load(arquivo)

# Extrair parágrafos de texto
nomes = []
for elem in odt.getElementsByType(P):
    texto = str(elem)
    if "SENTINELA" in texto or "PLANTÃO" in texto:  # Ajustar para o que identifica os nomes
        nomes.append(elem.firstChild.data)

# Mostrar os nomes extraídos
for nome in nomes:
    print("Nome extraído:", nome)
pernoite_doc = OpenDocumentText()

# Título
pernoite_doc.text.addElement(H(outlinelevel=1, text=f"Pernoite - {data_ontem}"))

# Inserir nomes
for nome in nomes:
    par = P(text=nome)
    pernoite_doc.text.addElement(par)

# Salvar
pernoite_doc.save(f"pernoites/pernoite_{data_ontem}.odt")
print("Pernoite gerado com sucesso.")
