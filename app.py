from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import gdown
from fpdf import FPDF
import os

app = Flask(__name__)

# ID do arquivo no Google Drive
GOOGLE_DRIVE_FILE_ID = "1mPYlc_uC3SfJnNQ_ToG6eVmn2ZYMhPCX"

def get_excel_from_google_drive():
    """Baixa a planilha do Google Drive e retorna o DataFrame."""
    url = f"https://drive.google.com/uc?id={GOOGLE_DRIVE_FILE_ID}"
    output_file = "patrimonio.xlsx"  # Nome temporário do arquivo baixado
    gdown.download(url, output_file, quiet=False)
    return pd.read_excel(output_file)

# Carregar a planilha no início do programa
df = get_excel_from_google_drive()

class PDF(FPDF):
    def __init__(self):
        super().__init__('P', 'mm', 'A4')  # Orientação retrato, milímetros, formato A4

    def header(self):
        # Cabeçalho centralizado
        self.set_font("Arial", "B", 12)
        self.cell(0, 6, "MINISTÉRIO DA DEFESA", ln=True, align="C")
        self.cell(0, 6, "COMANDO DA AERONÁUTICA", ln=True, align="C")
        self.cell(0, 6, "GRUPAMENTO DE APOIO DE LAGOA SANTA", ln=True, align="C")
        self.cell(0, 8, "GUIA DE MOVIMENTAÇÃO DE BEM MÓVEL PERMANENTE ENTRE AS SEÇÕES DO GAPLS", ln=True, align="C")
        self.ln(10)

    def fix_text(self, text):
        """Corrige caracteres incompatíveis com a codificação latin-1."""
        replacements = {
            "–": "-",  # Substituir travessão por hífen
            "“": '"',  # Substituir aspas abertas por aspas duplas
            "”": '"',  # Substituir aspas fechadas por aspas duplas
            "’": "'",  # Substituir apóstrofo por aspas simples
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text

    def add_table(self, dados_bmps):
        col_widths = [15, 55, 15, 30, 30, 30]
        headers = ["Nº BMP", "Nomenclatura", "Qtde", "Valor Atualizado", "Qtde a Movimentar", "Valor a Movimentar"]

        # Adicionar cabeçalho da tabela
        self.set_font("Arial", "B", 10)
        for width, header in zip(col_widths, headers):
            self.cell(width, 10, header, border=1, align="C")
        self.ln()

        # Adicionar as linhas da tabela
        self.set_font("Arial", size=10)
        for _, row in dados_bmps.iterrows():
            text = self.fix_text(row["NOMECLATURA/COMPONENTE"])
            self.cell(col_widths[0], 10, str(row["Nº BMP"]), border=1, align="C")
            self.cell(col_widths[1], 10, text, border=1, align="L")
            self.cell(col_widths[2], 10, str(row["QTD"]), border=1, align="C")
            self.cell(col_widths[3], 10, f"R$ {row['VL. ATUALIZ.']:.2f}".replace(".", ","), border=1, align="C")
            self.cell(col_widths[4], 10, str(row["Qtde a Movimentar"]), border=1, align="C")
            self.cell(col_widths[5], 10, f"R$ {row['Valor a Movimentar']:.2f}".replace(".", ","), border=1, align="C")
            self.ln()

    def add_details(self, secao_destino, chefia_origem, secao_origem, chefia_destino):
        text = f"""
Solicitação de Transferência:
Informo à Senhora Chefe do GAP-LS que os bens especificados estão inservíveis para uso neste setor, classificados como ociosos, recuperáveis, reparados ou novos - aguardando distribuição. Diante disso, solicito autorização para transferir o(s) Bem(ns) Móvel(is) Permanente(s) acima discriminado(s), atualmente sob minha guarda, para a Seção {secao_destino}.

{chefia_origem}
{secao_origem}

Confirmação da Seção de Destino:
Estou ciente da movimentação informada acima e, devido à necessidade do setor, solicito à Senhora Dirigente Máximo autorização para manter sob minha guarda os Bens Móveis Permanentes especificados.

{chefia_destino}
{secao_destino}

DO AGENTE DE CONTROLE INTERNO AO DIRIGENTE MÁXIMO
Informo à Senhora que, após conferência, foi verificado que esta guia cumpre o disposto no Módulo D do RADA-e e, conforme a alínea "d" do item 5.3 da ICA 179-1, encaminho para apreciação e se for o caso, autorização.

KARINA RAQUEL VALIMAREANU  Maj Int
Chefe da ACI

DESPACHO DA AGENTE DIRETOR
Autorizo a movimentação solicitada e determino:
1. Que a Seção de Registro realize a movimentação no SILOMS.
2. Que a Seção de Registro publique a movimentação no próximo aditamento a ser confeccionado, conforme o item 2.14.2, Módulo do RADA-e.
3. Que os detentores realizem a movimentação física do(s) bem(ns).

LUCIANA DO AMARAL CORREA  Cel Int
Dirigente Máximo
"""
        self.multi_cell(0, 8, self.fix_text(text))

@app.route("/", methods=["GET", "POST"])
def index():
    secoes_origem = df['Seção de Origem'].dropna().unique().tolist()
    secoes_destino = df['Seção de Destino'].dropna().unique().tolist()

    if request.method == "POST":
        bmp_numbers = request.form.get("bmp_numbers")
        secao_origem = request.form.get("secao_origem")
        secao_destino = request.form.get("secao_destino")
        chefia_origem = request.form.get("chefia_origem")
        chefia_destino = request.form.get("chefia_destino")

        quantidades_movimentadas = {}
        for key, value in request.form.items():
            if key.startswith("quantidade_"):
                bmp_key = key.split("_")[1]
                quantidades_movimentadas[bmp_key] = float(value) if value.strip() else 0

        if not (bmp_numbers and secao_origem and secao_destino and chefia_origem and chefia_destino):
            return render_template(
                "index.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Preencha todos os campos!",
            )

        bmp_list = [bmp.strip() for bmp in bmp_numbers.split(",")]
        dados_bmps = df[df["Nº BMP"].astype(str).isin(bmp_list)]
        if dados_bmps.empty:
            return render_template(
                "index.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Nenhum BMP encontrado para os números fornecidos.",
            )

        if not dados_bmps["CONTA"].eq("87 - MATERIAL DE CONSUMO DE USO DURADOURO").all():
            return render_template(
                "index.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Estes itens não pertencem à conta '87 - MATERIAL DE CONSUMO DE USO DURADOURO'."
            )

        dados_bmps["Qtde a Movimentar"] = dados_bmps["Nº BMP"].astype(str).map(quantidades_movimentadas).fillna(0)
        dados_bmps["Valor a Movimentar"] = dados_bmps.apply(
            lambda row: (row["VL. ATUALIZ."] / row["QTD"] * row["Qtde a Movimentar"]) if row["QTD"] > 0 else 0, axis=1
        )

        pdf = PDF()
        pdf.add_page()
        pdf.add_table(dados_bmps)
        pdf.add_details(secao_destino, chefia_origem, secao_origem, chefia_destino)

        output_path = "static/guia_circulacao_interna.pdf"
        pdf.output(output_path)
        return send_file(output_path, as_attachment=True)

    return render_template(
        "index.html", secoes_origem=secoes_origem, secoes_destino=secoes_destino
    )

@app.route("/autocomplete", methods=["POST"])
def autocomplete():
    data = request.get_json()
    bmp_numbers = data.get("bmp_numbers", [])

    if not bmp_numbers:
        return jsonify({"error": "Nenhum BMP fornecido!"}), 400

    response = {}
    for bmp in bmp_numbers:
        filtro_bmp = df[df["Nº BMP"].astype(str) == bmp]
        if not filtro_bmp.empty:
            secao_origem = filtro_bmp["Seção de Origem"].values[0]
            chefia_origem = filtro_bmp["Chefia de Origem"].values[0]
            response[bmp] = {
                "secao_origem": secao_origem,
                "chefia_origem": chefia_origem
            }
        else:
            response[bmp] = {"secao_origem": "", "chefia_origem": ""}

    return jsonify(response)

@app.route("/get_chefia", methods=["POST"])
def get_chefia():
    data = request.json
    secao = data.get("secao")
    tipo = data.get("tipo")

    if tipo == "destino":
        chefia = df[df['Seção de Destino'] == secao]['Chefia de Destino'].dropna().unique()
    elif tipo == "origem":
        chefia = df[df['Seção de Origem'] == secao]['Chefia de Origem'].dropna().unique()
    else:
        return jsonify({"error": "Tipo inválido!"}), 400

    return jsonify({"chefia": chefia.tolist()})
if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))  # Lê a variável PORT ou usa 5000 como padrão
    app.run(host="0.0.0.0", port=port)
