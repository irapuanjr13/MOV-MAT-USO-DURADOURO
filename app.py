from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from fpdf import FPDF
import gdown

app = Flask(__name__)

# ID do arquivo no Google Drive
GOOGLE_DRIVE_FILE_ID = "1mPYlc_uC3SfJnNQ_ToG6eVmn2ZYMhPCX"

# Função para baixar e carregar a planilha
def get_excel_from_google_drive():
    url = f"https://drive.google.com/uc?id={GOOGLE_DRIVE_FILE_ID}"
    output_file = "patrimonio.xlsx"
    gdown.download(url, output_file, quiet=False)
    return pd.read_excel(output_file)

# Carregar a planilha no início do programa
df = get_excel_from_google_drive()

class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 6, "MINISTÉRIO DA DEFESA", ln=True, align="C")
        self.cell(0, 6, "COMANDO DA AERONÁUTICA", ln=True, align="C")
        self.cell(0, 6, "GRUPAMENTO DE APOIO DE LAGOA SANTA", ln=True, align="C")
        self.cell(0, 8, "GUIA DE MOVIMENTAÇÃO DE BEM DE USO DURADOURO ENTRE AS SEÇÕES DO GAPLS", ln=True, align="C")
        self.ln(10)

    def add_table(self, dados_bmps):
        col_widths = [15, 55, 15, 30, 30, 30]
        headers = ["Nº BMP", "Nomenclatura", "Qtde", "Valor Atualizado", "Qtde a Movimentar", "Valor a Movimentar"]

        # Renderizar cabeçalhos
        self.set_font("Arial", "B", 9)
        for header, width in zip(headers, col_widths):
            self.cell(width, 10, header, border=1, align="C")
        self.ln()

        # Adicionar dados
        self.set_font("Arial", size=9)
        for _, row in dados_bmps.iterrows():
            self.cell(col_widths[0], 10, str(row["Nº BMP"]), border=1, align="C")
            self.cell(col_widths[1], 10, str(row["NOMECLATURA/COMPONENTE"]), border=1, align="L")
            self.cell(col_widths[2], 10, str(row["QTD"]), border=1, align="C")
            self.cell(col_widths[3], 10, f"R$ {row['VL. ATUALIZ.']:.2f}".replace(".", ","), border=1, align="C")
            self.cell(col_widths[4], 10, str(row["Qtde a Movimentar"]), border=1, align="C")
            self.cell(col_widths[5], 10, f"R$ {row['Valor a Movimentar']:.2f}".replace(".", ","), border=1, align="C")
            self.ln()

    def add_details(self, secao_destino, chefia_origem, secao_origem, chefia_destino):
        self.set_font("Arial", size=12)
        self.ln(10)
        text = f"""
Solicitação de Transferência:
Informo à Senhora Chefe do GAP-LS que os bens especificados estão inservíveis para uso neste setor, classificados como ociosos, recuperáveis, reparados ou novos - aguardando distribuição. Diante disso, solicito autorização para transferir o(s) Bem(ns) Móvel(is) Permanente(s) acima discriminado(s), atualmente sob minha guarda, para a Seção {secao_destino}.

{chefia_origem}
{secao_origem}

Confirmação da Seção de Destino:
Estou ciente da movimentação informada acima e, devido à necessidade do setor, solicito à Senhora Dirigente Máximo autorização para manter sob minha guarda os Bens Móveis Permanentes especificados.

{chefia_destino}
{secao_destino}

DO AGENTE DE CONTROLE INTERNO AO DIRIGENTE MÁXIMO:
Informo à Senhora que, após conferência, foi verificado que esta guia cumpre o disposto no Módulo D do RADA-e e, conforme a alínea "d" do item 5.3 da ICA 179-1, encaminho para apreciação e se for o caso, autorização.

KARINA RAQUEL VALIMAREANU  Maj Int
Chefe da ACI

DESPACHO DA AGENTE DIRETOR:
Autorizo a movimentação solicitada e determino:
1. Que a Seção de Registro realize a movimentação no SILOMS.
2. Que a Seção de Registro publique a movimentação no próximo aditamento a ser confeccionado, conforme o item 2.14.2, Módulo do RADA-e.
3. Que os detentores realizem a movimentação física do(s) bem(ns).

LUCIANA DO AMARAL CORREA  Cel Int
Dirigente Máximo
"""
        self.multi_cell(0, 8, text)

@app.route("/", methods=["GET", "POST"])
def index():
    secoes_origem = df["Seção de Origem"].dropna().unique().tolist()
    secoes_destino = df["Seção de Destino"].dropna().unique().tolist()

    if request.method == "POST":
        bmp_numbers = request.form.get("bmp_numbers", "").strip()
        secao_origem = request.form.get("secao_origem")
        secao_destino = request.form.get("secao_destino")
        chefia_origem = request.form.get("chefia_origem")
        chefia_destino = request.form.get("chefia_destino")

        quantidades_movimentadas = {}
        for key, value in request.form.items():
            if key.startswith("quantidade_"):
                bmp_key = key.split("_")[1]
                quantidades_movimentadas[bmp_key] = float(value) if value.strip() else 0

        bmp_list = [bmp.strip() for bmp in bmp_numbers.split(",") if bmp.strip()]
        dados_bmps = df[df["Nº BMP"].astype(str).isin(bmp_list)]

        if dados_bmps.empty:
            return render_template(
                "index.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Nenhum BMP encontrado ou inválido."
            )

        if not dados_bmps["CONTA"].eq("87 - MATERIAL DE CONSUMO DE USO DURADOURO").all():
            return render_template(
                "index.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Os itens não pertencem à conta '87 - MATERIAL DE CONSUMO DE USO DURADOURO'."
            )

        # Cálculo de valores
        dados_bmps["Qtde a Movimentar"] = dados_bmps["Nº BMP"].astype(str).map(quantidades_movimentadas).fillna(0)
        dados_bmps["Valor a Movimentar"] = dados_bmps.apply(
            lambda row: (row["VL. ATUALIZ."] / row["QTD"] * row["Qtde a Movimentar"]) if row["QTD"] > 0 else 0, axis=1
        )

        pdf = PDF()
        pdf.add_page()
        pdf.add_table(dados_bmps)
        pdf.add_details(secao_destino, chefia_origem, secao_origem, chefia_destino)  # Adicionando os detalhes ao PDF

        output_path = "static/guia_circulacao_interna.pdf"
        pdf.output(output_path)
        return send_file(output_path, as_attachment=True)

    return render_template(
        "index.html", secoes_origem=secoes_origem, secoes_destino=secoes_destino
    )

@app.route("/get_chefia", methods=["POST"])
def get_chefia():
    data = request.json
    secao = data.get("secao")
    tipo = data.get("tipo")

    if tipo == "origem":
        chefia = df[df['Seção de Origem'] == secao]['Chefia de Origem'].dropna().unique()
    elif tipo == "destino":
        chefia = df[df['Seção de Destino'] == secao]['Chefia de Destino'].dropna().unique()
    else:
        return jsonify({"error": "Tipo inválido"}), 400

    return jsonify({"chefia": chefia[0] if len(chefia) > 0 else ""})

if __name__ == "__main__":
    app.run(debug=True, port=5001)