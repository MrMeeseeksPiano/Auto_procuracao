from docxtpl import DocxTemplate
import streamlit as st
import pdfplumber
import re
from datetime import datetime
import io

st.title("Gerador de Procuração")

#arquivo_pdf = r'/Users/alanfontenele/Downloads/Documentos /UnB/Auto_procuracao/Procuracoes/Cópia de Contrato Vanderlei Alves de Sousa.pdf'
arquivo_pdf = st.file_uploader("Upload do contrato em PDF",type="pdf")

if arquivo_pdf:

    with pdfplumber.open(arquivo_pdf) as pdf:
        first_page = pdf.pages[0]
        pdf_text = first_page.extract_text()

    pdf_text = " ".join(pdf_text.split())
    dados = {}
    nome = re.search(r"(?<=CONTRATANTE:\s)(.*?)(?=,)", pdf_text).group(0)
    dados['nome'] = nome
    primeiro_nome = nome.split()[0]
    cpf = re.search(r"(\d{3}\.\d{3}\.\d{3}-\d{2})", pdf_text).group(0)
    dados['cpf'] = cpf
    #endereco = re.search(r"residente\s+e\s+domiciliado\s+em\s+(.*?CEP\s+\d{5}-\d{3})", pdf_text, re.DOTALL)
    dados['endereco'] = " ".join(m.group(1).split()).strip(", ") if (m := re.search(r"domiciliado\s+em\s+(.*?)\s*CEP", pdf_text, re.DOTALL | re.IGNORECASE)) else ""

    meses = {
        1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril',
        5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto',
        9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
    }

    hoje = datetime.now()
    #print(hoje)
    data_formatada = f"{hoje.day:02d} de {meses[hoje.month]} de {hoje.year}"
    dados['data'] = data_formatada

    if st.button("Gerar Procuração"):
        doc = DocxTemplate("ProcuraçãoTemplate.docx")
        doc.render(dados)

        output_stream = io.BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)

        st.download_button(
            label="Baixar Procuração",
            data=output_stream,
            file_name=f"Procuração - {primeiro_nome}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )




    #doc.save(f"Procuração {nome}.docx")



