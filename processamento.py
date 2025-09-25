# processamento.py
import os
import time
import pdfplumber
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from io import BytesIO
from werkzeug.utils import secure_filename

def gerar_pdf(caminho_entrada, bimestres):
    BASE_DIR = os.path.dirname(__file__)
    UPLOADS = os.path.join(BASE_DIR, "uploads")
    os.makedirs(UPLOADS, exist_ok=True)

    if bimestres is None:
        bimestre_str = "1"
    else:
        bimestre_str = str(bimestres).strip()

    mapping = {"1": 6, "2": 8, "3": 10, "4": 12}
    if bimestre_str not in mapping:
        raise ValueError(f"Bimestre inválido: {bimestres}")

    coluna_nota = mapping[bimestre_str]

    dados = []
    cabecalho_buffer = None
    with pdfplumber.open(caminho_entrada) as pdf:
        if not pdf.pages:
            raise ValueError("PDF sem páginas")
        try:
            first_page = pdf.pages[0]
            cabecalho_bbox = (0, 0, 841, 140)
            cabecalho_img = first_page.crop(cabecalho_bbox).to_image(resolution=300)
            cabecalho_buffer = BytesIO()
            cabecalho_img.save(cabecalho_buffer, format="PNG")
            cabecalho_buffer.seek(0)
        except Exception:
            cabecalho_buffer = None

        for page in pdf.pages:
            texto = page.extract_text() or ""
            tables = page.extract_tables() or []
            if not tables:
                continue
            table = tables[0]
            for row in table:
                if not row or len(row) <= coluna_nota:
                    continue
                matricula = (row[1] or "").strip()
                if not matricula.isdigit() or len(matricula) != 12:
                    continue
                status = (row[5] or "").strip().lower()
                nota_raw = (row[coluna_nota] or "").strip()
                if status == "cancelado":
                    continue
                if nota_raw == "-" or nota_raw == "":
                    continue
                nome = (row[2] or "").strip()
                dados.append([matricula, nome, nota_raw])

    df = pd.DataFrame(dados, columns=["Matrícula", "Nome", "Nota 1 Bim"])
    df["Nota 1 Bim"] = pd.to_numeric(
        df["Nota 1 Bim"].astype(str).str.replace(r"[^0-9,.\-]", "", regex=True).str.replace(",", "."),
        errors="coerce"
    )
    notas = df["Nota 1 Bim"].dropna()
    if notas.empty:
        media = mediana = desvio = nota_min = nota_max = 0.0
    else:
        media = notas.mean()
        mediana = notas.median()
        desvio = notas.std()
        nota_min = notas.min()
        nota_max = notas.max()

    stats_data = [
        ["Estatística", "Valor"],
        ["Média", f"{media:.2f}"],
        ["Mediana", f"{mediana:.2f}"],
        ["Desvio Padrão", f"{desvio:.2f}"],
        ["Nota Mínima", f"{nota_min:.2f}"],
        ["Nota Máxima", f"{nota_max:.2f}"]
    ]
    tabela_estatisticas = Table(stats_data, colWidths=[200, 100])
    estilo_tabela = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 1, colors.lightslategray),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ])
    tabela_estatisticas.setStyle(estilo_tabela)

    imagens = []
    buf = BytesIO()
    plt.figure()
    plt.hist(notas, bins=10, edgecolor='black')
    plt.title("Distribuição de Notas")
    plt.xlabel("Nota")
    plt.ylabel("Quantidade de Alunos")
    plt.grid(True)
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close()
    buf.seek(0)
    imagens.append(buf)

    df_sorted = df.sort_values("Nota 1 Bim")
    buf = BytesIO()
    plt.figure(figsize=(14, 6))
    cores_barras = df_sorted["Nota 1 Bim"].apply(lambda x: 'green' if pd.notna(x) and x >= 70 else 'red').tolist()
    plt.bar(df_sorted["Nome"], df_sorted["Nota 1 Bim"], color=cores_barras)
    plt.title("Notas dos Alunos (Ordenadas)", fontsize=16, weight='bold')
    plt.xlabel("Aluno", fontsize=12)
    plt.ylabel("Nota 1 Bim", fontsize=12)
    plt.xticks(rotation=90)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close()
    buf.seek(0)
    imagens.append(buf)

    situacao = {
        "Ruim (0-39)": int(df[df["Nota 1 Bim"] < 40].shape[0]),
        "Mediana (40-69)": int(df[(df["Nota 1 Bim"] >= 40) & (df["Nota 1 Bim"] < 70)].shape[0]),
        "Bom (70-85)": int(df[(df["Nota 1 Bim"] >= 70) & (df["Nota 1 Bim"] <= 85)].shape[0]),
        "Excelente (86-100)": int(df[df["Nota 1 Bim"] > 85].shape[0])
    }
    labels = list(situacao.keys())
    sizes = list(situacao.values())
    buf = BytesIO()
    plt.figure(figsize=(8, 8))
    plt.pie(sizes, labels=labels, startangle=90, wedgeprops=dict(width=0.4))
    plt.title("Situação dos Alunos", fontsize=16, weight='bold')
    plt.axis('equal')
    plt.savefig(buf, format='png', bbox_inches='tight')
    plt.close()
    buf.seek(0)
    imagens.append(buf)

    abaixo_70_df = df[df["Nota 1 Bim"] < 70].copy()
    abaixo_40 = abaixo_70_df[abaixo_70_df["Nota 1 Bim"] < 40].copy()
    entre_40_70 = abaixo_70_df[(abaixo_70_df["Nota 1 Bim"] >= 40) & (abaixo_70_df["Nota 1 Bim"] < 70)].copy()
    abaixo_40.insert(0, "#", range(1, len(abaixo_40) + 1))
    entre_40_70.insert(0, "#", range(1, len(entre_40_70) + 1))

    tabela_abaixo_40 = Table([["#", "Matrícula", "Nome", "Nota"]] + abaixo_40.values.tolist(), colWidths=[30, 90, 140, 70])
    tabela_entre_40_70 = Table([["#", "Matrícula", "Nome", "Nota"]] + entre_40_70.values.tolist(), colWidths=[30, 90, 140, 70])
    tabela_abaixo_40.setStyle(estilo_tabela)
    tabela_entre_40_70.setStyle(estilo_tabela)

    nome_base = secure_filename(os.path.basename(caminho_entrada))
    ts = int(time.time())
    nome_saida = f"{os.path.splitext(nome_base)[0]}_relatorio_bim{bimestre_str}_{ts}.pdf"
    caminho_saida = os.path.join(UPLOADS, nome_saida)

    pdf = canvas.Canvas(caminho_saida, pagesize=landscape(A4))
    largura, altura = landscape(A4)

    if cabecalho_buffer:
        try:
            pdf.drawImage(ImageReader(cabecalho_buffer), 0, altura - 140, width=largura, height=140)
        except Exception:
            pass

    pdf.rect(30, altura - 200, 381, 12)
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(60, altura - 197, f"INFORMAÇÕES GERAIS - BIMESTRE {bimestre_str}")
    w, h = tabela_estatisticas.wrap(300, 120)
    tabela_estatisticas.drawOn(pdf, 30, altura - 230 - h)
    pdf.showPage()

    pdf.rect(30, altura - 45, 381, 12)
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(145, altura - 43, "ALUNOS ENTRE 40 E 70")
    w_e, h_e = tabela_entre_40_70.wrap(381, altura)
    tabela_entre_40_70.drawOn(pdf, 30, altura - 60 - h_e)

    pdf.rect(431, altura - 45, 381, 12)
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(largura/2 + 135, altura - 43, "ALUNOS ABAIXO DE 40")
    w_a, h_a = tabela_abaixo_40.wrap(381, altura)
    tabela_abaixo_40.drawOn(pdf, 431, altura - 60 - h_a)

    pdf.showPage()

    imagens_por_pagina = 2
    largura_img = largura * 0.8
    altura_img = altura / imagens_por_pagina * 0.8

    for i in range(0, len(imagens), imagens_por_pagina):
        imagens_pagina = imagens[i:i+imagens_por_pagina]
        for j, img in enumerate(imagens_pagina):
            img.seek(0)
            imagem = ImageReader(img)
            iw, ih = imagem.getSize()
            escala = min(largura_img / iw, altura_img / ih)
            iw = iw * escala
            ih = ih * escala
            x = (largura - iw) / 2
            y = altura - (j + 1) * (altura / imagens_por_pagina) + ((altura / imagens_por_pagina - ih) / 2)
            pdf.drawImage(imagem, x, y, width=iw, height=ih)
        pdf.showPage()

    pdf.save()
    return caminho_saida