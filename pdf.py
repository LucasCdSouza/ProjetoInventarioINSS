import fitz
from utils import extrair_patrimonios

# Evita mensagens de erro do MuPDF no terminal
fitz.TOOLS.mupdf_display_errors(False)


# ==================================
# FUNÇÃO: Extrair patrimônios do PDF
# ==================================
def extrair_patrimonios_pdf(pdf_file):

    doc = fitz.open(pdf_file)
    encontrados = set()

    total_paginas = len(doc)

    for i, page in enumerate(doc):

        print(f"Página {i + 1} de {total_paginas}", end="\r")

        try:
            words = page.get_text("words")
        except:
            continue

        for w in words:
            texto = w[4]

            patrimonios = extrair_patrimonios(texto)

            for p in patrimonios:
                encontrados.add(p)

    doc.close()
    print()
    return encontrados


# ==================================
# FUNÇÃO: Marcar PDF
# ==================================
def marcar_pdf(pdf_file, set_excel, pdf_saida):

    doc = fitz.open(pdf_file)
    cor_verde = (0, 1, 0)

    for page in doc:

        crop = page.cropbox
        words = page.get_text("words")

        linhas_processadas = set()

        for w in words:
            x0, y0, x1, y1, texto = w[:5]

            patrimonios = extrair_patrimonios(texto)

            for patrimonio in patrimonios:

                if patrimonio in set_excel:

                    width_word = x1 - x0
                    height_word = y1 - y0

                    rect_final = None

                    # Caso horizontal
                    if width_word > height_word:
                        y_ref = round(y0, 1)

                        if y_ref not in linhas_processadas:
                            rect_final = fitz.Rect(crop.x0, y0 - 1, crop.x1, y1 + 1)
                            linhas_processadas.add(y_ref)

                    # Caso vertical
                    else:
                        x_ref = round(x0, 1)

                        if x_ref not in linhas_processadas:
                            rect_final = fitz.Rect(x0 - 1, crop.y0, x1 + 1, crop.y1)
                            linhas_processadas.add(x_ref)

                    if rect_final:
                        page.draw_rect(
                            rect_final,
                            color=None,
                            fill=cor_verde,
                            fill_opacity=0.35
                        )

    doc.save(pdf_saida)
    doc.close()