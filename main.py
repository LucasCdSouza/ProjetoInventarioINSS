import os
from excel import carregar_patrimonios_excel, marcar_excel
from pdf import extrair_patrimonios_pdf, marcar_pdf
from config import *


def main():

    # Cria pasta de saída
    os.makedirs("output", exist_ok=True)

    print("=" * 40)
    print("AUDITORIA DE INVENTÁRIO")
    print("=" * 40)

    # Ler Excel
    print("\n[1/4] Lendo Excel...")
    patrimonios_excel = carregar_patrimonios_excel(
        EXCEL_FILE,
        SHEET_NAME,
        COLUNA_PATRIMONIO
    )

    print(f"{len(patrimonios_excel)} patrimônios carregados.")

    # Ler PDF
    print("\n[2/4] Escaneando PDF...")
    patrimonios_pdf = extrair_patrimonios_pdf(PDF_FILE)

    print(f"{len(patrimonios_pdf)} patrimônios encontrados no PDF.")

    # Marcar Excel
    print("\n[3/4] Marcando Excel...")
    marcar_excel(
        EXCEL_FILE,
        SHEET_NAME,
        COLUNA_PATRIMONIO,
        patrimonios_pdf,
        EXCEL_SAIDA
    )

    print(f"Excel salvo em: {EXCEL_SAIDA}")

    # Marcar PDF
    print("\n[4/4] Marcando PDF...")
    marcar_pdf(
        PDF_FILE,
        patrimonios_excel,
        PDF_SAIDA
    )

    print(f"PDF salvo em: {PDF_SAIDA}")
    print("\nProcesso finalizado com sucesso.")


if __name__ == "__main__":
    main()