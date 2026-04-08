# ==============================
# CONFIGURAÇÕES DO PROJETO
# ==============================

# Nome do arquivo PDF que será analisado
PDF_FILE = "archives/ArquivosJuntos.pdf"

# Nome do arquivo Excel com o inventário
EXCEL_FILE = "archives/INVENTÁRIO.xlsx"

# Nome da aba da planilha
SHEET_NAME = "GEXSTM"

# Nome da coluna que será analisada
COLUNA_PATRIMONIO = "PATROMÔNIOS"

# Novo arquivo de Excel com as divergências marcadas
EXCEL_SAIDA = "output/INVENTARIO_ATUALIZADO_2025.xlsx"

# Nome do arquivo PDF com os patrimônios marcados
PDF_SAIDA = "output/INVENTARIO_ATUALIZADO_2025.pdf"

# Expressão que identifica os patrimônios (9 números)
PATRIMONIO_REGEX = r"\b\d{9}\b"