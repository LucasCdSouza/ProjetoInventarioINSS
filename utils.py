# Biblioteca para trabalhar com padrões de texto (regex)
import re

# Importa o padrão definido no config.py para identificar os números de patrimônio
from config import PATRIMONIO_REGEX


# ==================================
# FUNÇÃO: Extrair patrimônios
# ==================================
# Procura números de patrimônio dentro de um texto
def extrair_patrimonios(texto):

    # Encontra todos os números de 9 dígitos
    encontrados = re.findall(PATRIMONIO_REGEX, texto)

    return encontrados