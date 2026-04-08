# ProjetoInventarioINSS

Automatizando auditorias de inventário com Python, PDF e Excel

---

## Sobre o Projeto

O **ProjetoInventarioINSS** foi desenvolvido com o objetivo de resolver um problema real:  
comparar patrimônios registrados em planilhas com aqueles encontrados em documentos PDF — mesmo quando esses PDFs estão desorganizados ou com orientação irregular.

Com este projeto, o processo manual e demorado de conferência de inventário é transformado em uma **análise automatizada, rápida e confiável**.

---

## Tecnologias Utilizadas

-  Python  
-  PyMuPDF (fitz)  
-  Pandas  
-  OpenPyXL  
-  Regex  

---

##  Estrutura do Projeto

ProjetoInventarioINSS/
├── config.py      # Configurações gerais
├── main.py        # Execução principal
├── pdf.py         # Manipulação de PDF
├── excel.py       # Manipulação de Excel
├── utils.py       # Funções auxiliares
│
├── archives/      # Arquivos de entrada (obrigatoriamente os arquivos a serem analisados devem estar aqui)
└── output/        # Arquivos gerados (local de destino dos arquivos analisados)
---

## Como Funciona

O sistema realiza:

### 1️. Leitura do Excel
- Carrega os patrimônios
- Remove duplicados e vazios

### 2️. Análise do PDF
- Percorre todas as páginas
- Extrai patrimônios automaticamente
- Funciona mesmo com PDF desorganizado

### 3️. Comparação
- Identifica divergências entre os dados

### 4️. Geração de Saída
- Excel atualizado com marcações  
- PDF com itens destacados  

---

## Como Executar

### 1. Instale as dependências

```bash
pip install pandas openpyxl pymupdf
````
### 2. Configure os arquivos editando o "config.py" com o nome da planilha e do arquivo 

PDF_FILE = "archives/ArquivosJuntos.pdf"

EXCEL_FILE = "archives/INVENTÁRIO.xlsx"

### 3. Execução
execute com:

```bash
python main.py
````
