# ProjetoInventarioINSS

Automatizando auditorias de inventário com Python, PDF e Excel

---

## Sobre o Projeto

O **ProjetoInventarioINSS** foi desenvolvido com o objetivo de solucionar uma necessidade encontrada durante o meu estágio na área de Tecnologia da Informação na Gerência Executiva do INSS de Santa Maria (GEXSTM).

Durante os processos de conferência de inventário, é necessário conferir se os equipamentos e bens encontrados nas agências realmente correspondem aos patrimônios registrados no sistema de controle de inventário do INSS. Essa conferência é realizada por meio de relatórios em PDF e planilhas eletrônicas, onde era feita a conferência em um processo totalmente manual, o que além de ser demorado também acabava sendo sucetível a erros.

Com o intuito de tornar essa atividade mais rápida e confiável, foi desenvolvido este sistema, que realiza automaticamente a leitura dos itens com patrimônio que foram encontrados na GEXSTM e registrados em uma planilha do Excel e os compara com os itens já cadastrados no inventário, informações essas contidas nos relatórios em PDF extraídos do sistema do INSS.

Ao final do processo, o sistema identifica divergências entre os dados, permitindo uma análise mais eficiente e reduzindo significativamente o tempo necessário para auditorias e conferências de inventário.

---

## Aplicação das Tecnologias no Projeto

### Python
Linguagem utilizada para o desenvolvimento do sistema.

### Visual Studio Code
IDE utilizada para desenvolvimento do sistema.

### Pandas
Biblioteca utilizada para a leitura e tratamento dos dados da planilha Excel contendo os patrimônios.

### PyMuPDF (fitz)
Biblioteca nativa utilizada no módulo `pdf.py` para abrir e percorrer todas as páginas do relatório em PDF, ela extrai as palavras por meio do comando `page.get_text("words")` e gera um PDF atualizado com os itens encontrados já destacados.

### OpenPyXL
Utilizado em `excel.py` para manipular diretamente a planilha original, aplicar a formatação nas células e gerar um novo arquivo Excel com as divergências identificadas devido a classe nativa `PatternFill`.

### os

Biblioteca nativa do Python utilizada em `main.py`, ela é responsável pela manipulação e criação das pastas com os arquivos novos e já renomeados.

### re (Regex)

Biiiblioteca utilizada em `utils.py` para localizar os números de patrimônio dentro dos textos presentes tanto no PDF quanto no Excel. A função `re.findall()` permite identificar automaticamente padrões de 9 dígitos definidos em `config.py`.

---

## Estrutura do Projeto

```
ProjetoInventarioINSS/
│
├── config.py   # Configurações gerais
├── main.py     # Execução principal
├── pdf.py      # Manipulação de PDF
├── excel.py    # Manipulação de Excel
├── utils.py    # Funções auxiliares
│
├── archives/   # Arquivos de entrada (obrigatoriamente os arquivos a serem analisados devem estar aqui)
└── output/     # Arquivos gerados (local de destino dos arquivos analisados)
```

---

## Como Funciona

O sistema realiza:

### 1️. Leitura do Excel
- Carrega os patrimônios da planilha selecionada
- Remove duplicados e vazios

### 2️. Análise do PDF
- Percorre todas as páginas
- Extrai patrimônios automaticamente
- Funciona mesmo com PDF desorganizado

### 3️. Comparação
- Identifica divergências entre os dados

### 4️. Geração de Saída
- Excel atualizado com as marcações em amarelo de itens que não foram encontrados no pdf  
- PDF com itens encontrados na planilha destacados em verde

---

## Como Executar

### 1. Instale as dependências

```bash
pip install pandas openpyxl pymupdf
````
### 2. Configure os arquivos editando o "config.py" com o nome dos itens que serão analisados

PDF_FILE = "archives/DIGITE _NOME_DO_PDF.pdf"

EXCEL_FILE = "archives/DIGITE_O_NOME_DO_ARQUIVO_EXCEL_AQUI.xlsx"

SHEET_NAME = "DIGITE AQUI O NOME DA ABA DA PLANILHA QUE SERÁ ANALISADA"

COLUNA_PATRIMONIO = "DIGITE AQUI O NOME DA COLUNA DA PLANILHA QUE SERÁ ANALISADA"

EXCEL_SAIDA = "output/DIGITE_O_NOME_DO_NOVO_ARQUIVO_EXCEL_AQUI.xlsx"

PDF_SAIDA = "output/DIGITE_O_NOME_DO_NOVO_ARQUIVO_PDF_AQUI.pdf"

### 3. Execução do programa
execute com:

```bash
python main.py
````
