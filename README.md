# 📄 Automated Fillable PDF to Excel Extraction Pipeline

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)]()
[![Excel](https://img.shields.io/badge/Office%20365-Excel%20.XLSX-success.svg)]()

> Esteira automatizada em Python para processamento em lote (*batch processing*) de formulários PDF preenchíveis (**AcroForms / ISO 32000**), com extração de alta precisão de respostas digitadas, caixas de seleção normalizadas e assinaturas embutidas com rastreabilidade para auditoria.

---

## 🎯 Por que esta solução?

Muitas ferramentas tentam extrair formulários via **OCR ou parsers visuais de texto puro**, o que frequentemente causa:
* Perda de dados digitados (pois o texto reside na árvore `/AcroForm -> /Fields` e não na camada vetorial desenhada).
* Checkboxes ignorados ou interpretados como caracteres corrompidos.
* Assinaturas manuscritas ignoradas ou cortadas.

Esta ferramenta acessa **diretamente a árvore de objetos nativos do PDF**, garantindo **100% de precisão tabular**, processamento em menos de 100ms por documento e zero intervenção manual.

---

## 🔄 Workflow do Sistema

```text
               [ Lote de PDFs Preenchíveis (*.pdf) ]
                                 │
                                 ▼
                     1. main.py (Orquestrador)
         Varre a pasta de entrada e despacha cada documento
                                 │
                                 ▼
               2. extractor.py (Motor de Extração Híbrido)
       ┌─────────────────────────┼─────────────────────────┐
       ▼                         ▼                         ▼
 [Campos de Texto]      [Caixas de Seleção]      [Assinaturas Embutidas]
  Lê chave /V com        Normaliza /Yes e /Off    Isola objeto XObject /Image
  descoberta dinâmica    em 'Selected' e          Salva PNG em /signatures
  de títulos em          'Not Selected'           e gera string Base64
  Title Case                                      para auditoria
       └─────────────────────────┬─────────────────────────┘
                                 │
                                 ▼
                 3. excel_builder.py (Pós-Produção)
  Compila dados tabulares via Pandas e estiliza via OpenPyXL:
  • Cabeçalho corporativo Navy (#1B365D) e painéis congelados (A2)
  • Formatação condicional para checkboxes (verde/vermelho suave)
  • Auto-fit inteligente de largura de colunas
  • Hiperlinks nativos clicáveis para o PDF e para a imagem da assinatura
                                 │
                                 ▼
         [ output/Extracted_Form.xlsx ] + [ output/signatures/*.png ]
```

---

## ✨ Principais Funcionalidades

* **Zero-Config Dinâmico:** Se um novo formulário chegar com novos campos, o sistema descobre os nomes das colunas e formata os títulos automaticamente (*Title Case*), sem necessidade de reconfiguração manual.
* **Normalização Confiável de Checkboxes:** Converte automaticamente múltiplos padrões do mercado (`/Yes`, `/1`, `/Off`, `/0`, etc.) em `Selected` e `Not Selected`.
* **Auditoria Completa de Assinaturas:** Extrai assinaturas gráficas embutidas na página sem renderizar o documento inteiro, salvando o arquivo original na pasta de auditoria e fornecendo o hash Base64 na planilha.
* **Planilha Executiva Pronta para Tomada de Decisão:** Gera uma planilha `.xlsx` profissional estilizada com links que abrem o PDF de origem e a imagem da assinatura com apenas um clique.
* **Fábrica de Testes Embutida (`pdf_generator/`):** Acompanha um submódulo capaz de sintetizar formulários corporativos preenchidos e assinados para testes antes de receber a carga real.

---

## 📁 Estrutura de Pastas do Projeto

```text
PDF_Form_Data_Extraction/
├── sample_form_01.pdf             # PDFs de entrada na raiz
├── sample_form_02.pdf
├── sample_form_03.pdf
├── sample_form_04.pdf
│
├── config.py                      # Configurações globais e mapeamentos
├── extractor.py                   # Motor de extração (pypdf + PyMuPDF)
├── excel_builder.py               # Construtor e estilização do Excel (openpyxl)
├── main.py                        # Ponto de entrada CLI (executável)
├── requirements.txt               # Dependências do projeto
│
├── pdf_generator/                 # Fábrica de formulários sintéticos de teste
│   ├── create_test_forms.py
│   └── generate_signature.py
│
└── output/                        # Gerada automaticamente na execução
    ├── Extracted_Form.xlsx        # Planilha final consolidada
    └── signatures/                # Imagens das assinaturas extraídas (*.png / *.jpeg)
```

---

## 🚀 Como Utilizar (Passo a Passo)

### 1. Clonar o Repositório e Instalar Dependências

Certifique-se de ter o **Python 3.10 ou superior** instalado:

```powershell
# Instale as bibliotecas necessárias
pip install -r requirements.txt
```

### 2. Adicionar os PDFs para Processamento

Coloque todos os arquivos `.pdf` que deseja processar na raiz do projeto (ou no diretório configurado em `INPUT_PDF_DIR`).

> 💡 **Dica de Teste:** Se ainda não possui PDFs do seu cliente, execute o gerador incluído para criar formulários de teste realistas:
> ```powershell
> python pdf_generator\create_test_forms.py
> ```

### 3. Executar o Pipeline

Basta rodar o arquivo principal no seu terminal:

```powershell
python main.py
```

Você verá a saída em tempo real no console:

```text
=================================================================
  AUTOMATED FILLABLE PDF TO EXCEL EXTRACTION PIPELINE
=================================================================
Search Directory: C:\Projetos\PDF_Form_Data_Extraction
Total PDF Forms: 4
=================================================================
[001/004] Processing: sample_form_01.pdf... FINISHED
[002/004] Processing: sample_form_02.pdf... FINISHED
[003/004] Processing: sample_form_03.pdf... FINISHED
[004/004] Processing: sample_form_04.pdf... FINISHED
=================================================================
Generating Stylized Excel Worksheet and Linking Signatures...
SUCCESS! Worksheet successfully generated with 4 records:
----------Signatures saved in: C:\Projetos\output\signatures
Total Generating Time: 0.42 seconds
=================================================================
```

### 4. Consultar os Entregáveis

Abra a pasta `output/` gerada pelo script:
* **`output/Extracted_Form.xlsx`**: Sua planilha consolidada com todas as colunas, formatação corporativa e hiperlinks ativos.
* **`output/signatures/`**: Imagens isoladas de cada assinatura extraída dos documentos.

---

## ⚙️ Configurações Personalizadas (`config.py`)

No arquivo `config.py`, você pode ajustar o comportamento da automação conforme a necessidade do seu projeto:

```python
# Modo de captura de assinaturas:
# "hyperlink" -> Salva a imagem em disco e insere um link clicável na célula
# "base64"    -> Gera apenas a string Base64 dentro da célula
# "both"      -> Salva o arquivo, insere o link E guarda a string Base64 (Padrão)
SIGNATURE_CAPTURE_MODE = "both"

# Mapeamento explícito opcional (se vazio {}, descobre os nomes sozinhos):
FIELD_MAPPING = {
    # "nome_campo_no_pdf": "Nome Bonito na Coluna do Excel"
}
```

---

## 🛠️ Tecnologias Utilizadas

| Stack | Função |
| :--- | :--- |
| **Python 3.12** | Linguagem principal do projeto |
| **`pypdf`** | Leitura e navegação estruturada na árvore AcroForms |
| **`pymupdf` (MuPDF)** | Extração de objetos gráficos binários de baixo nível (XObject /Image) |
| **`pandas`** | Manipulação, estruturação e consolidação dos dados em matriz tabular |
| **`openpyxl`** | Criação, estilização, formatação condicional e hiperlinks no Excel |
| **`Pillow`** | Validação e suporte a processamento de imagens |

---

## 📄 Licença

Distribuído sob a licença MIT. Consulte o arquivo `LICENSE` para mais detalhes.
