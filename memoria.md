# 🧠 Memória do Projeto: Pipeline de Extração e Auditoria de Formulários PDF (AcroForms to Excel)

**Data de Conclusão:** 2026-10-09  
**Status:** ✅ Concluído / Pronto para Produção  
**Tags:** `python`, `pypdf`, `pymupdf`, `pandas`, `openpyxl`, `pillow`, `acroforms`, `pdf-extraction`, `data-automation`, `office-365`, `batch-processing`

---

## 1. 🎯 Objetivo do Projeto
Desenvolver uma esteira de automação industrial em Python para processamento em lote (*batch processing*) de mais de 50 formulários PDF preenchíveis corporativos (*AcroForms / ISO 32000*), garantindo 100% de acurácia na extração de:
1. **Respostas digitadas (*typed answers*)**: Strings textuais, identificadores funcionais, dados cadastrais e remunerações.
2. **Caixas de seleção (*ticked boxes*)**: Normalização determinística de estados de consentimento e conformidade (`Selected` vs `Not Selected`).
3. **Assinaturas embutidas (*embedded signatures*)**: Identificação de objetos gráficos binários de assinatura, extração em alta resolução, salvamento em disco e cálculo de hash Base64 para auditoria.
4. **Entregável Executivo**: Consolidação de todos os registros em uma única planilha Microsoft Excel (`.xlsx`), com cabeçalhos congelados, larguras autoajustadas, formatação condicional suave e hiperlinks nativos clicáveis para os PDFs e imagens extraídas.

---

## 2. ⚡ Problema Combatido
* **Falha de ferramentas tradicionais de OCR / Text Parsing**: Bibliotecas comuns como `pdfplumber` ou parsers de texto puro frequentemente falham ou sofrem com perda de dados em AcroForms, pois os valores preenchidos não residem na camada de texto desenhada da página, mas sim na árvore interna de dicionários `/Root -> /AcroForm -> /Fields -> /V`.
* **Inconsistência de Checkboxes**: Cada plataforma PDF (Adobe, DocuSign, Jotform) armazena checkboxes de forma distinta (`/Yes`, `/1`, `/Off`, `/0`, `/AS`). Sem uma camada de normalização inteligente, a análise posterior gera dados corrompidos ou ambíguos.
* **Perda de Rastreabilidade de Assinaturas**: Softwares convencionais ignoram assinaturas manuais embutidas ou requerem ferramentas pagas pesadas. A automação precisava extrair o artefato gráfico isolado sem renderizar o PDF inteiro como imagem (o que degradaria performance e consumiria memória excessiva).

---

## 3. 🛠️ Tech Stack Utilizada

| Tecnologia / Lib | Versão | Papel no Pipeline | Justificativa de Engenharia |
| :--- | :--- | :--- | :--- |
| **Python** | 3.12+ | Linguagem Base | Tipagem estática (`typing`), POO limpa e modularidade. |
| **`pypdf`** | >=4.3.0 | Leitura de AcroForms | Acesso nativo à árvore lógica de formulários PDF sem dependências externas em C. |
| **`pymupdf` (MuPDF)** | >=1.24.0 | Extração Gráfica de Baixo Nível | Motor ultraveloz em C/C++ para inspecionar a tabela `/XObject /Image`, extrair bytes de assinaturas por `xref` e renderizar testes. |
| **`pandas`** | >=2.2.0 | Consolidação Tabular | Normalização de schema, alinhamento de colunas e criação do `DataFrame` bidimensional em milissegundos. |
| **`openpyxl`** | >=3.1.2 | Estilização Avançada OpenXML | Controle do arquivo `.xlsx`, injeção de hiperlinks URI, formatação condicional e *freeze panes*. |
| **`Pillow` (PIL)** | >=10.0.0 | Manipulação Gráfica | Validação de formatos, dimensões de imagem e geração de grafismos de assinatura. |

---

## 4. 🔄 Pipeline de Execução do Software

```text
[Diretório Raiz: *.pdf]
          │
          ▼
1. Descoberta & Triagem (Path.glob("*.pdf") + sort alfabético determinístico)
          │
          ▼
2. Loop de Extração Híbrida (PDFFormExtractor):
   ├── A. Acesso a /AcroForm -> Lê chaves /V com fallback dinâmico ("Zero-Config")
   ├── B. Normalização de dados (Limpeza de "/" e padronização binária de checkboxes)
   └── C. Inspeção de Página (/XObject) -> Extrai bytes da imagem da assinatura (PNG/JPEG)
          ├── Salva arquivo atômico em output/signatures/{stem}_signature.{ext}
          └── Gera string Base64 em memória para validação criptográfica
          │
          ▼
3. Consolidação Tabular (pd.DataFrame(records))
          │
          ▼
4. Estilização Executiva (ExcelBuilder com openpyxl):
   ├── Cabeçalho Navy (#1B365D) com texto branco em negrito e altura 28pt
   ├── Congelamento de painel em A2 (Freeze Panes)
   ├── Formatação condicional: Verde (#E6F4EA) para Selected, Vermelho (#FCE8E6) para Not Selected
   ├── Criação de Hiperlinks nativos para os PDFs e imagens de assinatura
   └── Auto-fit inteligente de largura: min(max(len + 4, 14), 45) com trava para Base64
          │
          ▼
[Entregável Final: output/Extracted_Form.xlsx + output/signatures/*.png]
```

---

## 5. 🧱 Arquitetura dos Módulos

* **`config.py`**: Centraliza variáveis de ambiente, caminhos de I/O, mapeamento opcional de campos (`FIELD_MAPPING`) e nomes oficiais das colunas do sistema (`SYSTEM_COLUMNS`).
* **`extractor.py`**: Motor de extração encapsulado na classe `PDFFormExtractor`. Implementa tratamento defensivo de exceções e descoberta dinâmica de títulos em *Title Case*.
* **`excel_builder.py`**: Motor de pós-produção na classe `ExcelBuilder`. Transforma dados tabulares brutos em uma planilha corporativa pronta para decisão gerencial.
* **`main.py`**: Ponto de entrada CLI com medição de latência em tempo de execução, feedback visual em linha única (`end=""`) e retorno de códigos de erro de sistema operacional (`sys.exit(1)`).
* **`pdf_generator/`**: Submódulo de fábrica de dados sintéticos (*Mocking Factory*). Capaz de sintetizar formulários AcroForms com grid alinhado, campos interativos e assinaturas reais para testes de regressão antes do recebimento da carga real do cliente.

---

## 6. 🧗 Desafios Encontrados & Soluções Aplicadas

1. **Descompasso de Indexação Excel vs Python:**
   * *Desafio:* No Excel/OpenPyXL as colunas iniciam em 1 (`A=1`), enquanto no Pandas as colunas iniciam em 0.
   * *Solução:* Mapeamento matemático `df.columns[col_idx - 1]`, garantindo que cada célula receba os metadados corretos de sua respectiva coluna.
2. **Truncamento Visual e Inchaço por Base64:**
   * *Desafio:* A fórmula de autoajuste de colunas expandiria a coluna de Base64 para milhares de caracteres de largura, tornando o Excel ilegível.
   * *Solução:* Cláusula de guarda que identifica `config.SYSTEM_COLUMNS["signature_base64"]` e fixa a largura máxima em 25 caracteres.
3. **Formulários Pendentes de Assinatura:**
   * *Desafio:* Formulários que chegam sem assinatura embutida poderiam quebrar o pipeline por `NoneType` ou variáveis não atribuídas.
   * *Solução:* Retorno seguro de tripla `(sig_detected, sig_hyperlink, sig_base64_str)` com valores de fallback `"NO"` e `"N/A"`.

---

## 7. 💡 Principais Lições Aprendidas
* **Acesso Nativo a Especificações de Arquivo:** Ao lidar com PDFs corporativos, entender a especificação ISO 32000 e consultar `/AcroForm` diretamente é infinitamente mais veloz, confiável e determinístico do que recorrer a soluções caras ou instáveis de OCR.
* **Separar Camada de Dados de Camada de Apresentação:** Utilizar o Pandas para estruturar a matriz de dados e delegar a estilização fina ao OpenPyXL combina o melhor dos dois mundos: alta velocidade de processamento com qualidade estética executiva.
* **Construção de Ambiente de Testes Próprio (*Test Harness*):** A capacidade de gerar seus próprios dados de teste idênticos aos do cliente permite entregar a automação já testada, validada e funcional antes mesmo de receber os arquivos de produção.
