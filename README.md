# Excel Dynamic Data Validation with OpenPyXL

A Python utility that automates the creation of dropdown lists (**Data Validation**) in Excel workbooks using `openpyxl`. It dynamically maps columns between a target worksheet and a reference worksheet by matching their header names, automatically calculating the exact row range for each list.

## Features

- **Header-Based Matching:** Automatically links columns in the main sheet to their corresponding validation lists in a secondary sheet by column name, regardless of column order.
- **Dynamic Range Detection:** Finds the last non-empty row for each individual list column so dropdowns don't include trailing blank cells.
- **Custom Error & Input Prompts:** Configures user-friendly validation prompts and error alerts for each column.
- **Safe Re-runs:** Clears previous data validations on the target sheet before applying updated rules to avoid duplicates.

## Requirements

- Python 3.8+
- `openpyxl`

```bash
pip install openpyxl
```

## Usage

Run `main.py` pointing to your Excel file, specifying the target sheet name and the sheet containing the validation lists:

```python
from main import aplicar_validacao_por_coluna

aplicar_validacao_por_coluna(
    caminho_arquivo="teste.xlsx",
    nome_aba_principal="Cadastro",
    nome_aba_listas="Listas",
    linha_cabecalho=1,
    linha_max_aplicacao=500,
)
```
