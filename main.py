from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.worksheet.datavalidation import DataValidation


def aplicar_validacao_por_coluna(
    caminho_arquivo: str,
    nome_aba_principal: str,
    nome_aba_listas: str,
    linha_cabecalho: int = 1,
    linha_max_aplicacao: int = 1000,
) -> None:
    """
    Aplica validação de dados (lista suspensa) nas colunas da aba principal,
    buscando automaticamente a coluna de mesmo nome na aba de listas.
    """
    wb = load_workbook(caminho_arquivo)

    if nome_aba_principal not in wb.sheetnames:
        raise ValueError(f"Aba principal '{nome_aba_principal}' não encontrada.")
    if nome_aba_listas not in wb.sheetnames:
        raise ValueError(f"Aba de listas '{nome_aba_listas}' não encontrada.")

    ws_principal = wb[nome_aba_principal]
    ws_listas = wb[nome_aba_listas]

    # 1. Mapear as colunas da aba de listas: { "nome_da_coluna": (indice_coluna, ultima_linha_preenchida) }
    mapa_listas = {}
    for col_idx in range(1, ws_listas.max_column + 1):
        cabecalho = ws_listas.cell(row=linha_cabecalho, column=col_idx).value
        if cabecalho is not None:
            nome_normalizado = str(cabecalho).strip()

            # Descobrir até qual linha há dados preenchidos nessa coluna específica
            ultima_linha = linha_cabecalho
            for row_idx in range(ws_listas.max_row, linha_cabecalho, -1):
                if ws_listas.cell(row=row_idx, column=col_idx).value is not None:
                    ultima_linha = row_idx
                    break

            # Só adiciona se houver pelo menos 1 item abaixo do cabeçalho
            if ultima_linha > linha_cabecalho:
                mapa_listas[nome_normalizado] = (col_idx, ultima_linha)

    # Limpa validações anteriores da aba principal (opcional, evita duplicatas ao rodar novamente)
    ws_principal.data_validations.dataValidation.clear()

    # Nome da aba de listas formatado com aspas simples caso tenha espaços (ex: 'Base de Dados')
    aba_listas_quoted = quote_sheetname(nome_aba_listas)

    # 2. Percorrer as colunas da aba principal e criar a validação quando o nome coincidir
    for col_idx in range(1, ws_principal.max_column + 1):
        cabecalho_principal = ws_principal.cell(row=linha_cabecalho, column=col_idx).value
        if cabecalho_principal is None:
            continue

        nome_coluna = str(cabecalho_principal).strip()

        if nome_coluna in mapa_listas:
            col_lista_idx, ultima_linha_lista = mapa_listas[nome_coluna]
            letra_col_lista = get_column_letter(col_lista_idx)
            letra_col_principal = get_column_letter(col_idx)

            # Fórmula apontando para o intervalo exato daquela coluna na outra aba
            # Exemplo: ='Listas'!$A$2:$A$15
            formula_intervalo = (
                f"={aba_listas_quoted}!${letra_col_lista}${linha_cabecalho + 1}:"
                f"${letra_col_lista}${ultima_linha_lista}"
            )

            dv = DataValidation(
                type="list",
                formula1=formula_intervalo,
                allow_blank=True,
            )
            dv.errorTitle = "Valor inválido"
            dv.error = f"Por favor, selecione um valor válido para '{nome_coluna}'."
            dv.promptTitle = nome_coluna
            dv.prompt = "Selecione um item da lista."

            ws_principal.add_data_validation(dv)

            # Intervalo na aba principal onde a validação será aplicada (ex: A2:A1000)
            intervalo_destino = f"{letra_col_principal}{linha_cabecalho + 1}:{letra_col_principal}{linha_max_aplicacao}"
            dv.add(intervalo_destino)

            print(
                f"[OK] Coluna '{nome_coluna}' ({intervalo_destino}) -> "
                f"Lista em {formula_intervalo}"
            )

    wb.save(caminho_arquivo)
    print(f"\nArquivo salvo com sucesso em: {caminho_arquivo}")


def criar_arquivo_exemplo(caminho_arquivo: str) -> None:
    """Cria uma planilha de exemplo para testar a validação."""
    wb = Workbook()

    # Aba principal
    ws_principal = wb.active
    ws_principal.title = "Cadastro"
    ws_principal.append(["Nome", "Departamento", "Status", "Prioridade"])

    # Aba com as listas de validação (mesmos nomes das colunas que terão validação)
    ws_listas = wb.create_sheet(title="Listas")
    ws_listas.append(["Status", "Departamento", "Prioridade"])

    dados_listas = [
        ["Ativo", "Financeiro", "Alta"],
        ["Inativo", "Tecnologia", "Média"],
        ["Pendente", "RH", "Baixa"],
        [None, "Comercial", None],
        [None, "Operações", None],
    ]
    for linha in dados_listas:
        ws_listas.append(linha)

    wb.save(caminho_arquivo)


if __name__ == "__main__":
    arquivo = "teste.xlsx"

    # 1. Cria um arquivo de teste (remova se já tiver seu próprio arquivo .xlsx)
    #criar_arquivo_exemplo(arquivo)

    # 2. Aplica a validação relacionando os cabeçalhos das duas abas
    aplicar_validacao_por_coluna(
        caminho_arquivo=arquivo,
        nome_aba_principal="Cadastro",
        nome_aba_listas="Listas",
        linha_cabecalho=1,
        linha_max_aplicacao=500,
    )
