import openpyxl

def aplicar_layout(layout_arquivo, dados_arquivo, mapeamento):
    try:
        # Carregar o arquivo de layout de dados em formato xlsx
        layout_workbook = openpyxl.load_workbook(layout_arquivo)
        layout_sheet = layout_workbook.active

        # Carregar o arquivo de dados em formato xlsx
        dados_workbook = openpyxl.load_workbook(dados_arquivo)
        dados_sheet = dados_workbook.active

        # Criar um dicionário de mapeamento inverso (campo_destino: campo_origem)
        mapeamento_inverso = {v: k for k, v in mapeamento.items()}

        # Iterar sobre as linhas de dados
        for dados_row in dados_sheet.iter_rows(min_row=2, values_only=True):
            nova_linha = []

            # Aplicar o mapeamento
            for campo_destino in layout_sheet.iter_cols(max_col=len(mapeamento), values_only=True):
                campo_origem = mapeamento_inverso.get(campo_destino[0])
                if campo_origem:
                    valor = dados_row[dados_sheet[f"{campo_origem}1"].column - 1]
                    nova_linha.append(valor)
                else:
                    nova_linha.append(None)

            layout_sheet.append(nova_linha)

        # Salvar o arquivo de layout de dados resultante
        layout_workbook.save("layout_dados_resultante.xlsx")
        return "Mapeamento aplicado com sucesso."

    except Exception as e:
        return str(e)

# Exemplo de uso
layout_arquivo = 'layout_dados.xlsx'
dados_arquivo = 'dados.xlsx'
mapeamento = {
    'ID': 'ID',
    'DATA_VENCIMENTO': 'DATA VENCIMENTO',
    'FORNECEDOR': 'FORNECEDOR',
    # Adicione outros campos e mapeamentos conforme necessário
}

resultado = aplicar_layout(layout_arquivo, dados_arquivo, mapeamento)
print(resultado)
