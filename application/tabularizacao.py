from flask import current_app as app


def tabularizar(arquivo: str) -> str:
    """
    Função que automatiza a importação dos dados necessários e salva esses
    dados formatados em uma nova planilha.

    Função que recebe o nome do arquivo como parâmetro, cria uma nova planilha
    processada e retorna o nome da mesma.
    """
    
    import pandas as pd
    from datetime import date


    caminho = f'{app.config["UPLOAD_FOLDER"]}\\{arquivo}'
    print(caminho)
    nome_aba = 'LL ABERTO (T2T)'

    lista_bandeiras = [373, 389, 397, 415]

    lista_bandeiras_nomes = ['FOO', 'BAR', 'BAZ', 'QUX']

    unidade_54 = 423

    lista_colunas = [
        3, 6, 7, 9, 10, 11, 12, 13, 16, 17, 19, 21, 24, 26, 28,
        30, 32, 34, 35, 39, 41, 42, 43, 44, 46, 47, 48, 49, 51,
        52, 53, 54, 55, 56, 57, 58, 59, 61, 62, 64, 66, 68, 71
        ]
    de_para = {
        '501': '053', 
        '502':'049', 
        '503':'043', 
        '504':'046', 
        '506':'045', 
        '507':'051', 
        '509':'042', 
        '510':'050', 
        '511':'041', 
        '512': '044'
        }
    lista_de, lista_para = [], []
    for k, v in de_para.items():
        lista_de.append(k)
        lista_para.append(v)

    frames = []

    # Loop sobre as unidades dos grupos 1 e 2:
    # OBS: Grupo 1 termina em 267
    # OBS: Grupo 2 termina em 347
    for i in range(2, 347, 8):
        planilha = pd.read_excel(caminho, sheet_name=nome_aba, usecols=[i-1, i, i+1], skiprows=0, nrows=72)

        unidade = planilha.columns[1]
        
        unidade_num = unidade.split('-', 1)[0].strip()
        unidade_nome = unidade.split('-', 1)[1].strip()
        
        if len(unidade_num) == 1:
            unidade_num = '00' + unidade_num
        elif len(unidade_num) == 2:
            unidade_num = '0' + unidade_num
        elif len(unidade_num) == 3:
            unidade_num = unidade_num


        lista_descricao, lista_meta, lista_realizado = [], [], []

        for linha in lista_colunas:
            lista_descricao.append(planilha.iloc[linha, 0])
            lista_meta.append(planilha.iloc[linha, 1])
            lista_realizado.append(planilha.iloc[linha, 2])
            
        df = pd.DataFrame()

        df['Descrição'] = lista_descricao
        df['Meta'] = lista_meta
        df['Realizado'] = lista_realizado
        df.insert(0, 'Unidade', unidade_num)
        df.insert(1, 'Nome Unidade', unidade_nome)
        
        frames.append(df)

    for j in lista_bandeiras:
        planilha = pd.read_excel(caminho, sheet_name=nome_aba, usecols=[j-1, j, j+1], skiprows=0, nrows=72)

        unidade = lista_bandeiras_nomes[lista_bandeiras.index(j)]

        #unidade_num = '-'
        #unidade_nome = unidade.strip()

        unidade_num = unidade.split('-', 1)[0].strip()
        unidade_nome = unidade.split('-', 1)[1].strip()

        lista_descricao, lista_meta, lista_realizado = [], [], []

        for linha in lista_colunas:
            lista_descricao.append(planilha.iloc[linha, 0])
            lista_meta.append(planilha.iloc[linha, 1])
            lista_realizado.append(planilha.iloc[linha, 2])
            
        df = pd.DataFrame()

        df['Descrição'] = lista_descricao
        df['Meta'] = lista_meta
        df['Realizado'] = lista_realizado
        df.insert(0, 'Unidade', unidade_num)
        df.insert(1, 'Nome Unidade', unidade_nome)

        frames.append(df)

    # Unidade 54:
    planilha = pd.read_excel(caminho, sheet_name=nome_aba, usecols=[unidade_54-1, unidade_54, unidade_54+1], skiprows=0, nrows=72)

    unidade = planilha.columns[1]
            
    unidade_num = unidade.split('-', 1)[0].strip()
    unidade_nome = unidade.split('-', 1)[1].strip()

    if len(unidade_num) == 1:
        unidade_num = '00' + unidade_num
    elif len(unidade_num) == 2:
        unidade_num = '0' + unidade_num
    elif len(unidade_num) == 3:
        unidade_num = unidade_num

    lista_descricao, lista_meta, lista_realizado = [], [], []

    for linha in lista_colunas:
        lista_descricao.append(planilha.iloc[linha, 0])
        lista_meta.append(planilha.iloc[linha, 1])
        lista_realizado.append(planilha.iloc[linha, 2])
        
    df = pd.DataFrame()

    df['Descrição'] = lista_descricao
    df['Meta'] = lista_meta
    df['Realizado'] = lista_realizado
    df.insert(0, 'Unidade', unidade_num)
    df.insert(1, 'Nome Unidade', unidade_nome)

    frames.append(df)
    

    nova_planilha = pd.DataFrame()

    nova_planilha = pd.concat(frames)
    nova_planilha.reset_index(inplace=True, drop=True)

    nova_planilha['Unidade'] = nova_planilha['Unidade'].replace(lista_de, lista_para)

    nova_planilha.to_excel(
        f'{app.config["DOWNLOAD_FOLDER"]}\\Fechamento Tabularizado {date.today()}.xlsx',
        sheet_name=nome_aba,
        index=False
        )
    return f'Fechamento Tabularizado {date.today()}.xlsx'
