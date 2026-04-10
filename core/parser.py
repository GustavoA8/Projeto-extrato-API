import pandas as pd
from core.processors import (
    localizar_texto_df,
    pegar_valor_cabecalho,
    converter_valor_br,
    extrato_para_obj,
    ultima_celula_saldo,
    ultima_cedula_saldo_bradesco
)

def padronizar_extrato_itau(df):

    # 1 - achar linha do cabeçalho
    h = None

    for i, linha in df.iterrows():

        textos = [str(v).lower() for v in linha if pd.notna(v)]

        if ("data" in " ".join(textos) and
            "saldo" in " ".join(textos)):
            h = i
            break

    if h is None:
        print("Não achei linha de cabeçalho")
        return None


    # 2 - mapear colunas
    header = df.iloc[h]

    col_data = None
    col_lanc = None
    col_valor = None
    col_saldo = None

    for j, v in header.items():
        if pd.isna(v):
            continue

        v = str(v).lower()

        if "data" in v:
            col_data = j

        elif "lan" in v:
            col_lanc = j

        elif "valor" in v:
            col_valor = j

        elif "saldo" in v:
            col_saldo = j


    # 3 - validação
    if None in (col_data, col_lanc, col_valor, col_saldo):
        print("Mapeamento falhou:",
              col_data, col_lanc, col_valor, col_saldo)
        return None


    # 4 - cortar até último saldo
    linha_fim, _ = ultima_celula_saldo(df, col_saldo)

    dados = df.iloc[h+1 : linha_fim+1]


    # 5 - montar dataframe final
    novo = pd.DataFrame({
        "Data": dados.iloc[:, col_data],
        "Lançamento": dados.iloc[:, col_lanc],
        "Valor": dados.iloc[:, col_valor],
        "Saldo": dados.iloc[:, col_saldo],
    })

    return novo.dropna(how='all')

def padronizar_extrato_santander(df):

    # 1 - achar linha do cabeçalho
    h = None

    for i, linha in df.iterrows():

        textos = [str(v).lower() for v in linha if pd.notna(v)]

        if ("data" in " ".join(textos) and
            "saldo" in " ".join(textos)):
            h = i
            break

    if h is None:
        print("Não achei linha de cabeçalho")
        return None


    # 2 - mapear colunas
    header = df.iloc[h]

    col_data = None
    col_lanc = None
    col_valor = None
    col_saldo = None

    for j, v in header.items():
        if pd.isna(v):
            continue

        v = str(v).lower()

        if "data" in v:
            col_data = j

        elif "hist" in v:
            col_lanc = j

        elif "valor" in v:
            col_valor = j

        elif "saldo" in v:
            col_saldo = j


    # 3 - validação
    if None in (col_data, col_lanc, col_valor, col_saldo):
        print("Mapeamento falhou:",
              col_data, col_lanc, col_valor, col_saldo)
        return None


    # 4 - cortar até último saldo
    linha_fim, _ = ultima_celula_saldo(df, col_saldo)

    dados = df.iloc[h+1 : linha_fim+1]


    # 5 - montar dataframe final
    novo = pd.DataFrame({
        "Data": dados.iloc[:, col_data],
        "Lançamento": dados.iloc[:, col_lanc],
        "Valor": dados.iloc[:, col_valor],
        "Saldo": dados.iloc[:, col_saldo],
    })

    return novo.dropna(how='all')

def padronizar_extrato_bradesco(df):

    # 1 - achar linha do cabeçalho
    h = None

    for i, linha in df.iterrows():
        textos = [str(v).lower() for v in linha if pd.notna(v)]

        if "data" in " ".join(textos) and "saldo" in " ".join(textos):
            h = i
            break

    if h is None:
        print("Não achei linha de cabeçalho Bradesco")
        return None


    header = df.iloc[h]


    # 2 - mapear colunas
    col_data = None
    col_lanc = None
    col_credito = None
    col_debito = None
    col_saldo = None

    for j, v in header.items():
        if pd.isna(v):
            continue

        v = str(v).lower()

        if "data" in v:
            col_data = j

        elif "lan" in v or "hist" in v:
            col_lanc = j

        elif "crédito" in v or "credito" in v:
            col_credito = j

        elif "débito" in v or "debito" in v:
            col_debito = j

        elif "saldo" in v:
            col_saldo = j


    if None in (col_data, col_lanc, col_credito, col_debito, col_saldo):
        print("Mapeamento falhou Bradesco:",
              col_data, col_lanc, col_credito, col_debito, col_saldo)
        return None


    # 3 - achar linha TOTAL
    linha_fim = None

    for i in range(h, len(df)):
        linha = df.iloc[i]
        textos = [str(v).lower() for v in linha if pd.notna(v)]

        if any("total" in t for t in textos):
            linha_fim = i
            break

    if linha_fim is None:
        print("Não achei linha TOTAL do Bradesco")
        return None


    dados = df.iloc[h+1 : linha_fim].copy()


    # 4 - montar valor único com sinal
    valores = []

    for _, row in dados.iterrows():

        cred = converter_valor_br(row.iloc[col_credito])
        deb  = -(converter_valor_br(row.iloc[col_debito]))

        valor = cred - deb   # débito vira negativo

        valores.append(valor)


    # 5 - montar dataframe final
    novo = pd.DataFrame({
        "Data": dados.iloc[:, col_data],
        "Lançamento": dados.iloc[:, col_lanc],
        "Valor": valores,
        "Saldo": dados.iloc[:, col_saldo].apply(converter_valor_br),
    })


    return novo.dropna(how='all')

def invest_facil_bradesco(df):
     # 1 - achar linha do cabeçalho
    h = None
    linhaInvest = localizar_texto_df(df, 'Saldos Invest Fácil / Plus')
    print(f"Linha Invest Fácil: {linhaInvest[0][0]}")
    if not linhaInvest:
        print("Não achou 'Invest Fácil'")
        return None

    linha_inicio = linhaInvest[0][0]

    for i, linha in df.iloc[linha_inicio:].iterrows():
        textos = [str(v).lower() for v in linha if pd.notna(v)]

        if "data" in " ".join(textos) and "valor" in " ".join(textos):
            h = i
            break

    if h is None:
        print("Não achei linha de cabeçalho Bradesco")
        return None
    
    
    header = df.iloc[h]
    with open("saida.txt", "w", encoding="utf-8") as arquivo:
     arquivo.write(str(header))
    # 2 - mapear colunas
    col_data = None
    col_lanc = None
    col_val = None

    for j, v in header.items():
        if pd.isna(v):
            continue

        v = str(v).lower()

        if "data" in v:
            col_data = j

        elif "lan" in v or "hist" in v:
            col_lanc = j
    
        elif "valor" in v:
            col_val = j


    if None in (col_data, col_lanc, col_val):
        print("Mapeamento falhou Bradesco:",
              col_data, col_lanc, col_val)
        return None


    # 3 - achar linha TOTAL
    linha_fim = None

    for i in range(h, len(df)):
        linha = df.iloc[i]

        if linha.dropna().empty:
            linha_fim = i
            break

    # se não achou linha vazia, pega até o final
    if linha_fim is None:
       linha_fim = len(df)


    dados = df.iloc[h+1 : linha_fim].copy()




    # 5 - montar dataframe final
    novo = pd.DataFrame({
        "Data": dados[ col_data],
        "Lançamento": dados[ col_lanc],
        "Valor": dados[col_val]
    })


    return novo.dropna(how='all')
