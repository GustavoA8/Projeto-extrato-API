import pandas as pd


def localizar_texto_df(df, texto):
    encontrados = []

    for i, linha in df.iterrows():
        for j, valor in linha.items():
            if pd.notna(valor) and texto.lower() in str(valor).lower():
                encontrados.append((i, j, valor))

    return encontrados

def pegar_valor_cabecalho(df, linha, col_inicial):
    # percorre as colunas depois da encontrada
    for col in range(col_inicial + 1, len(df.columns)):
        valor = df.iloc[linha, col]

        if pd.notna(valor) and str(valor).strip() != "":
            return valor

    return None

def converter_valor_br(valor):
    if pd.isna(valor):
        return 0.0

    v = str(valor).strip()

    # remove pontos de milhar e troca vírgula por ponto
    v = v.replace(".", "").replace(",", ".")

    try:
        return float(v)
    except:
        return 0.0

def extrato_para_obj(df,banco=None):
    extrato = []

    for i, linha in df.iterrows():
        registro = {
            "data": linha["Data"],
            "lancamento": linha["Lançamento"],
            "valor": linha["Valor"],
            "saldo": None if banco == "Bradesco" else linha["Saldo"]
        }
        extrato.append(registro)
            

    return extrato

def ultima_celula_saldo(df, coluna):
    for i in range(len(df) - 1, -1, -1):
        valor = df.iloc[i, coluna]

        if pd.notna(valor) and str(valor).strip() != "":
            return i, valor

    return None, None

def ultima_cedula_saldo_bradesco(df, coluna):
    df = df.fillna("").astype(str)

    for i, row in df.iterrows():
        # procura a linha que contém a palavra TOTAL
        if row.str.contains("total", case=False).any():
            valor = df.iloc[i, coluna]

            # tratamento de formato brasileiro
            valor_tratado = str(valor).replace(".", "").replace(",", ".")

            try:
                return i, float(valor_tratado)
            except:
                return i, valor

    return None, None
    df = df.fillna("").astype(str)

    for i, row in df.iterrows():
        if row.str.contains("total", case=False).any():
            valor = df.iloc[i, coluna_saldo]

            # converte para número se estiver com vírgula
            valor = str(valor).replace(".", "").replace(",", ".")

            try:
                return float(valor)
            except:
                return valor

    return None
