import os
import sys
import pandas as pd
def carregar_mapa_grupos(banco):
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    caminho = os.path.join(base_path, "..", "grupo-aplic.xlsx")
    caminho= os.path.abspath(caminho)
    print(f"Caminho agrupamento: {caminho}")

    df_mapa = pd.read_excel(caminho, sheet_name=banco)

    df_mapa["Descrição"] = df_mapa["Descrição"].astype(str).str.upper().str.strip()

    return dict(zip(df_mapa["Descrição"], df_mapa["Grupo"]))

def separar_por_grupo(obj, banco, invest_obj=None):

    mapeamento = carregar_mapa_grupos(banco)

    aplicacoes_por_grupo = {}
    lancamentos_gerais = []
    print("Iniciando separação por grupo...")
    print(f"OBJ -> {obj}")
    print(f"INVEST_OBJ -> {invest_obj}")

    for reg in obj:

        descricao = str(reg["lancamento"]).upper().strip()
        grupo_encontrado = None

        for desc_mapa, grupo in mapeamento.items():
            if desc_mapa in descricao:
                grupo_encontrado = grupo
                break

        if grupo_encontrado:
            if grupo_encontrado not in aplicacoes_por_grupo:
                aplicacoes_por_grupo[grupo_encontrado] = []

            aplicacoes_por_grupo[grupo_encontrado].append(reg)
        elif 'tar' not in descricao.lower():
            lancamentos_gerais.append(reg)
    # ✅ GARANTE QUE O GRUPO EXISTE
    aplicacoes_por_grupo.setdefault("SALDO INVEST", [])
    if invest_obj:
      for reg in invest_obj:
          aplicacoes_por_grupo["SALDO INVEST"].append(reg)
          print(f"Processando investimento: {reg}")

    return aplicacoes_por_grupo, lancamentos_gerais
