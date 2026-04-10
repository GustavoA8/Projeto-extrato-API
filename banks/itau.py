import pandas as pd
from core.filtros import separador_lancamentos_aplic, separador_lancamentos_creditos, separador_lancamentos_debitos, separador_lancamentos_tar, separar_lancamentos_geral
from core.parser import invest_facil_bradesco, padronizar_extrato_bradesco, padronizar_extrato_itau
from core.processors import extrato_para_obj, localizar_texto_df, pegar_valor_cabecalho, ultima_cedula_saldo_bradesco, ultima_celula_saldo
from decimal import Decimal

from outputs.excel import gerar_excel

def main_itau(df):
    print("---------------------------------Sheets SALDOS----------------------------------------")
    nome_loc = localizar_texto_df(df, 'Nome:')
    

    print(nome_loc)
    
    for linha, coluna, valor in nome_loc:
        nomeCond = pegar_valor_cabecalho(df, linha, coluna)
        print(f"Nome encontrado: {valor}, Valor após: {nomeCond}")

    print(nomeCond)

    agenciaConta_loc = localizar_texto_df(df, 'Agência/Conta:')
    if not agenciaConta_loc:
        agencia_loc = localizar_texto_df(df, 'Agência:')
        conta_loc = localizar_texto_df(df, 'Conta:')
        for linha, coluna, valor in agencia_loc:
            agenciaCond = pegar_valor_cabecalho(df, linha, coluna)
            print(f"Agência encontrado: {valor}, Valor após: {agenciaCond}")
        for linha, coluna, valor in conta_loc:
            contaCond = pegar_valor_cabecalho(df, linha, coluna)
            print(f"Conta encontrada: {valor}, Valor após: {contaCond}")
        agenciaContaCond = f"{agenciaCond}/{contaCond}"
        print(f"Agência/Conta combinado: {agenciaContaCond}")
    else:
        print(agenciaConta_loc)
    for linha, coluna, valor in agenciaConta_loc:
        agenciaContaCond = pegar_valor_cabecalho(df, linha, coluna)
        print(f"Agência/Conta encontrado: {valor}, Valor após: {agenciaContaCond}")


    
    



    data_loc = localizar_texto_df(df, 'Periodo:')
    if not data_loc:
        data_loc = localizar_texto_df(df, 'Atualização:')


    print(data_loc)
    for linha, coluna, valor in data_loc:
        dataCond = pegar_valor_cabecalho(df, linha, coluna)
        print(f"Data encontrado: {valor}, Valor após: {dataCond}")

    saldoAnt_loc = localizar_texto_df(df, 'Saldo Anterior')
    print(saldoAnt_loc)

    for linha, coluna, valor in saldoAnt_loc:
        saldoAntCond = pegar_valor_cabecalho(df, linha, coluna)
        print(f"Saldo Anterior encontrado: {valor}, Valor após: {saldoAntCond}")

    sdo_inicial = saldoAntCond
    sdo_final = 0

    colunaSaldo = localizar_texto_df(df, 'Saldo (R$)')
    print(colunaSaldo)
    if colunaSaldo:
        
        coluna_index = colunaSaldo[0][1]
        linha_index, sdo_final = ultima_celula_saldo(df, coluna_index)
        print(f"Último saldo na coluna 'Saldo (R$)' encontrado na linha {linha_index}: {sdo_final}")
    else:
        print("Coluna 'Saldo (R$)' não encontrada.")
    print(f"\nResumo:\nNome: {nomeCond}\nAgência/Conta: {agenciaContaCond}\nData: {dataCond[:10]}\nSaldo Inicial: {sdo_inicial}\nSaldo Anterior: {sdo_inicial}\nSaldo Final: {sdo_final}")
    print("---------------------------------Sheets Tarifas----------------------------------------")
    print("Extrato:")
    extrato_df = padronizar_extrato_itau(df)
    obj = extrato_para_obj(extrato_df)
    tarifa = separador_lancamentos_tar(obj)
    total = 0
    for registro in tarifa:
        print("Data: ", registro['data'], " Lançamento: ", registro['lancamento'], " Valor: ", registro['valor'], " Saldo: ", registro['saldo'])
        total += float(registro['valor'])

    print("Qtde: ", len(tarifa), " Total Tarifas/Custas: ", total.__round__(2))

    aplicacao = separador_lancamentos_aplic(obj)
    print("---------------------------------Sheets Aplicações----------------------------------------")
    total_aplic = 0
    for registro in aplicacao:
        print("Data: ", registro['data'], " Lançamento: ", registro['lancamento'], " Valor: ", registro['valor'], " Saldo: ", registro['saldo'])
        total_aplic += float(registro['valor'])
    print("Qtde: ", len(aplicacao), " Total Aplicações: ", total_aplic.__round__(2))

    creditos = separador_lancamentos_creditos(obj)
    print("---------------------------------Sheets Créditos----------------------------------------")
    total_creditos = 0
    for registro in creditos:
        print("Data: ", registro['data'], " Lançamento: ", registro['lancamento'], " Valor: ", registro['valor'], " Saldo: ", registro['saldo'])
        total_creditos += float(registro['valor'])
    print("Qtde: ", len(creditos), " Total Créditos: ", total_creditos.__round__(2))
    
    debitos = separador_lancamentos_debitos(obj)
    print("---------------------------------Sheets Débitos----------------------------------------")
    total_debitos = 0
    for registro in debitos:
        print("Data: ", registro['data'], " Lançamento: ", registro['lancamento'], " Valor: ", registro['valor'], " Saldo: ", registro['saldo'])
        total_debitos += float(registro['valor'])
    print("Qtde: ", len(debitos), " Total Débitos: ", total_debitos.__round__(2))
    geral = separar_lancamentos_geral(obj)
    print(f"Itau: {df.attrs.get('arquivo_origem')}")
    
    gerar_excel(
    nome=nomeCond,
    banco="Itau",
    agencia_conta=agenciaContaCond,
    

    saldo_inicial=sdo_inicial,
    saldo_final=sdo_final,

    tarifas=tarifa,
    aplicacoes=aplicacao,
    creditos=creditos,
    debitos=debitos,
    geral=geral,
    arquivo_origem=df.attrs.get("arquivo_origem"),
    obj = obj,
    periodo=dataCond
    )
