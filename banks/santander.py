import pandas as pd
from core.filtros import separador_lancamentos_aplic, separador_lancamentos_credito, separador_lancamentos_creditos, separador_lancamentos_debitos, separador_lancamentos_tar, separar_lancamentos_geral
from core.parser import invest_facil_bradesco, padronizar_extrato_bradesco, padronizar_extrato_santander
from core.processors import extrato_para_obj, localizar_texto_df, pegar_valor_cabecalho, ultima_cedula_saldo_bradesco, ultima_celula_saldo
from decimal import Decimal

from outputs.excel import gerar_excel

def main_santander(df, condominio):
    print("---------------------------------Sheets SALDOS----------------------------------------")
    agencia_loc = localizar_texto_df(df, 'Agencia')
    agenciaCond = df.loc[agencia_loc[0][0] , agencia_loc[0][1]+ 1]

    conta_loc = localizar_texto_df(df, 'Conta')
    contaCond = df.loc[conta_loc[0][0] , conta_loc[0][1]+ 1]


    agenciaContaCond = f"{agenciaCond.astype(int)}/{contaCond}"
    print(f"Agência/Conta combinado: {agenciaContaCond}")


    
    



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
    print(f"\nResumo:\nAgência/Conta: {agenciaContaCond}\nSaldo Inicial: {sdo_inicial}\nSaldo Anterior: {sdo_inicial}\nSaldo Final: {sdo_final}")
    print("---------------------------------Sheets Tarifas----------------------------------------")
    print("Extrato:")
    extrato_df = padronizar_extrato_santander(df)
    obj = extrato_para_obj(extrato_df)
    tarifa = separador_lancamentos_tar(obj)
    total = 0
    for registro in tarifa:
        print("Data: ", registro['data'], " Lançamento: ", registro['lancamento'], " Valor: ", registro['valor'], " Saldo: ", registro['saldo'])
        total += float(registro['valor'])

    print("Qtde: ", len(tarifa), " Total Tarifas/Custas: ", total.__round__(2))

    
    print("---------------------------------Sheets CONTAMAX----------------------------------------")
    aplicacao = separador_lancamentos_aplic(obj)
    lista_aplicacao = []
    lista_resgate = []
    total_aplic = 0
    total_resgate = 0
   
 
    print("---------------------------------Sheets credito----------------------------------------")
    credito = separador_lancamentos_credito(obj)
    total_credito = 0
    for registro in credito:
        print("Data: ", registro['data'], " Lançamento: ", registro['lancamento'], " Valor: ", registro['valor'], " Saldo: ", registro['saldo'])
        total_credito += float(registro['valor'])
    print("Qtde: ", len(credito), " Total credito]: ", total_credito.__round__(2))
    
    debitos = separador_lancamentos_debitos(obj)
    print("---------------------------------Sheets Débitos----------------------------------------")
    total_debitos = 0
    for registro in debitos:
        print("Data: ", registro['data'], " Lançamento: ", registro['lancamento'], " Valor: ", registro['valor'], " Saldo: ", registro['saldo'])
        total_debitos += float(registro['valor'])
    print("Qtde: ", len(debitos), " Total Débitos: ", total_debitos.__round__(2))
    geral = separar_lancamentos_geral(obj) 

    gerar_excel(
    nome=condominio,
    banco="Santander",
    agencia_conta=agenciaContaCond,


    saldo_inicial=sdo_inicial,
    saldo_final=sdo_final,

    tarifas=tarifa,

    aplicacoes=aplicacao,

    creditos=credito,

    debitos=debitos,
    geral=geral,

    arquivo_origem=df.attrs.get("arquivo_origem"),
    obj = obj
)
