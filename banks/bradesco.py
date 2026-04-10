import pandas as pd
from core.filtros import separador_lancamentos_aplic, separador_lancamentos_creditos, separador_lancamentos_debitos, separador_lancamentos_tar, separar_lancamentos_geral
from core.parser import invest_facil_bradesco, padronizar_extrato_bradesco
from core.processors import extrato_para_obj, localizar_texto_df, pegar_valor_cabecalho, ultima_cedula_saldo_bradesco
from decimal import Decimal

from outputs.excel import gerar_excel

    

def main_bradesco(df, condominio):
    
    print("---------------------------------Sheets SALDOS----------------------------------------")
    print(df)
    print("INEVST FACIL: ", localizar_texto_df(df, 'Saldos Invest Fácil / Plus'))
    agenciaConta_loc = localizar_texto_df(df, 'Agência/Conta:')
    if not agenciaConta_loc:
        agencia_loc = localizar_texto_df(df, 'Agência:')
        conta_loc = localizar_texto_df(df, 'Conta:')
        for linha, coluna, valor in agencia_loc:
            pos_ag = valor.find("Agência:") + len("Agência:")
            agenciaCond = valor[pos_ag:pos_ag + 6].strip()

            print(f"Agência encontrado: {valor}, Valor após: {agenciaCond}")
            

            
        for linha, coluna, valor in conta_loc:
            pos_ct = valor.find("Conta:") + len("Conta:")
            contaCond = valor[pos_ct:pos_ct+10].strip()
            print(f"Conta encontrada: {valor}, Valor após: {contaCond}")
        agenciaContaCond = f"{agenciaCond}/{contaCond}"
        print(f"Agência/Conta combinado: {agenciaContaCond}")
    else:
        print(agenciaConta_loc)
    for linha, coluna, valor in agenciaConta_loc:
        agenciaContaCond = pegar_valor_cabecalho(df, linha, coluna)
        print(f"Agência/Conta encontrado: {valor}, Valor após: {agenciaContaCond}")

    saldoAnt_loc = localizar_texto_df(df, 'Saldo Anterior')
    print(saldoAnt_loc)
    saldoAntCond = 0
    for linha, coluna, valor in saldoAnt_loc:
        saldoAntCond = pegar_valor_cabecalho(df, linha, coluna)
        print(f"Saldo Anterior encontrado: {valor}, Valor após: {saldoAntCond}")
        break

    sdo_inicial = saldoAntCond
    sdo_final = 0

    colunaSaldo = localizar_texto_df(df, 'Saldo (R$)')
    print(colunaSaldo)
    if colunaSaldo:
        
        coluna_index = colunaSaldo[0][1]
        linha_index, sdo_final = ultima_cedula_saldo_bradesco(df, coluna_index)
        print(f"Último saldo na coluna 'Saldo (R$)' encontrado na linha {linha_index}: {sdo_final}")
    else:
        print("Coluna 'Saldo (R$)' não encontrada.")
    print(f"\nResumo:\nAgência/Conta: {agenciaContaCond}\nSaldo Inicial: {sdo_inicial}\nSaldo Anterior: {sdo_inicial}\nSaldo Final: {sdo_final}")
    print("---------------------------------Sheets Tarifas----------------------------------------")
    print("Extrato:")
    extrato_df = padronizar_extrato_bradesco(df)
    invest = invest_facil_bradesco(df)
    print("Invest Fácil:")
    
    dataTemp = invest.iloc[-1]['Data']
    nova_linha = {
        "Data": "01"+ dataTemp[2:],
        "Lançamento": "SALDO ANTERIOR",
        "Valor": str(saldoAntCond)
    }
    
    invest = pd.concat([pd.DataFrame([nova_linha]), invest], ignore_index=True)
    val = 0
    cont = 0
    for _, row in invest.iterrows():
        # print(f"Data: {row['Data']}, Lançamento: {row['Lançamento']}, Valor: {row['Valor']}")
        print("Valor:", row['Valor'])
        if  val is None:
            val = - Decimal(row['Valor'].replace('.', '').replace(',', '.'))

            continue
        novo_valor = val - Decimal(row['Valor'].replace('.', '').replace(',', '.')) 
        invest.at[cont, 'Valor'] = str(novo_valor) 
        val = Decimal(row['Valor'].replace('.', '').replace(',', '.'))
        cont+=1
    for _, row in invest.iterrows():
        print(f"Data: {row['Data']}, Lançamento: {row['Lançamento']}, Valor: {row['Valor']}")
        
    invest = invest.iloc[1:].reset_index(drop=True)

    obj_invest = extrato_para_obj(invest, "Bradesco")
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
    geral = separar_lancamentos_geral(obj)   
      
    print("---------------------------------Sheets Débitos----------------------------------------")
    total_debitos = 0
    for registro in debitos:
        print("Data: ", registro['data'], " Lançamento: ", registro['lancamento'], " Valor: ", registro['valor'], " Saldo: ", registro['saldo'])
        total_debitos += float(registro['valor'])
    print("Qtde: ", len(debitos), " Total Débitos: ", total_debitos.__round__(2))
   
    gerar_excel(
    nome=condominio,
    banco="Bradesco",
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
    invest = obj_invest
    )
