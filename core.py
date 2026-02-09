import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime



def looks_like_html(path, nbytes=512):
	try:
		with open(path, 'rb') as f:
			start = f.read(nbytes).lower()
		return b'<html' in start or b'<!doctype html' in start
	except Exception:
		return False

def read_with_fallback(path):
	# Try to read as a native Excel file first. If xlrd fails with BOF or
	# the content looks like HTML, try pandas.read_html as a fallback.
	try:
		if path.lower().endswith('.xlsx'):
			return pd.read_excel(path, header=None, engine='openpyxl')
		else:
			return pd.read_excel(path, header=None, engine='xlrd')
	except Exception as e:
		msg = str(e)
		# Common xlrd BOF/corrupt message contains 'Expected BOF' or 'Unsupported format'
		if 'expected bof' in msg.lower() or 'unsupported format' in msg.lower() or looks_like_html(path):
			print('Arquivo parece ser HTML ou está no formato "Web Page (HTML)". Tentando pd.read_html() como fallback...')
			try:
				tables = pd.read_html(path)
				if not tables:
					raise RuntimeError('Nenhuma tabela encontrada pelo read_html')
				# If multiple tables found, concatenate them (you may choose a specific one)
				df = pd.concat(tables, ignore_index=True)
				return df
			except Exception as e2:
				print('Falha ao tentar ler como HTML:', e2)
				raise
		else:
			raise

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

def extrato_para_obj(df):
	extrato = []

	for i, linha in df.iterrows():
		registro = {
			"data": linha["Data"],
			"lancamento": linha["Lançamento"],
			"valor": linha["Valor"],
			"saldo": linha["Saldo"],
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

def separador_lancamentos_tar(objeto_extrato):
	saldos = []
	for registro in objeto_extrato:
		if 'tar' in registro['lancamento'].lower() or 'custas' in registro['lancamento'].lower():
			saldos.append(registro)
	return saldos

def separador_lancamentos_aplic(objeto_extrato):
	aplic = []
	for registro in objeto_extrato:
		if 'aplic' in registro['lancamento'].lower():
			aplic.append(registro)
	return aplic

def separador_lancamentos_contamax(objeto_extrato):
	contamax = []
	for registro in objeto_extrato:
		if 'contamax' in registro['lancamento'].lower():
			contamax.append(registro)
	return contamax

def separador_lancamentos_creditos(objeto_extrato):
	creditos = []
	for registro in objeto_extrato:
		if float(registro['valor']) > 0 and registro['lancamento'].lower() != 'aplic' and "cdb" not in registro['lancamento'].lower() and "aplic" not in registro['lancamento'].lower():
			creditos.append(registro)
	return creditos

def separador_lancamentos_credito(objeto_extrato):
	credito = []
	for registro in objeto_extrato:
		if 'cr cob' in registro['lancamento'].lower():
			credito.append(registro)
	return credito

def separador_lancamentos_debitos(objeto_extrato):
	debitos = []
	for registro in objeto_extrato:
		if float(registro['valor']) < 0 and 'tar' not in registro['lancamento'].lower() and 'contamax' not in registro['lancamento'].lower():
			debitos.append(registro)
	return debitos

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


	
	



	data_loc = localizar_texto_df(df, 'Data:')
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
	gerar_excel_itau(
    nome=nomeCond,
    agencia_conta=agenciaContaCond,
    periodo=dataCond,

    saldo_inicial=sdo_inicial,
    saldo_final=sdo_final,

    tarifas=tarifa,
    aplicacoes=aplicacao,
    creditos=creditos,
    debitos=debitos,
	arquivo_origem=df.attrs.get("arquivo_origem")
    )

def main_santander(df):
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
	aplicacao = separador_lancamentos_contamax(obj)
	lista_aplicacao = []
	lista_resgate = []
	total_aplic = 0
	total_resgate = 0
	for i in aplicacao:
		if i["valor"] > 0:
			lista_resgate.append(i)
			total_resgate += float(i['valor'])
		else:
			lista_aplicacao.append(i)
			total_aplic += float(i['valor'])
	print("---------------------------------Aplicação----------------------------------------")
	for registro in lista_aplicacao:
		
		print("Data: ", registro['data'], " Lançamento: ", registro['lancamento'], " Valor: ", registro['valor'], " Saldo: ", registro['saldo'])
	print("Qtde: ", len(lista_aplicacao), " Total Aplicações: ", total_aplic.__round__(2))
	print("---------------------------------Resgate----------------------------------------")

	for registro in lista_resgate:
		
		print("Data: ", registro['data'], " Lançamento: ", registro['lancamento'], " Valor: ", registro['valor'], " Saldo: ", registro['saldo'])
	print("Qtde: ", len(lista_resgate), " Total Resgates: ", total_resgate.__round__(2))
	print("Resultado Líquido: ", (total_aplic + total_resgate).__round__(2))

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

	gerar_excel_santander(
    agencia_conta=agenciaContaCond,

    periodo="Período não identificado",   # ← você ainda não está extraindo data no Santander

    saldo_inicial=sdo_inicial,
    saldo_final=sdo_final,

    tarifas=tarifa,

    contamax_aplicacoes=lista_aplicacao,
    contamax_resgates=lista_resgate,

    credito=credito,

    debitos=debitos,

    arquivo_origem=df.attrs.get("arquivo_origem")
)

def main_bradesco(df):
	
	print("---------------------------------Sheets SALDOS----------------------------------------")
	print(df)
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
	gerar_excel_bradesco(
    agencia_conta=agenciaContaCond,

    saldo_inicial=sdo_inicial,
    saldo_final=sdo_final,
    tarifas=tarifa,
    aplicacoes=aplicacao,
    creditos=creditos,
    debitos=debitos,
	arquivo_origem=df.attrs.get("arquivo_origem")
    )

def gerar_excel_bradesco(
    agencia_conta,
    saldo_inicial,
    saldo_final,
    tarifas,
    aplicacoes,
    creditos,
    debitos,
    arquivo_origem
):

    # ===============================
    # FUNÇÕES AUXILIARES
    # ===============================

    def lista_para_df(lista):
        df = pd.DataFrame(lista)
    
        if df.empty:
            return df
    
        df = df.rename(columns={
            "data": "Data",
            "lancamento": "Lançamento",
            "valor": "Valor",
            "saldo": "Saldo"
        })

        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce")
        df["Saldo"] = pd.to_numeric(df["Saldo"], errors="coerce")

        return df.dropna(how="all")

    def estilo_padrao(ws):
        """Aplica o mesmo visual do template Santander/Itaú"""

        header = PatternFill("solid", fgColor="002060")   # azul escuro
        branco = Font(color="FFFFFF", bold=True)

        borda = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Cabeçalho
        for cell in ws[1]:
            cell.fill = header
            cell.font = branco
            cell.alignment = Alignment(horizontal="center")
            cell.border = borda

        # Corpo
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = borda

        # Ajustar largura
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter

            for cell in col:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except:
                    pass

            ws.column_dimensions[col_letter].width = max_length + 3

    # ===============================
    # DATAFRAMES
    # ===============================

    df_tarifas = lista_para_df(tarifas)
    df_aplic   = lista_para_df(aplicacoes)
    df_cred    = lista_para_df(creditos)
    df_deb     = lista_para_df(debitos)

    # ABA RESUMO (igual ao Itaú/Santander)
    df_resumo = pd.DataFrame([{
        "Agência/Conta": agencia_conta,
        "Saldo Inicial": saldo_inicial,
        "Saldo Final": saldo_final,
        "Total Tarifas": df_tarifas["Valor"].astype(float).sum() if not df_tarifas.empty else 0,
        "Total Aplicações": df_aplic["Valor"].astype(float).sum() if not df_aplic.empty else 0,
        "Total Créditos": df_cred["Valor"].astype(float).sum() if not df_cred.empty else 0,
        "Total Débitos": df_deb["Valor"].astype(float).sum() if not df_deb.empty else 0,
        "Arquivo Origem": arquivo_origem or ""
    }])

    # ===============================
    # NOME DO ARQUIVO
    # ===============================

    base = os.path.basename(arquivo_origem)
    nomeInicial, ext = os.path.splitext(base)
    nome_arquivo = f"BRADESCO_{arquivo_origem}.xlsx"

    # pegar pasta onde o programa está rodando
    pasta = os.path.dirname(arquivo_origem)

    caminho = os.path.join(pasta, nome_arquivo)

    # ===============================
    # GRAVAÇÃO
    # ===============================

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:

        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
        df_tarifas.to_excel(writer, sheet_name="Tarifas", index=False)
        df_aplic.to_excel(writer, sheet_name="Aplicacoes", index=False)
        df_cred.to_excel(writer, sheet_name="Creditos", index=False)
        df_deb.to_excel(writer, sheet_name="Debitos", index=False)

        wb = writer.book

        for aba in wb.sheetnames:
            ws = wb[aba]
            estilo_padrao(ws)

            # Formatar colunas de valor igual Itaú
            for col in ["C", "D"]:   # Valor e Saldo
                for cell in ws[col][1:]:
                    cell.number_format = 'R$ #,##0.00'

    print(f"\n✅ Excel Bradesco gerado no padrão Itaú/Santander:")
    print(caminho)

    return caminho

def gerar_excel_itau(
    nome,
    agencia_conta,
    periodo,
    saldo_inicial,
    saldo_final,
    tarifas,
    aplicacoes,
    creditos,
    debitos,
    arquivo_origem
):
    import os
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    wb = Workbook()

    base = os.path.basename(arquivo_origem)
    nome_sem_ext, ext = os.path.splitext(base)
    caminho = nome_sem_ext + "_extrato_conciliacao.xlsx"

    # ========== ESTILOS ==========
    font_titulo = Font(name='Arial', bold=True, size=11)
    font_cabecalho_tabela = Font(name='Arial', bold=True, size=10)
    font_normal = Font(name='Arial', size=10)
    font_total = Font(name='Arial', bold=True, size=10)
    font_negativo = Font(name='Arial', size=10, color="FF0000")
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    fill_cabecalho = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    fill_total = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    border_thin = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )

    def montar_aba(nome_aba, lista, titulo_secao):
        ws = wb.create_sheet(nome_aba)
        
        # Logo/Cabeçalho
        ws.merge_cells('B2:C2')
        ws['B2'] = 'Itaú'
        ws['B2'].font = Font(name='Arial', bold=True, size=14, color="0066CC")
        ws['B2'].alignment = align_left
        
        # Informações do cliente
        ws['B5'] = 'Nome:'
        ws['B5'].font = font_titulo
        ws['C5'] = nome
        ws['C5'].font = font_normal
        
        ws['D5'] = 'Agência/Conta:'
        ws['D5'].font = font_titulo
        ws['E5'] = agencia_conta
        ws['E5'].font = font_normal
        
        ws['B6'] = 'Data:'
        ws['B6'].font = font_titulo
        ws['C6'] = periodo
        ws['C6'].font = font_normal
        
        # Título da seção
        ws['B9'] = titulo_secao
        ws['B9'].font = Font(name='Arial', bold=True, size=11)
        
        # Cabeçalho da tabela
        linha = 12
        headers = ["Data", "Lançamento", "Valor (R$)", "Saldo (R$)"]
        colunas_letra = ["B", "C", "D", "E"]
        
        for col, texto in zip(colunas_letra, headers):
            cell = ws[f"{col}{linha}"]
            cell.value = texto
            cell.font = font_cabecalho_tabela
            cell.alignment = align_center
            cell.fill = fill_cabecalho
            cell.border = border_thin
        
        if not lista:
            ws[f"B{linha+1}"] = "Sem registros para esta categoria"
            ws[f"B{linha+1}"].font = Font(name='Arial', italic=True, size=10)
            ws.column_dimensions["B"].width = 12
            ws.column_dimensions["C"].width = 50
            ws.column_dimensions["D"].width = 18
            ws.column_dimensions["E"].width = 18
            return
        
        # Dados
        linha += 1
        linha_inicio_dados = linha
        
        for reg in lista:
            valor = float(reg["valor"])
            
            ws[f"B{linha}"] = reg["data"]
            ws[f"B{linha}"].font = font_normal
            ws[f"B{linha}"].alignment = align_center
            ws[f"B{linha}"].border = border_thin
            
            ws[f"C{linha}"] = reg["lancamento"]
            ws[f"C{linha}"].font = font_normal
            ws[f"C{linha}"].alignment = align_left
            ws[f"C{linha}"].border = border_thin
            
            ws[f"D{linha}"] = valor
            ws[f"D{linha}"].font = font_negativo if valor < 0 else font_normal
            ws[f"D{linha}"].alignment = align_right
            ws[f"D{linha}"].number_format = '#,##0.00'
            ws[f"D{linha}"].border = border_thin
            
            ws[f"E{linha}"] = reg["saldo"]
            ws[f"E{linha}"].font = font_normal
            ws[f"E{linha}"].alignment = align_right
            ws[f"E{linha}"].number_format = '#,##0.00'
            ws[f"E{linha}"].border = border_thin
            
            linha += 1
        
        # Linha de TOTAL
        ws.merge_cells(f'B{linha}:C{linha}')
        ws[f"B{linha}"] = "TOTAL"
        ws[f"B{linha}"].font = font_total
        ws[f"B{linha}"].alignment = align_center
        ws[f"B{linha}"].fill = fill_total
        ws[f"B{linha}"].border = border_thin
        
        ws[f"D{linha}"] = f'=SUM(D{linha_inicio_dados}:D{linha-1})'
        ws[f"D{linha}"].font = font_total
        ws[f"D{linha}"].alignment = align_right
        ws[f"D{linha}"].number_format = '#,##0.00'
        ws[f"D{linha}"].fill = fill_total
        ws[f"D{linha}"].border = border_thin
        
        ws[f"E{linha}"].fill = fill_total
        ws[f"E{linha}"].border = border_thin
        
        # Ajustar larguras
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 50
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 18

    # ========== ABA RESUMO ==========
    ws = wb.active
    ws.title = "Resumo"
    
    ws.merge_cells('B2:D2')
    ws['B2'] = 'RELATÓRIO DE CONCILIAÇÃO BANCÁRIA - ITAÚ'
    ws['B2'].font = Font(name='Arial', bold=True, size=13, color="0066CC")
    ws['B2'].alignment = align_center
    
    ws['B4'] = 'Nome:'
    ws['B4'].font = font_titulo
    ws['C4'] = nome
    ws['C4'].font = font_normal
    
    ws['B5'] = 'Agência/Conta:'
    ws['B5'].font = font_titulo
    ws['C5'] = agencia_conta
    ws['C5'].font = font_normal
    
    ws['B6'] = 'Período:'
    ws['B6'].font = font_titulo
    ws['C6'] = periodo
    ws['C6'].font = font_normal
    
    ws['B8'] = 'SALDOS'
    ws['B8'].font = Font(name='Arial', bold=True, size=11)
    
    ws['B9'] = 'Saldo Inicial:'
    ws['B9'].font = font_normal
    ws['C9'] = saldo_inicial
    ws['C9'].font = font_normal
    ws['C9'].number_format = '#,##0.00'
    
    ws['B10'] = 'Saldo Final:'
    ws['B10'].font = font_normal
    ws['C10'] = saldo_final
    ws['C10'].font = font_normal
    ws['C10'].number_format = '#,##0.00'
    
    ws['B12'] = 'TOTALIZADORES'
    ws['B12'].font = Font(name='Arial', bold=True, size=11)
    
    ws['B13'] = 'Total Tarifas/Custas:'
    ws['B13'].font = font_normal
    ws['C13'] = sum(float(x["valor"]) for x in tarifas) if tarifas else 0
    ws['C13'].font = Font(name='Arial', size=10, color="FF0000")
    ws['C13'].number_format = '#,##0.00'
    
    ws['B14'] = 'Total Créditos:'
    ws['B14'].font = font_normal
    ws['C14'] = sum(float(x["valor"]) for x in creditos) if creditos else 0
    ws['C14'].font = Font(name='Arial', size=10, color="008000")
    ws['C14'].number_format = '#,##0.00'
    
    ws['B15'] = 'Total Débitos:'
    ws['B15'].font = font_normal
    ws['C15'] = sum(float(x["valor"]) for x in debitos) if debitos else 0
    ws['C15'].font = Font(name='Arial', size=10, color="FF0000")
    ws['C15'].number_format = '#,##0.00'
    
    ws['B16'] = 'Total Aplicações:'
    ws['B16'].font = font_normal
    ws['C16'] = sum(float(x["valor"]) for x in aplicacoes) if aplicacoes else 0
    ws['C16'].font = font_normal
    ws['C16'].number_format = '#,##0.00'
    
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    
    # ========== CRIAR ABAS ==========
    montar_aba("Tarifas", tarifas, "APURAÇÃO TARIFAS BANCÁRIAS")
    montar_aba("Creditos", creditos, "CRÉDITOS")
    montar_aba("Debitos", debitos, "DÉBITOS")
    montar_aba("Aplicacoes", aplicacoes, "APLICAÇÕES")
    
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    
    wb.save(caminho)
    print(f"Arquivo Itaú estilizado gerado: {caminho}")


def gerar_excel_santander(
    
    agencia_conta,
    periodo,
    saldo_inicial,
    saldo_final,
    tarifas,
    contamax_aplicacoes,
    contamax_resgates,
    credito,
    debitos,
    arquivo_origem
):
    import os
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    wb = Workbook()

    base = os.path.basename(arquivo_origem)
    nome_sem_ext, ext = os.path.splitext(base)
    caminho = nome_sem_ext + "_extrato_conciliacao_santander.xlsx"

    # ========== ESTILOS SANTANDER ==========
    font_titulo = Font(name='Arial', bold=True, size=11)
    font_titulo_principal = Font(name='Arial', bold=True, size=14, color="4472C4")
    font_cabecalho_tabela = Font(name='Arial', bold=True, size=10, color="FFFFFF")
    font_normal = Font(name='Arial', size=10)
    font_total = Font(name='Arial', bold=True, size=10)
    font_negativo = Font(name='Arial', size=10, color="FF0000")
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    fill_cabecalho = PatternFill(start_color="7F7F7F", end_color="7F7F7F", fill_type="solid")
    fill_total = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_linha_alternada = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    border_thin = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )

    def montar_aba(nome_aba, lista, titulo_secao, mostrar_saldo=True):
        ws = wb.create_sheet(nome_aba)
        
        # Logo Santander
        ws.merge_cells('B2:C2')
        ws['B2'] = 'Santander'
        ws['B2'].font = Font(name='Arial', bold=True, size=14, color="EC0000")
        ws['B2'].alignment = align_left
        
        # Título "Internet Banking Empresarial"
        ws.merge_cells('D2:F2')
        ws['D2'] = 'Internet Banking Empresarial'
        ws['D2'].font = font_titulo_principal
        ws['D2'].alignment = align_right
        
        # Nome do condomínio
        ws.merge_cells('B5:C5')
        ws['B5'].font = font_titulo
        ws['B5'].alignment = align_left
        
        # Agência e Conta
        agencia, conta = agencia_conta.split('/')
        ws['D5'] = f'Agência: {agencia}'
        ws['D5'].font = font_normal
        ws['D5'].alignment = align_right
        
        ws['E5'] = f'Conta: {conta}'
        ws['E5'].font = font_normal
        ws['E5'].alignment = align_right
        
        # Título da seção
        ws.merge_cells('B7:F7')
        ws['B7'] = titulo_secao
        ws['B7'].font = Font(name='Arial', bold=True, size=11)
        ws['B7'].alignment = align_center
        
        # Cabeçalho da tabela
        linha = 9
        if mostrar_saldo:
            headers = ["Data", "Histórico", "Valor (R$)", "Valor (R$)"]
            colunas_letra = ["B", "C", "D", "E"]
        else:
            headers = ["Data", "Histórico", "Valor (R$)"]
            colunas_letra = ["B", "C", "D"]
        
        for col, texto in zip(colunas_letra, headers):
            cell = ws[f"{col}{linha}"]
            cell.value = texto
            cell.font = font_cabecalho_tabela
            cell.alignment = align_center
            cell.fill = fill_cabecalho
            cell.border = border_thin
        
        if not lista:
            ws[f"B{linha+1}"] = "Sem registros para esta categoria"
            ws[f"B{linha+1}"].font = Font(name='Arial', italic=True, size=10)
            ws.column_dimensions["B"].width = 12
            ws.column_dimensions["C"].width = 50
            ws.column_dimensions["D"].width = 18
            if mostrar_saldo:
                ws.column_dimensions["E"].width = 18
            return
        
        # Dados
        linha += 1
        linha_inicio_dados = linha
        
        for idx, reg in enumerate(lista):
            valor = float(reg["valor"])
            
            # Zebrar linhas (alternância de cores)
            if idx % 2 == 1:
                fill_linha = fill_linha_alternada
            else:
                fill_linha = PatternFill()
            
            ws[f"B{linha}"] = reg["data"]
            ws[f"B{linha}"].font = font_normal
            ws[f"B{linha}"].alignment = align_center
            ws[f"B{linha}"].border = border_thin
            ws[f"B{linha}"].fill = fill_linha
            
            ws[f"C{linha}"] = reg["lancamento"]
            ws[f"C{linha}"].font = font_normal
            ws[f"C{linha}"].alignment = align_left
            ws[f"C{linha}"].border = border_thin
            ws[f"C{linha}"].fill = fill_linha
            
            ws[f"D{linha}"] = valor
            ws[f"D{linha}"].font = font_negativo if valor < 0 else font_normal
            ws[f"D{linha}"].alignment = align_right
            ws[f"D{linha}"].number_format = '#,##0.00'
            ws[f"D{linha}"].border = border_thin
            ws[f"D{linha}"].fill = fill_linha
            
            if mostrar_saldo:
                ws[f"E{linha}"] = reg.get("saldo", "")
                ws[f"E{linha}"].font = font_normal
                ws[f"E{linha}"].alignment = align_right
                ws[f"E{linha}"].number_format = '#,##0.00'
                ws[f"E{linha}"].border = border_thin
                ws[f"E{linha}"].fill = fill_linha
            
            linha += 1
        
        # Linha de TOTAL
        if mostrar_saldo:
            ws.merge_cells(f'B{linha}:C{linha}')
        else:
            ws.merge_cells(f'B{linha}:B{linha}')
        
        ws[f"B{linha}"] = "TOTAL"
        ws[f"B{linha}"].font = font_total
        ws[f"B{linha}"].alignment = align_center
        ws[f"B{linha}"].fill = fill_total
        ws[f"B{linha}"].border = border_thin
        
        ws[f"D{linha}"] = f'=SUM(D{linha_inicio_dados}:D{linha-1})'
        ws[f"D{linha}"].font = font_total
        ws[f"D{linha}"].alignment = align_right
        ws[f"D{linha}"].number_format = '#,##0.00'
        ws[f"D{linha}"].fill = fill_total
        ws[f"D{linha}"].border = border_thin
        
        if mostrar_saldo:
            ws[f"E{linha}"].fill = fill_total
            ws[f"E{linha}"].border = border_thin
        
        # Ajustar larguras
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 50
        ws.column_dimensions["D"].width = 18
        if mostrar_saldo:
            ws.column_dimensions["E"].width = 18

    # ========== ABA RESUMO ==========
    ws = wb.active
    ws.title = "Resumo"
    
    # Logo Santander
    ws.merge_cells('B2:C2')
    ws['B2'] = 'Santander'
    ws['B2'].font = Font(name='Arial', bold=True, size=16, color="EC0000")
    ws['B2'].alignment = align_left
    
    # Título principal
    ws.merge_cells('B3:D3')
    ws['B3'] = 'RELATÓRIO DE CONCILIAÇÃO BANCÁRIA'
    ws['B3'].font = Font(name='Arial', bold=True, size=13, color="4472C4")
    ws['B3'].alignment = align_center
    
    ws['B5'] = 'Condomínio:'
    ws['B5'].font = font_titulo
    ws['C5'].font = font_normal
    
    ws['B6'] = 'Agência/Conta:'
    ws['B6'].font = font_titulo
    ws['C6'] = agencia_conta
    ws['C6'].font = font_normal
    
    ws['B7'] = 'Período:'
    ws['B7'].font = font_titulo
    ws['C7'] = periodo
    ws['C7'].font = font_normal
    
    ws['B9'] = 'SALDOS'
    ws['B9'].font = Font(name='Arial', bold=True, size=11)
    
    ws['B10'] = 'Saldo Inicial:'
    ws['B10'].font = font_normal
    ws['C10'] = saldo_inicial
    ws['C10'].font = font_normal
    ws['C10'].number_format = '#,##0.00'
    
    ws['B11'] = 'Saldo Final:'
    ws['B11'].font = font_normal
    ws['C11'] = saldo_final
    ws['C11'].font = font_normal
    ws['C11'].number_format = '#,##0.00'
    
    ws['B13'] = 'TOTALIZADORES'
    ws['B13'].font = Font(name='Arial', bold=True, size=11)
    
    ws['B14'] = 'Total Tarifas/Custas:'
    ws['B14'].font = font_normal
    ws['C14'] = sum(float(x["valor"]) for x in tarifas) if tarifas else 0
    ws['C14'].font = Font(name='Arial', size=10, color="FF0000")
    ws['C14'].number_format = '#,##0.00'
    
    ws['B15'] = 'Contamax - Aplicações:'
    ws['B15'].font = font_normal
    ws['C15'] = sum(float(x["valor"]) for x in contamax_aplicacoes) if contamax_aplicacoes else 0
    ws['C15'].font = Font(name='Arial', size=10, color="FF0000")
    ws['C15'].number_format = '#,##0.00'
    
    ws['B16'] = 'Contamax - Resgates:'
    ws['B16'].font = font_normal
    ws['C16'] = sum(float(x["valor"]) for x in contamax_resgates) if contamax_resgates else 0
    ws['C16'].font = Font(name='Arial', size=10, color="008000")
    ws['C16'].number_format = '#,##0.00'
    
    ws['B17'] = 'Contamax - Resultado:'
    ws['B17'].font = font_titulo
    total_contamax = (sum(float(x["valor"]) for x in contamax_aplicacoes) if contamax_aplicacoes else 0) + \
                     (sum(float(x["valor"]) for x in contamax_resgates) if contamax_resgates else 0)
    ws['C17'] = total_contamax
    ws['C17'].font = Font(name='Arial', size=10, bold=True, color="0000FF")
    ws['C17'].number_format = '#,##0.00'
    
    ws['B18'] = 'Total credito]:'
    ws['B18'].font = font_normal
    ws['C18'] = sum(float(x["valor"]) for x in credito) if credito else 0
    ws['C18'].font = Font(name='Arial', size=10, color="008000")
    ws['C18'].number_format = '#,##0.00'
    
    ws['B19'] = 'Total Débitos:'
    ws['B19'].font = font_normal
    ws['C19'] = sum(float(x["valor"]) for x in debitos) if debitos else 0
    ws['C19'].font = Font(name='Arial', size=10, color="FF0000")
    ws['C19'].number_format = '#,##0.00'
    
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 18
    
    # ========== CRIAR ABAS ==========
    montar_aba("Tarifas", tarifas, "APURAÇÃO TARIFAS BANCÁRIAS", mostrar_saldo=True)
    montar_aba("Contamax_Aplicacoes", contamax_aplicacoes, "CONTAMAX - APLICAÇÕES", mostrar_saldo=True)
    montar_aba("Contamax_Resgates", contamax_resgates, "CONTAMAX - RESGATES", mostrar_saldo=True)
    montar_aba("credito", credito, "credito] (CR COB)", mostrar_saldo=True)
    montar_aba("Debitos", debitos, "DÉBITOS", mostrar_saldo=True)
    
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    
    wb.save(caminho)
    print(f"Arquivo Santander estilizado gerado: {caminho}")