import pandas as pd
from openpyxl import load_workbook
import sys
import os


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

def padronizar_extrato(df):

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

def main():
	if len(sys.argv) > 1:
		path = sys.argv[1]
	else:
		# default to first .xls/.xlsx in cwd
		candidates = [f for f in os.listdir('.') if f.lower().endswith(('.xls', '.xlsx'))]
		if not candidates:
			print('Nenhum arquivo .xls/.xlsx encontrado na pasta atual. Passe o caminho como argumento.')
			print('Ex: python main2.py caminho\\para\\seu_arquivo.xls')
			sys.exit(1)
		path = candidates[0]

	print('Lendo arquivo:', path)

	if not os.path.exists(path):
		print('Arquivo não encontrado:', path)
		sys.exit(1)

	try:
		df = read_with_fallback(path)
	except Exception as e:
		print('Erro ao ler o arquivo:', e)
		print('\nSugestões:')
		print('- Abra o arquivo no Excel e salve como ".xlsx" (Salvar como -> Pasta de trabalho do Excel).')
		print('- Se o Excel oferecer recuperação ao abrir, escolha recuperar e depois salve como .xlsx.')
		print('- Se isso não ajudar, envie o arquivo para análise ou convertê-lo manualmente para um formato suportado.')
		sys.exit(1)

	print(df)
	

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
	print(f"\nResumo:\nNome: {nomeCond}\nAgência/Conta: {agenciaContaCond}\nData: {dataCond[:10]}\nSaldo Inicial: {sdo_inicial}\nSaldo Final: {sdo_final}")
	print("---------------------------------Sheets Tarifas----------------------------------------")
	print("Extrato:")
	extrato_df = padronizar_extrato(df)
	print(extrato_para_obj(extrato_df))
    


if __name__ == '__main__':
	main()