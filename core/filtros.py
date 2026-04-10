def separador_lancamentos_tar(objeto_extrato):
    saldos = []
    for registro in objeto_extrato:
        if 'tar' in registro['lancamento'].lower() or 'custas' in registro['lancamento'].lower():
            saldos.append(registro)
    return saldos

def separador_lancamentos_aplic(objeto_extrato):
    aplic = []
    for registro in objeto_extrato:
        if 'aplic' in registro['lancamento'].lower() :
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

def separar_lancamentos_geral(objeto_extrato):
    geral = []
    for registro in objeto_extrato:
        if 'tar' not in registro['lancamento'].lower() and 'contamax' not in registro['lancamento'].lower():
            geral.append(registro)
    return geral

def separador_lancamentos_debitos(objeto_extrato):
    debitos = []
    for registro in objeto_extrato:
        if float(registro['valor']) < 0 and 'tar' not in registro['lancamento'].lower() and 'contamax' not in registro['lancamento'].lower():
            debitos.append(registro)
    return debitos
