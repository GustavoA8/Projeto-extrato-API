from core.reader import read_with_fallback
from banks.itau import main_itau
from banks.santander import main_santander
from banks.bradesco import main_bradesco

def processar_arquivo(df, banco, condominio=None):
    if banco.lower() == "itau":
        return main_itau(df)

    elif banco.lower() == "santander":
        return main_santander(df, condominio)

    elif banco.lower() == "bradesco":
        return main_bradesco(df, condominio)

    else:
        raise ValueError("Banco não suportado")