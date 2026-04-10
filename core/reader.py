import pandas as pd

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
