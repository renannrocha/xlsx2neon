# excel_reader.py
import pandas as pd

def read_excel_tables(file_path, tables_info):
    """
    tables_info: lista de dicionários com keys:
      - sheet (nome ou índice da planilha)
      - start_row (linha do cabeçalho, ex: 2)
      - nrows (quantas linhas deve considerar após cabeçalho)
    """
    tables = []
    for info in tables_info:
        df = pd.read_excel(
            file_path,
            sheet_name=info['sheet'],
            header=info['start_row'] - 1,  # Pandas é zero-based
            nrows=info.get('nrows', None)
        )
        tables.append(df)
    return tables
