# main.py
from excel_reader import read_excel_tables
from db_config import engine

def create_and_insert_tables(tables, table_names):
    for df, table_name in zip(tables, table_names):
        df.to_sql(table_name, engine, if_exists='replace', index=False)
        print(f"Tabela '{table_name}' criada/inserida com sucesso.")

def main():
    file_path = input("Digite o caminho do arquivo Excel: ")

    num_tables = int(input("Quantas tabelas existem nesse Excel? "))

    tables_info = []
    table_names = []

    for i in range(num_tables):
        print(f"\nConfiguração da Tabela {i+1}")
        sheet = input("Nome ou índice da planilha: ")
        start_row = int(input("Linha do cabeçalho (ex: 2): "))
        table_name = input("Nome da tabela no banco: ")
        tables_info.append({'sheet': sheet, 'start_row': start_row})
        table_names.append(table_name)

    tables = read_excel_tables(file_path, tables_info)

    create_and_insert_tables(tables, table_names)

if __name__ == "__main__":
    main()
