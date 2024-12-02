import pandas as pd
from testreport import create_od_percent_table


def print_od_percent_table(file_path, sheet_name, main_merged_cells):
    df_od_percent = create_od_percent_table(file_path, sheet_name, main_merged_cells)
    print("Таблица с % ОД к просроченному портфелю и адресами:")
    print(df_od_percent)


if __name__ == "__main__":
    # Указание файлов и листов
    file_path = 'Отчёт 21.06.2024.xlsx'
    sheet_name = 'Сводная погашения NEW'
    main_merged_cells = ['30-', '30+', '60+', '90+', '180+', '365+']

    print_od_percent_table(file_path, sheet_name, main_merged_cells)
