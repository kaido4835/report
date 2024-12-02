import os
import pandas as pd
from data_processing import main, apply_structure_and_sorting


def main_script(input_file, sheet_name, report_file, report_sheet):
    division_column = 'Деления'  # Название колонки для деления DataFrame
    type_column = 'ТИП кредита'  # Название колонки типа
    product_kind_column = 'Вид продукта'  # Название колонки вида продукта
    result_column = 'Результат'  # Название колонки результата
    header_row = 0  # Номер строки, содержащей заголовки столбцов (если не первая строка, укажите соответствующее значение)

    # Получение обработанных данных
    processed_dataframes = main(input_file, sheet_name, division_column, type_column, product_kind_column,
                                result_column, report_file, report_sheet, header_row)

    # Применение структуры и сортировки к общему итоговому DataFrame
    overall_summary = pd.DataFrame()
    if processed_dataframes:
        for division, df in processed_dataframes.items():
            # Добавление данных к общему итоговому DataFrame
            overall_summary = overall_summary.add(df, fill_value=0)

        if not overall_summary.empty:
            overall_summary = apply_structure_and_sorting(overall_summary)
            processed_dataframes['Общий итог'] = overall_summary

        print("Data has been successfully processed.")
        print("DataFrames names and their content:")
        for name, df in processed_dataframes.items():
            print(f"\n{name}:\n{df}")
    else:
        print("No data was processed.")

    return processed_dataframes


if __name__ == "__main__":
    # Абсолютный путь к директории, где находятся файлы Excel
    base_dir = os.path.dirname(os.path.abspath(__file__))

    input_file = os.path.join(base_dir, 'База кредитная 2106 (1).xlsx')  # Абсолютный путь к файлу с данными
    sheet_name = 'База для свода'  # Название листа с данными
    report_file = os.path.join(base_dir, 'Отчёт 21.06.2024.xlsx')  # Абсолютный путь к файлу для поиска ключевых слов
    report_sheet = 'Сводная погашения NEW'  # Название листа для поиска ключевых слов

    processed_dataframes = main_script(input_file, sheet_name, report_file, report_sheet)
    # Здесь вы можете сохранить processed_dataframes или выполнить другие действия с ними.
