import pandas as pd
import os

# Словари для определения вида продукта
zalogovye = [
    "автокредиты", "автомикрозаймы физическим лицам", "cashloan",
    "Бизнес авто", "Бизнес ипотека", "ипотека"
]
bezzalogovye = [
    "микрозаймы физическим лицам", "карта рассрочка", "Овердрафт"
]


def load_excel(file_path, sheet_name, header_row=0):
    """
    Загружает данные из Excel файла и возвращает DataFrame.
    """
    if not os.path.isfile(file_path):
        print(f"Error: File {file_path} does not exist.")
        return None
    try:
        # Загрузка всех данных
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
        print(f"Successfully loaded {file_path} with sheet {sheet_name}")

        # Оставляем только нужные столбцы, если они присутствуют
        expected_columns = ['ТИП кредита', 'Результат', 'Деления']
        existing_columns = [col for col in expected_columns if col in df.columns]

        if not existing_columns:
            print(f"None of the expected columns found in {file_path} on sheet {sheet_name}.")
            return None

        df = df[existing_columns]

        # Проверка наличия столбца "Вид продукта" и заполнение его значениями при отсутствии
        if 'Вид продукта' not in df.columns and 'ТИП кредита' in df.columns:
            df['Вид продукта'] = df['ТИП кредита'].apply(assign_product_kind)

        return df
    except Exception as e:
        print(f"Error loading Excel file {file_path} on sheet {sheet_name}: {e}")
        return None


def assign_product_kind(credit_type):
    """
    Определяет вид продукта на основе типа кредита.
    """
    if credit_type in zalogovye:
        return 'Залоговый'
    elif credit_type in bezzalogovye:
        return 'Беззалоговый'
    else:
        return 'Неопределенный'


def count_unique_values(df, type_column, product_kind_column, result_column):
    """
    Подсчитывает количество уникальных значений в столбце результата для каждого уникального значения в столбце типа и вида продукта.
    """
    try:
        # Подсчет по типу кредита
        type_pivot_table = df.pivot_table(index=type_column, columns=result_column, aggfunc='size', fill_value=0,
                                          observed=False)
        type_pivot_table['Итог'] = type_pivot_table.sum(axis=1)  # Добавляем столбец с итогами по строкам

        # Подсчет по виду продукта
        product_kind_pivot_table = df.pivot_table(index=product_kind_column, columns=result_column, aggfunc='size',
                                                  fill_value=0, observed=False)
        product_kind_pivot_table['Итог'] = product_kind_pivot_table.sum(
            axis=1)  # Добавляем столбец с итогами по строкам

        # Замена названий в индексе
        type_pivot_table.rename(index={'Залоговый': 'Всего залоговые', 'Беззалоговый': 'Всего без залоговые'},
                                inplace=True)
        product_kind_pivot_table.rename(index={'Залоговый': 'Всего залоговые', 'Беззалоговый': 'Всего без залоговые'},
                                        inplace=True)

        # Объединение двух таблиц
        combined_pivot_table = pd.concat([type_pivot_table, product_kind_pivot_table])
        combined_pivot_table.index.name = None

        # Преобразование всех значений к числовому типу данных для корректного выполнения операций сложения
        combined_pivot_table = combined_pivot_table.apply(pd.to_numeric, errors='coerce').fillna(0)

        # Создание строки "Итого" с суммами
        combined_pivot_table.loc['Итого'] = combined_pivot_table.sum(numeric_only=True)
        if 'Всего залоговые' in combined_pivot_table.index and 'Всего без залоговые' in combined_pivot_table.index:
            combined_pivot_table.loc['Итого', 'Итог'] = combined_pivot_table.loc[
                ['Всего залоговые', 'Всего без залоговые'], 'Итог'].sum()
        else:
            combined_pivot_table.loc['Итого', 'Итог'] = combined_pivot_table['Итог'].sum()

        # Перемещение строки "Итого" наверх
        itogo_row = combined_pivot_table.loc['Итого']
        combined_pivot_table = combined_pivot_table.drop('Итого')
        combined_pivot_table = pd.concat([pd.DataFrame(itogo_row).T, combined_pivot_table])

        return combined_pivot_table
    except Exception as e:
        print(f"Error counting unique values: {e}")
        return None


def create_dataframes_by_division(df, division_column):
    """
    Создает отдельные DataFrame для каждой уникальной категории в столбце "Деления".
    """
    try:
        if division_column not in df.columns:
            print(f"Column '{division_column}' not found in DataFrame")
            print(f"Available columns: {df.columns.tolist()}")
            return None
        unique_divisions = df[division_column].unique()
        dataframes = {division: df[df[division_column] == division] for division in unique_divisions}
        return dataframes
    except Exception as e:
        print(f"Error creating dataframes by division: {e}")
        return None


def find_ordered_types_in_report(file_path, sheet_name, unique_types):
    """
    Ищет ключевые слова в указанном файле и листе и возвращает их в порядке появления, игнорируя остальные значения.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)  # Загружаем все столбцы
        print(f"Successfully loaded {file_path} with sheet {sheet_name} for ordered types.")

        ordered_types = []
        for cell_value in df.to_numpy().flatten():
            if cell_value in unique_types and cell_value not in ordered_types:
                ordered_types.append(cell_value)
        return ordered_types
    except Exception as e:
        print(f"Error loading Excel file {file_path} on sheet {sheet_name}: {e}")
        return []


def reorder_dataframe(df, type_column, ordered_types):
    """
    Упорядочивает DataFrame в соответствии с порядком значений в столбце type_column.
    """
    df[type_column] = pd.Categorical(df[type_column], categories=ordered_types, ordered=True)
    return df.sort_values(by=type_column)


def process_dataframes(dataframes, type_column, product_kind_column, result_column):
    """
    Обрабатывает каждый DataFrame для дальнейшего использования в другом модуле.
    """
    try:
        processed_dataframes = {}
        for division, df in dataframes.items():
            pivot_table = count_unique_values(df, type_column, product_kind_column, result_column)
            if pivot_table is not None:
                pivot_table = apply_structure_and_sorting(pivot_table)  # Применяем структуру и сортировку
                processed_dataframes[division] = pivot_table
                print(f"Data for division {division} successfully processed.")
            else:
                print(f"No data found for division {division}.")

        # Добавление общего итогового DataFrame
        overall_summary = pd.DataFrame()
        for division, df in processed_dataframes.items():
            # Преобразование всех значений к числовому типу данных перед добавлением
            df = df.apply(pd.to_numeric, errors='coerce').fillna(0)
            overall_summary = overall_summary.add(df, fill_value=0)

        if not overall_summary.empty:
            overall_summary = apply_structure_and_sorting(overall_summary)
            processed_dataframes['Общий итог'] = overall_summary

        return processed_dataframes
    except Exception as e:
        print(f"Error processing dataframes: {e}")
        return None


def apply_structure_and_sorting(df):
    """
    Применяет структуру и сортировку к итоговому DataFrame.
    """
    # Преобразование всех значений к числовому типу данных для корректного выполнения операций сложения
    df = df.apply(pd.to_numeric, errors='coerce').fillna(0)

    # Создание строки "Итого" с суммами
    df.loc['Итого'] = df.sum(numeric_only=True)
    if 'Всего залоговые' in df.index and 'Всего без залоговые' in df.index:
        df.loc['Итого', 'Итог'] = df.loc[['Всего залоговые', 'Всего без залоговые'], 'Итог'].sum()
    else:
        df.loc['Итого', 'Итог'] = df['Итог'].sum()

    # Перемещение строки "Итого" наверх
    itogo_row = df.loc['Итого']
    df = df.drop('Итого')
    df = pd.concat([pd.DataFrame(itogo_row).T, df])

    # Перемещение столбца "Итог" в конец
    column_order = [
        'Дал обещание', 'Не звонили', 'Дал номер клиента', 'Клиент заграницей',
        'Обещал связаться с клиентом', 'Связался с клиентом и сообщил', 'Частично оплатил', 'Не дозвон',
        'Бросил трубку', 'Другой номер', 'Не знаком с клиентом', 'Дело в суде', 'Клиент умер',
        'Отказывается от оплаты', 'Отказывается от разговора', 'Итог'
    ]

    cols = [col for col in column_order if col in df.columns]
    df = df[cols]

    return df


def main(input_file, sheet_name, division_column, type_column, product_kind_column, result_column, report_file,
         report_sheet, header_row=0):
    # Шаг 1: Загрузка данных из основного Excel файла
    df = load_excel(input_file, sheet_name, header_row)
    if df is None:
        return None

    print(f"Loaded DataFrame with columns: {df.columns}")

    # Шаг 2: Получение списка уникальных значений "ТИП кредита"
    unique_types = df[type_column].dropna().unique().tolist()
    print(f"Unique types in the base data: {unique_types}")

    # Шаг 3: Получение упорядоченного списка "ТИП кредита" из отчета
    ordered_types = find_ordered_types_in_report(report_file, report_sheet, unique_types)
    if not ordered_types:
        print(f"No ordered types found in report {report_file} on sheet {report_sheet}.")
        return None

    print(f"Ordered types from the report: {ordered_types}")

    # Шаг 4: Упорядочивание основного DataFrame в соответствии с упорядоченными типами
    df = reorder_dataframe(df, type_column, ordered_types)

    # Шаг 5: Создание DataFrame для каждой уникальной категории в столбце "Деления"
    dataframes = create_dataframes_by_division(df, division_column)
    if dataframes is None:
        return None

    # Шаг 6: Обработка и возврат обработанных DataFrame
    processed_dataframes = process_dataframes(dataframes, type_column, product_kind_column, result_column)
    if processed_dataframes is not None:
        print("Dataframes processed and ready for use in another module.")
    return processed_dataframes
