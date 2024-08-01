
# анализ эффективности торговых алгоритмов за 14 дней с выгрузкой в EXCEL
# в работе все криптовалютные пары  



import os
import pandas as pd
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import re

# Функция для поиска файла CSV в текущей директории
def find_csv_file():
    for file in os.listdir():
        if file.endswith(".csv"):
            return file
    raise FileNotFoundError("CSV файл не найден в текущей директории.")

# Функция для чтения и предварительной обработки данных
def preprocess_data(file_path):
    df = pd.read_csv(file_path)
    df.columns = df.columns.str.strip()
    df["Date/Time"] = df["Date/Time"].str.strip()
    df["Date/Time"] = pd.to_datetime(
        df["Date/Time"], format="%y/%m/%d %H:%M:%S", errors="coerce"
    )
    df = df.dropna(subset=["Date/Time"])
    df["Symbol"] = df["Symbol"].str.lower()
    return df

# Функция для анализа сделок за последние 14 дней
def analyze_trades_last_14_days(df, algorithm):
    last_14_days = df[df["Date/Time"] >= (pd.Timestamp.now() - pd.Timedelta(days=14))]
    last_14_days_algo = last_14_days[
        last_14_days["Opened"].str.contains(algorithm, na=False)
    ]

    trades_last_14_days = last_14_days_algo["Symbol"].value_counts().reset_index()
    trades_last_14_days.columns = ["Symbol", "Trades last 14 days"]

    positive_trades = last_14_days_algo[last_14_days_algo["PNLUSDT"] > 0]
    negative_trades = last_14_days_algo[last_14_days_algo["PNLUSDT"] < 0]

    positive_trades_count = positive_trades["Symbol"].value_counts().reset_index()
    positive_trades_count.columns = ["Symbol", "Positive Trades Count"]

    negative_trades_count = negative_trades["Symbol"].value_counts().reset_index()
    negative_trades_count.columns = ["Symbol", "Negative Trades Count"]

    positive_trades_sum = (
        positive_trades.groupby("Symbol")["PNLUSDT"].sum().reset_index()
    )
    positive_trades_sum.columns = ["Symbol", "Positive Trades Sum"]

    negative_trades_sum = (
        negative_trades.groupby("Symbol")["PNLUSDT"].sum().reset_index()
    )
    negative_trades_sum.columns = ["Symbol", "Negative Trades Sum"]

    return (
        trades_last_14_days,
        positive_trades_count,
        negative_trades_count,
        positive_trades_sum,
        negative_trades_sum,
    )

# Функция для создания итоговой таблицы
def create_final_table(
    algorithm,
    trades_last_14_days,
    positive_trades_sum,
    positive_trades_count,
    negative_trades_sum,
    negative_trades_count,
):
    final_result = trades_last_14_days.merge(
        positive_trades_sum, on="Symbol", how="left"
    )
    final_result = final_result.merge(positive_trades_count, on="Symbol", how="left")
    final_result = final_result.merge(negative_trades_sum, on="Symbol", how="left")
    final_result = final_result.merge(negative_trades_count, on="Symbol", how="left")
    final_result["Positive Trades Sum"] = final_result["Positive Trades Sum"].fillna(0)
    final_result["Negative Trades Sum"] = final_result["Negative Trades Sum"].fillna(0)
    final_result["Max PNLUSDT"] = (
        final_result["Positive Trades Sum"] + final_result["Negative Trades Sum"]
    )
    final_result = final_result[
        [
            "Symbol",
            "Max PNLUSDT",
            "Trades last 14 days",
            "Positive Trades Sum",
            "Positive Trades Count",
            "Negative Trades Sum",
            "Negative Trades Count",
        ]
    ]
    # Сортировка по столбцу Max PNLUSDT от большего к меньшему
    final_result = final_result.sort_values(by="Max PNLUSDT", ascending=False)
    return final_result

# Функция для анализа алгоритмов
def analyze_algorithm(df, algorithm):
    (
        trades_last_14_days,
        positive_trades_count,
        negative_trades_count,
        positive_trades_sum,
        negative_trades_sum,
    ) = analyze_trades_last_14_days(df, algorithm)
    final_result = create_final_table(
        algorithm,
        trades_last_14_days,
        positive_trades_sum,
        positive_trades_count,
        negative_trades_sum,
        negative_trades_count,
    )
    return final_result

# Функция для установки ширины столбцов по содержимому
def adjust_column_width(worksheet):
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column].width = adjusted_width

# Функция для сохранения данных в Excel и применения форматирования
def save_to_excel(algo_results, filename="tests_14.xlsx"):
    def sanitize_sheet_name(name):
        # Удаление всех недопустимых символов и сокращение длины до 31 символа
        name = re.sub(r"[^0-9A-Za-z_]", "", name)
        return name[:31]

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        for algo, result in algo_results.items():
            sanitized_algo = sanitize_sheet_name(algo)
            result.to_excel(writer, sheet_name=sanitized_algo, index=False)
            worksheet = writer.sheets[sanitized_algo]
            worksheet.auto_filter.ref = worksheet.dimensions
            worksheet.freeze_panes = worksheet["A2"]

            # Установка ширины столбцов по содержимому
            adjust_column_width(worksheet)

# Главная функция
def main():
    try:
        csv_file_path = find_csv_file()
        df = preprocess_data(csv_file_path)

        algorithms = df["Opened"].unique()
        algo_results = {}

        for algo in algorithms:
            if "test" in algo:
                final_result = analyze_algorithm(df, algo)
                algo_results[algo] = final_result

        if algo_results:
            save_to_excel(algo_results)
        else:
            print("Алгоритмы со словом 'test' не найдены.")

    except Exception as e:
        print(f"Произошла ошибка: {e}")

if __name__ == "__main__":
    main()
