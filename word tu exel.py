import tkinter as tk
from tkinter import ttk, filedialog
from docx import Document
import pandas as pd


def browse_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Word files", "*.docx")])
    if file_path:
        data_frames = load_tables_from_word(file_path)
        display_tables(data_frames)


def load_tables_from_word(file_path):
    document = Document(file_path)
    tables = document.tables
    data_frames = []

    for table in tables:
        data = []

        for i, row in enumerate(table.rows):
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text)
            data.append(row_data)

        df = pd.DataFrame(data)
        data_frames.append(df)

    return data_frames


def display_tables(data_frames):
    for i, df in enumerate(data_frames):
        table_frame = ttk.Frame(root)
        table_frame.grid(row=i * 2, column=0, sticky="nsew")

        label = ttk.Label(table_frame, text=f"Таблица {i + 1}")
        label.grid(row=0, column=0)

        tree = ttk.Treeview(table_frame, columns=list(
            range(len(df.columns))), show="headings", selectmode="extended")
        tree.grid(row=1, column=0, sticky="nsew")

        for col in range(len(df.columns)):
            tree.heading(col, text=f"Столбец {col + 1}")
            tree.column(col, stretch=True, anchor='center')

        for row in df.itertuples(index=False):
            tree.insert("", "end", values=row)

        scrollbar = ttk.Scrollbar(
            table_frame, orient="vertical", command=tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)

        root.grid_rowconfigure(i * 2, weight=1)
        root.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(1, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        separator = ttk.Frame(root, height=2, relief="sunken")
        separator.grid(row=i * 2 + 1, column=0, sticky="we", pady=5)

        # Добавление поля ввода для номеров столбцов
        columns_label = ttk.Label(
            table_frame, text="Столбцы (например, 2, 4):")
        columns_label.grid(row=2, column=0, padx=5, sticky="w")

        columns_entry = ttk.Entry(table_frame, width=10)
        columns_entry.grid(row=2, column=0, padx=5)

        # Добавление кнопки сохранения таблицы с выбранными столбцами
        save_button = ttk.Button(
            table_frame, text="Сохранить таблицу", command=lambda: save_table(df, columns_entry.get()))
        save_button.grid(row=2, column=0, padx=5, pady=10, sticky="e")


def save_table(df, columns_str):
    if columns_str:
        columns = columns_str.split(',')
        columns = [int(col.strip()) - 1 for col in columns]
        selected_columns_df = df.iloc[:, columns]
        save_path = filedialog.asksaveasfilename(
            filetypes=[("Excel files", "*.xlsx")], defaultextension=".xlsx")
        if save_path:
            selected_columns_df.to_excel(save_path, index=False)


root = tk.Tk()
root.title("Загрузка таблиц из Word-файла")

browse_button = ttk.Button(
    root, text="Выберите Word-файл", command=browse_file)
browse_button.grid(row=0, column=0, pady=10)

root.mainloop()
