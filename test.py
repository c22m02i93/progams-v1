import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import pymorphy2


def browse_file():
    global input_file
    input_file = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls")])
    if input_file:
        data_frames = load_tables_from_excel(input_file)
        display_tables(data_frames)


def load_tables_from_excel(file_path):
    data_frames = []

    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names

    for sheet_name in sheet_names:
        df = pd.read_excel(file_path, sheet_name)
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


def run_sdr():
    global input_file

    if not os.path.exists(input_file):
        print("Input file not found.")
        return

    df = pd.read_excel(input_file)

    processed_rows = process_dataframe(df)

    output_df = pd.DataFrame(processed_rows, columns=["Processed"])

    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx;*.xls")],
        title="Сохранить файл как"
    )

    if output_file:
        output_df.to_excel(output_file, index=False)
        messagebox.showinfo("Обработка завершена",
                            "Обработка таблицы завершена. Результаты сохранены в " + output_file)


def process_dataframe(df):
    processed_rows = []

    morph = pymorphy2.MorphAnalyzer()

    titles_dative = {
        "Митрополит": "Его Высокопреосвященству",
        "Архиепископ": "Его Высокопреосвященству",
        "Епископ": "Его Преосвященству"
    }

    for _, row in df.iterrows():
        cell_1 = row[2]
        cell_2 = row[3]

        if len(cell_1.split()) > 6:
            continue

        words = cell_1.split()
        title = words[0]
        region = " ".join(words[1:])

        first_name = morph.parse(cell_2)[0]
        first_name_dative_obj = first_name.inflect({'datv'})

        if first_name_dative_obj is None:
            continue  # Если слово не может быть просклонено, пропустите его

        first_name_dative = first_name_dative_obj.word.upper()

        title_dative = titles_dative.get(title)
        if title_dative is None:
            continue

        region_parts = region.split(" и ")

        region_dative_parts = []
        for part in region.split():
            inflected_word = morph.parse(part)[0].inflect({'datv'})
            if inflected_word is not None:
                region_dative_parts.append(inflected_word.word.capitalize())
            else:
                if part.lower() == "и":
                    region_dative_parts.append(part.lower())
                else:
                    region_dative_parts.append(part.capitalize())
        region_dative = " ".join(region_dative_parts)

        line_2 = f"{title_dative} {first_name_dative}"
        line_3 = f"{title}у {region_dative}"
        processed_row = f"{line_2}\n{line_3}"
        processed_rows.append(processed_row)

    return processed_rows

    morph = pymorphy2.MorphAnalyzer()

    titles_dative = {
        "Митрополит": "Его Высокопреосвященству",
        "Архиепископ": "Его Высокопреосвященству",
        "Епископ": "Его Преосвященству"
    }

    for _, row in df.iterrows():
        cell_1 = row[2]
        cell_2 = row[3]

        if len(cell_1.split()) > 6:
            continue

        words = cell_1.split()
        title = words[0]
        region = " ".join(words[1:])

        first_name = morph.parse(cell_2)[0]
        first_name_dative_obj = first_name.inflect({'datv'})

        if first_name_dative_obj is None:
            continue  # Если слово не может быть просклонено, пропустите его

        first_name_dative = first_name_dative_obj.word.upper()

        title_dative = titles_dative.get(title)
        if title_dative is None:
            continue

        region_dative = []
        for word in region.split():
            dative_word = morph.parse(word)[0].inflect({'datv'})
            if dative_word is not None:
                region_dative.append(dative_word.word.upper())
            else:
                region_dative.append(word.upper())
        region_dative = " ".join(region_dative)

        if title == "Митрополит":
            line_2 = f"Высокопреосвященнейшему {first_name_dative}"
        elif title == "Архиепископ":
            line_2 = f"Высокопреосвященнейшему {first_name_dative}"
        elif title == "Епископ":
            line_2 = f"Преосвященнейшему {first_name_dative}"
        else:
            continue

        line_3 = f"{title}у {region_dative}"
        processed_row = f"{title_dative},\n{line_2}\n{line_3}"
        processed_rows.append(processed_row)

    return processed_rows

    processed_rows = []

    morph = pymorphy2.MorphAnalyzer()

    titles_dative = {
        "Митрополит": "Его Высокопреосвященству",
        "Архиепископ": "Его Высокопреосвященству",
        "Епископ": "Его Преосвященству"
    }

    for _, row in df.iterrows():
        cell_1 = row[2]
        cell_2 = row[3]

        if len(cell_1.split()) > 6:
            continue

        words = cell_1.split()
        title = words[0]
        region = " ".join(words[1:])

        first_name = morph.parse(cell_2)[0]
        first_name_dative_obj = first_name.inflect({'datv'})

        if first_name_dative_obj is None:
            continue  # Если слово не может быть просклонено, пропустите его

        first_name_dative = first_name_dative_obj.word.upper()

        title_dative = titles_dative.get(title)
        if title_dative is None:
            continue

        region_dative = " ".join([morph.parse(word)[0].inflect(
            {'datv'}).word.capitalize() for word in region.split()])

        if title == "Митрополит":
            line_2 = f"Высокопреосвященнейшему {first_name_dative}"
        elif title == "Архиепископ":
            line_2 = f"Высокопреосвященнейшему {first_name_dative}"
        elif title == "Епископ":
            line_2 = f"Преосвященнейшему {first_name_dative}"
        else:
            continue

        line_3 = f"{title}у {region_dative}"
        processed_row = f"{title_dative},\n{line_2}\n{line_3}"
        processed_rows.append(processed_row)

    return processed_rows

    processed_rows = []

    morph = pymorphy2.MorphAnalyzer()

    for _, row in df.iterrows():
        cell_1 = row[2]
        cell_2 = row[3]

        if len(cell_1.split()) > 6:
            continue

        words = cell_1.split()
        title = words[0]
        region = " ".join(words[1:])

        first_name = morph.parse(cell_2)[0]
        first_name_dative_obj = first_name.inflect({'datv'})

        if first_name_dative_obj is None:
            continue  # Если слово не может быть просклонено, пропустите его

        first_name_dative = first_name_dative_obj.word.upper()

        title_dative = morph.parse(title)[0].inflect({'datv'}).word.upper()
        region_dative = " ".join([morph.parse(word)[0].inflect(
            {'datv'}).word.capitalize() for word in region.split()])

        if title == "Митрополит":
            line_1 = f"{title_dative},"
            line_2 = f"Высокопреосвященнейшему {first_name_dative}"
        elif title == "Архиепископ":
            line_1 = f"{title_dative},"
            line_2 = f"Высокопреосвященнейшему {first_name_dative}"
        elif title == "Епископ":
            line_1 = f"{title_dative},"
            line_2 = f"Преосвященнейшему {first_name_dative}"
        else:
            continue

        line_3 = f"{title}у {region_dative}"
        processed_row = f"{line_1}\n{line_2}\n{line_3}"
        processed_rows.append(processed_row)

    return processed_rows

    processed_rows = []

    morph = pymorphy2.MorphAnalyzer()

    for _, row in df.iterrows():
        cell_1 = row[2]
        cell_2 = row[3]

        if len(cell_1.split()) > 6:
            continue

        words = cell_1.split()
        title = words[0]
        region = " ".join(words[1:])

        first_name = morph.parse(cell_2)[0]
        first_name_dative_obj = first_name.inflect({'datv'})

        if first_name_dative_obj is None:
            continue  # Если слово не может быть просклонено, пропустите его

        first_name_dative = first_name_dative_obj.word.upper()

        if title == "Митрополит":
            line_1 = "Его Высокопреосвященству,"
            line_2 = f"Высокопреосвященнейшему {first_name_dative}"
        elif title == "Архиепископ":
            line_1 = "Его Высокопреосвященству,"
            line_2 = f"Высокопреосвященнейшему {first_name_dative}"
        elif title == "Епископ":
            line_1 = "Его Преосвященству,"
            line_2 = f"Преосвященнейшему {first_name_dative}"
        else:
            continue

        line_3 = f"{title}у {region}"
        processed_row = f"{line_1}\n{line_2}\n{line_3}"
        processed_rows.append(processed_row)

    return processed_rows


root = tk.Tk()
root.title("Загрузка таблиц из Excel-файла")

input_file = ""

browse_button = ttk.Button(
    root, text="Выберите Excel-файл", command=browse_file)
browse_button.grid(row=0, column=0, pady=10)

save_button = ttk.Button(root, text="Сохранить", command=run_sdr)
save_button.grid(row=1, column=0, pady=10)

root.mainloop()
