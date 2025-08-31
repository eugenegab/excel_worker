import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from cast_exceptions import FieldNotFound
from excel_processor import ExcelProcessor  # твой класс


def choose_file(entry_field: tk.Entry) -> None:
    """Открывает диалог выбора файла и записывает путь в поле ввода"""
    filepath = filedialog.askopenfilename(
        title="Выберите Excel-файл",
        filetypes=(("Excel файлы", "*.xlsx"),)
    )
    if filepath:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, filepath)


def enter_new_filepath(entry_field: tk.Entry) -> None:
    """Открывает диалог выбора пути и имени для нового файла"""
    filepath = filedialog.asksaveasfilename(
        title="Сохранить файл как...",
        defaultextension=".xlsx",
        filetypes=(("Excel файлы", "*.xlsx"),)
    )
    if filepath:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, filepath)


def run_processor(file_entry: tk.Entry, attr_entry: tk.Entry, value_entry: tk.Entry, output_entry: tk.Entry, fields: tk.Entry) -> None:
    """Запускает обработку Excel файла"""
    filepath = file_entry.get().strip()
    attr = attr_entry.get().strip()
    value = value_entry.get().strip()
    output_file = output_entry.get().strip()
    fields = [field.strip().lower() for field in fields.get().strip().split(',')]

    if not filepath:
        messagebox.showerror("Ошибка", "Вы не выбрали файл!")
        return
    elif not output_file:
        messagebox.showerror("Ошибка", "Вы не ввели путь и имя конечного файла!")
        return
    elif not attr:
        messagebox.showerror("Ошибка", "Вы не ввели название поля!")
        return
    elif not value:
        messagebox.showerror("Ошибка", "Вы не ввели значение для фильтрации!")
        return

    try:
        processor = ExcelProcessor(filepath=filepath, field_name=attr, value=value, output_path=output_file, wanted_fields=fields)
        message = processor.process()
        messagebox.showinfo(message, f"Файл обработан и сохранён в:\n{processor.output_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))


def create_gui():
    """Создание интерфейса"""
    root = tk.Tk()
    root.title("Excel Processor")
    root.geometry("600x250")
    root.resizable(False, False)

    # --- путь к файлу ---
    tk.Label(root, text="Файл Excel:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    file_entry = tk.Entry(root, width=40)
    file_entry.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Выбрать...", command=lambda: choose_file(file_entry)).grid(row=0, column=2, padx=5, pady=5)

    # --- имя файла с результатом ---
    tk.Label(root, text="Введите название нового файла").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    output_file = tk.Entry(root, width=40)
    output_file.grid(row=1, column=1, padx=5, pady=5, columnspan=2)
    tk.Button(root, text="Выбрать...", command=lambda: enter_new_filepath(output_file)).grid(row=1, column=2, padx=5, pady=5)

    # --- Необходимые поля таблицы ---
    tk.Label(root, text="Поля таблицы через запятую:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
    fields_entry = tk.Entry(root, width=40)
    fields_entry.grid(row=2, column=1, padx=5, pady=5, columnspan=2)

    # --- атрибут ---
    tk.Label(root, text="Атрибут:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
    attr_entry = tk.Entry(root, width=40)
    attr_entry.grid(row=3, column=1, padx=5, pady=5, columnspan=2)
    # combo = ttk.Combobox(root, values=ExcelProcessor.WANTED_FIELDS, state="readonly")
    # combo.grid(row=2, column=1, padx=5, pady=5, columnspan=2)
    # combo.current(0)

    # --- значение ---
    tk.Label(root, text="Значение:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
    value_entry = tk.Entry(root, width=40)
    value_entry.grid(row=4, column=1, padx=5, pady=5, columnspan=2)

    # --- кнопка запуска ---
    tk.Button(root, text="Запустить", command=lambda: run_processor(file_entry, attr_entry, value_entry, output_file, fields_entry), bg="lightgreen").grid(
        row=5, column=0, columnspan=3, pady=15
    )

    root.mainloop()


if __name__ == "__main__":
    create_gui()
