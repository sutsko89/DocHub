import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from docx import Document
import os
import subprocess
import platform
import re
import sys


def resource_path(relative_path):
    """ Получает абсолютный путь к ресурсу, работает для разработки и для PyInstaller """
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ---------- Настройки ----------
BG_COLOR = "#222222"
TEXT_COLOR = "#ffffff"
BUTTON_COLOR = "#ff7f27"
FONT = ("Segoe UI", 10)

root = tk.Tk()
root.title("Генератор договора поставки")
root.geometry("900x600")
root.configure(bg=BG_COLOR)
root.resizable(False, False)


# ---------- Функции валидации ----------
def validate_ogrn(ogrn):
    """Проверка ОГРН: только цифры, 13 символов"""
    if not ogrn:
        return False
    return ogrn.isdigit() and len(ogrn) == 13


def validate_ogrnip(ogrnip):
    """Проверка ОГРНИП: только цифры, 15 символов"""
    if not ogrnip:
        return False
    return ogrnip.isdigit() and len(ogrnip) == 15


def validate_inn(inn, is_ip=False):
    """Проверка ИНН: для ООО - 10 цифр, для ИП - 12 цифр"""
    if not inn:
        return False
    if is_ip:
        return inn.isdigit() and len(inn) == 12
    else:
        return inn.isdigit() and len(inn) == 10


def validate_kpp(kpp):
    """Проверка КПП: только цифры, 9 символов"""
    if not kpp:
        return False
    return kpp.isdigit() and len(kpp) == 9


def validate_email(email):
    """Проверка email по регулярному выражению"""
    if not email:
        return False
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None


def validate_phone(phone):
    """Проверка телефона: может содержать цифры и любые знаки, но не буквы"""
    if not phone:
        return False
    cleaned = re.sub(r'[\d\+\-\(\)\s\/\.]', '', phone)
    return cleaned == ''


def validate_account(account):
    """Проверка расчетного счета: 20 цифр"""
    if not account:
        return False
    return account.isdigit() and len(account) == 20


def validate_correspondent_account(account):
    """Проверка корреспондентского счета: 20 цифр"""
    if not account:
        return False
    return account.isdigit() and len(account) == 20


def validate_bik(bik):
    """Проверка БИК: 9 цифр"""
    if not bik:
        return False
    return bik.isdigit() and len(bik) == 9


def validate_all_fields(data, is_ip=False):
    """Проверка всех полей на корректность (для основного договора)"""
    errors = []

    if is_ip:
        if not validate_ogrnip(data.get("ОГРНИП", "")):
            errors.append("ОГРНИП должен содержать 15 цифр")
        if not validate_inn(data.get("ИНН", ""), is_ip=True):
            errors.append("ИНН ИП должен содержать 12 цифр")
    else:
        if not validate_ogrn(data.get("ОГРН", "")):
            errors.append("ОГРН должен содержать 13 цифр")
        if not validate_inn(data.get("ИНН", ""), is_ip=False):
            errors.append("ИНН должен содержать 10 цифр")
        if not validate_kpp(data.get("КПП", "")):
            errors.append("КПП должен содержать 9 цифр")

    if not validate_email(data.get("Email", "")):
        errors.append("Email имеет неверный формат")
    if not validate_phone(data.get("Телефон", "")):
        errors.append("Телефон может содержать только цифры и знаки: + - ( ) / . и пробелы")
    if not validate_account(data.get("Расчётный счёт", "")):
        errors.append("Расчётный счёт должен содержать 20 цифр")
    if not validate_correspondent_account(data.get("Корреспондентский счёт", "")):
        errors.append("Корреспондентский счёт должен содержать 20 цифр")
    if not validate_bik(data.get("БИК", "")):
        errors.append("БИК должен содержать 9 цифр")

    return errors


def validate_additional_agreement_required(data, is_ip=False, limit="", term=""):
    """Минимальная проверка для доп. соглашения."""
    errors = []

    if not data.get("Номер договора", "").strip():
        errors.append("Заполните 'Номер договора'")
    if not data.get("Дата (число)", "").strip():
        errors.append("Заполните 'Дата (число)'")
    if not data.get("Месяц (словом)", "").strip():
        errors.append("Заполните 'Месяц (словом)'")

    if is_ip:
        if not data.get("ФИО ИП", "").strip():
            errors.append("Заполните 'ФИО ИП'")
    else:
        if not data.get("Полное название компании (Покупь)", "").strip():
            errors.append("Заполните 'Полное название компании (Покупь)'")
        if not data.get("Сокращенное название компании", "").strip():
            errors.append("Заполните 'Сокращенное название компании'")
        if not data.get("ФИО директора", "").strip():
            errors.append("Заполните 'ФИО директора'")

    if not str(limit).strip():
        errors.append("Заполните 'Лимит'")
    if not str(term).strip():
        errors.append("Заполните 'Отсрочка платежа'")
    elif not str(term).strip().isdigit():
        errors.append("'Отсрочка платежа' должна быть числом")

    return errors


# ---------- Функции преобразования ФИО ----------
def format_director_name(full_name):
    """Преобразует ФИО из формата 'Фамилия Имя Отчество' в формат 'И.О. Фамилия'"""
    parts = full_name.strip().split()
    if len(parts) < 2:
        return full_name
    surname = parts[0]
    initials = ""
    for i in range(1, min(3, len(parts))):
        if parts[i]:
            initials += parts[i][0] + "."
    return f"{initials} {surname}"


def format_name_ip2(full_name):
    """Преобразует ФИО из формата 'Фамилия Имя Отчество' в формат 'Фамилия И.О.'"""
    parts = full_name.strip().split()
    if len(parts) < 2:
        return full_name
    surname = parts[0]
    initials = ""
    for i in range(1, min(3, len(parts))):
        if parts[i]:
            initials += parts[i][0] + "."
    return f"{surname} {initials}"


def format_name_ip3(full_name):
    """Преобразует ФИО из формата 'Фамилия Имя Отчество' в формат 'И.О. Фамилия'"""
    return format_director_name(full_name)


# ---------- Функция форматирования даты ----------
def format_date(day):
    """Форматирует число дня: добавляет ведущий ноль для чисел меньше 10"""
    day_str = str(day).strip()
    if day_str.isdigit():
        day_num = int(day_str)
        return f"{day_num:02d}"
    return day_str


# ---------- Числа прописью (RU) ----------
def number_to_words_ru(n: int) -> str:
    """Кардинальные числительные на русском (0..9999). Для term_in_words достаточно."""
    if n < 0:
        return "минус " + number_to_words_ru(-n)

    units = [
        "ноль", "один", "два", "три", "четыре",
        "пять", "шесть", "семь", "восемь", "девять"
    ]
    teens = [
        "десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать",
        "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"
    ]
    tens = [
        "", "", "двадцать", "тридцать", "сорок",
        "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"
    ]
    hundreds = [
        "", "сто", "двести", "триста", "четыреста",
        "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"
    ]

    if n < 10:
        return units[n]
    if 10 <= n < 20:
        return teens[n - 10]
    if 20 <= n < 100:
        t = n // 10
        u = n % 10
        return tens[t] if u == 0 else f"{tens[t]} {units[u]}"
    if 100 <= n < 1000:
        h = n // 100
        rest = n % 100
        return hundreds[h] if rest == 0 else f"{hundreds[h]} {number_to_words_ru(rest)}"
    if 1000 <= n < 10000:
        th = n // 1000
        rest = n % 1000

        if th == 1:
            th_part = "тысяча"
        elif th in (2, 3, 4):
            th_part = f"{units[th]} тысячи"
        else:
            th_part = f"{units[th] if th < 10 else number_to_words_ru(th)} тысяч"

        return th_part if rest == 0 else f"{th_part} {number_to_words_ru(rest)}"

    return str(n)


# ---------- Функция для открытия файла в зависимости от ОС ----------
def open_file(file_path):
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)  # noqa: S606
        elif platform.system() == "Darwin":
            subprocess.run(["open", file_path], check=False)
        else:
            subprocess.run(["xdg-open", file_path], check=False)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")


# ---------- DOCX: замена плейсхолдеров (в т.ч. если разбиты на runs) ----------
def replace_in_paragraph(paragraph, replacements: dict):
    runs = paragraph.runs
    if not runs:
        return

    def rebuild_full_text():
        return "".join(r.text for r in runs)

    full_text = rebuild_full_text()

    for key, value in replacements.items():
        key = str(key)
        value = "" if value is None else str(value)

        while key and (key in full_text):
            start = full_text.find(key)
            end = start + len(key)

            cur = 0
            start_run = None
            start_offset = 0
            end_run = None
            end_offset = 0

            for i, r in enumerate(runs):
                rt = r.text or ""
                next_cur = cur + len(rt)
                if start_run is None and start < next_cur:
                    start_run = i
                    start_offset = start - cur
                if end_run is None and end <= next_cur:
                    end_run = i
                    end_offset = end - cur
                    break
                cur = next_cur

            if start_run is None or end_run is None:
                break

            start_text = runs[start_run].text or ""
            end_text = runs[end_run].text or ""

            prefix = start_text[:start_offset]
            suffix = end_text[end_offset:]

            runs[start_run].text = prefix + value + suffix

            for j in range(start_run + 1, end_run + 1):
                runs[j].text = ""

            full_text = rebuild_full_text()


def replace_in_doc(doc, replacements: dict):
    for p in doc.paragraphs:
        replace_in_paragraph(p, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, replacements)

    for section in doc.sections:
        header = section.header
        footer = section.footer

        for p in header.paragraphs:
            replace_in_paragraph(p, replacements)
        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        replace_in_paragraph(p, replacements)

        for p in footer.paragraphs:
            replace_in_paragraph(p, replacements)
        for table in footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        replace_in_paragraph(p, replacements)


# ---------- Переменная для хранения типа контрагента ----------
is_ip = tk.BooleanVar(value=False)

# Глобальные переменные для активных полей даты
active_date_entry = None
active_month_entry = None
ooo_date_entry = None
ooo_month_entry = None
ip_date_entry = None
ip_month_entry = None

# ВАЖНО: раздельные словари, чтобы поля ИП не перезаписывали поля ООО
entries_ooo = {}
entries_ip = {}

labels = {}
ooo_widgets = []
ip_widgets = []


def get_active_entries():
    return entries_ip if is_ip.get() else entries_ooo


# ---------- Функция переключения между ООО и ИП ----------
def toggle_contractor_type():
    global active_date_entry, active_month_entry
    if is_ip.get():
        for widget in ooo_widgets:
            widget.grid_remove()
        for widget in ip_widgets:
            widget.grid()
        if "Фактический адрес" in labels and "Фактический адрес" in entries_ooo:
            labels["Фактический адрес"].grid_remove()
            entries_ooo["Фактический адрес"].grid_remove()
        active_date_entry = ip_date_entry
        active_month_entry = ip_month_entry
    else:
        for widget in ip_widgets:
            widget.grid_remove()
        for widget in ooo_widgets:
            widget.grid()
        if "Фактический адрес" in labels and "Фактический адрес" in entries_ooo:
            labels["Фактический адрес"].grid()
            entries_ooo["Фактический адрес"].grid()
        active_date_entry = ooo_date_entry
        active_month_entry = ooo_month_entry


# ---------- Основные функции ----------
def fill_current_date():
    today = datetime.now()
    day = today.day
    months = [
        "января", "февраля", "марта", "апреля", "мая", "июня",
        "июля", "августа", "сентября", "октября", "ноября", "декабря"
    ]
    month = months[today.month - 1]
    active_date_entry.delete(0, tk.END)
    active_date_entry.insert(0, format_date(day))
    active_month_entry.delete(0, tk.END)
    active_month_entry.insert(0, month)


def collect_data_from_active_fields():
    active_entries = get_active_entries()
    field_list = ip_fields if is_ip.get() else ooo_fields

    data = {}
    for label in field_list:
        widget = active_entries.get(label)
        if widget is None:
            continue
        if isinstance(widget, ttk.Combobox):
            data[label] = widget.get().strip()
        else:
            data[label] = widget.get().strip()

    if is_ip.get():
        data["Фактический адрес"] = data.get("Юридический адрес", "")

    return data


def build_replacements_for_contract(data: dict):
    formatted_date = format_date(data.get("Дата (число)", ""))

    if is_ip.get():
        name_ip = data.get("ФИО ИП", "")
        name_ip2 = format_name_ip2(name_ip)
        name_ip3 = format_name_ip3(name_ip)

        replacements = {
            "{{number}}": data.get("Номер договора", ""),
            "{{date}}": formatted_date,
            "{{mouth}}": data.get("Месяц (словом)", ""),
            "{{name_IP}}": name_ip,
            "{{name_IP2}}": name_ip2,
            "{{name_IP3}}": name_ip3,
            "{{OGRNIP}}": data.get("ОГРНИП", ""),
            "{{INN}}": data.get("ИНН", ""),
            "{{Urid_adress}}": data.get("Юридический адрес", ""),
            "{{fact_adress}}": data.get("Фактический адрес", ""),
            "{{E-mail}}": data.get("Email", ""),
            "{{Telephone}}": data.get("Телефон", ""),
            "{{r/schet}}": data.get("Расчётный счёт", ""),
            "{{bank}}": data.get("Название банка", ""),
            "{{k/schet}}": data.get("Корреспондентский счёт", ""),
            "{{BIK}}": data.get("БИК", "")
        }
    else:
        basis = data.get("Основание действий", "")
        if basis == "Устав":
            basis_text = "устава"
        elif basis == "Доверенность":
            basis_text = "доверенности"
        else:
            basis_text = ""

        director_name_short = format_director_name(data.get("ФИО директора", ""))
        full_company_name = data.get("Полное название компании (Покупь)", "")
        short_company_name = data.get("Сокращенное название компании", "")

        replacements = {
            "{{number}}": data.get("Номер договора", ""),
            "{{date}}": formatted_date,
            "{{mouth}}": data.get("Месяц (словом)", ""),
            "{{long_name_company}}": full_company_name,
            "{{short_name_company}}": short_company_name,
            "{{direct_name}}": data.get("ФИО директора", ""),
            "{{direct_name2}}": director_name_short,
            "{{ustava}}": basis_text,
            "{{OGRN}}": data.get("ОГРН", ""),
            "{{INN}}": data.get("ИНН", ""),
            "{{KPP}}": data.get("КПП", ""),
            "{{Urid_adress}}": data.get("Юридический адрес", ""),
            "{{fact_adress}}": data.get("Фактический адрес", ""),
            "{{E-mail}}": data.get("Email", ""),
            "{{Telephone}}": data.get("Телефон", ""),
            "{{r/schet}}": data.get("Расчётный счёт", ""),
            "{{bank}}": data.get("Название банка", ""),
            "{{k/schet}}": data.get("Корреспондентский счёт", ""),
            "{{BIK}}": data.get("БИК", "")
        }

    return replacements


def save_contract():
    try:
        active_entries = get_active_entries()
        contract_number = active_entries.get("Номер договора").get().strip() if active_entries.get("Номер договора") else ""

        if not contract_number:
            messagebox.showerror("Ошибка", "Пожалуйста, заполните поле 'Номер договора'")
            return

        template_path = resource_path("dogovor_template_ip.docx") if is_ip.get() else resource_path("dogovor_template.docx")
        if not os.path.exists(template_path):
            messagebox.showerror("Ошибка", f"Не найден шаблон:\n{template_path}")
            return

        data = collect_data_from_active_fields()

        validation_errors = validate_all_fields(data, is_ip.get())
        if validation_errors:
            error_message = "Обнаружены ошибки в данных:\n\n" + "\n".join(f"• {error}" for error in validation_errors)
            messagebox.showerror("Ошибка в данных", error_message)
            return

        try:
            doc = Document(template_path)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть шаблон:\n{e}")
            return

        replacements = build_replacements_for_contract(data)
        replace_in_doc(doc, replacements)

        if is_ip.get():
            short_name = data.get("ФИО ИП", "").strip()
        else:
            short_name = data.get("Сокращенное название компании", "").strip()

        if not short_name:
            short_name = f"№{contract_number}"
        else:
            invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
            for char in invalid_chars:
                short_name = short_name.replace(char, '')
            short_name = short_name.strip()

        default_filename = f"Договор поставки {short_name}.docx"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Документы Word", "*.docx"), ("Все файлы", "*.*")],
            initialfile=default_filename,
            title="Сохранить договор как..."
        )
        if not file_path:
            return

        try:
            doc.save(file_path)
            open_file(file_path)
            messagebox.showinfo("Успех", f"Договор сохранён и открыт:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка при сохранении", f"Не удалось сохранить договор:\n{e}")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить договор:\n{e}")


# ---------- Доп. соглашение ----------
def open_additional_agreement_window():
    win = tk.Toplevel(root)
    win.title("Создать доп. соглашение")
    win.geometry("520x220")
    win.configure(bg=BG_COLOR)
    win.resizable(False, False)
    win.grab_set()

    # Плейсхолдеры в UI убраны
    tk.Label(win, text="Лимит:", bg=BG_COLOR, fg=TEXT_COLOR, font=FONT).pack(anchor="w", padx=15, pady=(15, 5))
    limit_entry = tk.Entry(win, font=FONT, width=45)
    limit_entry.pack(anchor="w", padx=15)

    tk.Label(win, text="Отсрочка платежа, дней:", bg=BG_COLOR, fg=TEXT_COLOR, font=FONT).pack(anchor="w", padx=15, pady=(15, 5))
    term_var = tk.StringVar()
    term_entry = tk.Entry(win, font=FONT, width=45, textvariable=term_var)
    term_entry.pack(anchor="w", padx=15)

    term_words_label = tk.Label(win, text="Прописью: ", bg=BG_COLOR, fg=TEXT_COLOR, font=FONT)
    term_words_label.pack(anchor="w", padx=15, pady=(12, 0))

    def update_term_words(*args):
        val = term_var.get().strip()
        if val.isdigit():
            n = int(val)
            term_words_label.config(text=f"Прописью: {number_to_words_ru(n)}")
        elif val == "":
            term_words_label.config(text="Прописью: ")
        else:
            term_words_label.config(text="Прописью: (введите число)")

    term_var.trace_add("write", update_term_words)
    update_term_words()

    btn_frame = tk.Frame(win, bg=BG_COLOR)
    btn_frame.pack(side="bottom", fill="x", pady=15)

    btn_sub = tk.Frame(btn_frame, bg=BG_COLOR)
    btn_sub.pack(expand=True)

    def generate_additional_agreement():
        try:
            limit = limit_entry.get().strip()
            term = term_entry.get().strip()

            data = collect_data_from_active_fields()

            errs = validate_additional_agreement_required(data, is_ip.get(), limit=limit, term=term)
            if errs:
                messagebox.showerror("Ошибка", "Исправьте:\n\n" + "\n".join(f"• {e}" for e in errs), parent=win)
                return

            template_path = resource_path("additional_agreement_ip.docx") if is_ip.get() else resource_path("additional_agreement.docx")
            if not os.path.exists(template_path):
                messagebox.showerror("Ошибка", f"Не найден шаблон доп. соглашения:\n{template_path}", parent=win)
                return

            try:
                doc = Document(template_path)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось открыть шаблон:\n{e}", parent=win)
                return

            replacements = build_replacements_for_contract(data)

            term_int = int(term)
            replacements.update({
                "{{limit}}": limit,
                "{{term}}": str(term_int),
                "{{term_in_words}}": number_to_words_ru(term_int)
            })

            replace_in_doc(doc, replacements)

            contract_number = data.get("Номер договора", "").strip()
            if is_ip.get():
                short_name = data.get("ФИО ИП", "").strip()
            else:
                short_name = data.get("Сокращенное название компании", "").strip()

            if not short_name:
                short_name = f"№{contract_number}"
            else:
                invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
                for char in invalid_chars:
                    short_name = short_name.replace(char, '')
                short_name = short_name.strip()

            default_filename = f"Доп. соглашение {short_name}.docx"

            file_path = filedialog.asksaveasfilename(
                parent=win,
                defaultextension=".docx",
                filetypes=[("Документы Word", "*.docx"), ("Все файлы", "*.*")],
                initialfile=default_filename,
                title="Сохранить доп. соглашение как..."
            )
            if not file_path:
                return

            try:
                doc.save(file_path)
                open_file(file_path)
                messagebox.showinfo("Успех", f"Доп. соглашение сохранено и открыто:\n{file_path}", parent=win)
                win.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка при сохранении", f"Не удалось сохранить файл:\n{e}", parent=win)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать доп. соглашение:\n{e}", parent=win)

    tk.Button(
        btn_sub,
        text="Создать",
        command=generate_additional_agreement,
        bg=BUTTON_COLOR,
        fg="white",
        font=FONT,
        width=16
    ).pack(side="left", padx=10)

    tk.Button(
        btn_sub,
        text="Отмена",
        command=win.destroy,
        bg="#555555",
        fg="white",
        font=FONT,
        width=16
    ).pack(side="left", padx=10)

    limit_entry.focus_set()


# ---------- Создание интерфейса ----------
main_frame = tk.Frame(root, bg=BG_COLOR)
main_frame.pack(fill="both", expand=True, padx=15, pady=10)

type_frame = tk.Frame(main_frame, bg=BG_COLOR)
type_frame.pack(fill="x", pady=(0, 10))

tk.Label(type_frame, text="Тип контрагента:", bg=BG_COLOR, fg=TEXT_COLOR, font=FONT).pack(side="left", padx=(0, 10))
tk.Radiobutton(
    type_frame, text="Юр. лицо", variable=is_ip, value=False, command=toggle_contractor_type,
    bg=BG_COLOR, fg=TEXT_COLOR, selectcolor=BG_COLOR, font=FONT
).pack(side="left", padx=(0, 10))
tk.Radiobutton(
    type_frame, text="ИП", variable=is_ip, value=True, command=toggle_contractor_type,
    bg=BG_COLOR, fg=TEXT_COLOR, selectcolor=BG_COLOR, font=FONT
).pack(side="left")

fields_frame = tk.Frame(main_frame, bg=BG_COLOR)
fields_frame.pack(fill="both", expand=True)

def validate_digits_input(new_value):
    return new_value == "" or new_value.isdigit()

def validate_phone_input(new_value):
    if new_value == "":
        return True
    cleaned = re.sub(r'[\d\+\-\(\)\s\/\.]', '', new_value)
    return cleaned == "" or not any(c.isalpha() for c in cleaned)

digits_validator = root.register(validate_digits_input)
phone_validator = root.register(validate_phone_input)

# ---------- Поля для ООО ----------
ooo_fields = [
    "Номер договора",
    "Дата (число)",
    "Месяц (словом)",
    "Полное название компании (Покупь)",
    "Сокращенное название компании",
    "ФИО директора",
    "Основание действий",
    "Юридический адрес",
    "Фактический адрес",
    "Email",
    "Телефон",
    "ИНН",
    "КПП",
    "ОГРН",
    "Расчётный счёт",
    "Корреспондентский счёт",
    "Название банка",
    "БИК"
]

row = 0
for label in ooo_fields:
    label_widget = tk.Label(fields_frame, text=label, bg=BG_COLOR, fg=TEXT_COLOR, font=FONT, width=30, anchor="w")
    label_widget.grid(row=row, column=0, sticky="w", pady=2)
    labels[label] = label_widget

    if label == "Основание действий":
        combo = ttk.Combobox(fields_frame, values=["Устав", "Доверенность"], font=FONT, width=47, state="readonly")
        combo.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        combo.set("Устав")
        entries_ooo[label] = combo
        ooo_widgets.extend([label_widget, combo])
        ref_widget = combo

    elif label == "Телефон":
        entry = tk.Entry(fields_frame, font=FONT, width=50, validate="key", validatecommand=(phone_validator, '%P'))
        entry.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        entries_ooo[label] = entry
        ooo_widgets.extend([label_widget, entry])
        ref_widget = entry

    elif label in ["ИНН", "КПП", "ОГРН", "Расчётный счёт", "Корреспондентский счёт", "БИК"]:
        entry = tk.Entry(fields_frame, font=FONT, width=50, validate="key", validatecommand=(digits_validator, '%P'))
        entry.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        entries_ooo[label] = entry
        ooo_widgets.extend([label_widget, entry])
        ref_widget = entry

    else:
        entry = tk.Entry(fields_frame, font=FONT, width=50)
        entry.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        entries_ooo[label] = entry
        ooo_widgets.extend([label_widget, entry])
        ref_widget = entry

    if label == "Дата (число)":
        ooo_date_entry = ref_widget
    elif label == "Месяц (словом)":
        ooo_month_entry = ref_widget

    row += 1

# ---------- Поля для ИП ----------
ip_fields = [
    "Номер договора",
    "Дата (число)",
    "Месяц (словом)",
    "ФИО ИП",
    "Юридический адрес",
    "Email",
    "Телефон",
    "ИНН",
    "ОГРНИП",
    "Расчётный счёт",
    "Корреспондентский счёт",
    "Название банка",
    "БИК"
]

ip_row_mapping = {
    "Номер договора": 0,
    "Дата (число)": 1,
    "Месяц (словом)": 2,
    "ФИО ИП": 3,
    "Юридический адрес": 7,
    "Email": 9,
    "Телефон": 10,
    "ИНН": 11,
    "ОГРНИП": 13,
    "Расчётный счёт": 14,
    "Корреспондентский счёт": 15,
    "Название банка": 16,
    "БИК": 17
}

for label in ip_fields:
    target_row = ip_row_mapping[label]

    label_widget = tk.Label(fields_frame, text=label, bg=BG_COLOR, fg=TEXT_COLOR, font=FONT, width=30, anchor="w")
    label_widget.grid(row=target_row, column=0, sticky="w", pady=2)
    labels[label] = label_widget

    if label == "Телефон":
        entry = tk.Entry(fields_frame, font=FONT, width=50, validate="key", validatecommand=(phone_validator, '%P'))
    elif label in ["ИНН", "ОГРНИП", "Расчётный счёт", "Корреспондентский счёт", "БИК"]:
        entry = tk.Entry(fields_frame, font=FONT, width=50, validate="key", validatecommand=(digits_validator, '%P'))
    else:
        entry = tk.Entry(fields_frame, font=FONT, width=50)

    entry.grid(row=target_row, column=1, padx=10, pady=2, sticky="w")
    entries_ip[label] = entry

    if label == "Дата (число)":
        ip_date_entry = entry
    elif label == "Месяц (словом)":
        ip_month_entry = entry

    ip_widgets.extend([label_widget, entry])
    label_widget.grid_remove()
    entry.grid_remove()

for widget in ip_widgets:
    widget.grid_remove()

active_date_entry = ooo_date_entry
active_month_entry = ooo_month_entry

# ---------- Кнопки ----------
separator = tk.Frame(root, height=2, bg="#444444")
separator.pack(fill="x", padx=10, pady=5)

button_frame = tk.Frame(root, bg=BG_COLOR)
button_frame.pack(side="bottom", fill="x", pady=10)

button_subframe = tk.Frame(button_frame, bg=BG_COLOR)
button_subframe.pack(expand=True)

tk.Button(
    button_subframe,
    text="Автозаполнить дату",
    command=fill_current_date,
    bg=BUTTON_COLOR,
    fg="white",
    font=FONT,
    width=20
).pack(side="left", padx=10)

tk.Button(
    button_subframe,
    text="Сохранить договор",
    command=save_contract,
    bg=BUTTON_COLOR,
    fg="white",
    font=FONT,
    width=20
).pack(side="left", padx=10)

tk.Button(
    button_subframe,
    text="Создать доп. соглашение",
    command=open_additional_agreement_window,
    bg=BUTTON_COLOR,
    fg="white",
    font=FONT,
    width=24
).pack(side="left", padx=10)


# ---------- Универсальная поддержка копирования/вставки ----------
def copy_text(event=None):
    try:
        widget = root.focus_get()
        if isinstance(widget, (tk.Entry, ttk.Combobox)):
            if widget.selection_present():
                selected_text = widget.selection_get()
                root.clipboard_clear()
                root.clipboard_append(selected_text)
        return "break"
    except Exception:
        return "break"


def paste_text(event=None):
    try:
        widget = root.focus_get()
        if isinstance(widget, (tk.Entry, ttk.Combobox)):
            clipboard_text = root.clipboard_get()
            try:
                widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except Exception:
                pass
            widget.insert(tk.INSERT, clipboard_text)
    except Exception:
        pass
    return "break"


def cut_text(event=None):
    try:
        widget = root.focus_get()
        if isinstance(widget, (tk.Entry, ttk.Combobox)):
            if widget.selection_present():
                selected_text = widget.selection_get()
                root.clipboard_clear()
                root.clipboard_append(selected_text)
                widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
        return "break"
    except Exception:
        return "break"


def handle_keypress(event):
    if event.state & 0x4:  # Ctrl
        if event.keycode == 67:  # C
            copy_text()
            return "break"
        elif event.keycode == 86:  # V
            paste_text()
            return "break"
        elif event.keycode == 88:  # X
            cut_text()
            return "break"
        elif event.keycode == 65:  # A
            widget = root.focus_get()
            if isinstance(widget, (tk.Entry, ttk.Combobox)):
                widget.select_range(0, tk.END)
            return "break"


root.bind_all("<KeyPress>", lambda event: handle_keypress(event))


def create_context_menu(event):
    context_menu = tk.Menu(root, tearoff=0)
    context_menu.add_command(label="Копировать", command=copy_text)
    context_menu.add_command(label="Вставить", command=paste_text)
    context_menu.add_command(label="Вырезать", command=cut_text)
    context_menu.add_separator()
    context_menu.add_command(
        label="Выделить все",
        command=lambda: root.focus_get().select_range(0, tk.END)
        if isinstance(root.focus_get(), (tk.Entry, ttk.Combobox)) else None
    )
    context_menu.tk_popup(event.x_root, event.y_root)


for widget in list(entries_ooo.values()) + list(entries_ip.values()):
    if isinstance(widget, (tk.Entry, ttk.Combobox)):
        widget.bind("<Button-3>", create_context_menu)

root.mainloop()
