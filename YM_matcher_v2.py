"""Сопоставление поддержки Яндекс Музыки. Запуск: python YM_matcher_v2.py."""

from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from pathlib import Path
import re
import sys
from tempfile import NamedTemporaryFile
import traceback
import zipfile

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


INPUT_NAME = "zvonko_in.xlsx"
SOURCE_NAME = "ZVONKO Digital X Яндекс Музыка.xlsx"
OUTPUT_NAME = "zvonko_out_YM.xlsx"
MISSING_NAME = "Нет в zvonko - YM.xlsx"
ERROR_NAME = "Ошибки - YM.xlsx"


def application_dir():
    """В macOS входные файлы лежат рядом с .app, а не внутри Contents."""
    if not getattr(sys, "frozen", False):
        return Path(__file__).resolve().parent
    executable = Path(sys.executable).resolve()
    if sys.platform == "darwin":
        for parent in executable.parents:
            if parent.suffix.lower() == ".app":
                return parent.parent
    return executable.parent


def normalize(value):
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def release_key(value):
    text = str(value or "")
    # Несколько проходов позволяют убрать также вложенные круглые скобки.
    while re.search(r"\([^()]*\)", text):
        text = re.sub(r"\([^()]*\)", " ", text)
    return normalize(text)


def artist_tokens(value):
    parts = re.split(
        r",|;|\n|&|\b(?:feat|ft)\b\.?|\s+x\s+",
        str(value or ""), flags=re.IGNORECASE,
    )
    return frozenset(normalize(part) for part in parts if normalize(part))


def find_header(ws, source=False):
    artist_names = {"артист", "имя артиста", "artist"}
    release_names = {"название"} if source else {"релиз", "название релиза", "release"}
    for row in ws.iter_rows(max_row=min(ws.max_row, 50)):
        artists = [c.column for c in row if normalize(c.value) in artist_names]
        releases = [c.column for c in row if normalize(c.value) in release_names]
        if artists and releases:
            if len(artists) != 1 or len(releases) != 1:
                raise ValueError(f"Лист «{ws.title}»: неоднозначные заголовки в строке {row[0].row}.")
            return row[0].row, artists[0], releases[0]
    raise ValueError(f"Лист «{ws.title}»: не найдены заголовки артиста и релиза в первых 50 строках.")


def split_values(value):
    result = []
    for line in str(value or "").splitlines():
        line = re.sub(r"^[xх]\s+", "", line.strip(), flags=re.IGNORECASE).strip()
        if line and normalize(line) != "не поддержали":
            result.append(line)
    return result


def unique_values(values):
    result, seen = [], set()
    for value in values:
        key = normalize(value)
        if key and key not in seen:
            seen.add(key)
            result.append(value)
    return result


def support_values(ws, row, header_row):
    values = []
    for col in range(6, 13):
        value = ws.cell(row, col).value
        if col <= 8:
            if value is not None and str(value).strip():
                label = ws.cell(header_row, col).value
                if not normalize(label):
                    raise ValueError(f"Лист «{ws.title}»: нет заголовка поддержки в столбце {col}.")
                values.append(str(label).strip())
        else:
            values.extend(split_values(value))
    return unique_values(values)


def temporary_path(directory):
    with NamedTemporaryFile(dir=directory, prefix=".ym-", suffix=".xlsx", delete=False) as tmp:
        return Path(tmp.name)


def replace_formulas(path):
    """Удаляет формулы из XLSX, сохраняя их кеш и остальные части ZIP без изменений."""
    formulas = openpyxl.load_workbook(path, data_only=False)
    cached = openpyxl.load_workbook(path, data_only=True)
    count = 0
    try:
        for ws in formulas:
            for row in ws:
                for cell in row:
                    if cell.data_type == "f":
                        count += 1
                        if cached[ws.title][cell.coordinate].value is None:
                            raise ValueError(
                                f"Не сохранён результат формулы: {ws.title}!{cell.coordinate}. "
                                "Откройте источник в Excel, пересчитайте и сохраните его. "
                                "Исходный файл не изменён."
                            )
    finally:
        formulas.close()
        cached.close()
    if not count:
        return 0
    tmp = temporary_path(path.parent)
    try:
        removed = 0
        with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w") as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", info.filename):
                    # Обычные и общие формулы; значение <v> остаётся нетронутым.
                    data, n = re.subn(rb"<f\b[^>]*?/>|<f\b[^>]*>[^<]*</f>", b"", data)
                    removed += n
                dst.writestr(info, data)
        if removed != count:
            raise ValueError("Не удалось безопасно заменить все формулы. Источник не изменён.")
        check = openpyxl.load_workbook(tmp, data_only=False)
        try:
            if any(c.data_type == "f" for ws in check for row in ws for c in row):
                raise ValueError("После преобразования остались формулы. Источник не изменён.")
        finally:
            check.close()
        tmp.replace(path)
    finally:
        tmp.unlink(missing_ok=True)
    return count


@dataclass
class Record:
    row: int
    artists: frozenset
    release: str


def read_records(ws, header, file_name, issues):
    header_row, artist_col, release_col = header
    records = []
    for row in ws.iter_rows(min_row=header_row + 1):
        if not any(c.value is not None and str(c.value).strip() for c in row):
            continue
        number = row[0].row
        artist = ws.cell(number, artist_col).value
        release = ws.cell(number, release_col).value
        tokens, key = artist_tokens(artist), release_key(release)
        if not tokens or not key:
            reason = "Не заполнен артист" if not tokens else "Не заполнен релиз (после удаления скобок)"
            issues.append([file_name, ws.title, number, artist, release, reason])
            continue
        records.append(Record(number, tokens, key))
    return records


def make_index(records):
    index = defaultdict(set)
    for record in records:
        for artist in record.artists:
            index[(artist, record.release)].add(record.row)
    return index


def matching_rows(record, index):
    return sorted({row for artist in record.artists for row in index.get((artist, record.release), ())})


def prepare_playlist_column(ws, header):
    header_row, artist_col, release_col = header
    columns = [c.column for c in ws[header_row] if normalize(c.value) == "плейлисты"]
    if len(columns) > 1:
        raise ValueError("Во входном файле несколько столбцов «Плейлисты».")
    if columns and columns[0] != 3:
        # Сначала переносим существующий столбец в C, затем удаляем его прежнюю копию.
        old_col = columns[0]
        values = [(ws.cell(r, old_col).value, copy(ws.cell(r, old_col)._style)) for r in range(1, ws.max_row + 1)]
        ws.delete_cols(old_col)
        artist_col -= artist_col > old_col
        release_col -= release_col > old_col
        columns = []
    else:
        values = []
    if not columns:
        ws.insert_cols(3)
        artist_col += artist_col >= 3
        release_col += release_col >= 3
        for r, (value, style) in enumerate(values, 1):
            ws.cell(r, 3).value = value
            ws.cell(r, 3)._style = style
    ws.cell(header_row, 3).value = "Плейлисты"
    ws.cell(header_row, 3)._style = copy(ws.cell(header_row, release_col)._style)
    ws.column_dimensions["C"].width = 42
    for r in range(header_row + 1, ws.max_row + 1):
        ws.cell(r, 3).value = None
    return artist_col, release_col


def missing_workbook(source_ws, header_row, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = source_ws.title
    for name, dimension in source_ws.column_dimensions.items():
        target = ws.column_dimensions[name]
        target.width, target.hidden = dimension.width, dimension.hidden
        target.min, target.max = dimension.min, dimension.max
        target.outlineLevel = dimension.outlineLevel
    for target_row, source_row in enumerate([*range(1, header_row + 1), *rows], 1):
        ws.row_dimensions[target_row].height = source_ws.row_dimensions[source_row].height
        for source_cell in source_ws[source_row]:
            cell = ws.cell(target_row, source_cell.column, source_cell.value)
            # Между книгами копируем сами стили, а не их внутренние номера.
            cell.font = copy(source_cell.font)
            cell.fill = copy(source_cell.fill)
            cell.border = copy(source_cell.border)
            cell.alignment = copy(source_cell.alignment)
            cell.number_format = source_cell.number_format
            cell.protection = copy(source_cell.protection)
            cell.hyperlink = copy(source_cell.hyperlink)
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ws.max_column)}{ws.max_row}"
    return wb


def error_workbook(issues):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ошибки"
    ws.append(["Файл", "Лист", "Строка Excel", "Артист", "Релиз", "Причина"])
    for issue in issues:
        ws.append(issue)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for col, width in zip("ABCDEF", (48, 35, 15, 35, 45, 75)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return wb


def save_results(directory, books):
    """Сначала готовим все книги; только затем заменяем готовые результаты."""
    pending = []
    try:
        for name, book in books:
            tmp = temporary_path(directory)
            pending.append((tmp, directory / name))
            with tmp.open("wb") as stream:
                book.save(stream)
        for tmp, target in pending:
            tmp.replace(target)
    finally:
        for tmp, _ in pending:
            tmp.unlink(missing_ok=True)


def run(directory):
    directory = Path(directory)
    for name in (INPUT_NAME, SOURCE_NAME):
        if not (directory / name).is_file():
            raise FileNotFoundError(f"Не найден {name}. Положите оба Excel-файла рядом с программой: {directory}")
    issues = []
    books = []
    try:
        input_wb = openpyxl.load_workbook(directory / INPUT_NAME)
        books.append(input_wb)
        if len(input_wb.worksheets) != 1:
            raise ValueError("В zvonko_in.xlsx должен быть ровно один лист.")
        input_ws = input_wb.worksheets[0]
        input_header = find_header(input_ws)
        input_records = read_records(input_ws, input_header, INPUT_NAME, issues)
        formulas = replace_formulas(directory / SOURCE_NAME)
        source_wb = openpyxl.load_workbook(directory / SOURCE_NAME, data_only=True)
        books.append(source_wb)
        source_ws = source_wb.worksheets[0]
        source_header = find_header(source_ws, source=True)
        source_records = read_records(source_ws, source_header, SOURCE_NAME, issues)
        support = {r.row: support_values(source_ws, r.row, source_header[0]) for r in source_records}
        source_index = make_index(source_records)
        input_index = make_index(input_records)
        artist_col, release_col = prepare_playlist_column(input_ws, input_header)
        matched = 0
        for record in input_records:
            rows = matching_rows(record, source_index)
            values = unique_values(value for row in rows for value in support[row])
            cell = input_ws.cell(record.row, 3)
            cell.value = "\n".join(values) or None
            alignment = copy(input_ws.cell(record.row, release_col).alignment)
            alignment.vertical, alignment.wrap_text = "top", True
            cell.alignment = alignment
            if rows:
                matched += 1
            else:
                issues.append([INPUT_NAME, input_ws.title, record.row,
                               input_ws.cell(record.row, artist_col).value,
                               input_ws.cell(record.row, release_col).value,
                               "Нет совпадения по артисту и релизу в источнике"])
        missing_rows = [r.row for r in source_records if not matching_rows(r, input_index)]
        missing = missing_workbook(source_ws, source_header[0], missing_rows)
        errors = error_workbook(issues)
        books.extend([missing, errors])
        save_results(directory, [(OUTPUT_NAME, input_wb), (MISSING_NAME, missing), (ERROR_NAME, errors)])
        return (f"Обработка завершена. Совпадений: {matched} из {len(input_records)} корректных записей.\n"
                f"Нет во входном файле: {len(missing_rows)}. Записей в отчёте об ошибках: {len(issues)}.\n"
                f"Формул заменено: {formulas}.\nРезультаты: {directory}")
    finally:
        for book in books:
            book.close()


def main():
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    directory = application_dir()
    try:
        message = run(directory)
    except Exception as error:
        detail = traceback.format_exc()
        try:
            report = error_workbook([["Программа", "", "", "", "", str(error)]])
            try:
                save_results(directory, [(ERROR_NAME, report)])
            finally:
                report.close()
            (directory / "Ошибка запуска - YM.txt").write_text(detail, encoding="utf-8")
        except OSError:
            pass
        messagebox.showerror("Ошибка YM Matcher", f"{error}\n\nОбработка не завершена. Проверьте отчёт об ошибках.", parent=root)
        return 1
    else:
        messagebox.showinfo("YM Matcher", message, parent=root)
        return 0
    finally:
        root.destroy()


if __name__ == "__main__":
    raise SystemExit(main())
