# -*- coding: utf-8 -*-
"""
Сборка сводного Excel-файла из нескольких источников.
Использует Excel COM Interop (IronPython).
"""

import os
import System
from System.Globalization import CultureInfo
from System.Runtime.InteropServices import Marshal


def is_file_locked(file_path):
    """Проверить, заблокирован ли файл (открыт в другой программе)."""
    if not os.path.exists(file_path):
        return False
    try:
        with open(file_path, "a"):
            pass
        return False
    except IOError:
        return True


def _get_excel_app():
    """Создать экземпляр Excel.Application через COM."""
    excel_type = System.Type.GetTypeFromProgID("Excel.Application")
    xl = System.Activator.CreateInstance(excel_type)
    xl.Visible = False
    xl.DisplayAlerts = False
    return xl


def _release(xl):
    """Корректно освободить COM-объект Excel."""
    try:
        xl.Quit()
    except Exception:
        pass
    Marshal.ReleaseComObject(xl)


def _rgb_to_bgr(r, g, b):
    """COM Excel использует BGR-формат для цветов."""
    return (b << 16) | (g << 8) | r


# Цвета в BGR
_HEADER_COLOR = _rgb_to_bgr(0x44, 0x72, 0xC4)   # #4472C4
_WHITE = _rgb_to_bgr(0xFF, 0xFF, 0xFF)
_LABEL_COLOR = _rgb_to_bgr(0x33, 0x33, 0x33)     # #333333


def _cell(ws, row, col):
    """Получить ячейку через Cells.Item[row, col]."""
    return ws.Cells.Item[row, col]


def _range(ws, r1, c1, r2, c2):
    """Получить диапазон ячеек через Range[cell1, cell2]."""
    return ws.Range[_cell(ws, r1, c1), _cell(ws, r2, c2)]


def read_excel_file(file_path):
    """
    Прочитать внешний Excel-файл через COM Interop.

    Возвращает dict:
    {
        "name": str,
        "headers": [str, ...],
        "rows": [[str, ...], ...],
        "column_count": int
    }
    """
    original_culture = System.Threading.Thread.CurrentThread.CurrentCulture
    System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo("en-US")
    xl = _get_excel_app()
    try:
        wb = xl.Workbooks.Open(os.path.abspath(file_path))
        ws = wb.Worksheets.Item[1]

        used = ws.UsedRange
        rows_count = used.Rows.Count
        cols_count = used.Columns.Count

        if rows_count == 0:
            wb.Close(False)
            return {
                "name": os.path.basename(file_path),
                "headers": [],
                "rows": [],
                "column_count": 0,
            }

        values = used.Value2
        if values is None:
            wb.Close(False)
            return {
                "name": os.path.basename(file_path),
                "headers": [],
                "rows": [],
                "column_count": 0,
            }

        all_rows = []
        for r in range(1, rows_count + 1):
            row = []
            for c in range(1, cols_count + 1):
                try:
                    if rows_count == 1 and cols_count == 1:
                        val = values
                    elif rows_count == 1:
                        val = values.Item[c]
                    elif cols_count == 1:
                        val = values.Item[r]
                    else:
                        val = values.Item[r, c]
                    row.append(str(val) if val is not None else "")
                except Exception:
                    row.append("")
            all_rows.append(row)

        wb.Close(False)

        headers = all_rows[0] if all_rows else []
        rows = all_rows[1:] if len(all_rows) > 1 else []

        return {
            "name": os.path.basename(file_path),
            "headers": headers,
            "rows": rows,
            "column_count": len(headers),
        }
    finally:
        _release(xl)
        System.Threading.Thread.CurrentThread.CurrentCulture = original_culture


def build_export(sources_data, output_path):
    """
    Собрать все источники в один Excel-файл через COM Interop.

    sources_data: list of dict {"name", "headers", "rows", "column_count"}
    output_path: путь к итоговому файлу

    Возвращает общее количество строк данных.
    """
    original_culture = System.Threading.Thread.CurrentThread.CurrentCulture
    System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo("en-US")
    xl = _get_excel_app()
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets.Item[1]
        ws.Name = u"\u042d\u043a\u0441\u043f\u043e\u0440\u0442 \u0412\u041e\u0420"

        current_row = 1

        for i, src in enumerate(sources_data):
            col_count = max(src["column_count"], 1)

            # Название источника
            cell = _cell(ws, current_row, 1)
            cell.Value2 = src["name"]
            cell.Font.Bold = True
            cell.Font.Size = 11
            cell.Font.Color = _LABEL_COLOR
            if col_count > 1:
                rng = _range(ws, current_row, 1, current_row, col_count)
                rng.Merge()
                rng = None
            current_row += 1

            # Шапка
            for col_idx, header in enumerate(src["headers"], 1):
                cell = _cell(ws, current_row, col_idx)
                cell.Value2 = str(header) if header is not None else ""
                cell.Font.Bold = True
                cell.Font.Size = 10
                cell.Font.Color = _WHITE
                cell.Interior.Color = _HEADER_COLOR
                cell.HorizontalAlignment = -4108  # xlCenter
                cell.VerticalAlignment = -4108     # xlCenter
                cell.WrapText = True
                cell.Borders.LineStyle = 1         # xlContinuous
            current_row += 1

            # Данные
            for row_data in src["rows"]:
                for col_idx, val in enumerate(row_data, 1):
                    cell = _cell(ws, current_row, col_idx)
                    cell.Value2 = str(val) if val is not None else ""
                    cell.Font.Size = 10
                    cell.VerticalAlignment = -4108  # xlCenter
                    cell.WrapText = True
                    cell.Borders.LineStyle = 1
                current_row += 1

        # Автоподбор ширины столбцов
        if sources_data:
            max_col = max(s["column_count"] for s in sources_data)
            if max_col > 0:
                rng = _range(ws, 1, 1, current_row, max_col)
                rng.Columns.AutoFit()
                rng = None

        # Сохранить
        abs_path = os.path.abspath(output_path)
        if os.path.exists(abs_path):
            try:
                os.remove(abs_path)
            except Exception:
                pass
        missing = System.Type.Missing
        wb.SaveAs(
            abs_path,
            51,         # FileFormat: xlOpenXMLWorkbook
            missing,    # Password
            missing,    # WriteResPassword
            missing,    # ReadOnlyRecommended
            missing,    # CreateBackup
            1,          # AccessMode: xlNoChange
            missing,    # ConflictResolution
            missing,    # AddToMru
            missing,    # TextCodepage
            missing,    # TextVisualLayout
            missing,    # Local
        )
        wb.Close()

        return current_row - 1
    finally:
        _release(xl)
        System.Threading.Thread.CurrentThread.CurrentCulture = original_culture
