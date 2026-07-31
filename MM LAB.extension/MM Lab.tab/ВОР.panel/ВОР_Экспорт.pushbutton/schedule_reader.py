# -*- coding: utf-8 -*-
"""
Чтение данных из спецификаций Revit (ViewSchedule).
"""

import os
import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import (
    ViewSchedule,
    SectionType,
    FilteredElementCollector,
)


def get_all_schedules(doc):
    """Вернуть список (имя, ViewSchedule) всех спецификаций документа,
    отсортированный по имени, без шаблонов."""
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
    result = []
    for s in collector:
        if s.Name and not s.IsTemplate:
            result.append((s.Name, s))
    result.sort(key=lambda x: x[0])
    return result


def find_schedule_by_name(doc, name):
    """Найти спецификацию по имени. Возвращает ViewSchedule или None."""
    for sname, sched in get_all_schedules(doc):
        if sname == name:
            return sched
    return None


def read_schedule(schedule):
    """
    Извлечь данные из ViewSchedule через экспорт во временный файл.
    Надёжнее чем GetCellText — получает все столбцы включая длинный текст.

    Возвращает dict:
    {
        "name": str,
        "headers": [str, ...],
        "rows": [[str, ...], ...],
        "column_count": int
    }
    """
    import tempfile
    import shutil
    from Autodesk.Revit.DB import ViewScheduleExportOptions

    temp_dir = tempfile.mkdtemp()
    try:
        options = ViewScheduleExportOptions()
        options.ColumnHeaders = ExportColumnHeaders.OneRow

        schedule.Export(temp_dir, options)
        txt_name = schedule.Name + ".txt"
        txt_path = os.path.join(temp_dir, txt_name)

        if not os.path.exists(txt_path):
            # Фолбэк на GetCellText если файл не создался
            return _read_via_getcelltext(schedule)

        with open(txt_path, "r") as f:
            lines = f.readlines()

        if not lines:
            return {"name": schedule.Name, "headers": [], "rows": [], "column_count": 0}

        all_rows = []
        for line in lines:
            cells = line.rstrip("\r\n").split("\t")
            all_rows.append(cells)

        headers = all_rows[0] if all_rows else []
        rows = all_rows[1:] if len(all_rows) > 1 else []

        return {
            "name": schedule.Name,
            "headers": headers,
            "rows": rows,
            "column_count": len(headers),
        }
    except Exception:
        return _read_via_getcelltext(schedule)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _read_via_getcelltext(schedule):
    """Фолбэк: чтение через GetCellText (может терять длинный текст)."""
    table = schedule.GetTableData()

    headers = _read_section(table, SectionType.Header)
    body = _read_section(table, SectionType.Body)

    if (not headers or all(c.strip() == "" for r in headers for c in r)) and body:
        headers = body.pop(0) if body else []

    return {
        "name": schedule.Name,
        "headers": headers,
        "rows": body,
        "column_count": len(headers) if headers else (len(body[0]) if body else 0),
    }


def _read_section(table, section_type):
    """Прочитать все ячейки секции таблицы."""
    section = table.GetSectionData(section_type)
    if not section or section.NumberOfRows == 0:
        return []

    rows = []
    for r in range(section.NumberOfRows):
        row = []
        for c in range(section.NumberOfColumns):
            row.append(section.GetCellText(r, c))
        rows.append(row)
    return rows


def check_column_counts(doc, schedule_names):
    """
    Проверить количество столбцов у списка спецификаций.
    Возвращает [(name, column_count), ...].
    """
    results = []
    for name in schedule_names:
        sched = find_schedule_by_name(doc, name)
        if sched:
            data = read_schedule(sched)
            results.append((name, data["column_count"]))
    return results
