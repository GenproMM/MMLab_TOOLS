#! python3
# -*- coding: utf-8 -*-
"""Публикация Шаблона

Экспортирует открытый проект .rvt в шаблон .rte в общую папку публикаций:
чистит имя файла от суффикса пользователя, ставит маркер RD_ и убирает
за собой временные файлы и папки бэкапов.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Публикация\nШаблона"
__author__ = "MM LAB"

import os
import re
import shutil
import sys

_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton -> panel -> tab -> MM_LAB.extension
_EXTENSION_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")
if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import DetachFromCentralOption
from Autodesk.Revit.DB import ModelPathUtils
from Autodesk.Revit.DB import OpenOptions
from Autodesk.Revit.DB import SaveAsOptions
from Autodesk.Revit.UI import TaskDialog

import revit_compat


COMMAND_NAME = u"Публикация Шаблона"

# Корень общей папки публикаций RTE; год-подпапка берётся из версии Revit.
OUTPUT_ROOT = r"G:\Общие диски\02_BIM\2_Templates\1_Revit\1_Project\2_RTE"


def output_folder_for(revit_version):
    """Папка публикации для версии Revit: <OUTPUT_ROOT>\\<год>."""
    return os.path.join(OUTPUT_ROOT, str(revit_version))


def clean_and_format_filename(filename_without_ext, revit_version):
    """Убирает хвост с именем пользователя и ставит RD_ перед маркером версии.

    Маркер собирается из версии Revit: 2024 -> R24, 2020 -> R20.
    """
    cleaned = re.sub(r"_[a-zA-Z0-9\.\-]+$", "", filename_without_ext)

    marker = "R{0}".format(revit_version % 100)
    if marker in cleaned:
        escaped = re.escape(marker)
        if (not re.search(r"_RD_" + escaped + r"\b", cleaned)
                and not re.search(r"\bRD_" + escaped + r"\b", cleaned)):
            cleaned = cleaned.replace(marker, "RD_" + marker)

    return cleaned


def clean_up_temp_and_backups(folder_path):
    """Удаляет временные файлы сессии, папки бэкапов и нумерованные .0001.rte."""
    if not os.path.isdir(folder_path):
        return

    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)

        if os.path.isdir(item_path) and item.endswith("_backup"):
            try:
                shutil.rmtree(item_path)
            except Exception:
                pass
        elif os.path.isfile(item_path):
            if "_temp" in item or re.search(r"\.\d{4}\.(rte|rvt)$", item):
                try:
                    os.remove(item_path)
                except Exception:
                    pass


def publish_template(doc, app, revit_version):
    """Публикует текущий проект как .rte; возвращает имя созданного файла."""
    source_path = doc.PathName
    output_folder = output_folder_for(revit_version)

    if not os.path.isdir(output_folder):
        os.makedirs(output_folder)

    target_filename = clean_and_format_filename(
        os.path.splitext(os.path.basename(source_path))[0], revit_version)
    target_rte_path = os.path.join(output_folder, target_filename + ".rte")
    temp_copy_path = os.path.join(output_folder, target_filename + "_temp.rvt")

    temp_doc = None
    try:
        doc.Save()
        shutil.copyfile(source_path, temp_copy_path)

        open_options = OpenOptions()
        if doc.IsWorkshared:
            open_options.DetachFromCentralOption = (
                DetachFromCentralOption.DetachAndDiscardWorksets)
        else:
            open_options.DetachFromCentralOption = (
                DetachFromCentralOption.DoNotDetach)

        temp_model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(
            temp_copy_path)
        temp_doc = app.OpenDocumentFile(temp_model_path, open_options)

        save_options = SaveAsOptions()
        save_options.OverwriteExistingFile = True
        save_options.MaximumBackups = 1

        target_model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(
            target_rte_path)
        temp_doc.SaveAs(target_model_path, save_options)
    finally:
        # Фоновый документ закрываем всегда: открытый _temp.rvt не даст
        # ни удалить временный файл, ни перезапустить команду.
        if temp_doc is not None:
            temp_doc.Close(False)
        clean_up_temp_and_backups(output_folder)

    return target_filename + ".rte"


def main():
    """Точка входа кнопки: гейт версии, проверки документа, публикация."""
    revit_version = revit_compat.require_supported_version(COMMAND_NAME)

    uidoc = __revit__.ActiveUIDocument
    if uidoc is None or uidoc.Document is None:
        TaskDialog.Show(COMMAND_NAME, u"Открой проект Revit и повтори команду.")
        return
    doc = uidoc.Document

    if doc.IsFamilyDocument:
        TaskDialog.Show(
            COMMAND_NAME,
            u"Команда работает только с файлами проектов (.rvt).")
        return

    source_path = doc.PathName
    if not source_path or not source_path.lower().endswith(".rvt"):
        TaskDialog.Show(
            COMMAND_NAME,
            u"Сохрани проект на диск как .rvt и повтори команду.")
        return

    published = publish_template(doc, __revit__.Application, revit_version)

    TaskDialog.Show(
        COMMAND_NAME,
        u"Шаблон .rte успешно опубликован:\n{0}".format(published))


try:
    main()
except SystemExit:
    pass  # require_supported_version уже показал свой диалог
except Exception as ex:
    TaskDialog.Show(COMMAND_NAME, u"Ошибка:\n{0}".format(ex))
