#! python3
# -*- coding: utf-8 -*-
"""Публикация Шаблона

Публикует выбранный открытый проект сразу в двух форматах: как проект .rvt
в общую папку RVT и как шаблон .rte в папку публикаций RTE. Имя чистится от
суффикса пользователя, для .rte дополнительно ставится маркер RD_; временная
копия, папки бэкапов и нумерованные файлы убираются за собой.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Публикация\nШаблона"
__author__ = "MM LAB"

import os
import re
import shutil
import sys
import tempfile

_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton -> panel -> tab -> MM_LAB.extension
_EXTENSION_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")
if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Drawing import Font
from System.Drawing import FontStyle
from System.Drawing import Point
from System.Drawing import Size
from System.Windows.Forms import Button
from System.Windows.Forms import ComboBox
from System.Windows.Forms import ComboBoxStyle
from System.Windows.Forms import DialogResult
from System.Windows.Forms import Form
from System.Windows.Forms import FormBorderStyle
from System.Windows.Forms import FormStartPosition
from System.Windows.Forms import Label

from Autodesk.Revit.DB import DetachFromCentralOption
from Autodesk.Revit.DB import ModelPathUtils
from Autodesk.Revit.DB import OpenOptions
from Autodesk.Revit.DB import SaveAsOptions
from Autodesk.Revit.DB import WorksharingSaveAsOptions
from Autodesk.Revit.UI import TaskDialog

import revit_compat


COMMAND_NAME = u"Публикация Шаблона"

# Корень папки публикаций RTE; год-подпапка берётся из версии Revit.
OUTPUT_RTE_ROOT = r"G:\Общие диски\02_BIM\2_Templates\1_Revit\1_Project\2_RTE"
# Папка публикаций RVT — общая, без разбивки по годам.
OUTPUT_RVT_FOLDER = r"G:\Общие диски\02_BIM\2_Templates\1_Revit\1_Project\1_RVT"


class DocumentSelectionForm(Form):
    """Окно выбора открытого проекта Revit для публикации."""

    def __init__(self, doc_list):
        self.doc_map = {d.Title: d for d in doc_list}
        self.selected_doc = None
        self.InitializeComponent()

    def InitializeComponent(self):
        self.Text = u"Публикация .RVT и .RTE"
        self.Size = Size(450, 180)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        label = Label()
        label.Text = u"Выберите открытый проект для экспорта:"
        label.Location = Point(15, 15)
        label.Size = Size(400, 20)
        label.Font = Font("Segoe UI", 9.0, FontStyle.Regular)
        self.Controls.Add(label)

        self.combo = ComboBox()
        self.combo.Location = Point(15, 40)
        self.combo.Size = Size(400, 25)
        self.combo.DropDownStyle = ComboBoxStyle.DropDownList
        for title in sorted(self.doc_map.keys()):
            self.combo.Items.Add(title)
        if self.combo.Items.Count > 0:
            self.combo.SelectedIndex = 0
        self.Controls.Add(self.combo)

        btn_ok = Button()
        btn_ok.Text = u"Запустить"
        btn_ok.Location = Point(220, 85)
        btn_ok.Size = Size(90, 30)
        btn_ok.DialogResult = DialogResult.OK
        btn_ok.Click += self.on_ok
        self.Controls.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Text = u"Отмена"
        btn_cancel.Location = Point(325, 85)
        btn_cancel.Size = Size(90, 30)
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)

        self.AcceptButton = btn_ok
        self.CancelButton = btn_cancel

    def on_ok(self, sender, args):
        selected_title = self.combo.SelectedItem
        if selected_title in self.doc_map:
            self.selected_doc = self.doc_map[selected_title]


def rte_folder_for(revit_version):
    """Папка публикации .rte для версии Revit: <OUTPUT_RTE_ROOT>\\<год>."""
    return os.path.join(OUTPUT_RTE_ROOT, str(revit_version))


def remove_user_suffix(filename_without_ext):
    """Убирает хвост с именем пользователя (например _e.ahmedyanova)."""
    return re.sub(r"_[a-zA-Z0-9\.\-]+$", "", filename_without_ext)


def clean_and_format_filename(filename_without_ext, revit_version):
    """Убирает суффикс пользователя и ставит RD_ перед маркером версии.

    Маркер собирается из версии Revit: 2024 -> R24, 2020 -> R20.
    """
    cleaned = remove_user_suffix(filename_without_ext)

    marker = "R{0}".format(revit_version % 100)
    if marker in cleaned:
        escaped = re.escape(marker)
        if (not re.search(r"_RD_" + escaped + r"\b", cleaned)
                and not re.search(r"\bRD_" + escaped + r"\b", cleaned)):
            cleaned = cleaned.replace(marker, "RD_" + marker)

    return cleaned


def clean_up_temp_and_backups(folder_path):
    """Удаляет временные файлы сессии, папки бэкапов и нумерованные .0001.rte/.rvt."""
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


def get_eligible_documents(app):
    """Собирает все открытые и сохранённые на диске .rvt проекты.

    Загруженные RVT-связи отсеиваются: в app.Documents они лежат наравне
    с проектами и иначе попали бы в список выбора.
    """
    valid_docs = []
    for d in app.Documents:
        if d.IsFamilyDocument or d.IsLinked:
            continue
        path = d.PathName
        if path and path.lower().endswith(".rvt"):
            valid_docs.append(d)
    return valid_docs


def _save_copy_as(app, temp_model_path, target_path, detach_option, as_central):
    """Открывает временную копию с заданным Detach и сохраняет её в target_path."""
    temp_doc = None
    try:
        open_options = OpenOptions()
        open_options.DetachFromCentralOption = detach_option
        temp_doc = app.OpenDocumentFile(temp_model_path, open_options)

        save_options = SaveAsOptions()
        save_options.OverwriteExistingFile = True
        save_options.MaximumBackups = 1

        if as_central and temp_doc.IsWorkshared:
            ws_options = WorksharingSaveAsOptions()
            ws_options.SaveAsCentral = True
            save_options.SetWorksharingOptions(ws_options)

        target_model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(
            target_path)
        temp_doc.SaveAs(target_model_path, save_options)
    finally:
        # Фоновый документ закрываем всегда: открытая копия не даст ни удалить
        # временный файл, ни перезапустить команду.
        if temp_doc is not None:
            temp_doc.Close(False)


def publish_files(doc, app, revit_version):
    """Публикует проект как .rvt (в 1_RVT) и как .rte (в 2_RTE/<год>).

    Возвращает пару (путь к .rvt, путь к .rte).
    """
    source_path = doc.PathName
    rte_folder = rte_folder_for(revit_version)

    for folder in (OUTPUT_RVT_FOLDER, rte_folder):
        if not os.path.isdir(folder):
            os.makedirs(folder)

    raw_filename = os.path.splitext(os.path.basename(source_path))[0]

    # Для .rvt имя чистится только от суффикса пользователя, без маркера RD_.
    target_rvt_path = os.path.join(
        OUTPUT_RVT_FOLDER, remove_user_suffix(raw_filename) + ".rvt")
    # Для .rte к очищенному имени добавляется маркер RD_.
    target_rte_path = os.path.join(
        rte_folder,
        clean_and_format_filename(raw_filename, revit_version) + ".rte")

    temp_copy_path = os.path.join(
        tempfile.gettempdir(), raw_filename + "_publish_temp.rvt")

    try:
        doc.Save()
        shutil.copyfile(source_path, temp_copy_path)
        temp_model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(
            temp_copy_path)

        # .rvt — рабочие наборы сохраняем, файл публикуем как центральный.
        _save_copy_as(
            app, temp_model_path, target_rvt_path,
            DetachFromCentralOption.DetachAndPreserveWorksets
            if doc.IsWorkshared else DetachFromCentralOption.DoNotDetach,
            as_central=True)

        # .rte — рабочие наборы удаляем: шаблон не должен их тянуть за собой.
        _save_copy_as(
            app, temp_model_path, target_rte_path,
            DetachFromCentralOption.DetachAndDiscardWorksets
            if doc.IsWorkshared else DetachFromCentralOption.DoNotDetach,
            as_central=False)
    finally:
        if os.path.exists(temp_copy_path):
            try:
                os.remove(temp_copy_path)
            except Exception:
                pass

        clean_up_temp_and_backups(OUTPUT_RVT_FOLDER)
        clean_up_temp_and_backups(rte_folder)

    return target_rvt_path, target_rte_path


def main():
    """Точка входа кнопки: гейт версии, выбор проекта, публикация."""
    revit_version = revit_compat.require_supported_version(COMMAND_NAME)

    app = __revit__.Application

    eligible_docs = get_eligible_documents(app)
    if not eligible_docs:
        TaskDialog.Show(
            COMMAND_NAME,
            u"Не найдено ни одного сохранённого проекта (.rvt).\n\n"
            u"Убедись, что открытый файл проекта сохранён на диск.")
        return

    form = DocumentSelectionForm(eligible_docs)
    try:
        result = form.ShowDialog()
        selected_doc = form.selected_doc
    finally:
        form.Dispose()

    if result != DialogResult.OK or selected_doc is None:
        return

    rvt_path, rte_path = publish_files(selected_doc, app, revit_version)

    TaskDialog.Show(
        COMMAND_NAME,
        u"Публикация успешно завершена!\n\n"
        u"• Проект (.rvt):\n{0}\n\n"
        u"• Шаблон (.rte):\n{1}".format(rvt_path, rte_path))


try:
    main()
except SystemExit:
    pass  # require_supported_version уже показал свой диалог
except Exception as ex:
    TaskDialog.Show(COMMAND_NAME, u"Ошибка при работе скрипта:\n{0}".format(ex))
