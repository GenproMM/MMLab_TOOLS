#! python3
# -*- coding: utf-8 -*-
"""Приточный по классификации

Показывает найденные в проекте классификации систем воздуховодной сети
и позволяет для каждой классификации выбрать: включить, выключить или не
менять параметр «Приточный» у элементов с этой классификацией.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Приточный\nпо классификации"
__author__ = "GENPRO LAB"

# Канонический lib-бутстрап (D-15).
import os
import sys

_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton -> panel -> tab -> MM_LAB.extension
_EXTENSION_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")
if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import StorageType
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog
from System.Drawing import Point
from System.Drawing import Size
from System.Windows.Forms import AnchorStyles
from System.Windows.Forms import Button
from System.Windows.Forms import DataGridView
from System.Windows.Forms import DataGridViewAutoSizeColumnsMode
from System.Windows.Forms import DataGridViewComboBoxColumn
from System.Windows.Forms import DataGridViewEditMode
from System.Windows.Forms import DataGridViewSelectionMode
from System.Windows.Forms import DataGridViewTextBoxColumn
from System.Windows.Forms import DialogResult
from System.Windows.Forms import Form
from System.Windows.Forms import FormBorderStyle
from System.Windows.Forms import FormStartPosition
from System.Windows.Forms import Panel

import revit_compat
from ios_common_helpers import DECISION_TO_LABEL
from ios_common_helpers import LABEL_TO_DECISION
from ios_common_helpers import SupplyFlagDecision
from ios_common_helpers import collect_elements
from ios_common_helpers import get_default_supply_flag_decision
from ios_common_helpers import get_supply_flag_parameter
from ios_common_helpers import get_system_classification
from ios_common_helpers import is_writable
from ios_common_helpers import normalize_text
from ios_common_helpers import to_text


COMMAND_NAME = u"Приточный по классификации"


class SystemClassificationSelectionForm(Form):
    def __init__(self, options):
        Form.__init__(self)
        self._options = options
        self._grid = None
        self._build_ui()

    def _build_ui(self):
        self.Text = u"Настройка классификаций систем"
        self.ClientSize = Size(820, 560)
        self.MinimumSize = Size(820, 560)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.ShowInTaskbar = False

        grid = DataGridView()
        grid.Location = Point(12, 12)
        grid.Size = Size(796, 468)
        grid.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.MultiSelect = False
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grid.EditMode = DataGridViewEditMode.EditOnEnter
        grid.RowHeadersVisible = False
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        classification_column = DataGridViewTextBoxColumn()
        classification_column.Name = "classification"
        classification_column.HeaderText = u"Классификация системы"
        classification_column.ReadOnly = True

        decision_column = DataGridViewComboBoxColumn()
        decision_column.Name = "decision"
        decision_column.HeaderText = u"Как установить параметр \"Приточный\""
        for label in [u"Приток", u"Не приток", u"Не менять"]:
            decision_column.Items.Add(label)

        grid.Columns.Add(classification_column)
        grid.Columns.Add(decision_column)

        for option in self._options:
            index = grid.Rows.Add()
            grid.Rows[index].Cells[0].Value = option["classification"]
            grid.Rows[index].Cells[1].Value = DECISION_TO_LABEL[option["decision"]]

        button_panel = Panel()
        button_panel.Location = Point(0, 490)
        button_panel.Size = Size(820, 70)
        button_panel.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom

        ok_button = Button()
        ok_button.Text = u"Применить"
        ok_button.DialogResult = DialogResult.OK
        ok_button.Size = Size(140, 34)
        ok_button.Location = Point(536, 18)
        ok_button.Anchor = AnchorStyles.Right | AnchorStyles.Top

        cancel_button = Button()
        cancel_button.Text = u"Отмена"
        cancel_button.DialogResult = DialogResult.Cancel
        cancel_button.Size = Size(120, 34)
        cancel_button.Location = Point(688, 18)
        cancel_button.Anchor = AnchorStyles.Right | AnchorStyles.Top

        button_panel.Controls.Add(ok_button)
        button_panel.Controls.Add(cancel_button)

        self.AcceptButton = ok_button
        self.CancelButton = cancel_button

        self.Controls.Add(button_panel)
        self.Controls.Add(grid)

        self._grid = grid

    def get_selections(self):
        try:
            self._grid.EndEdit()
        except Exception:
            pass

        selections = []
        for row in self._grid.Rows:
            classification = to_text(row.Cells[0].Value)
            label = to_text(row.Cells[1].Value) or DECISION_TO_LABEL[SupplyFlagDecision.IGNORE]
            decision = LABEL_TO_DECISION.get(label, SupplyFlagDecision.IGNORE)
            selections.append({
                "classification": classification,
                "decision": decision,
            })
        return selections


def build_classification_options(targets):
    seen = {}

    for element in targets:
        classification = get_system_classification(element)
        if not classification:
            continue

        key = normalize_text(classification)
        if key and key not in seen:
            seen[key] = classification

    classifications = sorted(seen.values(), key=lambda item: item.lower())
    options = []
    for classification in classifications:
        options.append({
            "classification": classification,
            "decision": get_default_supply_flag_decision(classification),
        })

    return options


def main(doc):
    """Точка входа: настраивает параметр «Приточный» по выбранным классификациям систем."""
    revit_compat.require_supported_version(COMMAND_NAME)

    targets = collect_elements(
        doc,
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_DuctTerminal,
    )

    options = build_classification_options(targets)
    if not options:
        TaskDialog.Show(COMMAND_NAME, u"В проекте не найдено классификаций систем для обработки.")
        return

    form = SystemClassificationSelectionForm(options)
    try:
        if form.ShowDialog() != DialogResult.OK:
            return
        options = form.get_selections()
    finally:
        form.Dispose()

    selection_map = {}
    for option in options:
        selection_map[normalize_text(option["classification"])] = option["decision"]

    updated_count = 0
    unchanged_count = 0
    skipped_count = 0

    transaction = Transaction(doc, COMMAND_NAME)
    transaction.Start()
    try:
        for element in targets:
            classification = get_system_classification(element)
            normalized = normalize_text(classification)
            decision = selection_map.get(normalized)
            if not normalized or decision is None or decision == SupplyFlagDecision.IGNORE:
                skipped_count += 1
                continue

            parameter = get_supply_flag_parameter(element)
            if not is_writable(parameter) or parameter.StorageType != StorageType.Integer:
                skipped_count += 1
                continue

            target_value = 1 if decision == SupplyFlagDecision.SUPPLY else 0
            if parameter.AsInteger() == target_value:
                unchanged_count += 1
                continue

            parameter.Set(target_value)
            updated_count += 1

        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    TaskDialog.Show(
        COMMAND_NAME,
        u"Изменено: {0}\nБез изменений: {1}\nПропущено: {2}\nКлассификаций в настройке: {3}".format(
            updated_count,
            unchanged_count,
            skipped_count,
            len(options),
        ),
    )


def _entry():
    """Готовит doc/uidoc и вызывает main (doc/uidoc — параметрами, правило 18)."""
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None or uidoc.Document is None:
        TaskDialog.Show(COMMAND_NAME, u"Открой проект Revit и повтори команду.")
        return
    main(uidoc.Document)


try:
    _entry()
except SystemExit:
    pass  # require_supported_version уже показал свой диалог
except Exception as ex:
    TaskDialog.Show(COMMAND_NAME, u"Ошибка:\n{0}".format(ex))
