# -*- coding: utf-8 -*-
"""
Универсальное окно настроек скрипта.
Строит UI динамически по SETTINGS_SCHEMA из скрипта.
Для сложных окон скрипты определяют show_settings(doc, current_settings).
"""

import os
import sys
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import Window, Thickness, Visibility
from System.Windows.Controls import *
from System.Windows.Media import SolidColorBrush, Colors

from pyrevit import DB


class ScriptSettingsWindow(Window):
    """Окно настроек скрипта (generic, по SETTINGS_SCHEMA)."""

    def __init__(self, script_name, schema, current_settings, doc):
        self.Title = u"Настройки: {}".format(script_name)
        self.Width = 450
        self.Height = 550
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.Background = SolidColorBrush(Colors.White)
        self.result_settings = None
        self.doc = doc
        self.controls = {}  # key -> widget
        self._schema_defs = {}  # key -> schema item dict
        self._sheet_extras = {}  # key -> {sort_combo, hide_cb}

        stack = StackPanel()
        stack.Margin = Thickness(16)

        # Заголовок
        header = TextBlock()
        header.Text = u"Настройки скрипта"
        header.FontSize = 16
        header.FontWeight = System.Windows.FontWeights.Bold
        header.Margin = Thickness(0, 0, 0, 16)
        stack.Children.Add(header)

        # Scrollable settings area
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.MaxHeight = 420
        settings_panel = StackPanel()
        settings_panel.Margin = Thickness(0, 0, 0, 8)

        for item_def in schema:
            key = item_def.get("key", "")
            label = item_def.get("label", key)
            stype = item_def.get("type", "text")
            current = current_settings.get(key)

            self._schema_defs[key] = item_def

            lbl = TextBlock()
            lbl.Text = label
            lbl.FontSize = 13
            lbl.FontWeight = System.Windows.FontWeights.SemiBold
            lbl.Margin = Thickness(0, 8, 0, 4)
            settings_panel.Children.Add(lbl)

            if stype == "sheet_list":
                sortable = item_def.get("sortable", False)
                hide_opt = item_def.get("hide_unselected", False)
                sort_by = current_settings.get(key + "__sort_by", None)
                hide_val = current_settings.get(key + "__hide_unselected", False)
                ctrl = self._build_sheet_list(
                    current, sortable=sortable, hide_unselected=hide_opt,
                    sort_by=sort_by, hide_val=hide_val, key=key
                )
                self.controls[key] = ctrl
                settings_panel.Children.Add(ctrl)
            elif stype == "text":
                tb = TextBox()
                tb.Text = current or ""
                tb.Height = 28
                tb.FontSize = 13
                tb.Margin = Thickness(0, 0, 0, 4)
                self.controls[key] = tb
                settings_panel.Children.Add(tb)
            elif stype == "number":
                tb = TextBox()
                tb.Text = str(current) if current is not None else "0"
                tb.Height = 28
                tb.FontSize = 13
                tb.Margin = Thickness(0, 0, 0, 4)
                self.controls[key] = tb
                settings_panel.Children.Add(tb)
            elif stype == "checkbox":
                cb = CheckBox()
                cb.Content = item_def.get("checkbox_label", "")
                cb.FontSize = 12
                cb.IsChecked = bool(current) if current is not None else False
                cb.Margin = Thickness(0, 0, 0, 4)
                self.controls[key] = cb
                settings_panel.Children.Add(cb)
            elif stype == "select":
                cmb = ComboBox()
                cmb.Height = 28
                cmb.FontSize = 12
                cmb.Margin = Thickness(0, 0, 0, 4)
                for opt in item_def.get("options", []):
                    cmb.Items.Add(opt)
                if current and current in item_def.get("options", []):
                    cmb.SelectedItem = current
                elif cmb.Items.Count > 0:
                    cmb.SelectedIndex = 0
                self.controls[key] = cmb
                settings_panel.Children.Add(cmb)

        scroll.Content = settings_panel
        stack.Children.Add(scroll)

        # Кнопки
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 12, 0, 0)

        btn_ok = Button()
        btn_ok.Content = u"Сохранить"
        btn_ok.Width = 100
        btn_ok.Height = 30
        btn_ok.Margin = Thickness(0, 0, 8, 0)
        btn_ok.IsDefault = True
        btn_ok.Click += self._on_ok
        btn_panel.Children.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Content = u"Отмена"
        btn_cancel.Width = 100
        btn_cancel.Height = 30
        btn_cancel.IsCancel = True
        btn_panel.Children.Add(btn_cancel)

        stack.Children.Add(btn_panel)
        self.Content = stack

    def _get_sheet_params(self, sheets):
        """Собрать уникальные имена параметров из листов."""
        if not sheets:
            return []
        param_names = []
        seen = set()
        defaults = [u"Номер листа", u"Имя листа"]
        for d in defaults:
            param_names.append(d)
            seen.add(d)
        for sheet in sheets:
            for p in sheet.Parameters:
                pname = p.Definition.Name
                if pname not in seen:
                    seen.add(pname)
                    param_names.append(pname)
        return param_names

    def _get_param_value(self, sheet, param_name):
        """Получить строковое значение параметра листа."""
        if param_name == u"Номер листа":
            return sheet.SheetNumber or ""
        if param_name == u"Имя листа":
            return sheet.Name or ""
        for p in sheet.Parameters:
            if p.Definition.Name == param_name:
                return p.AsValueString() or p.AsString() or ""
        return ""

    def _build_sheet_list(self, selected_numbers, sortable=False,
                          hide_unselected=False, sort_by=None,
                          hide_val=False, key=""):
        """Построить список листов с чекбоксами."""
        border = Border()
        border.BorderBrush = SolidColorBrush(Colors.Gray)
        border.BorderThickness = Thickness(1)
        border.MaxHeight = 300

        outer = StackPanel()

        try:
            collector = DB.FilteredElementCollector(self.doc) \
                .OfClass(DB.ViewSheet) \
                .WhereElementIsNotElementType()
            sheets = list(collector.ToElements())
        except Exception:
            sheets = []

        # Toolbar
        toolbar = StackPanel()
        toolbar.Orientation = System.Windows.Controls.Orientation.Horizontal
        toolbar.Margin = Thickness(4, 4, 4, 2)

        sort_combo = None
        if sortable and sheets:
            lbl_sort = TextBlock()
            lbl_sort.Text = u"Сортировка: "
            lbl_sort.FontSize = 11
            lbl_sort.VerticalAlignment = System.Windows.VerticalAlignment.Center
            toolbar.Children.Add(lbl_sort)

            sort_combo = ComboBox()
            sort_combo.FontSize = 11
            sort_combo.MinWidth = 150
            sort_combo.Margin = Thickness(0, 0, 12, 0)
            param_names = self._get_sheet_params(sheets)
            for pn in param_names:
                sort_combo.Items.Add(pn)
            if sort_by and sort_by in param_names:
                sort_combo.SelectedItem = sort_by
            else:
                sort_combo.SelectedIndex = 0
            toolbar.Children.Add(sort_combo)

        hide_cb = None
        if hide_unselected:
            hide_cb = CheckBox()
            hide_cb.Content = u"Скрывать невыбранные"
            hide_cb.FontSize = 11
            hide_cb.IsChecked = hide_val
            hide_cb.VerticalAlignment = System.Windows.VerticalAlignment.Center
            toolbar.Children.Add(hide_cb)

        if toolbar.Children.Count > 0:
            outer.Children.Add(toolbar)

        sep = Separator()
        outer.Children.Add(sep)

        # Сохраняем ссылки на self (IronPython не позволяет setattr на WPF-объекты)
        self._sheet_extras[key] = {
            "sort_combo": sort_combo,
            "hide_cb": hide_cb,
            "sheets": sheets
        }

        # Список листов
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto

        panel = StackPanel()
        panel.Margin = Thickness(4)

        self._sheet_checks = []
        self._sheet_rows = []

        if sort_combo:
            sort_param = sort_combo.SelectedItem or u"Номер листа"
        else:
            sort_param = sort_by or u"Номер листа"

        sorted_sheets = sorted(
            sheets,
            key=lambda s: self._get_param_value(s, sort_param)
        )

        for sheet in sorted_sheets:
            row = StackPanel()
            row.Orientation = System.Windows.Controls.Orientation.Horizontal

            cb = CheckBox()
            num = sheet.SheetNumber
            name = sheet.Name
            cb.Content = u"{} - {}".format(num, name)
            cb.FontSize = 12
            cb.Tag = num
            cb.Margin = Thickness(2, 1, 2, 1)
            if selected_numbers and num in selected_numbers:
                cb.IsChecked = True

            idx = len(self._sheet_checks)
            self._sheet_checks.append(cb)
            self._sheet_rows.append(row)
            row.Children.Add(cb)

            if hide_unselected and hide_cb and hide_cb.IsChecked and not cb.IsChecked:
                row.Visibility = Visibility.Collapsed

            panel.Children.Add(row)

        # Скрытие
        if hide_unselected and hide_cb:
            def _on_hide_changed(s, evt):
                is_hidden = s.IsChecked
                for i, cb in enumerate(self._sheet_checks):
                    r = self._sheet_rows[i]
                    if is_hidden and not cb.IsChecked:
                        r.Visibility = Visibility.Collapsed
                    else:
                        r.Visibility = Visibility.Visible

            hide_cb.Checked += _on_hide_changed
            hide_cb.Unchecked += _on_hide_changed

            for i, cb in enumerate(self._sheet_checks):
                def _on_sheet_unchecked(s, evt, _idx=i):
                    if hide_cb.IsChecked and not s.IsChecked:
                        self._sheet_rows[_idx].Visibility = Visibility.Collapsed
                cb.Unchecked += _on_sheet_unchecked

        # Пересортировка
        if sortable and sort_combo:
            def _on_sort_changed(s, evt):
                param_name = s.SelectedItem
                if not param_name:
                    return

                states = {}
                for cb in self._sheet_checks:
                    states[cb.Tag] = cb.IsChecked

                panel.Children.Clear()
                self._sheet_checks = []
                self._sheet_rows = []

                new_sorted = sorted(
                    sheets,
                    key=lambda sh: self._get_param_value(sh, param_name)
                )

                for sheet in new_sorted:
                    row = StackPanel()
                    row.Orientation = System.Windows.Controls.Orientation.Horizontal

                    cb = CheckBox()
                    num = sheet.SheetNumber
                    name = sheet.Name
                    cb.Content = u"{} - {}".format(num, name)
                    cb.FontSize = 12
                    cb.Tag = num
                    cb.Margin = Thickness(2, 1, 2, 1)
                    if num in states:
                        cb.IsChecked = states[num]

                    new_idx = len(self._sheet_checks)
                    self._sheet_checks.append(cb)
                    self._sheet_rows.append(row)
                    row.Children.Add(cb)

                    if hide_unselected and hide_cb and hide_cb.IsChecked and not cb.IsChecked:
                        row.Visibility = Visibility.Collapsed

                    panel.Children.Add(row)

                # Перепривязать скрытие
                if hide_unselected and hide_cb:
                    for i, cb in enumerate(self._sheet_checks):
                        def _on_sheet_unchecked2(s, evt, _idx=i):
                            if hide_cb.IsChecked and not s.IsChecked:
                                self._sheet_rows[_idx].Visibility = Visibility.Collapsed
                        cb.Unchecked += _on_sheet_unchecked2

            sort_combo.SelectionChanged += _on_sort_changed

        if not sheets:
            txt = TextBlock()
            txt.Text = u"Листы не найдены"
            txt.Foreground = SolidColorBrush(Colors.Gray)
            txt.FontStyle = System.Windows.FontStyles.Italic
            panel.Children.Add(txt)

        scroll.Content = panel
        outer.Children.Add(scroll)
        border.Child = outer

        return border

    def _on_ok(self, sender, e):
        """Собрать настройки и закрыть."""
        settings = {}

        for key, ctrl in self.controls.items():
            if isinstance(ctrl, Border):
                # sheet_list
                selected = []
                for cb in self._sheet_checks:
                    if cb.IsChecked:
                        selected.append(cb.Tag)
                settings[key] = selected

                extras = self._sheet_extras.get(key, {})
                sort_combo = extras.get("sort_combo")
                hide_cb = extras.get("hide_cb")
                if sort_combo and sort_combo.SelectedItem:
                    settings[key + "__sort_by"] = sort_combo.SelectedItem
                if hide_cb is not None:
                    settings[key + "__hide_unselected"] = bool(hide_cb.IsChecked)
            elif isinstance(ctrl, CheckBox):
                settings[key] = bool(ctrl.IsChecked)
            elif isinstance(ctrl, ComboBox):
                settings[key] = ctrl.SelectedItem or ""
            elif isinstance(ctrl, TextBox):
                settings[key] = ctrl.Text

        self.result_settings = settings
        self.DialogResult = True

    def show_dialog(self):
        """Показать диалог и вернуть настройки."""
        if self.ShowDialog() == True:
            return self.result_settings
        return None
