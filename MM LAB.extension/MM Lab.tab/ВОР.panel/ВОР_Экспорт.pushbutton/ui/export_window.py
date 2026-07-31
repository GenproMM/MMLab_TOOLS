# -*- coding: utf-8 -*-
"""
Логика главного окна ВОР Экспорт.
Code-behind для export_window.xaml.
"""

import os
import sys
import json
import codecs
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")

import System
from System.Windows import (
    Window, Thickness, VerticalAlignment, Visibility, MessageBox,
    MessageBoxButton, MessageBoxImage, MessageBoxResult,
    TextAlignment, TextWrapping, TextTrimming,
    FontWeights, FontStyles,
)
from System.Windows.Controls import (
    ComboBox, ComboBoxItem, Button, StackPanel, Grid,
    TextBlock, Border, ContextMenu, MenuItem, Separator,
    ColumnDefinition,
)
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Colors, Stretch
from System.Windows.Shapes import Path as WpfPath

from pyrevit import revit, script
from pyrevit.forms import alert

doc = revit.doc
logger = script.get_logger()

# ---- Локальные модули (insert(0) — высший приоритет) ----
local_path = os.path.dirname(os.path.abspath(__file__))
parent_path = os.path.dirname(local_path)
if parent_path not in sys.path:
    sys.path.insert(0, parent_path)

# ---- Импорт общих модулей из ВОР_Валидатор (append — низший приоритет) ----
bundle_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
validator_path = os.path.normpath(
    os.path.join(bundle_path, "..", u"\u0412\u041e\u0420_\u0412\u0430\u043b\u0438\u0434\u0430\u0442\u043e\u0440.pushbutton")
)
if validator_path not in sys.path:
    sys.path.append(validator_path)

from core.user_paths import get_config_file
from core.settings_store import (
    load_last_selection, save_last_selection,
    get_sections, get_projects,
)

# ---- Импорт окна настроек валидатора ----
ui_path = os.path.join(validator_path, "ui")
if ui_path not in sys.path:
    sys.path.append(ui_path)
from settings_window import SettingsWindow

from schedule_reader import find_schedule_by_name, read_schedule, check_column_counts, get_all_schedules
from excel_builder import read_excel_file, build_export, is_file_locked
from ui.schedule_picker import SchedulePickerWindow


# ---- Шаблон иконок ----
_ROUNDED_ICON_TEMPLATE = XamlReader.Parse(
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="Button">'
    '<Border x:Name="Bd" CornerRadius="4" Background="#F0F0F0" '
    'BorderBrush="#D0D0D0" BorderThickness="1" Width="26" Height="26">'
    '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
    '</Border>'
    '<ControlTemplate.Triggers>'
    '<Trigger Property="IsMouseOver" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#D8D8D8"/>'
    '</Trigger>'
    '<Trigger Property="IsPressed" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#C8C8C8"/>'
    '</Trigger>'
    '</ControlTemplate.Triggers>'
    '</ControlTemplate>'
)


# ---- Конфиг: загрузка/сохранение export_configs ----

def _load_full_config():
    cf = get_config_file()
    if os.path.exists(cf):
        try:
            with codecs.open(cf, "r", "utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"sections": [], "projects": [], "rp_configs": {}}


def _save_full_config(full):
    with codecs.open(get_config_file(), "w", "utf-8") as f:
        json.dump(full, f, indent=2, ensure_ascii=False)


def _load_export_config(section, project):
    """Загрузить конфигурацию экспорта для пары раздел+проект."""
    key = u"{}|{}".format(section, project)
    full = _load_full_config()
    return full.get("export_configs", {}).get(key, None)


def _save_export_config(section, project, sources, output_path):
    """Сохранить конфигурацию экспорта, не трогая остальные ключи."""
    key = u"{}|{}".format(section, project)
    full = _load_full_config()
    ec = full.get("export_configs", {})
    ec[key] = {
        "sources": sources,
        "output_path": output_path or "",
    }
    full["export_configs"] = ec
    _save_full_config(full)


# ---- Главное окно ----

class ExportMainWindow(object):

    def __init__(self):
        self.sources = []
        self._suppress_auto_load = False

        xaml_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "export_window.xaml"
        )
        with codecs.open(xaml_path, "r", "utf-8") as f:
            xaml = f.read()

        self.window = XamlReader.Parse(xaml)

        # Найти элементы
        self.cmb_section = self.window.FindName("CmbSection")
        self.cmb_project = self.window.FindName("CmbProject")
        self.btn_settings = self.window.FindName("BtnSettings")
        self.txt_output_path = self.window.FindName("TxtOutputPath")
        self.btn_browse_output = self.window.FindName("BtnBrowseOutput")
        self.btn_add_source = self.window.FindName("BtnAddSource")
        self.sources_panel = self.window.FindName("SourcesPanel")
        self.warning_border = self.window.FindName("WarningBorder")
        self.txt_warning = self.window.FindName("TxtWarning")
        self.btn_export = self.window.FindName("BtnExport")

        # События
        self.cmb_section.SelectionChanged += self._on_selection_changed
        self.cmb_project.SelectionChanged += self._on_selection_changed
        self.btn_settings.Click += self._on_open_settings
        self.btn_browse_output.Click += self._on_browse_output
        self.btn_add_source.Click += self._on_add_source
        self.btn_export.Click += self._on_export
        self.txt_output_path.TextChanged += self._on_output_path_changed

        # Контекстное меню для кнопки "Добавить"
        self._ctx_menu = ContextMenu()
        mi_schedule = MenuItem()
        mi_schedule.Header = u"\u0414\u043e\u0431\u0430\u0432\u0438\u0442\u044c \u0441\u043f\u0435\u0446\u0438\u0444\u0438\u043a\u0430\u0446\u0438\u044e"
        mi_schedule.Click += lambda s, e: self._on_add_schedule()
        self._ctx_menu.Items.Add(mi_schedule)

        mi_excel = MenuItem()
        mi_excel.Header = u"\u0414\u043e\u0431\u0430\u0432\u0438\u0442\u044c Excel \u0444\u0430\u0439\u043b"
        mi_excel.Click += lambda s, e: self._on_add_excel()
        self._ctx_menu.Items.Add(mi_excel)

        self.btn_add_source.ContextMenu = self._ctx_menu

        # Заполнить комбобоксы
        self._suppress_auto_load = True
        self._populate_combos()
        self._restore_last_selection()
        self._suppress_auto_load = False

        # Загрузить конфиг для текущего выбора
        self._load_config_for_current_rp()

        # Глобальный перехват необработанных исключений — закрыть окно плагина
        from System.Windows import Application
        app = Application.Current
        if app:
            app.DispatcherUnhandledException += self._on_unhandled_exception

        self.window.ShowDialog()

    def _on_unhandled_exception(self, sender, args):
        import traceback as tb
        logger.error(tb.format_exception_only(type(args.Exception), args.Exception))
        args.Handled = True
        try:
            self.window.Close()
        except Exception:
            pass

    # ---- Комбобоксы ----

    def _populate_combos(self):
        sections = get_sections()
        for s in sections:
            item = ComboBoxItem()
            item.Content = s
            self.cmb_section.Items.Add(item)

        projects = get_projects()
        for p in projects:
            item = ComboBoxItem()
            item.Content = p
            self.cmb_project.Items.Add(item)

    def _restore_last_selection(self):
        section, project = load_last_selection()
        if section:
            self._set_combo_by_text(self.cmb_section, section)
        if project:
            self._set_combo_by_text(self.cmb_project, project)

    def _set_combo_by_text(self, combo, text):
        for i in range(combo.Items.Count):
            item = combo.Items[i]
            if item.Content == text:
                combo.SelectedIndex = i
                return

    def _get_selected_section(self):
        item = self.cmb_section.SelectedItem
        return item.Content if item else None

    def _get_selected_project(self):
        item = self.cmb_project.SelectedItem
        return item.Content if item else None

    # ---- События ----

    def _on_selection_changed(self, sender, e):
        if self._suppress_auto_load:
            return
        section = self._get_selected_section()
        project = self._get_selected_project()
        if section and project:
            save_last_selection(section, project)
        self._load_config_for_current_rp()

    def _on_output_path_changed(self, sender, e):
        self._auto_save()

    def _on_open_settings(self, sender, e):
        SettingsWindow()
        # Обновить комбобоксы после возможных изменений
        self._suppress_auto_load = True
        self.cmb_section.Items.Clear()
        self.cmb_project.Items.Clear()
        self._populate_combos()
        self._restore_last_selection()
        self._suppress_auto_load = False

    def _on_browse_output(self, sender, e):
        from System.Windows.Forms import SaveFileDialog
        dlg = SaveFileDialog()
        dlg.Filter = u"Excel \u0444\u0430\u0439\u043b\u044b (*.xlsx)|*.xlsx"
        dlg.DefaultExt = ".xlsx"
        if dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK:
            self.txt_output_path.Text = dlg.FileName

    def _on_add_source(self, sender, e):
        self._ctx_menu.IsOpen = True

    def _on_add_schedule(self):
        existing_names = [s["name"] for s in self.sources if s["type"] == "schedule"]
        picker = SchedulePickerWindow(doc, selected_names=existing_names)
        if picker.results:
            existing = {s["name"] for s in self.sources if s["type"] == "schedule"}
            for r in picker.results:
                if r["name"] not in existing:
                    self.sources.append({"type": "schedule", "name": r["name"]})
                    existing.add(r["name"])
            self._rebuild_sources_panel()
            self._auto_save()

    def _on_add_excel(self):
        from System.Windows.Forms import OpenFileDialog
        dlg = OpenFileDialog()
        dlg.Filter = u"Excel \u0444\u0430\u0439\u043b\u044b (*.xlsx)|*.xlsx"
        if dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK:
            source = {"type": "excel", "path": dlg.FileName}
            self.sources.append(source)
            self._rebuild_sources_panel()
            self._auto_save()

    def _on_export(self, sender, e):
        section = self._get_selected_section()
        project = self._get_selected_project()
        output_path = self.txt_output_path.Text.strip()

        if not section or not project:
            alert(u"\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u0440\u0430\u0437\u0434\u0435\u043b \u0438 \u043f\u0440\u043e\u0435\u043a\u0442")
            return

        if not self.sources:
            alert(u"\u0414\u043e\u0431\u0430\u0432\u044c\u0442\u0435 \u0445\u043e\u0442\u044f \u0431\u044b \u043e\u0434\u0438\u043d \u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a \u0434\u0430\u043d\u043d\u044b\u0445")
            return

        if not output_path:
            alert(u"\u0423\u043a\u0430\u0436\u0438\u0442\u0435 \u043f\u0443\u0442\u044c \u0432\u044b\u0433\u0440\u0443\u0437\u043a\u0438")
            return

        # Собрать данные из источников
        sources_data = []
        schedule_names = []
        missing = []

        for src in self.sources:
            if src["type"] == "schedule":
                sched = find_schedule_by_name(doc, src["name"])
                if not sched:
                    missing.append(src["name"])
                    continue
                schedule_names.append(src["name"])
                try:
                    data = read_schedule(sched)
                    sources_data.append(data)
                except Exception as ex:
                    logger.error(u"\u041e\u0448\u0438\u0431\u043a\u0430 \u0447\u0442\u0435\u043d\u0438\u044f '{}': {}".format(src["name"], ex))
                    missing.append(src["name"])

            elif src["type"] == "excel":
                if not os.path.exists(src["path"]):
                    missing.append(os.path.basename(src["path"]))
                    continue
                try:
                    data = read_excel_file(src["path"])
                    sources_data.append(data)
                except Exception as ex:
                    logger.error(u"\u041e\u0448\u0438\u0431\u043a\u0430 \u0447\u0442\u0435\u043d\u0438\u044f '{}': {}".format(src["path"], ex))
                    missing.append(os.path.basename(src["path"]))

        if missing:
            msg = u"\u041d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u044b:\n" + u"\n".join(missing)
            result = MessageBox.Show(
                msg + u"\n\n\u041f\u0440\u043e\u0434\u043e\u043b\u0436\u0438\u0442\u044c \u0441 \u043e\u0441\u0442\u0430\u0432\u0448\u0438\u043c\u0438\u0441\u044f?",
                u"\u041f\u0440\u0435\u0434\u0443\u043f\u0440\u0435\u0436\u0434\u0435\u043d\u0438\u0435",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
            )
            if result != MessageBoxResult.Yes:
                return

        if not sources_data:
            alert(u"\u041d\u0435\u0442 \u0434\u043e\u0441\u0442\u0443\u043f\u043d\u044b\u0445 \u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u043e\u0432 \u0434\u043b\u044f \u044d\u043a\u0441\u043f\u043e\u0440\u0442\u0430")
            return

        # Проверка столбцов спецификаций
        if len(schedule_names) > 1:
            col_counts = check_column_counts(doc, schedule_names)
            counts = [c for _, c in col_counts]
            if len(set(counts)) > 1:
                details = u", ".join(
                    u"{}: {}".format(n, c) for n, c in col_counts
                )
                self.txt_warning.Text = (
                    u"\u0421\u043f\u0435\u0446\u0438\u0444\u0438\u043a\u0430\u0446\u0438\u0438 \u0438\u043c\u0435\u044e\u0442 \u0440\u0430\u0437\u043d\u043e\u0435 \u043a\u043e\u043b\u0438\u0447\u0435\u0441\u0442\u0432\u043e \u0441\u0442\u043e\u043b\u0431\u0446\u043e\u0432: " + details
                )
                self.warning_border.Visibility = Visibility.Visible

                result = MessageBox.Show(
                    self.txt_warning.Text + u"\n\n\u041f\u0440\u043e\u0434\u043e\u043b\u0436\u0438\u0442\u044c \u044d\u043a\u0441\u043f\u043e\u0440\u0442?",
                    u"\u041d\u0435\u0441\u043e\u0432\u043f\u0430\u0434\u0435\u043d\u0438\u0435 \u0441\u0442\u043e\u043b\u0431\u0446\u043e\u0432",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning,
                )
                if result != MessageBoxResult.Yes:
                    return
            else:
                self.warning_border.Visibility = Visibility.Collapsed

        # Проверка: не открыт ли целевой файл в другой программе
        if is_file_locked(output_path):
            MessageBox.Show(
                u"\u0424\u0430\u0439\u043b '{}' \u043e\u0442\u043a\u0440\u044b\u0442 \u0432 \u0434\u0440\u0443\u0433\u043e\u0439 \u043f\u0440\u043e\u0433\u0440\u0430\u043c\u043c\u0435.\n\u0417\u0430\u043a\u0440\u043e\u0439\u0442\u0435 \u0444\u0430\u0439\u043b \u0438 \u043f\u043e\u0432\u0442\u043e\u0440\u0438\u0442\u0435 \u043f\u043e\u043f\u044b\u0442\u043a\u0443.".format(os.path.basename(output_path)),
                u"\u0424\u0430\u0439\u043b \u0437\u0430\u043d\u044f\u0442",
                MessageBoxButton.OK,
                MessageBoxImage.Warning,
            )
            self.window.Close()
            return

        # Экспорт
        try:
            total = build_export(sources_data, output_path)
            MessageBox.Show(
                u"\u042d\u043a\u0441\u043f\u043e\u0440\u0442 \u0437\u0430\u0432\u0435\u0440\u0448\u0451\u043d.\n\u0421\u0442\u0440\u043e\u043a: {}".format(total),
                u"\u0413\u043e\u0442\u043e\u0432\u043e",
                MessageBoxButton.OK,
                MessageBoxImage.Information,
            )
            os.startfile(output_path)
            self.window.Close()
        except Exception as ex:
            import traceback
            logger.error(u"\u041e\u0448\u0438\u0431\u043a\u0430 \u044d\u043a\u0441\u043f\u043e\u0440\u0442\u0430: {}".format(ex))
            logger.error(traceback.format_exc())
            alert(u"\u041e\u0448\u0438\u0431\u043a\u0430 \u044d\u043a\u0441\u043f\u043e\u0440\u0442\u0430:\n\u041f\u0443\u0442\u044c: {}\n{}".format(output_path, ex))
            self.window.Close()

    # ---- Панель источников ----

    def _rebuild_sources_panel(self):
        self.sources_panel.Children.Clear()
        for i, src in enumerate(self.sources):
            is_missing = self._check_missing(src)
            self._add_source_row(i, src, is_missing)

    def _check_missing(self, src):
        if src["type"] == "schedule":
            return find_schedule_by_name(doc, src["name"]) is None
        elif src["type"] == "excel":
            return not os.path.exists(src["path"])
        return True

    def _add_source_row(self, index, source, is_missing=False):
        row = Grid()
        row.Margin = Thickness(0, 2, 0, 2)
        row.VerticalAlignment = VerticalAlignment.Center

        if is_missing:
            row.Background = SolidColorBrush(
                System.Windows.Media.Color.FromArgb(30, 255, 200, 0)
            )

        # Колонка 0: имя (растягивается), Колонка 1: кнопки (auto)
        col_name = ColumnDefinition()
        col_name.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
        row.ColumnDefinitions.Add(col_name)

        col_btns = ColumnDefinition()
        col_btns.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Auto)
        row.ColumnDefinitions.Add(col_btns)

        # Левая часть: бейдж + текст
        name_panel = StackPanel()
        name_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        name_panel.VerticalAlignment = VerticalAlignment.Center
        System.Windows.Controls.Grid.SetColumn(name_panel, 0)

        badge = TextBlock()
        badge.Width = 24
        badge.TextAlignment = System.Windows.TextAlignment.Center
        badge.FontSize = 11
        badge.FontWeight = FontWeights.Bold
        badge.Margin = Thickness(4, 0, 8, 0)
        badge.VerticalAlignment = VerticalAlignment.Center

        if source["type"] == "schedule":
            badge.Text = "S"
            badge.Foreground = SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0, 122, 204)
            )
        else:
            badge.Text = "X"
            badge.Foreground = SolidColorBrush(
                System.Windows.Media.Color.FromRgb(76, 175, 80)
            )
        name_panel.Children.Add(badge)

        name_text = TextBlock()
        name_text.VerticalAlignment = VerticalAlignment.Center
        name_text.FontSize = 13
        name_text.Margin = Thickness(0, 0, 8, 0)

        if source["type"] == "schedule":
            display = source["name"]
        else:
            display = os.path.basename(source.get("path", ""))

        if is_missing:
            name_text.Text = display + u" (\u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d)"
            name_text.FontStyle = FontStyles.Italic
            name_text.Foreground = SolidColorBrush(Colors.Gray)
        else:
            name_text.Text = display

        name_panel.Children.Add(name_text)
        row.Children.Add(name_panel)

        # Правая часть: кнопки
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.VerticalAlignment = VerticalAlignment.Center
        System.Windows.Controls.Grid.SetColumn(btn_panel, 1)

        btn_up = self._create_icon_button(
            "M6,16 L12,8 L18,16",
            lambda s, e, idx=index: self._on_move_up(idx),
        )
        btn_up.ToolTip = u"\u0412\u0432\u0435\u0440\u0445"
        btn_panel.Children.Add(btn_up)

        btn_down = self._create_icon_button(
            "M6,8 L12,16 L18,8",
            lambda s, e, idx=index: self._on_move_down(idx),
        )
        btn_down.ToolTip = u"\u0412\u043d\u0438\u0437"
        btn_panel.Children.Add(btn_down)

        btn_delete = self._create_icon_button(
            "M6,6 L18,18 M18,6 L6,18",
            lambda s, e, idx=index: self._on_delete_source(idx),
        )
        btn_delete.ToolTip = u"\u0423\u0434\u0430\u043b\u0438\u0442\u044c"
        btn_panel.Children.Add(btn_delete)

        row.Children.Add(btn_panel)

        self.sources_panel.Children.Add(row)

    def _create_icon_button(self, path_data, click_handler):
        btn = Button()
        btn.Template = _ROUNDED_ICON_TEMPLATE
        btn.Margin = Thickness(2, 0, 0, 0)

        icon = WpfPath()
        icon.Data = System.Windows.Media.PathGeometry.Parse(path_data)
        icon.Stroke = SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100))
        icon.StrokeThickness = 1.5
        icon.Stretch = Stretch.Uniform
        icon.Width = 12
        icon.Height = 12

        btn.Content = icon
        btn.Click += click_handler
        return btn

    def _on_move_up(self, index):
        if index <= 0:
            return
        self.sources[index], self.sources[index - 1] = (
            self.sources[index - 1], self.sources[index]
        )
        self._rebuild_sources_panel()
        self._auto_save()

    def _on_move_down(self, index):
        if index >= len(self.sources) - 1:
            return
        self.sources[index], self.sources[index + 1] = (
            self.sources[index + 1], self.sources[index]
        )
        self._rebuild_sources_panel()
        self._auto_save()

    def _on_delete_source(self, index):
        if 0 <= index < len(self.sources):
            self.sources.pop(index)
            self._rebuild_sources_panel()
            self._auto_save()

    # ---- Конфиг ----

    def _load_config_for_current_rp(self):
        section = self._get_selected_section()
        project = self._get_selected_project()
        if not section or not project:
            return

        config = _load_export_config(section, project)
        if config:
            self.sources = config.get("sources", [])
            self.txt_output_path.Text = config.get("output_path", "")
        else:
            self.sources = []
            self.txt_output_path.Text = ""

        self.warning_border.Visibility = Visibility.Collapsed
        self._rebuild_sources_panel()

    def _auto_save(self):
        section = self._get_selected_section()
        project = self._get_selected_project()
        if not section or not project:
            return
        _save_export_config(section, project, self.sources, self.txt_output_path.Text.strip())
