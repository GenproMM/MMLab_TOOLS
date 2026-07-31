# -*- coding: utf-8 -*-
"""
Логика основного окна ВОР Валидатора.
Code-behind для main_window.xaml.
"""

import os
import sys
import imp
import clr
import codecs
import traceback

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import Window, GridLength, GridUnitType, Thickness, VerticalAlignment
from System.Windows.Controls import *
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Colors, Stretch
from System.Windows.Shapes import Path as WpfPath

from pyrevit import revit, script

bundle_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if bundle_path not in sys.path:
    sys.path.insert(0, bundle_path)

from core.config_manager import (
    load_rp_config, save_rp_config,
    load_script_settings, save_script_settings,
    resolve_script_paths
)
from core.registry import register_script
from core.section_project import get_available_sections, get_available_projects
from core.settings_store import load_last_selection, save_last_selection
from core.validation_engine import extract_script_metadata
from core.report_store import load_report

logger = script.get_logger()

_ROUNDED_ICON_TEMPLATE = XamlReader.Parse(
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="Button">'
    '<Border x:Name="Bd" CornerRadius="4" Background="#F0F0F0" '
    'BorderBrush="#D0D0D0" BorderThickness="1">'
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

# Иконка урны (Path geometry)
_TRASH_PATH_DATA = (
    "M2,4 L10,4 "
    "M3,4 L3.3,11 C3.3,11.6 3.7,12 4.3,12 L7.7,12 "
    "C8.3,12 8.7,11.6 8.7,11 L9,4 "
    "M4.5,4 L4.5,2.5 C4.5,2.2 4.7,2 5,2 L7,2 "
    "C7.3,2 7.5,2.2 7.5,2.5 L7.5,4 "
    "M5,6 L5,10 M6,6 L6,10 M7,6 L7,10"
)

# Иконка шестерёнки (Path geometry)
_GEAR_PATH_DATA = (
    "M6,2.5 L7,2.5 L7.3,3.5 L8.2,3.8 L9,3.1 L9.7,3.8 "
    "L9,4.7 L9.2,5.5 L10.3,5.8 L10.3,6.8 L9.2,7.1 "
    "L9,7.9 L9.7,8.8 L9,9.5 L8.1,8.8 L7.3,9.1 L7,10.2 "
    "L6,10.2 L5.7,9.1 L4.9,8.8 L4,9.5 L3.3,8.8 "
    "L4,7.9 L3.8,7.1 L2.7,6.8 L2.7,5.8 L3.8,5.5 "
    "L4,4.7 L3.3,3.8 L4,3.1 L4.9,3.8 L5.7,3.5 Z "
    "M6.5,4.8 A1.5,1.5,0,1,0,6.5,7.8 A1.5,1.5,0,1,0,6.5,4.8"
)

# Иконки стрелок вверх/вниз (шевроны) — для упорядочивания скриптов
_UP_PATH_DATA = "M6,16 L12,8 L18,16"
_DOWN_PATH_DATA = "M6,8 L12,16 L18,8"


def _create_icon_button(path_data, tooltip, tag, click_handler, stroke_color=None):
    """Создать кнопку с иконкой (Path)."""
    if stroke_color is None:
        stroke_color = SolidColorBrush(Colors.Gray)

    btn = Button()
    btn.Width = 26
    btn.Height = 26
    btn.Padding = Thickness(2)
    btn.Tag = tag
    btn.ToolTip = tooltip
    btn.Click += click_handler
    btn.VerticalAlignment = VerticalAlignment.Center
    btn.Template = _ROUNDED_ICON_TEMPLATE

    icon = WpfPath()
    icon.Data = System.Windows.Media.Geometry.Parse(path_data)
    icon.Stroke = stroke_color
    icon.StrokeThickness = 1.2
    icon.Stretch = Stretch.Uniform
    icon.Width = 14
    icon.Height = 14
    btn.Content = icon

    return btn


class MainWindow(Window):
    """Основное окно ВОР Валидатора."""

    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), "main_window.xaml")
        with codecs.open(xaml_path, "r", "utf-8") as f:
            xaml_content = f.read()

        window = XamlReader.Parse(xaml_content)

        self.window = window
        self.cmb_section = window.FindName("CmbSection")
        self.cmb_project = window.FindName("CmbProject")
        self.btn_settings = window.FindName("BtnSettings")
        self.btn_add_check = window.FindName("BtnAddCheck")
        self.btn_last_report = window.FindName("BtnLastReport")
        self.scripts_panel = window.FindName("ScriptsPanel")
        self.btn_run_validation = window.FindName("BtnRunValidation")
        self.txt_description = window.FindName("TxtDescription")
        self.description_border = window.FindName("DescriptionBorder")

        # Состояние
        self.custom_script_checks = {}    # name -> CheckBox
        self.custom_script_paths = {}     # name -> script_path
        self.script_descriptions = {}     # name -> description
        self.script_has_settings = {}     # name -> bool
        self.script_settings = {}         # name -> settings dict
        self.script_display_names = {}    # name -> display_name
        self.script_ids = {}              # name -> script_id (str or None)
        self.script_missing = {}          # name -> bool (True if file not found)
        self.script_order = []            # name[] в порядке отображения (детерминированный порядок)
        self._suppress_auto_load = False
        self.selected_script_name = None
        self.selected_script_row = None   # DockPanel

        # Привязка событий
        self.btn_settings.Click += self.on_open_settings
        self.cmb_section.SelectionChanged += self.on_section_or_project_changed
        self.cmb_project.SelectionChanged += self.on_section_or_project_changed
        self.btn_add_check.Click += self.on_add_check
        self.btn_last_report.Click += self.on_open_last_report
        self.btn_run_validation.Click += self.on_run_validation

        # Инициализация данных
        self._suppress_auto_load = True
        self.populate_sections()
        self.populate_projects()

        # Восстановление последнего выбора
        last_section, last_project = load_last_selection()
        if last_section:
            self.set_combo_by_text(self.cmb_section, last_section)
        if last_project:
            self.set_combo_by_text(self.cmb_project, last_project)
        self._suppress_auto_load = False

        section = self.get_selected_combo_value(self.cmb_section)
        project = self.get_selected_combo_value(self.cmb_project)
        if section and project:
            self.load_config_for_rp(section, project)
            save_last_selection(section, project)
        else:
            self.scripts_panel.Children.Clear()
            self._add_no_scripts_message()

        self.window.ShowDialog()

    # ================================================================
    # ЗАПОЛНЕНИЕ ДАННЫХ
    # ================================================================

    def populate_sections(self):
        """Заполнить ComboBox разделов."""
        self.cmb_section.Items.Clear()
        sections = get_available_sections()
        if not sections:
            item = ComboBoxItem()
            item.Content = "Создайте раздел..."
            item.Foreground = System.Windows.Media.Brushes.Gray
            item.IsEnabled = False
            self.cmb_section.Items.Add(item)
            return

        for sec in sections:
            item = ComboBoxItem()
            item.Content = sec
            self.cmb_section.Items.Add(item)
        self.cmb_section.SelectedIndex = 0

    def populate_projects(self):
        """Заполнить ComboBox проектов."""
        self.cmb_project.Items.Clear()
        projects = get_available_projects()
        if not projects:
            item = ComboBoxItem()
            item.Content = "Создайте проект..."
            item.Foreground = System.Windows.Media.Brushes.Gray
            item.IsEnabled = False
            self.cmb_project.Items.Add(item)
            return

        for proj in projects:
            item = ComboBoxItem()
            item.Content = proj
            self.cmb_project.Items.Add(item)
        self.cmb_project.SelectedIndex = 0

    # ================================================================
    # НАСТРОЙКИ
    # ================================================================

    def on_open_settings(self, sender, e):
        """Открыть окно настроек разделов и проектов."""
        from ui.settings_window import SettingsWindow

        prev_section = self.get_selected_combo_value(self.cmb_section)
        prev_project = self.get_selected_combo_value(self.cmb_project)

        SettingsWindow()

        self._suppress_auto_load = True
        self.populate_sections()
        self.populate_projects()

        if prev_section:
            self.set_combo_by_text(self.cmb_section, prev_section)
        if prev_project:
            self.set_combo_by_text(self.cmb_project, prev_project)

        self._suppress_auto_load = False

        section = self.get_selected_combo_value(self.cmb_section)
        project = self.get_selected_combo_value(self.cmb_project)
        if section and project:
            self.load_config_for_rp(section, project)

    # ================================================================
    # АВТОЗАГРУЗКА/СОХРАНЕНИЕ Р+П
    # ================================================================

    def on_section_or_project_changed(self, sender, e):
        """Автозагрузка конфигурации при смене Р или П."""
        if self._suppress_auto_load:
            return
        section = self.get_selected_combo_value(self.cmb_section)
        project = self.get_selected_combo_value(self.cmb_project)
        if section and project:
            self.load_config_for_rp(section, project)
            save_last_selection(section, project)

    def load_config_for_rp(self, section, project):
        """Загрузить и применить настройки для Р+П."""
        logger.info("Автозагрузка конфигурации для: {} | {}".format(section, project))

        config = load_rp_config(section, project)

        if config is None:
            logger.info("Нет сохранённых настроек для {} | {}".format(section, project))
            self.custom_script_checks.clear()
            self.custom_script_paths.clear()
            self.script_descriptions.clear()
            self.script_has_settings.clear()
            self.script_settings.clear()
            self.script_display_names.clear()
            self.script_ids.clear()
            self.script_missing.clear()
            self.script_order = []
            self.scripts_panel.Children.Clear()
            self._add_no_scripts_message()
            self._hide_description()
            return

        scripts = config.get("custom_scripts", [])
        self.custom_script_checks.clear()
        self.custom_script_paths.clear()
        self.script_descriptions.clear()
        self.script_has_settings.clear()
        self.script_settings.clear()
        self.script_display_names.clear()
        self.script_ids.clear()
        self.script_missing.clear()
        self.script_order = []
        self.scripts_panel.Children.Clear()

        if not scripts:
            self._add_no_scripts_message()
        else:
            # Резолв путей по ID если нужно
            scripts, warnings = resolve_script_paths(scripts)
            for w in warnings:
                logger.warning(w)

            # Дедупликация по ID
            seen_ids = set()
            for s in scripts:
                sid = s.get("id")
                if sid and sid in seen_ids:
                    logger.warning(u"Duplicate script id '{}' skipped".format(sid))
                    continue
                if sid:
                    seen_ids.add(sid)

                sname = s["name"]
                spath = s["path"]
                is_enabled = s.get("enabled", True)
                script_id = s.get("id")
                is_missing = not os.path.exists(spath)
                settings = {} if is_missing else load_script_settings(sname, section, project)
                self.add_script_to_ui(sname, spath, is_checked=is_enabled, settings=settings,
                                      script_id=script_id, is_missing=is_missing)

        self._hide_description()
        self.selected_script_name = None
        self.selected_script_row = None

    def _auto_save_current_config(self):
        """Автосохранение текущей конфигурации для выбранной Р+П."""
        section = self.get_selected_combo_value(self.cmb_section)
        project = self.get_selected_combo_value(self.cmb_project)
        if not section or not project:
            return

        all_scripts = []
        for sname in self.script_order:
            cb = self.custom_script_checks.get(sname)
            if cb is None:
                continue
            path = self.custom_script_paths.get(sname, "")
            entry = {
                "name": sname,
                "path": path,
                "enabled": bool(cb.IsChecked)
            }
            sid = self.script_ids.get(sname)
            if sid:
                entry["id"] = sid
            all_scripts.append(entry)

        save_rp_config(section, project, {"custom_scripts": all_scripts})

    # ================================================================
    # ДОБАВЛЕНИЕ ПРОВЕРОК
    # ================================================================

    def on_add_check(self, sender, e):
        """Добавить новую проверку — выбор py файла."""
        from System.Windows.Forms import OpenFileDialog

        section = self.get_selected_combo_value(self.cmb_section)
        project = self.get_selected_combo_value(self.cmb_project)

        if not section or not project:
            System.Windows.MessageBox.Show(
                "Сначала выберите раздел и проект.",
                "Ошибка",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning
            )
            return

        dlg = OpenFileDialog()
        dlg.Filter = "Python файлы|*.py"
        dlg.Title = "Выберите скрипт проверки"
        dlg.Multiselect = False
        scripts_dir = os.path.join(bundle_path, "scripts")
        if os.path.isdir(scripts_dir):
            dlg.InitialDirectory = scripts_dir

        if dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK:
            script_path = dlg.FileName
            base_name = os.path.splitext(os.path.basename(script_path))[0]
            # Для папочных скриптов (script.py) — берём имя из родительской папки
            if base_name == "script":
                parent = os.path.basename(os.path.dirname(script_path))
                script_name = parent if parent else base_name
            else:
                script_name = base_name

            if script_name in self.custom_script_checks:
                System.Windows.MessageBox.Show(
                    "Проверка '{}' уже существует.".format(script_name),
                    "Ошибка",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning
                )
                return

            # Извлекаем метаданные для получения SCRIPT_ID
            meta = extract_script_metadata(script_path)
            script_id = meta.get("script_id")

            # Проверяем SCRIPT_ID через реестр
            if script_id:
                ok, msg = register_script(
                    script_id, script_name, script_path, meta.get("name", script_name)
                )
                if not ok:
                    System.Windows.MessageBox.Show(
                        u"Скрипт с ID '{}' уже зарегистрирован для:\n\n{}\n\n"
                        u"Выберите другой скрипт или обновите реестр.".format(script_id, msg),
                        u"Дублирование ID",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Warning
                    )
                    return

            self.add_script_to_ui(script_name, script_path, is_checked=True,
                                  script_id=script_id)
            logger.info("Добавлен скрипт: {} -> {} (id={})".format(
                script_name, script_path, script_id or "none"))

    def add_script_to_ui(self, name, path, is_checked=False, settings=None,
                         script_id=None, is_missing=False):
        """Добавить скрипт в UI панель."""
        # Извлекаем метаданные из файла скрипта (если файл существует)
        display_name = name
        description = ""
        has_settings = False

        if not is_missing and os.path.exists(path):
            meta = extract_script_metadata(path)
            display_name = meta["name"]
            description = meta["description"]
            has_settings = meta["has_settings"]

        # Сохраняем состояние
        self.custom_script_paths[name] = path
        self.script_descriptions[name] = description
        self.script_has_settings[name] = has_settings
        self.script_display_names[name] = display_name
        self.script_ids[name] = script_id
        self.script_missing[name] = is_missing
        if name not in self.script_order:
            self.script_order.append(name)
        if settings:
            self.script_settings[name] = settings
        elif name not in self.script_settings:
            self.script_settings[name] = {}

        self._remove_no_scripts_message()

        # Контейнер строки
        dock = DockPanel()
        dock.Tag = name
        dock.Margin = Thickness(0, 3, 0, 3)
        dock.LastChildFill = True

        if is_missing:
            # Подсветка для ненайденных скриптов
            dock.Background = SolidColorBrush(
                System.Windows.Media.Color.FromArgb(30, 255, 200, 0)
            )

        # 1. Кнопка удаления (урна) — самая правая
        btn_del = _create_icon_button(
            _TRASH_PATH_DATA,
            "Удалить скрипт",
            name,
            self.on_delete_script,
            SolidColorBrush(Colors.Gray)
        )
        DockPanel.SetDock(btn_del, Dock.Right)
        dock.Children.Add(btn_del)

        # 2. Кнопка настроек (шестерёнка) — только если скрипт найден
        if has_settings and not is_missing:
            btn_gear = _create_icon_button(
                _GEAR_PATH_DATA,
                "Настройки скрипта",
                name,
                self.on_script_settings,
                SolidColorBrush(Colors.Gray)
            )
            DockPanel.SetDock(btn_gear, Dock.Right)
            dock.Children.Add(btn_gear)

        # 2b. Стрелки упорядочивания (вниз/вверх)
        if not is_missing:
            btn_down = _create_icon_button(
                _DOWN_PATH_DATA,
                u"\u0412\u043d\u0438\u0437",
                name,
                self.on_move_script_down,
                SolidColorBrush(Colors.Gray)
            )
            DockPanel.SetDock(btn_down, Dock.Right)
            dock.Children.Add(btn_down)

            btn_up = _create_icon_button(
                _UP_PATH_DATA,
                u"\u0412\u0432\u0435\u0440\u0445",
                name,
                self.on_move_script_up,
                SolidColorBrush(Colors.Gray)
            )
            DockPanel.SetDock(btn_up, Dock.Right)
            dock.Children.Add(btn_up)

        # 3. Чекбокс
        cb = CheckBox()
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Margin = Thickness(0, 0, 4, 0)
        if is_missing:
            cb.IsEnabled = False
            cb.IsChecked = False
        else:
            cb.IsChecked = is_checked
            cb.Checked += lambda s, e: self._auto_save_current_config()
            cb.Unchecked += lambda s, e: self._auto_save_current_config()
        self.custom_script_checks[name] = cb
        dock.Children.Add(cb)

        # 4. Название скрипта — клик показывает описание
        txt = TextBlock()
        if is_missing:
            txt.Text = display_name + u" (не найден)"
            txt.Foreground = SolidColorBrush(Colors.Gray)
            txt.FontStyle = System.Windows.FontStyles.Italic
        else:
            txt.Text = display_name
        txt.FontSize = 13
        txt.FontWeight = System.Windows.FontWeights.SemiBold
        txt.VerticalAlignment = VerticalAlignment.Center
        txt.Cursor = System.Windows.Input.Cursors.Hand
        txt.MouseLeftButtonUp += lambda s, e, n=name, d=dock: self._on_script_clicked(n, d)
        dock.Children.Add(txt)

        self.scripts_panel.Children.Add(dock)
        self._auto_save_current_config()

    def _on_script_clicked(self, name, dock_panel):
        """Обработчик клика на строке скрипта — показать описание."""
        # Снимаем выделение с предыдущей строки
        if self.selected_script_row and self.selected_script_row != dock_panel:
            self.selected_script_row.Background = SolidColorBrush(Colors.Transparent)

        # Выделяем текущую
        self.selected_script_name = name
        self.selected_script_row = dock_panel
        dock_panel.Background = SolidColorBrush(
            System.Windows.Media.Color.FromArgb(40, 0, 122, 204)
        )

        # Показываем описание
        desc = self.script_descriptions.get(name, "")
        if desc:
            self.txt_description.Text = desc
            self.description_border.Visibility = System.Windows.Visibility.Visible
        else:
            self._hide_description()

    def _hide_description(self):
        """Скрыть панель описания."""
        self.txt_description.Text = ""
        self.description_border.Visibility = System.Windows.Visibility.Collapsed

    def _remove_no_scripts_message(self):
        """Удалить сообщение 'нет скриптов' если оно есть."""
        to_remove = []
        for child in self.scripts_panel.Children:
            if isinstance(child, TextBlock):
                to_remove.append(child)
        for item in to_remove:
            self.scripts_panel.Children.Remove(item)

    def _add_no_scripts_message(self):
        """Добавить сообщение 'нет скриптов'."""
        txt = TextBlock()
        txt.Text = "Нет добавленных скриптов"
        txt.Foreground = System.Windows.Media.Brushes.Gray
        txt.FontStyle = System.Windows.FontStyles.Italic
        txt.Margin = Thickness(4, 8, 0, 8)
        self.scripts_panel.Children.Add(txt)

    def on_move_script_up(self, sender, e):
        """Сдвинуть скрипт на одну позицию вверх."""
        name = sender.Tag
        if not name or name not in self.script_order:
            return
        idx = self.script_order.index(name)
        if idx <= 0:
            return
        self.script_order[idx], self.script_order[idx - 1] = (
            self.script_order[idx - 1], self.script_order[idx]
        )
        self._rebuild_scripts_panel()
        self._auto_save_current_config()

    def on_move_script_down(self, sender, e):
        """Сдвинуть скрипт на одну позицию вниз."""
        name = sender.Tag
        if not name or name not in self.script_order:
            return
        idx = self.script_order.index(name)
        if idx >= len(self.script_order) - 1:
            return
        self.script_order[idx], self.script_order[idx + 1] = (
            self.script_order[idx + 1], self.script_order[idx]
        )
        self._rebuild_scripts_panel()
        self._auto_save_current_config()

    def _rebuild_scripts_panel(self):
        """Полная перерисовка списка скриптов в порядке self.script_order."""
        # Сохраняем текущие состояния перед очисткой
        saved = []
        for sname in self.script_order:
            cb = self.custom_script_checks.get(sname)
            saved.append({
                "name": sname,
                "path": self.custom_script_paths.get(sname, ""),
                "is_checked": bool(cb.IsChecked) if cb is not None else False,
                "settings": self.script_settings.get(sname, {}),
                "script_id": self.script_ids.get(sname),
                "is_missing": self.script_missing.get(sname, False),
            })

        selected_name = self.selected_script_name
        self.selected_script_name = None
        self.selected_script_row = None
        self.scripts_panel.Children.Clear()
        self.custom_script_checks.clear()

        if not saved:
            self._add_no_scripts_message()
            self._hide_description()
            return

        for s in saved:
            self.add_script_to_ui(
                s["name"], s["path"], is_checked=s["is_checked"],
                settings=s["settings"], script_id=s["script_id"],
                is_missing=s["is_missing"]
            )

        # add_script_to_ui вызывает _auto_save_current_config — это безопасно,
        # но мы не хотим лишних записей. Восстанавливаем выделение:
        if selected_name and selected_name in self.script_order:
            # Найдём DockPanel по Tag и выделим
            for child in self.scripts_panel.Children:
                if hasattr(child, "Tag") and child.Tag == selected_name:
                    self._on_script_clicked(selected_name, child)
                    break

    def on_delete_script(self, sender, e):
        """Удалить скрипт из текущей конфигурации."""
        script_name = sender.Tag
        if not script_name:
            return
        result = System.Windows.MessageBox.Show(
            "Удалить '{}' из текущей конфигурации?\n\nФайл скрипта не будет удалён.".format(
                self.script_display_names.get(script_name, script_name)
            ),
            "Удалить скрипт",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question
        )
        if result == System.Windows.MessageBoxResult.Yes:
            if script_name in self.custom_script_checks:
                del self.custom_script_checks[script_name]
            if script_name in self.custom_script_paths:
                del self.custom_script_paths[script_name]
            if script_name in self.script_descriptions:
                del self.script_descriptions[script_name]
            if script_name in self.script_has_settings:
                del self.script_has_settings[script_name]
            if script_name in self.script_settings:
                del self.script_settings[script_name]
            if script_name in self.script_display_names:
                del self.script_display_names[script_name]
            if script_name in self.script_ids:
                del self.script_ids[script_name]
            if script_name in self.script_missing:
                del self.script_missing[script_name]
            if script_name in self.script_order:
                self.script_order.remove(script_name)

            to_remove = []
            for child in self.scripts_panel.Children:
                if hasattr(child, 'Tag') and child.Tag == script_name:
                    to_remove.append(child)
            for item in to_remove:
                self.scripts_panel.Children.Remove(item)

            if len(self.custom_script_checks) == 0:
                self._add_no_scripts_message()

            if self.selected_script_name == script_name:
                self.selected_script_name = None
                self.selected_script_row = None
                self._hide_description()

            self._auto_save_current_config()
            logger.info("Скрипт '{}' удалён из текущей конфигурации".format(script_name))

    # ================================================================
    # НАСТРОЙКИ СКРИПТА
    # ================================================================

    @staticmethod
    def _format_error(ex):
        """Вернуть подробное сообщение об ошибке с traceback."""
        tb_lines = traceback.format_exc().splitlines()
        return u"{}\n\n---\n{}".format(str(ex), u"\n".join(tb_lines))

    def on_script_settings(self, sender, e):
        """Открыть окно настроек скрипта."""
        script_name = sender.Tag
        if not script_name:
            return

        script_path = self.custom_script_paths.get(script_name)
        if not script_path or not os.path.exists(script_path):
            System.Windows.MessageBox.Show(
                "Файл скрипта не найден.",
                "Ошибка",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning
            )
            return

        # Загружаем модуль чтобы получить SETTINGS_SCHEMA
        try:
            module_name = "_settings_{}".format(script_name)
            if module_name in sys.modules:
                del sys.modules[module_name]

            script_dir = os.path.dirname(script_path)
            if script_dir not in sys.path:
                sys.path.insert(0, script_dir)
            if bundle_path not in sys.path:
                sys.path.insert(0, bundle_path)

            base = os.path.splitext(os.path.basename(script_path))[0]
            mod_file, mod_path, mod_desc = imp.find_module(base, [script_dir])
            module = imp.load_module(module_name, mod_file, mod_path, mod_desc)
            if mod_file:
                mod_file.close()
        except Exception as ex:
            System.Windows.MessageBox.Show(
                u"Ошибка загрузки скрипта: {}".format(self._format_error(ex)),
                "Ошибка",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error
            )
            return

        # Проверяем: скрипт может определить своё окно настроек
        show_fn = getattr(module, "show_settings", None)
        if callable(show_fn):
            current = self.script_settings.get(script_name, {})
            try:
                new_settings = show_fn(revit.doc, current)
            except Exception as ex:
                System.Windows.MessageBox.Show(
                    u"Ошибка окна настроек: {}".format(self._format_error(ex)),
                    u"Ошибка",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error
                )
                return
        else:
            # Generic-окно по SETTINGS_SCHEMA
            schema = getattr(module, "SETTINGS_SCHEMA", [])
            if not schema:
                System.Windows.MessageBox.Show(
                    u"У скрипта нет настраиваемых параметров.",
                    u"Информация",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information
                )
                return

            current = self.script_settings.get(script_name, {})

            from ui.script_settings_window import ScriptSettingsWindow
            dlg = ScriptSettingsWindow(
                self.script_display_names.get(script_name, script_name),
                schema, current, revit.doc
            )
            new_settings = dlg.show_dialog()

        if new_settings is not None:
            self.script_settings[script_name] = new_settings

            # Автосохранение настроек в отдельный файл скрипта
            section = self.get_selected_combo_value(self.cmb_section)
            project = self.get_selected_combo_value(self.cmb_project)
            if section and project:
                save_script_settings(script_name, section, project, new_settings)
                logger.info("Настройки скрипта '{}' автосохранены".format(script_name))

    # ================================================================
    # ЗАПУСК ВАЛИДАЦИИ
    # ================================================================

    def on_run_validation(self, sender, e):
        """Запустить валидацию."""
        section = self.get_selected_combo_value(self.cmb_section)
        project = self.get_selected_combo_value(self.cmb_project)

        selected_scripts = []
        for sname in self.script_order:
            cb = self.custom_script_checks.get(sname)
            if cb is None:
                continue
            if cb.IsChecked and not self.script_missing.get(sname, False):
                path = self.custom_script_paths.get(sname, "")
                settings = self.script_settings.get(sname, {})
                selected_scripts.append({
                    "name": sname,
                    "path": path,
                    "settings": settings
                })

        if not selected_scripts:
            System.Windows.MessageBox.Show(
                "Не выбрано ни одной проверки.\n\nДобавьте и отметьте скрипты.",
                "ВОР Валидатор",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning
            )
            return

        logger.info(u"Запуск валидации: раздел='{}', проект='{}'".format(section, project))
        logger.info(u"Скрипты: {}".format([s["name"] for s in selected_scripts]))

        # Главное окно остаётся открытым; окно прогона (modeless) открывается поверх.
        from ui.run_window import RunWindow
        RunWindow(scripts=selected_scripts, section=section, project=project,
                  mode="run", owner=self.window)

    def on_open_last_report(self, sender, e):
        """Открыть окно просмотра последнего отчёта (без выполнения скриптов)."""
        report = load_report()
        if not report:
            System.Windows.MessageBox.Show(
                u"Отчётов пока нет.\n\nЗапустите проверку, чтобы получить отчёт.",
                u"ВОР Валидатор",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information
            )
            return

        section = report.get("section", "")
        project = report.get("project", "")

        from ui.run_window import RunWindow
        RunWindow(scripts=None, section=section, project=project,
                  mode="view", report=report, owner=self.window)

    # ================================================================
    # УТИЛИТЫ
    # ================================================================

    def get_selected_combo_value(self, combo):
        """Получить выбранное значение из ComboBox."""
        if combo.SelectedItem and combo.SelectedItem.IsEnabled:
            return combo.SelectedItem.Content
        return ""

    def set_combo_by_text(self, combo, text):
        """Выбрать элемент ComboBox по тексту."""
        for item in combo.Items:
            if isinstance(item, ComboBoxItem) and item.Content == text:
                combo.SelectedItem = item
                return
