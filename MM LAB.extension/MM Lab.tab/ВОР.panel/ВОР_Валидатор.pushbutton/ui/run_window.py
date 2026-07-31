# -*- coding: utf-8 -*-
"""
Окно прогона проверки ВОР Валидатора.

Два режима:
  - mode="run"  : живой прогон выбранных скриптов с прогресс-баром
                  по каждому, кнопкой Стоп и кнопкой «Открыть результат».
  - mode="view" : холодный просмотр последнего отчёта (загружен из
                  report_last.json) без выполнения скриптов.

Архитектурное ограничение (важно):
  Revit API нельзя вызывать из фонового потока, поэтому реальная
  многопоточность невозможна. Живой прогресс реализован через yield
  UI-потоку между скриптами: после каждой строки вызываем
  Dispatcher.Invoke(Background), чтобы WPF успел перерисовать
  ProgressBar и обработать клик Стопа. Внутри одного run() прогресс
  не двигается — это физически невозможно. Стоп срабатывает только
  между скриптами (после завершения текущего).
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
from System.Windows import (
    Window, Thickness, FontWeights, FontStyles,
    HorizontalAlignment, VerticalAlignment,
    Visibility as WVisibility
)
from System.Windows.Controls import (
    StackPanel, DockPanel, Grid, ColumnDefinition, RowDefinition,
    GridLength, GridUnitType, TextBlock, Button, Border, CheckBox,
    ScrollViewer, ScrollBarVisibility, Orientation, ProgressBar
)
from System.Windows.Media import SolidColorBrush, Colors, Color, Stretch
from System.Windows.Shapes import Path as WpfPath
from System.Windows.Markup import XamlReader
from System.Windows.Threading import DispatcherPriority

from pyrevit import revit, script, DB

bundle_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if bundle_path not in sys.path:
    sys.path.insert(0, bundle_path)

from core.validation_engine import run_custom_script
from core.report_store import save_report, load_report

logger = script.get_logger()

# ── Design System ──────────────────────────────────────────────
_BG_WINDOW = Color.FromArgb(255, 245, 245, 245)
_BG_CARD = Colors.White
_BRUSH_BORDER = SolidColorBrush(Color.FromArgb(255, 204, 204, 204))

_FG_TITLE = SolidColorBrush(Color.FromArgb(255, 51, 51, 51))
_FG_SUBTITLE = SolidColorBrush(Color.FromArgb(255, 85, 85, 85))
_FG_BODY = SolidColorBrush(Color.FromArgb(255, 68, 68, 68))
_FG_MUTED = SolidColorBrush(Colors.Gray)

_FG_SUCCESS = SolidColorBrush(Color.FromArgb(255, 56, 142, 60))
_FG_ERROR = SolidColorBrush(Color.FromArgb(255, 211, 47, 47))
_FG_WARNING = SolidColorBrush(Color.FromArgb(255, 255, 152, 0))
_FG_INFO = SolidColorBrush(Color.FromArgb(255, 33, 150, 243))

_BG_SUCCESS_LIGHT = SolidColorBrush(Color.FromArgb(40, 56, 142, 60))
_BG_ERROR_LIGHT = SolidColorBrush(Color.FromArgb(40, 211, 47, 47))
_BG_RUNNING_LIGHT = SolidColorBrush(Color.FromArgb(30, 33, 150, 243))
_BG_SKIPPED_LIGHT = SolidColorBrush(Color.FromArgb(20, 128, 128, 128))

# Иконки статусов (Path geometry, stroke-only)
_CHECK_PATH_DATA = "M4,9 L7.5,12.5 L14,5"        # галочка ✓
_CROSS_PATH_DATA = "M5,5 L13,13 M13,5 L5,13"    # крестик ✗
_WARN_PATH_DATA = "M9,2 L16,14 L2,14 Z M9,7 L9,10 M9,12 L9,12.5"  # треугольник ⚠
_INFO_PATH_DATA = "M9,2 A7,7,0,1,0,9,16 A7,7,0,1,0,9,2 M9,7 L9,7.5 M9,9 L9,13"

_ROUNDED_ICON_TEMPLATE = XamlReader.Parse(
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="Button">'
    '<Border x:Name="Bd" CornerRadius="3" Background="#F0F0F0" '
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

_PRIMARY_BUTTON_TEMPLATE = XamlReader.Parse(
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="Button">'
    '<Border x:Name="Bd" Background="#007ACC" CornerRadius="4" Padding="10,5">'
    '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
    '</Border>'
    '<ControlTemplate.Triggers>'
    '<Trigger Property="IsMouseOver" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#005A9E"/>'
    '</Trigger>'
    '<Trigger Property="IsPressed" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#004578"/>'
    '</Trigger>'
    '</ControlTemplate.Triggers>'
    '</ControlTemplate>'
)

_SECONDARY_BUTTON_TEMPLATE = XamlReader.Parse(
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="Button">'
    '<Border x:Name="Bd" Background="#E0E0E0" BorderBrush="#CCCCCC" '
    'BorderThickness="1" CornerRadius="3" Padding="10,5">'
    '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
    '</Border>'
    '<ControlTemplate.Triggers>'
    '<Trigger Property="IsMouseOver" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#D0D0D0"/>'
    '</Trigger>'
    '<Trigger Property="IsPressed" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#C0C0C0"/>'
    '</Trigger>'
    '</ControlTemplate.Triggers>'
    '</ControlTemplate>'
)


def _make_icon(path_data, brush, size=16, stroke=1.6):
    """Создать stroke-only иконку (Path)."""
    icon = WpfPath()
    icon.Data = System.Windows.Media.Geometry.Parse(path_data)
    icon.Stroke = brush
    icon.StrokeThickness = stroke
    icon.Stretch = Stretch.Uniform
    icon.Width = size
    icon.Height = size
    icon.HorizontalAlignment = HorizontalAlignment.Center
    icon.VerticalAlignment = VerticalAlignment.Center
    return icon


def _make_icon_button(content, tooltip, handler, template=None):
    """Создать текстовую кнопку-действие."""
    btn = Button()
    btn.Content = content
    btn.ToolTip = tooltip
    btn.Height = 28
    btn.MinWidth = 70
    btn.VerticalAlignment = VerticalAlignment.Center
    btn.Template = template if template is not None else _SECONDARY_BUTTON_TEMPLATE
    btn.FontSize = 12
    btn.Padding = Thickness(8, 0, 8, 0)
    if handler:
        btn.Click += handler
    return btn


# ── Строка одного скрипта ──────────────────────────────────────

class _ScriptRow(object):
    """Визуальная строка + состояние одного скрипта в окне прогона."""

    def __init__(self, parent, script_info):
        self.parent = parent
        self.script_info = script_info
        self.name = script_info.get("name", "")
        self.result = None          # ValidationResult после выполнения
        self.executed = False       # был ли выполнен (не пропущен стопом)

        self._build()

    def _build(self):
        # Корневой Border (карточка строки)
        self.border = Border()
        self.border.Background = SolidColorBrush(_BG_CARD)
        self.border.BorderBrush = _BRUSH_BORDER
        self.border.BorderThickness = Thickness(0, 0, 0, 1)
        self.border.Padding = Thickness(10, 8, 10, 8)
        self.border.Margin = Thickness(0, 0, 0, 4)

        # Внутренний Grid: [иконка][имя+сообщение ......][действие]
        grid = Grid()
        col_icon = ColumnDefinition()
        col_icon.Width = GridLength(1, GridUnitType.Auto)
        col_main = ColumnDefinition()
        col_main.Width = GridLength(1, GridUnitType.Star)
        col_act = ColumnDefinition()
        col_act.Width = GridLength(1, GridUnitType.Auto)
        grid.ColumnDefinitions.Add(col_icon)
        grid.ColumnDefinitions.Add(col_main)
        grid.ColumnDefinitions.Add(col_act)

        # Иконка статуса (колонка 0)
        self.icon_holder = TextBlock()
        self.icon_holder.Width = 22
        self.icon_holder.Height = 22
        self.icon_holder.VerticalAlignment = VerticalAlignment.Center
        self.icon_holder.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(self.icon_holder, 0)
        # Изначально — нейтральная точка (ожидание)
        self.icon_holder.Content = ""
        grid.Children.Add(self.icon_holder)

        # Центр: имя + сообщение + прогресс (колонка 1)
        main = StackPanel()
        main.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(main, 1)

        self.txt_name = TextBlock()
        self.txt_name.Text = self.name
        self.txt_name.FontSize = 13
        self.txt_name.FontWeight = FontWeights.SemiBold
        self.txt_name.Foreground = _FG_TITLE
        main.Children.Add(self.txt_name)

        self.txt_message = TextBlock()
        self.txt_message.Text = ""          # сообщение появится после выполнения
        self.txt_message.FontSize = 11
        self.txt_message.Foreground = _FG_MUTED
        self.txt_message.TextWrapping = System.Windows.TextWrapping.Wrap
        self.txt_message.Margin = Thickness(0, 2, 0, 0)
        self.txt_message.Visibility = WVisibility.Collapsed
        main.Children.Add(self.txt_message)

        self.progress = ProgressBar()
        self.progress.Height = 4
        self.progress.IsIndeterminate = True
        self.progress.Visibility = WVisibility.Collapsed
        self.progress.Margin = Thickness(0, 4, 12, 0)
        main.Children.Add(self.progress)

        grid.Children.Add(main)

        # Кнопка действия (колонка 2)
        self.action_panel = StackPanel()
        self.action_panel.Orientation = Orientation.Horizontal
        self.action_panel.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(self.action_panel, 2)
        grid.Children.Add(self.action_panel)

        self.border.Child = grid

    # ── Состояния ──

    def set_pending(self):
        """Ожидание: ещё не выполнялся."""
        self._set_icon(None)
        self.txt_message.Visibility = WVisibility.Collapsed
        self.progress.Visibility = WVisibility.Collapsed
        self.border.Background = SolidColorBrush(_BG_CARD)

    def set_running(self):
        """Выполняется сейчас."""
        self._set_icon(None)
        self.txt_message.Text = u"\u0412\u044b\u043f\u043e\u043b\u043d\u044f\u0435\u0442\u0441\u044f..."  # "Выполняется..."
        self.txt_message.Foreground = _FG_INFO
        self.txt_message.Visibility = WVisibility.Visible
        self.progress.Visibility = WVisibility.Visible
        self.border.Background = _BG_RUNNING_LIGHT

    def set_done(self, result):
        """Завершён с результатом ValidationResult."""
        self.result = result
        self.executed = True
        self.progress.Visibility = WVisibility.Collapsed

        passed = bool(result.passed)
        msg = result.message or ""

        # Сообщение: N проблем + длительность
        elem_count = len(result.elements) if result.elements else 0
        dur = getattr(result, "duration_ms", 0) or 0
        dur_s = u"{:.1f} \u0441".format(dur / 1000.0) if dur >= 1000 else u"{} \u043c\u0441".format(dur)

        if not passed:
            # Ошибка выполнения vs реальный FAIL
            is_error = msg.lower().startswith(u"\u043e\u0448\u0438\u0431\u043a\u0430")  # "ошибка"
            if is_error:
                self._set_icon(_make_icon(_WARN_PATH_DATA, _FG_WARNING))
                self.border.Background = _BG_SKIPPED_LIGHT
                detail = u"\u041e\u0448\u0438\u0431\u043a\u0430 \u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u044f"  # "Ошибка выполнения"
            else:
                self._set_icon(_make_icon(_CROSS_PATH_DATA, _FG_ERROR))
                self.border.Background = _BG_ERROR_LIGHT
                detail = u"\u041f\u0440\u043e\u0431\u043b\u0435\u043c: {}".format(elem_count)  # "Проблем: N"
            full_msg = u"{}\n{}\n({})".format(detail, msg, dur_s)
            self.txt_message.Text = full_msg
            self.txt_message.Foreground = _FG_ERROR if not is_error else _FG_WARNING
        else:
            self._set_icon(_make_icon(_CHECK_PATH_DATA, _FG_SUCCESS))
            self.border.Background = _BG_SUCCESS_LIGHT
            self.txt_message.Text = u"OK \u2014 {}\n({})".format(msg if msg else u"\u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u043f\u0440\u043e\u0439\u0434\u0435\u043d\u0430", dur_s)
            self.txt_message.Foreground = _FG_SUCCESS

        self.txt_message.Visibility = WVisibility.Visible
        self._refresh_action_button()

    def set_skipped(self):
        """Пропущен из-за останова."""
        self.executed = False
        self.progress.Visibility = WVisibility.Collapsed
        self._set_icon(None)
        self.txt_message.Text = u"\u041d\u0435 \u0432\u044b\u043f\u043e\u043b\u043d\u044f\u043b\u0441\u044f (\u043e\u0441\u0442\u0430\u043d\u043e\u0432)"  # "Не выполнялся (останов)"
        self.txt_message.Foreground = _FG_MUTED
        self.txt_message.Visibility = WVisibility.Visible
        self.border.Background = _BG_SKIPPED_LIGHT
        self._clear_action_button()

    def set_view_mode(self, script_dict):
        """Холодный режим просмотра из отчёта (script_dict из report_last.json)."""
        self.result = None  # в режиме view нет живого объекта ValidationResult
        self._script_dict = script_dict
        self.executed = True

        status = script_dict.get("status", "passed")
        msg = script_dict.get("message", "")
        elem_count = script_dict.get("element_count", 0)
        dur = script_dict.get("duration_ms", 0) or 0
        dur_s = u"{:.1f} \u0441".format(dur / 1000.0) if dur >= 1000 else u"{} \u043c\u0441".format(dur)
        has_window = script_dict.get("has_results_window", False)

        if status == "passed":
            self._set_icon(_make_icon(_CHECK_PATH_DATA, _FG_SUCCESS))
            self.border.Background = _BG_SUCCESS_LIGHT
            detail = u"OK"
            self.txt_message.Foreground = _FG_SUCCESS
        elif status == "error":
            self._set_icon(_make_icon(_WARN_PATH_DATA, _FG_WARNING))
            self.border.Background = _BG_SKIPPED_LIGHT
            detail = u"\u041e\u0448\u0438\u0431\u043a\u0430 \u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u044f"
            self.txt_message.Foreground = _FG_WARNING
        else:  # failed
            self._set_icon(_make_icon(_CROSS_PATH_DATA, _FG_ERROR))
            self.border.Background = _BG_ERROR_LIGHT
            detail = u"\u041f\u0440\u043e\u0431\u043b\u0435\u043c: {}".format(elem_count)
            self.txt_message.Foreground = _FG_ERROR

        self.txt_message.Text = u"{}\n{}\n({})".format(detail, msg, dur_s)
        self.txt_message.Visibility = WVisibility.Visible
        self._refresh_action_button_view(elem_count, has_window)

    # ── Кнопка действия ──

    def _clear_action_button(self):
        self.action_panel.Children.Clear()

    def _set_icon(self, icon):
        """Заменить контент иконки статуса (или скрыть)."""
        self.icon_holder.Content = icon if icon is not None else ""

    def _refresh_action_button(self):
        """В режиме run: показать «Открыть результат», если есть что показать."""
        self._clear_action_button()
        if self.result is None:
            return
        # Есть что показывать, если: есть элементы ИЛИ скрипт упал (сообщение)
        has_elements = bool(self.result.elements)
        has_window = bool(getattr(self.result, "has_results_window", False))
        if has_elements or has_window:
            btn = _make_icon_button(
                u"\u041e\u0442\u043a\u0440\u044b\u0442\u044c",  # "Открыть"
                u"\u041e\u0442\u043a\u0440\u044b\u0442\u044c \u0440\u0435\u0437\u0443\u043b\u044c\u0442\u0430\u0442",
                lambda s, e, si=self.script_info: self.parent._on_open_result_run(si),
                template=_SECONDARY_BUTTON_TEMPLATE,
            )
            self.action_panel.Children.Add(btn)

    def _refresh_action_button_view(self, elem_count, has_window):
        """В режиме view: показать «Открыть результат» из сохранённых элементов."""
        self._clear_action_button()
        if elem_count > 0 or has_window:
            btn = _make_icon_button(
                u"\u041e\u0442\u043a\u0440\u044b\u0442\u044c",
                u"\u041e\u0442\u043a\u0440\u044b\u0442\u044c \u0440\u0435\u0437\u0443\u043b\u044c\u0442\u0430\u0442",
                lambda s, e, sd=self._script_dict: self.parent._on_open_result_view(sd),
                template=_SECONDARY_BUTTON_TEMPLATE,
            )
            self.action_panel.Children.Add(btn)


# ── Окно прогона ───────────────────────────────────────────────

class RunWindow(Window):

    def __init__(self, scripts=None, section="", project="", mode="run", report=None,
                 owner=None):
        """
        Args:
            scripts: list[dict] — [{name, path, settings, id?}] (режим run)
            section, project: str — текущие раздел/проект
            mode: "run" | "view"
            report: dict — загруженный отчёт (режим view); если None и mode=view,
                           грузится из report_last.json
            owner: родительское окно (для CenterOwner)
        """
        xaml_path = os.path.join(os.path.dirname(__file__), "run_window.xaml")
        with codecs.open(xaml_path, "r", "utf-8") as f:
            xaml_content = f.read()
        window = XamlReader.Parse(xaml_content)

        self.window = window
        self.txt_title = window.FindName("TxtTitle")
        self.txt_subtitle = window.FindName("TxtSubtitle")
        self.scripts_panel = window.FindName("ScriptsPanel")
        self.btn_stop = window.FindName("BtnStop")
        self.btn_close = window.FindName("BtnClose")

        self.mode = mode
        self.section = section
        self.project = project
        self.scripts = scripts or []
        self.rows = []
        self.results = []
        self._run_index = 0
        self._stop_requested = False

        self.btn_close.Click += self._on_close
        self.btn_stop.Click += self._on_stop

        if owner is not None:
            try:
                window.Owner = owner
            except Exception:
                pass

        if mode == "view":
            self._init_view(report)
        else:
            self._init_run()

        window.Show()

    # ── Режим RUN ──

    def _init_run(self):
        self.txt_title.Text = u"\u041f\u0440\u043e\u0433\u043e\u043d \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438"
        self.txt_subtitle.Text = u"{} | {}".format(self.section, self.project)
        self.btn_stop.Visibility = WVisibility.Visible

        for si in self.scripts:
            row = _ScriptRow(self, si)
            row.set_pending()
            self.rows.append(row)
            self.scripts_panel.Children.Add(row.border)

        # Стартуем прогон через BeginInvoke (асинхронно), чтобы окно успело
        # отрисоваться (Show() выполнится после возврата из __init__).
        self.window.Dispatcher.BeginInvoke(
            DispatcherPriority.Background,
            System.Action(self._start_run)
        )

    def _start_run(self):
        self._run_index = 0
        self._schedule_next()

    def _schedule_next(self):
        """Асинхронно поставить следующий шаг в очередь Dispatcher.

        BeginInvoke возвращает управление немедленно в цикл сообщений WPF,
        поэтому UI успевает перерисоваться (прогресс) и обработать клик Стопа.
        """
        self.window.Dispatcher.BeginInvoke(
            DispatcherPriority.Background,
            System.Action(self._run_step)
        )

    def _run_step(self):
        # Проверка останова — между скриптами
        if self._stop_requested:
            self._on_run_complete(was_stopped=True)
            return

        if self._run_index >= len(self.rows):
            self._on_run_complete(was_stopped=False)
            return

        row = self.rows[self._run_index]
        row.set_running()
        # BeginInvoke: ставим _execute_current в очередь и сразу возвращаемся.
        # WPF успеет отрисовать ProgressBar ДО того, как начнётся тяжёлый run().
        self.window.Dispatcher.BeginInvoke(
            DispatcherPriority.Background,
            System.Action(self._execute_current)
        )

    def _execute_current(self):
        if self._stop_requested:
            self._on_run_complete(was_stopped=True)
            return

        row = self.rows[self._run_index]
        try:
            result = run_custom_script(row.script_info, self.section, self.project)
        except Exception as ex:
            from core.validation_engine import ValidationResult
            result = ValidationResult(
                check_name=row.name,
                passed=False,
                message=u"\u041e\u0448\u0438\u0431\u043a\u0430: {}".format(ex)
            )
            logger.error(u"RunWindow: {}".format(traceback.format_exc()))

        self.results.append(result)
        row.set_done(result)
        self._run_index += 1
        self._schedule_next()

    def _on_run_complete(self, was_stopped):
        self.btn_stop.Visibility = WVisibility.Collapsed

        if was_stopped:
            # Отметить невыполненные строки как пропущенные
            for i in range(self._run_index, len(self.rows)):
                self.rows[i].set_skipped()

        # Сохранить отчёт (только выполненные)
        done_results = [r.result for r in self.rows if r.executed and r.result is not None]
        total = len(self.rows)
        save_report(self.section, self.project, done_results,
                    was_stopped=was_stopped, scripts_total=total)

        # Итоговый подзаголовок
        passed = sum(1 for r in done_results if r.passed)
        failed = sum(1 for r in done_results if not r.passed)
        stopped_tag = u"  [\u043e\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043e]" if was_stopped else u""
        self.txt_subtitle.Text = u"{} | {}  \u2014  OK: {}, \u043f\u0440\u043e\u0431\u043b\u0435\u043c: {}{}".format(
            self.section, self.project, passed, failed, stopped_tag
        )

    def _on_stop(self, sender, e):
        self._stop_requested = True
        self.btn_stop.IsEnabled = False
        self.btn_stop.Content = u"\u041e\u0441\u0442\u0430\u043d\u0430\u0432\u043b\u0438\u0432\u0430\u0435\u043c..."

    # ── Режим VIEW ──

    def _init_view(self, report):
        if report is None:
            report = load_report()

        if not report:
            self.txt_title.Text = u"\u041e\u0442\u0447\u0451\u0442\u043e\u0432 \u043f\u043e\u043a\u0430 \u043d\u0435\u0442"
            self.txt_subtitle.Text = u"\u0417\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0443, \u0447\u0442\u043e\u0431\u044b \u043f\u043e\u043b\u0443\u0447\u0438\u0442\u044c \u043e\u0442\u0447\u0451\u0442"
            self.btn_stop.Visibility = WVisibility.Collapsed
            return

        ts = report.get("timestamp", "")
        was_stopped = report.get("was_stopped", False)
        self.section = report.get("section", "")
        self.project = report.get("project", "")

        self.txt_title.Text = u"\u041f\u043e\u0441\u043b\u0435\u0434\u043d\u0438\u0439 \u043e\u0442\u0447\u0451\u0442"
        stopped_tag = u"  [\u043e\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043e]" if was_stopped else u""
        self.txt_subtitle.Text = u"{} | {}  \u2014  \u043e\u0442 {}{}".format(
            self.section, self.project, ts, stopped_tag
        )
        self.btn_stop.Visibility = WVisibility.Collapsed

        for sd in report.get("scripts", []):
            row = _ScriptRow(self, {"name": sd.get("name", "")})
            row.set_view_mode(sd)
            self.rows.append(row)
            self.scripts_panel.Children.Add(row.border)

    # ── Открытие результата ──

    def _load_script_module(self, script_info):
        """Вернуть модуль скрипта для show_results.

        ВАЖНО: модуль НЕ перезагружается, если уже загружен в этой сессии —
        иначе module-level _last_results_data (кэш данных прогона) сбросится
        в None и show_results не сможет открыть окно. Берём из sys.modules.
        """
        script_path = script_info.get("path")
        if not script_path or not os.path.exists(script_path):
            return None
        try:
            module_name = "custom_{}".format(os.path.splitext(script_info.get("name", "x"))[0])
            # Переиспользуем уже загруженный модуль — в нём живой _last_results_data
            if module_name in sys.modules:
                return sys.modules[module_name]
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
            return module
        except Exception:
            logger.error(u"RunWindow load module: {}".format(traceback.format_exc()))
            return None

    def _on_open_result_run(self, script_info):
        """Открыть результат в режиме run: живой объект результата есть."""
        # Найти результат для этого скрипта
        result = None
        for r in self.results:
            if r.check_name == script_info.get("name"):
                result = r
                break
        if result is None:
            return

        # Сначала пробуем show_results (кастомное окно скрипта)
        module = self._load_script_module(script_info)
        if module is not None and hasattr(module, "show_results"):
            try:
                settings = script_info.get("settings", {})
                module.show_results(revit.doc, self.section, self.project, settings)
                return
            except Exception:
                logger.error(u"show_results: {}".format(traceback.format_exc()))

        # Иначе — стандартный results-шаблон по элементам
        self._show_standard_results(result.elements, script_info.get("name", ""))

    def _on_open_result_view(self, script_dict):
        """Открыть результат в режиме view: данных в памяти нет, только отчёт."""
        name = script_dict.get("name", "")
        element_ids = script_dict.get("element_ids", [])
        # Восстанавливаем ElementId из int
        ids = []
        for eid_int in element_ids:
            try:
                ids.append(DB.ElementId(int(eid_int)))
            except Exception:
                pass
        self._show_standard_results(ids, name)

    def _show_standard_results(self, element_ids, script_name):
        """Показать стандартное results-окно по списку ElementId."""
        if not element_ids:
            System.Windows.MessageBox.Show(
                u"\u041d\u0435\u0442 \u044d\u043b\u0435\u043c\u0435\u043d\u0442\u043e\u0432 \u0434\u043b\u044f \u043f\u0440\u043e\u0441\u043c\u043e\u0442\u0440\u0430.",
                script_name,
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information
            )
            return

        doc = revit.doc
        # Группируем как одну категорию "Проблемы (N)"
        elements = []
        for eid in element_ids:
            try:
                el = doc.GetElement(eid)
                if el is None:
                    continue
                elements.append({"id": eid, "name": self._element_label(el)})
            except Exception:
                pass

        results = {u"\u041f\u0440\u043e\u0431\u043b\u0435\u043c\u044b ({})".format(len(elements)): elements}
        self._open_results_window(doc, results, title=script_name)

    @staticmethod
    def _element_label(el):
        """Человекочитаемая подпись элемента для списка."""
        try:
            cat = el.Category
            cat_name = cat.Name if cat is not None else u"\u042d\u043b\u0435\u043c\u0435\u043d\u0442"
        except Exception:
            cat_name = u"\u042d\u043b\u0435\u043c\u0435\u043d\u0442"
        try:
            type_name = ""
            t = el.Document.GetElement(el.GetTypeId())
            if t is not None:
                type_name = t.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                type_name = type_name.AsString() if type_name else t.Name
            elif hasattr(el, "Name"):
                type_name = el.Name
        except Exception:
            type_name = ""
        if type_name:
            return u"{}: {}".format(cat_name, type_name)
        return cat_name

    def _open_results_window(self, doc, results, title=""):
        """Открыть стандартное ResultsWindow (modeless singleton через __main__)."""
        try:
            from ui.standard_results_window import StandardResultsWindow
            import __main__
            attr = "_vor_validator_results_window"
            old = getattr(__main__, attr, None)
            if old:
                try:
                    old.Close()
                except Exception:
                    pass
            win = StandardResultsWindow(doc, results, title=title or u"\u0420\u0435\u0437\u0443\u043b\u044c\u0442\u0430\u0442\u044b")
            setattr(__main__, attr, win)
            win.Show()
        except Exception:
            logger.error(u"_open_results_window: {}".format(traceback.format_exc()))

    # ── Закрытие ──

    def _on_close(self, sender, e):
        self.window.Close()
