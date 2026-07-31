# -*- coding: utf-8 -*-
"""
Окно настроек для скрипта проверки заполненности GP_01_Коды.
Две группы категорий: семейства и материалы.
"""

import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import (
    Window, Thickness,
    FontWeights, HorizontalAlignment, VerticalAlignment
)
from System.Windows.Controls import (
    StackPanel, TextBlock, CheckBox, Button, Border, ScrollViewer,
    ScrollBarVisibility, Separator
)
from System.Windows.Media import SolidColorBrush, Colors

from pyrevit import DB

PARAM_FAMILY_CODE_UNIT = u"01_GP_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b"
PARAM_FAMILY_WORK_TYPE = u"01_GP_\u041a\u043e\u0434\u0412\u0438\u0434\u0430\u0420\u0430\u0431\u043e\u0442\u044b_\u0422\u0438\u043f"
PARAM_MATERIAL_CODE = u"01_GP_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b_\u041c\u0430\u0442"

_GRAY = SolidColorBrush(Colors.Gray)
_WHITE = SolidColorBrush(Colors.White)
_BORDER_GRAY = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 200, 200, 200))
_HEADER_FAMILY = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 33, 150, 243))
_HEADER_MATERIAL = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 152, 0))


def _detect_categories(doc):
    """Найти категории с нужными параметрами."""
    result = {"family": [], "material": []}
    family_bics = set()
    material_bics = set()

    target_params = {
        PARAM_FAMILY_CODE_UNIT: "family",
        PARAM_FAMILY_WORK_TYPE: "family",
        PARAM_MATERIAL_CODE: "material",
    }

    try:
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            definition = iterator.Key
            group = target_params.get(definition.Name)
            if group is None:
                continue
            binding = iterator.Current
            cat_set = binding.Categories
            cat_iter = cat_set.GetEnumerator()
            while cat_iter.MoveNext():
                cat = cat_iter.Current
                bic_int = int(cat.BuiltInCategory)
                if group == "family" and bic_int not in family_bics:
                    family_bics.add(bic_int)
                    result["family"].append((cat.Name, bic_int))
                elif group == "material" and bic_int not in material_bics:
                    material_bics.add(bic_int)
                    result["material"].append((cat.Name, bic_int))
    except Exception:
        pass

    result["family"].sort(key=lambda x: x[0])
    result["material"].sort(key=lambda x: x[0])
    return result


class KodySettingsWindow(Window):

    def __init__(self, doc, current_settings):
        self.Title = u"\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438: \u0417\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u043e\u0441\u0442\u044c GP_01_\u041a\u043e\u0434\u044b"
        self.Width = 450
        self.Height = 540
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.Background = _WHITE
        self.doc = doc
        self.result_settings = None

        self._category_checks = {}
        saved_categories = (current_settings or {}).get("categories", {})
        detected = _detect_categories(doc)

        # ==== UI ====
        root = StackPanel()
        root.Margin = Thickness(16)

        header = TextBlock()
        header.Text = u"\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u0437\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u043e\u0441\u0442\u0438 GP_01_\u041a\u043e\u0434\u044b"
        header.FontSize = 14
        header.FontWeight = FontWeights.Bold
        header.Margin = Thickness(0, 0, 0, 4)
        root.Children.Add(header)

        hint = TextBlock()
        hint.Text = u"\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438 \u0434\u043b\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438"
        hint.FontSize = 11
        hint.FontStyle = System.Windows.FontStyles.Italic
        hint.Foreground = _GRAY
        hint.Margin = Thickness(0, 0, 0, 8)
        root.Children.Add(hint)

        # Кнопки Выбрать все / Снять все
        toggle_panel = StackPanel()
        toggle_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        toggle_panel.Margin = Thickness(0, 0, 0, 8)

        btn_select_all = Button()
        btn_select_all.Content = u"\u0412\u044b\u0431\u0440\u0430\u0442\u044c \u0432\u0441\u0435"
        btn_select_all.FontSize = 11
        btn_select_all.Height = 24
        btn_select_all.Width = 100
        btn_select_all.Margin = Thickness(0, 0, 8, 0)
        btn_select_all.Click += self._on_select_all
        toggle_panel.Children.Add(btn_select_all)

        btn_deselect_all = Button()
        btn_deselect_all.Content = u"\u0421\u043d\u044f\u0442\u044c \u0432\u0441\u0435"
        btn_deselect_all.FontSize = 11
        btn_deselect_all.Height = 24
        btn_deselect_all.Width = 100
        btn_deselect_all.Click += self._on_deselect_all
        toggle_panel.Children.Add(btn_deselect_all)

        root.Children.Add(toggle_panel)

        # === Семейства ===
        family_header = TextBlock()
        family_header.Text = u"\u041a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438 \u0441\u0435\u043c\u0435\u0439\u0441\u0442\u0432 (01_GP_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b + 01_GP_\u041a\u043e\u0434\u0412\u0438\u0434\u0430\u0420\u0430\u0431\u043e\u0442\u044b_\u0422\u0438\u043f)"
        family_header.FontSize = 12
        family_header.FontWeight = FontWeights.SemiBold
        family_header.Foreground = _HEADER_FAMILY
        family_header.Margin = Thickness(0, 4, 0, 4)
        root.Children.Add(family_header)

        family_scroll = ScrollViewer()
        family_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        family_scroll.MaxHeight = 150
        family_panel = StackPanel()

        if not detected["family"]:
            empty_msg = TextBlock()
            empty_msg.Text = u"\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u044b \u043d\u0435 \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u043d\u044b \u043d\u0438 \u043a \u043e\u0434\u043d\u043e\u0439 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438."
            empty_msg.FontSize = 12
            empty_msg.Foreground = _GRAY
            empty_msg.Margin = Thickness(0, 4, 0, 4)
            family_panel.Children.Add(empty_msg)

        for cat_name, bic in detected["family"]:
            cb = CheckBox()
            cb.Content = cat_name
            cb.FontSize = 13
            cb.Margin = Thickness(0, 4, 0, 0)
            cb.IsChecked = saved_categories.get(cat_name, True)
            family_panel.Children.Add(cb)
            self._category_checks[cat_name] = cb

        family_scroll.Content = family_panel
        root.Children.Add(family_scroll)

        # Разделитель
        sep = Separator()
        sep.Margin = Thickness(0, 8, 0, 8)
        root.Children.Add(sep)

        # === Материалы ===
        mat_header = TextBlock()
        mat_header.Text = u"\u041a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438 \u043c\u0430\u0442\u0435\u0440\u0438\u0430\u043b\u043e\u0432 (01_GP_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b_\u041c\u0430\u0442)"
        mat_header.FontSize = 12
        mat_header.FontWeight = FontWeights.SemiBold
        mat_header.Foreground = _HEADER_MATERIAL
        mat_header.Margin = Thickness(0, 0, 0, 4)
        root.Children.Add(mat_header)

        mat_scroll = ScrollViewer()
        mat_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        mat_scroll.MaxHeight = 150
        mat_panel = StackPanel()

        if not detected["material"]:
            empty_msg2 = TextBlock()
            empty_msg2.Text = u"\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 \u043d\u0435 \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u043d \u043d\u0438 \u043a \u043e\u0434\u043d\u043e\u0439 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438."
            empty_msg2.FontSize = 12
            empty_msg2.Foreground = _GRAY
            empty_msg2.Margin = Thickness(0, 4, 0, 4)
            mat_panel.Children.Add(empty_msg2)

        for cat_name, bic in detected["material"]:
            cb = CheckBox()
            cb.Content = cat_name
            cb.FontSize = 13
            cb.Margin = Thickness(0, 4, 0, 0)
            cb.IsChecked = saved_categories.get(cat_name, True)
            mat_panel.Children.Add(cb)
            self._category_checks[cat_name] = cb

        mat_scroll.Content = mat_panel
        root.Children.Add(mat_scroll)

        # Кнопки
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 12, 0, 0)

        btn_ok = Button()
        btn_ok.Content = u"\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c"
        btn_ok.Width = 100
        btn_ok.Height = 30
        btn_ok.Margin = Thickness(0, 0, 8, 0)
        btn_ok.IsDefault = True
        btn_ok.Click += self._on_ok
        btn_panel.Children.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Content = u"\u041e\u0442\u043c\u0435\u043d\u0430"
        btn_cancel.Width = 100
        btn_cancel.Height = 30
        btn_cancel.IsCancel = True
        btn_panel.Children.Add(btn_cancel)

        root.Children.Add(btn_panel)
        self.Content = root

    def _on_select_all(self, sender, e):
        for cb in self._category_checks.values():
            cb.IsChecked = True

    def _on_deselect_all(self, sender, e):
        for cb in self._category_checks.values():
            cb.IsChecked = False

    def _on_ok(self, sender, e):
        categories = {}
        for cat_name, cb in self._category_checks.items():
            categories[cat_name] = bool(cb.IsChecked)

        self.result_settings = {"categories": categories}
        self.DialogResult = True

    def show_dialog(self):
        if self.ShowDialog() == True:
            return self.result_settings
        return None
