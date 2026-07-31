# -*- coding: utf-8 -*-
"""
Кастомное окно настроек для скрипта 'Проверка марки по шаблону'.
Демонстрирует паттерн show_settings() с собственным WPF-окном.
"""

import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import Window, Thickness
from System.Windows.Controls import *
from System.Windows.Media import SolidColorBrush, Colors


class CheckSettingsWindow(Window):
    """Окно выбора категории и шаблона марки."""

    def __init__(self, doc, current_settings):
        self.Title = u"\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438: \u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u043c\u0430\u0440\u043a\u0438 \u043f\u043e \u0448\u0430\u0431\u043b\u043e\u043d\u0443"
        self.Width = 400
        self.Height = 280
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.Background = SolidColorBrush(Colors.White)
        self.result_settings = None
        self.doc = doc

        # Текущие значения
        cur_category = current_settings.get("category", "")
        cur_pattern = current_settings.get("pattern", "")

        stack = StackPanel()
        stack.Margin = Thickness(16)

        # Категория
        lbl_cat = TextBlock()
        lbl_cat.Text = u"\u041a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044f \u044d\u043b\u0435\u043c\u0435\u043d\u0442\u043e\u0432:"
        lbl_cat.FontSize = 13
        lbl_cat.FontWeight = System.Windows.FontWeights.SemiBold
        lbl_cat.Margin = Thickness(0, 0, 0, 4)
        stack.Children.Add(lbl_cat)

        self._category_combo = ComboBox()
        self._category_combo.Height = 28
        self._category_combo.FontSize = 12
        self._category_combo.Margin = Thickness(0, 0, 0, 12)
        categories = [u"\u0421\u0442\u0435\u043d\u044b", u"\u0414\u0432\u0435\u0440\u0438", u"\u041e\u043a\u043d\u0430", u"\u041f\u0435\u0440\u0435\u043a\u0440\u044b\u0442\u0438\u044f", u"\u041a\u043e\u043b\u043e\u043d\u043d\u044b"]
        for cat in categories:
            self._category_combo.Items.Add(cat)
        if cur_category and cur_category in categories:
            self._category_combo.SelectedItem = cur_category
        else:
            self._category_combo.SelectedIndex = 0
        stack.Children.Add(self._category_combo)

        # Шаблон марки
        lbl_pat = TextBlock()
        lbl_pat.Text = u"\u0428\u0430\u0431\u043b\u043e\u043d \u043c\u0430\u0440\u043a\u0438 (\u043f\u0440\u0435\u0444\u0438\u043a\u0441):"
        lbl_pat.FontSize = 13
        lbl_pat.FontWeight = System.Windows.FontWeights.SemiBold
        lbl_pat.Margin = Thickness(0, 0, 0, 4)
        stack.Children.Add(lbl_pat)

        self._pattern_box = TextBox()
        self._pattern_box.Text = cur_pattern or ""
        self._pattern_box.Height = 28
        self._pattern_box.FontSize = 12
        self._pattern_box.Margin = Thickness(0, 0, 0, 16)
        stack.Children.Add(self._pattern_box)

        # Кнопки
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right

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

        stack.Children.Add(btn_panel)
        self.Content = stack

    def _on_ok(self, sender, e):
        """Сохранить настройки."""
        self.result_settings = {
            "category": self._category_combo.SelectedItem or "",
            "pattern": self._pattern_box.Text or ""
        }
        self.DialogResult = True

    def show_dialog(self):
        """Показать диалог и вернуть настройки."""
        if self.ShowDialog() == True:
            return self.result_settings
        return None
