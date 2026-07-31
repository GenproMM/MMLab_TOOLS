---
name: VorUICreator
description: >
  Generates WPF window classes (settings, results, confirm, input, run/dashboard) for the
  pyRevit ВОР Валидатор plugin, following the canonical design system and handling all
  IronPython 2.7 WPF constraints. Use when creating or editing any file that subclasses
  WPF Window or builds WPF controls — triggers: "создать окно", "create window", "build UI",
  "настроить интерфейс", "построить окно результатов", "настройки окна", "добавить диалог",
  "make dialog", "settings/results/input/confirmation window", any file in ui/*.py or
  scripts/*/lib/*.py. Also use for the run/dashboard batch window (progress + Stop) and when
  a window freezes or needs modeless Show().
version: 0.2.0
---

# VOR UI Creator

## Purpose

Generate WPF window classes for the ВОР Валидатор plugin. Every generated window must follow the canonical design system — colors, typography, layout — so windows from script `lib/` folders are visually consistent with the main plugin windows built from XAML.

## Code vs XAML

**Default: pure Python code (Window subclass).** Use for all script windows in `lib/` subfolders. Single file, no external XAML. Dynamic content built programmatically.

**XAML string + XamlReader.Parse:** only for main plugin windows in `ui/` folder when complex `Window.Resources` styles or `ControlTemplate` triggers are needed. Requires two files (`.py` + `.xaml`).

## Design System Constants

Every generated window file must include these module-level constants after imports. They are copy-pasted, not imported — scripts loaded via `imp.load_module` have isolated namespaces and cannot reliably import shared modules.

```python
# ── Colors ──
from System.Windows.Media import SolidColorBrush, Colors, Color

_BG_WINDOW = SolidColorBrush(Color.FromArgb(255, 245, 245, 245))    # #F5F5F5
_BG_CARD = SolidColorBrush(Colors.White)
_BG_PANEL = SolidColorBrush(Color.FromArgb(255, 250, 250, 250))     # #FAFAFA
_BRUSH_BORDER = SolidColorBrush(Color.FromArgb(255, 204, 204, 204)) # #CCCCCC

_BG_PRIMARY = SolidColorBrush(Color.FromArgb(255, 0, 122, 204))     # #007ACC
_BG_PRIMARY_HOVER = SolidColorBrush(Color.FromArgb(255, 0, 90, 158)) # #005A9E
_FG_ON_PRIMARY = SolidColorBrush(Colors.White)

_BG_SECONDARY = SolidColorBrush(Color.FromArgb(255, 224, 224, 224))  # #E0E0E0
_BG_SECONDARY_HOVER = SolidColorBrush(Color.FromArgb(255, 208, 208, 208)) # #D0D0D0
_BRUSH_SECONDARY_BORDER = SolidColorBrush(Color.FromArgb(255, 204, 204, 204)) # #CCCCCC

_FG_TITLE = SolidColorBrush(Color.FromArgb(255, 51, 51, 51))        # #333333
_FG_SUBTITLE = SolidColorBrush(Color.FromArgb(255, 85, 85, 85))     # #555555
_FG_DESCRIPTION = SolidColorBrush(Color.FromArgb(255, 102, 102, 102)) # #666666
_FG_MUTED = SolidColorBrush(Colors.Gray)

_FG_ERROR = SolidColorBrush(Color.FromArgb(255, 211, 47, 47))       # #D32F2F
_FG_SUCCESS = SolidColorBrush(Color.FromArgb(255, 56, 142, 60))     # #388E3C
_FG_WARNING = SolidColorBrush(Color.FromArgb(255, 255, 152, 0))     # #FF9800
_FG_INFO = SolidColorBrush(Color.FromArgb(255, 33, 150, 243))       # #2196F3

_BG_SELECTED = SolidColorBrush(Color.FromArgb(40, 0, 122, 204))     # #007ACC ~16%
```

Set `Background = _BG_WINDOW` (#F5F5F5) on every window. White feels harsh and breaks visual consistency with the main plugin windows.

## Styled Button Templates

Code-built windows need `XamlReader.Parse` ControlTemplate strings for hover/pressed effects. Define at module level:

```python
from System.Windows.Markup import XamlReader

_PRIMARY_BUTTON_TEMPLATE = XamlReader.Parse(
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="Button">'
    '<Border x:Name="Bd" Background="#007ACC" CornerRadius="4" Padding="12,6">'
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
```

## Helper Functions

Include these in every generated file. They produce styled WPF controls matching the design system:

```python
def _make_title(text, font_size=16):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = font_size
    tb.FontWeight = FontWeights.Bold
    tb.Foreground = _FG_TITLE
    tb.Margin = Thickness(0, 0, 0, 12)
    return tb

def _make_section(text):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = 14
    tb.FontWeight = FontWeights.SemiBold
    tb.Foreground = _FG_TITLE
    tb.Margin = Thickness(0, 8, 0, 4)
    return tb

def _make_hint(text):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = 11
    tb.FontStyle = FontStyles.Italic
    tb.Foreground = _FG_MUTED
    tb.Margin = Thickness(0, 0, 0, 8)
    return tb

def _make_label(text):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = 13
    tb.FontWeight = FontWeights.SemiBold
    tb.Foreground = _FG_TITLE
    tb.Margin = Thickness(0, 8, 0, 4)
    return tb

def _make_primary_button(text, handler, width=100, height=30):
    btn = Button()
    btn.Content = text
    btn.Width = width
    btn.Height = height
    btn.Foreground = _FG_ON_PRIMARY
    btn.FontWeight = FontWeights.SemiBold
    btn.Template = _PRIMARY_BUTTON_TEMPLATE
    btn.IsDefault = True
    if handler:
        btn.Click += handler
    return btn

def _make_secondary_button(text, handler=None, width=100, height=30):
    btn = Button()
    btn.Content = text
    btn.Width = width
    btn.Height = height
    btn.Template = _SECONDARY_BUTTON_TEMPLATE
    btn.IsCancel = True
    if handler:
        btn.Click += handler
    return btn

def _make_button_panel(save_text, save_handler, cancel_text=None):
    cancel_text = cancel_text or u"\u041e\u0442\u043c\u0435\u043d\u0430"
    panel = StackPanel()
    panel.Orientation = Orientation.Horizontal
    panel.HorizontalAlignment = HorizontalAlignment.Right
    panel.Margin = Thickness(0, 12, 0, 0)

    save = _make_primary_button(save_text, save_handler)
    save.IsDefault = True
    save.Margin = Thickness(0, 0, 8, 0)
    panel.Children.Add(save)

    cancel = _make_secondary_button(cancel_text)
    panel.Children.Add(cancel)
    return panel

def _make_scroll_list(max_height=300):
    scroll = ScrollViewer()
    scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
    scroll.MaxHeight = max_height
    panel = StackPanel()
    scroll.Content = panel
    return scroll, panel

def _make_card_border():
    border = Border()
    border.Background = _BG_CARD
    border.BorderBrush = _BRUSH_BORDER
    border.BorderThickness = Thickness(1)
    border.CornerRadius = CornerRadius(3)
    border.Padding = Thickness(8, 6, 8, 6)
    return border
```

## Dialog Pattern

Every modal dialog must follow this universal layout:

1. `StackPanel` root with `Margin = Thickness(16)`
2. Title `TextBlock` — `_make_title(text, font_size=16)` for window, `_make_section(text)` for sub-sections
3. Optional hint — `_make_hint(text)`
4. Content area — direct controls or `_make_scroll_list()` for lists
5. Button panel at bottom-right — `_make_button_panel(save_text, save_handler)`

```python
root = StackPanel()
root.Margin = Thickness(16)
root.Children.Add(_make_title(u"\u0417\u0430\u0433\u043e\u043b\u043e\u0432\u043e\u043a"))
# ... content ...
root.Children.Add(_make_button_panel(u"\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c", self._on_save))
self.Content = root
```

### Settings Dialog Layout

Settings windows use the standard `StackPanel` dialog pattern with a scrollable card in the middle. The `MaxHeight` on `ScrollViewer` prevents the window from growing infinitely:

1. Title + hint
2. Toggle buttons row (Select All / Deselect All)
3. Hide-unchecked checkbox
4. Scrollable card with checkbox list (`_make_scroll_list()` + `_make_card_border()`)
5. Button panel at bottom-right

```python
root = StackPanel()
root.Margin = Thickness(16)

root.Children.Add(_make_title(u"\u0417\u0430\u0433\u043e\u043b\u043e\u0432\u043e\u043a"))
root.Children.Add(_make_hint(u"\u041f\u043e\u0434\u0441\u043a\u0430\u0437\u043a\u0430"))

# Toggle row
toggle_panel = StackPanel()
toggle_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
toggle_panel.Margin = Thickness(0, 0, 0, 4)
toggle_panel.Children.Add(btn_select_all)
toggle_panel.Children.Add(btn_deselect_all)
root.Children.Add(toggle_panel)

# Hide unchecked
root.Children.Add(self._hide_unchecked_cb)

# Scrollable card with checkboxes
scroll, cats_panel = _make_scroll_list(max_height=300)
card = _make_card_border()
# ... add checkboxes to cats_panel ...
card.Child = cats_panel
scroll.Content = card
root.Children.Add(scroll)

root.Children.Add(_make_button_panel(u"\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c", self._on_save))
self.Content = root
```

**Checkbox list must be wrapped inside `_make_card_border()`** for visual grouping.

**Hide unchecked toggle**: Every settings window with category checkboxes must include a "Скрыть невыбранные" checkbox that collapses unchecked items:

```python
self._hide_unchecked_cb = CheckBox()
self._hide_unchecked_cb.Content = u"\u0421\u043a\u0440\u044b\u0442\u044c \u043d\u0435\u0432\u044b\u0431\u0440\u0430\u043d\u043d\u044b\u0435"
self._hide_unchecked_cb.FontSize = 12
self._hide_unchecked_cb.Foreground = _FG_SUBTITLE
self._hide_unchecked_cb.Margin = Thickness(0, 4, 0, 8)
self._hide_unchecked_cb.Checked += self._on_hide_unchecked
self._hide_unchecked_cb.Unchecked += self._on_hide_unchecked

def _on_hide_unchecked(self, sender, e):
    hide = bool(self._hide_unchecked_cb.IsChecked)
    for name, cb in self._checks.items():
        if hide and not bool(cb.IsChecked):
            cb.Visibility = System.Windows.Visibility.Collapsed
        else:
            cb.Visibility = System.Windows.Visibility.Visible
```

`_on_select_all` and `_on_deselect_all` should call `_on_hide_unchecked` afterward — otherwise checkboxes toggled by the bulk action keep their previous visibility and the list looks stale.

## Window Types

| Type | Size | Resize | Modal | Typical Use |
|------|------|--------|-------|-------------|
| Settings | 400-500 x 350-550 | CanResize (MinWidth=350, MinHeight=300) | Modal | Script settings, category selection |
| Results | 600-900 x 450-650 | CanResize | Modeless | Validation results, element lists |
| Run (dashboard) | ~640 x 560 | CanResize | Modeless | Batch run: per-script progress + status + "Open result" button + Stop |
| Confirm | 350-450 x 150-250 | NoResize | Modal | Yes/No questions |
| Input | 350-400 x 160-250 | NoResize | Modal | Single text/number entry |

**Settings dialog**: `show_dialog()` returns settings dict on Save, `None` on Cancel.

**Results window**: `Show()` (modeless), not `ShowDialog()`. Store the single live reference on a **unique** `__main__` attribute per script (e.g. `__main__._myscript_results_window`) so it survives across calls without colliding with other scripts' windows; close-and-recreate on re-open. Include a `_select_elements()` helper for Revit model selection.

**Run / dashboard window** (`ui/run_window.py`): modeless, drives a batch run. Per-script row shows a `ProgressBar` (indeterminate) while executing, then a status icon (✓/✗/⚠) + message + an "Открыть результат" button. Execution is not multithreaded — Revit API cannot run off the UI thread — so the window yields to the WPF `Dispatcher` between scripts. Use `Dispatcher.BeginInvoke(DispatcherPriority.Background, System.Action(step))` rather than `Dispatcher.Invoke(...)`. The reason matters: `Invoke` is synchronous, blocking the calling thread until the step completes, so WPF never gets a chance to repaint the ProgressBar or process the Stop click — the UI freezes for the entire run. `BeginInvoke` posts the step to the queue and returns immediately, letting the message loop repaint between steps. Stop takes effect only between scripts (checked at the start of each step); it cannot interrupt a script mid-`run()`. See `ui/run_window.py` for the reference implementation.

**Confirm dialog**: static `show(message, title)` classmethod returning `True`/`False`.

**Input form**: `show_dialog()` returns entered string on OK, `None` on Cancel.

## IronPython WPF Constraints

- **.NET enums cannot be imported by name** — use fully qualified names: `System.Windows.ResizeMode.CanResize`, `System.Windows.Visibility.Collapsed`, `System.Windows.Controls.Orientation.Horizontal`, `System.Windows.Controls.ScrollBarVisibility.Auto`. A bare `from System.Windows import ResizeMode` (or `from System.Windows.Controls import Dock`) fails with `CreateInstance() takes at most 4 arguments (2 given)`.
- **Avoid `from System.Windows.Controls import *`** — wildcard brings in classes but breaks enum resolution. Import controls by explicit name instead.
- **Avoid `DockPanel` as root layout** — it requires the `Dock` enum (`System.Windows.Controls.Dock.Top/Bottom`) on every child via `DockPanel.SetDock()`. Use `StackPanel` root instead — simpler and enum-free.
- **`Thickness` takes 1 or 4 arguments**: `Thickness(8)` for uniform, `Thickness(8, 6, 8, 6)` for left/top/right/bottom. The 2-argument form `Thickness(8, 6)` does not exist in .NET and throws `CreateInstance() takes at most 4 arguments (2 given)`.
- `IsChecked` is `Nullable<bool>` — wrap it: `bool(cb.IsChecked)`
- Set `DialogResult` directly: `self.DialogResult = True`
- Lambda closures: `lambda s, e, n=name: handler(n)` — default args capture loop variables
- WPF objects reject arbitrary Python attributes — store refs in `self._dict` instead
- `FontStyles.Italic` from `System.Windows`
- `CornerRadius` from `System.Windows`
- Use `codecs.open(path, "r", "utf-8")` for XAML files. The `encoding=` kwarg is unsupported in IronPython 2.7 and raises.
- Cyrillic in runtime strings: use `\uXXXX` escapes (e.g. `u"\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c"`)

## Required Imports Boilerplate

```python
# -*- coding: utf-8 -*-
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import (
    Window, Thickness, FontWeights, FontStyles, CornerRadius,
    HorizontalAlignment, VerticalAlignment
)
from System.Windows.Controls import (
    StackPanel, TextBlock, CheckBox, Button, Border,
    ScrollViewer, ScrollBarVisibility
)
from System.Windows.Media import SolidColorBrush, Colors, Color
from System.Windows.Markup import XamlReader
```

## Interaction with VorScriptCreator

VorUICreator generates the window class code. VorScriptCreator defines the script contract, metadata, file placement, and the `run()` / `show_results()` / `show_settings()` hooks. When generating a window for a validation script:

**Settings window** (`scripts/<Name>/lib/settings_window.py`):
1. The parent script defines `HAS_SETTINGS = True` and `show_settings(doc, current_settings)` which imports and calls the window class
2. `show_dialog()` returns a settings dict or `None`

**Results window** (`scripts/<Name>/lib/results_window.py`):
1. The parent script defines `show_results(doc, section, project, settings)` which imports and calls the window class — this is called by the run window's "Открыть результат" button
2. The parent's `run()` should not open the window itself — it only computes and caches data in a module-level `_last_results_data` that `show_results()` reads back. Opening from `run()` blocks the dashboard and fires a window on every run.
3. Follow the Results window type above: modeless `Show()`, unique per-script `__main__` singleton, `_select_elements()` helper.

**Run / dashboard window** (`ui/run_window.py`) — engine-owned, generated once. Do not regenerate per script.

## Generation Rules

The generated file is self-contained: it carries the design-system constants, button templates, and `_make_*` helpers inline (IronPython isolates namespaces, so these are pasted, not imported — they're emitted by the skill automatically). The rules below cover the non-obvious decisions. For the rest, copy the nearest example in `examples/` and adapt it — examples beat rules.

1. **Window layout**: `Background = _BG_WINDOW`, dialog pattern `StackPanel(16) → title → content → bottom-right buttons`.
2. **`show_dialog()` contract**: returns settings dict on `DialogResult=True`, `None` otherwise.
3. **Results windows**: include a `_select_elements(elem_ids)` method that calls `uidoc.Selection.SetElement_ids(...)` for click-to-select in Revit.
4. **Dynamic control refs**: store in `self._`-prefixed dicts (e.g. `self._checks`). WPF objects reject arbitrary Python attributes.
5. **Settings windows**: `ResizeMode.CanResize` with `MinWidth=350, MinHeight=300` so users can expand to see all checkboxes; `StackPanel` root with `_make_scroll_list(max_height=300)` to cap height.
6. **Category-checkbox settings**: wrap the list in `_make_card_border()` and include a "hide unchecked" toggle (`\u0421\u043a\u0440\u044b\u0442\u044c \u043d\u0435\u0432\u044b\u0431\u0440\u0430\u043d\u043d\u044b\u0435`) that collapses unchecked items; `_on_select_all`/`_on_deselect_all` call `_on_hide_unchecked` afterward.
7. **Cyrillic**: all runtime `u"..."` strings use `\uXXXX` escapes (see IronPython WPF Constraints).
8. **Run/dashboard window**: see the Run window type above for `Dispatcher.BeginInvoke` + Stop. Do not regenerate `ui/run_window.py` per script — it's engine-owned.

## Additional Resources

### Reference Files
- **`references/design-system.md`** — complete token reference: colors, typography, spacing, button specs
- **`references/wpf-patterns.md`** — IronPython WPF code recipes: Grid, Expander, ListBox, event handlers

### Examples
- **`examples/settings-dialog/settings_window.py`** — modal settings with checkboxes, Select All/Deselect All
- **`examples/results-window/results_window.py`** — modeless results with Expander per category, Revit selection
- **`examples/confirm-dialog/confirm_dialog.py`** — Yes/No confirmation dialog
- **`examples/input-form/input_form.py`** — single text input with OK/Cancel
- **`ui/run_window.py`** + **`ui/run_window.xaml`** — the live run/dashboard window (batch run, per-script progress, Stop, "Открыть результат"). Reference implementation for the Run window type.
- **`ui/standard_results_window.py`** — the engine's fallback results window for scripts without `show_results`
