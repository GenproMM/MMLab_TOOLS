# WPF Code Patterns for IronPython 2.7

Practical patterns for building WPF windows programmatically in the ВОР Валидатор plugin.

## Module-Level ControlTemplate Parsing

Parse XAML ControlTemplate strings at module level (after WPF assemblies loaded). This is the only way to get hover/pressed effects in code-built windows:

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
```

Assign to button via `.Template`:
```python
btn = Button()
btn.Content = u"\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c"
btn.Template = _PRIMARY_BUTTON_TEMPLATE
```

## Grid Construction (Verbose IronPython)

Grid requires explicit `ColumnDefinition` / `RowDefinition` objects:

```python
from System.Windows import GridLength, GridUnitType
from System.Windows.Controls import Grid, ColumnDefinition, RowDefinition

grid = Grid()

col1 = ColumnDefinition()
col1.Width = GridLength(1, GridUnitType.Star)
col2 = ColumnDefinition()
col2.Width = GridLength(1, GridUnitType.Star)
grid.ColumnDefinitions.Add(col1)
grid.ColumnDefinitions.Add(col2)

row1 = RowDefinition()
row1.Height = GridLength(1, GridUnitType.Star)
grid.RowDefinitions.Add(row1)

# Position elements
Grid.SetColumn(element, 0)
Grid.SetRow(element, 0)
grid.Children.Add(element)
```

For auto-sized rows/columns (no explicit width/height needed):
```python
row = RowDefinition()
row.Height = GridLength.Auto  # GridLength.Auto is a static property
```

## DockPanel Pattern

Used for modeless results windows (header docked top, buttons docked bottom, content fills):

```python
from System.Windows.Controls import DockPanel

root = DockPanel()
root.Margin = Thickness(12)

# Top header
header_panel = StackPanel()
DockPanel.SetDock(header_panel, System.Windows.Controls.Dock.Top)
header_panel.Margin = Thickness(0, 0, 0, 8)
# ... add header content ...
root.Children.Add(header_panel)

# Bottom buttons
btn_panel = StackPanel()
btn_panel.Orientation = Orientation.Horizontal
btn_panel.HorizontalAlignment = HorizontalAlignment.Right
DockPanel.SetDock(btn_panel, System.Windows.Controls.Dock.Bottom)
# ... add buttons ...
root.Children.Add(btn_panel)

# Center content (added LAST to fill remaining space)
content = ScrollViewer()
# ... set content ...
root.Children.Add(content)
```

## Expander for Categorized Results

Collapsible sections for grouped results:

```python
from System.Windows.Controls import Expander

expander = Expander()
expander.Header = u"\u041a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044f (5)"
expander.IsExpanded = True

listbox = ListBox()
listbox.FontSize = 12
# ... add items ...
expander.Content = listbox
```

## ListBox with Revit Element Selection

Click on item selects element in Revit model:

```python
listbox = ListBox()
listbox.FontSize = 12
listbox.BorderThickness = Thickness(0)

for info in elements:
    item = ListBoxItem()
    item.Content = info["name"]
    item.Tag = info["id"]  # DB.ElementId
    listbox.Items.Add(item)

def _on_list_select(sender, e):
    selected = sender.SelectedItem
    if selected and hasattr(selected, 'Tag') and selected.Tag:
        self._select_elements([selected.Tag])

listbox.SelectionChanged += _on_list_select

# Helper method:
def _select_elements(self, elem_ids):
    from pyrevit import revit
    from System.Collections.Generic import List
    try:
        uidoc = revit.uidoc
        ids = List[DB.ElementId]()
        for eid in elem_ids:
            ids.Add(eid)
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        pass
```

## Modal Dialog Lifecycle

```python
class MyDialog(Window):
    def __init__(self, doc, current_settings):
        # ... build UI ...
        self.result_settings = None
        btn_save = _make_primary_button(u"\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c", self._on_save)
        btn_save.IsDefault = True
        btn_cancel = _make_secondary_button(u"\u041e\u0442\u043c\u0435\u043d\u0430")
        btn_cancel.IsCancel = True

    def _on_save(self, sender, e):
        self.result_settings = {"key": "value"}
        self.DialogResult = True

    def show_dialog(self):
        if self.ShowDialog() == True:
            return self.result_settings
        return None
```

## Modeless Window Singleton

For results windows that should not duplicate. Called from the script's `show_results(doc, section, project, settings)` hook, not from `run()` — opening a window inside `run()` blocks the run window's dashboard and fires a window on every run:

```python
# In the parent script, called by show_results():
def _show_results_window(doc, results_data):
    import __main__
    attr_name = "_myscript_results_window"   # unique per script!
    existing = getattr(__main__, attr_name, None)
    if existing:
        try:
            existing.Close()
        except Exception:
            pass

    from results_window import ResultsWindow
    win = ResultsWindow(doc, results_data)
    win.Show()
    setattr(__main__, attr_name, win)
```

## Event Handlers in Loops

Capture loop variables with default arguments:

```python
# WRONG — all handlers reference the last item
for name in items:
    btn.Click += lambda s, e: self._handle(name)  # bug: name is last value

# CORRECT — default arg captures current value
for name in items:
    btn.Click += lambda s, e, n=name: self._handle(n)
```

## Null-Checking WPF Values

```python
# IsChecked is Nullable<bool> — always cast
is_checked = bool(cb.IsChecked)

# ComboBox selection might be None
value = combo.SelectedItem or ""

# Use getattr for safe attribute access on WPF objects
tag = getattr(item, 'Tag', None)
```

## Cyrillic String Escaping

All Cyrillic text in IronPython runtime must use `\uXXXX` escapes when loaded via `imp.load_module`:

```python
# Russian text → Unicode escapes
u"\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c"      # Сохранить
u"\u041e\u0442\u043c\u0435\u043d\u0430"                          # Отмена
u"\u0412\u044b\u0431\u0440\u0430\u0442\u044c \u0432\u0441\u0435"  # Выбрать все
u"\u0421\u043d\u044f\u0442\u044c \u0432\u0441\u0435"             # Снять все
u"\u0417\u0430\u043a\u0440\u044b\u0442\u044c"                     # Закрыть
u"\u0414\u0430"                                                    # Да
u"\u041d\u0435\u0442"                                              # Нет
u"\u041e\u041a"                                                    # ОК
```

## Window Subclass Pattern

The default pattern for all script windows:

```python
class MyWindow(Window):
    def __init__(self, doc, current_settings):
        self.Title = u"..."
        self.Width = 400
        self.Height = 300
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.Background = _BG_WINDOW  # Always #F5F5F5
        self.result_settings = None

        # Build UI
        root = StackPanel()
        root.Margin = Thickness(16)
        # ... add children ...
        self.Content = root
```

## Toggle All Pattern (Select All / Deselect All)

Common pattern for checkbox lists:

```python
self._checks = {}  # name → CheckBox

# ... populate checkboxes ...
for name in items:
    cb = CheckBox()
    cb.Content = name
    cb.FontSize = 13
    cb.IsChecked = saved.get(name, True)
    cb.Margin = Thickness(0, 4, 0, 0)
    self._checks[name] = cb

def _on_select_all(self, sender, e):
    for cb in self._checks.values():
        cb.IsChecked = True

def _on_deselect_all(self, sender, e):
    for cb in self._checks.values():
        cb.IsChecked = False

def _collect_settings(self):
    return {name: bool(cb.IsChecked) for name, cb in self._checks.items()}
```

## Hide Unchecked Pattern

For settings windows with category checkboxes, include a toggle that hides unchecked items:

```python
from System.Windows import Visibility

# In __init__, after toggle row:
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
            cb.Visibility = Visibility.Collapsed
        else:
            cb.Visibility = Visibility.Visible
```

**Important**: `_on_select_all` and `_on_deselect_all` must call `_on_hide_unchecked` after toggling to refresh visibility. When "Deselect all" is clicked while hide-unchecked is active, all items collapse — the user can uncheck the toggle to see everything again.

## Rebuilding UI Dynamically

For windows that need to refresh (e.g., results with exclude/return):

```python
def _rebuild(self):
    self.Content = None
    self._build_ui()
```

Clear and rebuild the entire content tree. Simple and reliable in IronPython WPF.
