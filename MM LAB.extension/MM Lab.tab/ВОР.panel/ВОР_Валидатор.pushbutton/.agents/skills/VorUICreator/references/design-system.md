# Design System Reference

Complete token reference for ВОР Валидатор UI. All values extracted from the canonical XAML styles in `ui/main_window.xaml` and `ui/settings_window.xaml`.

## Color Tokens

### Backgrounds

| Token | Hex | ARGB | Brush Construction | Usage |
|-------|-----|------|--------------------|-------|
| `_BG_WINDOW` | `#F5F5F5` | (255, 245, 245, 245) | `SolidColorBrush(Color.FromArgb(255, 245, 245, 245))` | Every window `Background` |
| `_BG_CARD` | `#FFFFFF` | White | `SolidColorBrush(Colors.White)` | Card borders, inner panels, list containers |
| `_BG_PANEL` | `#FAFAFA` | (255, 250, 250, 250) | `SolidColorBrush(Color.FromArgb(255, 250, 250, 250))` | Description borders, info panels |
| `_BG_SELECTED` | `#007ACC` 16% | (40, 0, 122, 204) | `SolidColorBrush(Color.FromArgb(40, 0, 122, 204))` | Selected row highlight |

### Primary Action (Blue)

| Token | Hex | Usage |
|-------|-----|-------|
| `_BG_PRIMARY` | `#007ACC` | Primary button background, Add button |
| `_BG_PRIMARY_HOVER` | `#005A9E` | Primary button hover |
| — (pressed) | `#004578` | Primary button pressed |
| `_FG_ON_PRIMARY` | White | Text/icon color on primary backgrounds |

### Secondary Action (Gray)

| Token | Hex | Usage |
|-------|-----|-------|
| `_BG_SECONDARY` | `#E0E0E0` | Secondary button background |
| `_BG_SECONDARY_HOVER` | `#D0D0D0` | Secondary button hover |
| — (pressed) | `#C0C0C0` | Secondary button pressed |
| `_BRUSH_SECONDARY_BORDER` | `#CCCCCC` | Secondary button border |

### Text

| Token | Hex | Usage |
|-------|-----|-------|
| `_FG_TITLE` | `#333333` | Page titles, section headers, field labels |
| `_FG_SUBTITLE` | `#555555` | Sub-section text (e.g. "Scripts check") |
| `_FG_DESCRIPTION` | `#666666` | Description text, secondary info |
| `_FG_MUTED` | `Colors.Gray` | Hint text, empty states, disabled text |

### Semantic

| Token | Hex | Usage |
|-------|-----|-------|
| `_FG_ERROR` | `#D32F2F` | Problem headers, error indicators |
| `_FG_SUCCESS` | `#388E3C` | Success/exception headers |
| `_FG_WARNING` | `#FF9800` | Warning indicators, material section headers |
| `_FG_INFO` | `#2196F3` | Info indicators, family section headers |

### Semantic Backgrounds (for results cells)

| Name | Hex | Usage |
|------|-----|-------|
| Error cell BG | `#FFEBEE` | Problem cell background |
| Success cell BG | `#E8F5E9` | Exception/success cell background |

### Borders

| Token | Hex | Usage |
|-------|-----|-------|
| `_BRUSH_BORDER` | `#CCCCCC` | Card borders, list borders, separator lines |

## Typography Scale

| Role | FontSize | FontWeight | Foreground | Bottom Margin |
|------|----------|------------|------------|---------------|
| Page title | 18 | Bold | `_FG_TITLE` | 12 |
| Window title | 16 | Bold | `_FG_TITLE` | 12-16 |
| Section header | 14 | SemiBold | `_FG_TITLE` | 4-8 |
| Sub-section | 14 | SemiBold | `_FG_SUBTITLE` | 4 |
| Field label | 13 | SemiBold | `_FG_TITLE` | 4 |
| Body text | 13 | Normal | `_FG_TITLE` | — |
| Checkbox text | 13 | Normal | `_FG_TITLE` | — |
| ComboBox text | 13 | Normal | `_FG_TITLE` | — |
| Description | 12 | Normal | `_FG_DESCRIPTION` | — |
| ListBox item | 11-12 | Normal | `_FG_TITLE` | — |
| Hint text | 11 | Italic | `_FG_MUTED` | 8 |
| Toggle button | 11 | Normal | `_FG_TITLE` | — |
| Inline action | 10-11 | Normal | `_FG_ON_PRIMARY` or `_FG_TITLE` | — |

## Spacing and Layout

### Margins

| Constant | Value (px) | Usage |
|----------|-----------|-------|
| Root margin | 16 | `StackPanel.Margin = Thickness(16)` on root container |
| Group spacing | 12 | Between major UI sections |
| Sub-group spacing | 8 | Between related controls, after hint text |
| Label-to-input | 4 | Bottom margin on label before input |
| Button row top | 12 | `Thickness(0, 12, 0, 0)` on button panel |
| Save button right | 8 | `Thickness(0, 0, 8, 0)` on Save button (gap before Cancel) |

### Button Dimensions

| Button Type | Width | Height | Padding | CornerRadius | FontSize | FontWeight |
|-------------|-------|--------|---------|-------------|----------|------------|
| Primary action (large) | Auto (stretch) | 42 | 12,6 | 4 | 14 | SemiBold |
| Primary dialog | 100 | 30 | 12,6 | 4 | — | SemiBold |
| Secondary dialog | 100 | 30 | 10,5 | 3 | — | — |
| Toggle (Select/Deselect all) | 100 | 24 | — | — | 11 | — |
| Inline action (Exclude/Return) | 75 | 22 | — | — | 10 | — |
| Icon button | 26x26 | — | — | 4 | — | — |

### Input Dimensions

| Control | Height | FontSize | Padding |
|---------|--------|----------|---------|
| TextBox | 28 | 13 | — |
| ComboBox | 30 | 13 | 6,4 |
| CheckBox | Auto | 13 | 6,4 |
| ListBox | Auto | 11-12 | — |

### ScrollViewer

| Context | MaxHeight |
|---------|-----------|
| Settings category list | 250 |
| Generic settings | 250-300 |
| Results element list | 340-420 |
| Settings window (projects/sections) | 300 |

### Border / Card

| Property | Value |
|----------|-------|
| BorderBrush | `_BRUSH_BORDER` (#CCCCCC) |
| BorderThickness | 1 |
| CornerRadius | 3 |
| Background (card) | `_BG_CARD` (White) |
| Background (panel) | `_BG_PANEL` (#FAFAFA) |
| Padding | 8,6 |

## Window Configurations

| Type | Width | Height | ResizeMode | StartupLocation | Modal |
|------|-------|--------|------------|-----------------|-------|
| Main plugin | 520 | 620 | CanResize | CenterScreen | Modal |
| Settings management | 700 | 500 | NoResize | CenterScreen | Modal |
| Script settings | 400-500 | 350-550 | CanResize | CenterScreen | Modal |
| Script results | 600-900 | 450-650 | CanResize | CenterScreen | Modeless |
| Run / dashboard | 640 | 560 | CanResize | CenterOwner | Modeless |
| Confirmation | 350-450 | 150-250 | NoResize | CenterOwner | Modal |
| Input form | 350-400 | 160-250 | NoResize | CenterScreen | Modal |
