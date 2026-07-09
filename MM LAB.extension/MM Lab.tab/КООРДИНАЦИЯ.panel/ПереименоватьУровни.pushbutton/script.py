# -*- coding: utf-8 -*-

from Autodesk.Revit import DB
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document

# ----------------------------------------------------------------------
# Параметр
# ----------------------------------------------------------------------

param_name = forms.ask_for_string(
    default="GP_01_Этаж_Имя",
    prompt="Введите параметр для записи номера этажа:",
    title="Нумерация уровней"
)

if not param_name:
    forms.alert("Отменено", exitscript=True)

# ----------------------------------------------------------------------
# Формат номера
# ----------------------------------------------------------------------

format_option = forms.ask_for_one_item(
    ["1", "01", "001"],
    default="01",
    prompt="Выберите формат числа:",
    title="Формат"
)

if not format_option:
    forms.alert("Отменено", exitscript=True)

# ----------------------------------------------------------------------
# Префиксы имен уровней
# ----------------------------------------------------------------------

prefix_up = forms.ask_for_string(
    default="KR_L",
    prompt="Введите префикс для 1-го этажа и выше:",
    title="Префикс верхних уровней"
)

if prefix_up is None:
    forms.alert("Отменено", exitscript=True)

prefix_up = prefix_up.strip()

prefix_down = forms.ask_for_string(
    default="KR_B",
    prompt="Введите префикс для уровней ниже 1-го этажа:",
    title="Префикс нижних уровней"
)

if prefix_down is None:
    forms.alert("Отменено", exitscript=True)

prefix_down = prefix_down.strip()


# ----------------------------------------------------------------------
# Формат номера этажа
# ----------------------------------------------------------------------

def fmt(n):
    if format_option == "01":
        return "{:02d}".format(n)
    elif format_option == "001":
        return "{:03d}".format(n)
    else:
        return "{}".format(n)


# ----------------------------------------------------------------------
# Формат отметки уровня
# ----------------------------------------------------------------------

def format_elevation(level):
    # Перевод из внутренних единиц Revit (футы) в метры
    elev = DB.UnitUtils.ConvertFromInternalUnits(
        level.Elevation,
        DB.UnitTypeId.Meters
    )

    # Исключаем -0.000 из-за погрешностей
    if abs(elev) < 0.0005:
        elev = 0.0

    value = "{:.3f}".format(abs(elev)).replace(".", ",")

    if elev > 0:
        return "+{}".format(value)
    elif elev < 0:
        return "-{}".format(value)
    else:
        return value


# ----------------------------------------------------------------------
# Уровни
# ----------------------------------------------------------------------

levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))

if not levels:
    forms.alert("Нет уровней", exitscript=True)

levels.sort(key=lambda x: x.Elevation)

# ----------------------------------------------------------------------
# Выбор базового уровня
# ----------------------------------------------------------------------

base_level = forms.SelectFromList.show(
    levels,
    name_attr="Name",
    title="Выберите уровень, который будет = 1",
    multiselect=False
)

if not base_level:
    forms.alert("Отменено", exitscript=True)

base_index = levels.index(base_level)

# ----------------------------------------------------------------------
# Запись
# ----------------------------------------------------------------------

t = DB.Transaction(doc, "Нумерация уровней")
t.Start()

errors = []


def set_val(lvl, val, prefix):
    value = fmt(val)

    # Запись номера этажа в параметр
    p = lvl.LookupParameter(param_name)

    if p is None:
        errors.append("{} - отсутствует параметр '{}'".format(lvl.Name, param_name))
    elif p.IsReadOnly:
        errors.append("{} - параметр только для чтения".format(lvl.Name))
    else:
        try:
            p.Set(value)
        except Exception as ex:
            errors.append("{} - ошибка записи параметра ({})".format(lvl.Name, ex))

    # Переименование уровня
    old_name = lvl.Name

    try:
        new_name = "{}{}_{}".format(
            prefix,
            value,
            format_elevation(lvl)
        )
        lvl.Name = new_name
    except Exception as ex:
        errors.append("{} - не удалось изменить имя ({})".format(old_name, ex))


# ----------------------------------------------------------------------
# Нумерация
# ----------------------------------------------------------------------

# Базовый уровень = 1
set_val(base_level, 1, prefix_up)

# Выше базового
num = 2
for lvl in levels[base_index + 1:]:
    set_val(lvl, num, prefix_up)
    num += 1

# Ниже базового
num = 1
for lvl in reversed(levels[:base_index]):
    set_val(lvl, num, prefix_down)
    num += 1

t.Commit()

# ----------------------------------------------------------------------
# Результат
# ----------------------------------------------------------------------

if errors:
    forms.alert(
        "Готово с ошибками:\n\n{}".format("\n".join(errors)),
        title="Нумерация уровней"
    )
else:
    forms.alert(
        "Уровни успешно пронумерованы.",
        title="Нумерация уровней"
    )