#! python3
# -*- coding: utf-8 -*-
"""Совместимость Revit API для скриптов MM LAB (единый compat-модуль).

Единственное место версионных ветвлений Revit API в репозитории:
детекция версии Revit, fail-fast на неподдерживаемых версиях,
канонический каскад чтения параметров (pythonnet-обходы),
ElementId Int64 (Revit 2024), единицы измерения (2020 vs 2022+),
создание перекрытий (Floor.Create, 2022+) и .NET-interop.

Скрипты кнопок зовут стабильные хелперы этого модуля, а не сырой
версионный API. Ветвления вида ``if version >= ...`` вне compat
запрещены конвенцией MM LAB.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет

Пример:
    from revit_compat import require_supported_version, get_parameter
    from Autodesk.Revit.DB import BuiltInParameter

    COMMAND_NAME = u"Моя кнопка"
    version = require_supported_version(COMMAND_NAME)
    parameter = get_parameter(
        element, BuiltInParameter.ROOM_AREA, u"Площадь", u"Area"
    )

Модуль не открывает транзакций и сам не пишет в модель: открыть и
закрыть Transaction обязан вызывающий скрипт.
"""

import builtins
import os
import sys

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import CurveArray
from Autodesk.Revit.DB import CurveLoop
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import Floor
from Autodesk.Revit.DB import UnitUtils
from Autodesk.Revit.UI import TaskDialog

import System
from System.Collections.Generic import List


# Поддерживаемые версии Revit (D-02). Новые версии добавлять здесь
# и в ветки хелперов ниже (процедура /mm-new-compat).
SUPPORTED_VERSIONS = (2020, 2022, 2024)

# Кеш: BuiltInParameter (int) -> имя параметра из Definition.Name.
_BIP_NAME_CACHE = {}

# Ленивая карта единиц измерения; строится при первом вызове convert_*.
_UNITS_MAP = None

# Версия Revit, валидированная require_supported_version: версионные
# хелперы (_units_map, create_floor) переиспользуют её вместо повторной
# детекции — повторная детекция теряет явный аргумент ``revit=`` и на
# хосте без __revit__/HOST_APP молча уводит в ветку 2022+.
_VALIDATED_VERSION = None


# --- Версия Revit -----------------------------------------------------------


def _version_number(revit_object):
    """Достаёт целочисленную версию Revit из объекта приложения.

    Поддерживает UIApplication (``.Application.VersionNumber``),
    Application (``.VersionNumber``) и pyrevit.HOST_APP (``.version``).
    Возвращает int либо None, если версию извлечь не удалось.
    """
    if revit_object is None:
        return None

    try:
        return int(revit_object.Application.VersionNumber)
    except Exception:
        pass

    try:
        return int(revit_object.VersionNumber)
    except Exception:
        pass

    try:
        return int(revit_object.version)
    except Exception:
        pass

    return None


def get_revit_version(revit=None):
    """Возвращает версию Revit (например, 2024) либо None.

    Каскад источников (``__revit__`` внутри lib-модулей — негарантированный
    контракт, поэтому только через builtins и с фолбэками):
    1. явный аргумент ``revit`` (UIApplication/Application из скрипта);
    2. ``builtins.__revit__`` — объект, который pyRevit инжектит скрипту;
    3. ``pyrevit.HOST_APP`` — хост-платформа (если пакет pyrevit доступен);
    4. None — версию определить не удалось.
    """
    version = _version_number(revit)
    if version is not None:
        return version

    version = _version_number(getattr(builtins, "__revit__", None))
    if version is not None:
        return version

    try:
        from pyrevit import HOST_APP
    except ImportError:
        HOST_APP = None

    return _version_number(HOST_APP)


def require_supported_version(command_name, revit=None):
    """Fail-fast по версии Revit (D-03): вызывать в начале main().

    Если версия не определена или не входит в SUPPORTED_VERSIONS —
    показывает TaskDialog с перечнем поддерживаемых версий и мягко
    завершает скрипт через SystemExit. Иначе возвращает версию (int)
    и кеширует её в _VALIDATED_VERSION: дальнейшие версионные ветвления
    модуля (_units_map, create_floor) используют именно эту версию,
    даже если она была получена через явный аргумент ``revit=``.
    """
    global _VALIDATED_VERSION
    version = get_revit_version(revit)
    if version not in SUPPORTED_VERSIONS:
        supported_text = u" / ".join(str(item) for item in SUPPORTED_VERSIONS)
        version_text = str(version) if version is not None else u"не определена"
        message = (
            u"Эта кнопка поддерживает Revit {0}.\n"
            u"Текущая версия: {1}.\n"
            u"Обратись в GENPRO LAB."
        ).format(supported_text, version_text)
        TaskDialog.Show(command_name, message)
        raise SystemExit(message)
    _VALIDATED_VERSION = version
    return version


def _effective_version():
    """Версия Revit для версионных ветвлений хелперов модуля.

    Сначала — версия, валидированная require_supported_version
    (переживает и путь с явным аргументом ``revit=``), затем повторная
    детекция get_revit_version(); None — версию определить не удалось.
    """
    if _VALIDATED_VERSION is not None:
        return _VALIDATED_VERSION
    return get_revit_version()


# --- Параметры (D-04: pythonnet-обходы) --------------------------------------


def _bip_to_lookup_name(document, built_in_parameter):
    """Конвертирует BuiltInParameter в стабильное имя для LookupParameter.

    document.GetElement(ElementId(bip)) -> ParameterElement ->
    GetDefinition().Name; имя кешируется в _BIP_NAME_CACHE — каждый
    BuiltInParameter резолвится один раз за сессию (обход из
    ios_common_helpers). При неудаче кешируется None.
    """
    bip_int = int(built_in_parameter)
    if bip_int not in _BIP_NAME_CACHE:
        name = None
        try:
            param_element = document.GetElement(ElementId(bip_int))
            if param_element is not None:
                name = param_element.GetDefinition().Name
        except Exception:
            name = None
        _BIP_NAME_CACHE[bip_int] = name
    return _BIP_NAME_CACHE[bip_int]


def get_parameter(element, built_in_parameter, *fallback_names):
    """Читает параметр по BuiltInParameter каноническим каскадом.

    Единственный канонический вариант обхода в репозитории (D-01);
    третий вариант не создавать. Каскад:
    1. прямой вызов element.get_Parameter(bip) — на pythonnet 3.x
       падает с TypeError (enum воспринимается как int);
    2. явный выбор перегрузки get_Parameter.__overloads__[BuiltInParameter]
       (обход «Мокрых зон»);
    3. BuiltInParameter -> имя параметра через ParameterElement (кеш)
       -> element.LookupParameter(имя) (обход ios_common_helpers);
    4. element.LookupParameter по каждому из fallback_names
       (локализованные имена, передаются вызывающим);
    5. None — параметр не найден.

    Возвращает Parameter либо None.
    """
    # Шаг 1: прямой вызов (IronPython и часть pythonnet-сред).
    try:
        parameter = element.get_Parameter(built_in_parameter)
        if parameter is not None:
            return parameter
    except TypeError:
        pass

    # Шаг 2: явный выбор перегрузки (pythonnet 3.x).
    try:
        parameter = element.get_Parameter.__overloads__[BuiltInParameter](
            built_in_parameter
        )
        if parameter is not None:
            return parameter
    except (TypeError, AttributeError, KeyError):
        pass

    # Шаг 3: BuiltInParameter -> имя -> LookupParameter (кеш на модуль).
    try:
        document = getattr(element, "Document", None)
        if document is not None:
            name = _bip_to_lookup_name(document, built_in_parameter)
            if name:
                parameter = element.LookupParameter(name)
                if parameter is not None:
                    return parameter
    except Exception:
        pass

    # Шаг 4: локализованные имена из аргументов.
    for fallback_name in fallback_names:
        if not fallback_name:
            continue
        try:
            parameter = element.LookupParameter(fallback_name)
        except Exception:
            parameter = None
        if parameter is not None:
            return parameter

    # Шаг 5: параметр не найден.
    return None


def get_shared_parameter(element, guid):
    """Читает общий (shared) параметр по GUID.

    Принимает str или System.Guid (строка оборачивается в System.Guid).
    Прямой вызов element.get_Parameter(guid) с тем же
    __overloads__-фолбэком, что и в get_parameter.
    Возвращает Parameter либо None.
    """
    if isinstance(guid, str):
        guid = System.Guid(guid)

    try:
        return element.get_Parameter(guid)
    except TypeError:
        pass

    try:
        return element.get_Parameter.__overloads__[System.Guid](guid)
    except (TypeError, AttributeError, KeyError):
        return None


# --- ElementId: Revit 2024 = Int64 -------------------------------------------


def element_id_value(element_id):
    """Возвращает числовое значение ElementId (int); None -> -1.

    Revit 2024+ — свойство .Value (Int64); Revit 2020/2022 —
    .IntegerValue. Семантика повторяет ios_common_helpers дословно.
    """
    if element_id is None:
        return -1

    try:
        return element_id.Value
    except Exception:
        return element_id.IntegerValue


def make_element_id(id_value):
    """Создаёт ElementId из числа с учётом Int64 (Revit 2024+).

    Сначала пробует конструктор ElementId(System.Int64), затем
    ElementId(int) для старых версий API.
    """
    try:
        return ElementId(System.Int64(id_value))
    except Exception:
        return ElementId(int(id_value))


# --- Units: 2020 = DisplayUnitType, 2022/2024 = ForgeTypeId/UnitTypeId -------


def _units_map():
    """Ленивая карта единиц измерения: ключ -> объект единицы Revit API.

    Строится при первом вызове convert_from_internal/convert_to_internal,
    чтобы импорт модуля не зависел от версии Revit. Единственное
    версионное ветвление Units: Revit <= 2020 использует DisplayUnitType
    (удалён в 2022), новее — UnitTypeId (ForgeTypeId). Версия берётся
    через _effective_version (кеш require_supported_version, затем
    повторная детекция); если версия не определена (None), берётся
    современная ветка 2022+.
    """
    global _UNITS_MAP
    if _UNITS_MAP is None:
        version = _effective_version()
        if version is not None and version <= 2020:
            from Autodesk.Revit.DB import DisplayUnitType

            _UNITS_MAP = {
                "mm": DisplayUnitType.DUT_MILLIMETERS,
                "cm": DisplayUnitType.DUT_CENTIMETERS,
                "m": DisplayUnitType.DUT_METERS,
                "m2": DisplayUnitType.DUT_SQUARE_METERS,
                "m3": DisplayUnitType.DUT_CUBIC_METERS,
            }
        else:
            from Autodesk.Revit.DB import UnitTypeId

            _UNITS_MAP = {
                "mm": UnitTypeId.Millimeters,
                "cm": UnitTypeId.Centimeters,
                "m": UnitTypeId.Meters,
                "m2": UnitTypeId.SquareMeters,
                "m3": UnitTypeId.CubicMeters,
            }
    return _UNITS_MAP


def _unit_object(unit_key):
    """Возвращает объект единицы по ключу; неизвестный ключ -> ValueError."""
    units = _units_map()
    if unit_key not in units:
        raise ValueError(
            u"Неизвестный ключ единиц: {0!r}. Доступные ключи: {1}.".format(
                unit_key, u", ".join(sorted(units))
            )
        )
    return units[unit_key]


def convert_from_internal(value, unit_key):
    """Конвертирует значение из внутренних единиц Revit.

    unit_key: "mm" | "cm" | "m" | "m2" | "m3".
    Сигнатура UnitUtils.ConvertFromInternalUnits(value, unit) одинакова
    во всех поддерживаемых версиях — меняется только тип unit
    (DisplayUnitType в 2020, ForgeTypeId/UnitTypeId в 2022+).
    """
    return UnitUtils.ConvertFromInternalUnits(value, _unit_object(unit_key))


def convert_to_internal(value, unit_key):
    """Конвертирует значение во внутренние единицы Revit.

    unit_key: "mm" | "cm" | "m" | "m2" | "m3".
    """
    return UnitUtils.ConvertToInternalUnits(value, _unit_object(unit_key))


# --- Floor: 2020 = doc.Create.NewFloor, 2022+ = Floor.Create -----------------


def create_floor(doc, curve_loops, floor_type_id, level_id):
    """Создаёт перекрытие по списку CurveLoop на всех поддерживаемых версиях.

    Транзакцию НЕ открывает — открыть и закрыть Transaction обязан
    вызывающий скрипт.

    Revit 2022/2024: Floor.Create(doc, IList[CurveLoop], floor_type_id,
    level_id). Revit 2020: doc.Create.NewFloor(CurveArray, ...) — эта
    ветка не поддерживает отверстия: используется только ПЕРВЫЙ контур
    из curve_loops, остальные игнорируются. Версия — через
    _effective_version (кеш require_supported_version).

    Возвращает созданный Floor.
    """
    version = _effective_version()
    if version is not None and version <= 2020:
        first_loop = None
        for loop in curve_loops:
            first_loop = loop
            break
        if first_loop is None:
            raise ValueError(u"create_floor: пустой список контуров curve_loops.")

        curve_array = CurveArray()
        for curve in first_loop:
            curve_array.Append(curve)

        floor_type = doc.GetElement(floor_type_id)
        level = doc.GetElement(level_id)
        return doc.Create.NewFloor(curve_array, floor_type, level, False)

    return Floor.Create(
        doc, to_net_list(curve_loops, CurveLoop), floor_type_id, level_id
    )


# --- pythonnet interop --------------------------------------------------------


def to_net_list(items, net_type):
    """Собирает System.Collections.Generic.List[net_type] из последовательности.

    Многие методы Revit API принимают строго IList[T]: питоновский
    список на pythonnet в них не маршалится автоматически.
    """
    net_list = List[net_type]()
    for item in items:
        net_list.Add(item)
    return net_list


def enum_from_int(enum_type, int_value):
    """Приводит int к .NET enum через System.Enum.ToObject.

    На pythonnet 3.x неявный каст int -> enum запрещён; ToObject —
    канонический явный способ получить значение enum из числа.
    """
    return System.Enum.ToObject(enum_type, int_value)


def iter_count(sequence):
    """Безопасно считает элементы .NET IList / питоновской последовательности.

    На pythonnet 3.x обращение к .Count у IList может падать (кейс
    коммита 3fcf888 — краш «Мокрых зон»), а len() поддержан не для всех
    .NET-коллекций. Каскад: len() -> .Count -> подсчёт итерацией.
    """
    try:
        return len(sequence)
    except TypeError:
        pass

    try:
        return sequence.Count
    except Exception:
        pass

    count = 0
    for _ in sequence:
        count += 1
    return count


# --- Vendored-библиотеки (корневой lib/) --------------------------------------


def ensure_vendor_lib():
    """Подключает каталог vendored-библиотек (<репозиторий>/lib) к sys.path.

    Там лежат vendored-библиотеки (openpyxl, et_xmlfile). Вызывать после
    канонического lib-бутстрапа и только если кнопке действительно нужен
    vendored-пакет. Возвращает путь к каталогу либо None, если каталог
    не найден.
    """
    # lib -> MM LAB.extension -> корень репозитория
    repo_root = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))
    vendor_dir = os.path.join(repo_root, "lib")
    if not os.path.isdir(vendor_dir):
        return None
    if vendor_dir not in sys.path:
        sys.path.append(vendor_dir)
    return vendor_dir
