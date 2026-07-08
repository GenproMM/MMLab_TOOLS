#! python3
# -*- coding: utf-8 -*-

import math

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import ConnectorProfileType
from Autodesk.Revit.DB import Domain
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import FlowDirectionType
from Autodesk.Revit.DB import MEPCurve
from Autodesk.Revit.DB import StorageType
from Autodesk.Revit.DB.ExternalService import ExternalServiceRegistry
from Autodesk.Revit.DB.ExternalService import ExternalServices
from Autodesk.Revit.DB.Mechanical import DuctSystemType
from Autodesk.Revit.UI import TaskDialog

try:
    text_type = unicode
except NameError:
    text_type = str


UNDEFINED_LABELS = set([u"неопределено", u"notdefined"])
UNDEFINED_SERVER_LABELS = set([u"неопределено", u"notdefined", u"undefined", u"none"])
SUPPLY_LABELS = set([u"supplyair", u"приток", u"приточныйвоздух", u"приточный"])
NOT_SUPPLY_LABELS = set([
    u"exhaustair",
    u"удаляемыйвоздух",
    u"отработанныйвоздух",
    u"вытяжнойвоздух",
    u"вытяжка",
    u"returnair",
    u"возвратныйвоздух",
    u"рециркуляция",
])


class SupplyFlagDecision(object):
    SUPPLY = "supply"
    NOT_SUPPLY = "not_supply"
    IGNORE = "ignore"


DECISION_TO_LABEL = {
    SupplyFlagDecision.SUPPLY: u"Приток",
    SupplyFlagDecision.NOT_SUPPLY: u"Не приток",
    SupplyFlagDecision.IGNORE: u"Не менять",
}

LABEL_TO_DECISION = dict((value, key) for key, value in DECISION_TO_LABEL.items())


def to_text(value):
    if value is None:
        return u""

    if isinstance(value, text_type):
        return value

    try:
        return text_type(value)
    except Exception:
        return text_type(str(value))


def element_id_value(element_id):
    if element_id is None:
        return -1

    try:
        return element_id.Value
    except Exception:
        return element_id.IntegerValue


def normalize_text(value):
    text = to_text(value)
    if not text:
        return u""

    return u"".join(ch.lower() for ch in text if ch.isalnum())


def show_error(command_name, ex):
    TaskDialog.Show(command_name, u"Ошибка:\n{0}".format(to_text(ex)))


def get_document(command_name):
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None or uidoc.Document is None:
        TaskDialog.Show(command_name, u"Открой проект Revit и повтори команду.")
        return None

    return uidoc.Document


def collect_elements(document, *categories):
    result = []
    for category in categories:
        collector = FilteredElementCollector(document).OfCategory(category).WhereElementIsNotElementType()
        for element in collector:
            result.append(element)
    return result


def collect_additional_flow_elements(document):
    return collect_elements(
        document,
        BuiltInCategory.OST_DuctCurves,
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_DuctTerminal,
        BuiltInCategory.OST_FlexDuctCurves,
    )


def is_writable(parameter):
    return parameter is not None and not parameter.IsReadOnly


def nearly_equal(first, second, tolerance):
    return abs(first - second) < tolerance


def get_parameter_by_names(element, *names):
    for name in names:
        if not name:
            continue

        parameter = element.LookupParameter(name)
        if parameter is not None:
            return parameter

    return None


# Кеш: BuiltInParameter (int) → английское имя параметра из Definition.Name
_BIP_NAME_CACHE = {}


def _bip_to_lookup_name(document, built_in_parameter):
    """Конвертирует BuiltInParameter в стабильное имя для LookupParameter.

    Использует document.GetElement(ElementId(bip)) → ParameterElement →
    Definition.Name. Имя кешируется — каждый BuiltInParameter резолвится
    один раз за сессию.

    LookupParameter(string) имеет единственную перегрузку — нет проблем
    с неоднозначностью на pythonnet (Revit 2024+).
    """
    bip_int = int(built_in_parameter)
    if bip_int not in _BIP_NAME_CACHE:
        name = None
        param_element = document.GetElement(ElementId(bip_int))
        if param_element is not None:
            try:
                name = param_element.GetDefinition().Name
            except Exception:
                name = None
        _BIP_NAME_CACHE[bip_int] = name
    return _BIP_NAME_CACHE[bip_int]


def get_parameter(element, built_in_parameter, *names):
    # Получаем параметр через LookupParameter по имени из ParameterElement.
    # Никаких перегрузок — LookupParameter имеет ровно одну сигнатуру (string).
    name = _bip_to_lookup_name(element.Document, built_in_parameter)
    if name:
        parameter = element.LookupParameter(name)
        if parameter is not None:
            return parameter

    # Fallback: поиск по локализованным именам из аргументов
    return get_parameter_by_names(element, *names)


def get_parameter_text(parameter):
    if parameter is None:
        return u""

    try:
        value_string = parameter.AsValueString()
        if value_string:
            return to_text(value_string)
    except Exception:
        pass

    storage_type = parameter.StorageType
    if storage_type == StorageType.String:
        return to_text(parameter.AsString())
    if storage_type == StorageType.Integer:
        return to_text(parameter.AsInteger())
    if storage_type == StorageType.Double:
        return to_text(parameter.AsDouble())
    if storage_type == StorageType.ElementId:
        return to_text(element_id_value(parameter.AsElementId()))

    return u""


def get_supply_flag_parameter(element):
    return get_parameter_by_names(
        element,
        u"Приточный",
        u"Приток",
        u"Supply",
        u"Supply Flag",
    )


def is_undefined_text(value):
    return normalize_text(value) in UNDEFINED_LABELS


def is_parameter_undefined(parameter):
    if parameter is None:
        return False

    storage_type = parameter.StorageType
    if storage_type == StorageType.Integer:
        return parameter.AsInteger() == 0 or is_undefined_text(get_parameter_text(parameter))
    if storage_type == StorageType.ElementId:
        return element_id_value(parameter.AsElementId()) == element_id_value(ElementId.InvalidElementId) \
            or is_undefined_text(get_parameter_text(parameter))
    if storage_type == StorageType.String:
        return is_undefined_text(parameter.AsString()) or is_undefined_text(parameter.AsValueString())

    return is_undefined_text(get_parameter_text(parameter))


def try_set_parameter_to_undefined(parameter):
    if not is_writable(parameter):
        return False

    if is_parameter_undefined(parameter):
        return True

    try:
        if parameter.SetValueString(u"Не определено") and is_parameter_undefined(parameter):
            return True
    except Exception:
        pass

    try:
        if parameter.SetValueString(u"Not Defined") and is_parameter_undefined(parameter):
            return True
    except Exception:
        pass

    try:
        if parameter.StorageType == StorageType.Integer:
            return parameter.Set(0)
        if parameter.StorageType == StorageType.ElementId:
            return parameter.Set(ElementId.InvalidElementId)
        if parameter.StorageType == StorageType.String:
            return parameter.Set(u"Не определено")
    except Exception:
        return False

    return False


def get_loss_method_parameters(element):
    parameters = []
    for built_in_parameter, names in [
        (BuiltInParameter.RBS_DUCT_LOSS_METHOD_SERVER_PARAM, [u"Метод расчета потерь", u"Loss Method"]),
        (BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM, [u"Метод расчета потерь", u"Loss Method"]),
        (BuiltInParameter.RBS_DUCT_TERMINAL_LOSS_METHOD_SERVER_PARAM, [u"Метод расчета потерь", u"Loss Method"]),
    ]:
        parameter = get_parameter(element, built_in_parameter, *names)
        if parameter is not None:
            parameters.append(parameter)

    return parameters


def get_duct_loss_method_not_defined_server_id():
    service = ExternalServiceRegistry.GetService(ExternalServices.BuiltInExternalServices.DuctFittingAndAccessoryPressureDropService)
    if service is None:
        return None

    server_ids = list(service.GetRegisteredServerIds())
    for server_id in server_ids:
        server = service.GetServer(server_id)
        if server is None:
            continue

        server_name = normalize_text(server.GetName())
        if server_name in UNDEFINED_SERVER_LABELS:
            return server_id

    for server_id in server_ids:
        server = service.GetServer(server_id)
        if server is None:
            continue

        server_name = normalize_text(server.GetName())
        if "undefined" in server_name or "notdefined" in server_name or "неопредел" in server_name:
            return server_id

    return None


def ensure_loss_method_undefined(family_instance, not_defined_server_id):
    changed = False

    for parameter in get_loss_method_parameters(family_instance):
        if not is_writable(parameter):
            continue

        if is_parameter_undefined(parameter):
            continue

        set_success = False

        if not_defined_server_id is not None and parameter.StorageType == StorageType.ElementId:
            try:
                set_success = parameter.Set(not_defined_server_id)
            except Exception:
                set_success = False

            if set_success and is_parameter_undefined(parameter):
                changed = True
                continue

        if try_set_parameter_to_undefined(parameter) and is_parameter_undefined(parameter):
            changed = True

    return changed


def get_default_supply_flag_decision(classification):
    normalized = normalize_text(classification)

    if normalized in SUPPLY_LABELS:
        return SupplyFlagDecision.SUPPLY
    if normalized in NOT_SUPPLY_LABELS:
        return SupplyFlagDecision.NOT_SUPPLY

    if "supply" in normalized or "прит" in normalized:
        return SupplyFlagDecision.SUPPLY
    if "exhaust" in normalized or "return" in normalized or "выт" in normalized or "удал" in normalized:
        return SupplyFlagDecision.NOT_SUPPLY

    return SupplyFlagDecision.IGNORE


def get_connector_manager(element):
    if element is None:
        return None

    if isinstance(element, MEPCurve):
        return element.ConnectorManager

    model = getattr(element, "MEPModel", None)
    if model is not None:
        return model.ConnectorManager

    return None


def get_hvac_connectors(element):
    manager = get_connector_manager(element)
    if manager is None:
        return []

    result = []
    for connector in manager.Connectors:
        if connector.Domain == Domain.DomainHvac and connector.Direction != FlowDirectionType.In:
            result.append(connector)

    return result


def get_system_classification(element):
    classification = normalize_text(
        get_parameter_text(
            get_parameter(
                element,
                BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM,
                u"Классификация системы",
            )
        )
    )
    if classification:
        return classification

    for connector in get_hvac_connectors(element):
        try:
            system_type = connector.DuctSystemType
            if system_type != DuctSystemType.UndefinedSystemType:
                return to_text(system_type)
        except Exception:
            pass

    return u""


def get_connector_area(connector):
    if connector.Shape == ConnectorProfileType.Round:
        return math.pi * connector.Radius * connector.Radius
    if connector.Shape == ConnectorProfileType.Rectangular or connector.Shape == ConnectorProfileType.Oval:
        return connector.Width * connector.Height
    return 0.0


def set_additional_flow_value(document, target_value):
    updated_count = 0

    for element in collect_additional_flow_elements(document):
        # LookupParameter по имени из ParameterElement — без перегрузок get_Parameter.
        name = _bip_to_lookup_name(document, BuiltInParameter.RBS_ADDITIONAL_FLOW)
        if name:
            parameter = element.LookupParameter(name)
        else:
            parameter = get_parameter_by_names(element, "Additional Flow", u"Доп. расход")

        if not is_writable(parameter) or parameter.StorageType != StorageType.Double:
            continue

        current_value = parameter.AsDouble()
        if nearly_equal(current_value, target_value, 1e-9):
            continue

        parameter.Set(target_value)
        updated_count += 1

    return updated_count
