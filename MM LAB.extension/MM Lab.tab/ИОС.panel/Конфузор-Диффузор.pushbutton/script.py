# coding: utf-8
"""Общие функции для кнопок GenproMEP в pyRevit (IronPython)."""

import math

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import Domain
from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import FlowDirectionType
from Autodesk.Revit.DB import MEPCurve
from Autodesk.Revit.DB import StorageType
from Autodesk.Revit.DB import ConnectorProfileType
from Autodesk.Revit.DB.ExternalService import ExternalServiceRegistry
from Autodesk.Revit.DB.ExternalService import ExternalServices
from Autodesk.Revit.DB.Mechanical import DuctSystemType
from Autodesk.Revit.UI import TaskDialog
from System.Drawing import Point
from System.Drawing import Size
from System.Windows.Forms import AnchorStyles
from System.Windows.Forms import Button
from System.Windows.Forms import DataGridView
from System.Windows.Forms import DataGridViewAutoSizeColumnsMode
from System.Windows.Forms import DataGridViewComboBoxColumn
from System.Windows.Forms import DataGridViewEditMode
from System.Windows.Forms import DataGridViewSelectionMode
from System.Windows.Forms import DataGridViewTextBoxColumn
from System.Windows.Forms import DialogResult
from System.Windows.Forms import DockStyle
from System.Windows.Forms import Form
from System.Windows.Forms import FormBorderStyle
from System.Windows.Forms import FormStartPosition
from System.Windows.Forms import Panel


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


def get_parameter(element, built_in_parameter, *names):
    parameter = element.get_Parameter(built_in_parameter)
    if parameter is not None:
        return parameter

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
    seen = set()

    candidates = [
        element.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_PARAM),
        element.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM),
        element.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SETTINGS),
        element.LookupParameter(u"Loss Method"),
        element.LookupParameter(u"Метод определения потерь"),
        element.LookupParameter(u"Loss Method Settings"),
        element.LookupParameter(u"Настройки метода определения потерь"),
    ]

    for parameter in candidates:
        if parameter is None:
            continue

        key = u"{0}|{1}|{2}".format(
            to_text(parameter.Definition.Name if parameter.Definition else u""),
            to_text(parameter.StorageType),
            element_id_value(parameter.Id),
        )
        if key in seen:
            continue

        seen.add(key)
        parameters.append(parameter)

    return parameters


def get_duct_loss_method_not_defined_server_id():
    service = ExternalServiceRegistry.GetService(
        ExternalServices.BuiltInExternalServices.DuctFittingAndAccessoryPressureDropService
    )
    if service is None:
        return None

    for server_id in service.GetRegisteredServerIds():
        server = service.GetServer(server_id)
        if server is None:
            continue

        if normalize_text(server.GetName()) in UNDEFINED_SERVER_LABELS:
            return to_text(server.GetServerId())

    return None


def ensure_loss_method_undefined(family_instance, not_defined_server_id):
    handled_server_parameter = False
    server_parameter = family_instance.get_Parameter(BuiltInParameter.RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM)

    if is_writable(server_parameter):
        handled_server_parameter = True
        current_value = to_text(server_parameter.AsString())
        if current_value.lower() == to_text(not_defined_server_id).lower() or is_parameter_undefined(server_parameter):
            return "already"

        if not_defined_server_id and server_parameter.Set(not_defined_server_id):
            return "updated"

    found_any = False
    changed_any = False

    for parameter in get_loss_method_parameters(family_instance):
        found_any = True
        if is_parameter_undefined(parameter):
            continue

        if try_set_parameter_to_undefined(parameter):
            changed_any = True

    if not found_any:
        if handled_server_parameter:
            return "already"
        return "skipped"

    if changed_any:
        return "updated"

    for parameter in get_loss_method_parameters(family_instance):
        if is_parameter_undefined(parameter):
            return "already"

    return "skipped"


def get_default_supply_flag_decision(classification):
    normalized = normalize_text(classification)
    if normalized in SUPPLY_LABELS:
        return SupplyFlagDecision.SUPPLY
    if normalized in NOT_SUPPLY_LABELS:
        return SupplyFlagDecision.NOT_SUPPLY
    return SupplyFlagDecision.IGNORE


def get_connector_manager(element):
    try:
        if getattr(element, "MEPModel", None) is not None and element.MEPModel.ConnectorManager is not None:
            return element.MEPModel.ConnectorManager
    except Exception:
        pass

    if isinstance(element, MEPCurve):
        return element.ConnectorManager

    return None


def get_hvac_connectors(element):
    result = []
    connector_manager = get_connector_manager(element)
    if connector_manager is None:
        return result

    for connector in connector_manager.Connectors:
        if connector.Domain == Domain.DomainHvac:
            result.append(connector)

    return result


def get_system_classification(element):
    classification = get_parameter_text(
        get_parameter(
            element,
            BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM,
            u"System Classification",
            u"Классификация системы",
        )
    )
    if classification:
        return classification

    classification = get_parameter_text(
        get_parameter(
            element,
            BuiltInParameter.RBS_DUCT_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM,
            u"System Classification",
            u"Классификация системы",
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
        parameter = element.get_Parameter(BuiltInParameter.RBS_ADDITIONAL_FLOW)
        if not is_writable(parameter) or parameter.StorageType != StorageType.Double:
            continue

        current_value = parameter.AsDouble()
        if nearly_equal(current_value, target_value, 1e-9):
            continue

        parameter.Set(target_value)
        updated_count += 1

    return updated_count


class SystemClassificationSelectionForm(Form):
    def __init__(self, options):
        Form.__init__(self)
        self._options = options
        self._grid = None
        self._build_ui()

    def _build_ui(self):
        self.Text = u"Настройка классификаций систем"
        self.ClientSize = Size(820, 560)
        self.MinimumSize = Size(820, 560)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.ShowInTaskbar = False

        grid = DataGridView()
        grid.Location = Point(12, 12)
        grid.Size = Size(796, 468)
        grid.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.MultiSelect = False
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grid.EditMode = DataGridViewEditMode.EditOnEnter
        grid.RowHeadersVisible = False
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        classification_column = DataGridViewTextBoxColumn()
        classification_column.Name = "classification"
        classification_column.HeaderText = u"Классификация системы"
        classification_column.ReadOnly = True

        decision_column = DataGridViewComboBoxColumn()
        decision_column.Name = "decision"
        decision_column.HeaderText = u"Как установить параметр \"Приточный\""
        for label in [u"Приток", u"Не приток", u"Не менять"]:
            decision_column.Items.Add(label)

        grid.Columns.Add(classification_column)
        grid.Columns.Add(decision_column)

        for option in self._options:
            index = grid.Rows.Add()
            grid.Rows[index].Cells[0].Value = option["classification"]
            grid.Rows[index].Cells[1].Value = DECISION_TO_LABEL[option["decision"]]

        button_panel = Panel()
        button_panel.Location = Point(0, 490)
        button_panel.Size = Size(820, 70)
        button_panel.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom

        ok_button = Button()
        ok_button.Text = u"Применить"
        ok_button.DialogResult = DialogResult.OK
        ok_button.Size = Size(140, 34)
        ok_button.Location = Point(536, 18)
        ok_button.Anchor = AnchorStyles.Right | AnchorStyles.Top

        cancel_button = Button()
        cancel_button.Text = u"Отмена"
        cancel_button.DialogResult = DialogResult.Cancel
        cancel_button.Size = Size(120, 34)
        cancel_button.Location = Point(688, 18)
        cancel_button.Anchor = AnchorStyles.Right | AnchorStyles.Top

        button_panel.Controls.Add(ok_button)
        button_panel.Controls.Add(cancel_button)

        self.AcceptButton = ok_button
        self.CancelButton = cancel_button

        self.Controls.Add(button_panel)
        self.Controls.Add(grid)

        self._grid = grid

    def get_selections(self):
        try:
            self._grid.EndEdit()
        except Exception:
            pass

        selections = []
        for row in self._grid.Rows:
            classification = to_text(row.Cells[0].Value)
            label = to_text(row.Cells[1].Value) or DECISION_TO_LABEL[SupplyFlagDecision.IGNORE]
            decision = LABEL_TO_DECISION.get(label, SupplyFlagDecision.IGNORE)
            selections.append({
                "classification": classification,
                "decision": decision,
            })
        return selections


# coding: utf-8
"""Определяет, является ли переход конфузором или диффузором."""


from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import FlowDirectionType
from Autodesk.Revit.DB import StorageType
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog




COMMAND_NAME = u"Конфузор/Диффузор"
AREA_TOLERANCE = 1e-6


try:
    doc = get_document(COMMAND_NAME)
    if doc:
        confuser_count = 0
        diffuser_count = 0
        skipped_count = 0

        fittings = collect_elements(doc, BuiltInCategory.OST_DuctFitting)

        transaction = Transaction(doc, COMMAND_NAME)
        transaction.Start()
        try:
            for element in fittings:
                family_instance = element if hasattr(element, "MEPModel") else None
                if family_instance is None or family_instance.MEPModel is None:
                    skipped_count += 1
                    continue

                try:
                    part_type = family_instance.MEPModel.PartType
                except Exception:
                    skipped_count += 1
                    continue

                if normalize_text(to_text(part_type)) != "transition":
                    skipped_count += 1
                    continue

                hvac_connectors = get_hvac_connectors(family_instance)
                if len(hvac_connectors) != 2:
                    skipped_count += 1
                    continue

                in_connector = None
                out_connector = None
                for connector in hvac_connectors:
                    if connector.Direction == FlowDirectionType.In:
                        in_connector = connector
                    elif connector.Direction == FlowDirectionType.Out:
                        out_connector = connector

                if in_connector is None or out_connector is None:
                    skipped_count += 1
                    continue

                in_area = get_connector_area(in_connector)
                out_area = get_connector_area(out_connector)
                if in_area <= 0.0 or out_area <= 0.0 or nearly_equal(in_area, out_area, AREA_TOLERANCE):
                    skipped_count += 1
                    continue

                confuser_parameter = get_parameter_by_names(family_instance, u"Конфузор", u"Confuser")
                if not is_writable(confuser_parameter) or confuser_parameter.StorageType != StorageType.Integer:
                    skipped_count += 1
                    continue

                is_confuser = in_area > out_area
                new_value = 1 if is_confuser else 0
                if confuser_parameter.AsInteger() != new_value:
                    confuser_parameter.Set(new_value)

                if is_confuser:
                    confuser_count += 1
                else:
                    diffuser_count += 1

            transaction.Commit()
        except Exception:
            transaction.RollBack()
            raise

        total = confuser_count + diffuser_count
        TaskDialog.Show(
            COMMAND_NAME,
            u"Классифицировано переходов: {0}\nКонфузоров: {1}\nДиффузоров: {2}\nПропущено: {3}".format(
                total,
                confuser_count,
                diffuser_count,
                skipped_count,
            )
        )
except Exception as ex:
    show_error(COMMAND_NAME, ex)
