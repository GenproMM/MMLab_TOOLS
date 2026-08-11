#! python3
# -*- coding: utf-8 -*-

import clr

clr.AddReference("RevitAPIUI")

from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI import TaskDialogCommonButtons


def alert(message, title=u"Сообщение"):
    dialog = TaskDialog(title)
    dialog.MainContent = message
    dialog.CommonButtons = TaskDialogCommonButtons.Ok
    dialog.Show()