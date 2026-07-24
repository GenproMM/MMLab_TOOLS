import os
import sys
import requests
from Autodesk.Revit.DB import *
from pyrevit import forms

EXTENSION_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
sys.path.append(os.path.join(EXTENSION_ROOT, "lib"))


def get_room_value(element):
    try:
        param = element.LookupParameter("GP_23_Назначение")
        if param and param.HasValue:
            return param.AsString()
    except:
        pass
    return ""
