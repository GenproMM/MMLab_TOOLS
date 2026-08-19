#! python3
# -*- coding: utf-8 -*-
"""Замена CAD-геометрии

Заменяет DWG-импорты в семействе на лёгкие нативные боксы, сохраняя
коннекторы (тип, профиль, параметры потока и системы) на новой геометрии.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Замена\nCAD-геометрии"
__author__ = "Артём Брежнев"

# Канонический lib-бутстрап (D-15).
import os
import sys

_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton -> panel -> tab -> MM_LAB.extension
_EXTENSION_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")
if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")
clr.AddReference("System.Core")

from Autodesk.Revit.DB import (
    XYZ, Plane, SketchPlane, Line, CurveArray, CurveArrArray,
    FilteredElementCollector, ImportInstance, ConnectorElement,
    ConnectorProfileType, Domain, PlanarFace, Solid, GeometryInstance,
    Mesh, ElementId, Options, ViewDetailLevel, Transaction,
    BuiltInParameter, CADLinkType, PerformanceAdviser
)
from Autodesk.Revit.DB.Mechanical import DuctSystemType
from Autodesk.Revit.DB.Plumbing import PipeSystemType
from Autodesk.Revit.UI import TaskDialog
import System
from System.Collections.Generic import List, HashSet

import revit_compat


COMMAND_NAME = u"Замена CAD-геометрии"


def safe_len(coll):
    try:
        return len(coll)
    except Exception:
        try:
            return coll.Count
        except Exception:
            return 0


def mesh_bbox(mesh):
    if mesh.NumberOfVertices == 0: return None, None
    BIG = 1e9
    mn = XYZ( BIG,  BIG,  BIG)
    mx = XYZ(-BIG, -BIG, -BIG)
    for v in mesh.Vertices:
        mn = XYZ(min(mn.X, v.X), min(mn.Y, v.Y), min(mn.Z, v.Z))
        mx = XYZ(max(mx.X, v.X), max(mx.Y, v.Y), max(mx.Z, v.Z))
    return mn, mx


def bbox_inside(inner, outer, tol=0.005):
    imn, imx = inner; omn, omx = outer
    return (omn.X-tol <= imn.X and omn.Y-tol <= imn.Y and omn.Z-tol <= imn.Z
        and omx.X+tol >= imx.X and omx.Y+tol >= imx.Y and omx.Z+tol >= imx.Z)


def create_box(doc, mn, mx):
    dz = max(mx.Z - mn.Z, 0.01)
    sp = SketchPlane.Create(doc,
        Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0, 0, mn.Z)))
    p0 = XYZ(mn.X, mn.Y, mn.Z); p1 = XYZ(mx.X, mn.Y, mn.Z)
    p2 = XYZ(mx.X, mx.Y, mn.Z); p3 = XYZ(mn.X, mx.Y, mn.Z)
    lp = CurveArray()
    lp.Append(Line.CreateBound(p0, p1)); lp.Append(Line.CreateBound(p1, p2))
    lp.Append(Line.CreateBound(p2, p3)); lp.Append(Line.CreateBound(p3, p0))
    arr = CurveArrArray(); arr.Append(lp)
    return doc.FamilyCreate.NewExtrusion(True, arr, sp, dz)


def create_tiny(doc, origin, bz, bx):
    HALF, DEPTH = 0.08, -0.003
    if bx is None or bx.GetLength() < 0.001:
        ref = XYZ.BasisX if abs(bz.X) < 0.9 else XYZ.BasisY
        bx = bz.CrossProduct(ref).Normalize()
    by = bz.CrossProduct(bx).Normalize()
    bx = by.CrossProduct(bz).Normalize()
    hx = bx.Multiply(HALF); hy = by.Multiply(HALF)
    c0 = origin.Subtract(hx).Subtract(hy)
    c1 = origin.Add(hx).Subtract(hy)
    c2 = origin.Add(hx).Add(hy)
    c3 = origin.Subtract(hx).Add(hy)
    sp = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(bz, origin))
    lp = CurveArray()
    lp.Append(Line.CreateBound(c0, c1)); lp.Append(Line.CreateBound(c1, c2))
    lp.Append(Line.CreateBound(c2, c3)); lp.Append(Line.CreateBound(c3, c0))
    arr = CurveArrArray(); arr.Append(lp)
    return doc.FamilyCreate.NewExtrusion(True, arr, sp, DEPTH)


def find_face_ref(elem, target_n, opts):
    best_ref, best_dot = None, -2.0
    for obj in elem.get_Geometry(opts):
        if isinstance(obj, Solid):
            for f in obj.Faces:
                if isinstance(f, PlanarFace):
                    dot = f.FaceNormal.DotProduct(target_n)
                    if dot > best_dot:
                        best_dot, best_ref = dot, f.Reference
    return best_ref, best_dot


def delete_import_categories(doc):
    deleted = 0
    try:
        cats = doc.Settings.Categories
        import_cats = []
        for c in cats:
            try:
                if c.SubCategories is not None:
                    for sub in c.SubCategories:
                        import_cats.append(sub.Id)
            except Exception: pass
        for cid in import_cats:
            try:
                if doc.GetElement(cid) is not None:
                    doc.Delete(cid)
                    deleted += 1
            except Exception: pass
    except Exception: pass
    return deleted


def purge_unused_native(doc):
    if not hasattr(doc, "GetUnusedElements"):
        return None
    total = 0
    iters = []
    empty_set = HashSet[ElementId]()
    for i in range(8):
        try:
            unused = doc.GetUnusedElements(empty_set)
            n = safe_len(unused)
            if n == 0:
                iters.append("итер{}: 0".format(i+1)); break
            try:
                deleted = doc.Delete(unused)
                cnt = safe_len(deleted)
            except Exception:
                cnt = 0
                for eid in unused:
                    try:
                        doc.Delete(eid); cnt += 1
                    except Exception: pass
            total += cnt
            iters.append("итер{}: {}".format(i+1, cnt))
            if cnt == 0: break
        except Exception as e:
            iters.append("итер{} err: {}".format(i+1, str(e)[:60])); break
    return total, " / ".join(iters)


def purge_unused_adviser(doc, max_iter=6):
    PURGE_GUID = 'e8c63650-70b7-435a-9010-ec97660c1bda'
    try:
        adviser = PerformanceAdviser.GetPerformanceAdviser()
    except Exception: return 0, "no adviser"
    target_guid = System.Guid(PURGE_GUID)
    rule_id = None
    for r in adviser.GetAllRuleIds():
        if r.Guid.Equals(target_guid):
            rule_id = r; break
    if rule_id is None: return 0, "no rule"
    rule_ids = List[type(rule_id)]()
    rule_ids.Add(rule_id)
    total = 0; iters = []
    for i in range(max_iter):
        try:
            failures = adviser.ExecuteRules(doc, rule_ids)
            if safe_len(failures) == 0:
                iters.append("итер{}: 0".format(i+1)); break
            ids = failures[0].GetFailingElements()
            if safe_len(ids) == 0:
                iters.append("итер{}: 0".format(i+1)); break
            try:
                deleted = doc.Delete(ids)
                cnt = safe_len(deleted)
            except Exception:
                cnt = 0
                for eid in ids:
                    try: doc.Delete(eid); cnt += 1
                    except Exception: pass
            total += cnt
            iters.append("итер{}: {}".format(i+1, cnt))
        except Exception as e:
            iters.append("итер{} err: {}".format(i+1, str(e)[:60])); break
    return total, " / ".join(iters)


def main(doc):
    """Точка входа: заменяет DWG-импорты семейства на нативные боксы."""
    revit_compat.require_supported_version(COMMAND_NAME)

    if not doc.IsFamilyDocument:
        TaskDialog.Show(COMMAND_NAME, "Открой .rfa файл в редакторе семейств")
        return

    log = []

    saved = []
    for ce in FilteredElementCollector(doc).OfClass(ConnectorElement).ToElements():
        try:
            origin = ce.Origin
            cs = ce.CoordinateSystem
            d = {"id": ce.Id, "origin": origin,
                 "basisZ": cs.BasisZ, "basisX": cs.BasisX}
            try:    d["domain"]  = ce.Domain
            except Exception: d["domain"]  = None
            try:    d["profile"] = ce.Shape
            except Exception: d["profile"] = ConnectorProfileType.Round

            d["radius"], d["width"], d["height"] = None, None, None
            if d["profile"] == ConnectorProfileType.Round:
                try: d["radius"] = ce.Radius
                except Exception: pass
            elif d["profile"] == ConnectorProfileType.Rectangular:
                try:
                    d["width"]  = ce.Width
                    d["height"] = ce.Height
                except Exception: pass

            d["sys_type_int"] = None
            try:
                if d["domain"] == Domain.DomainHvac:
                    p = revit_compat.get_parameter(ce, BuiltInParameter.RBS_DUCT_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM)
                    if p is None or not p.HasValue:
                        p = revit_compat.get_parameter(ce, BuiltInParameter.RBS_DUCT_CONNECTOR_SYSTEM_TYPE_PARAM)
                    if p and p.HasValue: d["sys_type_int"] = p.AsInteger()
                elif d["domain"] == Domain.DomainPiping:
                    p = revit_compat.get_parameter(ce, BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM)
                    if p is None or not p.HasValue:
                        p = revit_compat.get_parameter(ce, BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_TYPE_PARAM)
                    if p and p.HasValue: d["sys_type_int"] = p.AsInteger()
            except Exception: pass

            d["flow"] = None
            try:
                p = revit_compat.get_parameter(ce, BuiltInParameter.RBS_DUCT_CONNECTOR_FLOW_DIRECTION_PARAM)
                if p is None or not p.HasValue:
                    p = revit_compat.get_parameter(ce, BuiltInParameter.RBS_PIPE_CONNECTOR_FLOW_DIRECTION_PARAM)
                if p and p.HasValue: d["flow"] = p.AsInteger()
            except Exception: pass

            saved.append(d)
        except Exception as e:
            log.append("read: " + str(e))
    log.append("Считано коннекторов: {}".format(len(saved)))

    imports = list(FilteredElementCollector(doc).OfClass(ImportInstance).ToElements())

    if not imports:
        TaskDialog.Show(COMMAND_NAME,
            "DWG-импортов не найдено.\nКоннекторов: {}".format(len(saved)))
        return

    opts = Options()
    opts.ComputeReferences = True
    opts.DetailLevel = ViewDetailLevel.Fine

    bboxes = []
    for imp in imports:
        try:
            for obj in imp.get_Geometry(opts):
                if isinstance(obj, GeometryInstance):
                    for item in obj.GetInstanceGeometry():
                        if isinstance(item, Mesh):
                            mn, mx = mesh_bbox(item)
                            if mn and mx: bboxes.append((mn, mx))
        except Exception as e:
            log.append("DWG geom: " + str(e))

    if not bboxes:
        for imp in imports:
            bb = imp.get_BoundingBox(None)
            if bb: bboxes.append((bb.Min, bb.Max))
        log.append("Fallback bbox: {}".format(len(bboxes)))
    else:
        log.append("mesh-кубиков: {}".format(len(bboxes)))

    before = len(bboxes)
    filtered = []
    for i, b in enumerate(bboxes):
        inside_other = False
        for j, other in enumerate(bboxes):
            if i == j: continue
            if bbox_inside(b, other):
                if i > j or not bbox_inside(other, b):
                    inside_other = True; break
        if not inside_other:
            filtered.append(b)
    bboxes = filtered
    log.append("После фильтра вложенных: {} (было {})".format(len(bboxes), before))

    boxes_made = 0
    tiny_ids = []
    new_count = 0
    dwg_del = 0
    type_del = 0
    cat_del = 0
    purged_count = 0
    purge_log = ""

    try:
        # ─── TX1: боксы + tiny ───
        t1 = Transaction(doc, "Замена CAD: боксы и tiny")
        t1.Start()
        try:
            for mn, mx in bboxes:
                try:
                    create_box(doc, mn, mx)
                    boxes_made += 1
                except Exception as e:
                    log.append("box: " + str(e))

            for s in saved:
                try:
                    t = create_tiny(doc, s["origin"], s["basisZ"], s["basisX"])
                    tiny_ids.append(t.Id)
                except Exception as e:
                    tiny_ids.append(None)
                    log.append("tiny: " + str(e))
            t1.Commit()
        except Exception:
            if t1.HasStarted(): t1.RollBack()
            raise

        log.append("Боксов: {}, tiny: {}".format(
            boxes_made, sum(1 for x in tiny_ids if x is not None)))

        # ─── TX2: коннекторы + удалить старые + DWG + типы ───
        t2 = Transaction(doc, "Замена CAD: коннекторы и удаление DWG")
        t2.Start()
        try:
            for i, s in enumerate(saved):
                if i >= len(tiny_ids) or tiny_ids[i] is None: continue
                try:
                    tiny = doc.GetElement(tiny_ids[i])
                    if tiny is None: continue

                    face_ref, dot = find_face_ref(tiny, s["basisZ"], opts)
                    if face_ref is None or dot < 0.5:
                        log.append("conn[{}]: dot={:.2f}".format(i, dot)); continue

                    profile = s["profile"] or ConnectorProfileType.Round
                    domain  = s["domain"]
                    ce_new  = None

                    if domain == Domain.DomainHvac:
                        ce_new = ConnectorElement.CreateDuctConnector(
                            doc, DuctSystemType.SupplyAir, profile, face_ref)
                    elif domain == Domain.DomainPiping:
                        ce_new = ConnectorElement.CreatePipeConnector(
                            doc, PipeSystemType.SupplyHydronic, face_ref)
                    else:
                        ce_new = ConnectorElement.CreateDuctConnector(
                            doc, DuctSystemType.SupplyAir,
                            ConnectorProfileType.Round, face_ref)

                    if s["sys_type_int"] is not None:
                        try:
                            if domain == Domain.DomainHvac:
                                p = revit_compat.get_parameter(ce_new, BuiltInParameter.RBS_DUCT_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM)
                                if p is None or p.IsReadOnly:
                                    p = revit_compat.get_parameter(ce_new, BuiltInParameter.RBS_DUCT_CONNECTOR_SYSTEM_TYPE_PARAM)
                                if p and not p.IsReadOnly:
                                    p.Set(s["sys_type_int"])
                            elif domain == Domain.DomainPiping:
                                p = revit_compat.get_parameter(ce_new, BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_CLASSIFICATION_PARAM)
                                if p is None or p.IsReadOnly:
                                    p = revit_compat.get_parameter(ce_new, BuiltInParameter.RBS_PIPE_CONNECTOR_SYSTEM_TYPE_PARAM)
                                if p and not p.IsReadOnly:
                                    p.Set(s["sys_type_int"])
                        except Exception as e:
                            log.append("sys_type[{}]: {}".format(i, str(e)[:60]))

                    try:
                        if s["radius"]:
                            p = revit_compat.get_parameter(ce_new, BuiltInParameter.CONNECTOR_RADIUS)
                            if p and not p.IsReadOnly: p.Set(s["radius"])
                    except Exception: pass
                    try:
                        if s["width"]:
                            p = revit_compat.get_parameter(ce_new, BuiltInParameter.CONNECTOR_WIDTH)
                            if p and not p.IsReadOnly: p.Set(s["width"])
                        if s["height"]:
                            p = revit_compat.get_parameter(ce_new, BuiltInParameter.CONNECTOR_HEIGHT)
                            if p and not p.IsReadOnly: p.Set(s["height"])
                    except Exception: pass

                    if s["flow"] is not None:
                        try:
                            p = revit_compat.get_parameter(ce_new, BuiltInParameter.RBS_DUCT_CONNECTOR_FLOW_DIRECTION_PARAM)
                            if p is None or p.IsReadOnly:
                                p = revit_compat.get_parameter(ce_new, BuiltInParameter.RBS_PIPE_CONNECTOR_FLOW_DIRECTION_PARAM)
                            if p and not p.IsReadOnly:
                                p.Set(s["flow"])
                        except Exception: pass

                    new_count += 1
                except Exception as e:
                    log.append("create[{}]: {}".format(i, str(e)))

            for s in saved:
                try:
                    if doc.GetElement(s["id"]) is not None:
                        doc.Delete(s["id"])
                except Exception: pass

            cad_type_ids = set()
            for imp in imports:
                try:
                    tid = imp.GetTypeId()
                    if tid != ElementId.InvalidElementId:
                        cad_type_ids.add(tid)
                except Exception: pass

            for imp in imports:
                try:
                    doc.Delete(imp.Id)
                    dwg_del += 1
                except Exception as e:
                    log.append("del DWG: " + str(e))

            for tid in cad_type_ids:
                try:
                    if doc.GetElement(tid) is not None:
                        doc.Delete(tid)
                        type_del += 1
                except Exception as e:
                    log.append("del CADtype: " + str(e))

            try:
                for ct in FilteredElementCollector(doc).OfClass(CADLinkType).ToElements():
                    try:
                        doc.Delete(ct.Id)
                        type_del += 1
                    except Exception: pass
            except Exception: pass

            cat_del = delete_import_categories(doc)
            log.append("Удалено категорий: {}".format(cat_del))

            t2.Commit()
        except Exception:
            if t2.HasStarted(): t2.RollBack()
            raise

        # ─── TX3: PURGE UNUSED ───
        t3 = Transaction(doc, "Замена CAD: Purge Unused")
        t3.Start()
        try:
            native_result = purge_unused_native(doc)
            if native_result is not None:
                purged_count, purge_log = native_result
                purge_log = "GetUnusedElements: " + purge_log
            else:
                purged_count, purge_log = purge_unused_adviser(doc, max_iter=6)
                purge_log = "PerformanceAdviser: " + purge_log

            if native_result is not None:
                extra, extra_log = purge_unused_adviser(doc, max_iter=3)
                if extra > 0:
                    purged_count += extra
                    purge_log += " || PerformanceAdviser: " + extra_log

            t3.Commit()
        except Exception:
            if t3.HasStarted(): t3.RollBack()
            raise

    except Exception as e:
        TaskDialog.Show(COMMAND_NAME + ": ошибка", "Произошла ошибка: " + str(e))
    else:
        summary = (
            "Коннекторов было:    {}\n"
            "Коннекторов создано: {}\n"
            "Кубиков:             {}\n"
            "DWG инстансов:       {}\n"
            "CAD типов:           {}\n"
            "Импорт-категорий:    {}\n"
            "Purge всего удалено: {}\n"
            "Purge итерации:      {}\n\n"
            "Не забудь 'Сохранить как' чтобы файл реально похудел."
        ).format(len(saved), new_count, boxes_made, dwg_del,
                 type_del, cat_del, purged_count, purge_log)
        TaskDialog.Show(COMMAND_NAME + ": готово", summary)


def _entry():
    """Готовит doc/uidoc и вызывает main (doc/uidoc — параметрами, правило 18)."""
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None or uidoc.Document is None:
        TaskDialog.Show(COMMAND_NAME, u"Открой проект Revit и повтори команду.")
        return
    main(uidoc.Document)


try:
    _entry()
except SystemExit:
    pass  # require_supported_version уже показал свой диалог
except Exception as ex:
    TaskDialog.Show(COMMAND_NAME, u"Ошибка:\n{0}".format(ex))
