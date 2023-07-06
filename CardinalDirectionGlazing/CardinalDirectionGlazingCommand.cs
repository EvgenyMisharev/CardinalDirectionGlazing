using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CardinalDirectionGlazing
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CardinalDirectionGlazingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Selection sel = commandData.Application.ActiveUIDocument.Selection;
            Document linkDoc = null;

            Guid windowsAreaNorthGuid = new Guid("820af414-f6ec-472d-887c-a2046a0c5988");
            Guid windowsAreaSouthGuid = new Guid("81ab8e02-45c6-4d26-b0e5-a6736b0c352d");

            Guid windowsAreaWestGuid = new Guid("65fe3416-f836-48ff-bed9-3fdf2126a1f9");
            Guid windowsAreaEastGuid = new Guid("fc33c487-9bbb-43d6-a7f4-aba5f9638fe3");

            Guid windowsAreaNorthwestGuid = new Guid("f78f8a53-cea7-4e00-955c-3748aa7a37c7");
            Guid windowsAreaNortheastGuid = new Guid("b8120c53-0793-4932-bc71-845302914573");

            Guid windowsAreaSouthwestGuid = new Guid("3ff1f178-2cff-4b54-a0d3-eee58fa1622c");
            Guid windowsAreaSoutheastGuid = new Guid("c5e261ae-68f5-4a91-a55f-8686d278f5ab");

            List<Space> spaceList = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .WhereElementIsNotElementType()
                .Cast<Space>()
                .Where(s => s.Area > 0)
                .ToList();

            //Проверка наличия параметров
            if(spaceList.Count != 0)
            {
                Parameter windowsAreaNorthParameter = spaceList.First().get_Parameter(windowsAreaNorthGuid);
                if(windowsAreaNorthParameter == null)
                {
                    TaskDialog.Show("Revit", "Параметр пространства \"ПлощадьОкон_С\" не найден! Добавьте параметр в соответствии с инструкцией!");
                    return Result.Cancelled;
                }

                Parameter windowsAreaSouthParameter = spaceList.First().get_Parameter(windowsAreaSouthGuid);
                if (windowsAreaSouthParameter == null)
                {
                    TaskDialog.Show("Revit", "Параметр пространства \"ПлощадьОкон_Ю\" не найден! Добавьте параметр в соответствии с инструкцией!");
                    return Result.Cancelled;
                }

                Parameter windowsAreaWestParameter = spaceList.First().get_Parameter(windowsAreaWestGuid);
                if (windowsAreaWestParameter == null)
                {
                    TaskDialog.Show("Revit", "Параметр пространства \"ПлощадьОкон_З\" не найден! Добавьте параметр в соответствии с инструкцией!");
                    return Result.Cancelled;
                }

                Parameter windowsAreaEastParameter = spaceList.First().get_Parameter(windowsAreaEastGuid);
                if (windowsAreaEastParameter == null)
                {
                    TaskDialog.Show("Revit", "Параметр пространства \"ПлощадьОкон_В\" не найден! Добавьте параметр в соответствии с инструкцией!");
                    return Result.Cancelled;
                }

                Parameter windowsAreaNorthwestParameter = spaceList.First().get_Parameter(windowsAreaNorthwestGuid);
                if (windowsAreaNorthwestParameter == null)
                {
                    TaskDialog.Show("Revit", "Параметр пространства \"ПлощадьОкон_СЗ\" не найден! Добавьте параметр в соответствии с инструкцией!");
                    return Result.Cancelled;
                }

                Parameter windowsAreaNortheastParameter = spaceList.First().get_Parameter(windowsAreaNortheastGuid);
                if (windowsAreaNortheastParameter == null)
                {
                    TaskDialog.Show("Revit", "Параметр пространства \"ПлощадьОкон_СВ\" не найден! Добавьте параметр в соответствии с инструкцией!");
                    return Result.Cancelled;
                }

                Parameter windowsAreaSouthwestParameter = spaceList.First().get_Parameter(windowsAreaSouthwestGuid);
                if (windowsAreaSouthwestParameter == null)
                {
                    TaskDialog.Show("Revit", "Параметр пространства \"ПлощадьОкон_ЮЗ\" не найден! Добавьте параметр в соответствии с инструкцией!");
                    return Result.Cancelled;
                }

                Parameter windowsAreaSoutheastParameter = spaceList.First().get_Parameter(windowsAreaSoutheastGuid);
                if (windowsAreaSoutheastParameter == null)
                {
                    TaskDialog.Show("Revit", "Параметр пространства \"ПлощадьОкон_ЮВ\" не найден! Добавьте параметр в соответствии с инструкцией!");
                    return Result.Cancelled;
                }
            }

            List<RevitLinkInstance> revitLinkInstanceList = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();
            if(revitLinkInstanceList.Count == 0)
            {
                TaskDialog.Show("Ravit", "В проекте отсутствуют связанные файлы!");
                return Result.Cancelled;
            }

            CardinalDirectionGlazingWPF cardinalDirectionGlazingWPF = new CardinalDirectionGlazingWPF(revitLinkInstanceList);
            cardinalDirectionGlazingWPF.ShowDialog();
            if (cardinalDirectionGlazingWPF.DialogResult != true)
            {
                return Result.Cancelled;
            }

            if (cardinalDirectionGlazingWPF.SelectedRevitLinkInstance == null)
            {
                TaskDialog.Show("Ravit", "Связанный файл не выбран!");
                return Result.Cancelled;
            }

            linkDoc = cardinalDirectionGlazingWPF.SelectedRevitLinkInstance.GetLinkDocument();
            Transform transform = cardinalDirectionGlazingWPF.SelectedRevitLinkInstance.GetTotalTransform();
            ProjectPosition position = linkDoc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);
            Transform trueNorthTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, -position.Angle, XYZ.Zero);
            XYZ trueNorthBasisY = trueNorthTransform.OfVector(XYZ.BasisY);
            XYZ trueNorthBasisX = trueNorthTransform.OfVector(XYZ.BasisX);

            string spacesForProcessingButtonName = cardinalDirectionGlazingWPF.SpacesForProcessingButtonName;
            if(spacesForProcessingButtonName == "radioButton_Selected")
            {
                spaceList = new List<Space>();
                SpaceSelectionFilter spaceSelectionFilter = new SpaceSelectionFilter();
                IList<Reference> selSpaces = null;
                spaceList = GetSpacesFromCurrentSelection(doc, sel);
                if (spaceList.Count == 0)
                {
                    try
                    {
                        selSpaces = sel.PickObjects(ObjectType.Element, spaceSelectionFilter, "Выберите пространства!");
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        return Result.Cancelled;
                    }

                    foreach (Reference roomRef in selSpaces)
                    {
                        spaceList.Add(doc.GetElement(roomRef) as Space);
                    }
                }
            }

            List<FamilyInstance> windowsList = new FilteredElementCollector(linkDoc)
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .ToList();

            List<Wall> curtainWallsList = new FilteredElementCollector(linkDoc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .Where(w => w.CurtainGrid != null)
                .ToList();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Остекление по сторонам света");
                foreach (Space space in spaceList)
                {
                    double windowsAreaNorth = 0;
                    double windowsAreaSouth = 0;

                    double windowsAreaWest = 0;
                    double windowsAreaEast = 0;

                    double windowsAreaNorthwest = 0;
                    double windowsAreaNortheast = 0;

                    double windowsAreaSouthwest = 0;
                    double windowsAreaSoutheast = 0;

                    Solid spaceSolid = null;
                    SolidCurveIntersectionOptions intersectOptions = new SolidCurveIntersectionOptions();

                    GeometryElement geomRoomElement = space.get_Geometry(new Options());
                    foreach (GeometryObject geomObj in geomRoomElement)
                    {
                        spaceSolid = geomObj as Solid;
                        if (spaceSolid != null) break;
                    }
                    if (spaceSolid != null)
                    {
                        foreach (FamilyInstance window in windowsList)
                        {
                            bool flag = false;
                            BoundingBoxXYZ windowBoundingBox = window.get_BoundingBox(null);
                            if (windowBoundingBox == null) continue;
                            XYZ windowCenter = (windowBoundingBox.Max + windowBoundingBox.Min) / 2;
                            Curve lineA = Line.CreateBound(windowCenter, windowCenter + (700 / 304.8) * window.FacingOrientation.Normalize()) as Curve;
                            Curve lineB = Line.CreateBound(windowCenter, windowCenter + (700 / 304.8) * window.FacingOrientation.Normalize().Negate()) as Curve;
                            Curve transformdlineA = lineA.CreateTransformed(transform) as Curve;
                            Curve transformdlineB = lineB.CreateTransformed(transform) as Curve;
                            SolidCurveIntersection intersection = spaceSolid.IntersectWithCurve(transformdlineB, intersectOptions);
                            if (intersection.SegmentCount > 0)
                            {
                                foreach (Space secondSpace in spaceList)
                                {
                                    if (secondSpace.Id == space.Id) continue;

                                    Solid secondSpaceSolid = null;

                                    GeometryElement secondGeomRoomElement = secondSpace.get_Geometry(new Options());
                                    foreach (GeometryObject secondGeomObj in secondGeomRoomElement)
                                    {
                                        secondSpaceSolid = secondGeomObj as Solid;
                                        if (secondSpaceSolid != null) break;
                                    }
                                    if (secondSpaceSolid != null)
                                    {
                                        SolidCurveIntersection secondIntersection = secondSpaceSolid.IntersectWithCurve(transformdlineA, intersectOptions);
                                        if (secondIntersection.SegmentCount > 0)
                                        {
                                            flag = true;
                                        }
                                    }
                                }
                                if (flag == true) continue;

                                double roughHeight = 0;
                                double roughWidth = 0;
                                double caseworkHeight = 0;
                                double caseworkWidth = 0;
                                double maxHeight = 0;
                                double maxWidth = 0;

                                if (window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM) != null)
                                {
                                    roughHeight = window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM).AsDouble();
                                }
                                if (window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM) != null)
                                {
                                    roughWidth = window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM).AsDouble();
                                }
                                if (window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_HEIGHT) != null)
                                {
                                    caseworkHeight = window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_HEIGHT).AsDouble();
                                }
                                if (window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_WIDTH) != null)
                                {
                                    caseworkWidth = window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_WIDTH).AsDouble();
                                }

                                if (roughHeight >= caseworkHeight)
                                {
                                    maxHeight = roughHeight;
                                }
                                else
                                {
                                    maxHeight = caseworkHeight;
                                }

                                if (roughWidth >= caseworkWidth)
                                {
                                    maxWidth = roughWidth;
                                }
                                else
                                {
                                    maxWidth = caseworkWidth;
                                }

                                double windowArea = maxHeight * maxWidth;

                                XYZ lineDirection = (lineA as Line).Direction;
                                if(lineDirection.AngleTo(trueNorthBasisY) <= Math.PI / 8)
                                {
                                    windowsAreaNorth += windowArea;
                                }
                                else if (lineDirection.AngleTo(trueNorthBasisY.Negate()) <= Math.PI / 8)
                                {
                                    windowsAreaSouth += windowArea;
                                }
                                else if (lineDirection.AngleTo(trueNorthBasisX.Negate()) <= Math.PI / 8)
                                {
                                    windowsAreaWest += windowArea;
                                }
                                else if (lineDirection.AngleTo(trueNorthBasisX) <= Math.PI / 8)
                                {
                                    windowsAreaEast += windowArea;
                                }

                                else if (lineDirection.AngleTo((trueNorthBasisY + trueNorthBasisX.Negate())/2) < Math.PI / 8)
                                {
                                    windowsAreaNorthwest += windowArea;
                                }
                                else if (lineDirection.AngleTo((trueNorthBasisY + trueNorthBasisX) / 2) < Math.PI / 8)
                                {
                                    windowsAreaNortheast += windowArea;
                                }

                                else if (lineDirection.AngleTo((trueNorthBasisY.Negate() + trueNorthBasisX.Negate()) / 2) < Math.PI / 8)
                                {
                                    windowsAreaSouthwest += windowArea;
                                }
                                else if (lineDirection.AngleTo((trueNorthBasisY.Negate() + trueNorthBasisX) / 2) < Math.PI / 8)
                                {
                                    windowsAreaSoutheast += windowArea;
                                }
                            }
                        }

                        foreach (Wall wall in curtainWallsList)
                        {
                            List<ElementId> CurtainPanelsIdList = wall.CurtainGrid.GetPanelIds().ToList();
                            foreach (ElementId panelId in CurtainPanelsIdList)
                            {
                                Panel panel = null;
                                FamilyInstance doorwindows = null;
                                panel = linkDoc.GetElement(panelId) as Panel;
                                double panelArea = 0;

                                if (panel == null)
                                {
                                    doorwindows = linkDoc.GetElement(panelId) as FamilyInstance;
                                    if (doorwindows != null)
                                    {
                                        bool flag = false;
                                        BoundingBoxXYZ windowBoundingBox = doorwindows.get_BoundingBox(null);
                                        if (windowBoundingBox == null) continue;
                                        XYZ windowCenter = (windowBoundingBox.Max + windowBoundingBox.Min) / 2;
                                        Curve lineA = Line.CreateBound(windowCenter, windowCenter + (700 / 304.8) * doorwindows.FacingOrientation.Normalize()) as Curve;
                                        Curve lineB = Line.CreateBound(windowCenter, windowCenter + (700 / 304.8) * doorwindows.FacingOrientation.Normalize().Negate()) as Curve;
                                        Curve transformdlineA = lineA.CreateTransformed(transform) as Curve;
                                        Curve transformdlineB = lineB.CreateTransformed(transform) as Curve;

                                        SolidCurveIntersection intersectionB = spaceSolid.IntersectWithCurve(transformdlineB, intersectOptions);
                                        if (intersectionB.SegmentCount > 0)
                                        {
                                            foreach (Space secondSpace in spaceList)
                                            {
                                                if (secondSpace.Id == space.Id) continue;

                                                Solid secondSpaceSolid = null;

                                                GeometryElement secondGeomRoomElement = secondSpace.get_Geometry(new Options());
                                                foreach (GeometryObject secondGeomObj in secondGeomRoomElement)
                                                {
                                                    secondSpaceSolid = secondGeomObj as Solid;
                                                    if (secondSpaceSolid != null) break;
                                                }
                                                if (secondSpaceSolid != null)
                                                {
                                                    SolidCurveIntersection secondIntersection = secondSpaceSolid.IntersectWithCurve(transformdlineA, intersectOptions);
                                                    if (secondIntersection.SegmentCount > 0)
                                                    {
                                                        flag = true;
                                                    }
                                                }
                                            }
                                            if (flag == true) continue;

                                            double curtainWallPanelsHeight = 0;
                                            double curtainWallPanelsWidth = 0;

                                            if (doorwindows.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_HEIGHT) != null)
                                            {
                                                curtainWallPanelsHeight = doorwindows.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_HEIGHT).AsDouble();
                                            }
                                            if (doorwindows.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_WIDTH) != null)
                                            {
                                                curtainWallPanelsWidth = doorwindows.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_WIDTH).AsDouble();
                                            }

                                            panelArea += curtainWallPanelsHeight * curtainWallPanelsWidth;

                                            XYZ lineDirection = (lineA as Line).Direction;
                                            if (lineDirection.AngleTo(trueNorthBasisY) <= Math.PI / 8)
                                            {
                                                windowsAreaNorth += panelArea;
                                            }
                                            else if (lineDirection.AngleTo(trueNorthBasisY.Negate()) <= Math.PI / 8)
                                            {
                                                windowsAreaSouth += panelArea;
                                            }
                                            else if (lineDirection.AngleTo(trueNorthBasisX.Negate()) <= Math.PI / 8)
                                            {
                                                windowsAreaWest += panelArea;
                                            }
                                            else if (lineDirection.AngleTo(trueNorthBasisX) <= Math.PI / 8)
                                            {
                                                windowsAreaEast += panelArea;
                                            }

                                            else if (lineDirection.AngleTo((trueNorthBasisY + trueNorthBasisX.Negate()) / 2) < Math.PI / 8)
                                            {
                                                windowsAreaNorthwest += panelArea;
                                            }
                                            else if (lineDirection.AngleTo((trueNorthBasisY + trueNorthBasisX) / 2) < Math.PI / 8)
                                            {
                                                windowsAreaNortheast += panelArea;
                                            }

                                            else if (lineDirection.AngleTo((trueNorthBasisY.Negate() + trueNorthBasisX.Negate()) / 2) < Math.PI / 8)
                                            {
                                                windowsAreaSouthwest += panelArea;
                                            }
                                            else if (lineDirection.AngleTo((trueNorthBasisY.Negate() + trueNorthBasisX) / 2) < Math.PI / 8)
                                            {
                                                windowsAreaSoutheast += panelArea;
                                            }
                                        }
                                    }
                                }

                                if (panel != null)
                                {
                                    if (panel.Symbol.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_CONSTRUCTION_TYPE) != null)
                                    {
                                        if (panel.Symbol.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_CONSTRUCTION_TYPE).AsString() != "Остекление")
                                        {
                                            continue;
                                        }
                                    }
                                    bool flag = false;
                                    BoundingBoxXYZ panelBoundingBox = panel.get_BoundingBox(null);
                                    if (panelBoundingBox == null) continue;
                                    XYZ panelCenter = (panelBoundingBox.Max + panelBoundingBox.Min) / 2;
                                    Curve lineA = Line.CreateBound(panelCenter, panelCenter + (700 / 304.8) * panel.FacingOrientation.Normalize()) as Curve;
                                    Curve lineB = Line.CreateBound(panelCenter, panelCenter + (700 / 304.8) * panel.FacingOrientation.Normalize().Negate()) as Curve;
                                    Curve transformdlineA = lineA.CreateTransformed(transform) as Curve;
                                    Curve transformdlineB = lineB.CreateTransformed(transform) as Curve;
                                    
                                    SolidCurveIntersection intersectionB = spaceSolid.IntersectWithCurve(transformdlineB, intersectOptions);
                                    if (intersectionB.SegmentCount > 0)
                                    {
                                        foreach (Space secondSpace in spaceList)
                                        {
                                            if (secondSpace.Id == space.Id) continue;

                                            Solid secondSpaceSolid = null;

                                            GeometryElement secondGeomRoomElement = secondSpace.get_Geometry(new Options());
                                            foreach (GeometryObject secondGeomObj in secondGeomRoomElement)
                                            {
                                                secondSpaceSolid = secondGeomObj as Solid;
                                                if (secondSpaceSolid != null) break;
                                            }
                                            if (secondSpaceSolid != null)
                                            {
                                                SolidCurveIntersection secondIntersection = secondSpaceSolid.IntersectWithCurve(transformdlineA, intersectOptions);
                                                if (secondIntersection.SegmentCount > 0)
                                                {
                                                    flag = true;
                                                }
                                            }
                                        }
                                        if (flag == true) continue;

                                        XYZ lineDirection = (lineA as Line).Direction;
                                        panelArea = panel.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                                        if (lineDirection.AngleTo(trueNorthBasisY) <= Math.PI / 8)
                                        {
                                            windowsAreaNorth += panelArea;
                                        }
                                        else if (lineDirection.AngleTo(trueNorthBasisY.Negate()) <= Math.PI / 8)
                                        {
                                            windowsAreaSouth += panelArea;
                                        }
                                        else if (lineDirection.AngleTo(trueNorthBasisX.Negate()) <= Math.PI / 8)
                                        {
                                            windowsAreaWest += panelArea;
                                        }
                                        else if (lineDirection.AngleTo(trueNorthBasisX) <= Math.PI / 8)
                                        {
                                            windowsAreaEast += panelArea;
                                        }

                                        else if (lineDirection.AngleTo((trueNorthBasisY + trueNorthBasisX.Negate()) / 2) < Math.PI / 8)
                                        {
                                            windowsAreaNorthwest += panelArea;
                                        }
                                        else if (lineDirection.AngleTo((trueNorthBasisY + trueNorthBasisX) / 2) < Math.PI / 8)
                                        {
                                            windowsAreaNortheast += panelArea;
                                        }

                                        else if (lineDirection.AngleTo((trueNorthBasisY.Negate() + trueNorthBasisX.Negate()) / 2) < Math.PI / 8)
                                        {
                                            windowsAreaSouthwest += panelArea;
                                        }
                                        else if (lineDirection.AngleTo((trueNorthBasisY.Negate() + trueNorthBasisX) / 2) < Math.PI / 8)
                                        {
                                            windowsAreaSoutheast += panelArea;
                                        }
                                    }
                                }
                            }
                        }

                    }

                    space.get_Parameter(windowsAreaNorthGuid).Set(windowsAreaNorth);
                    space.get_Parameter(windowsAreaSouthGuid).Set(windowsAreaSouth);

                    space.get_Parameter(windowsAreaWestGuid).Set(windowsAreaWest);
                    space.get_Parameter(windowsAreaEastGuid).Set(windowsAreaEast);

                    space.get_Parameter(windowsAreaNorthwestGuid).Set(windowsAreaNorthwest);
                    space.get_Parameter(windowsAreaNortheastGuid).Set(windowsAreaNortheast);

                    space.get_Parameter(windowsAreaSouthwestGuid).Set(windowsAreaSouthwest);
                    space.get_Parameter(windowsAreaSoutheastGuid).Set(windowsAreaSoutheast);
                }
                t.Commit();
            }
            return Result.Succeeded;
        }
        private static List<Space> GetSpacesFromCurrentSelection(Document doc, Selection sel)
        {
            ICollection<ElementId> selectedIds = sel.GetElementIds();
            List<Space> tempSpacessList = new List<Space>();
            foreach (ElementId roomId in selectedIds)
            {
                if (doc.GetElement(roomId) is Space
                    && null != doc.GetElement(roomId).Category
                    && doc.GetElement(roomId).Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_MEPSpaces))
                {
                    tempSpacessList.Add(doc.GetElement(roomId) as Space);
                }
            }
            return tempSpacessList;
        }
    }
}
