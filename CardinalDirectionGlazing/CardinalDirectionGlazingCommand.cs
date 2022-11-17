using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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


            //Выбор связанного файла
            RevitLinkInstanceSelectionFilter selFilterRevitLinkInstance = new RevitLinkInstanceSelectionFilter();
            Reference selRevitLinkInstance = null;
            try
            {
                selRevitLinkInstance = sel.PickObject(ObjectType.Element, selFilterRevitLinkInstance, "Выберите связанный файл!");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }

            IEnumerable<RevitLinkInstance> revitLinkInstance = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Where(li => li.Id == selRevitLinkInstance.ElementId)
                .Cast<RevitLinkInstance>();
            if (revitLinkInstance.Count() == 0)
            {
                TaskDialog.Show("Ravit", "Связанный файл не найден!");
                return Result.Cancelled;
            }
            linkDoc = revitLinkInstance.First().GetLinkDocument();
            Transform transform = revitLinkInstance.First().GetTotalTransform();

            List<Space> spaceList = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .WhereElementIsNotElementType()
                .Cast<Space>()
                .Where(s => s.Area > 0)
                .ToList();


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Север помнит");
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
                        List<FamilyInstance> windowsOnSpaceLevelList = new FilteredElementCollector(linkDoc)
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

                        foreach (FamilyInstance window in windowsOnSpaceLevelList)
                        {
                            BoundingBoxXYZ windowBoundingBox = window.get_BoundingBox(null);
                            if (windowBoundingBox == null) continue;
                            XYZ windowCenter = (windowBoundingBox.Max + windowBoundingBox.Min) / 2;
                            Curve lineA = Line.CreateBound(windowCenter, windowCenter + (700 / 304.8) * window.FacingOrientation.Normalize()) as Curve;
                            Curve lineB = Line.CreateBound(windowCenter, windowCenter + (700 / 304.8) * window.FacingOrientation.Normalize().Negate()) as Curve;
                            Curve transformdlineB = lineB.CreateTransformed(transform) as Curve;
                            SolidCurveIntersection intersection = spaceSolid.IntersectWithCurve(transformdlineB, intersectOptions);
                            if (intersection.SegmentCount > 0)
                            {
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
                                if(lineDirection.AngleTo(XYZ.BasisY) <= Math.PI / 8)
                                {
                                    windowsAreaNorth += windowArea;
                                }
                                else if (lineDirection.AngleTo(XYZ.BasisY.Negate()) <= Math.PI / 8)
                                {
                                    windowsAreaSouth += windowArea;
                                }
                                else if (lineDirection.AngleTo(XYZ.BasisX.Negate()) <= Math.PI / 8)
                                {
                                    windowsAreaWest += windowArea;
                                }
                                else if (lineDirection.AngleTo(XYZ.BasisX) <= Math.PI / 8)
                                {
                                    windowsAreaEast += windowArea;
                                }

                                else if (lineDirection.AngleTo((XYZ.BasisY + XYZ.BasisX.Negate())/2) < Math.PI / 8)
                                {
                                    windowsAreaNorthwest += windowArea;
                                }
                                else if (lineDirection.AngleTo((XYZ.BasisY + XYZ.BasisX) / 2) < Math.PI / 8)
                                {
                                    windowsAreaNortheast += windowArea;
                                }

                                else if (lineDirection.AngleTo((XYZ.BasisY.Negate() + XYZ.BasisX.Negate()) / 2) < Math.PI / 8)
                                {
                                    windowsAreaSouthwest += windowArea;
                                }
                                else if (lineDirection.AngleTo((XYZ.BasisY.Negate() + XYZ.BasisX) / 2) < Math.PI / 8)
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
                                        BoundingBoxXYZ windowBoundingBox = doorwindows.get_BoundingBox(null);
                                        if (windowBoundingBox == null) continue;
                                        XYZ windowCenter = (windowBoundingBox.Max + windowBoundingBox.Min) / 2;
                                        Curve lineA = Line.CreateBound(windowCenter, windowCenter + (700 / 304.8) * doorwindows.FacingOrientation.Normalize()) as Curve;
                                        Curve lineB = Line.CreateBound(windowCenter, windowCenter + (700 / 304.8) * doorwindows.FacingOrientation.Normalize().Negate()) as Curve;
                                        Curve transformdlineB = lineB.CreateTransformed(transform) as Curve;

                                        SolidCurveIntersection intersectionB = spaceSolid.IntersectWithCurve(transformdlineB, intersectOptions);
                                        if (intersectionB.SegmentCount > 0)
                                        {
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
                                            if (lineDirection.AngleTo(XYZ.BasisY) <= Math.PI / 8)
                                            {
                                                windowsAreaNorth += panelArea;
                                            }
                                            else if (lineDirection.AngleTo(XYZ.BasisY.Negate()) <= Math.PI / 8)
                                            {
                                                windowsAreaSouth += panelArea;
                                            }
                                            else if (lineDirection.AngleTo(XYZ.BasisX.Negate()) <= Math.PI / 8)
                                            {
                                                windowsAreaWest += panelArea;
                                            }
                                            else if (lineDirection.AngleTo(XYZ.BasisX) <= Math.PI / 8)
                                            {
                                                windowsAreaEast += panelArea;
                                            }

                                            else if (lineDirection.AngleTo((XYZ.BasisY + XYZ.BasisX.Negate()) / 2) < Math.PI / 8)
                                            {
                                                windowsAreaNorthwest += panelArea;
                                            }
                                            else if (lineDirection.AngleTo((XYZ.BasisY + XYZ.BasisX) / 2) < Math.PI / 8)
                                            {
                                                windowsAreaNortheast += panelArea;
                                            }

                                            else if (lineDirection.AngleTo((XYZ.BasisY.Negate() + XYZ.BasisX.Negate()) / 2) < Math.PI / 8)
                                            {
                                                windowsAreaSouthwest += panelArea;
                                            }
                                            else if (lineDirection.AngleTo((XYZ.BasisY.Negate() + XYZ.BasisX) / 2) < Math.PI / 8)
                                            {
                                                windowsAreaSoutheast += panelArea;
                                            }
                                        }
                                    }
                                }

                                if (panel != null)
                                {
                                    BoundingBoxXYZ panelBoundingBox = panel.get_BoundingBox(null);
                                    if (panelBoundingBox == null) continue;
                                    XYZ panelCenter = (panelBoundingBox.Max + panelBoundingBox.Min) / 2;
                                    Curve lineA = Line.CreateBound(panelCenter, panelCenter + (700 / 304.8) * panel.FacingOrientation.Normalize()) as Curve;
                                    Curve lineB = Line.CreateBound(panelCenter, panelCenter + (700 / 304.8) * panel.FacingOrientation.Normalize().Negate()) as Curve;
                                    Curve transformdlineB = lineB.CreateTransformed(transform) as Curve;
                                    
                                    SolidCurveIntersection intersectionB = spaceSolid.IntersectWithCurve(transformdlineB, intersectOptions);
                                    if (intersectionB.SegmentCount > 0)
                                    {
                                        XYZ lineDirection = (lineA as Line).Direction;
                                        panelArea = panel.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                                        if (lineDirection.AngleTo(XYZ.BasisY) <= Math.PI / 8)
                                        {
                                            windowsAreaNorth += panelArea;
                                        }
                                        else if (lineDirection.AngleTo(XYZ.BasisY.Negate()) <= Math.PI / 8)
                                        {
                                            windowsAreaSouth += panelArea;
                                        }
                                        else if (lineDirection.AngleTo(XYZ.BasisX.Negate()) <= Math.PI / 8)
                                        {
                                            windowsAreaWest += panelArea;
                                        }
                                        else if (lineDirection.AngleTo(XYZ.BasisX) <= Math.PI / 8)
                                        {
                                            windowsAreaEast += panelArea;
                                        }

                                        else if (lineDirection.AngleTo((XYZ.BasisY + XYZ.BasisX.Negate()) / 2) < Math.PI / 8)
                                        {
                                            windowsAreaNorthwest += panelArea;
                                        }
                                        else if (lineDirection.AngleTo((XYZ.BasisY + XYZ.BasisX) / 2) < Math.PI / 8)
                                        {
                                            windowsAreaNortheast += panelArea;
                                        }

                                        else if (lineDirection.AngleTo((XYZ.BasisY.Negate() + XYZ.BasisX.Negate()) / 2) < Math.PI / 8)
                                        {
                                            windowsAreaSouthwest += panelArea;
                                        }
                                        else if (lineDirection.AngleTo((XYZ.BasisY.Negate() + XYZ.BasisX) / 2) < Math.PI / 8)
                                        {
                                            windowsAreaSoutheast += panelArea;
                                        }
                                    }
                                }
                            }
                        }

                    }

                    space.LookupParameter("ПлощадьОкон_С").Set(windowsAreaNorth);
                    space.LookupParameter("ПлощадьОкон_Ю").Set(windowsAreaSouth);

                    space.LookupParameter("ПлощадьОкон_З").Set(windowsAreaWest);
                    space.LookupParameter("ПлощадьОкон_В").Set(windowsAreaEast);

                    space.LookupParameter("ПлощадьОкон_СЗ").Set(windowsAreaNorthwest);
                    space.LookupParameter("ПлощадьОкон_СВ").Set(windowsAreaNortheast);

                    space.LookupParameter("ПлощадьОкон_ЮЗ").Set(windowsAreaSouthwest);
                    space.LookupParameter("ПлощадьОкон_ЮВ").Set(windowsAreaSoutheast);
                }
                t.Commit();
            }
            return Result.Succeeded;
        }
    }
}
