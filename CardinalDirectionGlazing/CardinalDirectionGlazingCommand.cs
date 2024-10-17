using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace CardinalDirectionGlazing
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CardinalDirectionGlazingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                _ = GetPluginStartInfo();
            }
            catch { }

            Document doc = commandData.Application.ActiveUIDocument.Document;
            Selection sel = commandData.Application.ActiveUIDocument.Selection;

            // GUIDы параметров окон по сторонам света
            Guid windowsAreaNorthGuid = new Guid("820af414-f6ec-472d-887c-a2046a0c5988");
            Guid windowsAreaSouthGuid = new Guid("81ab8e02-45c6-4d26-b0e5-a6736b0c352d");
            Guid windowsAreaWestGuid = new Guid("65fe3416-f836-48ff-bed9-3fdf2126a1f9");
            Guid windowsAreaEastGuid = new Guid("fc33c487-9bbb-43d6-a7f4-aba5f9638fe3");
            Guid windowsAreaNorthwestGuid = new Guid("f78f8a53-cea7-4e00-955c-3748aa7a37c7");
            Guid windowsAreaNortheastGuid = new Guid("b8120c53-0793-4932-bc71-845302914573");
            Guid windowsAreaSouthwestGuid = new Guid("3ff1f178-2cff-4b54-a0d3-eee58fa1622c");
            Guid windowsAreaSoutheastGuid = new Guid("c5e261ae-68f5-4a91-a55f-8686d278f5ab");

            // Получаем список связанных файлов
            List<RevitLinkInstance> revitLinkInstanceList = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            // Открываем окно WPF для выбора опций
            CardinalDirectionGlazingWPF cardinalDirectionGlazingWPF = new CardinalDirectionGlazingWPF(revitLinkInstanceList);
            cardinalDirectionGlazingWPF.ShowDialog();
            if (cardinalDirectionGlazingWPF.DialogResult != true)
            {
                return Result.Cancelled;
            }

            // Проверка выбранного связанного файла
            if (cardinalDirectionGlazingWPF.SpacesOrRoomsForProcessingButtonName == "radioButton_Spaces")
            {
                // Если выбрана обработка пространств, связанный файл обязателен
                if (cardinalDirectionGlazingWPF.SelectedRevitLinkInstance == null)
                {
                    TaskDialog.Show("Revit", "Связанный файл не выбран! Для обработки пространств необходим связанный файл.");
                    return Result.Cancelled;
                }
            }


            Document linkDoc = null;
            Transform transform = null;

            // Всегда берем истинный север из текущего документа
            ProjectPosition position = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);
            Transform trueNorthTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, -position.Angle, XYZ.Zero);
            XYZ trueNorthBasisY = trueNorthTransform.OfVector(XYZ.BasisY);
            XYZ trueNorthBasisX = trueNorthTransform.OfVector(XYZ.BasisX);

            // Если выбран связанный файл, используем его для обработки, но истинный север берем из текущего документа
            if (cardinalDirectionGlazingWPF.SelectedRevitLinkInstance != null)
            {
                linkDoc = cardinalDirectionGlazingWPF.SelectedRevitLinkInstance.GetLinkDocument();
                transform = cardinalDirectionGlazingWPF.SelectedRevitLinkInstance.GetTotalTransform();
            }

            // Выбираем обработку: пространства или помещения
            List<Element> elementsList = new List<Element>();
            // Проверка, обрабатываются ли пространства или помещения
            if (cardinalDirectionGlazingWPF.SpacesOrRoomsForProcessingButtonName == "radioButton_Spaces")
            {
                if (cardinalDirectionGlazingWPF.SpacesForProcessingButtonName == "radioButton_Selected")
                {
                    // Получаем выбранные пространства
                    List<Space> selectedSpaces = GetSpacesFromCurrentSelection(doc, sel);

                    // Если ничего не выбрано, даем пользователю возможность выбрать пространства
                    if (selectedSpaces.Count == 0)
                    {
                        try
                        {
                            IList<Reference> selectedReferences = sel.PickObjects(ObjectType.Element, new SpaceSelectionFilter(), "Выберите пространства");
                            selectedSpaces = selectedReferences.Select(r => doc.GetElement(r.ElementId) as Space).ToList();
                        }
                        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                        {
                            return Result.Cancelled;
                        }
                    }

                    // Преобразуем пространства в список элементов для дальнейшей обработки
                    elementsList = selectedSpaces.Cast<Element>().ToList();
                }
                else
                {
                    // Обрабатываем все пространства
                    elementsList = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_MEPSpaces)
                        .WhereElementIsNotElementType()
                        .ToList();
                }
            }
            else if (cardinalDirectionGlazingWPF.SpacesOrRoomsForProcessingButtonName == "radioButton_Rooms")
            {
                if (cardinalDirectionGlazingWPF.SpacesForProcessingButtonName == "radioButton_Selected")
                {
                    // Получаем выбранные помещения
                    List<Room> selectedRooms = GetRoomsFromCurrentSelection(doc, sel);

                    // Если ничего не выбрано, даем пользователю возможность выбрать помещения
                    if (selectedRooms.Count == 0)
                    {
                        try
                        {
                            IList<Reference> selectedReferences = sel.PickObjects(ObjectType.Element, new RoomSelectionFilter(), "Выберите помещения");
                            selectedRooms = selectedReferences.Select(r => doc.GetElement(r.ElementId) as Room).ToList();
                        }
                        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                        {
                            return Result.Cancelled;
                        }
                    }

                    // Преобразуем помещения в список элементов для дальнейшей обработки
                    elementsList = selectedRooms.Cast<Element>().ToList();
                }
                else
                {
                    // Обрабатываем все помещения
                    elementsList = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Rooms)
                        .WhereElementIsNotElementType()
                        .ToList();
                }
            }

            if (elementsList.Count == 0)
            {
                TaskDialog.Show("Revit", "Не найдены пространства или помещения для обработки.");
                return Result.Cancelled;
            }

            // Проверка наличия параметров в первом элементе
            if (elementsList.Count != 0)
            {
                Element firstElement = elementsList.First();

                // Проверка каждого параметра по отдельности и вывод сообщения, если параметр отсутствует
                if (firstElement.get_Parameter(windowsAreaNorthGuid) == null)
                {
                    TaskDialog.Show("Revit", "Параметр \"ПлощадьОкон_С\" не найден! Добавьте параметр.");
                    return Result.Cancelled;
                }

                if (firstElement.get_Parameter(windowsAreaSouthGuid) == null)
                {
                    TaskDialog.Show("Revit", "Параметр \"ПлощадьОкон_Ю\" не найден! Добавьте параметр.");
                    return Result.Cancelled;
                }

                if (firstElement.get_Parameter(windowsAreaWestGuid) == null)
                {
                    TaskDialog.Show("Revit", "Параметр \"ПлощадьОкон_З\" не найден! Добавьте параметр.");
                    return Result.Cancelled;
                }

                if (firstElement.get_Parameter(windowsAreaEastGuid) == null)
                {
                    TaskDialog.Show("Revit", "Параметр \"ПлощадьОкон_В\" не найден! Добавьте параметр.");
                    return Result.Cancelled;
                }

                if (firstElement.get_Parameter(windowsAreaNorthwestGuid) == null)
                {
                    TaskDialog.Show("Revit", "Параметр \"ПлощадьОкон_СЗ\" не найден! Добавьте параметр.");
                    return Result.Cancelled;
                }

                if (firstElement.get_Parameter(windowsAreaNortheastGuid) == null)
                {
                    TaskDialog.Show("Revit", "Параметр \"ПлощадьОкон_СВ\" не найден! Добавьте параметр.");
                    return Result.Cancelled;
                }

                if (firstElement.get_Parameter(windowsAreaSouthwestGuid) == null)
                {
                    TaskDialog.Show("Revit", "Параметр \"ПлощадьОкон_ЮЗ\" не найден! Добавьте параметр.");
                    return Result.Cancelled;
                }

                if (firstElement.get_Parameter(windowsAreaSoutheastGuid) == null)
                {
                    TaskDialog.Show("Revit", "Параметр \"ПлощадьОкон_ЮВ\" не найден! Добавьте параметр.");
                    return Result.Cancelled;
                }
            }

            // Инициализируем списки для окон и стен с остеклением
            List<FamilyInstance> windowsList = new List<FamilyInstance>();
            List<Wall> curtainWallsList = new List<Wall>();

            // Инициализируем списки для окон и стен из связанного документа
            List<FamilyInstance> linkedWindowsList = new List<FamilyInstance>();
            List<Wall> linkedCurtainWallsList = new List<Wall>();

            if (cardinalDirectionGlazingWPF.SpacesOrRoomsForProcessingButtonName == "radioButton_Spaces")
            {
                // Если обрабатываем пространства, берем окна и стены только из связанного файла
                if (linkDoc != null)
                {
                    linkedWindowsList = new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Windows)
                        .OfClass(typeof(FamilyInstance))
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(w => w.SuperComponent == null)
                        .ToList();

                    linkedCurtainWallsList = new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Walls)
                        .OfClass(typeof(Wall))
                        .WhereElementIsNotElementType()
                        .Cast<Wall>()
                        .Where(w => w.CurtainGrid != null)
                        .ToList();
                }
                else
                {
                    TaskDialog.Show("Revit", "Связанный файл для обработки пространств не найден.");
                    return Result.Cancelled;
                }
            }
            else if (cardinalDirectionGlazingWPF.SpacesOrRoomsForProcessingButtonName == "radioButton_Rooms")
            {
                // Если обрабатываем помещения, берем окна и стены как из связанного файла, так и из текущего документа

                // Окна и стены из текущего документа
                windowsList = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Windows)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(w => w.SuperComponent == null)
                    .ToList();

                curtainWallsList = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .Where(w => w.CurtainGrid != null)
                    .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Наружный витраж")
                    .ToList();

                // Если выбран связанный файл, добавляем окна и стены из связанного файла
                if (linkDoc != null)
                {
                    linkedWindowsList = new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Windows)
                        .OfClass(typeof(FamilyInstance))
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(w => w.SuperComponent == null)
                        .ToList();

                    linkedCurtainWallsList = new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Walls)
                        .OfClass(typeof(Wall))
                        .WhereElementIsNotElementType()
                        .Cast<Wall>()
                        .Where(w => w.CurtainGrid != null)
                        .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Наружный витраж")
                        .ToList();
                }
            }

            // Начинаем транзакцию для обновления данных в Revit
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Остекление по сторонам света");

                foreach (Element element in elementsList)
                {
                    double windowsAreaNorth = 0;
                    double windowsAreaSouth = 0;
                    double windowsAreaWest = 0;
                    double windowsAreaEast = 0;
                    double windowsAreaNorthwest = 0;
                    double windowsAreaNortheast = 0;
                    double windowsAreaSouthwest = 0;
                    double windowsAreaSoutheast = 0;

                    Solid elementSolid = GetSolidFromElement(element);
                    if (elementSolid == null) continue;

                    // Обработка окон из текущего документа (не используем трансформацию)
                    foreach (FamilyInstance window in windowsList)
                    {
                        double roughHeight = 0;
                        double roughWidth = 0;
                        double caseworkHeight = 0;
                        double caseworkWidth = 0;
                        double maxHeight = 0;
                        double maxWidth = 0;

                        // Получаем параметры ROUGH HEIGHT и ROUGH WIDTH
                        if (window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM) != null)
                        {
                            roughHeight = window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM).AsDouble();
                        }
                        if (window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM) != null)
                        {
                            roughWidth = window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM).AsDouble();
                        }

                        // Получаем параметры CASEWORK для высоты и ширины
                        if (window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_HEIGHT) != null)
                        {
                            caseworkHeight = window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_HEIGHT).AsDouble();
                        }
                        if (window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_WIDTH) != null)
                        {
                            caseworkWidth = window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_WIDTH).AsDouble();
                        }

                        // Выбираем максимальные значения высоты и ширины
                        maxHeight = Math.Max(roughHeight, caseworkHeight);
                        maxWidth = Math.Max(roughWidth, caseworkWidth);

                        double windowArea = maxHeight * maxWidth;

                        BoundingBoxXYZ windowBoundingBox = window.get_BoundingBox(null);
                        if (windowBoundingBox == null) continue;

                        XYZ windowCenter = (windowBoundingBox.Max + windowBoundingBox.Min) / 2;
                        Curve windowCurve = Line.CreateBound(windowCenter, windowCenter + window.FacingOrientation.Negate() * 700 / 304.8);

                        SolidCurveIntersection intersection = elementSolid.IntersectWithCurve(windowCurve, new SolidCurveIntersectionOptions());
                        if (intersection.SegmentCount > 0)
                        {
                            UpdateWindowAreas(ref windowsAreaNorth, ref windowsAreaSouth, ref windowsAreaWest, ref windowsAreaEast, ref windowsAreaNorthwest, ref windowsAreaNortheast, ref windowsAreaSouthwest, ref windowsAreaSoutheast, windowArea, window.FacingOrientation, trueNorthBasisX, trueNorthBasisY);
                        }
                    }

                    // Обработка панелей витражей из текущего документа (не используем трансформацию)
                    foreach (Wall wall in curtainWallsList)
                    {
                        foreach (ElementId panelId in wall.CurtainGrid.GetPanelIds())
                        {
                            Panel panel = doc.GetElement(panelId) as Panel;
                            if (panel == null || !IsPanelGlazing(panel)) continue;

                            BoundingBoxXYZ panelBoundingBox = panel.get_BoundingBox(null);
                            if (panelBoundingBox == null) continue;

                            XYZ panelCenter = (panelBoundingBox.Max + panelBoundingBox.Min) / 2;

                            // Создаем две линии: одну в направлении FacingOrientation, другую в противоположном направлении
                            Curve panelCurveForward = Line.CreateBound(panelCenter, panelCenter + panel.FacingOrientation * 700 / 304.8);
                            Curve panelCurveBackward = Line.CreateBound(panelCenter, panelCenter + panel.FacingOrientation.Negate() * 700 / 304.8);

                            SolidCurveIntersectionOptions intersectionOptions = new SolidCurveIntersectionOptions();

                            // Проверяем пересечение в прямом направлении
                            SolidCurveIntersection intersectionForward = elementSolid.IntersectWithCurve(panelCurveForward, intersectionOptions);
                            bool intersected = false;

                            if (intersectionForward.SegmentCount > 0)
                            {
                                double panelArea = panel.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                                UpdateWindowAreas(ref windowsAreaNorth, ref windowsAreaSouth, ref windowsAreaWest, ref windowsAreaEast, ref windowsAreaNorthwest, ref windowsAreaNortheast, ref windowsAreaSouthwest, ref windowsAreaSoutheast, panelArea, panel.FacingOrientation.Negate(), trueNorthBasisX, trueNorthBasisY);
                                intersected = true;
                            }

                            // Проверяем пересечение в противоположном направлении, если не было пересечения в прямом
                            if (!intersected)
                            {
                                SolidCurveIntersection intersectionBackward = elementSolid.IntersectWithCurve(panelCurveBackward, intersectionOptions);
                                if (intersectionBackward.SegmentCount > 0)
                                {
                                    double panelArea = panel.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                                    UpdateWindowAreas(ref windowsAreaNorth, ref windowsAreaSouth, ref windowsAreaWest, ref windowsAreaEast, ref windowsAreaNorthwest, ref windowsAreaNortheast, ref windowsAreaSouthwest, ref windowsAreaSoutheast, panelArea, panel.FacingOrientation, trueNorthBasisX, trueNorthBasisY);
                                }
                            }
                        }
                    }

                    // Обработка окон из связанного документа (используем transform)
                    foreach (FamilyInstance window in linkedWindowsList)
                    {
                        double roughHeight = 0;
                        double roughWidth = 0;
                        double caseworkHeight = 0;
                        double caseworkWidth = 0;
                        double maxHeight = 0;
                        double maxWidth = 0;

                        // Получаем параметры ROUGH HEIGHT и ROUGH WIDTH
                        if (window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM) != null)
                        {
                            roughHeight = window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM).AsDouble();
                        }
                        if (window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM) != null)
                        {
                            roughWidth = window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM).AsDouble();
                        }

                        // Получаем параметры CASEWORK для высоты и ширины
                        if (window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_HEIGHT) != null)
                        {
                            caseworkHeight = window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_HEIGHT).AsDouble();
                        }
                        if (window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_WIDTH) != null)
                        {
                            caseworkWidth = window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_WIDTH).AsDouble();
                        }

                        // Выбираем максимальные значения высоты и ширины
                        maxHeight = Math.Max(roughHeight, caseworkHeight);
                        maxWidth = Math.Max(roughWidth, caseworkWidth);

                        double windowArea = maxHeight * maxWidth;

                        BoundingBoxXYZ windowBoundingBox = window.get_BoundingBox(null);
                        if (windowBoundingBox == null) continue;

                        XYZ windowCenter = (windowBoundingBox.Max + windowBoundingBox.Min) / 2;
                        windowCenter = transform.OfPoint(windowCenter); // Применяем трансформацию для окна из связанного файла
                        Curve windowCurve = Line.CreateBound(windowCenter, windowCenter + transform.OfVector(window.FacingOrientation.Negate()) * 700 / 304.8);

                        SolidCurveIntersection intersection = elementSolid.IntersectWithCurve(windowCurve, new SolidCurveIntersectionOptions());
                        if (intersection.SegmentCount > 0)
                        {
                            UpdateWindowAreas(ref windowsAreaNorth, ref windowsAreaSouth, ref windowsAreaWest, ref windowsAreaEast, ref windowsAreaNorthwest, ref windowsAreaNortheast, ref windowsAreaSouthwest, ref windowsAreaSoutheast, windowArea, transform.OfVector(window.FacingOrientation), trueNorthBasisX, trueNorthBasisY);
                        }
                    }

                    // Обработка панелей витражей из связанного документа (используем transform)
                    foreach (Wall wall in linkedCurtainWallsList)
                    {
                        foreach (ElementId panelId in wall.CurtainGrid.GetPanelIds())
                        {
                            Panel panel = linkDoc.GetElement(panelId) as Panel;
                            if (panel == null || !IsPanelGlazing(panel)) continue;

                            BoundingBoxXYZ panelBoundingBox = panel.get_BoundingBox(null);
                            if (panelBoundingBox == null) continue;

                            XYZ panelCenter = (panelBoundingBox.Max + panelBoundingBox.Min) / 2;
                            panelCenter = transform.OfPoint(panelCenter); // Применяем трансформацию для панели из связанного файла

                            // Создаем две линии: одну в направлении FacingOrientation, другую в противоположном направлении
                            Curve panelCurveForward = Line.CreateBound(panelCenter, panelCenter + transform.OfVector(panel.FacingOrientation) * 700 / 304.8);
                            Curve panelCurveBackward = Line.CreateBound(panelCenter, panelCenter + transform.OfVector(panel.FacingOrientation.Negate()) * 700 / 304.8);

                            SolidCurveIntersectionOptions intersectionOptions = new SolidCurveIntersectionOptions();

                            // Проверяем пересечение в прямом направлении
                            SolidCurveIntersection intersectionForward = elementSolid.IntersectWithCurve(panelCurveForward, intersectionOptions);
                            bool intersected = false;

                            if (intersectionForward.SegmentCount > 0)
                            {
                                double panelArea = panel.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                                UpdateWindowAreas(ref windowsAreaNorth, ref windowsAreaSouth, ref windowsAreaWest, ref windowsAreaEast, ref windowsAreaNorthwest, ref windowsAreaNortheast, ref windowsAreaSouthwest, ref windowsAreaSoutheast, panelArea, transform.OfVector(panel.FacingOrientation.Negate()), trueNorthBasisX, trueNorthBasisY);
                                intersected = true;
                            }

                            // Проверяем пересечение в противоположном направлении, если не было пересечения в прямом
                            if (!intersected)
                            {
                                SolidCurveIntersection intersectionBackward = elementSolid.IntersectWithCurve(panelCurveBackward, intersectionOptions);
                                if (intersectionBackward.SegmentCount > 0)
                                {
                                    double panelArea = panel.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                                    UpdateWindowAreas(ref windowsAreaNorth, ref windowsAreaSouth, ref windowsAreaWest, ref windowsAreaEast, ref windowsAreaNorthwest, ref windowsAreaNortheast, ref windowsAreaSouthwest, ref windowsAreaSoutheast, panelArea, transform.OfVector(panel.FacingOrientation), trueNorthBasisX, trueNorthBasisY);
                                }
                            }
                        }
                    }

                    // Установка значений параметров для каждого элемента
                    element.get_Parameter(windowsAreaNorthGuid)?.Set(windowsAreaNorth);
                    element.get_Parameter(windowsAreaSouthGuid)?.Set(windowsAreaSouth);
                    element.get_Parameter(windowsAreaWestGuid)?.Set(windowsAreaWest);
                    element.get_Parameter(windowsAreaEastGuid)?.Set(windowsAreaEast);
                    element.get_Parameter(windowsAreaNorthwestGuid)?.Set(windowsAreaNorthwest);
                    element.get_Parameter(windowsAreaNortheastGuid)?.Set(windowsAreaNortheast);
                    element.get_Parameter(windowsAreaSouthwestGuid)?.Set(windowsAreaSouthwest);
                    element.get_Parameter(windowsAreaSoutheastGuid)?.Set(windowsAreaSoutheast);
                }

                t.Commit();
            }

            return Result.Succeeded;
        }

        // Дополнительные методы для получения Solid, вычисления площадей и проверки панели
        private Solid GetSolidFromElement(Element element)
        {
            GeometryElement geomElement = element.get_Geometry(new Options());
            foreach (GeometryObject geomObj in geomElement)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                    return solid;
            }
            return null;
        }

        private void UpdateWindowAreas(
            ref double north, ref double south, ref double west, ref double east,
            ref double northwest, ref double northeast, ref double southwest, ref double southeast,
            double area, XYZ orientation, XYZ trueNorthBasisX, XYZ trueNorthBasisY)
        {
            // Рассчитываем углы до каждой из сторон света
            double angleToNorth = orientation.AngleTo(trueNorthBasisY);
            double angleToSouth = orientation.AngleTo(trueNorthBasisY.Negate());
            double angleToWest = orientation.AngleTo(trueNorthBasisX.Negate());
            double angleToEast = orientation.AngleTo(trueNorthBasisX);

            // Рассчитываем углы до промежуточных направлений
            double angleToNorthwest = orientation.AngleTo(trueNorthBasisY + trueNorthBasisX.Negate());
            double angleToNortheast = orientation.AngleTo(trueNorthBasisY + trueNorthBasisX);
            double angleToSouthwest = orientation.AngleTo(trueNorthBasisY.Negate() + trueNorthBasisX.Negate());
            double angleToSoutheast = orientation.AngleTo(trueNorthBasisY.Negate() + trueNorthBasisX);

            // Условие для основного направления Север
            if (angleToNorth <= Math.PI / 8)
            {
                north += area;
            }
            // Условие для основного направления Юг
            else if (angleToSouth <= Math.PI / 8)
            {
                south += area;
            }
            // Условие для основного направления Запад
            else if (angleToWest <= Math.PI / 8)
            {
                west += area;
            }
            // Условие для основного направления Восток
            else if (angleToEast <= Math.PI / 8)
            {
                east += area;
            }
            // Условие для промежуточного направления Северо-Запад
            else if (angleToNorthwest <= Math.PI / 8)
            {
                northwest += area;
            }
            // Условие для промежуточного направления Северо-Восток
            else if (angleToNortheast <= Math.PI / 8)
            {
                northeast += area;
            }
            // Условие для промежуточного направления Юго-Запад
            else if (angleToSouthwest <= Math.PI / 8)
            {
                southwest += area;
            }
            // Условие для промежуточного направления Юго-Восток
            else if (angleToSoutheast <= Math.PI / 8)
            {
                southeast += area;
            }
        }

        private bool IsPanelGlazing(Panel panel)
        {
            string constructionType = panel.Symbol.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_CONSTRUCTION_TYPE)?.AsString();
            return constructionType == "Остекление";
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

        private static List<Room> GetRoomsFromCurrentSelection(Document doc, Selection sel)
        {
            ICollection<ElementId> selectedIds = sel.GetElementIds();
            List<Room> tempRoomsList = new List<Room>();
            foreach (ElementId roomId in selectedIds)
            {
                if (doc.GetElement(roomId) is Room
                    && null != doc.GetElement(roomId).Category
                    && doc.GetElement(roomId).Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Rooms))
                {
                    tempRoomsList.Add(doc.GetElement(roomId) as Room);
                }
            }
            return tempRoomsList;
        }

        private static async Task GetPluginStartInfo()
        {
            // Получаем сборку, в которой выполняется текущий код
            Assembly thisAssembly = Assembly.GetExecutingAssembly();
            string assemblyName = "CardinalDirectionGlazing";
            string assemblyNameRus = "Остекление по сторонам";
            string assemblyFolderPath = Path.GetDirectoryName(thisAssembly.Location);

            int lastBackslashIndex = assemblyFolderPath.LastIndexOf("\\");
            string dllPath = assemblyFolderPath.Substring(0, lastBackslashIndex + 1) + "PluginInfoCollector\\PluginInfoCollector.dll";

            Assembly assembly = Assembly.LoadFrom(dllPath);
            Type type = assembly.GetType("PluginInfoCollector.InfoCollector");

            if (type != null)
            {
                // Создание экземпляра класса
                object instance = Activator.CreateInstance(type);

                // Получение метода CollectPluginUsageAsync
                var method = type.GetMethod("CollectPluginUsageAsync");

                if (method != null)
                {
                    // Вызов асинхронного метода через reflection
                    Task task = (Task)method.Invoke(instance, new object[] { assemblyName, assemblyNameRus });
                    await task;  // Ожидание завершения асинхронного метода
                }
            }
        }
    }
}
