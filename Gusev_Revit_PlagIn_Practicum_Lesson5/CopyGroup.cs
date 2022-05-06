using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Gusev_Revit_PlagIn_Practicum_Lesson5
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            List<String> levelNames = new List<String>();
            levelNames.Add("Уровень 1");
            levelNames.Add("Уровень 2");

            List<Level> levels = GetLevels(doc, levelNames);

            List<Wall> walls = CreateWalls(doc, 10000, 5000, 0, 0, levels.ElementAt(0), levels.ElementAt(1));


            Transaction transaction = new Transaction(doc, "Установка двери");
            transaction.Start();
            AddDoor(doc, levels.ElementAt(0), walls[0]);
            transaction.Commit();

            Transaction transaction1 = new Transaction(doc, "Установка окон");
            transaction.Start();
            AddWindow(doc, levels.ElementAt(0), walls[1], 1000);
            AddWindow(doc, levels.ElementAt(0), walls[2], 1000);
            AddWindow(doc, levels.ElementAt(0), walls[3], 1000);
            transaction.Commit();



            return Result.Succeeded;
        }

        private void AddWindow(Document doc, Level level, Wall wall, double _height)
        {
            double height = UnitUtils.ConvertToInternalUnits(_height, UnitTypeId.Millimeters);

            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 1220 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();
            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ pointHeight = new XYZ(0, 0, height);
            XYZ point = (point1 + point2) / 2;
            point = point + pointHeight;

            if (!windowType.IsActive)
                windowType.Activate();

            doc.Create.NewFamilyInstance(point, windowType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

        }

        private void AddDoor(Document doc, Level level, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();
            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2)/2;

            if (!doorType.IsActive)
                doorType.Activate();

            doc.Create.NewFamilyInstance(point, doorType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural) ;

        }

        // Метод получающий существующие уровни из документа, имена которых перечислены в входном списке

        public List<Level> GetLevels(Document doc, List<String> levelNames)
        {
            List<Level> listNamedlevel = new List<Level>();
            List<Level> listlevel = new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .OfType<Level>()
                            .ToList();
            foreach (String levelName in levelNames)
            {
                try
                {
                    listNamedlevel.Add(listlevel.FirstOrDefault(x => x.Name.Equals(levelName)));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return listNamedlevel;
        }

        public List<Wall> CreateWalls(Document doc, double _width, double _depth, double x, double y,
                                    Level baseLevel, Level upperLevel)
        {
            double width = UnitUtils.ConvertToInternalUnits(_width, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(_depth, UnitTypeId.Millimeters);
            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(x-dx, y-dy, 0));
            points.Add(new XYZ(x+dx, y-dy, 0));
            points.Add(new XYZ(x+dx, y+dy, 0));
            points.Add(new XYZ(x-dx, y+dy, 0));
            points.Add(new XYZ(x-dx, y-dy, 0));

            List<Wall> walls = new List<Wall>();

            Transaction transaction = new Transaction(doc, "Построение стен");
            transaction.Start();
            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, baseLevel.Id, false);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(upperLevel.Id);
                walls.Add(wall);
            }

            transaction.Commit();

            return walls;
        }
    }
}
