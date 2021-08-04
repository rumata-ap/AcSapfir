using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
using SapfirLib;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[assembly: CommandClass(typeof(AcSapfir.SapfirDrafter))]
namespace AcSapfir
{
   public class SapfirDrafter
   {
      private SapfirLib.Application appSpf;
      private SapfirDoc docSpf;
      private AutoProject projSpf;
      private AutoStorey storeySpf;
      private AutoFactory factorySpf;
      private AutoModel paramModelSpf;
      private AutoModel3D drawModelSpf;
      private AutoMultiContour sectionSpf;
      private AutoLine lineSpf;
      private AutoPolyLine polylineSpf;
      private ModelsTypes modelType;
      private Models3dTypes model3DType;
      private AutoVectContour contSpf;
      //private AcSsetsBuilder ssBld = new AcSsetsBuilder();
      //private AcDrafter acadDrafter = new AcDrafter();
      private object buf1;
      private object buf2;
      private object buf3;

      public SapfirDrafter()
      {
         //buf1 = Array();
         //buf2 = Array();
         //buf3 = Array();

         appSpf = new SapfirLib.Application(); //CreateObject("Sapfir.Application")
         appSpf.Visible = 1;
         docSpf = appSpf.GetActiveDoc();

         if (docSpf.CountProjects == 0) projSpf = docSpf.NewProject();
         else projSpf = docSpf.GetActiveProject();

         if (projSpf.CountStorey == 0) storeySpf = projSpf.NewStorey("1-й этаж");
         else storeySpf = docSpf.GetActiveStorey();
      }

      void CreateWalls(Line line, double t)
      {
         //paramModelSpf = storeySpf.GetModelByID(97);
         paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_WALL);
         polylineSpf = paramModelSpf.GetAxisLine();
         //buf2 = new object[6];
         //int res = polylineSpf.GetPoints(ref buf2);
         buf1 = new object[] { line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0, line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0 };
         polylineSpf.SetPoints(buf1);
         paramModelSpf.Parameter["M_THICKNESS"] = t;
         paramModelSpf.Parameter["M_CONSTYPE"] = "M_CTP_BEAR";
         paramModelSpf.Parameter["M_CR_FILL"] = 11854048;
         paramModelSpf.Parameter["M_CR_LINE"] = 7369618;
         paramModelSpf.Parameter["M_MATERIAL"] = "ab7d2fa2-df44-4d8f-8d46-71b7fb919bde";
         paramModelSpf.RegenModel();
      }

      void CreateSlabs(Polyline line, double t)
      {
         paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_SLAB);
         polylineSpf = paramModelSpf.GetAxisLine();
         buf1 = new object[line.NumberOfVertices * 3];
         object[] crd = (object[])buf1;
         for (int i = 0; i < line.NumberOfVertices; i++)
         {
            Point2d point = line.GetPoint2dAt(i);
            crd[3 * i] = point.X * 0.001;
            crd[3 * i + 1] = point.Y * 0.001;
            crd[3 * i + 2] = 0;
         }
         polylineSpf.SetPoints(crd);
         paramModelSpf.Parameter["M_THICKNESS"] = t;
         paramModelSpf.Parameter["M_CONSTYPE"] = "M_CTP_BEAR";
         paramModelSpf.Parameter["M_CR_FILL"] = 12832210;
         paramModelSpf.Parameter["M_CR_LINE"] = 2318692;
         paramModelSpf.Parameter["M_MATERIAL"] = "79f404bd-50e9-4058-b664-c159ffdd0ce8";
         paramModelSpf.Parameter["M_TYPE_LEVEL"] = "M_LEVEL_UP";
         paramModelSpf.RegenModel();
      }

      void CreateSlabHoles(Polyline line, int ids)
      {
         paramModelSpf = storeySpf.GetModelByID(ids);
         AutoModel holeModelSp = paramModelSpf.NewHole((int)ModelsTypes.TM_HOLE);
         polylineSpf = holeModelSp.GetAxisLine();
         buf1 = new object[line.NumberOfVertices * 3];
         object[] crd = (object[])buf1;
         for (int i = 0; i < line.NumberOfVertices; i++)
         {
            Point2d point = line.GetPoint2dAt(i);
            crd[3 * i] = point.X * 0.001;
            crd[3 * i + 1] = point.Y * 0.001;
            crd[3 * i + 2] = 0;
         }
         polylineSpf.SetPoints(crd);
         holeModelSp.SetPosition((double)crd[0], (double)crd[1], 0);

         paramModelSpf.RegenModel();
      }

      void CreateFoundSlabs(Polyline line, double t)
      {
         paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_SLAB_1);
         polylineSpf = paramModelSpf.GetAxisLine();
         buf1 = new object[line.NumberOfVertices * 3];
         object[] crd = (object[])buf1;
         for (int i = 0; i < line.NumberOfVertices; i++)
         {
            Point2d point = line.GetPoint2dAt(i);
            crd[3 * i] = point.X * 0.001;
            crd[3 * i + 1] = point.Y * 0.001;
            crd[3 * i + 2] = 0;
         }
         polylineSpf.SetPoints(crd);
         paramModelSpf.Parameter["M_THICKNESS"] = t;
         paramModelSpf.Parameter["M_CONSTYPE"] = "M_CTP_BEAR";
         paramModelSpf.Parameter["M_CR_FILL"] = 9876679;
         paramModelSpf.Parameter["M_CR_LINE"] = 2785991;
         paramModelSpf.Parameter["M_MATERIAL"] = "239d15bb-4253-4546-89f1-2d90300a3c79";
         paramModelSpf.Parameter["M_TYPE_LEVEL"] = "M_LEVEL_DN";
         paramModelSpf.RegenModel();
      }

      List<int> GetWalls(out int count)
      {
         List<int> ids = new List<int>();
         for (int i = 0; i < storeySpf.CountModel; i++)
         {
            paramModelSpf = storeySpf.GetModelByIndex(i);
            if (paramModelSpf.TypeModel == 10) ids.Add(paramModelSpf.ID);
         }
         count = ids.Count;
         return ids;
      }

      void CreateDoors(Line line, IEnumerable<int> ids, double h)
      {
         foreach (int item in ids)
         {
            paramModelSpf = storeySpf.GetModelByID(item);
            polylineSpf = paramModelSpf.GetAxisLine();
            lineSpf = polylineSpf.GetLine(0);
            buf1 = new object[6];
            int res = lineSpf.GetPoints(ref buf1);
            object[] crd = (object[])buf1;
            Line2d l = new Line2d(new Point2d((double)crd[0] * 1000, (double)crd[1] * 1000), new Point2d((double)crd[3] * 1000, (double)crd[4] * 1000));
            Line2d d = new Line2d(new Point2d(line.StartPoint.X, line.StartPoint.Y), new Point2d(line.EndPoint.X, line.EndPoint.Y));
            Tolerance tol = new Tolerance(1E-8, 1E-8);

            if (l.IsColinearTo(d, tol))
            {
               modelType = ModelsTypes.TM_DOOR;
               AutoModel modelDoor = paramModelSpf.NewHole((int)modelType);
               modelDoor.Parameter["B"] = line.Length * 0.001;
               modelDoor.Parameter["H"] = h;
               Point2d lsp = new Point2d((double)crd[0] * 1000, (double)crd[1] * 1000);
               Point2d dsp = new Point2d(line.StartPoint.X, line.StartPoint.Y);
               Point2d dep = new Point2d(line.EndPoint.X, line.EndPoint.Y);
               modelDoor.Parameter["M_POS"] = Math.Min(lsp.GetDistanceTo(dsp), lsp.GetDistanceTo(dep)) * 0.001;

               //modelDoor.RegenModel();
               //paramModelSpf.RegenModel();
            }
         }
      }

      void CreatePointLoads(DBPoint point, double p1)
      {
         paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_LOAD_1);
         paramModelSpf.SetPosition(point.Position.X * 0.001, point.Position.Y * 0.001, 0);
         paramModelSpf.Parameter["M_LOAD_1"] = p1;
         paramModelSpf.RegenModel();
      }

      void CreateLineLoads(Line line, double p1)
      {
         AutoParametersDlg dlg = new AutoParametersDlg();
         dlg.DoModal();
         paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_LOAD_2);
         polylineSpf = paramModelSpf.GetAxisLine();
         buf1 = new object[] { line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0, line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0 };
         polylineSpf.AddLine((int)Models3dTypes.TM3_LINE, buf1);
         paramModelSpf.Parameter["M_LOAD_1"] = p1;
         paramModelSpf.Parameter["M_LOAD_2"] = p1;
         paramModelSpf.RegenModel();
      }
      
      void CreateAxes(Line line, int num)
      {         
         buf1 = new object[] { line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0, line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0 };
         AutoObjDim axObj = storeySpf.NewModel((int)ModelsTypes.TM_DIMENSION);
         axObj.SetDimParam((int)Models3dTypes.DIM_AXIS, (int)Models3dTypes.DIM_DIR_XY, num.ToString(), buf1);
      }
      
      void CreateLineLoads(Line line, double p1, double p2)
      {
         paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_LOAD_2);
         polylineSpf = paramModelSpf.GetAxisLine();
         buf1 = new object[] { line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0, line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0 };
         polylineSpf.AddLine((int)Models3dTypes.TM3_LINE, buf1);
         paramModelSpf.Parameter["M_LOAD_1"] = p1;
         paramModelSpf.Parameter["M_LOAD_2"] = p2;
         paramModelSpf.RegenModel();
      }

      [CommandMethod("Sapfir_AXES", CommandFlags.UsePickSet)]
      public void Sapfir_AXES()
      {
         AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateAxes);
      }
      
      [CommandMethod("Sapfir_Point_LOADS", CommandFlags.UsePickSet)]
      public void Sapfir_Point_LOADS()
      {
         Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
         PromptDoubleResult result = acDocEd.GetDouble("Введите значение нагрузки в тоннах: ");
         if (result.Status == PromptStatus.OK) AcUtilites.ActionOnPoints(AcUtilites.Selection(), CreatePointLoads, result.Value);
      }
      
      [CommandMethod("Sapfir_Line_LOADS", CommandFlags.UsePickSet)]
      public void Sapfir_Line_LOADS()
      {
         Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
         PromptDoubleResult result = acDocEd.GetDouble("Введите значение нагрузки в тоннах: ");
         if (result.Status == PromptStatus.OK) AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateLineLoads, result.Value);
      }
      
      [CommandMethod("Sapfir_Line_LOADS2", CommandFlags.UsePickSet)]
      public void Sapfir_Line_LOADS2()
      {
         Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
         PromptDoubleResult result = acDocEd.GetDouble("Введите значение нагрузки в начале в тоннах: ");
         PromptDoubleResult result1 = acDocEd.GetDouble("Введите значение нагрузки в конце в тоннах: ");
         if (result.Status == PromptStatus.OK && result1.Status == PromptStatus.OK)
         {
            AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateLineLoads, result.Value, result1.Value);
         }
      }

      [CommandMethod("Sapfir_SLABS", CommandFlags.UsePickSet)]
      public void Sapfir_SLABS()
      {
         Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
         PromptDoubleResult result = acDocEd.GetDouble("Введите толщину плит в метрах: ");
         if (result.Status == PromptStatus.OK) AcUtilites.ActionOnPolylines(AcUtilites.Selection(), CreateSlabs, result.Value);     
      }

      [CommandMethod("Sapfir_SLAB_Holes", CommandFlags.UsePickSet)]
      public void Sapfir_SLAB_holes()
      {
         Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
         PromptIntegerResult result = acDocEd.GetInteger("Введите идентификатор (ID) плиты для размещения отверстий: ");
         if (result.Status == PromptStatus.OK) AcUtilites.ActionOnPolylines(AcUtilites.Selection(), CreateSlabHoles, result.Value);     
      }

      [CommandMethod("Sapfir_Found_SLABS", CommandFlags.UsePickSet)]
      public void Sapfir_Found_SLABS()
      {
         Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
         PromptDoubleResult result = acDocEd.GetDouble("Введите толщину фундаментных плит в метрах: ");
         if (result.Status == PromptStatus.OK) AcUtilites.ActionOnPolylines(AcUtilites.Selection(), CreateFoundSlabs, result.Value);       
      }

      [CommandMethod("Sapfir_WALLS", CommandFlags.UsePickSet)]
      public void Sapfir_WALLS()
      {
         Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
         PromptDoubleResult result = acDocEd.GetDouble("Введите толщину стен в метрах: ");
         if (result.Status == PromptStatus.OK) AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateWalls, result.Value);      
      }

      [CommandMethod("Sapfir_DOORS", CommandFlags.UsePickSet)]
      public void Sapfir_doors()
      {
         Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
         PromptDoubleResult result = acDocEd.GetDouble("Введите высоту проема в метрах: ");
         if (result.Status == PromptStatus.OK) AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateDoors, GetWalls(out int cnt), result.Value);     
      }

      [CommandMethod("Sapfir_test", CommandFlags.UsePickSet)]
      public void Sapfir_test()
      {
         
      }
   }

   // типы 3D моделей
   internal enum Models3dTypes
   {
      TM3_LINE = 1, // 1 отрезок
      TM3_ARC = 2, // 2 дуга
      TM3_BEZIER = 3, // 3 Безье
      TM3_POLYLINE = 4, // 4 полилиния
      TM3_MESH3D = 5, // 5 гранная модель трёхмерного объекта
      TM3_MODEL3D = 6, // 6 модель 3D
      TM3_TEXT3D = 7, // 7 текст 3D
      TM3_FACE = 8, // 8 грань меша
      TM3_DIM3D = 9, // 9 обозначение размера

      TM3_TRIANGLE = 10, // 10 пространственный треугольник
      TM3_DIMENSION = 11, // 11 модель размера
      TM3_POLY = 12, // 12 полигон/штриховка
      TM3_AXIS = 13, // 13 ось с маркировкой в кружочке
      TM3_FE_BAR = 14, // 14 - конечный элемент стержень
      TM3_FE_TRIANGLE = 15, // 15 - конечный элемент треугольник
      TM3_FE_QUAD = 16, // 16 - конечный элемент четырёхугольник
      TM3_FE_TETRAEDR = 17, // 17 - конечный элемент тетраэдр
      TM3_FE_PRYSM_3 = 18, // 18 - конечный элемент трёхгранная призма
      TM3_FE_PRYSM_4 = 19, // 19 - конечный элемент универсальный объёмный восьмиузловой
      TM3_FE_SUPER = 20, // 20 - конечный элемент суперэлемент
      TM3_ELLIPSE = 21, // 21 - эллипс (дуга эллипса)
      TM3_POINT = 22, // 22 - трёхмерная точка (маркер)
      TM_DRAFT = 100, // 100 - чертеж

      // Параметры размеров
      DIM_ELEV = 1,
      DIM_LINEAR = 2,
      DIM_CHAIN = 3,
      DIM_RADIAL = 4,
      DIM_ANGULAR = 5,
      DIM_NOTE = 6,
      DIM_DIAMETR = 7,
      DIM_AXIS = 8,
      DIM_POINT = 9,
      DIM_ARC = 10,
      DIM_MARKER_CIRCLE = 11,

      DIM_DIR_X = 1,
      DIM_DIR_Y = 2,
      DIM_DIR_D = 3,
      DIM_DIR_XY = 3,
      DIM_DIR_Z = 4
   }

   // Типы параметрических объектов
   internal enum ModelsTypes
   {
      TM_NONE = 0, // Неопределенный тип модели
      TM_SITE = 1, // Участок (массив зданий) -
      TM_PROJECT = 3, // Здание (проект)
      TM_STOREY = 4, // Этаж
      TM_BLOCK = 5, // БЛОК
      TM_FEAPROJECT = 6, // Конструкции
      TM_NODE = 7, // Сборочный узел

      TM_WALL = 10, // Стена
      TM_WALL_1 = 65546, // Перегородка
      TM_SLAB = 11, // Перекрытие
      TM_SLAB_1 = 65547, // Фундаментая плита
      TM_COLUMN = 12, // Колонна
      TM_BEAM = 13, // Балка
      TM_BAR = 14, // Стержень
      TM_ROOF = 15, // Крыша
      TM_STAIRS = 16, // Лестница
      TM_AXES = 17, // Оси координационные
      TM_LINE = 18, // Линия (построения)
      TM_ZONE = 19, // Зона (помещение)
      TM_DIMENSION = 20, // Размер
      TM_LOAD = 21, // нагрузка сила
      TM_MOMENT = 22, // нагрузка момент
      TM_CAPITAL = 23, // Капитель
      TM_REF = 24, // ссылочный объект (фантом)
      TM_UNDERCAP = 25, // подколонник
      TM_TRUSS = 26, // ферма

      TM_LOAD_1 = 65557, // Точечная нагрузка
      TM_LOAD_2 = 131093, // Линейная нагрузка
      TM_LOAD_4 = 262165, // Штамп нагрузки
                          //TM_STAIRS_1 = "E09D069F-22FA-4aa8-B788-AF7E558F1798" ' Одномаршевая лестница

      TM_RECESS = 29, // Ниша
      TM_HOLE = 30, // Проем
      TM_WINDOW = 31, // Окно
      TM_DOOR = 32, // Дверь
      TM_MODVISION = 33, // контейнер моделей для видов документирования (разрезы, планы, фасады...)
      TM_MODVISIONSECTION = 34, // для вида "вертикальный разрез/сечение"
      TM_MODVISIONPLAN = 35, // для вида "план"
      TM_MODVISION3D = 36, // Для вида "3D"
      TM_MODVISIONCONSTRUCT = 37, // Для вида "Конструирование"

      TM_SPHERE = 60, // Сфера
      TM_PIVOTAL = 61, // Поверхнсть вращения
      TM_PRISM = 62, // Призма
      TM_PROFILE = 63, // Профиль (кинематическая поверхность)
      TM_CONE = 64, // Конус
      TM_HIPPAR = 65, // Гиперболический параболоид
      TM_LIGHT = 66, // Источник света
      TM_POLY = 67, // заштрихованный полигон
      TM_TEXT = 68, // текст
      TM_CYLINDER = 69, // Цилиндрическая поверхность
      TM_SURFACE = 70, // Кинематические поверхности: генерация по траекториям и образующим

      TM_WIND = 71, // Ветер для приложения ветровой нагрузки
      TM_ARM_AREA = 72, // Армирование плиты (обозначения фонового армирования и участков раскладки дополнительной арматуры)
      TM_ARM_SLAB = 73, // Армирование плиты, результаты расчёта, КЭ, мозаики и шкалы, параметры и т.п.
      TM_KARKAS = 74, // Армирование плиты, каркас поперечной арматуры
      TM_PUNCH = 75, // Армирование плиты, узел продавливания
      TM_ARM_WALL = 76, // Армирование стены: результаты расчёта, КЭ, мозаики и шкалы, параметры
      TM_ARM_BAR = 77, // Арматурный стержень: прямой или изогнутый (Г-элемент, П-элемент и т.п.)
      TM_ARM_ZONE = 78, // Зона армирования диафрагмы(стены) отдельными арматурными стержнями (возможно, отогнутыми)
      TM_ARM_COLUMN = 79, // Армирование колонны: результаты расчёта, параметры
      TM_FEASCHEMA = 80, // Расчетная модель
      TM_FEABAR = 81, // Идеализированная модель стержня
      TM_FEASHELL = 82, // Идеализированная модель оболочки
      TM_MULTICONT = 83, // массив параметрических контуров - CAutoMultiContour

      TM_OTHER = 128 // Иные Объекты
   }
   // Управление именованными параметрами
   internal enum ParameterOptionsTypes
   {
      NPA_GROUP_OPEN = 65536,
      NPA_GROUP_CLOSE = 131072,
      NPA_READONLY = 262144,
      //NPA_OPTIONAL = long.Parse("&H00080000"),
      //NPA_UNDEFINED = long.Parse("&H00100000"),
      //NPA_VALUE_LIST = long.Parse("&H00200000"),
      //NPA_SPIN_CONTROL = long.Parse("&H00400000"),
      //NPA_SERVICE = long.Parse("&H00800000"),
      //NPA_DONT_SAVE = long.Parse("&H01000000"),
      //NPA_AUTO_REGEN = long.Parse("&H02000000"), //изменение параметра приводит к необходимости автоматической регенерации при пом. скрипта
      //NPA_HIDDEN = long.Parse("&H10000000"), // скрытый параметр - пропускается в функции GetParameters если не разрешается брать скрытые параметры (по умочанию - не разрешается)
      //NPA_HIDE_DLG = long.Parse("&H40000000"), // для параметра имеющего свойство NPA_AUTO_REGEN - функцию регенерации FunRegen выполнять без открытия HTML диалога
      //NPA_PRIVATE = long.Parse("&H08000000"), // параметры для частного использования объектом (не отображается в списке общим способом)

      //NPA_PAR_TYPE = long.Parse("&H000000FF"),
      //NPA_PAR_UNIT = long.Parse("&H0000FF00"),

      //NPA_PT_STRING = long.Parse("&H01"),
      //NPA_PT_INT = long.Parse("&H03"),
      //NPA_PT_FLOAT = long.Parse("&H04"),
      //NPA_PT_COLORREF = long.Parse("&H05"),
      //NPA_PT_FOLDER = long.Parse("&H06"),
      //NPA_PT_BOOLEAN = long.Parse("&H07"),
      //NPA_PT_HEADER = long.Parse("&H08"),
      //NPA_PT_MNEMO = long.Parse("&H09"),
      //NPA_PT_FILEPATH = long.Parse("&H0A"),
      //NPA_PT_FILEPATH_3D = long.Parse("&H1A"),
      //NPA_PT_TYPELIBPATH = long.Parse("&H1B"), // путь в папку с моделями, используется совместно с: NPA_PT_FILEPATH, NPA_PT_FILEPATH_3D, NPA_PT_FILEPATH_IMAGE
      //NPA_PU_LOADAREA = long.Parse("&H2600"), // Нагрузка распределённая:

      //NPA_PT_FILEPATH_IMAGE = long.Parse("&H2A"),
      //NPA_PT_RGBA = long.Parse("&H0B"),
      //NPA_PT_COLORMAT = long.Parse("&H0C"),
      //NPA_PT_SLIDER = long.Parse("&H0D"),
      //NPA_PT_MATERIAL = long.Parse("&H0E"),
      //NPA_PT_FONT = long.Parse("&H0F"),
      ////NPA_PT_FILEFILTER = CLng("&H10")
      //NPA_PT_LAYER = long.Parse("&H11"),
      //NPA_PT_LAYER_COMBI = long.Parse("&H12"),
      //NPA_PT_SCALE = long.Parse("&H13"),
      //NPA_PT_SORTAMENT = long.Parse("&H14"),
      //NPA_PT_GROUPDATA = long.Parse("&H18"),

      OBP_NAMES = 4,
      OBP_COMMENTS = 8,
      OBP_VALUES = 16,
      //OBP_SET_SERIALIZE = long.Parse("&H40"),
      //OBP_GET_HIDDEN = long.Parse("&H100"), // получить скрытые параметры (автоматом включается для скриптов)

      //DEL_ALL = long.Parse("&H0"), //- удалить все параметры
      //DEL_INPUT = long.Parse("&H1"), //- удалить заданные параметры
      //DEL_RANGE = long.Parse("&H2"), //- удалить все между заданными параметрами и задано два параметра
      //DEL_NOT_PRIVATE = long.Parse("&H4"), //- удалить все кроме приватных
      //DEL_AUTO_REGEN = long.Parse("&H8"), //- удалить все AUTO_REGEN

      //// Для одномерных величин
      ////  (длины малые, средние, большие):
      //NPA_PU_D1_S = long.Parse("&H1100"),
      //NPA_PU_D1_M = long.Parse("&H2100"),
      //NPA_PU_D1_L = long.Parse("&H3100"),

      //// Для двумерных величин
      ////  (площади малые, средние, большие):
      //NPA_PU_D2_S = long.Parse("&H1200"),
      //NPA_PU_D2_M = long.Parse("&H2200"),
      //NPA_PU_D2_L = long.Parse("&H3200"),

      //// Для объёмных величин
      ////  (объёмы малые, средние, большие):
      //NPA_PU_D3_S = long.Parse("&H1300"),
      //NPA_PU_D3_M = long.Parse("&H2300"),
      //NPA_PU_D3_L = long.Parse("&H3300"),

      //// Угловые величины:
      //NPA_PU_ANGLE = long.Parse("&H0400"),

      //// Временные интервалы:
      //NPA_PU_TIME = long.Parse("&H0500"),

      //// Силы:
      //NPA_PU_FORCE = long.Parse("&H0600"),

      //// Масса:
      //NPA_PU_MASS = long.Parse("&H0700")

   }

}
