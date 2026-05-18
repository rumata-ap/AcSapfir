using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
using SapfirLib;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

[assembly: CommandClass(typeof(AcSapfir.SapfirDrafter))]

namespace AcSapfir
{
    /// <summary>
    /// Основной класс для создания объектов в Sapfir из геометрии AutoCAD.
    /// Содержит команды AutoCAD ([CommandMethod]) и логику взаимодействия
    /// с библиотекой SapfirLib (COM-автоматизация САПФИР).
    /// </summary>
    /// <remarks>
    /// При каждом вызове команды AutoCAD создаётся новый экземпляр класса,
    /// который устанавливает соединение с Sapfir через COM.
    /// Координаты масштабируются из миллиметров (AutoCAD) в метры (Sapfir)
    /// с коэффициентом 0.001.
    /// </remarks>
    public class SapfirDrafter
    {
        /// <summary>Приложение Sapfir (COM-объект).</summary>
        private SapfirLib.Application appSpf;

        /// <summary>Активный документ Sapfir.</summary>
        private SapfirDoc docSpf;

        /// <summary>Активный проект (здание) Sapfir.</summary>
        private AutoProject projSpf;

        /// <summary>Активный этаж Sapfir, на котором создаются объекты.</summary>
        private AutoStorey storeySpf;

        /// <summary>Текущая параметрическая модель Sapfir (стена, плита и т.д.).</summary>
        private AutoModel paramModelSpf;

        /// <summary>Текущая осевая линия Sapfir.</summary>
        private AutoPolyLine polylineSpf;

        /// <summary>Буфер для обмена координатами с SapfirLib.</summary>
        private object buf1;

        /// <summary>Выбранный пользователем GUID материала для текущей команды.</summary>
        private string selectedMaterialGuid;

        string SelectMaterial(string elementType)
        {
            var form = new MaterialSelectorForm(docSpf, elementType);
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(form);
            if (form.DialogResult == System.Windows.Forms.DialogResult.OK && form.SelectedGuid != null)
            {
                selectedMaterialGuid = form.SelectedGuid;
                AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                    "\n  Материал: {0}", form.SelectedName);
                return form.SelectedGuid;
            }
            selectedMaterialGuid = null;
            return null;
        }

        /// <summary>
        /// Инициализирует новый экземпляр класса <see cref="SapfirDrafter"/>.
        /// Устанавливает соединение с приложением Sapfir, получает активный документ,
        /// проект и этаж, или создаёт их, если они отсутствуют.
        /// </summary>
        public SapfirDrafter()
        {
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                appSpf = new SapfirLib.Application();
                appSpf.Visible = 1;

                var activeView = appSpf.GetActiveSapfirView();
                if (activeView != null)
                {
                    docSpf = activeView.GetDocument();
                }

                if (docSpf == null)
                {
                    docSpf = appSpf.GetActiveDoc();
                }

                if (docSpf == null) docSpf = appSpf.NewDocument();

                if (docSpf.CountProjects == 0) projSpf = docSpf.NewProject();
                else projSpf = docSpf.GetActiveProject();

                if (projSpf.CountStorey == 0) storeySpf = projSpf.NewStorey("1-й этаж");
                else storeySpf = docSpf.GetActiveStorey();
            }
            catch (System.Exception ex)
            {
                acDocEd.WriteMessage("\nОшибка подключения к САПФИР: " + ex.Message);
                acDocEd.WriteMessage("\nУбедитесь, что САПФИР запущен и в нём открыт документ.");
                throw;
            }
        }

        // ─────────────────────────────────────────────────────────────
        //  Создание конструктивных элементов Sapfir
        // ─────────────────────────────────────────────────────────────

        /// <summary>
        /// Создаёт стены в Sapfir на основе линии AutoCAD.
        /// Координаты масштабируются из миллиметров в метры (×0.001).
        /// </summary>
        /// <param name="line">Линия AutoCAD, представляющая ось стены.</param>
        /// <param name="t">Толщина стены в метрах.</param>
        void CreateWalls(Line line, double t)
        {
            paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_WALL);
            polylineSpf = paramModelSpf.GetAxisLine();
            buf1 = new object[]
            {
                line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0,
                line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0
            };
            polylineSpf.SetPoints(buf1);
            paramModelSpf.Parameter["M_THICKNESS"] = t;
            paramModelSpf.Parameter["M_CONSTYPE"] = "M_CTP_BEAR";
            paramModelSpf.Parameter["M_CR_FILL"] = 11854048;
            paramModelSpf.Parameter["M_CR_LINE"] = 7369618;
            if (selectedMaterialGuid != null)
                paramModelSpf.Parameter["M_MATERIAL"] = selectedMaterialGuid;
            else
                paramModelSpf.Parameter["M_MATERIAL"] = "ab7d2fa2-df44-4d8f-8d46-71b7fb919bde";
            paramModelSpf.RegenModel();
            AcUtilites.SetSapfirId(line, paramModelSpf.ID);
            AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                "\n  Стена ID={0}", paramModelSpf.ID);
        }

        void CreateLines(Line line, double _)
        {
            paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_LINE);
            polylineSpf = paramModelSpf.GetAxisLine();
            buf1 = new object[]
            {
                line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0,
                line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0
            };
            polylineSpf.AddLine((int)Models3dTypes.TM3_LINE, buf1);
            paramModelSpf.RegenModel();
            AcUtilites.SetSapfirId(line, paramModelSpf.ID);
            AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                "\n  Линия ID={0}", paramModelSpf.ID);
        }

        /// <summary>
        /// Создаёт плиты перекрытия в Sapfir на основе полилинии AutoCAD.
        /// Координаты масштабируются из миллиметров в метры (×0.001).
        /// </summary>
        /// <param name="line">Полилиния AutoCAD, представляющая контур плиты.</param>
        /// <param name="t">Толщина плиты в метрах.</param>
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
            paramModelSpf.Parameter["M_MATERIAL"] = selectedMaterialGuid ?? "79f404bd-50e9-4058-b664-c159ffdd0ce8";
            paramModelSpf.Parameter["M_TYPE_LEVEL"] = "M_LEVEL_UP";
            paramModelSpf.RegenModel();
            AcUtilites.SetSapfirId(line, paramModelSpf.ID);
            AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                "\n  Плита ID={0}", paramModelSpf.ID);
        }

        /// <summary>
        /// Создаёт отверстия в плитах Sapfir на основе полилинии AutoCAD.
        /// Координаты масштабируются из миллиметров в метры (×0.001).
        /// </summary>
        /// <param name="line">Полилиния AutoCAD, представляющая контур отверстия.</param>
        /// <param name="ids">Идентификатор (ID) плиты, в которой создаётся отверстие.</param>
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

        /// <summary>
        /// Создаёт фундаментные плиты в Sapfir на основе полилинии AutoCAD.
        /// Координаты масштабируются из миллиметров в метры (×0.001).
        /// </summary>
        /// <param name="line">Полилиния AutoCAD, представляющая контур фундаментной плиты.</param>
        /// <param name="t">Толщина фундаментной плиты в метрах.</param>
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
            paramModelSpf.Parameter["M_MATERIAL"] = selectedMaterialGuid ?? "239d15bb-4253-4546-89f1-2d90300a3c79";
            paramModelSpf.Parameter["M_TYPE_LEVEL"] = "M_LEVEL_DN";
            paramModelSpf.RegenModel();
            AcUtilites.SetSapfirId(line, paramModelSpf.ID);
            AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                "\n  Фунд.плита ID={0}", paramModelSpf.ID);
        }

        /// <summary>
        /// Получает список идентификаторов (ID) стен, созданных в текущем этаже Sapfir.
        /// </summary>
        /// <param name="count">Выходной параметр, содержащий количество найденных стен.</param>
        /// <returns>Список идентификаторов стен.</returns>
        List<int> GetWalls(out int count)
        {
            List<int> ids = new List<int>();
            for (int i = 0; i < storeySpf.CountModel; i++)
            {
                paramModelSpf = storeySpf.GetModelByIndex(i);
                if (paramModelSpf.TypeModel == (int)ModelsTypes.TM_WALL) ids.Add(paramModelSpf.ID);
            }
            count = ids.Count;
            return ids;
        }

        /// <summary>
        /// Создаёт двери в стенах Sapfir на основе линии AutoCAD, представляющей проём.
        /// Ищет коллинеарные стены и создаёт в них проёмы типа «дверь».
        /// Координаты масштабируются из миллиметров в метры (×0.001).
        /// </summary>
        /// <param name="line">Линия AutoCAD, представляющая проём двери.</param>
        /// <param name="ids">Коллекция идентификаторов стен, в которых следует искать место для двери.</param>
        /// <param name="h">Высота проёма двери в метрах.</param>
        void CreateDoors(Line line, IEnumerable<int> ids, double h)
        {
            foreach (int item in ids)
            {
                paramModelSpf = storeySpf.GetModelByID(item);
                polylineSpf = paramModelSpf.GetAxisLine();
                AutoLine spfLine = polylineSpf.GetLine(0);
                buf1 = new object[6];
                int res = spfLine.GetPoints(ref buf1);
                object[] crd = (object[])buf1;

                Line2d wallLine = new Line2d(
                    new Point2d((double)crd[0] * 1000, (double)crd[1] * 1000),
                    new Point2d((double)crd[3] * 1000, (double)crd[4] * 1000));
                Line2d doorLine = new Line2d(
                    new Point2d(line.StartPoint.X, line.StartPoint.Y),
                    new Point2d(line.EndPoint.X, line.EndPoint.Y));
                Tolerance tol = new Tolerance(1E-8, 1E-8);

                if (wallLine.IsColinearTo(doorLine, tol))
                {
                    AutoModel modelDoor = paramModelSpf.NewHole((int)ModelsTypes.TM_DOOR);
                    modelDoor.Parameter["B"] = line.Length * 0.001;
                    modelDoor.Parameter["H"] = h;
                    Point2d wallStart = new Point2d((double)crd[0] * 1000, (double)crd[1] * 1000);
                    Point2d doorStart = new Point2d(line.StartPoint.X, line.StartPoint.Y);
                    Point2d doorEnd = new Point2d(line.EndPoint.X, line.EndPoint.Y);
                    modelDoor.Parameter["M_POS"] = Math.Min(wallStart.GetDistanceTo(doorStart), wallStart.GetDistanceTo(doorEnd)) * 0.001;
                }
            }
        }

        /// <summary>
        /// Создаёт точечные нагрузки в Sapfir на основе точки AutoCAD.
        /// Координаты масштабируются из миллиметров в метры (×0.001).
        /// </summary>
        /// <param name="point">Точка AutoCAD, обозначающая место приложения нагрузки.</param>
        /// <param name="p1">Значение точечной нагрузки.</param>
        void CreatePointLoads(DBPoint point, double p1)
        {
            paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_LOAD_1);
            paramModelSpf.SetPosition(point.Position.X * 0.001, point.Position.Y * 0.001, 0);
            paramModelSpf.Parameter["M_LOAD_1"] = p1;
            paramModelSpf.RegenModel();
        }

        /// <summary>
        /// Создаёт линейные нагрузки в Sapfir на основе линии AutoCAD
        /// с равномерным распределением (одинаковое значение в начале и конце).
        /// Координаты масштабируются из миллиметров в метры (×0.001).
        /// </summary>
        /// <param name="line">Линия AutoCAD, представляющая линию приложения нагрузки.</param>
        /// <param name="p1">Равномерно распределённая нагрузка на линии.</param>
        void CreateLineLoads(Line line, double p1)
        {
            AutoParametersDlg dlg = new AutoParametersDlg();
            dlg.DoModal();
            paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_LOAD_2);
            polylineSpf = paramModelSpf.GetAxisLine();
            buf1 = new object[]
            {
                line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0,
                line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0
            };
            polylineSpf.AddLine((int)Models3dTypes.TM3_LINE, buf1);
            paramModelSpf.Parameter["M_LOAD_1"] = p1;
            paramModelSpf.Parameter["M_LOAD_2"] = p1;
            paramModelSpf.RegenModel();
        }

        /// <summary>
        /// Создаёт координационные оси в Sapfir на основе линии AutoCAD.
        /// Координаты масштабируются из миллиметров в метры (×0.001).
        /// </summary>
        /// <param name="line">Линия AutoCAD, представляющая ось.</param>
        /// <param name="num">Номер оси (порядковый индекс при выборе).</param>
        void CreateAxes(Line line, int num)
        {
            buf1 = new object[]
            {
                line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0,
                line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0
            };
            AutoObjDim axObj = storeySpf.NewModel((int)ModelsTypes.TM_DIMENSION);
            axObj.SetDimParam((int)Models3dTypes.DIM_AXIS, (int)Models3dTypes.DIM_DIR_XY, num.ToString(), buf1);
        }

        /// <summary>
        /// Создаёт линейные нагрузки в Sapfir с разными значениями в начале и конце линии
        /// (трапециевидная эпюра).
        /// Координаты масштабируются из миллиметров в метры (×0.001).
        /// </summary>
        /// <param name="line">Линия AutoCAD, представляющая линию приложения нагрузки.</param>
        /// <param name="p1">Нагрузка в начале линии.</param>
        /// <param name="p2">Нагрузка в конце линии.</param>
        void CreateLineLoads(Line line, double p1, double p2)
        {
            paramModelSpf = storeySpf.NewModel((int)ModelsTypes.TM_LOAD_2);
            polylineSpf = paramModelSpf.GetAxisLine();
            buf1 = new object[]
            {
                line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0,
                line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0
            };
            polylineSpf.AddLine((int)Models3dTypes.TM3_LINE, buf1);
            paramModelSpf.Parameter["M_LOAD_1"] = p1;
            paramModelSpf.Parameter["M_LOAD_2"] = p2;
            paramModelSpf.RegenModel();
        }

        // ─────────────────────────────────────────────────────────────
        //  Вспомогательные методы
        // ─────────────────────────────────────────────────────────────

        void ResolveActiveContext()
        {
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;

            var activeView = appSpf.GetActiveSapfirView();
            if (activeView != null)
                docSpf = activeView.GetDocument();

            if (docSpf == null)
                docSpf = appSpf.GetActiveDoc();

            if (docSpf == null)
                docSpf = appSpf.NewDocument();

            if (docSpf.CountProjects == 0)
                projSpf = docSpf.NewProject();
            else
                projSpf = docSpf.GetActiveProject();

            if (projSpf.CountStorey == 0)
                storeySpf = projSpf.NewStorey("1-й этаж");
            else
                storeySpf = docSpf.GetActiveStorey();

            acDocEd.WriteMessage("\n[Sapfir] Активно: {0} / проект ID={1} / этажей: {2}",
                docSpf.Title, projSpf.ID, projSpf.CountStorey);
        }

        // ─────────────────────────────────────────────────────────────
        //  Команды AutoCAD
        // ─────────────────────────────────────────────────────────────

        [CommandMethod("Sapfir_AXES", CommandFlags.UsePickSet)]
        public void Sapfir_AXES()
        {
            ResolveActiveContext();
            AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateAxes);
        }

        [CommandMethod("Sapfir_Point_LOADS", CommandFlags.UsePickSet)]
        public void Sapfir_Point_LOADS()
        {
            ResolveActiveContext();
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
            PromptDoubleResult result = acDocEd.GetDouble("Введите значение нагрузки в тоннах: ");
            if (result.Status == PromptStatus.OK) AcUtilites.ActionOnPoints(AcUtilites.Selection(), CreatePointLoads, result.Value);
        }

        [CommandMethod("Sapfir_Line_LOADS", CommandFlags.UsePickSet)]
        public void Sapfir_Line_LOADS()
        {
            ResolveActiveContext();
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
            PromptDoubleResult result = acDocEd.GetDouble("Введите значение нагрузки в тоннах: ");
            if (result.Status == PromptStatus.OK) AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateLineLoads, result.Value);
        }

        [CommandMethod("Sapfir_Line_LOADS2", CommandFlags.UsePickSet)]
        public void Sapfir_Line_LOADS2()
        {
            ResolveActiveContext();
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
            ResolveActiveContext();
            selectedMaterialGuid = SelectMaterial("Плиты перекрытия");
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
            PromptDoubleResult result = acDocEd.GetDouble("Введите толщину плит в метрах: ");
            if (result.Status == PromptStatus.OK) AcUtilites.ActionOnPolylines(AcUtilites.Selection(), CreateSlabs, result.Value);
        }

        [CommandMethod("Sapfir_SLAB_Holes", CommandFlags.UsePickSet)]
        public void Sapfir_SLAB_holes()
        {
            ResolveActiveContext();
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
            PromptIntegerResult result = acDocEd.GetInteger("Введите идентификатор (ID) плиты для размещения отверстий: ");
            if (result.Status == PromptStatus.OK) AcUtilites.ActionOnPolylines(AcUtilites.Selection(), CreateSlabHoles, result.Value);
        }

        [CommandMethod("Sapfir_Found_SLABS", CommandFlags.UsePickSet)]
        public void Sapfir_Found_SLABS()
        {
            ResolveActiveContext();
            selectedMaterialGuid = SelectMaterial("Фундаментные плиты");
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
            PromptDoubleResult result = acDocEd.GetDouble("Введите толщину фундаментных плит в метрах: ");
            if (result.Status == PromptStatus.OK) AcUtilites.ActionOnPolylines(AcUtilites.Selection(), CreateFoundSlabs, result.Value);
        }

        [CommandMethod("Sapfir_WALLS", CommandFlags.UsePickSet)]
        public void Sapfir_WALLS()
        {
            ResolveActiveContext();
            selectedMaterialGuid = SelectMaterial("Стены");
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
            PromptDoubleResult result = acDocEd.GetDouble("Введите толщину стен в метрах: ");
            if (result.Status == PromptStatus.OK) AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateWalls, result.Value);
        }

        [CommandMethod("Sapfir_LINES", CommandFlags.UsePickSet)]
        public void Sapfir_LINES()
        {
            ResolveActiveContext();
            AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateLines, 0.0);
        }

        [CommandMethod("Sapfir_WALL_EDIT", CommandFlags.UsePickSet)]
        public void Sapfir_WALL_EDIT()
        {
            ResolveActiveContext();
            ObjectId[] objIds = AcUtilites.Selection();
            if (objIds == null) return;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForRead);
                    int spfId = AcUtilites.GetSapfirId(acEnt);
                    if (spfId == 0) continue;

                    try
                    {
                        paramModelSpf = storeySpf.GetModelByID(spfId);
                        polylineSpf = paramModelSpf.GetAxisLine();

                        if (acEnt is Polyline polyline)
                        {
                            int n = polyline.NumberOfVertices;
                            object[] coords = new object[n * 3];
                            for (int i = 0; i < n; i++)
                            {
                                Point2d pt = polyline.GetPoint2dAt(i);
                                coords[3 * i] = pt.X * 0.001;
                                coords[3 * i + 1] = pt.Y * 0.001;
                                coords[3 * i + 2] = 0.0;
                            }
                            polylineSpf.SetPoints(coords);
                        }
                        else if (acEnt is Line line)
                        {
                            buf1 = new object[]
                            {
                                line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0,
                                line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0
                            };
                            polylineSpf.SetPoints(buf1);
                        }
                        else continue;

                        paramModelSpf.RegenModel();
                    }
                    catch { }
                }
                acTrans.Commit();
            }
        }

        [CommandMethod("Sapfir_DOORS", CommandFlags.UsePickSet)]
        public void Sapfir_doors()
        {
            ResolveActiveContext();
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
            PromptDoubleResult result = acDocEd.GetDouble("Введите высоту проема в метрах: ");
            if (result.Status == PromptStatus.OK) AcUtilites.ActionOnLines(AcUtilites.Selection(), CreateDoors, GetWalls(out int cnt), result.Value);
        }

        [CommandMethod("Sapfir_STOREYS", CommandFlags.UsePickSet)]
        public void Sapfir_STOREYS()
        {
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;

            SapfirDoc activeDoc = null;

            var activeView = appSpf.GetActiveSapfirView();
            if (activeView != null)
                activeDoc = activeView.GetDocument();

            if (activeDoc == null)
                activeDoc = appSpf.GetActiveDoc();

            if (activeDoc == null)
            {
                acDocEd.WriteMessage("\nНет активного документа Sapfir");
                return;
            }

            ObjectId[] objIds = AcUtilites.Selection();
            List<Tuple<string, double>> data = AcUtilites.GetMTextsData(objIds);
            if (data.Count == 0) return;

            var sorted = data.OrderBy(s => s.Item2).ToList();

            AutoProject proj = activeDoc.CountProjects > 0
                ? activeDoc.GetActiveProject()
                : activeDoc.NewProject();

            while (proj.CountStorey > 0)
                proj.DelStoreyByIndex(0);

            foreach (var item in sorted)
            {
                AutoStorey newStorey = proj.NewStorey(item.Item1);
                newStorey.Parameter["M_LEVEL"] = item.Item2 * 0.001;
            }
        }

        [CommandMethod("Sapfir_COLUMNS", CommandFlags.UsePickSet)]
        public void Sapfir_COLUMNS()
        {
            ResolveActiveContext();
            selectedMaterialGuid = SelectMaterial("Колонны");
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;

            ObjectId[] objIds = AcUtilites.Selection();
            if (objIds == null) return;

            PromptKeywordOptions kwOpts = new PromptKeywordOptions(
                "\nЗадайте угол поворота сечения [Авто/по Двум точкам]: ", "Auto TwoPoints");
            kwOpts.Keywords.Default = "Auto";
            kwOpts.AllowNone = true;
            PromptResult kwResult = acDocEd.GetKeywords(kwOpts);

            bool manualAngle = false;
            double manualAngleRad = 0;

            if (kwResult.Status == PromptStatus.OK && kwResult.StringResult == "TwoPoints")
            {
                PromptPointResult pt1 = acDocEd.GetPoint("\nПервая точка направления угла: ");
                if (pt1.Status != PromptStatus.OK) return;
                PromptPointOptions pt2Opts = new PromptPointOptions("\nВторая точка направления угла: ");
                pt2Opts.BasePoint = pt1.Value;
                pt2Opts.UseBasePoint = true;
                PromptPointResult pt2 = acDocEd.GetPoint(pt2Opts);
                if (pt2.Status != PromptStatus.OK) return;
                double dx = pt2.Value.X - pt1.Value.X;
                double dy = pt2.Value.Y - pt1.Value.Y;
                manualAngleRad = Math.Atan2(dy, dx);
                manualAngle = true;
            }

            double storeyHeight = 3.0;
            try { storeyHeight = (double)storeySpf.Parameter["M_LEVEL"]; } catch { }

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForRead);
                    if (!(acEnt is Polyline polyline)) continue;

                    int n = polyline.NumberOfVertices;
                    bool isClosed = polyline.Closed;
                    Point2d firstPt = polyline.GetPoint2dAt(0);
                    Point2d lastPt = polyline.GetPoint2dAt(n - 1);
                    if (!isClosed && Math.Abs(firstPt.X - lastPt.X) < 1e-9
                                  && Math.Abs(firstPt.Y - lastPt.Y) < 1e-9)
                        isClosed = true;

                    if (n < 2) continue;

                    double cx = 0, cy = 0;
                    for (int i = 0; i < n; i++)
                    {
                        Point2d pt = polyline.GetPoint2dAt(i);
                        cx += pt.X * 0.001;
                        cy += pt.Y * 0.001;
                    }
                    cx /= n;
                    cy /= n;

                    int vertCount = isClosed ? n + 1 : n;
                    object[] verts = new object[vertCount * 3];
                    for (int i = 0; i < n; i++)
                    {
                        Point2d pt = polyline.GetPoint2dAt(i);
                        verts[3 * i] = (pt.X * 0.001) - cx;
                        verts[3 * i + 1] = (pt.Y * 0.001) - cy;
                        verts[3 * i + 2] = 0.0;
                    }
                    if (isClosed)
                    {
                        verts[3 * n] = verts[0];
                        verts[3 * n + 1] = verts[1];
                        verts[3 * n + 2] = 0.0;
                    }

                    double angle;
                    if (manualAngle)
                    {
                        angle = manualAngleRad;
                    }
                    else
                    {
                        double maxSideLen = 0;
                        angle = 0;
                        for (int i = 0; i < n - 1; i++)
                        {
                            double dx = (double)verts[3 * (i + 1)] - (double)verts[3 * i];
                            double dy = (double)verts[3 * (i + 1) + 1] - (double)verts[3 * i + 1];
                            double len = Math.Sqrt(dx * dx + dy * dy);
                            if (len > maxSideLen)
                            {
                                maxSideLen = len;
                                angle = Math.Atan2(dy, dx);
                            }
                        }
                    }

                    AutoModel column = storeySpf.NewModel((int)ModelsTypes.TM_COLUMN);
                    column.Parameter["M_POS_X"] = cx;
                    column.Parameter["M_POS_Y"] = cy;

                    var multiCont = column.GetMultiContour();
                    if (multiCont != null)
                    {
                        var cont = multiCont.GetContour();
                        if (cont != null)
                        {
                            var pl = cont.NewPolyLine();
                            pl.SetPoints(verts);
                            pl.Closed = 0;
                        }
                        multiCont.Parameter["M_SIZE_GRO"] = 0;
                        multiCont.Parameter["M_SIZE_GRC"] = 0;
                        multiCont.AddToLibPrj();
                    }

                    column.Parameter["M_HEIGHT"] = storeyHeight;
                    if (selectedMaterialGuid != null)
                        column.Parameter["M_MATERIAL"] = selectedMaterialGuid;
                    column.Parameter["M_ANGLE"] = angle * 180.0 / Math.PI;
                    column.RegenModel();

                    Entity writable = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite);
                    AcUtilites.SetSapfirId(writable, column.ID);
                    AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                        "\n  Колонна ID={0}, угол={1:F1}°", column.ID, angle * 180.0 / Math.PI);
                }
                acTrans.Commit();
            }
        }

        [CommandMethod("Sapfir_WINDOWS", CommandFlags.UsePickSet)]
        public void Sapfir_WINDOWS()
        {
            ResolveActiveContext();
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
            PromptDoubleResult result = acDocEd.GetDouble("Введите высоту окна в метрах: ");

            if (result.Status != PromptStatus.OK) return;

            ObjectId[] objIds = AcUtilites.Selection();
            if (objIds == null) return;

            List<int> wallIds = GetWalls(out int cnt);
            if (cnt == 0) return;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForRead);
                    if (!(acEnt is Line line)) continue;

                    foreach (int item in wallIds)
                    {
                        paramModelSpf = storeySpf.GetModelByID(item);
                        polylineSpf = paramModelSpf.GetAxisLine();
                        AutoLine spfLine = polylineSpf.GetLine(0);
                        buf1 = new object[6];
                        spfLine.GetPoints(ref buf1);
                        object[] crd = (object[])buf1;

                        Line2d wallLine = new Line2d(
                            new Point2d((double)crd[0] * 1000, (double)crd[1] * 1000),
                            new Point2d((double)crd[3] * 1000, (double)crd[4] * 1000));
                        Line2d winLine = new Line2d(
                            new Point2d(line.StartPoint.X, line.StartPoint.Y),
                            new Point2d(line.EndPoint.X, line.EndPoint.Y));
                        Tolerance tol = new Tolerance(1E-8, 1E-8);

                        if (wallLine.IsColinearTo(winLine, tol))
                        {
                            AutoModel modelWindow = paramModelSpf.NewHole((int)ModelsTypes.TM_WINDOW);
                            modelWindow.Parameter["B"] = line.Length * 0.001;
                            modelWindow.Parameter["H"] = result.Value;
                            Point2d wallStart = new Point2d((double)crd[0] * 1000, (double)crd[1] * 1000);
                            Point2d winStart = new Point2d(line.StartPoint.X, line.StartPoint.Y);
                            Point2d winEnd = new Point2d(line.EndPoint.X, line.EndPoint.Y);
                            modelWindow.Parameter["M_POS"] = Math.Min(
                                wallStart.GetDistanceTo(winStart),
                                wallStart.GetDistanceTo(winEnd)) * 0.001;
                        }
                    }
                }
                acTrans.Commit();
            }
        }

        [CommandMethod("Sapfir_STOREYS_WIZ", CommandFlags.Session)]
        public void Sapfir_STOREYS_WIZ()
        {
            var form = new StoreyGeneratorForm(appSpf, docSpf);
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(form);
            if (form.Result == null || form.Result.Count == 0) return;

            ResolveActiveContext();

            var proj = docSpf.CountProjects > 0
                ? docSpf.GetActiveProject()
                : docSpf.NewProject();

            while (proj.CountStorey > 0)
                proj.DelStoreyByIndex(0);

            foreach (var item in form.Result)
            {
                AutoStorey newStorey = proj.NewStorey(item.Item1);
                newStorey.Parameter["M_LEVEL"] = item.Item2;
            }
        }

        [CommandMethod("Sapfir_BEAMS", CommandFlags.UsePickSet)]
        public void Sapfir_BEAMS()
        {
            ResolveActiveContext();
            selectedMaterialGuid = SelectMaterial("Балки");
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;

            PromptDoubleResult hResult = acDocEd.GetDouble("\nВведите высоту сечения балки в метрах: ");
            if (hResult.Status != PromptStatus.OK) return;
            PromptDoubleResult bResult = acDocEd.GetDouble("\nВведите ширину сечения балки в метрах: ");
            if (bResult.Status != PromptStatus.OK) return;

            double h = hResult.Value;
            double b = bResult.Value;

            ObjectId[] objIds = AcUtilites.Selection();
            if (objIds == null) return;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForRead);
                    if (!(acEnt is Line line)) continue;

                    double cx = (line.StartPoint.X + line.EndPoint.X) * 0.0005;
                    double cy = (line.StartPoint.Y + line.EndPoint.Y) * 0.0005;

                    AutoModel beam = storeySpf.NewModel((int)ModelsTypes.TM_BEAM);
                    beam.Parameter["M_POS_X"] = cx;
                    beam.Parameter["M_POS_Y"] = cy;

                    polylineSpf = beam.GetAxisLine();
                    buf1 = new object[]
                    {
                        line.StartPoint.X * 0.001, line.StartPoint.Y * 0.001, 0,
                        line.EndPoint.X * 0.001, line.EndPoint.Y * 0.001, 0
                    };
                    polylineSpf.SetPoints(buf1);

                    var multiCont = beam.GetMultiContour();
                    if (multiCont != null)
                    {
                        var cont = multiCont.GetContour();
                        if (cont != null)
                        {
                            double hw = b * 0.5;
                            double hh = h * 0.5;
                            object[] secVerts = new object[]
                            {
                                -hw, -hh, 0.0,
                                 hw, -hh, 0.0,
                                 hw,  hh, 0.0,
                                -hw,  hh, 0.0,
                                -hw, -hh, 0.0
                            };
                            var pl = cont.NewPolyLine();
                            pl.SetPoints(secVerts);
                            pl.Closed = 0;
                        }
                        multiCont.Parameter["M_SIZE_GRO"] = 0;
                        multiCont.Parameter["M_SIZE_GRC"] = 0;
                        multiCont.AddToLibPrj();
                    }

                    if (selectedMaterialGuid != null)
                        beam.Parameter["M_MATERIAL"] = selectedMaterialGuid;

                    beam.RegenModel();

                    Entity writable = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite);
                    AcUtilites.SetSapfirId(writable, beam.ID);
                    acDocEd.WriteMessage("\n  Балка ID={0}, h={1:F2}м, b={2:F2}м", beam.ID, h, b);
                }
                acTrans.Commit();
            }
        }

        [CommandMethod("Sapfir_test", CommandFlags.UsePickSet)]
        public void Sapfir_test()
        {
        }
    }

    /// <summary>
    /// Типы 3D моделей Sapfir, используемые при создании 3D-представлений элементов.
    /// </summary>
    internal enum Models3dTypes
    {
        TM3_LINE = 1,
        TM3_ARC = 2,
        TM3_BEZIER = 3,
        TM3_POLYLINE = 4,
        TM3_MESH3D = 5,
        TM3_MODEL3D = 6,
        TM3_TEXT3D = 7,
        TM3_FACE = 8,
        TM3_DIM3D = 9,

        TM3_TRIANGLE = 10,
        TM3_DIMENSION = 11,
        TM3_POLY = 12,
        TM3_AXIS = 13,
        TM3_FE_BAR = 14,
        TM3_FE_TRIANGLE = 15,
        TM3_FE_QUAD = 16,
        TM3_FE_TETRAEDR = 17,
        TM3_FE_PRYSM_3 = 18,
        TM3_FE_PRYSM_4 = 19,
        TM3_FE_SUPER = 20,
        TM3_ELLIPSE = 21,
        TM3_POINT = 22,
        TM_DRAFT = 100,

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

    /// <summary>
    /// Типы параметрических объектов Sapfir, используемые при создании
    /// стен, плит, колонн, нагрузок, проёмов и других элементов.
    /// </summary>
    internal enum ModelsTypes
    {
        TM_NONE = 0,
        TM_SITE = 1,
        TM_PROJECT = 3,
        TM_STOREY = 4,
        TM_BLOCK = 5,
        TM_FEAPROJECT = 6,
        TM_NODE = 7,

        TM_WALL = 10,
        TM_WALL_1 = 65546,
        TM_SLAB = 11,
        TM_SLAB_1 = 65547,
        TM_COLUMN = 12,
        TM_BEAM = 13,
        TM_BAR = 14,
        TM_ROOF = 15,
        TM_STAIRS = 16,
        TM_AXES = 17,
        TM_LINE = 18,
        TM_ZONE = 19,
        TM_DIMENSION = 20,
        TM_LOAD = 21,
        TM_MOMENT = 22,
        TM_CAPITAL = 23,
        TM_REF = 24,
        TM_UNDERCAP = 25,
        TM_TRUSS = 26,

        TM_LOAD_1 = 65557,
        TM_LOAD_2 = 131093,
        TM_LOAD_4 = 262165,

        TM_RECESS = 29,
        TM_HOLE = 30,
        TM_WINDOW = 31,
        TM_DOOR = 32,
        TM_MODVISION = 33,
        TM_MODVISIONSECTION = 34,
        TM_MODVISIONPLAN = 35,
        TM_MODVISION3D = 36,
        TM_MODVISIONCONSTRUCT = 37,

        TM_SPHERE = 60,
        TM_PIVOTAL = 61,
        TM_PRISM = 62,
        TM_PROFILE = 63,
        TM_CONE = 64,
        TM_HIPPAR = 65,
        TM_LIGHT = 66,
        TM_POLY = 67,
        TM_TEXT = 68,
        TM_CYLINDER = 69,
        TM_SURFACE = 70,

        TM_WIND = 71,
        TM_ARM_AREA = 72,
        TM_ARM_SLAB = 73,
        TM_KARKAS = 74,
        TM_PUNCH = 75,
        TM_ARM_WALL = 76,
        TM_ARM_BAR = 77,
        TM_ARM_ZONE = 78,
        TM_ARM_COLUMN = 79,
        TM_FEASCHEMA = 80,
        TM_FEABAR = 81,
        TM_FEASHELL = 82,
        TM_MULTICONT = 83,

        TM_OTHER = 128
    }

    /// <summary>
    /// Опции управления именованными параметрами в Sapfir.
    /// </summary>
    internal enum ParameterOptionsTypes
    {
        NPA_GROUP_OPEN = 65536,
        NPA_GROUP_CLOSE = 131072,
        NPA_READONLY = 262144,

        OBP_NAMES = 4,
        OBP_COMMENTS = 8,
        OBP_VALUES = 16
    }
}