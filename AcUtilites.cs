using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AcSapfir
{
    /// <summary>
    /// Вспомогательный класс для работы с объектами AutoCAD.
    /// Предоставляет методы выбора примитивов и итерации по ним
    /// с выполнением заданного действия для каждого примитива нужного типа.
    /// </summary>
    internal static class AcUtilites
    {
        /// <summary>
        /// Выполняет выбор объектов в AutoCAD.
        /// Сначала пытается использовать текущий набор выбора (PickFirst),
        /// если он пуст — запрашивает выбор у пользователя.
        /// </summary>
        /// <returns>
        /// Массив идентификаторов выбранных объектов (ObjectId).
        /// Может возвращать null, если выбор отменён.
        /// </returns>
        public static ObjectId[] Selection()
        {
            Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
            PromptSelectionOptions ssOption = new PromptSelectionOptions
            {
                MessageForAdding = "Выберите примитивы для создания объектов САПФИР:"
            };

            PromptSelectionResult ssPrompt = acDocEd.SelectImplied();
            if (ssPrompt.Status != PromptStatus.OK)
            {
                ssPrompt = acDocEd.GetSelection(ssOption);
            }

            if (ssPrompt.Status == PromptStatus.OK && ssPrompt.Value != null)
            {
                return ssPrompt.Value.GetObjectIds();
            }

            return null;
        }

        /// <summary>
        /// Выполняет заданное действие для всех выбранных точек (DBPoint).
        /// </summary>
        /// <param name="objIds">Массив идентификаторов объектов для обработки.</param>
        /// <param name="action">Делегат действия, принимающий точку и числовой параметр.</param>
        /// <param name="parameter">Дополнительный числовой параметр для передачи в действие.</param>
        public static void ActionOnPoints(ObjectId[] objIds, Action<DBPoint, double> action, double parameter)
        {
            if (objIds == null) return;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                    if (acEnt is DBPoint dbPoint)
                    {
                        action(dbPoint, parameter);
                    }
                }
                acTrans.Commit();
            }
        }

        /// <summary>
        /// Выполняет заданное действие для всех выбранных линий (Line),
        /// передавая числовой параметр.
        /// </summary>
        /// <param name="objIds">Массив идентификаторов объектов для обработки.</param>
        /// <param name="action">Делегат действия, принимающий линию и числовой параметр.</param>
        /// <param name="parameter">Дополнительный числовой параметр для передачи в действие.</param>
        public static void ActionOnLines(ObjectId[] objIds, Action<Line, double> action, double parameter)
        {
            if (objIds == null) return;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                    if (acEnt is Line line)
                    {
                        action(line, parameter);
                    }
                }
                acTrans.Commit();
            }
        }

        /// <summary>
        /// Выполняет заданное действие для всех выбранных линий, передавая индекс объекта.
        /// </summary>
        /// <param name="objIds">Массив идентификаторов объектов для обработки.</param>
        /// <param name="action">Делегат действия, принимающий линию и её порядковый номер в списке.</param>
        public static void ActionOnLines(ObjectId[] objIds, Action<Line, int> action)
        {
            if (objIds == null) return;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;
            List<ObjectId> ids = new List<ObjectId>(objIds);

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                    if (acEnt is Line line)
                    {
                        action(line, ids.IndexOf(acObjId));
                    }
                }
                acTrans.Commit();
            }
        }

        /// <summary>
        /// Выполняет заданное действие для всех выбранных линий, передавая два числовых параметра.
        /// Используется для нагрузок с разными значениями в начале и конце линии.
        /// </summary>
        /// <param name="objIds">Массив идентификаторов объектов для обработки.</param>
        /// <param name="action">Делегат действия, принимающий линию и два числовых параметра.</param>
        /// <param name="parameter">Первый числовой параметр (например, нагрузка в начале).</param>
        /// <param name="parameter1">Второй числовой параметр (например, нагрузка в конце).</param>
        public static void ActionOnLines(ObjectId[] objIds, Action<Line, double, double> action, double parameter, double parameter1)
        {
            if (objIds == null) return;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                    if (acEnt is Line line)
                    {
                        action(line, parameter, parameter1);
                    }
                }
                acTrans.Commit();
            }
        }

        /// <summary>
        /// Выполняет заданное действие для всех выбранных линий, передавая
        /// коллекцию целочисленных идентификаторов и числовой параметр.
        /// Используется при создании дверей, когда нужно знать ID стен.
        /// </summary>
        /// <param name="objIds">Массив идентификаторов объектов для обработки.</param>
        /// <param name="action">Делегат действия, принимающий линию, коллекцию int-параметров и числовой параметр.</param>
        /// <param name="par">Коллекция целочисленных параметров (например, ID стен).</param>
        /// <param name="parameter">Дополнительный числовой параметр.</param>
        public static void ActionOnLines(ObjectId[] objIds, Action<Line, IEnumerable<int>, double> action, IEnumerable<int> par, double parameter)
        {
            if (objIds == null) return;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                    if (acEnt is Line line)
                    {
                        action(line, par, parameter);
                    }
                }
                acTrans.Commit();
            }
        }

        /// <summary>
        /// Выполняет заданное действие для всех выбранных полилиний (Polyline),
        /// передавая числовой параметр.
        /// </summary>
        /// <param name="objIds">Массив идентификаторов объектов для обработки.</param>
        /// <param name="action">Делегат действия, принимающий полилинию и числовой параметр.</param>
        /// <param name="parameter">Дополнительный числовой параметр для передачи в действие (например, толщина плиты).</param>
        public static void ActionOnPolylines(ObjectId[] objIds, Action<Polyline, double> action, double parameter)
        {
            if (objIds == null) return;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                    if (acEnt is Polyline polyline)
                    {
                        action(polyline, parameter);
                    }
                }
                acTrans.Commit();
            }
        }

        /// <summary>
        /// Выполняет заданное действие для всех выбранных полилиний, передавая целочисленный идентификатор.
        /// Используется при создании отверстий в плитах, где нужен ID плиты.
        /// </summary>
        /// <param name="objIds">Массив идентификаторов объектов для обработки.</param>
        /// <param name="action">Делегат действия, принимающий полилинию и целочисленный идентификатор.</param>
        /// <param name="id">Дополнительный целочисленный идентификатор (например, ID плиты).</param>
        public static void ActionOnPolylines(ObjectId[] objIds, Action<Polyline, int> action, int id)
        {
            if (objIds == null) return;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                    if (acEnt is Polyline polyline)
                    {
                        action(polyline, id);
                    }
                }
                acTrans.Commit();
            }
        }

        /// <summary>
        /// Извлекает данные из выбранных многострочных текстов (MText):
        /// текстовое содержимое и Y-координату точки вставки каждого объекта.
        /// Y-координата используется как абсолютная отметка в миллиметрах.
        /// </summary>
        /// <param name="objIds">Массив идентификаторов объектов для обработки.</param>
        /// <returns>
        /// Список кортежей (имя, Y-координата), где имя — очищенный от форматирования текст MText,
        /// а Y-координата — вертикальная координата точки вставки в пространстве модели (мм).
        /// </returns>
        public static List<Tuple<string, double>> GetMTextsData(ObjectId[] objIds)
        {
            List<Tuple<string, double>> result = new List<Tuple<string, double>>();

            if (objIds == null) return result;

            Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                foreach (ObjectId acObjId in objIds)
                {
                    Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForRead);
                    if (acEnt is MText mtext)
                    {
                        string text = StripMTextFormat(mtext.Text);
                        double y = mtext.Location.Y;
                        result.Add(new Tuple<string, double>(text, y));
                    }
                }
                acTrans.Commit();
            }

            return result;
        }

        /// <summary>
        /// Удаляет коды форматирования MText из строки, возвращая чистый текст.
        /// Убирает коды типа { \A, \P, \H, \W, \T, \S, \C, \f, \~, а также фигурные скобки.
        /// </summary>
        /// <param name="rawText">Строка MText с форматированием.</param>
        /// <returns>Очищенный текст без кодов форматирования.</returns>
        internal static string StripMTextFormat(string rawText)
        {
            if (string.IsNullOrEmpty(rawText)) return rawText;

            string result = rawText;

            // Удаление обратных кодов форматирования: \A, \P, \H, \W, \T, \S, \C, \f, \L, \l, \O, \o, \K, \k, \~ и т.д.
            result = Regex.Replace(result, @"\\[A-Za-z][^;]*;?", "");

            // Удаление \P (перенос строки) — заменяем на пробел
            result = result.Replace(@"\P", " ");

            // Удаление фигурных скобок
            result = result.Replace("{", "").Replace("}", "");

            // Удаление оставшихся обратных слешей с одиночными буквами
            result = Regex.Replace(result, @"\\[A-Za-z]{1,2}", "");

            return result.Trim();
        }
    }
}