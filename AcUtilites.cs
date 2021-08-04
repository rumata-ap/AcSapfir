using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AcSapfir
{
   class AcUtilites
   {
      public static ObjectId[] Selection()
      {
         Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
         // Get the PickFirst selection set
         PromptSelectionResult ssPrompt;
         PromptSelectionOptions ssOption = new PromptSelectionOptions();
         ssOption.MessageForAdding = "Выберите примитивы для создания объектов САПФИР:";
         ssPrompt = acDocEd.SelectImplied();
         SelectionSet sset;

         if (ssPrompt.Status == PromptStatus.OK) sset = ssPrompt.Value;
         else { ssPrompt = acDocEd.GetSelection(ssOption); sset = ssPrompt.Value; }

         ObjectId[] acObjIds = null;
         if (sset != null) acObjIds = sset.GetObjectIds();

         return acObjIds;
      }

      public static void ActionOnPoints(ObjectId[] objIds, Action<DBPoint, double> action, double parameter)
      {
         Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;
         Editor acDocEd = AcApp.DocumentManager.MdiActiveDocument.Editor;
         if (objIds != null)
         {
            // Starts a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
               try
               {
                  // Gets information about each object
                  foreach (ObjectId acObjId in objIds)
                  {
                     Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                     //acDocEd.WriteMessage("\n" + acEnt.GetType().Name);
                     switch (acEnt.GetType().Name)
                     {
                        case "DBPoint":
                           DBPoint line = acEnt as DBPoint;
                           action(line, parameter);
                           break;
                        default:
                           break;
                     }

                  }
                  acTrans.Commit();
               }
               finally
               {
                  acTrans.Dispose();
               }
            }
         }
      }

      public static void ActionOnLines(ObjectId[] objIds, Action<Line, double> action, double parameter)
      {
         Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;
         
         if (objIds != null)
         {
            // Starts a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
               try
               {
                  // Gets information about each object
                  foreach (ObjectId acObjId in objIds)
                  {
                     Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                     //acDocEd.WriteMessage("\n" + acEnt.GetType().Name);
                     switch (acEnt.GetType().Name)
                     {
                        case "Line":
                           Line line = acEnt as Line;
                           action(line, parameter);
                           break;
                        case "Polyline":
                           Polyline polyline = acEnt as Polyline;
                           break;
                        case "Polyline3d":
                           Point3dCollection ptsAc = new Point3dCollection();
                           Polyline3d polyline3d = acEnt as Polyline3d;
                           polyline3d.GetStretchPoints(ptsAc);
                           break;
                        default:
                           break;
                     }

                  }
                  acTrans.Commit();
               }
               finally
               {
                  acTrans.Dispose();
               }
            }
         }
      }

      public static void ActionOnLines(ObjectId[] objIds, Action<Line, int> action)
      {
         Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;
         List<ObjectId> ids = new List<ObjectId>(objIds);
         if (objIds != null)
         {
            // Starts a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
               try
               {
                  // Gets information about each object
                  foreach (ObjectId acObjId in objIds)
                  {
                     Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                     //acDocEd.WriteMessage("\n" + acEnt.GetType().Name);
                     switch (acEnt.GetType().Name)
                     {
                        case "Line":
                           Line line = acEnt as Line;
                           action(line, ids.IndexOf(acObjId));
                           break;
                        default:
                           break;
                     }

                  }
                  acTrans.Commit();
               }
               finally
               {
                  acTrans.Dispose();
               }
            }
         }
      }

      public static void ActionOnLines(ObjectId[] objIds, Action<Line, double, double> action, double parameter, double parameter1)
      {
         Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

         if (objIds != null)
         {
            // Starts a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
               try
               {
                  // Gets information about each object
                  foreach (ObjectId acObjId in objIds)
                  {
                     Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                     //acDocEd.WriteMessage("\n" + acEnt.GetType().Name);
                     switch (acEnt.GetType().Name)
                     {
                        case "Line":
                           Line line = acEnt as Line;
                           action(line, parameter, parameter1);
                           break;
                        case "Polyline":
                           Polyline polyline = acEnt as Polyline;
                           break;
                        case "Polyline3d":
                           Point3dCollection ptsAc = new Point3dCollection();
                           Polyline3d polyline3d = acEnt as Polyline3d;
                           polyline3d.GetStretchPoints(ptsAc);
                           break;
                        default:
                           break;
                     }

                  }
                  acTrans.Commit();
               }
               finally
               {
                  acTrans.Dispose();
               }
            }
         }
      }

      public static void ActionOnLines(ObjectId[] objIds, Action<Line, IEnumerable<int>, double> action, IEnumerable<int> par, double parameter)
      {
         Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

         if (objIds != null)
         {
            // Starts a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
               try
               {
                  // Gets information about each object
                  foreach (ObjectId acObjId in objIds)
                  {
                     Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                     //acDocEd.WriteMessage("\n" + acEnt.GetType().Name);
                     switch (acEnt.GetType().Name)
                     {
                        case "Line":
                           Line line = acEnt as Line;
                           action(line, par, parameter);
                           break;
                        case "Polyline":
                           Polyline polyline = acEnt as Polyline;
                           break;
                        case "Polyline3d":
                           Point3dCollection ptsAc = new Point3dCollection();
                           Polyline3d polyline3d = acEnt as Polyline3d;
                           polyline3d.GetStretchPoints(ptsAc);
                           break;
                        default:
                           break;
                     }

                  }
                  acTrans.Commit();
               }
               finally
               {
                  acTrans.Dispose();
               }
            }
         }
      }

      public static void ActionOnPolylines(ObjectId[] objIds, Action<Polyline, double> action, double parameter)
      {
         Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

         if (objIds != null)
         {
            // Starts a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
               try
               {
                  // Gets information about each object
                  foreach (ObjectId acObjId in objIds)
                  {
                     Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                     //acDocEd.WriteMessage("\n" + acEnt.GetType().Name);
                     switch (acEnt.GetType().Name)
                     {
                        case "Line":
                           Line line = acEnt as Line;                          
                           break;
                        case "Polyline":
                           Polyline polyline = acEnt as Polyline;
                           action(polyline, parameter);
                           break;
                        case "Polyline3d":
                           Point3dCollection ptsAc = new Point3dCollection();
                           Polyline3d polyline3d = acEnt as Polyline3d;
                           polyline3d.GetStretchPoints(ptsAc);
                           break;
                        default:
                           break;
                     }

                  }
                  acTrans.Commit();
               }
               finally
               {
                  acTrans.Dispose();
               }
            }
         }
      }

      public static void ActionOnPolylines(ObjectId[] objIds, Action<Polyline, int> action, int id)
      {
         Database acCurDb = AcApp.DocumentManager.MdiActiveDocument.Database;

         if (objIds != null)
         {
            // Starts a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
               try
               {
                  // Gets information about each object
                  foreach (ObjectId acObjId in objIds)
                  {
                     Entity acEnt = (Entity)acTrans.GetObject(acObjId, OpenMode.ForWrite, true);
                     //acDocEd.WriteMessage("\n" + acEnt.GetType().Name);
                     switch (acEnt.GetType().Name)
                     {
                        case "Line":
                           Line line = acEnt as Line;
                           break;
                        case "Polyline":
                           Polyline polyline = acEnt as Polyline;
                           action(polyline, id);
                           break;
                        case "Polyline3d":
                           Point3dCollection ptsAc = new Point3dCollection();
                           Polyline3d polyline3d = acEnt as Polyline3d;
                           polyline3d.GetStretchPoints(ptsAc);
                           break;
                        default:
                           break;
                     }

                  }
                  acTrans.Commit();
               }
               finally
               {
                  acTrans.Dispose();
               }
            }
         }
      }
   }
}
