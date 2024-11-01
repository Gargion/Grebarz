using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace Grebarz古霸子
{

    public class GZ_Commands_Erase
    {
        public static Database db = HostApplicationServices.WorkingDatabase;
        public static Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        public static Editor ed = doc.Editor;

        [CommandMethod("gzEE")]
        public void GzEE()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + System.Environment.NewLine +
                         "\n指令說明: 隨選隨刪。" + System.Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 啟動事務處理
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                bool continueSelection = true;

                while (continueSelection)
                {
                    // 提示用戶圈選物件
                    PromptSelectionResult selResult = ed.GetSelection();

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 遍歷所選的物件，並進行刪除
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                            ent.Erase();
                        }
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 結束後返回消息
            ed.WriteMessage("\n已刪除所有選定物件。\n");
        }

        [CommandMethod("gzEEH")]
        public void GzEEH()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 刪除選中的填充線物件。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                bool continueSelection = true;

                while (continueSelection)
                {
                    // 提示用戶圈選剖面線（HATCH）
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選填充線:"
                    };

                    // 設置過濾條件，只選擇 HATCH 類型的物件
                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "HATCH")
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);

                    // 提示用戶進行選擇
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 刪除所有選中的剖面線（HATCH）
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entityToErase = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                            entityToErase.Erase(); // 刪除剖面線
                        }

                        // 提示用戶繼續圈選剖面線
                        ed.WriteMessage("\n圈選填充線:");
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "選定的填充線已刪除。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzEEP")]
        public void GzEEP()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 刪除所選的所有點物件。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                bool continueSelection = true;

                while (continueSelection)
                {
                    // 提示用戶圈選點
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選點:"
                    };

                    // 過濾條件，只選擇 "POINT" 類型的物件
                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "POINT")
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);

                    // 提示用戶進行選擇
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 刪除所有選中的點
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entityToErase = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                            entityToErase.Erase(); // 刪除點
                        }

                        // 提示用戶繼續圈選
                        ed.WriteMessage("\n圈選點:");
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "選定的點物件已刪除。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzEEE")]
        public void GzEEE()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 刪除指定類型的所有物件。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 提示用戶選擇物件來指定類型
                PromptEntityOptions entityOpts = new PromptEntityOptions("\n選取物件，指定類型:");
                PromptEntityResult entityResult = ed.GetEntity(entityOpts);

                if (entityResult.Status == PromptStatus.OK)
                {
                    ObjectId selectedId = entityResult.ObjectId;
                    Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);

                    // 獲取物件類型（如 LINE, CIRCLE 等）
                    string entityType = ent.GetType().Name;

                    // 提示用戶圈選相同類型的物件
                    ed.WriteMessage("\n圈選物件:");
                    bool continueSelection = true;

                    while (continueSelection)
                    {
                        // 設置過濾條件，選擇與指定類型相同的物件
                        TypedValue[] filter = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.Start, entityType)
                        };
                        SelectionFilter selectionFilter = new SelectionFilter(filter);
                        PromptSelectionResult selResult = ed.GetSelection(selectionFilter);

                        if (selResult.Status == PromptStatus.OK)
                        {
                            SelectionSet ss = selResult.Value;

                            // 刪除選中的物件
                            foreach (ObjectId objId in ss.GetObjectIds())
                            {
                                Entity entityToErase = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                                entityToErase.Erase(); // 刪除物件
                            }

                            // 提示用戶繼續圈選
                            ed.WriteMessage("\n圈選物件:");
                        }
                        else
                        {
                            continueSelection = false; // 用戶取消選擇時退出循環
                        }
                    }

                    tr.Commit(); // 提交事務
                }
            }

            // 程序結束後輸出提示
            "指定類型的物件已刪除。".Gz_StatusBarMsg();
        }





    }
}
