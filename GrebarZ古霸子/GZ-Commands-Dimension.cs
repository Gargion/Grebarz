using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace Grebarz古霸子
{

    public class GZ_Commands_Dimension
    {
        public static Database db = HostApplicationServices.WorkingDatabase;
        public static Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        public static Editor ed = doc.Editor;

        [CommandMethod("gzEED")]
        public void GzEED()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 刪除所選的標註物件。" + Environment.NewLine +
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
                    // 提示用戶圈選標註物件
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選標註:"
                    };

                    // 過濾條件，只選擇標註類型的物件
                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "DIMENSION")
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);

                    // 提示用戶進行選擇
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 刪除所有選中的標註物件
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entityToErase = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                            entityToErase.Erase(); // 刪除標註
                        }

                        // 提示用戶繼續圈選標註
                        ed.WriteMessage("\n圈選標註:");
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "選定的標註物件已刪除。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzEED0")]
        public void GzEED0()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 圈選範圍，刪除[內定測量值]的標註。" + Environment.NewLine +
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
                    // 提示用戶圈選標註
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選標註:"
                    };

                    // 設置過濾條件，選擇包含或不包含文字的標註物件
                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "DIMENSION"),
                        new TypedValue((int)DxfCode.Text, ""),
                        new TypedValue((int)DxfCode.Text, "*<>*")
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);

                    // 提示用戶進行選擇
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 刪除所有選中的標註
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entityToErase = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                            entityToErase.Erase(); // 刪除標註
                        }

                        // 提示用戶繼續圈選標註
                        ed.WriteMessage("\n圈選標註:");
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "選定的標註物件已刪除。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzEED00")]
        public void GzEED00()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 刪除包含或不包含文字的標註物件。" + Environment.NewLine +
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
                    // 提示用戶圈選標註
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選標註:"
                    };

                    // 設置過濾條件，選擇標註類型的物件，過濾沒有文字和有文字的標註
                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "DIMENSION"),
                        new TypedValue((int)DxfCode.Text, ""),
                        new TypedValue((int)DxfCode.Text, "*<>*")
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);

                    // 提示用戶進行選擇
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 刪除所有選中的標註
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entityToErase = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                            entityToErase.Erase(); // 刪除標註
                        }

                        // 提示用戶繼續圈選標註
                        ed.WriteMessage("\n圈選標註:");
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "選定的標註物件已刪除。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzEED1")]
        public void GzEED1()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 圈選範圍，刪除[非內定測量值]的標註。" + Environment.NewLine +
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
                    // 提示用戶圈選標註
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選標註:"
                    };

                    // 設置過濾條件，選擇標註類型的物件
                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "DIMENSION")
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);

                    // 提示用戶進行選擇
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 刪除所有符合條件的標註
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Dimension dim = (Dimension)tr.GetObject(objId, OpenMode.ForWrite);

                            // 獲取標註的文本內容
                            string dimText = dim.DimensionText;

                            // 如果文本不為空或包含 "<>"，則刪除該標註
                            if (!string.IsNullOrEmpty(dimText) || dimText.Contains("<>"))
                            {
                                dim.Erase(); // 刪除標註
                            }
                        }

                        // 提示用戶繼續圈選標註
                        ed.WriteMessage("\n圈選標註:");
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "符合條件的標註已刪除。".Gz_StatusBarMsg();
        }



    }
}
