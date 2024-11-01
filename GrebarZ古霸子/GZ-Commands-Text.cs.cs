using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace Grebarz古霸子
{

    public class GZ_Commands_Text
    {
        public static Database db = HostApplicationServices.WorkingDatabase;
        public static Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        public static Editor ed = doc.Editor;

        private static string OLDTXT_OLD = null;
        private static string NEWTXT_OLD = null;

        [CommandMethod("gzEET")]
        public void GzEET()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 刪除所選的文字和屬性定義。" + Environment.NewLine +
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
                    // 提示用戶圈選文字和屬性定義
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選文字:"
                    };

                    // 過濾條件，只選擇 "*TEXT" 和 "ATTDEF" 類型的物件
                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "TEXT"),
                        new TypedValue((int)DxfCode.Start, "MTEXT"),
                        new TypedValue((int)DxfCode.Start, "ATTDEF")
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);

                    // 提示用戶進行選擇
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 刪除所有選中的文字和屬性定義
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entityToErase = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                            entityToErase.Erase(); // 刪除文字或屬性定義
                        }

                        // 提示用戶繼續圈選
                        ed.WriteMessage("\n圈選文字:");
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "選定的文字和屬性定義已刪除。".Gz_StatusBarMsg();
        }

        [CommandMethod("gz103")]
        public void Gz103()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 替換所選文本或標註中的舊字串為新字串。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 取得舊的字串
            string OLDTXT;
            if (OLDTXT_OLD == null)
            {
                PromptResult oldTxtResult = ed.GetString("\n指定舊的字串:");
                if (oldTxtResult.Status != PromptStatus.OK) return;
                OLDTXT = oldTxtResult.StringResult;
            }
            else
            {
                PromptResult oldTxtResult = ed.GetString($"\n指定舊的字串<{OLDTXT_OLD}>:");
                OLDTXT = oldTxtResult.StringResult == "" ? OLDTXT_OLD : oldTxtResult.StringResult;
            }
            OLDTXT_OLD = OLDTXT;

            // 取得新的字串
            string NEWTXT;
            if (NEWTXT_OLD == null)
            {
                PromptResult newTxtResult = ed.GetString("\n指定新的字串:");
                if (newTxtResult.Status != PromptStatus.OK) return;
                NEWTXT = newTxtResult.StringResult;
            }
            else
            {
                PromptResult newTxtResult = ed.GetString($"\n指定新的字串<{NEWTXT_OLD}>:");
                NEWTXT = newTxtResult.StringResult == "" ? NEWTXT_OLD : newTxtResult.StringResult;
                if (string.Equals(NEWTXT, "NIL", StringComparison.OrdinalIgnoreCase))
                {
                    NEWTXT = "";
                }
            }
            NEWTXT_OLD = NEWTXT;

            // 提示用戶圈選文字或標註
            ed.WriteMessage("\n圈選文字、標註:");

            // 啟動事務處理
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                bool continueSelection = true;

                while (continueSelection)
                {
                    // 選擇文字、標註和多引線
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選文字、標註:"
                    };

                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "MTEXT"),
                        new TypedValue((int)DxfCode.Start, "TEXT"),
                        new TypedValue((int)DxfCode.Start, "DIMENSION"),
                        new TypedValue((int)DxfCode.Start, "MULTILEADER")
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);

                    // 提示用戶進行選擇
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 遍歷所選物件並進行文字替換
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entity = (Entity)tr.GetObject(objId, OpenMode.ForWrite);

                            string entityText = GetEntityText(entity);
                            if (string.IsNullOrEmpty(entityText)) continue;

                            // 替換文本
                            entityText = ReplaceText(OLDTXT, NEWTXT, entityText);

                            // 設置新的文本
                            SetEntityText(entity, entityText);
                        }

                        // 提示用戶繼續圈選
                        ed.WriteMessage("\n圈選文字、標註:");
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "文字替換完成。".Gz_StatusBarMsg();
        }

        [CommandMethod("gz103Q")]
        public void Gz103Q()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 將選定的文字或標註中的指定字串替換為空字串。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 取得舊的字串
            string OLDTXT;
            if (OLDTXT_OLD == null)
            {
                PromptResult oldTxtResult = ed.GetString("\n指定字串:");
                if (oldTxtResult.Status != PromptStatus.OK) return;
                OLDTXT = oldTxtResult.StringResult;
            }
            else
            {
                PromptResult oldTxtResult = ed.GetString($"\n指定字串<{OLDTXT_OLD}>:");
                OLDTXT = oldTxtResult.StringResult == "" ? OLDTXT_OLD : oldTxtResult.StringResult;
            }
            OLDTXT_OLD = OLDTXT;

            // 新的字串為空
            string NEWTXT = "";

            // 提示用戶圈選文字或標註
            ed.WriteMessage("\n圈選文字、標註:");

            // 啟動事務處理
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                bool continueSelection = true;

                while (continueSelection)
                {
                    // 選擇文字、標註和多引線
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選文字、標註:"
                    };

                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "MTEXT"),
                        new TypedValue((int)DxfCode.Start, "TEXT"),
                        new TypedValue((int)DxfCode.Start, "DIMENSION"),
                        new TypedValue((int)DxfCode.Start, "MULTILEADER")
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);

                    // 提示用戶進行選擇
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 遍歷所選物件並進行字串替換
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entity = (Entity)tr.GetObject(objId, OpenMode.ForWrite);

                            string entityText = GetEntityText(entity);
                            if (string.IsNullOrEmpty(entityText)) continue;

                            // 替換文本
                            entityText = ReplaceText(OLDTXT, NEWTXT, entityText);

                            // 設置新的文本
                            SetEntityText(entity, entityText);
                        }

                        // 提示用戶繼續圈選
                        ed.WriteMessage("\n圈選文字、標註:");
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "文字替換完成。".Gz_StatusBarMsg();
        }

        [CommandMethod("gz1034")]
        public void Gz1034()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 圈選文字或標註，移除指定字首。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 取得舊的字串
            string OLDTXT;
            if (OLDTXT_OLD == null)
            {
                PromptResult oldTxtResult = ed.GetString("\n指定字首:");
                if (oldTxtResult.Status != PromptStatus.OK) return;
                OLDTXT = oldTxtResult.StringResult;
            }
            else
            {
                PromptResult oldTxtResult = ed.GetString($"\n指定字首<{OLDTXT_OLD}>:");
                OLDTXT = oldTxtResult.StringResult == "" ? OLDTXT_OLD : oldTxtResult.StringResult;
            }
            OLDTXT_OLD = OLDTXT;

            // 新的字串為空
            string NEWTXT = "";

            // 提示用戶圈選文字或標註
            ed.WriteMessage("\n圈選文字、標註:");

            // 啟動事務處理
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                bool continueSelection = true;

                while (continueSelection)
                {
                    // 選擇文字、標註和多引線
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選文字、標註:"
                    };

                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "MTEXT"),
                        new TypedValue((int)DxfCode.Start, "TEXT"),
                        new TypedValue((int)DxfCode.Start, "DIMENSION"),
                        new TypedValue((int)DxfCode.Start, "MULTILEADER")
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);

                    // 提示用戶進行選擇
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 遍歷所選物件並進行字串替換
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entity = (Entity)tr.GetObject(objId, OpenMode.ForWrite);

                            string entityText = GetEntityText(entity);
                            if (string.IsNullOrEmpty(entityText)) continue;

                            // 替換文本
                            entityText = ReplaceText(OLDTXT, NEWTXT, entityText);

                            // 設置新的文本
                            SetEntityText(entity, entityText);
                        }

                        // 提示用戶繼續圈選
                        ed.WriteMessage("\n圈選文字、標註:");
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            ("字首為: " + OLDTXT + " 的字首已移除").Gz_StatusBarMsg();
        }



        // 獲取實體文本
        private string GetEntityText(Entity entity)
        {
            if (entity is MText mtext)
            {
                return mtext.Contents;
            }
            else if (entity is DBText dbText)
            {
                return dbText.TextString;
            }
            else if (entity is Dimension dim)
            {
                return dim.DimensionText;
            }
            else if (entity is MLeader mleader)
            {
                return mleader.MText.Contents;
            }
            return null;
        }

        // 設置實體文本
        private void SetEntityText(Entity entity, string newText)
        {
            if (entity is MText mtext)
            {
                mtext.Contents = newText;
            }
            else if (entity is DBText dbText)
            {
                dbText.TextString = newText;
            }
            else if (entity is Dimension dim)
            {
                dim.DimensionText = newText;
            }
            else if (entity is MLeader mleader)
            {
                MText leaderText = mleader.MText;
                leaderText.Contents = newText;
                mleader.MText = leaderText;
            }
        }

        // 替換文本
        private string ReplaceText(string oldText, string newText, string sourceText)
        {
            return sourceText.Replace(oldText, newText);
        }

    }
}
