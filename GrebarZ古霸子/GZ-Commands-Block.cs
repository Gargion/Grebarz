using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace Grebarz古霸子
{

    public class GZ_Commands_Block
    {
        public static Database db = HostApplicationServices.WorkingDatabase;
        public static Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        public static Editor ed = doc.Editor;

        [CommandMethod("gzEEB")]
        public void GzEEB()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 圈選範圍，指定圖塊，並刪除相同名稱的圖塊。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 提示用戶圈選範圍，選擇 "INSERT"（圖塊）的所有物件
                PromptSelectionOptions selOpts = new PromptSelectionOptions();
                selOpts.MessageForAdding = "\n圈選範圍: ";
                TypedValue[] filter = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Start, "INSERT") // 選擇所有圖塊
                };
                SelectionFilter selectionFilter = new SelectionFilter(filter);
                PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                if (selResult.Status != PromptStatus.OK)
                {
                    return; // 如果未選擇範圍，則退出
                }

                SelectionSet ss = selResult.Value; // 圈選範圍內的所有圖塊

                // 開始讓用戶選擇指定的圖塊
                bool continueSelection = true;
                while (continueSelection)
                {
                    PromptEntityOptions entityOpts = new PromptEntityOptions("\n指定圖塊: ");
                    entityOpts.SetRejectMessage("\n必須選擇圖塊.");
                    entityOpts.AddAllowedClass(typeof(BlockReference), true); // 只允許選擇圖塊

                    PromptEntityResult entityResult = ed.GetEntity(entityOpts);

                    if (entityResult.Status == PromptStatus.OK)
                    {
                        ObjectId selectedId = entityResult.ObjectId;
                        BlockReference blockRef = (BlockReference)tr.GetObject(selectedId, OpenMode.ForRead);

                        // 獲取圖塊名稱
                        string blockName = blockRef.Name;

                        // 如果圖塊名以 "*" 開頭，則逐個刪除
                        if (blockName.StartsWith("*"))
                        {
                            foreach (ObjectId objId in ss.GetObjectIds())
                            {
                                BlockReference blkRef = (BlockReference)tr.GetObject(objId, OpenMode.ForWrite);
                                if (blkRef.Name == blockName)
                                {
                                    blkRef.Erase(); // 刪除圖塊
                                }
                            }
                            // 程序結束後輸出提示
                            ("圖塊名稱: " + blockName + "已刪除。").Gz_StatusBarMsg();
                        }
                        else
                        {
                            // 使用過濾器再次選擇相同名稱的圖塊
                            TypedValue[] blockFilter = new TypedValue[]
                            {
                                new TypedValue((int)DxfCode.BlockName, blockName)
                            };
                            SelectionFilter blockSelectionFilter = new SelectionFilter(blockFilter);
                            PromptSelectionResult blockSelResult = ed.SelectAll(blockSelectionFilter);

                            if (blockSelResult.Status == PromptStatus.OK)
                            {
                                SelectionSet ssx = blockSelResult.Value;

                                // 刪除所有符合條件的圖塊
                                foreach (ObjectId objId in ssx.GetObjectIds())
                                {
                                    BlockReference blkRef = (BlockReference)tr.GetObject(objId, OpenMode.ForWrite);
                                    blkRef.Erase(); // 刪除圖塊
                                }
                                // 程序結束後輸出提示
                                ("圖塊名稱: " + blockName + "已刪除。").Gz_StatusBarMsg();
                            }
                        }
                    }
                    else
                    {
                        // 程序結束後輸出提示
                        ("使用者取消動作。").Gz_StatusBarMsg();
                        continueSelection = false; // 如果用戶取消選擇，結束循環
                    }
                }

                tr.Commit(); // 提交事務

            }


        }

        [CommandMethod("gzEEBB")]
        public void GzEEBB()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 圈選範圍，並刪除範圍中所有圖塊。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 使用 AutoCAD 的 UNDO 命令來啟動一個 "BE" 標記
            doc.SendStringToExecute("._UNDO _BE ", true, false, false);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                bool continueSelection = true;

                while (continueSelection)
                {
                    // 提示用戶圈選圖塊
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選圖塊: "
                    };
                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.Start, "INSERT") // 只選擇圖塊
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);
                    PromptSelectionResult selResult = ed.GetSelection(selOpts, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;

                        // 刪除所選圖塊
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entity = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                            entity.Erase(); // 刪除圖塊
                        }
                    }
                    else
                    {
                        continueSelection = false; // 當用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 使用 AutoCAD 的 UNDO 命令來結束 "BE" 標記
            doc.SendStringToExecute("._UNDO _END ", true, false, false);

            // 程序結束後輸出完成提示
            "圖塊刪除操作已完成。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzBXX")]
        public void GzBXX()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 先指定圖塊，再圈選分解範圍，將指定名稱的所有圖塊分解。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 提示用戶選擇圖塊
                PromptEntityOptions entityOpts = new PromptEntityOptions("\n選取物件，指定圖塊:");
                entityOpts.SetRejectMessage("\n必須選擇圖塊。");
                entityOpts.AddAllowedClass(typeof(BlockReference), true);
                PromptEntityResult entityResult = ed.GetEntity(entityOpts);

                if (entityResult.Status == PromptStatus.OK)
                {
                    ObjectId selectedId = entityResult.ObjectId;
                    BlockReference blockRef = (BlockReference)tr.GetObject(selectedId, OpenMode.ForRead);

                    // 獲取圖塊名稱
                    string blockName = blockRef.Name;

                    // 提示用戶圈選圖塊
                    ed.WriteMessage("\n圈選圖塊:");
                    bool continueSelection = true;

                    while (continueSelection)
                    {
                        // 設置過濾條件，只選擇與指定圖塊名稱相同的圖塊
                        TypedValue[] filter = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.BlockName, blockName),
                            new TypedValue((int)DxfCode.Start, "INSERT") // 限制為圖塊
                        };
                        SelectionFilter selectionFilter = new SelectionFilter(filter);
                        PromptSelectionResult selResult = ed.GetSelection(selectionFilter);

                        if (selResult.Status == PromptStatus.OK)
                        {
                            SelectionSet ss = selResult.Value;

                            // 將每個圖塊進行分解
                            foreach (ObjectId objId in ss.GetObjectIds())
                            {
                                BlockReference blkRef = (BlockReference)tr.GetObject(objId, OpenMode.ForWrite);

                                // 分解圖塊
                                DBObjectCollection explodedObjects = new DBObjectCollection();
                                blkRef.Explode(explodedObjects);

                                // 將分解的物件添加到模型空間
                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                                foreach (Entity ent in explodedObjects)
                                {
                                    btr.AppendEntity(ent);
                                    tr.AddNewlyCreatedDBObject(ent, true);
                                }

                                // 刪除原圖塊
                                blkRef.Erase();
                            }

                            // 提示用戶繼續圈選圖塊
                            ed.WriteMessage("\n圈選圖塊:");
                        }
                        else
                        {
                            continueSelection = false; // 當用戶取消選擇時退出循環
                        }
                    }

                    tr.Commit(); // 提交事務
                }
            }

            // 程序結束後輸出完成提示
            "指定圖塊已分解並刪除。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzBEE")]
        public void GzBEE()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 先指定圖塊，再圈選刪除範圍，範圍內所有指定圖塊刪除。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 提示用戶選擇圖塊
                PromptEntityOptions entityOpts = new PromptEntityOptions("\n選取物件，指定圖塊:");
                entityOpts.SetRejectMessage("\n必須選擇圖塊。");
                entityOpts.AddAllowedClass(typeof(BlockReference), true); // 限制為圖塊選擇
                PromptEntityResult entityResult = ed.GetEntity(entityOpts);

                if (entityResult.Status == PromptStatus.OK)
                {
                    ObjectId selectedId = entityResult.ObjectId;
                    BlockReference blockRef = (BlockReference)tr.GetObject(selectedId, OpenMode.ForRead);

                    // 獲取圖塊名稱
                    string blockName = blockRef.Name;

                    // 提示用戶圈選同名圖塊
                    ed.WriteMessage("\n圈選圖塊:");
                    bool continueSelection = true;

                    while (continueSelection)
                    {
                        // 設置過濾條件，只選擇指定圖塊名稱的圖塊
                        TypedValue[] filter = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.BlockName, blockName),
                            new TypedValue((int)DxfCode.Start, "INSERT") // 只選擇圖塊
                        };
                        SelectionFilter selectionFilter = new SelectionFilter(filter);
                        PromptSelectionResult selResult = ed.GetSelection(selectionFilter);

                        if (selResult.Status == PromptStatus.OK)
                        {
                            SelectionSet ss = selResult.Value;

                            // 刪除所有選中的圖塊
                            foreach (ObjectId objId in ss.GetObjectIds())
                            {
                                Entity entityToErase = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                                entityToErase.Erase(); // 刪除圖塊
                            }

                            // 提示用戶繼續圈選圖塊
                            ed.WriteMessage("\n圈選圖塊:");
                        }
                        else
                        {
                            continueSelection = false; // 當用戶取消選擇時，退出循環
                        }
                    }

                    tr.Commit(); // 提交事務
                }
            }

            // 程序結束後輸出完成提示
            "指定圖塊已刪除。".Gz_StatusBarMsg();
        }


    }
}
