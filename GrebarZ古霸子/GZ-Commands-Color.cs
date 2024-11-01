using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace Grebarz古霸子
{

    public class GZ_Commands_Color
    {
        public static Database db = HostApplicationServices.WorkingDatabase;
        public static Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        public static Editor ed = doc.Editor;

        [CommandMethod("gzCNL")]
        public void GzCNL()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 更改選取範圍內指定圖層物件的顏色。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 打開 AutoCAD 顏色選擇對話框，並讓用戶選擇顏色
            Color color = Color.FromColorIndex(ColorMethod.ByAci, 256); // 默認顏色
            PromptIntegerOptions colorOptions = new PromptIntegerOptions("\n選擇顏色 (0-256): ")
            {
                DefaultValue = 256,
                AllowNegative = false,
                AllowNone = true,
                LowerLimit = 0,
                UpperLimit = 256
            };
            PromptIntegerResult colorResult = ed.GetInteger(colorOptions);

            if (colorResult.Status == PromptStatus.OK)
            {
                color = Color.FromColorIndex(ColorMethod.ByAci, (short)colorResult.Value);
            }

            // 提示用戶圈選範圍內的物件
            PromptSelectionOptions selOpts = new PromptSelectionOptions();
            selOpts.MessageForAdding = "\n圈選範圍: ";
            PromptSelectionResult selResult = ed.GetSelection(selOpts);

            if (selResult.Status != PromptStatus.OK)
            {
                return; // 如果未選擇範圍，則退出
            }

            SelectionSet ss = selResult.Value;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 開始讓用戶選擇圖層並更改顏色
                bool continueSelection = true;
                while (continueSelection)
                {
                    // 提示用戶選擇物件來指定圖層
                    PromptEntityOptions entityOpts = new PromptEntityOptions("\n指定圖層: ");
                    PromptEntityResult entityResult = ed.GetEntity(entityOpts);

                    if (entityResult.Status == PromptStatus.OK)
                    {
                        ObjectId selectedId = entityResult.ObjectId;
                        Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);

                        // 獲取選中的物件圖層名稱
                        string selectedLayer = ent.Layer;

                        // 選擇指定圖層上的所有物件
                        TypedValue[] filter = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.LayerName, selectedLayer)
                        };
                        SelectionFilter selectionFilter = new SelectionFilter(filter);
                        PromptSelectionResult selectionResult = ed.SelectAll(selectionFilter);

                        if (selectionResult.Status == PromptStatus.OK)
                        {
                            SelectionSet ssX = selectionResult.Value;

                            // 更改指定圖層的物件顏色
                            foreach (ObjectId objId in ssX.GetObjectIds())
                            {
                                Entity entityToChange = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                                entityToChange.Color = color; // 更改顏色
                            }
                        }
                    }
                    else
                    {
                        continueSelection = false; // 結束選擇
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出完成訊息
            "圖層物件的顏色已更改。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzEEC")]
        public void GzEEC()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 圈選範圍內的物件，並根據指定顏色刪除物件。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 提示用戶選擇物件範圍
            PromptSelectionOptions selOpts = new PromptSelectionOptions
            {
                MessageForAdding = "\n圈選範圍: "
            };
            PromptSelectionResult selResult = ed.GetSelection(selOpts);

            if (selResult.Status != PromptStatus.OK)
            {
                return; // 如果未選擇範圍，則退出
            }

            SelectionSet ss = selResult.Value;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                bool continueSelection = true;

                while (continueSelection)
                {
                    // 提示用戶指定顏色
                    PromptEntityOptions entityOpts = new PromptEntityOptions("\n指定顏色: ");
                    entityOpts.SetRejectMessage("\n必須選擇一個物件。");
                    entityOpts.AddAllowedClass(typeof(Entity), true); // 允許所有實體類型

                    PromptEntityResult entityResult = ed.GetEntity(entityOpts);

                    if (entityResult.Status == PromptStatus.OK)
                    {
                        ObjectId selectedId = entityResult.ObjectId;
                        Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);

                        // 獲取物件顏色
                        Color color = ent.Color;
                        short colorIndex = color.ColorIndex;

                        // 如果顏色為 NULL，默認設為 256（ByLayer）
                        if (colorIndex == 0)
                        {
                            colorIndex = 256; // 默認顏色為 ByLayer
                        }

                        // 使用顏色過濾器來選擇相同顏色的物件
                        TypedValue[] filter = null;
                        if (colorIndex == 256)
                        {
                            // 顏色為 ByLayer 的過濾器
                            filter = new TypedValue[]
                            {
                                new TypedValue((int)DxfCode.Color, colorIndex),
                                new TypedValue((int)DxfCode.LayerName, ent.Layer)
                            };
                        }
                        else
                        {
                            // 顏色為其他顏色的過濾器
                            filter = new TypedValue[]
                            {
                                new TypedValue((int)DxfCode.Color, colorIndex)
                            };
                        }

                        SelectionFilter selectionFilter = new SelectionFilter(filter);
                        PromptSelectionResult selectionResult = ed.SelectAll(selectionFilter);

                        if (selectionResult.Status == PromptStatus.OK)
                        {
                            SelectionSet ssX = selectionResult.Value;

                            // 刪除選中的物件
                            foreach (ObjectId objId in ssX.GetObjectIds())
                            {
                                Entity entityToErase = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                                entityToErase.Erase(); // 刪除物件
                            }
                        }
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇，退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出完成提示
            "選定顏色的物件已刪除。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzCEE")]
        public void GzCEE()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 先指定顏色，再圈選刪除範圍，刪除範圍內指定顏色的所有物件。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 提示用戶選擇一個物件來指定顏色
                PromptEntityOptions entityOpts = new PromptEntityOptions("\n選取物件，指定顏色:");
                PromptEntityResult entityResult = ed.GetEntity(entityOpts);

                if (entityResult.Status == PromptStatus.OK)
                {
                    ObjectId selectedId = entityResult.ObjectId;
                    Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);

                    // 獲取所選物件的顏色
                    short colorIndex = ent.Color.ColorIndex;

                    // 如果顏色為 NULL，設置為 256（表示 ByLayer）
                    if (colorIndex == 0)
                    {
                        colorIndex = 256; // 默認顏色 ByLayer
                    }

                    // 提示用戶選擇要刪除的物件
                    ed.WriteMessage("\n圈選物件:");

                    bool continueSelection = true;

                    while (continueSelection)
                    {
                        // 設置過濾條件，只選擇與指定顏色相同的物件
                        TypedValue[] filter = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.Color, colorIndex)
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
                            continueSelection = false; // 當用戶取消選擇時退出循環
                        }
                    }

                    tr.Commit(); // 提交事務
                }
            }

            // 程序結束後輸出提示
            "與指定顏色相同的物件已刪除。".Gz_StatusBarMsg();
        }

        [CommandMethod("gz306")]
        public void Gz306()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 將指定圖層的物件顏色設為隨層（ByLayer）。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 提示用戶圈選範圍
                PromptSelectionOptions selOpts = new PromptSelectionOptions
                {
                    MessageForAdding = "\n圈選範圍:"
                };
                PromptSelectionResult selResult = ed.GetSelection(selOpts);

                if (selResult.Status != PromptStatus.OK)
                {
                    return; // 如果未圈選範圍，退出
                }

                SelectionSet ss = selResult.Value;

                // 提示用戶選擇一個物件來指定圖層
                bool continueSelection = true;

                while (continueSelection)
                {
                    // 提示用戶選擇物件來指定圖層
                    PromptEntityOptions entityOpts = new PromptEntityOptions("\n指向物件，同圖層顏色隨層:");
                    PromptEntityResult entityResult = ed.GetEntity(entityOpts);

                    if (entityResult.Status == PromptStatus.OK)
                    {
                        ObjectId selectedId = entityResult.ObjectId;
                        Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);

                        // 獲取物件的圖層名稱
                        string layerName = ent.Layer;

                        // 遍歷圈選範圍內的物件
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity entityToModify = (Entity)tr.GetObject(objId, OpenMode.ForWrite);

                            // 如果物件的圖層與選擇的圖層相同，將顏色設為 ByLayer (256)
                            if (entityToModify.Layer == layerName)
                            {
                                entityToModify.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256);
                            }
                        }
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消選擇時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "圖層顏色設置完成，已設為隨層（ByLayer）。".Gz_StatusBarMsg();
        }


    }
}
