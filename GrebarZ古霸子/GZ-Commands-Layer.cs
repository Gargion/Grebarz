using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace Grebarz古霸子
{

    public class GZ_Commands_Layer
    {
        public static Database db = HostApplicationServices.WorkingDatabase;
        public static Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        public static Editor ed = doc.Editor;
        public static string LAY_P = "nnnnil";//紀錄凍結圖層用
        public static string LAY_LOCK = "nnnnil";//紀錄上鎖圖層用

        [CommandMethod("gz305")]
        public void Command305()
        {
            string tip = "\n--------------------" + Environment.NewLine + "指令說明: 將當前圖層，變為選取物件之圖層" + Environment.NewLine + "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 提示用戶選擇一個物件
            PromptEntityOptions options = new PromptEntityOptions("\n指向物件圖層為目前層:");
            PromptEntityResult result = ed.GetEntity(options);

            // 檢查用戶是否選擇了有效的物件
            if (result.Status == PromptStatus.OK)
            {
                ObjectId selectedId = result.ObjectId;

                using (Transaction tr = selectedId.Database.TransactionManager.StartTransaction())
                {
                    // 獲取選擇的物件
                    Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);

                    // 獲取圖層表
                    LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                    // 獲取物件的圖層名稱
                    string layerName = ent.Layer;

                    // 設置當前圖層
                    if (layerTable.Has(layerName) == true)
                    {
                        db.Clayer = layerTable[layerName];
                        tr.Commit();

                        string msg = "\n--------------------" + Environment.NewLine + "當前的圖層已設為：" + layerName;
                        msg.Gz_StatusBarMsg();
                    }
                }
            }
        }

        [CommandMethod("gz315")]
        public void Command315()
        {
            string tip = "\n--------------------" + Environment.NewLine + "指令說明: 選取目標圖層，將圈選物件之圖層指定為：目標圖層" + Environment.NewLine + "\n--------------------";
            tip.Gz_StatusBarMsg();
            // 提示用戶選擇一個物件
            PromptEntityOptions entityOptions = new PromptEntityOptions("\n選取物件，指定目標圖層:");
            PromptEntityResult entityResult = ed.GetEntity(entityOptions);

            // 檢查用戶是否選擇了有效的物件
            if (entityResult.Status == PromptStatus.OK)
            {
                ObjectId selectedEntityId = entityResult.ObjectId;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 獲取選擇的物件
                    Entity selectedEntity = (Entity)tr.GetObject(selectedEntityId, OpenMode.ForRead);

                    // 取得物件的圖層名稱
                    string targetLayerName = selectedEntity.Layer;

                    // 提示圈選多個物件
                    PromptSelectionOptions selectionOptions = new PromptSelectionOptions();
                    selectionOptions.MessageForAdding = "\n圈選物件:";

                    PromptSelectionResult selectionResult = ed.GetSelection(selectionOptions);

                    if (selectionResult.Status == PromptStatus.OK)
                    {
                        SelectionSet selectedSet = selectionResult.Value;

                        // 修改選取物件的圖層
                        foreach (SelectedObject selObj in selectedSet)
                        {
                            if (selObj != null)
                            {
                                Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForWrite) as Entity;
                                if (ent != null)
                                {
                                    ent.Layer = targetLayerName;
                                }
                            }
                        }

                        tr.Commit();

                        string msg = "\n--------------------" + Environment.NewLine + "已將選取的物件設為圖層：" + targetLayerName;
                        msg.Gz_StatusBarMsg();
                    }
                }
            }
        }
        
        [CommandMethod("gz3150")]
        public void ChangeToLayerZero()
        {
            string tip = "\n--------------------" + Environment.NewLine + "指令說明: 圈選圖層，圖層指定為：0圖層" + Environment.NewLine + "\n--------------------";
            tip.Gz_StatusBarMsg();
            PromptSelectionOptions selOptions = new PromptSelectionOptions();
            selOptions.MessageForAdding = "\n圈選物件:";
            PromptSelectionResult selResult;

            while ((selResult = ed.GetSelection(selOptions)).Status == PromptStatus.OK)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    SelectionSet ss = selResult.Value;
                    foreach (SelectedObject so in ss)
                    {
                        if (so != null)
                        {
                            Entity entity = (Entity)tr.GetObject(so.ObjectId, OpenMode.ForWrite);
                            entity.Layer = "0"; // Set layer to "0"
                        }
                    }
                    tr.Commit();
                }
            }
        }

        [CommandMethod("gz3151")]
        public void ChangeTextPropertiesBasedOnSelection()
        {
            // 指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 選取文字並圈選具有相同字串的文字，然後設置其圖層和顏色。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 選取文字，指定字串
            PromptEntityOptions options = new PromptEntityOptions("\n選取文字，指定字串:");
            options.SetRejectMessage("\n必須選擇文字物件。");
            options.AddAllowedClass(typeof(DBText), true);
            PromptEntityResult result = ed.GetEntity(options);

            if (result.Status == PromptStatus.OK)
            {
                ObjectId selectedId = result.ObjectId;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    DBText selectedText = tr.GetObject(selectedId, OpenMode.ForRead) as DBText;

                    if (selectedText != null)
                    {
                        string txt = selectedText.TextString;
                        string enlay = selectedText.Layer;
                        int encol = selectedText.ColorIndex;
                        string colorString = (encol == 0) ? "byblock" : (encol == 256) ? "bylayer" : encol.ToString();

                        // 設置圈選的文字
                        PromptSelectionOptions selOptions = new PromptSelectionOptions();
                        selOptions.MessageForAdding = "\n圈選文字:";
                        ed.WriteMessage(selOptions.MessageForAdding);

                        TypedValue[] filter = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.Start, "TEXT"),
                            new TypedValue((int)DxfCode.Text, txt)
                        };
                        SelectionFilter selFilter = new SelectionFilter(filter);
                        PromptSelectionResult selResult = ed.GetSelection(selOptions, selFilter);

                        if (selResult.Status == PromptStatus.OK)
                        {
                            SelectionSet ss = selResult.Value;
                            foreach (ObjectId id in ss.GetObjectIds())
                            {
                                DBText text = tr.GetObject(id, OpenMode.ForWrite) as DBText;
                                if (text != null)
                                {
                                    text.Layer = enlay;
                                    text.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, (short)encol);
                                }
                            }
                            tr.Commit();
                            string msg = "\n文字物件屬性已更新完成。";
                            msg.Gz_StatusBarMsg();
                        }
                    }
                }
            }
        }

        [CommandMethod("gz315C")]
        public void ChangeToCurrentLayer()
        {
            // 指令說明
            string tip = "\n--------------------" + Environment.NewLine + 
                "\n指令說明: 選擇物件並將其圖層更改為當前圖層"
                + Environment.NewLine + "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 提示用戶選擇一個或多個物件
            PromptSelectionOptions options = new PromptSelectionOptions();
            options.MessageForAdding = "\n圈選物件:";
            PromptSelectionResult result = ed.GetSelection(options);

            // 檢查用戶是否選擇了有效的物件
            if (result.Status == PromptStatus.OK)
            {
                SelectionSet selectionSet = result.Value;
                string currentLayer = db.Clayer.ToString();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject selectedObject in selectionSet)
                    {
                        if (selectedObject != null)
                        {
                            Entity ent = (Entity)tr.GetObject(selectedObject.ObjectId, OpenMode.ForWrite);
                            ent.Layer = currentLayer; // 設置物件的圖層為當前圖層
                        }
                    }
                    tr.Commit();
                }

                string msg = "\n--------------------" + Environment.NewLine + "選定物件的圖層已更改為當前圖層。" + Environment.NewLine + "--------------------";
                msg.Gz_StatusBarMsg();
            }
        }

        [CommandMethod("gz325")]
        public void ChangeSelectedObjectsLayerToCurrent()
        {
            // 指令說明
            string tip = "\n--------------------" + Environment.NewLine + 
                "\n指令說明: 選擇物件，並將相同圖層的物件更改至當前圖層" +
                Environment.NewLine + "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 提示用戶選擇一個物件
            PromptEntityOptions options = new PromptEntityOptions("\n選取物件，指定圖層:");
            PromptEntityResult result = ed.GetEntity(options);

            // 檢查用戶是否選擇了有效的物件
            if (result.Status == PromptStatus.OK)
            {
                ObjectId selectedId = result.ObjectId;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 獲取選擇的物件
                    Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);
                    string targetLayer = ent.Layer; // 獲取選定物件的圖層
                    string currentLayer = db.Clayer.ToString();

                    // 提示用戶圈選相同圖層的物件
                    PromptSelectionOptions selOptions = new PromptSelectionOptions();
                    selOptions.MessageForAdding = "\n圈選物件:";
                    TypedValue[] filter = new TypedValue[]
                    {
                        new TypedValue((int)DxfCode.LayerName, targetLayer)
                    };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);
                    PromptSelectionResult selResult = ed.GetSelection(selOptions, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet selectionSet = selResult.Value;

                        foreach (SelectedObject selectedObject in selectionSet)
                        {
                            if (selectedObject != null)
                            {
                                Entity objEntity = (Entity)tr.GetObject(selectedObject.ObjectId, OpenMode.ForWrite);
                                objEntity.Layer = currentLayer; // 設置物件的圖層為當前圖層
                            }
                        }
                        tr.Commit();

                        string msg = "\n--------------------" + Environment.NewLine + "選定物件的圖層已更改為當前圖層。" + Environment.NewLine + "--------------------";
                        msg.Gz_StatusBarMsg();
                    }
                }
            }
        }

        [CommandMethod("gz316")]
        public void ChangeSelectedObjectsColor()
        {
            // 指令說明
            string tip = "\n--------------------" + Environment.NewLine + "\n指令說明: 選擇物件並將圈選的物件更改為相同顏色" + Environment.NewLine + "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 提示用戶選擇一個物件
            PromptEntityOptions options = new PromptEntityOptions("\n選取物件，指定顏色:");
            PromptEntityResult result = ed.GetEntity(options);

            // 檢查用戶是否選擇了有效的物件
            if (result.Status == PromptStatus.OK)
            {
                ObjectId selectedId = result.ObjectId;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 獲取選擇的物件
                    Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);
                    
                    short colorIndex = (short)ent.ColorIndex;

                    // 判斷顏色是否是 ByLayer 或 ByBlock
                    string colorDescription = (colorIndex == 0) ? "byblock" : (colorIndex == 256) ? "bylayer" : colorIndex.ToString();

                    // 提示用戶圈選物件
                    PromptSelectionOptions selOptions = new PromptSelectionOptions();
                    selOptions.MessageForAdding = "\n圈選物件:";
                    PromptSelectionResult selResult = ed.GetSelection(selOptions);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet selectionSet = selResult.Value;

                        foreach (SelectedObject selectedObject in selectionSet)
                        {
                            if (selectedObject != null)
                            {
                                Entity objEntity = (Entity)tr.GetObject(selectedObject.ObjectId, OpenMode.ForWrite);
                                if (colorDescription == "bylayer")
                                {
                                    objEntity.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256);
                                }
                                else if (colorDescription == "byblock")
                                {
                                    objEntity.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                }
                                else
                                {
                                    objEntity.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex);
                                }
                            }
                        }
                        tr.Commit();

                        string msg = "\n--------------------" + Environment.NewLine + "選定的物件顏色已更改。" + Environment.NewLine + "--------------------";
                        msg.Gz_StatusBarMsg();
                    }
                }
            }
        }

        [CommandMethod("gz326")]
        public void ChangeSelectedObjectsTypeAndColor()
        {
            // 指令說明
            string tip = "\n--------------------" + Environment.NewLine + "\n指令說明: 選擇物件並將圈選的物件更改為相同類型和顏色" + Environment.NewLine + "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 提示用戶選擇一個物件
            PromptEntityOptions options = new PromptEntityOptions("\n選取物件，指定類型、顏色:");
            PromptEntityResult result = ed.GetEntity(options);

            // 檢查用戶是否選擇了有效的物件
            if (result.Status == PromptStatus.OK)
            {
                ObjectId selectedId = result.ObjectId;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 獲取選擇的物件
                    Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);
                    short colorIndex = (short)ent.ColorIndex;

                    // 確定顏色為 ByLayer 或 ByBlock
                    string colorDescription = (colorIndex == 0) ? "byblock" : (colorIndex == 256) ? "bylayer" : colorIndex.ToString();

                    // 獲取物件類型
                    string entityType = ent.GetType().Name;

                    // 提示用戶圈選物件
                    PromptSelectionOptions selOptions = new PromptSelectionOptions();
                    selOptions.MessageForAdding = "\n圈選物件:";
                    TypedValue[] filter = { new TypedValue((int)DxfCode.Start, entityType) };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);
                    PromptSelectionResult selResult = ed.GetSelection(selOptions, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet selectionSet = selResult.Value;

                        foreach (SelectedObject selectedObject in selectionSet)
                        {
                            if (selectedObject != null)
                            {
                                Entity objEntity = (Entity)tr.GetObject(selectedObject.ObjectId, OpenMode.ForWrite);
                                if (colorDescription == "bylayer")
                                {
                                    objEntity.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256);
                                }
                                else if (colorDescription == "byblock")
                                {
                                    objEntity.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                }
                                else
                                {
                                    objEntity.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex);
                                }
                            }
                        }
                        tr.Commit();

                        string msg = "\n--------------------" + Environment.NewLine + "選定的物件類型和顏色已更改。" + Environment.NewLine + "--------------------";
                        msg.Gz_StatusBarMsg();
                    }
                }
            }
        }

        [CommandMethod("gz336")]
        public void ChangeObjectsToLayerColor()
        {
            // 指令說明
            string tip = "\n--------------------" + Environment.NewLine + "\n指令說明: 選取物件並將圈選的物件更改為相同圖層顏色" + Environment.NewLine + "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 提示用戶選擇一個物件
            PromptEntityOptions options = new PromptEntityOptions("\n選取物件，指定圖層顏色:");
            PromptEntityResult result = ed.GetEntity(options);

            // 檢查用戶是否選擇了有效的物件
            if (result.Status == PromptStatus.OK)
            {
                ObjectId selectedId = result.ObjectId;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 獲取選擇的物件
                    Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);
                    short colorIndex = (short)ent.ColorIndex;

                    // 確定顏色為 ByLayer 或 ByBlock
                    string colorDescription = (colorIndex == 0) ? "byblock" : (colorIndex == 256) ? "bylayer" : colorIndex.ToString();

                    // 獲取物件的圖層名稱
                    string layerName = ent.Layer;

                    // 提示用戶圈選物件
                    PromptSelectionOptions selOptions = new PromptSelectionOptions();
                    selOptions.MessageForAdding = "\n圈選物件:";
                    TypedValue[] filter = { new TypedValue((int)DxfCode.LayerName, layerName) };
                    SelectionFilter selectionFilter = new SelectionFilter(filter);
                    PromptSelectionResult selResult = ed.GetSelection(selOptions, selectionFilter);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet selectionSet = selResult.Value;

                        foreach (SelectedObject selectedObject in selectionSet)
                        {
                            if (selectedObject != null)
                            {
                                Entity objEntity = (Entity)tr.GetObject(selectedObject.ObjectId, OpenMode.ForWrite);
                                if (colorDescription == "bylayer")
                                {
                                    objEntity.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256);
                                }
                                else if (colorDescription == "byblock")
                                {
                                    objEntity.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0);
                                }
                                else
                                {
                                    objEntity.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex);
                                }
                            }
                        }
                        tr.Commit();

                        string msg = "\n--------------------" + Environment.NewLine + "選定的物件圖層顏色已更改。" + Environment.NewLine + "--------------------";
                        msg.Gz_StatusBarMsg();
                    }
                }
            }
        }

        [CommandMethod("gz314")]
        public void FreezeUnselectedLayers()
        {
            // 指令說明
            string tip = "\n--------------------" + Environment.NewLine + "\n指令說明: 圈選物件，凍結未選圖層" + Environment.NewLine + "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 獲取當前圖層的 ObjectId
            ObjectId currentLayerId = db.Clayer;

            string currentLayerName = string.Empty;

            // 使用事務來獲取當前圖層的名稱
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTableRecord currentLayerRecord = (LayerTableRecord)tr.GetObject(currentLayerId, OpenMode.ForRead);
                currentLayerName = currentLayerRecord.Name; // 獲取當前圖層的名稱
                tr.Commit();
            }

            // 提示用戶圈選物件
            ed.WriteMessage("\n圈選物件，凍結未選圖層");
            PromptSelectionResult selectionResult = ed.GetSelection();

            if (selectionResult.Status == PromptStatus.OK)
            {
                SelectionSet selectedSet = selectionResult.Value;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 凍結所有層，除了當前層
                    var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                    foreach (ObjectId layerId in layerTable)
                    {
                        LayerTableRecord layerRecord = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                        if (layerRecord.Name != currentLayerName) // 排除當前圖層
                        {
                            layerRecord.IsFrozen = true; // 凍結層
                        }
                    }

                    // 遍歷選定的物件
                    for (int i = 0; i < selectedSet.Count; i++)
                    {
                        // 獲取物件 ID
                        ObjectId entityId = selectedSet.GetObjectIds()[i];
                        Entity entity = (Entity)tr.GetObject(entityId, OpenMode.ForRead);
                        string layerName = entity.Layer;

                        // 解凍選定物件的層
                        if (layerName != currentLayerName)
                        {
                            LayerTableRecord selectedLayer = (LayerTableRecord)tr.GetObject(layerTable[layerName], OpenMode.ForWrite);
                            selectedLayer.IsFrozen = false; // 解凍層
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage("\n--------------------" + Environment.NewLine + "所有未選的圖層已凍結 (當前圖層不可凍結)。" + Environment.NewLine + "--------------------");
            }
        }

        [CommandMethod("gz304")]
        public void FreezeLayerOfSelectedObjects()
        {
            // 簡單的說明文字
            string tip = "\n--------------------" + Environment.NewLine + "\n指令說明: 凍結選取物件圖層" + Environment.NewLine + "\n--------------------";
            tip.Gz_StatusBarMsg();

            
            PromptEntityOptions options = new PromptEntityOptions("\n凍結選取物件圖層: ");
            PromptEntityResult result;

            while (true)
            {
                result = ed.GetEntity(options);
                if (result.Status != PromptStatus.OK) break;

                ObjectId selectedId = result.ObjectId;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 獲取選擇的物件
                    Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);
                    string layerName = ent.Layer;

                    // 添加圖層到列表
                    LAY_P = string.Concat(LAY_P, ",", layerName);

                    // 凍結圖層
                    LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (layerTable.Has(layerName))
                    {
                        LayerTableRecord layerRecord = (LayerTableRecord)tr.GetObject(layerTable[layerName], OpenMode.ForWrite);
                        layerRecord.IsFrozen = true; // 凍結圖層
                        LAY_P = "nnnnil";
                    }

                    tr.Commit();
                }
            }
            ed.WriteMessage("\n凍結的圖層: " + LAY_P);
        }

        [CommandMethod("gz304Q")]
        public void ThawAllLayers()
        {
            // 這裡加上簡單的說明文字
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 解凍所有圖層" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 解凍所有圖層
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId layerId in layerTable)
                {
                    LayerTableRecord layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    layer.IsFrozen = false; // 解凍圖層
                }
                tr.Commit();
            }

            ed.WriteMessage("\n所有圖層已解凍。");
        }

        [CommandMethod("gz304T")]
        public void ListFrozenLayers()
        {
            // 這裡加上簡單的說明文字
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 列出所有凍結的圖層" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 獲取當前圖層的 ObjectId
            ObjectId currentLayerId = db.Clayer;

            string currentLayerName = string.Empty;

            // 使用事務來獲取當前圖層的名稱
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTableRecord currentLayerRecord = (LayerTableRecord)tr.GetObject(currentLayerId, OpenMode.ForRead);
                currentLayerName = currentLayerRecord.Name; // 獲取當前圖層的名稱
                tr.Commit();
            }

            List<string> frozenLayers = new List<string>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId layerId in layerTable)
                {
                    LayerTableRecord layerRecord = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    if (layerRecord.IsFrozen && layerRecord.Name != currentLayerName)
                    {
                        frozenLayers.Add(layerRecord.Name); // 將凍結的層名稱加入列表
                    }
                }

                tr.Commit();
            }

            // 如果有凍結的圖層，呼叫圖層選單顯示函數
            if (frozenLayers.Count > 0)
            {
                WPF.ShowWPF<WPF_LayerSelection>(frozenLayers);
            }
            else
            {
                string msg = "沒有找到凍結中的圖層。";
                msg.Gz_MsgBoxInf();
            }
        }

        [CommandMethod("gz304P")]
        public void Gz304P()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 解凍前次由指令gz304凍結的圖層" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 判斷 LAY_P 是否為 "nnnnil"
            if (LAY_P == "nnnnil")
            {
                string msg = "找不到由指令gz304凍結的圖層。";
                msg.Gz_MsgBoxWarning();
                return; // 退出程序
            }

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                if (layerTable.Has(LAY_P))
                {
                    // 將 LAY_P 所指定的圖層設置為解凍
                    LayerTableRecord layerRecord = (LayerTableRecord)tr.GetObject(layerTable[LAY_P], OpenMode.ForWrite);
                    layerRecord.IsFrozen = false;
                }

                tr.Commit();
            }
        }

        [CommandMethod("gz324")]
        public void Gz324()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 鎖住選取物件的圖層。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 重置 LAY_LOCK 變量
                LAY_LOCK = "nnnnil";

                // 進入循環以選取物件並鎖住圖層
                bool continueSelection = true;
                while (continueSelection)
                {
                    PromptEntityOptions options = new PromptEntityOptions("\n鎖住選取物件圖層: ");
                    PromptEntityResult result = ed.GetEntity(options);

                    if (result.Status == PromptStatus.OK)
                    {
                        ObjectId selectedId = result.ObjectId;
                        Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);

                        // 獲取選取物件的圖層名稱
                        string layerName = ent.Layer;

                        // 將圖層名稱加入 LAY_LOCK 變量
                        LAY_LOCK += "," + layerName;

                        // 鎖住該圖層
                        LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                        if (layerTable.Has(layerName))
                        {
                            LayerTableRecord layerRecord = (LayerTableRecord)tr.GetObject(layerTable[layerName], OpenMode.ForWrite);
                            layerRecord.IsLocked = true; // 鎖住圖層
                        }
                    }
                    else
                    {
                        // 如果使用者沒有選取物件，退出循環
                        continueSelection = false;
                    }
                }

                tr.Commit(); // 提交變更
            }
            // 程序結束，輸出完成提示
            "鎖住圖層的操作已完成。".Gz_StatusBarMsg();
        }

        [CommandMethod("gz324Q")]
        public void Gz324Q()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 解鎖所有圖層。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 獲取圖層表
                LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                // 解鎖所有圖層
                foreach (ObjectId layerId in layerTable)
                {
                    LayerTableRecord layerRecord = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    if (layerRecord.IsLocked)
                    {
                        layerRecord.IsLocked = false; // 解鎖圖層
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            string msg = "\n所有圖層已解鎖。";
            msg.Gz_StatusBarMsg();
        }

        [CommandMethod("gzDLAY")]
        public void GzDLAY()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 刪除與所選物件同一圖層的所有物件。" + Environment.NewLine +
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
                    // 提示用戶選擇一個物件
                    PromptEntityOptions options = new PromptEntityOptions("\n指向物件，同圖層物件全部刪除: ");
                    PromptEntityResult result = ed.GetEntity(options);

                    if (result.Status == PromptStatus.OK)
                    {
                        ObjectId selectedId = result.ObjectId;
                        Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);

                        // 獲取選取物件的圖層名稱
                        string layerName = ent.Layer;

                        // 選取該圖層上的所有物件
                        TypedValue[] filter = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.LayerName, layerName)
                        };
                        SelectionFilter selectionFilter = new SelectionFilter(filter);
                        PromptSelectionResult selectionResult = ed.SelectAll(selectionFilter);

                        if (selectionResult.Status == PromptStatus.OK)
                        {
                            SelectionSet ss = selectionResult.Value;

                            // 刪除所有與所選物件同圖層的物件
                            foreach (ObjectId objId in ss.GetObjectIds())
                            {
                                Entity layerEntity = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                                layerEntity.Erase(); // 刪除物件
                            }
                        }
                    }
                    else
                    {
                        // 當用戶不再選擇物件時退出循環
                        continueSelection = false;
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 結束時輸出完成訊息
            "同圖層物件刪除操作已完成。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzEEL")]
        public void GzEEL()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 圈選範圍內的物件，並指定圖層的物件全部刪除。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 提示用戶圈選範圍
                PromptSelectionOptions selOpts = new PromptSelectionOptions();
                selOpts.MessageForAdding = "\n圈選範圍: ";
                PromptSelectionResult selResult = ed.GetSelection(selOpts);

                if (selResult.Status != PromptStatus.OK)
                {
                    return; // 如果未選擇範圍，則退出
                }

                SelectionSet ssAll = selResult.Value; // 圈選範圍內的所有物件

                List<string> selectedLayers = new List<string>();
                bool continueSelection = true;

                while (continueSelection)
                {
                    // 提示用戶選擇物件並指定圖層
                    PromptEntityOptions entityOpts = new PromptEntityOptions("\n圈選物件，指定圖層: ");
                    PromptEntityResult entityResult = ed.GetEntity(entityOpts);

                    if (entityResult.Status == PromptStatus.OK)
                    {
                        ObjectId selectedId = entityResult.ObjectId;
                        Entity ent = (Entity)tr.GetObject(selectedId, OpenMode.ForRead);

                        // 將選中的物件的圖層加入清單
                        string selectedLayer = ent.Layer;
                        if (!selectedLayers.Contains(selectedLayer))
                        {
                            selectedLayers.Add(selectedLayer);
                        }

                        // 遍歷所有已圈選的物件，選取與指定圖層相同的物件
                        List<ObjectId> objectsToErase = new List<ObjectId>();
                        foreach (ObjectId objId in ssAll.GetObjectIds())
                        {
                            Entity layerEntity = (Entity)tr.GetObject(objId, OpenMode.ForRead);
                            if (selectedLayers.Contains(layerEntity.Layer))
                            {
                                objectsToErase.Add(objId);
                            }
                        }

                        // 刪除與指定圖層相同的物件
                        foreach (ObjectId objId in objectsToErase)
                        {
                            Entity eraseEntity = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                            eraseEntity.Erase(); // 刪除物件
                        }
                    }
                    else
                    {
                        continueSelection = false; // 當用戶未選擇物件時退出循環
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束時輸出提示
            "指定圖層的物件已刪除。".Gz_StatusBarMsg();
        }

        [CommandMethod("gzLEE")]
        public void GzLEE()
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 先圈選指定(多個)圖層，再圈選刪除範圍，範圍內被指定的(多個)圖層物件刪除。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            List<string> selectedLayers = new List<string>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                bool continueSelection = true;

                // 讓用戶選擇圖層
                while (continueSelection)
                {
                    // 提示用戶選擇物件來指定圖層
                    PromptSelectionOptions selOpts = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n圈選物件，可指定多個圖層:"
                    };
                    PromptSelectionResult selResult = ed.GetSelection(selOpts);

                    if (selResult.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selResult.Value;
                        foreach (ObjectId objId in ss.GetObjectIds())
                        {
                            Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForRead);
                            string layerName = ent.Layer;

                            // 將圖層名加入選擇列表中
                            if (!selectedLayers.Contains(layerName))
                            {
                                selectedLayers.Add(layerName);
                            }
                        }
                    }
                    else
                    {
                        continueSelection = false; // 用戶取消操作，退出循環
                    }
                }

                // 將選擇的圖層名稱轉換為逗號分隔的字符串
                string layersToDelete = string.Join(",", selectedLayers.ToArray());

                // 如果已選擇了圖層，進行範圍選擇和刪除
                if (selectedLayers.Count > 0)
                {
                    // 提示用戶選擇範圍，刪除與指定圖層相同的物件
                    ed.WriteMessage("\n圈選刪除範圍:");
                    continueSelection = true;

                    while (continueSelection)
                    {
                        TypedValue[] filter = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.LayerName, layersToDelete)
                        };
                        SelectionFilter selectionFilter = new SelectionFilter(filter);
                        PromptSelectionResult selResult = ed.GetSelection(selectionFilter);

                        if (selResult.Status == PromptStatus.OK)
                        {
                            SelectionSet ssToDelete = selResult.Value;

                            foreach (ObjectId objId in ssToDelete.GetObjectIds())
                            {
                                Entity entityToErase = (Entity)tr.GetObject(objId, OpenMode.ForWrite);
                                entityToErase.Erase(); // 刪除物件
                            }
                        }
                        else
                        {
                            continueSelection = false; // 用戶取消操作，退出循環
                        }
                    }
                }

                tr.Commit(); // 提交事務
            }

            // 程序結束後輸出提示
            "刪除操作已完成。".Gz_StatusBarMsg();
        }

        
        

    }
}
