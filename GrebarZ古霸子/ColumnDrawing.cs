using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using Exception = System.Exception;

namespace Grebarz古霸子
{
    public class ColumnDrawing
    {
        private Dictionary<string, object> originalSettings = new Dictionary<string, object>();

        // 保存AutoCAD系統變量的字段
        public double VTEnable { get; private set; }      // 視圖切換設置
        public double Osmode { get; private set; }        // 捕捉模式設置
        public string Clayer { get; private set; }        // 當前圖層
        public string Cecolor { get; private set; }       // 當前顏色設置
        public double Dimzin { get; private set; }        // 標註設置
        public double DrawOrderCtl { get; private set; }  // 繪圖順序控制
        public double FilletRadius { get; private set; }  // 圓角半徑設置

        // 用於標記是否讀取了搭接數據
        public string KeyReadJS { get; private set; }

        // 定義 FL_KG 為字典，用來存儲樓層名稱及其對應的數據
        private Dictionary<string, int> FL_KG = new Dictionary<string, int>();

        // 定義樓層高度與混凝土強度的對照表
        private Dictionary<string, double> floorHeights;  // 用於存儲樓層名稱與樓層高度
        private Dictionary<string, string> concreteStrengths;  // 用於存儲樓層名稱與混凝土強度


        // 主執行入口
        [CommandMethod("gzCC")]
        public void Execute()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // 1. 初始化設置
                InitializeSettings();

                // 2. 讀取樓層數據（包括高度和混凝土強度）
                ed.WriteMessage("\n正在讀取樓層配置...");
                ReadFloorConfig();

                // 檢查是否成功讀取樓層數據
                if (floorHeights == null || floorHeights.Count == 0)
                {
                    ed.WriteMessage("\n未能成功讀取樓層配置。");
                    return;
                }

                // 3. 處理圖形中的多段線框線
                ed.WriteMessage("\n正在處理框線...");
                Process701Q();

                // 4. 選擇並處理續接數據
                ed.WriteMessage("\n正在讀取接續數據...");
                List<ConnectorData> connectors = ReadJSData();

                if (connectors == null || connectors.Count == 0)
                {
                    ed.WriteMessage("\n未能讀取接續數據。");
                }

                // 5. 繪製標註和處理續接信息
                ed.WriteMessage("\n繪製續接標註...");
                foreach (var connector in connectors)
                {
                    // 假設 DrawConnection 和 AnnotateConnections 是處理續接的繪圖和標註方法
                    DrawColumn(connector);
                    AnnotateConnections(connector);
                }

                // 6. 恢復設置
                ed.WriteMessage("\n恢復設置...");
                RestoreSettings();

                // 最終完成提示
                ed.WriteMessage("\n操作完成！");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n操作失敗: {ex.Message}");
            }
        }


        // 保存原始設置
        public void SaveAutoCADSettings()
        {
            // 獲取需要保存的系統變量並存入字典
            originalSettings["clayer"] = Application.GetSystemVariable("clayer");
            originalSettings["cecolor"] = Application.GetSystemVariable("cecolor");
            originalSettings["osmode"] = Application.GetSystemVariable("osmode");
            originalSettings["filletrad"] = Application.GetSystemVariable("filletrad");
            originalSettings["ltscale"] = Application.GetSystemVariable("ltscale");
            originalSettings["dimscale"] = Application.GetSystemVariable("dimscale");
            originalSettings["textsize"] = Application.GetSystemVariable("textsize");

            // 提示保存完成
            Document doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n原始系統變量已保存。");
        }

        // 初始化繪圖設置
        private void InitializeSettings()
        {
            // 關閉命令回顯
            Application.SetSystemVariable("cmdecho", 0);
            Application.SetSystemVariable("attreq", 1);
            Application.SetSystemVariable("nomutt", 0);
            Application.SetSystemVariable("mirrtext", 0);

            // 開始撤銷記錄
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute("undo be ", true, false, false);

            // 保存當前設置
            double vtEnable = Convert.ToDouble(Application.GetSystemVariable("VTENABLE"));
            double osmode = Convert.ToDouble(Application.GetSystemVariable("osmode"));
            string clayer = Application.GetSystemVariable("clayer").ToString();
            string cecolor = Application.GetSystemVariable("cecolor").ToString();
            double dimzin = Convert.ToDouble(Application.GetSystemVariable("dimzin"));
            double drawOrderCtl = Convert.ToDouble(Application.GetSystemVariable("draworderctl"));

            // 更新設置
            Application.SetSystemVariable("VTENABLE", 0);
            Application.SetSystemVariable("osmode", 0);
            Application.SetSystemVariable("dimzin", 8);

            if (drawOrderCtl > 0)
            {
                Application.SetSystemVariable("draworderctl", 0);
            }

            // 保存設置值供後續恢復使用
            SaveSettings();
        }

        // 保存當前設置
        private void SaveSettings()
        {
            // 獲取並保存需要恢復的系統變量
            this.VTEnable = Convert.ToDouble(Application.GetSystemVariable("VTENABLE"));
            this.Osmode = Convert.ToDouble(Application.GetSystemVariable("osmode"));
            this.Clayer = Application.GetSystemVariable("clayer").ToString();
            this.Cecolor = Application.GetSystemVariable("cecolor").ToString();
            this.Dimzin = Convert.ToDouble(Application.GetSystemVariable("dimzin"));
            this.DrawOrderCtl = Convert.ToDouble(Application.GetSystemVariable("draworderctl"));

            // 保存圓角半徑
            this.FilletRadius = Convert.ToDouble(Application.GetSystemVariable("filletrad"));

            // 檢查並顯示保存的設置
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n保存當前系統設置。");
        }

        public List<ConnectorData> ReadJSData()
        {
            // 從 AutoCAD 獲取當前圖紙的路徑
            string dwgPrefix = Application.GetSystemVariable("dwgprefix").ToString();
            string fileName = dwgPrefix + "7CD.XLS";  // 使用字符串拼接路徑

            List<ConnectorData> connectorDataList = new List<ConnectorData>();

            try
            {
                // 使用 ClosedXML 打開 Excel 文件
                using (var workbook = new XLWorkbook(fileName))
                {
                    // 獲取名為 "接續" 的工作表
                    var worksheet = workbook.Worksheet("接續");

                    // 假設數據從第二行開始（第一行是標題）
                    int row = 2;
                    while (!worksheet.Cell(row, 1).IsEmpty())
                    {
                        // A欄: 接續類型
                        string connectorType = worksheet.Cell(row, 1).GetString();

                        // B欄: X 位置
                        double positionX = worksheet.Cell(row, 2).GetDouble();

                        // C欄: Y 位置
                        double positionY = worksheet.Cell(row, 3).GetDouble();

                        // D欄: 寬度
                        double width = worksheet.Cell(row, 4).GetDouble();

                        // E欄: 高度
                        double height = worksheet.Cell(row, 5).GetDouble();

                        // F欄: 主筋尺寸
                        string mainBarSize = worksheet.Cell(row, 6).GetString();

                        // G欄: 主筋數量
                        int mainBarCount = worksheet.Cell(row, 7).GetValue<int>(); // 使用 GetValue<int>()

                        // H欄: 箍筋尺寸
                        string stirrupSize = worksheet.Cell(row, 8).GetString();

                        // I欄: 箍筋數量
                        int stirrupCount = worksheet.Cell(row, 9).GetValue<int>(); // 使用 GetValue<int>()

                        // J欄: 是否有上續接
                        bool isConnectedAbove = worksheet.Cell(row, 10).GetValue<bool>(); // 使用 GetValue<bool>()

                        // K欄: 是否有下續接
                        bool isConnectedBelow = worksheet.Cell(row, 11).GetValue<bool>(); // 使用 GetValue<bool>()

                        // 創建並填充 ConnectorData 對象
                        ConnectorData connectorData = new ConnectorData
                        {
                            ConnectorType = connectorType,
                            PositionX = positionX,
                            PositionY = positionY,
                            Width = width,
                            Height = height,
                            MainBarSize = mainBarSize,
                            MainBarCount = mainBarCount,
                            StirrupSize = stirrupSize,
                            StirrupCount = stirrupCount,
                            IsConnectedAbove = isConnectedAbove,
                            IsConnectedBelow = isConnectedBelow
                        };

                        // 添加到列表
                        connectorDataList.Add(connectorData);
                        row++;  // 移動到下一行
                    }
                }

                // 標記接續數據讀取成功
                Document doc = Application.DocumentManager.MdiActiveDocument;
                doc.Editor.WriteMessage("\n接續數據讀取完成。");

            }
            catch (Exception ex)
            {
                // 讀取過程中如果出現錯誤，提示用戶
                Document doc = Application.DocumentManager.MdiActiveDocument;
                doc.Editor.WriteMessage($"\n讀取接續數據失敗: {ex.Message}");
            }

            return connectorDataList;
        }












        // 單個柱子的繪圖邏輯
        private void DrawColumn(ConnectorData connector)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 獲取模型空間塊表記錄
                BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // 1. 繪製柱子的矩形（假設柱子是矩形）
                Point3d basePoint = new Point3d(connector.PositionX, connector.PositionY, 0);  // 使用 ConnectorData 的位置
                double width = connector.Width;
                double height = connector.Height;

                Polyline columnRect = new Polyline(4);

                // 添加頂點，創建矩形
                columnRect.AddVertexAt(0, new Point2d(basePoint.X, basePoint.Y), 0, 0, 0);
                columnRect.AddVertexAt(1, new Point2d(basePoint.X + width, basePoint.Y), 0, 0, 0);
                columnRect.AddVertexAt(2, new Point2d(basePoint.X + width, basePoint.Y + height), 0, 0, 0);
                columnRect.AddVertexAt(3, new Point2d(basePoint.X, basePoint.Y + height), 0, 0, 0);
                columnRect.Closed = true;

                // 設置圖層
                columnRect.Layer = "COLUMN";
                btr.AppendEntity(columnRect);
                tr.AddNewlyCreatedDBObject(columnRect, true);

                // 2. 繪製主筋信息標註（假設主筋在圖形下方）
                string mainBarInfo = $"主筋: {connector.MainBarSize} x {connector.MainBarCount}";
                DBText mainBarText = new DBText
                {
                    Position = new Point3d(basePoint.X + width / 2, basePoint.Y - 20, 0),
                    TextString = mainBarInfo,
                    Height = 15,
                    Layer = "ANNOTATION"
                };
                btr.AppendEntity(mainBarText);
                tr.AddNewlyCreatedDBObject(mainBarText, true);

                // 3. 繪製箍筋信息標註
                string stirrupInfo = $"箍筋: {connector.StirrupSize} x {connector.StirrupCount}";
                DBText stirrupText = new DBText
                {
                    Position = new Point3d(basePoint.X + width / 2, basePoint.Y - 40, 0),
                    TextString = stirrupInfo,
                    Height = 15,
                    Layer = "ANNOTATION"
                };
                btr.AppendEntity(stirrupText);
                tr.AddNewlyCreatedDBObject(stirrupText, true);

                // 提交事務
                tr.Commit();
            }
        }



        //處裡主筋續接、箍筋
        private void AnnotateConnections(ConnectorData connector)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 獲取模型空間
                BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                Point3d basePoint = new Point3d(connector.PositionX, connector.PositionY, 0);

                // 1. 標註上續接（如果有）
                if (connector.IsConnectedAbove)
                {
                    Line upperConnectionLine = new Line(
                        new Point3d(basePoint.X, basePoint.Y + connector.Height, 0),
                        new Point3d(basePoint.X + connector.Width / 2, basePoint.Y + connector.Height + 50, 0)
                    );
                    upperConnectionLine.Layer = "ANNOTATION";
                    btr.AppendEntity(upperConnectionLine);
                    tr.AddNewlyCreatedDBObject(upperConnectionLine, true);

                    // 添加標註文字
                    CreateText($"上部續接: {connector.ConnectorType}",
                               new Point3d(basePoint.X + connector.Width / 2, basePoint.Y + connector.Height + 60, 0),
                               15, "ANNOTATION");
                }

                // 2. 標註下續接（如果有）
                if (connector.IsConnectedBelow)
                {
                    Line lowerConnectionLine = new Line(
                        new Point3d(basePoint.X, basePoint.Y, 0),
                        new Point3d(basePoint.X + connector.Width / 2, basePoint.Y - 50, 0)
                    );
                    lowerConnectionLine.Layer = "ANNOTATION";
                    btr.AppendEntity(lowerConnectionLine);
                    tr.AddNewlyCreatedDBObject(lowerConnectionLine, true);

                    // 添加標註文字
                    CreateText($"下部續接: {connector.ConnectorType}",
                               new Point3d(basePoint.X + connector.Width / 2, basePoint.Y - 60, 0),
                               15, "ANNOTATION");
                }

                tr.Commit();
            }

            // 提示續接標註完成
            ed.WriteMessage($"\n續接標註完成: {connector.ConnectorType}");
        }



        //private double CalculateElevationForColumn(int index, List<ColumnData> columnList)
        //{
        //    double elevation = 0;

        //    // 只有當不是第一根柱子時才進行累加
        //    if (index > 0)
        //    {
        //        // 累加之前所有樓層的高度
        //        elevation = columnList.Take(index).Sum(c => c.FloorHeight);
        //    }

        //    return elevation;
        //}

        

        ////箍筋標註
        //private void AnnotateStirrup(ColumnData column)
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;

        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        BlockTableRecord btr = GetBlockTableRecord(tr);

        //        // 添加箍筋標註
        //        DBText stirrupText = new DBText
        //        {
        //            Position = new Point3d(column.Position.X, column.Position.Y - 20, 0),
        //            TextString = $"箍筋: {column.StirrupSize} x {column.StirrupCount}",
        //            Height = 15,
        //            Layer = "COLDAT"
        //        };
        //        btr.AppendEntity(stirrupText);
        //        tr.AddNewlyCreatedDBObject(stirrupText, true);

        //        tr.Commit();
        //    }
        //}

        public ColumnDrawing()
        {
            // 初始化對照表
            floorHeights = new Dictionary<string, double>();
            concreteStrengths = new Dictionary<string, string>();
        }

        //這個方法將對應於 LISP 中 READ-FL-C 函數的邏輯，讀取樓層參數。
        private void ReadFloorConfig()
        {
            // 從 AutoCAD 獲取當前圖紙的路徑
            string dwgPrefix = Application.GetSystemVariable("dwgprefix").ToString();
            string fileName = dwgPrefix + "7CD.XLS";  // 使用字符串拼接路徑

            // 清空之前的對照表
            floorHeights.Clear();
            concreteStrengths.Clear();

            try
            {
                // 使用 ClosedXML 打開 Excel 文件
                using (var workbook = new XLWorkbook(fileName))
                {
                    // 獲取名為 "樓層" 的工作表
                    var worksheet = workbook.Worksheet("樓層");

                    // 假設數據從第二行開始（第一行是標題）
                    int row = 2;
                    while (!worksheet.Cell(row, 1).IsEmpty())
                    {
                        // A欄是樓層名稱
                        string floorName = worksheet.Cell(row, 1).GetString();

                        // B欄是樓層高度，轉換為 double
                        double floorHeight = worksheet.Cell(row, 2).GetDouble();

                        // C欄是混凝土強度，存儲為字符串
                        string concreteStrength = worksheet.Cell(row, 3).GetString();

                        // 添加樓層名稱和高度到對照表
                        floorHeights[floorName] = floorHeight;

                        // 添加樓層名稱和混凝土強度到對照表
                        concreteStrengths[floorName] = concreteStrength;

                        row++;  // 移動到下一行
                    }
                }

                // 提示已讀取完成
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n樓層信息與混凝土強度讀取完成。");
            }
            catch (Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"讀取樓層 Excel 文件失敗: {ex.Message}");
            }
        }

        


        private void Process701Q()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // 選擇 "文字_fl_c" 圖層上的所有多段線實體
            var selection = ed.SelectAll(new SelectionFilter(new[]
            {
        new TypedValue((int)DxfCode.LayerName, "文字_fl_c"),
        new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
    }));

            if (selection.Status != PromptStatus.OK)
            {
                ed.WriteMessage("不存在 '文字_fl_c' 圖層上的框線。");
                return;
            }

            // 如果樓層高度和混凝土強度的對照表尚未加載，則需要讀取樓層信息
            if (floorHeights == null || floorHeights.Count == 0)
            {
                ReadFloorConfig();  // 加載樓層數據
            }

            // 用於存儲排序的數據
            List<SortedData> sortedList = new List<SortedData>();

            // 遍歷選中的多段線實體
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                foreach (ObjectId objId in selection.Value.GetObjectIds())
                {
                    Polyline polyline = tr.GetObject(objId, OpenMode.ForRead) as Polyline;

                    if (polyline != null)
                    {
                        // 獲取框線的邊界點 PT1 和 PT2
                        Point3d pt1, pt2;
                        GetBoxCCM(polyline, out pt1, out pt2);

                        // 假設 GetFloorNameFromPolyline 根據 Polyline 的屬性來判斷樓層名稱
                        string floorName = GetFloorNameFromPolyline(polyline);

                        // 從樓層高度和混凝土強度的對照表中查找對應的數據
                        if (floorHeights.ContainsKey(floorName))
                        {
                            double floorHeight = floorHeights[floorName];
                            string concreteStrength = concreteStrengths.ContainsKey(floorName) ? concreteStrengths[floorName] : "N/A";

                            // 將數據存入 SortedData
                            SortedData data = new SortedData
                            {
                                FloorNumber = floorName,  // 樓層名稱
                                FloorHeight = floorHeight,  // 樓層高度
                                ConcreteStrength = concreteStrength,  // 混凝土強度
                                Points = new List<Point3d> { pt1, pt2 }
                            };

                            sortedList.Add(data);
                        }
                    }
                }
                tr.Commit();
            }

            // 對數據進行排序處理
            sortedList = SortColumnOrRow(sortedList);

            // 繪製或處理排序後的數據
            DrawSortedData(sortedList);

            // 放大視圖以適應範圍
            ZoomToFit();
        }

        private void GetBoxCCM(Entity entity, out Point3d pt1, out Point3d pt2)
        {
            // 獲取邊界框
            Extents3d? bounds = entity.GeometricExtents;
            if (bounds.HasValue)
            {
                pt1 = bounds.Value.MinPoint;
                pt2 = bounds.Value.MaxPoint;
            }
            else
            {
                pt1 = Point3d.Origin;
                pt2 = Point3d.Origin;
            }
        }

        //負責根據已排序的樓層數據繪製圖形，包括標註樓層高度、混凝土強度以及多段線的繪製。
        private void DrawSortedData(List<SortedData> sortedData)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 獲取模型空間塊表記錄
                BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                foreach (SortedData data in sortedData)
                {
                    // 1. 繪製樓層框線
                    if (data.Points.Count >= 2)
                    {
                        Polyline polyline = new Polyline(4);
                        polyline.AddVertexAt(0, new Point2d(data.Points[0].X, data.Points[0].Y), 0, 0, 0);
                        polyline.AddVertexAt(1, new Point2d(data.Points[1].X, data.Points[0].Y), 0, 0, 0);
                        polyline.AddVertexAt(2, new Point2d(data.Points[1].X, data.Points[1].Y), 0, 0, 0);
                        polyline.AddVertexAt(3, new Point2d(data.Points[0].X, data.Points[1].Y), 0, 0, 0);
                        polyline.Closed = true;
                        polyline.Layer = "FLOOR_BOUNDARY";
                        btr.AppendEntity(polyline);
                        tr.AddNewlyCreatedDBObject(polyline, true);
                    }

                    // 2. 標註樓層高度和混凝土強度
                    string annotationText = $"樓層: {data.FloorNumber}, 高度: {data.FloorHeight} mm, 混凝土: {data.ConcreteStrength}";
                    Point3d annotationPosition = new Point3d(
                        (data.Points[0].X + data.Points[1].X) / 2,
                        (data.Points[0].Y + data.Points[1].Y) / 2,
                        0
                    );
                    CreateText(annotationText, annotationPosition, 15, "ANNOTATION");
                }

                // 提交事務，保存繪製的圖形和標註
                tr.Commit();
            }

            ed.WriteMessage("\n樓層數據繪製和標註已完成。");
        }

        //將 AutoCAD 視圖縮放到合適的範圍，以便將所有繪製的對象都顯示在視圖中。
        private void ZoomToFit()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 獲取當前視口
                ViewTableRecord view = ed.GetCurrentView();

                // 計算所有對象的邊界框
                Extents3d extents = new Extents3d();
                bool hasValidExtents = false;

                BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId objId in modelSpace)
                {
                    Entity entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;

                    if (entity != null && entity.Visible)
                    {
                        // 更新邊界框
                        if (hasValidExtents)
                        {
                            extents.AddExtents(entity.GeometricExtents);
                        }
                        else
                        {
                            extents = entity.GeometricExtents;
                            hasValidExtents = true;
                        }
                    }
                }

                // 如果存在有效的邊界框
                if (hasValidExtents)
                {
                    // 計算縮放比率，讓視圖適應所有對象
                    view.CenterPoint = new Point2d(
                        (extents.MinPoint.X + extents.MaxPoint.X) / 2.0,
                        (extents.MinPoint.Y + extents.MaxPoint.Y) / 2.0
                    );
                    view.Height = extents.MaxPoint.Y - extents.MinPoint.Y;
                    view.Width = extents.MaxPoint.X - extents.MinPoint.X;

                    // 為了留出邊界，將視圖放大一點
                    view.Height *= 1.1;
                    view.Width *= 1.1;

                    // 更新視圖
                    ed.SetCurrentView(view);
                }

                tr.Commit();
            }

            ed.WriteMessage("\n視圖已縮放到適應所有對象。");
        }



        // 恢復保存的系統設置
        private void RestoreSettings()
        {
            // 恢復 VTENABLE 設置
            Application.SetSystemVariable("VTENABLE", this.VTEnable);

            // 恢復捕捉模式
            Application.SetSystemVariable("osmode", this.Osmode);

            // 恢復圖層和顏色
            Application.SetSystemVariable("clayer", this.Clayer);
            Application.SetSystemVariable("cecolor", this.Cecolor);

            // 恢復標註設置
            Application.SetSystemVariable("dimzin", this.Dimzin);

            // 恢復繪圖順序控制設置
            if (this.DrawOrderCtl > 0)
            {
                Application.SetSystemVariable("draworderctl", this.DrawOrderCtl);
            }

            // 恢復圓角半徑
            Application.SetSystemVariable("filletrad", this.FilletRadius);

            // 提示用戶設置已恢復
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n系統設置已恢復。");
        }

        
        private void SetVariable(string variableName, object value)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // 檢查系統變量是否存在
                if (Application.GetSystemVariable(variableName) != null)
                {
                    // 設置系統變量的值
                    Application.SetSystemVariable(variableName, value);
                    ed.WriteMessage($"\n設置變量: {variableName} = {value}");
                }
                else
                {
                    ed.WriteMessage($"\n變量 {variableName} 不存在。");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n設置變量 {variableName} 時發生錯誤: {ex.Message}");
            }
        }



        private string GetFloorNameFromPolyline(Polyline polyline)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            string floorName = string.Empty;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 1. 嘗試從附近的文字標註中提取樓層名稱
                TypedValue[] filterList = new TypedValue[]
                {
            new TypedValue((int)DxfCode.Start, "TEXT"),  // 尋找文本實體
            new TypedValue((int)DxfCode.LayerName, "樓層標註")  // 假設樓層標註的圖層名為 "樓層標註"
                };

                SelectionFilter filter = new SelectionFilter(filterList);
                PromptSelectionResult selection = ed.SelectAll(filter);

                if (selection.Status == PromptStatus.OK)
                {
                    foreach (SelectedObject obj in selection.Value)
                    {
                        DBText text = tr.GetObject(obj.ObjectId, OpenMode.ForRead) as DBText;

                        if (text != null)
                        {
                            // 檢查文本是否位於多段線的附近
                            double distance = polyline.GeometricExtents.MinPoint.DistanceTo(text.Position);

                            if (distance < 500)  // 假設500單位內的標註是有效的樓層名稱
                            {
                                floorName = text.TextString.Trim();
                                break;  // 找到樓層名稱後退出
                            }
                        }
                    }
                }

                // 2. 如果沒有找到標註，基於多段線的 Y 坐標推測樓層
                if (string.IsNullOrEmpty(floorName))
                {
                    Point3d minPoint = polyline.GeometricExtents.MinPoint;
                    Point3d maxPoint = polyline.GeometricExtents.MaxPoint;

                    // 使用 Y 坐標範圍推測樓層
                    double avgY = (minPoint.Y + maxPoint.Y) / 2;

                    // 假設某些 Y 坐標範圍對應特定的樓層
                    if (avgY >= 0 && avgY < 3000)
                    {
                        floorName = "1F";  // 假設 0 - 3000 單位範圍內是 1F
                    }
                    else if (avgY >= 3000 && avgY < 6000)
                    {
                        floorName = "2F";  // 假設 3000 - 6000 單位範圍內是 2F
                    }
                    // 可以根據實際需求添加更多的條件來判斷樓層
                }

                tr.Commit();
            }

            // 返回提取到的樓層名稱，找不到則返回空字符串
            return floorName;
        }

        private List<SortedData> SortColumnOrRow(List<SortedData> data)
        {
            data.Sort((a, b) =>
            {
                int result = a.FloorNumber.CompareTo(b.FloorNumber);
                if (result == 0)
                {
                    // 根據列和行進行排序
                    result = a.Points[0].X.CompareTo(b.Points[0].X);
                    if (result == 0)
                    {
                        result = a.Points[0].Y.CompareTo(b.Points[0].Y);
                    }
                }
                return result;
            });

            return data;
        }


        

        private ObjectId CreateText(string textString, Point3d position, double height, string layer)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 獲取模型空間
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // 創建文本
                DBText text = new DBText
                {
                    Position = position,
                    TextString = textString,
                    Height = height,
                    Layer = layer
                };

                // 添加文本到模型空間
                btr.AppendEntity(text);
                tr.AddNewlyCreatedDBObject(text, true);

                tr.Commit();
                return text.ObjectId;
            }
        }

        

    }
}
