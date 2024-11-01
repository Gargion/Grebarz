using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Grebarz古霸子
{
    public class DrawingProcessor
    {
        public void ProcessDrawing()
        {
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 繪製根據 Excel 資料生成的圖形。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 初始化AutoCAD對象
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Excel 文件的讀取操作
            string excelFilePath = (string)Application.GetSystemVariable("dwgprefix") + "7CD.XLS";
            var dataFromExcel = ReadExcelData(excelFilePath, "樓層");
            var joinDataFromExcel = ReadExcelData(excelFilePath, "搭接");

            // 繪製圖形邏輯
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 設置圖層
                SetLayer("REC_COL", 7);
                SetLayer("COLDAT", 7);
                SetLayer("HOOP", 5);
                SetLayer("HOOPTXT", 6);

                // 開始繪製文字、圖形、尺寸等
                foreach (var dataRow in dataFromExcel)
                {
                    Point3d startPoint = new Point3d(0, 0, 0); // 根據需要初始化起始點
                    // 繪製直線
                    DrawLine(db, tr, startPoint, new Point3d(100, 100, 0));

                    // 繪製文字
                    DrawText(db, tr, "標註文字", new Point3d(50, 50, 0), 10, "標註");
                }

                // 提交事務
                tr.Commit();
            }
        }

        private void SetLayer(string layerName, short colorIndex)
        {
            // 設置 AutoCAD 圖層的邏輯
        }

        private void DrawLine(Database db, Transaction tr, Point3d startPt, Point3d endPt)
        {
            Line line = new Line(startPt, endPt);
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            btr.AppendEntity(line);
            tr.AddNewlyCreatedDBObject(line, true);
        }

        private void DrawText(Database db, Transaction tr, string textContent, Point3d position, double height, string layer)
        {
            DBText text = new DBText
            {
                TextString = textContent,
                Height = height,
                Position = position,
                Layer = layer
            };
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            btr.AppendEntity(text);
            tr.AddNewlyCreatedDBObject(text, true);
        }

        private List<object[]> ReadExcelData(string filePath, string sheetName)
        {
            // Excel 文件讀取邏輯
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[sheetName];

            List<object[]> data = new List<object[]>();
            Excel.Range usedRange = worksheet.UsedRange;

            for (int row = 1; row <= usedRange.Rows.Count; row++)
            {
                object[] rowData = new object[usedRange.Columns.Count];
                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    rowData[col - 1] = (usedRange.Cells[row, col] as Excel.Range).Value;
                }
                data.Add(rowData);
            }

            workbook.Close();
            excelApp.Quit();

            return data;
        }
    }
}
