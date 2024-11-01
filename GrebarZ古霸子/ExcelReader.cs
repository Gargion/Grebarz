using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Application = Microsoft.Office.Interop.Excel.Application;
using Exception = System.Exception;

namespace Grebarz古霸子
{
    public static class ExcelReader 
    {
        public static string KeyReadJs { get; private set; }
        public static object[,] ReadFlC() //從 Excel 文件(7CD.XLS)中讀取[樓層]數據。
        {
            // 獲取當前圖形文件的路徑
            string dwgPrefix = HostApplicationServices.WorkingDatabase.Filename;

            // Excel 文件名稱
            string fileName = Path.Combine(Path.GetDirectoryName(dwgPrefix), "7CD.XLS");

            // 檢查 Excel 文件是否存在
            if (!File.Exists(fileName))
            {
                string errorMsg = $"未找到 Excel 文件: {fileName}";
                errorMsg.Gz_MsgBoxError();
                return null;
            }

            try
            {
                // 初始化 Excel 應用程序
                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(fileName);
                Worksheet worksheet = workbook.Sheets["樓層"] as Worksheet;

                // 假設我們要讀取工作表的所有數據
                Range usedRange = worksheet.UsedRange;
                object[,] data = usedRange.Value2;

                // 關閉 Excel
                workbook.Close(false);
                excelApp.Quit();

                return data; // 返回 Excel 數據
            }
            catch (Exception ex)
            {
                string errorMsg = "讀取 Excel 文件時發生錯誤: " + ex.Message;
                errorMsg.Gz_MsgBoxError();
                return null;
            }
        }

        public static void ReadJsC()//從 Excel 文件(7CD.XLS)中讀取[搭接]數據。
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 從 Excel 文件(7CD.XLS)中讀取[搭接]數據。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 獲取當前圖形文件的路徑
            string dwgPrefix = HostApplicationServices.WorkingDatabase.Filename;

            // 組合 Excel 文件路徑
            string fileName = Path.Combine(Path.GetDirectoryName(dwgPrefix), "7CD.XLS");

            if (!File.Exists(fileName))
            {
                string errorMsg = $"未找到 Excel 文件: {fileName}";
                errorMsg.Gz_MsgBoxError();
                return;
            }

            try
            {
                // 打開 Excel 文件
                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(fileName);
                Worksheet worksheet = workbook.Sheets["搭接"] as Worksheet;

                // 假設你需要從工作表中讀取數據
                Range usedRange = worksheet.UsedRange;
                object[,] data = usedRange.Value2;

                // 調用自定義方法處理數據
                ReadBaseDatCol(data);

                // 設置 KeyReadJs 為 "1"
                KeyReadJs = "1";

                // 關閉 Excel
                workbook.Close(false);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                string errorMsg = "讀取 Excel 文件時發生錯誤: " + ex.Message;
                errorMsg.Gz_MsgBoxError();
            }
        }



        // 這是處理 Excel 數據的自定義方法
        private static void ReadBaseDatCol(object[,] data)
        {
            // 在這裡處理 data，這與 LISP 中的 (READ_BASEDAT_COL LIST_DATA) 等效
            // 您可以根據需求來處理數據，例如輸出數據、篩選等。
            for (int row = 1; row <= data.GetLength(0); row++)
            {
                for (int col = 1; col <= data.GetLength(1); col++)
                {
                    Console.WriteLine($"Data at [{row}, {col}]: {data[row, col]}");
                }
            }
        }
    }


}
