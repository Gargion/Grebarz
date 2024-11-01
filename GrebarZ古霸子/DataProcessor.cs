using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Grebarz古霸子
{
    public static class DataProcessor
    {
        public static List<object> LenJsCol2800 = new List<object>();
        public static List<object> LenJsCol4200 = new List<object>();
        public static List<object> LenJsCol2800Up = new List<object>();
        public static List<object> LenJsCol4200Up = new List<object>();
        public static List<object> LenUnit = new List<object>();
        public static List<object> Len904200 = new List<object>();

        public static void ReadBaseDatCol(object[,] data) // 處理 Excel 數據進行搭接計算。
        {
            // 輸出指令說明
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 處理 Excel 數據進行搭接計算。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 初始化數據列表
            LenJsCol2800.Clear();
            LenJsCol4200.Clear();
            LenJsCol2800Up.Clear();
            LenJsCol4200Up.Clear();
            LenUnit.Clear();
            Len904200.Clear();

            int K = 0;

            // 跳過前兩行數據，開始解析
            for (int i = 2; i < data.GetLength(0); i++)
            {
                object[] row = Enumerable.Range(0, data.GetLength(1)).Select(j => data[i, j]).ToArray();

                if (row.Length >= 5)
                {
                    // 根據 NTH 提取數據並模仿 LISP 中的 CONS 函數
                    var dataJs = new List<object> { new List<object> { 210, Convert.ToInt32(row[0]) }, Convert.ToInt32(row[4]) };
                    var dataJsUp = new List<object> { new List<object> { 210, Convert.ToInt32(row[0]) }, Convert.ToInt32(row[1]) };

                    // 根據 Excel 數據行處理
                    LenJsCol2800.Add(dataJs);
                    LenJsCol2800Up.Add(dataJsUp);
                }

                if (row.Length >= 7)
                {
                    var data90 = new List<object> { Convert.ToInt32(row[0]), Convert.ToInt32(row[6]) };
                    var dataUnit = new List<object> { Convert.ToInt32(row[0]), Convert.ToDouble(row[7]) };

                    Len904200.Add(data90);
                    LenUnit.Add(dataUnit);
                }

                // 繼續根據 LISP 的順序處理更多數據
            }

            // 進行類似 LISP 的反轉數據操作
            LenJsCol2800.Reverse();
            LenJsCol4200.Reverse();
            LenJsCol2800Up.Reverse();
            LenJsCol4200Up.Reverse();
            Len904200.Reverse();
        }
    }
}
