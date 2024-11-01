using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Grebarz古霸子
{
    public class ConnectorData
    {
        // 位置屬性
        public double PositionX { get; set; }   // 柱 X 向軸號
        public double PositionY { get; set; }   // 柱 Y 向軸號

        // 尺寸屬性
        public double Width { get; set; }       // 柱子的寬度
        public double Height { get; set; }      // 柱子的高度

        // 主筋和箍筋屬性
        public int MainBarCount { get; set; }   // 主筋數量
        public string MainBarSize { get; set; } // 主筋尺寸
        public int StirrupCount { get; set; }   // 箍筋數量
        public string StirrupSize { get; set; } // 箍筋尺寸

        // 續接屬性
        public string ConnectorType { get; set; }  // 續接類型，例如 "搭接"
        public bool IsConnectedAbove { get; set; } // 是否有上部續接
        public bool IsConnectedBelow { get; set; } // 是否有下部續接

        // 構造函數
        public ConnectorData()
        {
            // 默認值，可以根據實際情況設置
            PositionX = 0;
            PositionY = 0;
            Width = 500;   // 默認柱子寬度
            Height = 3000; // 默認柱子高度
            MainBarCount = 4;
            MainBarSize = "#8";
            StirrupCount = 6;
            StirrupSize = "#4";
            ConnectorType = "續接";
            IsConnectedAbove = false;
            IsConnectedBelow = false;
        }
    }
}
