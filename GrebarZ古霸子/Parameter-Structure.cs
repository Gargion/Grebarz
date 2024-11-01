using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Grebarz古霸子
{
    public class Column
    {
        // 柱子基本屬性
        public string ColumnName { get; set; }      // 柱的名稱或編號
        public string FloorLevel { get; set; }      // 所在樓層名稱
        public double PositionX { get; set; }       // 柱子的X方向位置
        public double PositionY { get; set; }       // 柱子的Y方向位置
        public double SizeX { get; set; }           // 柱子的X方向尺寸（寬度）
        public double SizeY { get; set; }           // 柱子的Y方向尺寸（深度）
        public double ColumnHeight { get; set; }    // 柱子的總高度

        // 樓層相關信息
        public double FloorHeight { get; set; }     // 樓層高度
        public int FloorNumber { get; set; }        // 樓層編號（用於排序和處理）

        // 主筋和箍筋信息
        public int MainBarCount { get; set; }       // 主筋數量
        public string MainBarSpec { get; set; }     // 主筋尺寸（例如 "#8"）
        public string StirrupSpec { get; set; }     // 箍筋尺寸
        public int StirrupCount { get; set; }       // 箍筋數量

        // 續接相關
        public bool IsConnectedAbove { get; set; }  // 是否有上部續接
        public bool IsConnectedBelow { get; set; }  // 是否有下部續接

        // 繪圖位置屬性
        public Point3d Position => new Point3d(PositionX, PositionY, 0);  // 繪製時的位置

        // 標註信息
        public string AnnotationText { get; set; }  // 標註文本（例如：樓層名稱、主筋信息等）
    }

}
