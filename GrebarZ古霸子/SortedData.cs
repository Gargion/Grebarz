using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Grebarz古霸子
{
    public class SortedData
    {
        public string FloorNumber { get; set; }         // 樓層名稱
        public double FloorHeight { get; set; }         // 樓層高度
        public string ConcreteStrength { get; set; }    // 混凝土強度
        public List<Point3d> Points { get; set; }       // 邊界點（通常是 PT1 和 PT2）

        // 構造函數
        public SortedData()
        {
            Points = new List<Point3d>();  // 初始化邊界點列表
        }

        // 方法：方便打印和顯示 SortedData
        public override string ToString()
        {
            return $"樓別: {FloorNumber}, 樓高: {FloorHeight}, fc': {ConcreteStrength}, Points: {Points.Count} points";
        }
    }

}
