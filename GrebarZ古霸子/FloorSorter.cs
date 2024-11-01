using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Grebarz古霸子
{
    public class FloorSorter
    {
        public List<string> gzFL_Sort(List<List<string>> floorData)
        {
            string tip = "\n--------------------" + Environment.NewLine +
                         "\n指令說明: 依樓層分組進行圖層排序。" + Environment.NewLine +
                         "\n--------------------";
            tip.Gz_StatusBarMsg();

            // 初始化樓層列表
            List<string> flSort = new List<string>();
            List<string> flB = new List<string>();    // 地下層 (B開頭)
            List<string> flS = new List<string>();    // 樓層 (結尾為F)
            List<string> flR = new List<string>();    // 屋頂層 (R或P開頭)

            // 依據輸入數據篩選出唯一的樓層名稱
            foreach (var floor in floorData)
            {
                string floorName = floor[1]; // 獲取第二個元素作為樓層名稱
                if (!flSort.Contains(floorName))
                {
                    flSort.Add(floorName);
                }
            }

            // 分類樓層，B*為地下層，*F且非B和R開頭為普通樓層，R*, P*為屋頂層
            foreach (var floor in flSort)
            {
                if (floor.StartsWith("B"))
                {
                    flB.Add(floor);
                }
                else if (floor.EndsWith("F") && !floor.StartsWith("B") && !floor.StartsWith("R"))
                {
                    flS.Add(floor.Replace("F", ""));  // 移除F以便排序
                }
                else if (floor.StartsWith("R") || floor.StartsWith("P"))
                {
                    flR.Add(floor);
                }
            }

            // 依數字順序排序普通樓層
            flS = flS.OrderBy(f => double.Parse(f)).Select(f => f + "F").Reverse().ToList(); // 按大小排序後倒序並加回F

            // 合併地下層、普通樓層和屋頂層
            List<string> sortedFloors = new List<string>();
            sortedFloors.AddRange(flB);
            sortedFloors.AddRange(flS);
            sortedFloors.AddRange(flR);

            return sortedFloors;
        }

    }
}
