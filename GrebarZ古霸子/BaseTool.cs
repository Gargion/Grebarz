using System;
using System.Collections.Generic;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace Grebarz古霸子
{
    public static partial class BaseTool
    {
        

        /// <summary>
        /// 角度轉換為弧度
        /// </summary>
        /// <param name="degree">角度值</param>
        /// <returns>弧度</returns>
        public static double DegreeToAngle(this double degree)
        {
            return degree * Math.PI / 180;
        }
        /// <summary>
        /// 弧度轉換為角度
        /// </summary>
        /// <param name="angle">弧度值</param>
        /// <returns>角度</returns>
        public static double AngleToDegree(this double angle)
        {
            return angle * 180 / Math.PI;
        }
        /// <summary>
        /// 判斷是否三點成一線
        /// </summary>
        /// <param name="firstPoint">第一點</param>
        /// <param name="secondPoint">第二點</param>
        /// <param name="thirdPoint">第三點</param>
        /// <returns>是.否</returns>
        public static bool IsOnSameLine(this Point3d firstPoint, Point3d secondPoint, Point3d thirdPoint)
        {
            Vector3d vSecToFst = secondPoint.GetVectorTo(firstPoint);
            Vector3d vSecToTrd = secondPoint.GetVectorTo(thirdPoint);
            if (vSecToFst.GetAngleTo(vSecToTrd) == 0 || vSecToFst.GetAngleTo(vSecToTrd) == Math.PI)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// 獲得自己至目標點連線與x軸的夾角
        /// </summary>
        /// <param name="starPoint">起點</param>
        /// <param name="endPoint">目標點</param>
        /// <returns>double角度</returns>
        public static double GetAngelToXAxis(this Point3d starPoint, Point3d endPoint)
        {
            //創建x軸向量
            Vector3d vXAxis = new Vector3d(1, 0, 0);

            //獲取起終兩點連線的向量
            Vector3d vSptToEpt = starPoint.GetVectorTo(endPoint);
            return vSptToEpt.Y > 0 ? vXAxis.GetAngleTo(vSptToEpt) : -vXAxis.GetAngleTo(vSptToEpt);
        }
        /// <summary>
        /// 給兩點，求距離
        /// </summary>
        /// <param name="point1">第一點</param>
        /// <param name="point2">第二點</param>
        /// <returns>double數值</returns>
        public static double GetDistanceBetweenTwoPoint(this Point3d point1,Point3d point2)
        {
            return Math.Sqrt((point1.X - point2.X) * (point1.X - point2.X) + (point1.Y - point2.Y) * (point1.Y - point2.Y) + (point1.Z - point2.Z) * (point1.Z - point2.Z));
        }
        /// <summary>
        /// 給兩點，求中點座標
        /// </summary>
        /// <param name="point1">第一點</param>
        /// <param name="point2">第二點</param>
        /// <returns>點座標</returns>
        public static Point3d GetCenterPointBetweenTwoPoint(this Point3d point1, Point3d point2)
        {
            return new Point3d((point1.X + point2.X) / 2, (point1.Y + point2.Y) / 2, (point1.Z + point2.Z) / 2);
        }
        


    }
}
