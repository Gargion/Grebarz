using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Text;

namespace Grebarz古霸子
{
    public static partial class PropertyTool
    {
        //繪圖資料庫，重要，不可移除
        public static Database db = HostApplicationServices.WorkingDatabase;
        /// <summary>
        /// 改變圖型顏色
        /// </summary>
        /// <param name="c1Id">圖形的ObjectId</param>
        /// <param name="colorIndex">顏色索引</param>
        /// <returns></returns>
        private static void ChangeEntityColor(this ObjectId c1Id, short colorIndex)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打開塊表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

                //打開塊表紀錄
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                //添加圖形到塊表紀錄
                Entity ent1 = (Entity)c1Id.GetObject(OpenMode.ForWrite);
                ent1.ColorIndex = colorIndex;

                //提交事務
                trans.Commit();
            }
        }
        /// <summary>
        /// 改變圖型顏色
        /// </summary>
        /// <param name="ent">圖形的ObjectId</param>
        /// <param name="colorIndex">顏色索引</param>
        private static void ChangeEntityColor(this Entity ent, short colorIndex)
        {
            //判斷圖形的IsNewlyObject
            if (ent.IsNewObject)
            {
                ent.ColorIndex = colorIndex;
            }
            else
            {
                ent.ObjectId.ChangeEntityColor(colorIndex);
            }
        }
        /// <summary>
        /// 移動圖形
        /// </summary>
        /// <param name="entId">圖形的ObjectId</param>
        /// <param name="startPoint">參考基準點</param>
        /// <param name="targetPoint">參考目標點</param>
        private static void MoveEntity(this ObjectId entId, Point3d startPoint, Point3d targetPoint)
        {
            using(Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打開塊表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

                //打開塊表紀錄
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                //打開圖形
                Entity ent = (Entity)entId.GetObject(OpenMode.ForWrite);

                //創建新向量
                Vector3d vector = startPoint.GetVectorTo(targetPoint);
                Matrix3d mt = Matrix3d.Displacement(vector);
                ent.TransformBy(mt);

                //提交事務
                trans.Commit();
            }
        }
        /// <summary>
        /// 移動圖形
        /// </summary>
        /// <param name="ent">圖形對象</param>
        /// <param name="startPoint">參考基準點</param>
        /// <param name="targetPoint">參考目標點</param>
        private static void MoveEntity(this Entity ent, Point3d startPoint, Point3d targetPoint)
        {
            //判斷圖形對象IsNewlyObject
            if (ent.IsNewObject)
            {
                //計算變換矩陣
                Vector3d vector = startPoint.GetVectorTo(targetPoint);
                Matrix3d mt = Matrix3d.Displacement(vector);
                ent.TransformBy(mt);

            }
            else
            {
                ent.ObjectId.MoveEntity(startPoint, targetPoint);
            }
        }
        /// <summary>
        /// 複製圖形
        /// </summary>
        /// <param name="entId">圖形的ObjectId</param>
        /// <param name="startPoint">參考基準點</param>
        /// <param name="targetPoint">參考目標點</param>
        private static Entity CopyEntity(this ObjectId entId, Point3d startPoint, Point3d targetPoint)
        {
            //創建新的圖形對象
            Entity entR;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打開塊表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

                //打開塊表紀錄
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                //打開圖形
                Entity ent = (Entity)entId.GetObject(OpenMode.ForWrite);

                //創建新向量
                Vector3d vector = startPoint.GetVectorTo(targetPoint);
                Matrix3d mt = Matrix3d.Displacement(vector);
                entR = ent.GetTransformedCopy(mt);
                btr.AppendEntity(entR);
                trans.AddNewlyCreatedDBObject(entR, true);

                //提交事務
                trans.Commit();
            }
            return entR;
        }
        /// <summary>
        /// 複製圖形
        /// </summary>
        /// <param name="ent">圖形對象</param>
        /// <param name="startPoint">參考基準點</param>
        /// <param name="targetPoint">參考目標點</param>
        /// <returns></returns>
        private static Entity CopyEntity(this Entity ent, Point3d startPoint, Point3d targetPoint)
        {
            //創建新的圖形對象
            Entity entR;

            //判斷圖形對象IsNewlyObject
            if (ent.IsNewObject)
            {
                //計算變換矩陣
                Vector3d vector = startPoint.GetVectorTo(targetPoint);
                Matrix3d mt = Matrix3d.Displacement(vector);
                entR = ent.GetTransformedCopy(mt);

            }
            else
            {
                entR = ent.ObjectId.CopyEntity(startPoint, targetPoint);
            }
            return entR;
        }
        /// <summary>
        /// 旋轉圖形
        /// </summary>
        /// <param name="entId">圖形的ObjectId</param>
        /// <param name="center">旋轉基準點</param>
        /// <param name="degree">旋轉角度</param>
        private static void RotaeEntity(this ObjectId entId, Point3d center, double degree)
        {
            
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打開塊表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

                //打開塊表紀錄
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                //打開圖形
                Entity ent = (Entity)entId.GetObject(OpenMode.ForWrite);

                //創建新向量

                Matrix3d mt = Matrix3d.Rotation(degree.DegreeToAngle(), Vector3d.ZAxis, center);
                ent.GetTransformedCopy(mt);

                //提交事務
                trans.Commit();
            }
           
        }
        /// <summary>
        /// 旋轉圖形
        /// </summary>
        /// <param name="ent">圖形對象</param>
        /// <param name="center">旋轉基準點</param>
        /// <param name="degree">旋轉角度</param>
        private static void RotaeEntity(this Entity ent, Point3d center, double degree)
        {
            //判斷圖形對象IsNewlyObject
            if (ent.IsNewObject)
            {
                //計算變換矩陣
                
                Matrix3d mt = Matrix3d.Rotation(degree.DegreeToAngle(),Vector3d.ZAxis,center);
                ent.TransformBy(mt);

            }
            else
            {
                ent.ObjectId.RotaeEntity(center, degree);
            }
        }
        /// <summary>
        /// 鏡射圖形
        /// </summary>
        /// <param name="entId">圖形的ObjectId</param>
        /// <param name="point1">第一點</param>
        /// <param name="point2">第二點</param>
        /// <param name="isEraseSource">是否刪除原圖</param>
        /// <returns></returns>
        private static Entity MirrorEntity(this ObjectId entId, Point3d point1, Point3d point2, bool isEraseSource)
        {
            //創建一個物件對象
            Entity entR;
            //計算鏡像的參考矩陣
            Matrix3d mt = Matrix3d.Mirroring(new Line3d(point1, point2));
            using(Transaction trans = entId.Database.TransactionManager.StartTransaction())
            {
                //打開原圖象
                Entity ent = (Entity)trans.GetObject(entId, OpenMode.ForWrite);
                entR = ent.GetTransformedCopy(mt);
                //判斷是否刪除對象
                if (isEraseSource)
                {
                    ent.Erase();
                }
                trans.Commit();
            }
            return entR;
        }
        /// <summary>
        /// 鏡射圖形
        /// </summary>
        /// <param name="ent">圖形對象</param>
        /// <param name="point1">第一點</param>
        /// <param name="point2">第二點</param>
        /// <param name="isEraseSource">是否刪除原圖</param>
        private static void MirrorEntity(this Entity ent, Point3d point1, Point3d point2, bool isEraseSource)
        {
            //聲明對象
            Entity entR;
            if (ent.IsNewObject)
            {
                //計算鏡像的參考矩陣
                Matrix3d mt = Matrix3d.Mirroring(new Line3d(point1, point2));
                entR = ent.GetTransformedCopy(mt);

            }
            else
            {
                entR = ent.ObjectId.MirrorEntity(point1, point2, isEraseSource);
            }
        }
        /// <summary>
        /// 縮放圖形
        /// </summary>
        /// <param name="entId">圖形的ObjectId</param>
        /// <param name="basePoint">基準點</param>
        /// <param name="facter">縮放比例</param>
        private static void ScaleEntity(this ObjectId entId, Point3d basePoint, double facter)
        {
            Matrix3d mt = Matrix3d.Scaling(facter, basePoint);
            //啟動事務處理
            using(Transaction trans = entId.Database.TransactionManager.StartTransaction())
            {
                //打開要縮放的圖形對象
                Entity ent = (Entity)entId.GetObject(OpenMode.ForWrite);
                ent.TransformBy(mt);
                trans.Commit();
            }
        }
        /// <summary>
        /// 縮放圖形
        /// </summary>
        /// <param name="ent">圖形對象</param>
        /// <param name="basePoint">基準點</param>
        /// <param name="facter">縮放比例</param>
        private static void ScaleEntity(this Entity ent, Point3d basePoint, double facter)
        {
            if (ent.IsNewObject)
            {
                //計算縮放矩陣
                Matrix3d mt = Matrix3d.Scaling(facter, basePoint);
                ent.TransformBy(mt);
            }
            else
            {
                ent.ObjectId.ScaleEntity(basePoint, facter);
            }

        }
        /// <summary>
        /// 刪除物件
        /// </summary>
        /// <param name="entId">圖形的ObjectId</param>
        private static void EraseEntity(this ObjectId entId)
        {
            //啟動事務處理
            using (Transaction trans = entId.Database.TransactionManager.StartTransaction())
            {
                //打開要縮放的圖形對象
                Entity ent = (Entity)entId.GetObject(OpenMode.ForWrite);
                ent.Erase();
                trans.Commit();
            }
        }
    }
}
