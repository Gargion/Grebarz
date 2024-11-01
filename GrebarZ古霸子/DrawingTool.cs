using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Text;

namespace Grebarz古霸子
{
    public static partial class DrawingTool
    {
        /// <summary>
        /// 將圖形添加到圖行文件中
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="ent">圖形對象</param>
        /// <returns>圖形的ObjectId</returns>
        public static ObjectId AddEntityToModelSpace(this Database db, Entity ent)
        {
            //聲明ObjectId，用於返回
            ObjectId entId = ObjectId.Null;

            //開啟事務處理
            using(Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打開塊表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

                //打開塊表紀錄
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                //添加圖形到塊表紀錄
                entId = btr.AppendEntity(ent);

                //更新數據
                trans.AddNewlyCreatedDBObject(ent, true);

                //提交事務
                trans.Commit();
            }
            return entId;
        }
        /// <summary>
        /// 將圖形添加到圖行文件中
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="ent">圖形對象，可變參數</param>
        /// <returns>圖形的ObjectId，數組返回</returns>
        public static ObjectId[] AddEntityToModelSpace(this Database db, params Entity[] ent)
        {
            //聲明ObjectId，用於返回
            ObjectId[] entId = new ObjectId[ent.Length];

            //開啟事務處理
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打開塊表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

                //打開塊表紀錄
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                //添加圖形到塊表紀錄
                for(int i = 0; i < ent.Length; i++)
                {
                    //添加圖形到塊表紀錄
                    entId[i] = btr.AppendEntity(ent[i]);

                    //更新數據庫
                    trans.AddNewlyCreatedDBObject(ent[i], true);
                }
                //提交事務
                trans.Commit();
            }
            return entId;
        }
        /// <summary>
        /// 繪製直線
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="startPoint">起點座標</param>
        /// <param name="endPoint">終點座標</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddLine(this Database db, Point3d startPoint, Point3d endPoint)
        {
            return db.AddEntityToModelSpace(new Line(startPoint, endPoint));
        }
        /// <summary>
        /// 繪製直線
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="startPoint">起點座標</param>
        /// <param name="length">與X軸正方向的角度</param>
        /// <param name="degree"></param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddLine(this Database db, Point3d startPoint, double length, double degree)
        {
            double X = startPoint.X + length * Math.Cos(degree.DegreeToAngle());
            double Y = startPoint.Y + length * Math.Sin(degree.DegreeToAngle());

            Point3d endPoint = new Point3d(X,Y,0);

            return db.AddEntityToModelSpace(new Line(startPoint, endPoint));
        }
        /// <summary>
        /// 繪製弧-基本方法
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="center">弧心點</param>
        /// <param name="radius">弧半徑</param>
        /// <param name="startDegree">起始角度</param>
        /// <param name="endDegree">終止角度</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddArc(this Database db, Point3d center, double radius, double startDegree, double endDegree)
        {
            return db.AddEntityToModelSpace(new Arc(center, radius, startDegree.DegreeToAngle(), endDegree.DegreeToAngle()));
        }
        /// <summary>
        /// 繪製弧-指定三個點
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="startpoint">起點</param>
        /// <param name="pointOnArc">弧上任一點</param>
        /// <param name="endpoint">終點</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddArc(this Database db, Point3d startpoint, Point3d pointOnArc, Point3d endpoint)
        {
            //判斷三點是否同一線
            if (startpoint.IsOnSameLine(pointOnArc, endpoint))
            {
                return ObjectId.Null;
            }

            //創建幾何類對象
            CircularArc3d cArc = new CircularArc3d(startpoint, pointOnArc, endpoint);

            //要取得兩個角度，要用向量來求得 (角度為x軸與 起、終點 的向量夾角)
            double sAng = cArc.Center.GetAngelToXAxis(startpoint);
            double eAng = cArc.Center.GetAngelToXAxis(endpoint);
            return db.AddEntityToModelSpace(new Arc(cArc.Center, cArc.Radius, sAng, eAng));
        }
        /// <summary>
        /// 繪製弧-指定弧心.起點.夾角
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="center">弧心</param>
        /// <param name="startPoint">起點</param>
        /// <param name="degree">夾角</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddArc(this Database db,Point3d center,Point3d startPoint,double degree)
        {
            //獲取半徑
            double radius = center.GetDistanceBetweenTwoPoint(startPoint);

            //獲取起點角度
            double sAng = center.GetAngelToXAxis(startPoint);

            //創建圓弧對象
            Arc arc = new Arc(center, radius, sAng, sAng+degree);
            return db.AddEntityToModelSpace(arc);
        }
        /// <summary>
        /// 繪製圓-給圓心.半徑
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="center">圓心</param>
        /// <param name="radius">半徑</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddCircle(this Database db, Point3d center, double radius)
        {
            return db.AddEntityToModelSpace(new Circle(center, new Vector3d(0, 0, 1), radius));
        }

        /// <summary>
        /// 繪製圓-給兩個點
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="point1">第一點</param>
        /// <param name="point2">第二點</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddCircle(this Database db, Point3d point1, Point3d point2)
        {
            //獲取中心點
            Point3d center = point1.GetCenterPointBetweenTwoPoint(point2);
            //獲取半徑
            double radius = center.GetDistanceBetweenTwoPoint(point1);

            return db.AddEntityToModelSpace(new Circle(center,new Vector3d(0,0,1), radius));
        }
        /// <summary>
        /// 繪製圓-給三個點
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="point1">第一點</param>
        /// <param name="point2">第二點</param>
        /// <param name="point3">第三點</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddCircle(this Database db, Point3d point1, Point3d point2, Point3d point3)
        {
            //判斷三點不共線
            if (point1.IsOnSameLine(point2, point3))
            {
                return ObjectId.Null;
            }
            CircularArc3d cArc = new CircularArc3d(point1, point2, point3);
            return db.AddCircle(cArc.Center, cArc.Radius);
        }
        /// <summary>
        /// 繪製聚合線(多段直線)
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="isclosed">是否閉合</param>
        /// <param name="constantWidth">線寬</param>
        /// <param name="vertices">端點</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddPolyLine(this Database db, bool isclosed,double constantWidth, params Point2d[] vertices)
        {
            if(vertices.Length < 2)
            {
                return ObjectId.Null;
            }
            //創建一個聚合線對象
            Polyline pLine = new Polyline();

            //添加聚合線端點
            for(int i = 0; i < vertices.Length; i++)
            {
                pLine.AddVertexAt(i, vertices[i], 0, 0, 0);
            }

            //判斷是否閉合
            if (isclosed)
            {
                pLine.Closed = true;
            }

            //設置聚合線寬
            pLine.ConstantWidth = constantWidth;

            return db.AddEntityToModelSpace(pLine);
        }
        /// <summary>
        /// 繪製矩形(給對角兩點)
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="point1">第一點</param>
        /// <param name="point2">第二點</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddRectangle(this Database db, Point2d point1, Point2d point2)
        {
            //創建聚合線與點
            Polyline pLine = new Polyline();
            Point2d p1 = new Point2d(Math.Min(point1.X, point2.X), Math.Min(point1.Y, point2.Y));
            Point2d p2 = new Point2d(Math.Max(point1.X, point2.X), Math.Min(point1.Y, point2.Y));
            Point2d p3 = new Point2d(Math.Max(point1.X, point2.X), Math.Max(point1.Y, point2.Y));
            Point2d p4 = new Point2d(Math.Min(point1.X, point2.X), Math.Max(point1.Y, point2.Y));

            //定義節點
            pLine.AddVertexAt(0, p1, 0, 0, 0);
            pLine.AddVertexAt(1, p2, 0, 0, 0);
            pLine.AddVertexAt(2, p3, 0, 0, 0);
            pLine.AddVertexAt(3, p4, 0, 0, 0);

            //設置閉合
            pLine.Closed = true;

            return db.AddEntityToModelSpace(pLine);

        }
        /// <summary>
        /// 繪製正多邊形
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="center">多邊形圓心</param>
        /// <param name="radius">多邊形圓半徑</param>
        /// <param name="sideNum">邊數</param>
        /// <param name="startAngle">起始角度</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddPolygon(this Database db, Point2d center, double radius, int sideNum, double startAngle)
        {
            //創建聚合線
            Polyline pLine = new Polyline();

            //判斷邊數必須大於2
            if (sideNum < 3)
            {
                return ObjectId.Null;
            }
            Point2d[] point = new Point2d[sideNum];
            double angel = startAngle.DegreeToAngle();
            //求每個頂點的座標
            for (int i = 0; i < sideNum; i++)
            {
                point[i] =new Point2d( center.X + radius * Math.Cos(angel),center.Y + radius * Math.Sin(angel));
                pLine.AddVertexAt(i, point[i], 0, 0, 0);
                angel += Math.PI * 2 / sideNum;
            }
            pLine.Closed = true;

            return db.AddEntityToModelSpace(pLine);
        }
        /// <summary>
        /// 繪製橢圓
        /// </summary>
        /// <param name="db">圖形數據庫</param>
        /// <param name="center">橢圓中心</param>
        /// <param name="majorRadius">長軸長度</param>
        /// <param name="shortRadius">短軸長度</param>
        /// <param name="degree">長軸與x軸夾角</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddEllipse(this Database db,Point3d center,double majorRadius,double shortRadius,double degree)
        {
            //計算相關參數
            double ratio = shortRadius / majorRadius;
            Vector3d majorAxis = new Vector3d(majorRadius * Math.Cos(degree.DegreeToAngle()), majorRadius * Math.Sin(degree.DegreeToAngle()),0);
            //創建橢圓對象
            Ellipse elli = new Ellipse(center,Vector3d.ZAxis,majorAxis,ratio,0,Math.PI*2);
            return db.AddEntityToModelSpace(elli);

        }
       
        /// <summary>
        /// 圖案填充-自訂樣式
        /// </summary>
        /// <param name="cID">邊界圖形ObjectId</param>
        /// <param name="typeName">填充樣式名稱</param>
        /// <param name="scale">比例</param>
        /// <param name="degree">填充線角度</param>
        /// <param name="bkColor">背景色</param>
        /// <param name="hatchColorIndex">填充線顏色</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddCustomHatch(this ObjectId cID, string typeName,double scale,double degree,Color bkColor,int hatchColorIndex)
        {  
            if(cID == null)
            {
                return ObjectId.Null;
            }

            Database db = HostApplicationServices.WorkingDatabase;
            ObjectId hatchId = ObjectId.Null;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    //創建填充線對象
                    Hatch hatch = new Hatch();
                    //設置填充線比例
                    hatch.PatternScale = scale;
                    //設置背景色
                    hatch.BackgroundColor = bkColor;
                    //設置填充圖案顏色
                    hatch.ColorIndex = hatchColorIndex;
                    //設置填充線類型和樣式
                    hatch.SetHatchPattern(HatchPatternType.PreDefined, typeName);
                    //加入圖形數據庫
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    hatchId = btr.AppendEntity(hatch);
                    trans.AddNewlyCreatedDBObject(hatch, true);
                    //設置填充角度
                    hatch.PatternAngle = degree.DegreeToAngle();
                    //設置關聯
                    hatch.Associative = true;

                    //設置透明度，1~90，空=不透明
                    hatch.Transparency = new Transparency();

                    ObjectIdCollection obIDs = new ObjectIdCollection { cID };

                    //設置邊界圖形和填充方式
                    hatch.AppendLoop(HatchLoopTypes.Outermost, obIDs);
                    

                    //計算填充並顯示
                    hatch.EvaluateHatch(true);
                    //提交事務
                    trans.Commit();
                }
                catch(System.Exception ex)
                {
                    //處理異常
                    string errInf = "創建填充線時發生錯誤: " + ex.Message + " " + cID.ToString();
                    errInf.Gz_MsgBoxError();
                }
                
            }
            return hatchId;
        }
        /// <summary>
        /// 圖案填充-實心
        /// </summary>
        /// <param name="cID">邊界圖形ObjectId</param>
        /// <param name="bkColor">背景色</param>
        /// <param name="transparency">透明度</param>
        /// <returns>ObjectId</returns>
        public static ObjectId AddSolidHatch(this ObjectId cID, Color bkColor,int transparency)
        {
            if (cID == null)
            {
                return ObjectId.Null;
            }

            Database db = HostApplicationServices.WorkingDatabase;
            ObjectId hatchId = ObjectId.Null;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    //創建填充線對象
                    Hatch hatch = new Hatch();
                    //設置填充線比例
                    hatch.PatternScale = 1;
                    //設置背景色
                    hatch.BackgroundColor = bkColor;
                    
                    //設置填充線類型和樣式
                    hatch.SetHatchPattern(HatchPatternType.PreDefined, HatchTool.HatchPatternName.solid);
                    //加入圖形數據庫
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    hatchId = btr.AppendEntity(hatch);
                    trans.AddNewlyCreatedDBObject(hatch, true);
                    
                    //設置關聯
                    hatch.Associative = true;

                    //設置透明度，1~90，空=不透明
                    if(transparency == 0)
                    {
                        hatch.Transparency = new Transparency();
                    }
                    else if(transparency < 91 && transparency > 0)
                    {
                        hatch.Transparency = new Transparency((byte)transparency);
                    }
                    else
                    {
                        string errMsg = "填充線未能畫出，無效的透明度: " + transparency;
                        errMsg.Gz_MsgBoxError();
                        return ObjectId.Null;
                    }
                    ObjectIdCollection obIDs = new ObjectIdCollection { cID };

                    //設置邊界圖形和填充方式
                    hatch.AppendLoop(HatchLoopTypes.Outermost, obIDs);


                    //計算填充並顯示
                    hatch.EvaluateHatch(true);
                    //提交事務
                    trans.Commit();
                }
                catch (System.Exception ex)
                {
                    //處理異常
                    string errInf = "創建填充線時發生錯誤: " + ex.Message + " " + cID.ToString();
                    errInf.Gz_MsgBoxError();
                }

            }
            return hatchId;
        }

        /// <summary>
        /// 在 AutoCAD 中輸入命令或 AutoLISP 代碼。
        /// </summary>
        /// <param name="command">要輸入的命令或 AutoLISP 代碼。</param>
        public static void TypeInCommand(string command)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;

            // 在命令行中執行傳入的命令
            doc.SendStringToExecute(command + " ", true, false, false);
        }


    }
}
