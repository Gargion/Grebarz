using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using System.IO;
using System.Collections.Generic;
using Autodesk.AutoCAD.Colors;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Color = Autodesk.AutoCAD.Colors.Color;
using System.Windows;
using Autodesk.Windows;

namespace Grebarz古霸子
{
    public class LineExam
    {
        //繪圖資料庫，重要，不可移除
        public static Database db = HostApplicationServices.WorkingDatabase;
        public static Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        public static Editor ed = doc.Editor;

        [CommandMethod("gz001")]

        public void gz001()
        {
            //功能鎖
            if (!LicenseValidator.IsLicenseValid)
            {
                MessagerTool.CommandLocked();
                return;
            }

            //程序開始
            PromptPointOptions ppo = new PromptPointOptions("請指定第一個點:");
            PromptPointResult ppr = ppo.Gz_GetPoint();
            Point3d p1 = new Point3d(0, 0, 0);
            Point3d p2 = new Point3d(0,0,0);
            switch (ppr.Status)
            {
                case PromptStatus.OK:
                    p1 = ppr.Value;
                    ppo.BasePoint = p1;
                    ppo.UseBasePoint = true;
                    ppo.Message = "請指定第二個點";
                    ppr = ppo.Gz_GetPoint();
                    switch (ppr.Status)
                    {
                        case PromptStatus.OK:
                            p2 = ppr.Value;
                            db.AddLine(p1, p2);
                            break;
                    }


                    break;
                default:
                    MessagerTool.ForceEnd();
                    break;
            }
            
        }

        [CommandMethod("gz002")]

        public void gz002()
        {
            Arc arc1 = new Arc();
            arc1.Center = new Point3d(0, 0, 0);
            arc1.StartAngle = -Math.PI / 4;
            arc1.EndAngle = Math.PI / 4;
            arc1.Radius = 500;

            Arc arc2 = new Arc(new Point3d(50,50,0),20,45,90);
            Arc arc3 = new Arc(new Point3d(50, 50, 0), new Vector3d(0,0,1),20, Math.PI/4, Math.PI/2);

            db.AddEntityToModelSpace(arc1);
        }

        [CommandMethod("gz003")]

        public void gz003()
        {
            //透過三個點取得弧
            Point3d startpoint = new Point3d(100, 100, 0);
            Point3d endpoint = new Point3d(200, 500, 0);
            Point3d pointOnArc = new Point3d(150, 100, 0);

            db.AddArc(startpoint, pointOnArc, endpoint);
        }

        [CommandMethod("gz004")]

        public void gz004()
        {

            
            ObjectId cId = db.AddCircle(new Point3d(50, 50, 0), 200);
            ObjectId cId2 = db.AddCircle(new Point3d(500, 500, 0), 200);

            cId.AddCustomHatch(HatchTool.HatchPatternName.jislc20,10,0,Color.FromColorIndex(ColorMethod.ByLayer,4),1);
            cId2.AddSolidHatch(Color.FromColorIndex(ColorMethod.ByLayer, 1), 120);
        }
        
        [CommandMethod("gz005")]

        public void gz005()
        {
            db.AddPolyLine(false, 0,  new Point2d(50, 10), new Point2d(50, 0), new Point2d(160, 170));
        }

        [CommandMethod("gz006")]
        public void gz006()
        {


            db.AddRectangle(new Point2d(100, 100), new Point2d(500, 300));
        }

        [CommandMethod("gz007")]
        public void gz007()
        {
            db.AddPolygon(new Point2d(0, 0), 300, 7,45);
            db.AddPolygon(new Point2d(600, 0), 300, 7, 90);
            db.AddPolygon(new Point2d(900, 0), 300, 7, 0);
        }

        [CommandMethod("gz008")]

        public void gz008()
        {
            //定義文件名
            string fileName = @"C:\Users\aaaay\Desktop\CAD二次開發教學 C#NET\匯入測試檔案.txt";
            //把文件內容讀取到數組
            string[] contents = File.ReadAllLines(fileName);
            //創建List對象，將數據進行整理
            List<List<string>> List = new List<List<string>>();
            for (int i = 0; i < contents.Length; i++)
            {
                string[] cont = contents[i].Split(new char[] { ' ' });
                List<string> subList = new List<string>();
                for(int j = 0; j < cont.Length; j++)
                {
                    subList.Add(cont[j]);
                }
                List.Add(subList);
            }

            //創建聚合線對象
            Polyline pLine = new Polyline();
            for(int i = 0; i < List.Count; i++)
            {
                //數據轉換
                double X, Y;
                bool bx = double.TryParse(List[i][0], out X);
                bool by = double.TryParse(List[i][1], out Y);
                if(bx == false || bx == false)
                {
                    string text = "外部文件內容有誤！繪製圖型失敗。";
                    text.Gz_MsgBoxError();
                    return;
                }
                pLine.AddVertexAt(i, new Point2d(X,Y), 0, 0, 0);
            }
            db.AddEntityToModelSpace(pLine);
        }

        [CommandMethod("gz009")]
        public void gz009()
        {
            string str1 = "資訊提示";
            string str2 = "錯誤提示";
            string str3 = "警告提示";
            string str4 = "問題提示";
            str1.Gz_MsgBoxInf();
            str2.Gz_MsgBoxError();
            str3.Gz_MsgBoxWarning();
            str4.Gz_MsgBoxQuestion();
        }

        [CommandMethod("gz010")]
        public void gz010()
        {
            //RibbonControl ribbonControl = ComponentManager.Ribbon;
            //ribbonControl.AddTab("文字類", "Gz.RibbonId.01", true);
            //ribbonControl.AddTab("物件類", "Gz.RibbonId.02", true);
            //ribbonControl.AddTab("圖層類", "Gz.RibbonId.03", true);
        }
    }
}
