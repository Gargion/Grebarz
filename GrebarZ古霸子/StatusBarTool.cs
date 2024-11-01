using System;
using System.Collections.Generic;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace Grebarz古霸子
{
    public static partial class StatusBarTool
    {
        /// <summary>
        /// 狀態欄提示文字
        /// </summary>
        /// <param name="text">輸入信息</param>
        public static void Gz_StatusBarMsg(this string text)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage(text);
        }
        public static void Gz_StatusBarExplainMsg(this string text)
        {
            string tip = "\n--------------------" + Environment.NewLine +
                     "\n指令說明: " + text + Environment.NewLine +
                     "\n--------------------";

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage(tip);
        }

    }
}