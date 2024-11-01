using System;
using System.Collections.Generic;
using System.Text;

namespace Grebarz古霸子
{
    public static partial class MessagerTool
    {
        private static readonly string FORCE_END = "--------------------------" + Environment.NewLine + "程序被強制結束。";
        private static readonly string PROGRESS_SUCCESS = "--------------------------" + Environment.NewLine + "程序順利完成！";
        public static readonly string LICENSE_SUCCESS = "--------------------------" + Environment.NewLine + "License Key 通過驗證，歡迎使用 " + Core.Gz_GetSystemName() + " |版本: " + Core.Gz_GetSystemVersion();
        private static readonly string LICENSE_FAILD = "糟糕...您的 License Key 已失效，請至官方網站上進行購買。" + Environment.NewLine + Core.Gz_GetProductSite();
        private static readonly string LICENSE_SAVED = "--------------------------" + Environment.NewLine + "License Key 已保存 ";
        private static readonly string COMMAND_LOCKED = "--------------------------" + Environment.NewLine + "License Key 未通過驗證，功能無法使用。 ";

        public static void ForceEnd()
        {
            StatusBarTool.Gz_StatusBarMsg(FORCE_END);
        }

        public static void ProgessSuccess()
        {
            StatusBarTool.Gz_StatusBarMsg(PROGRESS_SUCCESS);
        }

        public static void LicenseSuccess()
        {
            StatusBarTool.Gz_StatusBarMsg(LICENSE_SUCCESS);
        }

        public static void LicenseFaild()
        {
            Gz_MsgBoxError(LICENSE_FAILD);
        }

        public static void LicenseSaved()
        {
            Gz_MsgBoxInf(LICENSE_SAVED);
        }

        public static void CommandLocked()
        {
            StatusBarTool.Gz_StatusBarMsg(COMMAND_LOCKED);
        }

        /// <summary>
        /// 視窗提示文字
        /// </summary>
        /// <param name="text">輸入信息</param>
        public static void Gz_MsgBoxInf(this string text)
        {
            string caption = Core.Gz_GetSystemName() + "==提示==";

            //Application.ShowAlertDialog(str);
            System.Windows.Forms.MessageBox.Show(text, caption, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }


        public static void Gz_MsgBoxWarning(this string text)
        {
            string caption = Core.Gz_GetSystemName() + "==警告==";

            //Application.ShowAlertDialog(str);
            System.Windows.Forms.MessageBox.Show(text, caption, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
        }
        public static void Gz_MsgBoxError(this string text)
        {
            string caption = Core.Gz_GetSystemName() + "==錯誤==";

            //Application.ShowAlertDialog(str);
            System.Windows.Forms.MessageBox.Show(text, caption, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
        }
        public static void Gz_MsgBoxQuestion(this string text)
        {
            string caption = Core.Gz_GetSystemName() + "==問題==";

            //Application.ShowAlertDialog(str);
            System.Windows.Forms.MessageBox.Show(text, caption, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
        }
    }
}
