using Autodesk.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using System.Windows;
using Application = System.Windows.Application;

namespace Grebarz古霸子
{
    public class App_Initializer : IExtensionApplication
    {
        // 許可驗證器實例
        private static LicenseValidator licenseValidator;

        public void Initialize()
        {
            // 確保 WPF 應用程式啟動
            if (Application.Current == null)
            {
                Application app = new Application();
                app.ShutdownMode = ShutdownMode.OnExplicitShutdown;
            }

            // 啟動許可驗證
            licenseValidator = new LicenseValidator();
            licenseValidator.ValidateLicense();


            // 啟動時關閉預設 Ribbon
            DrawingTool.TypeInCommand("RIBBONCLOSE ");

            // 加載自定義 Ribbon
            //AddCustomRibbon();

            // 其他初始化代碼可以放在這裡
        }

        public void Terminate()
        {
            // 清理代碼或處理終止事件
        }
    }
}