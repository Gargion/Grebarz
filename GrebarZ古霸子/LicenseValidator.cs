using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using Exception = System.Exception;

namespace Grebarz古霸子
{
    public class LicenseValidator
    {
        public static bool IsLicenseValid { get; private set; } // 靜態變量

        // 初始化方法
        public async void ValidateLicense()
        {
            // 確保 WPF 應用程序啟動
            System.Windows.Application app = System.Windows.Application.Current;
            if (app == null)
            {
                app = new System.Windows.Application();

                app.ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown; // 確保不會因窗口關閉而結束
                
            }

            string licenseKey = LicenseHelper.GetLicenseKey();
            bool licenseKeyCheck = await LicenseKeyCheck(licenseKey);

            // 如果許可無效，持續要求使用者輸入
            while (!licenseKeyCheck)
            {
                // 使用 Dispatcher 呼叫 WPF 窗口
                bool shouldExit = await app.Dispatcher.InvokeAsync(() =>
                {
                    WPF_LicenseInput window = new WPF_LicenseInput();
                    if (window.ShowDialog() == true)
                    {
                        // 保存輸入的許可金鑰
                        LicenseHelper.SaveLicenseKey(window.EnteredLicenseKey);
                        MessagerTool.LicenseSaved();
                        licenseKey = LicenseHelper.GetLicenseKey();
                        return false; // 不退出循環，繼續驗證
                    }
                    else
                    {
                        return true; // 用戶取消，退出循環
                    }
                });

                if (shouldExit)
                {
                    break;
                }
                licenseKeyCheck = await LicenseKeyCheck(licenseKey);
            }

            // 設置狀態欄訊息
            StatusBarTool.Gz_StatusBarMsg("\nDLL 已載入: 初始化完成....\n");
        }



        // 檢查許可金鑰的方法
        public static async Task<bool> LicenseKeyCheck(string licenseKey)
        {
            WooCommerceLicenseManager licenseManager = new WooCommerceLicenseManager();
            bool isLicenseValid = await licenseManager.ValidateLicenseKeyAsync(licenseKey);

            IsLicenseValid = isLicenseValid; // 更新靜態變量

            if (isLicenseValid)
            {
                StatusBarTool.Gz_StatusBarMsg(MessagerTool.LICENSE_SUCCESS);
                return true;
            }
            else
            {
                MessagerTool.LicenseFaild();
                return false;
            }
        }

        public static class LicenseHelper
        {
            private static readonly string LicenseFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Grebarz", "License.txt");

            public static string GetLicenseKey()
            {
                try
                {
                    if (File.Exists(LicenseFilePath))
                    {
                        string str = "您的授權金鑰: " + File.ReadAllText(LicenseFilePath).Trim();
                        str.Gz_MsgBoxInf();
                        return File.ReadAllText(LicenseFilePath).Trim();
                    }
                    else
                    {
                        string str = LicenseFilePath + " 路徑中沒找到檔案";
                        //str.Gz_MsgBoxInf();
                        return "0";
                    }
                }
                catch (Exception ex)
                {
                    String str = "Error reading license key: " + ex.Message;
                    str.Gz_MsgBoxError();
                }

                return string.Empty;
            }

            public static void SaveLicenseKey(string licenseKey)
            {
                try
                {
                    // 確保目錄存在
                    Directory.CreateDirectory(Path.GetDirectoryName(LicenseFilePath));
                    File.WriteAllText(LicenseFilePath, licenseKey);

                    string str = "授權金鑰已儲存。";
                    str.Gz_MsgBoxInf();

                }
                catch (Exception ex)
                {
                    String str = "Error saving license key: " + ex.Message;
                    str.Gz_MsgBoxError();
                }
            }
        }

    }
}
