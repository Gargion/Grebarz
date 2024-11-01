using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Grebarz古霸子
{
    class WPF
    {
        /// <summary>
        /// 顯示指定類型的 WPF 窗口，確保 WPF Application 上下文存在。
        /// </summary>
        /// <typeparam name="T">要顯示的 WPF 窗口類型</typeparam>
        /// <param name="constructorArgs">傳遞給窗口構造函數的參數</param>
        public static void ShowWPF<T>(params object[] constructorArgs) where T : Window
        {
            // 確保有一個 WPF Application 上下文
            if (Application.Current == null)
            {
                Application app = new Application
                {
                    ShutdownMode = ShutdownMode.OnExplicitShutdown
                };

                // 使用 Dispatcher 確保 UI 線程能夠運行
                app.Dispatcher.Invoke(() =>
                {
                    // 使用反射創建窗口實例
                    T window = (T)Activator.CreateInstance(typeof(T), constructorArgs);
                    window.ShowDialog();
                });
            }
            else
            {
                // 使用現有的 Application 來顯示 WPF 窗口
                Application.Current.Dispatcher.Invoke(() =>
                {
                    T window = (T)Activator.CreateInstance(typeof(T), constructorArgs);
                    window.ShowDialog();
                });
            }
        }
    }   
}
