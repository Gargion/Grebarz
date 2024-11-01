using Autodesk.Windows;
using Autodesk.AutoCAD.Runtime;
using System.Windows.Input;
using System;
using System.Windows.Media.Imaging;
using System.Windows.Interop;
using System.Windows.Media;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;

namespace Grebarz古霸子
{
    public class RibbonTool
    {
        [CommandMethod("AddRibbon")]
        public void CreateRibbon()
        {
            //強制打開
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("RIBBON ", true, false, false);

            // 創建標籤
            RibbonTab tab1 = AddTab("Grebarz'古霸子'系統", "MY_RIBBON_TAB");

            // 創建面板
            RibbonPanelSource panelSource = AddPanel(tab1, "常用工具");

            // 創建按鈕並綁定到自定義命令
            AddButton(panelSource, "執行命令", "MyCustomCommand", null);

            // 創建下拉選單並綁定到自定義命令
            RibbonPanelSource panel1 = AddPanel(tab1, "選擇工具");

            List<(string, string, Bitmap)> options = new List<(string, string, Bitmap)>
                {
                    ("選項1", "CommandForOption1", null),  // 指定圖像和綁定命令
                    ("選項2", "CommandForOption2", null),
                    ("選項3", "CommandForOption3", null)  // 沒有指定圖像，使用預設圖像
                };

            AddCombo(panel1, "選擇一項", options);

            // 創建分割按鈕並綁定到自定義命令
            RibbonPanelSource panel2 = AddPanel(tab1, "動作工具");

            Dictionary<string, (string, Bitmap)> splitCommands = new Dictionary<string, (string, Bitmap)>
            {
                { "動作1", ("ActionCommand1", null) },
                { "動作2", ("ActionCommand2", null) }
            };

            AddSplitButton(panel2, "選擇動作", splitCommands);

            //創建切換按鈕並綁定到自定義命令
            RibbonPanelSource panel3 = AddPanel(tab1, "狀態工具");

            AddToggleButton(panel3, "開啟/關閉功能", "EnableCommand", "DisableCommand");


        }

        // 創建標籤的方法
        public RibbonTab AddTab(string tabTitle, string tabId)
        {
            var ribbonControl = ComponentManager.Ribbon;

            if (ribbonControl != null)
            {
                // 檢查是否已經存在該標籤
                foreach (RibbonTab existingTab in ribbonControl.Tabs)
                {
                    if (existingTab.Title == tabTitle)
                    {
                        return existingTab; // 如果已經存在，則直接返回
                    }
                }

                // 創建新的 Ribbon 標籤
                RibbonTab tab = new RibbonTab
                {
                    Title = tabTitle,
                    Id = tabId
                };
                ribbonControl.Tabs.Add(tab);
                return tab;
            }

            return null;
        }

        // 創建面板的方法
        public RibbonPanelSource AddPanel(RibbonTab tab, string panelTitle)
        {
            // 創建新的 Ribbon 面板
            RibbonPanelSource panelSource = new RibbonPanelSource
            {
                Title = panelTitle
            };
            RibbonPanel panel = new RibbonPanel
            {
                Source = panelSource
            };
            tab.Panels.Add(panel);
            return panelSource;
        }

        // 創建按鈕並綁定命令的方法
        public void AddButton(RibbonPanelSource panelSource, string buttonText, string commandName, Bitmap icon = null)
        {
            RibbonButton button = new RibbonButton
            {
                Text = buttonText,
                ShowText = true
            };

            // 設置圖像，如果沒有指定，使用預設圖片
            button.LargeImage = ConvertBitmapToImageSource(icon ?? Properties.Resources.Default_icon);

            button.CommandHandler = new RibbonCommandHandler(commandName);

            // 將按鈕添加到面板
            panelSource.Items.Add(button);
        }

        // 創建下拉選單的方法並綁定命令
        public void AddCombo(RibbonPanelSource panelSource, string comboName, List<(string Option, string Command, Bitmap Icon)> options)
        {
            RibbonCombo combo = new RibbonCombo
            {
                Text = comboName + Environment.NewLine,
                ShowText = true,
                Width = 150
            };

            foreach (var option in options)
            {
                RibbonItem comboItem = new RibbonItem
                {
                    Text = option.Option,
                    IsEnabled = true,
                    Image = ConvertBitmapToImageSource(option.Icon ?? Properties.Resources.Default_icon) // 使用指定或預設圖片
                };

                combo.Items.Add(comboItem);
            }

            // 設定事件來處理選擇改變
            combo.CurrentChanged += (s, e) =>
            {
                var selectedItem = combo.Current as RibbonItem;
                if (selectedItem != null)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n已選擇：{selectedItem.Text}");
                    // 執行綁定的命令
                    var command = options.FirstOrDefault(opt => opt.Option == selectedItem.Text).Command;
                    if (!string.IsNullOrEmpty(command))
                    {
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute(command + " ", true, false, false);
                    }
                }
            };

            panelSource.Items.Add(combo);
        }

        // 創建切換按鈕並綁定兩個狀態下的命令
        public void AddToggleButton(RibbonPanelSource panelSource, string buttonText, string commandOn, string commandOff, Bitmap icon = null)
        {
            RibbonToggleButton toggleButton = new RibbonToggleButton
            {
                Text = buttonText,
                ShowText = true,
                IsChecked = false, // 初始狀態
                LargeImage = ConvertBitmapToImageSource(icon ?? Properties.Resources.Default_icon) // 設置圖像
            };

            // 設定按鈕點擊事件
            toggleButton.CommandHandler = new ToggleCommandHandler(commandOn, commandOff, toggleButton);

            panelSource.Items.Add(toggleButton);
        }

        // 創建分割按鈕並綁定命令
        public void AddSplitButton(RibbonPanelSource panelSource, string buttonText, Dictionary<string, (string Command, Bitmap Icon)> commands)
        {
            RibbonSplitButton splitButton = new RibbonSplitButton
            {
                Text = buttonText,
                ShowText = true
            };

            foreach (var command in commands)
            {
                RibbonButton button = new RibbonButton
                {
                    Text = command.Key,
                    ShowText = true,
                    CommandHandler = new RibbonCommandHandler(command.Value.Command),
                    LargeImage = ConvertBitmapToImageSource(command.Value.Icon ?? Properties.Resources.Default_icon) // 使用指定或預設圖片
                };

                splitButton.Items.Add(button);
            }

            panelSource.Items.Add(splitButton);
        }

        // 將 Bitmap 轉換為 ImageSource
        private ImageSource ConvertBitmapToImageSource(Bitmap bitmap)
        {
            var hBitmap = bitmap.GetHbitmap();
            return Imaging.CreateBitmapSourceFromHBitmap(
                hBitmap,
                IntPtr.Zero,
                System.Windows.Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }
    }

    // 通用的指令處理
    public class RibbonCommandHandler : ICommand
    {
        private readonly string _commandName;

        public RibbonCommandHandler(string commandName)
        {
            _commandName = commandName;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter) => true;

        public void Execute(object parameter)
        {
            // 執行 AutoCAD 命令
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute(_commandName + " ", true, false, false);
        }
    }

    // 切換按鈕的指令處理
    public class ToggleCommandHandler : ICommand
    {
        private readonly string _commandOn;
        private readonly string _commandOff;
        private readonly RibbonToggleButton _toggleButton;

        public ToggleCommandHandler(string commandOn, string commandOff, RibbonToggleButton toggleButton)
        {
            _commandOn = commandOn;
            _commandOff = commandOff;
            _toggleButton = toggleButton;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter) => true;

        public void Execute(object parameter)
        {
            // 根據按鈕狀態選擇執行的命令
            string commandToExecute = _toggleButton.IsChecked == true ? _commandOn : _commandOff;
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute(commandToExecute + " ", true, false, false);
        }
    }
}
