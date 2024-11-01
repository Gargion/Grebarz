using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Grebarz古霸子
{
    public partial class WPF_LayerSelection : Window
    {
        Database db = HostApplicationServices.WorkingDatabase;

        public WPF_LayerSelection(List<string> frozenLayers)
        {
            InitializeComponent();
            LayerListBox.ItemsSource = frozenLayers; // 假設你有一個 ListBox 控件來顯示層
        }

        private void UnfreezeButton_Click(object sender, RoutedEventArgs e)
        {
            // 獲取選定的圖層
            var selectedLayers = LayerListBox.SelectedItems.Cast<string>().ToList();

            // 進行解凍操作
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                foreach (var layerName in selectedLayers)
                {
                    var layerId = layerTable[layerName];
                    LayerTableRecord layerRecord = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                    layerRecord.IsFrozen = false; // 解凍層
                }
                tr.Commit();
            }

            this.Close(); // 關閉窗口
        }
    }
}
