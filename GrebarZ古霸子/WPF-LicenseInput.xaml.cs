using System.Windows;

namespace Grebarz古霸子
{
    public partial class WPF_LicenseInput : Window
    {
        public string EnteredLicenseKey { get; private set; }

        public WPF_LicenseInput()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            EnteredLicenseKey = LicenseTextBox.Text;
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
