using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SmartCrossRef
{
    /// <summary>
    /// Interaction logic for winAbout.xaml
    /// </summary>
    public partial class winAbout : System.Windows.Window
    {
        public winAbout()
        {
            InitializeComponent();
        }

        private void BtnCoffee_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Replace 'YOUR_USERNAME' with your actual Buy Me a Coffee page handle
                string url = "https://www.buymeacoffee.com/YOUR_USERNAME";

                // Safely opens the default system web browser
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open the link: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
