using System;
using System.Collections.Generic;
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

namespace qual
{
    /// <summary>
    /// Interaction logic for qualHome.xaml
    /// </summary>
    public partial class qualHome : Page
    {
        private bool fileSet = false;
        string fileSource;
        public qualHome()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (fileSet)
            {
                qualDisplay display = new qualDisplay(fileSource);
                this.NavigationService.Navigate(display);
            }
        }

        // handles "open file" button
        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();

            // Launch OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = openFileDlg.ShowDialog();

            // Get the selected file name and display in a TextBox.
            // Load content of file in a TextBlock
            if (result == true)
            {
                FileNameTextBox.Text = openFileDlg.FileName;
                fileSet = true;
                fileSource = openFileDlg.FileName;
            }
        }

        private void FileNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // nothing
        }
    }

}
