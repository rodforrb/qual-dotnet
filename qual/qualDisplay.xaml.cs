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
using Excel;

namespace qual
{
    // Interaction logic for qualDisplay.xaml
    public partial class qualDisplay : Page
    {
        Parser parser;      // backend parser for display page
        String fileSource;  // file source given by previous page
        public qualDisplay(String fileName)
        {
            /* Constructor to initialize parser for 'Display' page */
            fileSource = fileName;
            parser = new Parser();  // instantiate a parser
            InitializeComponent();  // initialize page
            LoadWorksheet();        // read data from a worksheet
            LoadEmails();           // read emails from helper file
            displayGrid.ItemsSource = LoadEntries(); // attach row entries to the displayGrid
        }

        private void LoadWorksheet()
        {
            /* Update parser with a given qualification sheet */
            worksheet ws = Workbook.Worksheets(fileSource).ElementAt(0);
            parser.ParseSheet(ws);
        }

        private void LoadEmails()
        {
            /* update parser with external list of emails */
            String emailFile = "emails.xlsx";
            worksheet ws = Workbook.Worksheets(emailFile).ElementAt(0);
            parser.ParseEmails(ws);
        } 
        private List<RowEntry> LoadEntries()
        {
            /* Get employee entry list from parser */
            return parser.GetEntries();
        }

        private void displayGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // no use for now
        }

        private void btnSend_Click(object sender, RoutedEventArgs e)
        {
            /* "Send" button clicked inside a row */
            // acquire actual RowEntry which was clicked
            RowEntry row = (RowEntry)((Button)e.Source).DataContext;
            // initiate an email
            parser.SendEmail(row);

            // update data with 'emailed' status
            // TODO

            // reload/refresh display
            displayGrid.ItemsSource = null;
            displayGrid.ItemsSource = LoadEntries();
        }

        //private void btnOpen_Click(object sender, RoutedEventArgs e)
        //{    
        //    // Create OpenFileDialog
        //    Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();

        //    // Launch OpenFileDialog by calling ShowDialog method
        //    Nullable<bool> result = openFileDlg.ShowDialog();

        //    // Get the selected file name and display in a TextBox.
        //    // Load content of file in a TextBlock
        //    if (result == true)
        //    {
        //        FileNameTextBox.Text = openFileDlg.FileName;
        //    }
        //}
    }
}
