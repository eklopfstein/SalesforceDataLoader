using System;
using System.Diagnostics;
using System.Linq;
using System.Windows;

namespace DataLoadUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private String excelFile;
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// On click event to open dialog for users to select the desired excel file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFileSearch_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";

            // Display OpenFileDialog by calling ShowDialog method 
            bool? result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                excelFile = dlg.FileName;
                tbxExcelFile.Text = excelFile.Split('\\').Last();
            }
        }

        /// <summary>
        /// On click event to try to upload the test data to the given Salesforce
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpload_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Check for any incomplete fields
                if (tbxExcelFile.Text == "" || tbxUsername.Text == "" || pbxPassword.Password == "" || pbxToken.Password == "")
                {
                    // If there are any, tell the user to fill them out and do not try to upload
                    lblError.Content = "Please fill out all textboxes";
                    lblError.Visibility = Visibility.Visible;
                }
                else
                {
                    // If all fields are filled out, try to run the upload
                    lblError.Visibility = Visibility.Hidden;
                    String batchFile = Environment.CurrentDirectory + "\\runUpload.bat";
                    String arguments = excelFile + " " + tbxUsername.Text + " " + pbxPassword.Password + " " + pbxToken.Password + " " + rbnCreateUsers.IsChecked.ToString();
                    tbxLog.Text = ExecuteUpload(batchFile, arguments);
                }
            }
            catch (Exception ex)
            {
                lblError.Content = ex.Message;
                lblError.Visibility = Visibility.Visible;
            }
        }

        /// <summary>
        /// Takes an excel file and uploads the contents to a Salesforce sandbox
        /// </summary>
        /// <param name="fileName">Full path to excel file</param>
        /// <param name="arguments">String of Salesforce login information formatted like, "username password securityToken True/False"
        ///         The last value is wether or not to create user objects</param>
        /// <returns>Log messages from the execution</returns>
        public static String ExecuteUpload(String fileName, String arguments)
        {
            Process dataUpload = new Process();

            // Redirect the output stream of the child process.
            dataUpload.StartInfo.UseShellExecute = false;
            dataUpload.StartInfo.RedirectStandardOutput = true;
            dataUpload.StartInfo.RedirectStandardError = true;
            dataUpload.StartInfo.FileName = fileName;
            dataUpload.StartInfo.Arguments = arguments;
            dataUpload.Start();

            // Read the output stream first and then wait.
            String output = dataUpload.StandardOutput.ReadToEnd();
            dataUpload.WaitForExit();
            return output;
        }
    }
}
