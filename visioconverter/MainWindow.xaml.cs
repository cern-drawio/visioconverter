using System;
using System.IO;
using System.Windows;

using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Diagnostics;

namespace visioconverter
{
    public partial class MainWindow : Window
    {
        static String[] VSD_FILES = null;
        static String OUTPUT_FILES_PATH = null;
        public MainWindow()
        {
            if (CheckExecution())
            {
                ShowMessage("It looks like the application is already running in your system\n\nShutting down...", 
                    "Visioconverter already running", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
            else
            {
                InitializeComponent();
            }
        }

        private Boolean CheckExecution()
        {
            foreach (Process process in Process.GetProcesses())
            {
                try
                {
                    Process current = Process.GetCurrentProcess();
                    String pname = process.ProcessName.ToLower();
                    if (pname.IndexOf(current.ProcessName) >= 0 && process.Id != current.Id)
                        return true;
                }
                catch (Exception exception)
                {
                    CreateLog(exception.ToString());
                    String logPath = AppDomain.CurrentDomain.BaseDirectory + "error.log";
                    ShowMessage("An error occurred while checking the processes\nA log file has been created in " + logPath,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                }
            }

            return false;
        }

        private MessageBoxResult ShowMessage(String text, String title, MessageBoxButton options, MessageBoxImage image)
        {
            return MessageBox.Show(text, title, options, image);
        }

        private void BtnOpenFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFilesDialog = new OpenFileDialog();
            openFilesDialog.Multiselect = true;
            openFilesDialog.Filter = "VSD Files|*.vsd";
            if (openFilesDialog.ShowDialog() == true)
            {
                VSD_FILES = openFilesDialog.FileNames;
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("------------------------------");
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("Files selected: " + VSD_FILES.Length);
                infoBox.AppendText(Environment.NewLine);
            }

        }

        private void BtnSavePath_Click(object sender, RoutedEventArgs e)
        {
            var outputFileDialog = new CommonOpenFileDialog();
            outputFileDialog.IsFolderPicker = true;
            if (outputFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                OUTPUT_FILES_PATH = outputFileDialog.FileName;
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("------------------------------");
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("New output folder selected:");
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText(OUTPUT_FILES_PATH);
                infoBox.AppendText(Environment.NewLine);
            }
        }

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            btnConvert.IsEnabled = false;
            if (VSD_FILES == null)
            {
                ShowMessage("Please select at least one file",
                    "No files selected",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            else if (OUTPUT_FILES_PATH == null)
            {
                ShowMessage("Please select the output folder",
                    "No output folder selected",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            else
            {
                MessageBoxResult result = ShowMessage("The conversion will start now\nNote that all Visio instances must be closed\nDo you want to continue?",
                    "Confirmation message",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                if(result == MessageBoxResult.Yes)
                    Convert();
            }
            btnConvert.IsEnabled = true;
        }

        public void CreateLog(String exception)
        {
            using (StreamWriter writer = File.AppendText("error.log"))
            {
                writer.Write("\r\nLog Entry : ");
                writer.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                writer.WriteLine("  :");
                writer.WriteLine("  :{0}", exception);
                writer.WriteLine("-------------------------------");
            }
        }

        private void TextListener(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            infoBox.ScrollToEnd();
        }

        private void Convert()
        {
            infoBox.AppendText(Environment.NewLine);
            infoBox.AppendText("------------------------------");
            infoBox.AppendText(Environment.NewLine);
            infoBox.AppendText("Starting conversion process");
            infoBox.AppendText(Environment.NewLine);

            Microsoft.Office.Interop.Visio.InvisibleApp VisioInst = null;

            try
            {
                // Using COM to call visio to convert the files
                Type VisioType = Type.GetTypeFromProgID("Visio.InvisibleApp");

                foreach (String file in VSD_FILES)
                {
                    if (file.Length > 0)
                    {
                        string filename = Path.GetFileNameWithoutExtension(file);
                    
                        // VisioInst instances are created in each iteration to avoid RPC server issues
                        VisioInst = (Microsoft.Office.Interop.Visio.InvisibleApp)Activator.CreateInstance(VisioType);

                        infoBox.AppendText("Opening: " + filename + ".vsd");
                        infoBox.AppendText(Environment.NewLine);

                        // Open .vsd file
                        var doc = VisioInst.Documents.Open(Path.GetFullPath(file));

                        infoBox.AppendText("Saving: " + filename + ".vsdx");
                        infoBox.AppendText(Environment.NewLine);

                        // Save .vsdx file
                        doc.SaveAs(OUTPUT_FILES_PATH + "\\" + filename + ".vsdx");

                        // Close document
                        doc.Close();

                        // Close visio instance
                        VisioInst.Quit();
                        VisioInst = null;

                        infoBox.AppendText("Done!");
                        infoBox.AppendText(Environment.NewLine);
                        infoBox.AppendText(Environment.NewLine);
                    }
                }
            }
            catch (ArgumentNullException exception)
            {
                ShowMessage("This application requires Visio to make the conversion",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("------------------------------"); ;
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("The conversion process failed"); ;
                infoBox.AppendText(Environment.NewLine);
            }
            catch (Exception exception)
            {
                CreateLog(exception.ToString());
                String logPath = AppDomain.CurrentDomain.BaseDirectory + "error.log";
                ShowMessage("An error occurred while checking the processes\nA log file has been created in " + logPath,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("------------------------------"); ;
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("The conversion process failed"); ;
                infoBox.AppendText(Environment.NewLine);
            }
            finally
            {
                // Close visio if an unexpected error occured
                if (VisioInst != null)
                {
                    VisioInst.Quit();
                }
            }

            infoBox.AppendText(Environment.NewLine);
            infoBox.AppendText("------------------------------"); ;
            infoBox.AppendText(Environment.NewLine);
            infoBox.AppendText("All tasks finished"); ;
            infoBox.AppendText(Environment.NewLine);

        }

        private void btnConvert_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ClickCount > 1)
            {
                //here you would probably want to include code that is called by your
                //mouse down event handler.
                e.Handled = true;
            }
        }
    }
}
