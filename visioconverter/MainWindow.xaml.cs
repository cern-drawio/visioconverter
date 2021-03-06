﻿using System;
using System.IO;
using System.Windows;

using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;

namespace visioconverter
{
    public partial class MainWindow : Window
    {
        const string openingText = "Opening: {0}";
        const string savingText = "Saving: {0}";

        static string[] VSD_FILES = null;
        static string OUTPUT_FILES_PATH = null;

        BackgroundWorker worker = null;
        
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

        private void OnContentRendered(object sender, EventArgs e)
        {
            worker = new BackgroundWorker
            {
                WorkerReportsProgress = true
            };
            worker.DoWork += Convert;
            worker.ProgressChanged += OnProgressChanged;
            worker.RunWorkerCompleted += OnConversionCompleted;
        }

        void OnProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
        }

        void OnConversionCompleted(object sender, EventArgs e)
        {
            btnOpenFiles.IsEnabled = true;
            btnOpenFolder.IsEnabled = true;
            btnSavePath.IsEnabled = true;
            btnConvert.IsEnabled = true;
        }

        // Check if the application is already running
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

        // Shows a message box with the information specified
        private MessageBoxResult ShowMessage(String text, String title, MessageBoxButton options, MessageBoxImage image)
        {
            return MessageBox.Show(text, title, options, image);
        }

        // InputFiles button handler
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
                infoBox.AppendText("Number of files selected: " + VSD_FILES.Length);
                infoBox.AppendText(Environment.NewLine);
            }
        }

        // OutputFolder button handler
        private async void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            var openFolderDialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true
            };

            if (openFolderDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                // Set progressbar as indeterminate while loading
                progressBar.IsIndeterminate = true;

                // Load files in a different thread so the UI can is not blocked.
                VSD_FILES = await Task.Run(() => {
                    return GetFolderFiles(openFolderDialog.FileName);
                });

                // Set progrssbar back to normal
                progressBar.IsIndeterminate = false;

                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("------------------------------");
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("Number of files selected: " + VSD_FILES.Length);
                infoBox.AppendText(Environment.NewLine);
            }
        }

        // Get all the files inside te specified folder
        private string[] GetFolderFiles(string folder)
        {
            List<string> files = new List<string>();
            foreach (string file in Directory.GetFiles(folder, "*.vsd", SearchOption.AllDirectories).Where(item => item.EndsWith(".vsd")))
            {
                files.Add(file);
            }
            return files.ToArray();
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

        // Convert button handler
        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            btnOpenFiles.IsEnabled = false;
            btnOpenFolder.IsEnabled = false;
            btnSavePath.IsEnabled = false;
            btnConvert.IsEnabled = false;
            progressBar.Value = 0;
            if (VSD_FILES == null)
            {
                ShowMessage("Please select at least one file",
                    "No files selected",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            else
            {
                MessageBoxResult result = ShowMessage("The conversion will start now.\n" +
                    "Note that all Visio instances must be closed.\n\n" +
                    "If no output folder has been selected the converted file will be stored next to the source file." +
                    "Any existing file with the same name (.vsdx) will be overwritten.\n\n" +
                    "Do you want to continue?",
                    "Confirmation message",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                if(result == MessageBoxResult.Yes)
                    worker.RunWorkerAsync();
            }
        }

        // Create log file containing the exception occurred
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

        // Start the file conversion
        private void Convert(object sender, DoWorkEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(() => 
            {
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("------------------------------");
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("Starting conversion process");
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText(Environment.NewLine);
            });

            Microsoft.Office.Interop.Visio.InvisibleApp VisioInst = null;

            try
            {
                // Using COM to call visio to convert the files
                Type VisioType = Type.GetTypeFromProgID("Visio.InvisibleApp");
                
                // Create visio instance to open all the files
                VisioInst = (Microsoft.Office.Interop.Visio.InvisibleApp)Activator.CreateInstance(VisioType);

                // If an Output Folder was selected, the save to that folder
                // Otherwise, save in the same location of the source file
                Func<string, string> SaveLocation = null;
                if (OUTPUT_FILES_PATH == null)
                {
                    SaveLocation = (filePath) =>  string.Format("{0}x", filePath);
                }
                else
                {
                    SaveLocation = (filePath) => string.Format("{0}\\{1}x", OUTPUT_FILES_PATH, Path.GetFileName(filePath));
                }

                //foreach (String file in VSD_FILES)
                for (int i = 0; i < VSD_FILES.Length; i++)
                {
                    var file = VSD_FILES[i];

                    try
                    {
                        if (file.Length > 0)
                        {
                            string openLocation = Path.GetFullPath(file);

                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                infoBox.AppendText(string.Format(openingText, openLocation));
                                infoBox.AppendText(Environment.NewLine);
                            });

                            // Open .vsd file
                            var doc = VisioInst.Documents.Open(openLocation);
                            string saveLocation = SaveLocation(openLocation);

                            // Add wait time to avoid the "The RPC server is unavailable" exception
                            Thread.Sleep(1000);

                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                infoBox.AppendText(string.Format(savingText, saveLocation));
                                infoBox.AppendText(Environment.NewLine);
                            });

                            // Save .vsdx file
                            doc.SaveAs(saveLocation);

                            // Close document
                            doc.Close();

                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                infoBox.AppendText("Done!");
                                infoBox.AppendText(Environment.NewLine);
                                infoBox.AppendText(Environment.NewLine);
                            });
                        }
                    }
                    catch(Exception loopException)
                    {
                        CreateLog(loopException.ToString());
                        String logPath = AppDomain.CurrentDomain.BaseDirectory + "error.log";

                        ShowMessage("An error occurred while while converting this file\nA log file has been created in " + logPath + "\nClick OK to continue.",
                                    "Error",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);

                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            infoBox.AppendText(Environment.NewLine);
                            infoBox.AppendText(string.Format("An error occurred while converting this file, please check the logs.")); ;
                            infoBox.AppendText(Environment.NewLine);
                            infoBox.AppendText(Environment.NewLine);
                        });
                    }

                    (sender as BackgroundWorker).ReportProgress((i + 1) * 100 / VSD_FILES.Length);
                }

                Application.Current.Dispatcher.Invoke(() =>
                {
                    infoBox.AppendText("All the files were converted!");
                    infoBox.AppendText(Environment.NewLine);
                    infoBox.AppendText("The application can now be closed.");
                    infoBox.AppendText(Environment.NewLine);
                });
            }
            catch (ArgumentNullException)
            {
                ShowMessage("This application requires Visio to make the conversion.",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                Application.Current.Dispatcher.Invoke(() =>
                {
                    infoBox.AppendText(Environment.NewLine);
                    infoBox.AppendText("------------------------------"); ;
                    infoBox.AppendText(Environment.NewLine);
                    infoBox.AppendText("The conversion process failed."); ;
                    infoBox.AppendText(Environment.NewLine);
                });
            }
            catch (Exception exception)
            {
                CreateLog(exception.ToString());
                String logPath = AppDomain.CurrentDomain.BaseDirectory + "error.log";
                ShowMessage("An error occurred while checking the processes\nA log file has been created in " + logPath,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                Application.Current.Dispatcher.Invoke(() =>
                {
                    infoBox.AppendText(Environment.NewLine);
                    infoBox.AppendText("------------------------------"); ;
                    infoBox.AppendText(Environment.NewLine);
                    infoBox.AppendText("The conversion process failed."); ;
                    infoBox.AppendText(Environment.NewLine);
                });
            }
            finally
            {
                // Close visio if an unexpected error occured
                if (VisioInst != null)
                {
                    // Close visio instance
                    VisioInst.Quit();
                    VisioInst = null;
                }
            }

            Application.Current.Dispatcher.Invoke(() =>
            {
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("------------------------------"); ;
                infoBox.AppendText(Environment.NewLine);
                infoBox.AppendText("All tasks finished"); ;
                infoBox.AppendText(Environment.NewLine);
            });

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

        public string HelpText
        {
            get { return "My Text \n Your Text"; }
        }
    }
}
