//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs">(c) Controlled Vocabulary on CodePlex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
    using System;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Windows;
    using System.Windows.Forms;
    using System.Windows.Input;
    using Microsoft.VisualBasic.FileIO;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            // Check if the target directory exists, if not, create it.
            if (Directory.Exists(target.FullName) == false)
            {
                Directory.CreateDirectory(target.FullName);
            }

            // Copy each file into it's new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                fi.CopyTo(Path.Combine(target.ToString(), fi.Name), true);
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo subdir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(subdir.Name);
                CopyAll(subdir, nextTargetSubDir);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // get the buttons install location
            this.labelAppData.Content = string.Format(CultureInfo.InvariantCulture, @"{0}\Controlled Vocabulary", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
            this.GetButtons();

            try
            {
                Process[] processes = Process.GetProcessesByName("OUTLOOK");
                this.labelOutookRunning.Visibility = processes.Length > 0 ? System.Windows.Visibility.Visible : System.Windows.Visibility.Hidden;
            }
            catch
            {
                // Do Nothing
            }
        }

        private void GetButtons()
        {
            this.listBoxButtons.Items.Clear();
            DirectoryInfo d = new DirectoryInfo(this.labelAppData.Content + @"\Buttons");
            foreach (DirectoryInfo buttonDir in d.GetDirectories())
            {
                this.listBoxButtons.Items.Add(buttonDir.Name);
            }

            if (this.listBoxButtons.Items.Count > 0)
            {
                this.listBoxButtons.SelectedIndex = 0;
            }
        }

        private void labelAppData_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            Process.Start(this.labelAppData.Content.ToString());
        }

        private void labelCodePlex_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            Process.Start(@"http://controlledvocabulary.codeplex.com");
        }

        private void labelBlog_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            Process.Start(@"http://www.freetodev.com");
        }

        private void buttonRemove_Click(object sender, RoutedEventArgs e)
        {
            if (this.listBoxButtons.SelectedIndex >= 0)
            {
                if (Directory.Exists(this.labelAppData.Content + @"\Buttons\" + this.listBoxButtons.SelectedValue))
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(this.labelAppData.Content + @"\Buttons\" + this.listBoxButtons.SelectedValue, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
                    this.GetButtons();
                }
            }
        }

        private void buttonAdd_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new FolderBrowserDialog { Description = "Choose a button Folder" };
            if (dlg.ShowDialog(this.GetIWin32Window()) == System.Windows.Forms.DialogResult.OK)
            {
                DirectoryInfo source = new DirectoryInfo(dlg.SelectedPath);
                DirectoryInfo destination = new DirectoryInfo(this.labelAppData.Content + @"\Buttons\" + source.Name);
                CopyAll(source, destination);
                this.GetButtons();
            }
        }
    }
}
