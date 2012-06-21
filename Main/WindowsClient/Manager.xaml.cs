//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Manager.xaml.cs">(c) Controlled Vocabulary on CodePlex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Windows;
    using System.Windows.Forms;
    using System.Windows.Input;
    using System.Xml;
    using Microsoft.VisualBasic.FileIO;
    using MessageBox = System.Windows.Forms.MessageBox;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class Manager
    {
        private readonly List<CheckedListBoxItem> checkedListItems = new List<CheckedListBoxItem>();
        private XmlDocument xdoc;
        private bool initializing = true;

        /// <summary>
        /// Initializes a new instance of the Manager class
        /// </summary>
        public Manager()
        {
            InitializeComponent();
            this.checkBoxAutoUpdate.IsChecked = Convert.ToBoolean(StaticHelper.GetApplicationSetting("AutoUpdate"));
        }

        /// <summary>
        /// Initializes a new instance of the Manager class
        /// </summary>
        /// <param name="cvcfPath">cvcfPath</param>
        public Manager(string cvcfPath)
        {
            InitializeComponent();
            this.textBoxDiscover.Text = cvcfPath;
            this.DiscoverConfig();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.initializing = false;

            // get the buttons install location
            this.labelAppData.Content = string.Format(CultureInfo.InvariantCulture, @"{0}\Controlled Vocabulary", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
            this.GetButtons();

            try
            {
                Process[] processes = Process.GetProcessesByName("OUTLOOK");
                if (processes.Length > 0)
                {
                    this.labelOutlookRunning.Visibility = System.Windows.Visibility.Visible;
                }
            }
            catch
            {
                // Do Nothing
            }

            this.checkBoxCallMailto.IsChecked = Convert.ToBoolean(StaticHelper.GetApplicationSetting("CallMailtoProtocol"));
            this.checkBoxCopySubject.IsChecked = Convert.ToBoolean(StaticHelper.GetApplicationSetting("CopySubjectToClipboard"));
            this.textBoxMasterEmailAccount.Text = StaticHelper.GetApplicationSetting("MasterEmailAccount");
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

        private void buttonRemove_Click(object sender, RoutedEventArgs e)
        {
            if (this.listBoxButtons.SelectedIndex >= 0)
            {
                if (Directory.Exists(this.labelAppData.Content + @"\Buttons\" + this.listBoxButtons.SelectedValue))
                {
                    if (MessageBox.Show("Would you like to delete " + this.listBoxButtons.SelectedItem + "?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                    {
                        try
                        {
                            Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(this.labelAppData.Content + @"\Buttons\" + this.listBoxButtons.SelectedValue, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
                            this.GetButtons();
                        }
                        catch (Exception)
                        {
                            // do nothing
                        }
                    }
                }
            }
        }

        private void buttonDiscover_Click_1(object sender, RoutedEventArgs e)
        {
            this.DiscoverConfig();
        }

        private void DiscoverConfig()
        {
            if (!string.IsNullOrEmpty(this.textBoxDiscover.Text))
            {
                this.listboxDiscovered.ItemsSource = null;
                this.listboxDiscovered.Items.Clear();
                this.checkedListItems.Clear();

                try
                {
                    this.xdoc = new XmlDocument();
                    this.xdoc.Load(this.textBoxDiscover.Text);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Discovery Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                XmlNodeList buttons = this.xdoc.SelectNodes("/cvcf/button");
                if (buttons != null)
                {
                    foreach (CheckedListBoxItem item in from XmlNode buttonNode in buttons select new CheckedListBoxItem { Name = buttonNode.Attributes["name"].Value, IsChecked = true, SourcePath = buttonNode.Attributes["sourcePath"].Value })
                    {
                        this.checkedListItems.Add(item);
                    }
                }

                this.listboxDiscovered.ItemsSource = this.checkedListItems;
            }
            else
            {
                MessageBox.Show("Please enter a file path or URL", "Data Required", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void buttonAddDiscovered_Click_1(object sender, RoutedEventArgs e)
        {
            foreach (CheckedListBoxItem item in this.checkedListItems)
            {
                if (item.IsChecked)
                {
                    if (!StaticHelper.DeployZippedButton(item.SourcePath, item.Name))
                    {
                        MessageBox.Show("Deploy Failed. Check your Event Log", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            this.GetButtons();
        }

        private void buttonCheckForUpdates_Click(object sender, RoutedEventArgs e)
        {
            if (StaticHelper.CheckForMenuXmlUpdates())
            {
                MessageBox.Show("Update is complete", "Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Update failed. Check your Event Log", "Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void checkBoxAutoUpdate_Checked(object sender, RoutedEventArgs e)
        {
            if (this.initializing)
            {
                return;
            }

            StaticHelper.SetApplicationSetting("AutoUpdate", this.checkBoxAutoUpdate.IsChecked.ToString());
        }

        private void checkBoxAutoUpdate_Unchecked(object sender, RoutedEventArgs e)
        {
            if (this.initializing)
            {
                return;
            }

            StaticHelper.SetApplicationSetting("AutoUpdate", this.checkBoxAutoUpdate.IsChecked.ToString());
        }

        private void checkBoxCallMailto_Unchecked(object sender, RoutedEventArgs e)
        {
            if (this.initializing)
            {
                return;
            }

            StaticHelper.SetApplicationSetting("CallMailtoProtocol", this.checkBoxCallMailto.IsChecked.ToString());
        }

        private void checkBoxCallMailto_Checked(object sender, RoutedEventArgs e)
        {
            if (this.initializing)
            {
                return;
            }

            StaticHelper.SetApplicationSetting("CallMailtoProtocol", this.checkBoxCallMailto.IsChecked.ToString());
        }

        private void checkBoxCopySubject_Unchecked(object sender, RoutedEventArgs e)
        {
            if (this.initializing)
            {
                return;
            }

            StaticHelper.SetApplicationSetting("CopySubjectToClipboard", this.checkBoxCopySubject.IsChecked.ToString());
        }

        private void checkBoxCopySubject_Checked(object sender, RoutedEventArgs e)
        {
            if (this.initializing)
            {
                return;
            }

            StaticHelper.SetApplicationSetting("CopySubjectToClipboard", this.checkBoxCopySubject.IsChecked.ToString());
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StaticHelper.SetApplicationSetting("MasterEmailAccount", this.textBoxMasterEmailAccount.Text);
        }
    }
}
