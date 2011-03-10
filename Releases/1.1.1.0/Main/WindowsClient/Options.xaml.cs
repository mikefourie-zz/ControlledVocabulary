//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Options.xaml.cs">(c) Controlled Vocabulary on CodePlex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
    using System.Windows;
    using ControlledVocabulary.Properties;

    /// <summary>
    /// Interaction logic for Options.xaml
    /// </summary>
    public partial class Options : Window
    {
        public Options()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.checkBoxCallMailto.IsChecked = Settings.Default.CallMailTo;
            this.checkBoxCopySubject.IsChecked = Settings.Default.CopyToClipboard;
        }

        private void checkBoxCallMailto_Checked(object sender, RoutedEventArgs e)
        {
            this.SaveSettings();
        }

        private void SaveSettings()
        {
            Settings.Default.CallMailTo = (bool)this.checkBoxCallMailto.IsChecked;
            Settings.Default.CopyToClipboard = (bool)this.checkBoxCopySubject.IsChecked;
            Settings.Default.Save();
        }

        private void checkBoxCopySubject_Checked(object sender, RoutedEventArgs e)
        {
            this.SaveSettings();
        }
    }
}
