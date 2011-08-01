//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs">(c) Controlled Vocabulary on CodePlex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Reflection;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using ControlledVocabulary.Properties;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Initializes a new instance of the MainWindow class
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            this.Width = Convert.ToInt32(Settings.Default.WindowWidth);
            this.Height = Convert.ToInt32(Settings.Default.WindowHeight);
            FileVersionInfo versionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
            this.Title += " - " + new Version(versionInfo.FileMajorPart, versionInfo.FileMinorPart, versionInfo.FileBuildPart, versionInfo.FilePrivatePart);
        }

        private static void Guidance(object sender, RoutedEventArgs e)
        {
            MenuItem m = (MenuItem)sender;
            string[] idParts = m.Uid.Split(new[] { StaticHelper.SplitSequence }, StringSplitOptions.RemoveEmptyEntries);
            string guidanceUrl = StaticHelper.GetGuidanceUrl(idParts[0]);
            System.Diagnostics.Process.Start(guidanceUrl);
        }

        private static void Launch(object sender, RoutedEventArgs e)
        {
            MenuItem m = (MenuItem)sender;
            System.Diagnostics.Process.Start(m.Tag.ToString());
        }

        private static void Send(object sender, RoutedEventArgs e)
        {
            MenuItem m = (MenuItem)sender;

            try
            {
                string[] idParts = m.Uid.Split(new[] { StaticHelper.SplitSequence }, StringSplitOptions.RemoveEmptyEntries);

                // Get the recipients
                string[] recipients = StaticHelper.GetRecipients(idParts[0], m.Uid);
                string mailto = "mailto:" + recipients[0];
                mailto += "?subject=" + m.Tag;

                if (!string.IsNullOrEmpty(recipients[1]))
                {
                    mailto += "&cc=" + recipients[1];
                }

                if (!string.IsNullOrEmpty(recipients[2]))
                {
                    mailto += "&bcc=" + recipients[2];
                }

                if (Convert.ToBoolean(StaticHelper.GetApplicationSetting("CopySubjectToClipboard")))
                {
                    Clipboard.SetText(m.Tag.ToString());
                }

                if (Convert.ToBoolean(StaticHelper.GetApplicationSetting("CallMailtoProtocol")))
                {
                    Process.Start(mailto);
                }
            }
            catch (System.Exception ex)
            {
                StaticHelper.LogMessage(MessageType.Error, ex.ToString());
                throw;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // check for updates
            if (Convert.ToBoolean(StaticHelper.GetApplicationSetting("AutoUpdate")))
            {
                if (string.IsNullOrEmpty(StaticHelper.GetApplicationSetting("LastUpdateCheckDate")))
                {
                    StaticHelper.CheckForMenuXmlUpdates();
                }
                else
                {
                    DateTime lastcheck = Convert.ToDateTime(StaticHelper.GetApplicationSetting("LastUpdateCheckDate"));
                    TimeSpan t = DateTime.Now - lastcheck;
                    if (t.Days >= Convert.ToInt32(StaticHelper.GetApplicationSetting("UpdateCheckFrequency")))
                    {
                        StaticHelper.CheckForMenuXmlUpdates();
                    }
                }
            }

            // get the buttons
            StaticHelper.LogMessage(MessageType.Info, "Getting buttons");
            menu[] buttons = StaticHelper.GetControlledVocabularyMenus();

            // build the buttons
            StaticHelper.LogMessage(MessageType.Info, "Building menu");
            this.BuildMenu(buttons);

            if (this.menu1.Items.Count == 0)
            {
                this.ShowManager();
            }
        }

        private void BuildMenu(IEnumerable<menu> buttons)
        {
            this.menu1.Items.Clear();
            foreach (menu customButton in buttons)
            {
                MenuItem newMenuItem = new MenuItem { Header = customButton.label };
                foreach (var item in customButton.Items)
                {
                    if (item is button)
                    {
                        button bb = (button)item;
                        MenuItem submenu = new MenuItem { Header = bb.label, Uid = bb.id, Tag = bb.tag };
                        switch (bb.onAction)
                        {
                            case "SendNormal":
                            case "SendHigh":
                            case "SendLow":
                                submenu.Click += Send;
                                break;
                            case "Guidance":
                                submenu.Click += Guidance;
                                break;
                            case "Launch":
                                submenu.Click += Launch;
                                break;
                            default:
                                submenu.IsEnabled = false;
                                break;
                        }

                        newMenuItem.Items.Add(submenu);
                    }
                    else if (item is menu)
                    {
                        menu m = (menu)item;
                        MenuItem submenu = new MenuItem { Header = m.label };

                        foreach (var b in m.Items)
                        {
                            if (b is button)
                            {
                                button bb = (button)b;
                                MenuItem subsmenu = new MenuItem { Header = bb.label, Uid = bb.id, Tag = bb.tag };
                                switch (bb.onAction)
                                {
                                    case "SendNormal":
                                    case "SendHigh":
                                    case "SendLow":
                                        subsmenu.Click += Send;
                                        break;
                                    case "Guidance":
                                        subsmenu.Click += Guidance;
                                        break;
                                    case "Launch":
                                        subsmenu.Click += Launch;
                                        break;
                                    default:
                                        subsmenu.IsEnabled = false;
                                        break;
                                }

                                submenu.Items.Add(subsmenu);
                            }
                            else if (b is menuSeparator)
                            {
                                submenu.Items.Add(new Separator());
                            }
                        }

                        newMenuItem.Items.Add(submenu);
                    }
                    else if (item is menuSeparator)
                    {
                        newMenuItem.Items.Add(new Separator());
                    }
                }

                this.menu1.Items.Add(newMenuItem);
            }
        }

        private void labelBlog_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            Process.Start(@"http://mikefourie.wordpress.com");
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            Settings.Default.WindowHeight = this.Height.ToString();
            Settings.Default.WindowWidth = this.Width.ToString();
            Settings.Default.Save();
        }

        private void labelManager_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            this.ShowManager();
        }

        private void ShowManager()
        {
            Manager managerWindow = new Manager();
            managerWindow.ShowDialog();
            StaticHelper.LogMessage(MessageType.Info, "Building menu");

            // get the buttons
            StaticHelper.LogMessage(MessageType.Info, "Getting buttons");
            menu[] buttons = StaticHelper.GetControlledVocabularyMenus();

            // build the buttons
            StaticHelper.LogMessage(MessageType.Info, "Building menu");
            this.BuildMenu(buttons);
        }

        private void labelCodePlex_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            Process.Start(@"http://controlledvocabulary.codeplex.com");
        }
    }
}
