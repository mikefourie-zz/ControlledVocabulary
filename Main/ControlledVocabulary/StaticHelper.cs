//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="StaticHelper.cs">(c) Controlled Vocabulary on CodePlex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
    using System;
    using System.Diagnostics;
    using System.Drawing;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Xml.Serialization;

    public static class StaticHelper
    {
        public const string SplitSequence = "...";

        public static void CheckForUpdates()
        {
            try
            {
                // Get the installation path
                DirectoryInfo installationPath = GetInstallationPath();

                XmlSerializer deserializer = new XmlSerializer(typeof(ButtonConfiguration));
                ButtonConfiguration buttonConfig;

                // Iterate over Add-ins found
                DirectoryInfo buttonRoot = new DirectoryInfo(Path.Combine(installationPath.FullName, "Buttons"));
                DirectoryInfo[] buttons = buttonRoot.GetDirectories();
                foreach (FileInfo file in buttons.Select(button => new FileInfo(Path.Combine(button.FullName, "config.xml"))))
                {
                    // open the configuration file for the button
                    using (FileStream buttonStream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read))
                    {
                        buttonConfig = (ButtonConfiguration)deserializer.Deserialize(buttonStream);
                    }

                    // if an onlineUrl is present, then we look for an update
                    if (!string.IsNullOrEmpty(buttonConfig.onlineUrl))
                    {
                        string currentMenu;
                        using (TextReader tr = new StreamReader(Path.Combine(file.DirectoryName, @"button.xml")))
                        {
                            currentMenu = tr.ReadToEnd();
                        }

                        using (System.Net.WebClient client = new System.Net.WebClient())
                        {
                            // prevent file caching by windows
                            client.CachePolicy = new System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore);

                            Stream myStream = client.OpenRead(buttonConfig.onlineUrl);
                            if (myStream != null)
                            {
                                using (StreamReader sr = new StreamReader(myStream))
                                {
                                    string latestMenu = sr.ReadToEnd();
                                    if (latestMenu != currentMenu)
                                    {
                                        using (TextWriter tw = new StreamWriter(Path.Combine(file.DirectoryName, @"button.xml")))
                                        {
                                            LogMessage(MessageType.Info, "Updating menu with online updates");
                                            tw.Write(latestMenu);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage(MessageType.Warning, "Update check failed." + ex.Message);

                // swallow the error.
            }
        }

        public static DirectoryInfo GetInstallationPath()
        {
            DirectoryInfo installationPath = new DirectoryInfo(string.Format(CultureInfo.InvariantCulture, @"{0}\Controlled Vocabulary", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)));
            string paths = string.Empty;
            if (!installationPath.Exists)
            {
                string message = string.Format(CultureInfo.InvariantCulture, "Installation Path not found: {0}", paths);
                LogMessage(MessageType.Error, message);
                throw new ArgumentException(message);
            }

            return installationPath;
        }

        public static Bitmap GetImage(string image)
        {
            if (string.IsNullOrEmpty(image))
            {
                LogMessage(MessageType.Error, "GetImage received invalid image parameter");
                throw new ArgumentException("GetImage received invalid image parameter");
            }

            // First we need to find which Button was clicked
            string[] imageParts = image.Split(new[] { SplitSequence }, StringSplitOptions.RemoveEmptyEntries);

            // Get the installation path
            DirectoryInfo installationPath = GetInstallationPath();

            // Get the image file
            FileInfo f = new FileInfo(Path.Combine(installationPath.FullName, @"Buttons\" + imageParts[0] + @"\images\" + imageParts[1]));
            if (!f.Exists)
            {
                string message = string.Format(CultureInfo.InvariantCulture, "Image file not found: {0}", f.FullName);
                LogMessage(MessageType.Error, message);
                throw new ArgumentException(message);
            }

            return new Bitmap(f.FullName);
        }

        public static string[] GetRecipients(string buttonId, string controlId)
        {
            // Get the installation path
            DirectoryInfo installationPath = GetInstallationPath();
            string[] recipients = new string[3];
            bool foundrecipients = false;

            XmlSerializer deserializer = new XmlSerializer(typeof(menu));
            menu cvmenu;
            FileInfo f = new FileInfo(Path.Combine(installationPath.FullName, @"Buttons\" + buttonId + @"\button.xml"));
            using (FileStream buttonStream = new FileStream(f.FullName, FileMode.Open, FileAccess.Read))
            {
                cvmenu = (menu)deserializer.Deserialize(buttonStream);
            }

            foreach (var item in cvmenu.Items)
            {
                if (foundrecipients)
                {
                    break;
                }

                if (item is button)
                {
                    button b = (button)item;
                    if (b.id == controlId)
                    {
                        if (!string.IsNullOrEmpty(b.toRecipients))
                        {
                            recipients[0] = b.toRecipients;
                            recipients[1] = b.ccRecipients;
                            recipients[2] = b.bccRecipients;
                            foundrecipients = true;
                        }

                        break;
                    }
                }
                else if (item is menu)
                {
                    menu m = (menu)item;
                    foreach (var b in m.Items)
                    {
                        if (b is button)
                        {
                            button bb = (button)b;
                            if (bb.id == controlId)
                            {
                                if (!string.IsNullOrEmpty(bb.toRecipients))
                                {
                                    recipients[0] = bb.toRecipients;
                                    recipients[1] = bb.ccRecipients;
                                    recipients[2] = bb.bccRecipients;
                                    foundrecipients = true;
                                }

                                break;
                            }
                        }
                    }
                }
            }

            if (!foundrecipients)
            {
                deserializer = new XmlSerializer(typeof(ButtonConfiguration));

                ButtonConfiguration buttonConfig;
                f = new FileInfo(Path.Combine(installationPath.FullName, @"Buttons\" + buttonId + @"\config.xml"));
                using (FileStream buttonStream = new FileStream(f.FullName, FileMode.Open, FileAccess.Read))
                {
                    buttonConfig = (ButtonConfiguration)deserializer.Deserialize(buttonStream);
                }

                recipients[0] = buttonConfig.toRecipients;
                recipients[1] = buttonConfig.ccRecipients;
                recipients[2] = buttonConfig.bccRecipients;
            }

            return recipients;
        }

        public static string GetFromAccount(string buttonId)
        {
            // Get the installation path
            DirectoryInfo installationPath = GetInstallationPath();
            XmlSerializer deserializer = new XmlSerializer(typeof(ButtonConfiguration));

            ButtonConfiguration buttonConfig;
            FileInfo f = new FileInfo(Path.Combine(installationPath.FullName, @"Buttons\" + buttonId + @"\config.xml"));
            using (FileStream buttonStream = new FileStream(f.FullName, FileMode.Open, FileAccess.Read))
            {
                buttonConfig = (ButtonConfiguration)deserializer.Deserialize(buttonStream);
            }

            return buttonConfig.from;
        }

        public static string GetGuidanceUrl(string buttonId)
        {
            // Get the installation path
            DirectoryInfo installationPath = GetInstallationPath();

            XmlSerializer deserializer = new XmlSerializer(typeof(ButtonConfiguration));
            ButtonConfiguration buttonConfig;
            FileInfo f = new FileInfo(Path.Combine(installationPath.FullName, @"Buttons\" + buttonId + @"\config.xml"));
            using (FileStream buttonStream = new FileStream(f.FullName, FileMode.Open, FileAccess.Read))
            {
                buttonConfig = (ButtonConfiguration)deserializer.Deserialize(buttonStream);
            }

            return buttonConfig.guidanceUrl;
        }

        public static menu[] GetControlledVocabularyMenus()
        {
            // Get the installation path
            DirectoryInfo installationPath = StaticHelper.GetInstallationPath();

            // Iterate over Add-ins found
            DirectoryInfo buttonRoot = new DirectoryInfo(Path.Combine(installationPath.FullName, "Buttons"));
            DirectoryInfo[] buttons = buttonRoot.GetDirectories();

            menu[] menus = new menu[buttons.Length];
            int i = 0;
            foreach (FileInfo file in buttons.Select(button => new FileInfo(Path.Combine(button.FullName, "button.xml"))))
            {
                XmlSerializer deserializer = new XmlSerializer(typeof(menu));
                using (FileStream buttonStream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read))
                {
                    menus[i] = (menu)deserializer.Deserialize(buttonStream);
                }
                
                i++;
            }

            return menus;
        }

        public static void LogMessage(MessageType messageType, string error)
        {
            if (messageType == MessageType.Error)
            {
                using (EventLog eventLog = new EventLog())
                {
                    eventLog.Source = "ControlledVocabulary";
                    eventLog.Log = "Application";
                    eventLog.WriteEntry(error);
                }

                return;
            }

            // Get the installation path
            DirectoryInfo installationPath = StaticHelper.GetInstallationPath();
            if (!File.Exists(installationPath + @"\enablelogging.txt"))
            {
                return;
            }

            DirectoryInfo logDirectory = new DirectoryInfo(@"C:\ControlledVocabularyLog");
            if (!logDirectory.Exists)
            {
                logDirectory.Create();
            }

            using (TextWriter tw = new StreamWriter(logDirectory.FullName + @"\Log.txt", true, Encoding.UTF8))
            {
                tw.WriteLine(string.Format(CultureInfo.InvariantCulture, "{0} - {1}: {2}", DateTime.Now, messageType, error));
            }
        }
    }
}
