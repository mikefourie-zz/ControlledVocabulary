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
    using System.Management;
    using System.Net;
    using System.Net.Cache;
    using System.Text;
    using System.Xml;
    using System.Xml.Serialization;
    using Ionic.Zip;

    public static class StaticHelper
    {
        public const string SplitSequence = "...";

        public static bool UpgradeButton(ButtonConfiguration buttonConfig, FileSystemInfo file)
        {
            LogMessage(MessageType.Info, "Checking Menu Versions");

            try
            {
                using (var client = GetPreconfiguredWebClient())
                {
                    if (string.IsNullOrEmpty(buttonConfig.versionUrl))
                    {
                        return false;
                    }

                    Stream myStream = client.OpenRead(buttonConfig.versionUrl);
                    if (myStream != null)
                    {
                        using (StreamReader sr = new StreamReader(myStream))
                        {
                            string latestVersion = sr.ReadToEnd();
                            if (latestVersion != buttonConfig.currentVersion)
                            {
                                // Download the latest version
                                DeployZippedButton(buttonConfig.sourceUrl, file.Name);
                                return true;
                            }
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogMessage(MessageType.Error, "Update check failed." + ex.Message);

                return false;

                // swallow the error.
            }
        }

        public static void DeployZippedButton(string sourceUrl, string buttonName)
        {
            using (var client = GetPreconfiguredWebClient())
            {
                client.DownloadFile(sourceUrl, StaticHelper.GetCachePath().FullName + @"\" + buttonName + ".zip");
            }

            RemoveContent(new DirectoryInfo(Path.Combine(StaticHelper.GetButtonsPath().FullName, buttonName)));                                
            using (ZipFile zip = ZipFile.Read(StaticHelper.GetCachePath().FullName + @"\" + buttonName + ".zip"))
            {
                foreach (ZipEntry e in zip)
                {
                    e.Extract(StaticHelper.GetButtonsPath().FullName, ExtractExistingFileAction.OverwriteSilently);
                }
            }

            RemoveContent(StaticHelper.GetCachePath());
        }

        public static void RemoveContent(DirectoryInfo dir)
        {
            if (!dir.Exists)
            {
                return;
            }

            LogMessage(MessageType.Info, string.Format(CultureInfo.CurrentCulture, "Removing Content from Folder: {0}", dir.FullName));
            FileSystemInfo[] infos = dir.GetFileSystemInfos("*");
            foreach (FileSystemInfo i in infos)
            {
                // Check to see if this is a DirectoryInfo object.
                if (i is DirectoryInfo)
                {
                    string dirObject = string.Format(CultureInfo.CurrentCulture, "win32_Directory.Name='{0}'", i.FullName);
                    using (ManagementObject mdir = new ManagementObject(dirObject))
                    {
                        mdir.Get();
                        ManagementBaseObject outParams = mdir.InvokeMethod("Delete", null, null);

                        // ReturnValue should be 0, else failure
                        if (outParams != null)
                        {
                            if (Convert.ToInt32(outParams.Properties["ReturnValue"].Value, CultureInfo.CurrentCulture) != 0)
                            {
                                LogMessage(MessageType.Error, string.Format(CultureInfo.CurrentCulture, "Directory deletion error: ReturnValue: {0}", outParams.Properties["ReturnValue"].Value));
                                return;
                            }
                        }
                        else
                        {
                            LogMessage(MessageType.Error, "The ManagementObject call to invoke Delete returned null.");
                            return;
                        }
                    }
                }
                else if (i is FileInfo)
                {
                    // First make sure the file is writable.
                    FileAttributes fileAttributes = System.IO.File.GetAttributes(i.FullName);

                    // If readonly attribute is set, reset it.
                    if ((fileAttributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        System.IO.File.SetAttributes(i.FullName, fileAttributes ^ FileAttributes.ReadOnly);
                    }

                    if (i.Exists)
                    {
                        System.IO.File.Delete(i.FullName);
                    }
                }
            }
        }

        public static void CheckForMenuXmlUpdates()
        {
            LogMessage(MessageType.Info, "Checking for Menu content updates");
            try
            {
                DirectoryInfo installationPath = GetInstallationPath();
                XmlSerializer deserializer = new XmlSerializer(typeof(ButtonConfiguration));

                // Iterate over Add-ins found
                DirectoryInfo buttonRoot = new DirectoryInfo(Path.Combine(installationPath.FullName, "Buttons"));
                DirectoryInfo[] buttons = buttonRoot.GetDirectories();
                foreach (FileInfo file in buttons.Select(button => new FileInfo(Path.Combine(button.FullName, "config.xml"))))
                {
                    // open the configuration file for the button
                    ButtonConfiguration buttonConfig;
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

                        using (var client = GetPreconfiguredWebClient())
                        {
                            // check if we need to update the whole button or just look for structure updates.
                            if (!UpgradeButton(buttonConfig, file))
                            {
                                Stream myStream = client.OpenRead(buttonConfig.onlineUrl);
                                if (myStream != null)
                                {
                                    using (StreamReader sr = new StreamReader(myStream))
                                    {
                                        string latestMenu = sr.ReadToEnd();
                                        if (latestMenu != currentMenu)
                                        {
                                            FileInfo f = new FileInfo(Path.Combine(file.DirectoryName, @"button.xml"));

                                            // First make sure the file is writable.
                                            FileAttributes fileAttributes = File.GetAttributes(f.FullName);

                                            // If readonly attribute is set, reset it.
                                            if ((fileAttributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                                            {
                                                File.SetAttributes(f.FullName, fileAttributes ^ FileAttributes.ReadOnly);
                                            }

                                            using (TextWriter tw = new StreamWriter(f.FullName))
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

                StaticHelper.SetApplicationSetting("LastUpdateCheckDate", DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                LogMessage(MessageType.Error, "Update check failed." + ex.Message);

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

        public static DirectoryInfo GetCachePath()
        {
            DirectoryInfo cachePath = new DirectoryInfo(string.Format(CultureInfo.InvariantCulture, @"{0}\Controlled Vocabulary\Cache", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)));
            if (!cachePath.Exists)
            {
                cachePath.Create();
            }

            return cachePath;
        }

        public static DirectoryInfo GetButtonsPath()
        {
            DirectoryInfo buttonsPath = new DirectoryInfo(string.Format(CultureInfo.InvariantCulture, @"{0}\Controlled Vocabulary\Buttons", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)));
            string paths = string.Empty;
            if (!buttonsPath.Exists)
            {
                string message = string.Format(CultureInfo.InvariantCulture, "Buttons Path not found: {0}", paths);
                LogMessage(MessageType.Error, message);
                throw new ArgumentException(message);
            }

            return buttonsPath;
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

        public static string GetTemplate(string template)
        {
            if (string.IsNullOrEmpty(template))
            {
                LogMessage(MessageType.Error, "GetTemplate received invalid template parameter");
                throw new ArgumentException("GetTemplate received invalid template parameter");
            }

            // First we need to find which Button was clicked
            string[] templateParts = template.Split(new[] { SplitSequence }, StringSplitOptions.RemoveEmptyEntries);

            // Get the installation path
            DirectoryInfo installationPath = GetInstallationPath();

            // Get the image file
            FileInfo f = new FileInfo(Path.Combine(installationPath.FullName, @"Buttons\" + templateParts[0] + @"\Templates\" + templateParts[1] + ".html"));
            if (!f.Exists)
            {
                return string.Empty;
            }

            return System.IO.File.ReadAllText(f.FullName);
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

        public static string GetStandardSuffix(string buttonId)
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

            return buttonConfig.standardSuffix;
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
                    try
                    {
                        menus[i] = (menu)deserializer.Deserialize(buttonStream);
                    }
                    catch (Exception ex)
                    {
                        LogMessage(MessageType.Error, ex.ToString());
                    }
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

        public static string GetApplicationSetting(string settingName)
        {
            XmlDocument xdoc = new XmlDocument();

            // Get the installation path
            DirectoryInfo installationPath = GetInstallationPath();
            xdoc.Load(Path.Combine(installationPath.FullName, "settings.xml"));
            XmlNode node = xdoc.SelectSingleNode(string.Format("/ControlledVocabularySettings/setting[@name='{0}']", settingName));
            if (node == null)
            {
                LogMessage(MessageType.Error, string.Format("SettingName: {0} not found in settings.xml", settingName));
                return string.Empty;
            }

            return node.Attributes["value"].Value;
        }

        public static void SetApplicationSetting(string settingName, string settingValue)
        {
            XmlDocument xdoc = new XmlDocument();
            DirectoryInfo installationPath = GetInstallationPath();
            string fileName = Path.Combine(installationPath.FullName, "settings.xml");
            xdoc.Load(fileName);

            FileAttributes fileAttributes = System.IO.File.GetAttributes(fileName);

            // If readonly attribute is set, reset it.
            if ((fileAttributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                System.IO.File.SetAttributes(fileName, fileAttributes ^ FileAttributes.ReadOnly);
            }

            XmlNode node = xdoc.SelectSingleNode(string.Format("/ControlledVocabularySettings/setting[@name='{0}']", settingName));
            if (node == null)
            {
                // Create a new node.
                XmlElement elem = xdoc.CreateElement("setting");
                elem.SetAttribute("name", settingName);
                elem.SetAttribute("value", settingValue);

                // Add the node to the document.
                XmlElement root = xdoc.DocumentElement;
                root.AppendChild(elem);
            }
            else
            {
                node.Attributes["value"].Value = settingValue;
            }

            xdoc.Save(fileName);
        }

        private static WebClient GetPreconfiguredWebClient()
        {
            return new WebClient { CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore), UseDefaultCredentials = true };
        }
    }
}
