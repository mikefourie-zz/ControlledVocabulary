//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Outlook2010CVAddIn.cs">(c) Controlled Vocabulary on Codeplex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace Outlook2010CV
{
    using System;
    using System.Drawing;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using ControlledVocabulary;
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Outlook;
    using Application = Microsoft.Office.Interop.Outlook.Application;
    using Office = Microsoft.Office.Core;
    using Outlook = Microsoft.Office.Interop.Outlook;

    [ComVisible(true)]
    public class Outlook2010CVAddIn : IRibbonExtensibility
    {
        private string cachedRibbon;

        public Bitmap LoadImages(string image)
        {
            return StaticHelper.GetImage(image);
        }

        // TODO: Bug how do get this thing to only show in the main form and not create phantom windows
        ////public bool GetVisible(Office.IRibbonControl control)
        ////{
        ////    // only show the tab in the mai Outlook Explorer view
        ////    return control.Context is Outlook.Explorer;
        ////}

        public string GetCustomUI(string ribbonID)
        {
            if (!string.IsNullOrEmpty(this.cachedRibbon))
            {
                return this.cachedRibbon;    
            }

            StaticHelper.CheckForUpdates();

            // Get the installation path
            DirectoryInfo installationPath = StaticHelper.GetInstallationPath();

            // Get the Outlook2010.xml file
            FileInfo f = new FileInfo(Path.Combine(installationPath.FullName, @"Templates\Outlook2010.xml"));
            if (!f.Exists)
            {
                string message = string.Format(CultureInfo.InvariantCulture, "File not found: {0}", f.FullName);
                StaticHelper.LogMessage(MessageType.Error, message);
                throw new ArgumentException(message);
            }

            string ribbonXml;
            using (TextReader tr = new StreamReader(f.FullName))
            {
                ribbonXml = tr.ReadToEnd();
            }

            // Iterate over Add-ins found
            DirectoryInfo buttonRoot = new DirectoryInfo(Path.Combine(installationPath.FullName, "Buttons"));
            DirectoryInfo[] buttons = buttonRoot.GetDirectories();
            StringBuilder buttonXml = new StringBuilder();
            foreach (FileInfo file in buttons.Select(button => new FileInfo(Path.Combine(button.FullName, "button.xml"))))
            {
                if (!file.Exists)
                {
                    StaticHelper.LogMessage(MessageType.Error, string.Format(CultureInfo.InvariantCulture, "File not found: {0}", file.FullName));
                    continue;
                }

                using (TextReader tr = new StreamReader(file.FullName))
                {
                    buttonXml.Append(tr.ReadToEnd());
                }
            }

            // Inject the Add-ins using regular expression
            Regex regEx = new Regex("DLBUTTONPLACHOLDER_DONOTREMOVE");
            this.cachedRibbon = regEx.Replace(ribbonXml, buttonXml.ToString());

            return this.cachedRibbon;
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            // ThisAddIn.ribbon = ribbonUI;
        }

        public void SendNormal(IRibbonControl control)
        {
            // First we need to find which Button was clicked
            string[] idParts = control.Id.Split(new[] { StaticHelper.SplitSequence }, StringSplitOptions.RemoveEmptyEntries);
            this.Send(idParts[0], control.Tag, OlImportance.olImportanceNormal);
        }

        public void SendHigh(IRibbonControl control)
        {
            string[] idParts = control.Id.Split(new[] { StaticHelper.SplitSequence }, StringSplitOptions.RemoveEmptyEntries);
            this.Send(idParts[0], control.Tag, OlImportance.olImportanceHigh);
        }

        public void SendLow(IRibbonControl control)
        {
            string[] idParts = control.Id.Split(new[] { StaticHelper.SplitSequence }, StringSplitOptions.RemoveEmptyEntries);
            this.Send(idParts[0], control.Tag, OlImportance.olImportanceLow);
        }

        public void Send(string buttonId, string subject, OlImportance importance)
        {
            Application outlookApp = new ApplicationClass();
            MailItem newEmail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

            // Get the recipients
            string[] recipients = StaticHelper.GetRecipients(buttonId);
            newEmail.To = recipients[0];
            newEmail.CC = recipients[1];
            newEmail.BCC = recipients[2];

            newEmail.Subject = subject;
            newEmail.Importance = importance;
            newEmail.Display(true);
        }
    }
}