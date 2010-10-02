//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Outlook2010CVAddIn.cs">(c) Controlled Vocabulary on CodePlex, 2010. All other rights reserved.</copyright>
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
            StaticHelper.LogMessage(MessageType.Info, "Getting Custom UI");
            try
            {
                if (!string.IsNullOrEmpty(this.cachedRibbon))
                {
                    return this.cachedRibbon;
                }

                StaticHelper.LogMessage(MessageType.Info, "Checking for Updates");
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

                // now for a bit of a hack. need to move to objects in a future release
                string readXml = buttonXml.ToString();
                Regex regEx = new Regex(" toRecipients=\"[^\"]*\"");
                readXml = regEx.Replace(readXml, string.Empty);
                regEx = new Regex(" ccRecipients=\"[^\"]*\"");
                readXml = regEx.Replace(readXml, string.Empty);
                regEx = new Regex(" bccRecipients=\"[^\"]*\"");
                readXml = regEx.Replace(readXml, string.Empty);

                // Inject the Add-ins using regular expression
                regEx = new Regex("DLBUTTONPLACHOLDER_DONOTREMOVE");
                this.cachedRibbon = regEx.Replace(ribbonXml, readXml);
                StaticHelper.LogMessage(MessageType.Info, "Ribbon = " + this.cachedRibbon);

                return this.cachedRibbon;
            }
            catch (System.Exception ex)
            {
                StaticHelper.LogMessage(MessageType.Error, ex.ToString());
                throw;
            }
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
        }

        public void Guidance(IRibbonControl control)
        {
            string[] idParts = control.Id.Split(new[] { StaticHelper.SplitSequence }, StringSplitOptions.RemoveEmptyEntries);
            string guidanceUrl = StaticHelper.GetGuidanceUrl(idParts[0]);
            System.Diagnostics.Process.Start(guidanceUrl);
        }

        public void SendNormal(IRibbonControl control)
        {
            this.Send(control, control.Tag, OlImportance.olImportanceNormal);
        }

        public void SendHigh(IRibbonControl control)
        {
            this.Send(control, control.Tag, OlImportance.olImportanceHigh);
        }

        public void SendLow(IRibbonControl control)
        {
            this.Send(control, control.Tag, OlImportance.olImportanceLow);
        }

        public void Send(IRibbonControl control, string subject, OlImportance importance)
        {
            try
            {
                string[] idParts = control.Id.Split(new[] { StaticHelper.SplitSequence }, StringSplitOptions.RemoveEmptyEntries);
                Application outlookApp = new ApplicationClass();
                MailItem newEmail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

                // Get the recipients
                string[] recipients = StaticHelper.GetRecipients(idParts[0], control.Id);
                newEmail.To = recipients[0];
                newEmail.CC = recipients[1];
                newEmail.BCC = recipients[2];
                newEmail.Subject = subject;
                newEmail.Importance = importance;

                string from = StaticHelper.GetFromAccount(idParts[0]);
                if (!string.IsNullOrEmpty(from))
                {
                    // Retrieve the account that has the specific SMTP address.
                    Account account = GetAccountForEmailAddress(outlookApp, from);
                    if (account != null)
                    {
                        // Use this account to send the e-mail.
                        newEmail.SendUsingAccount = account;
                    }
                }
                
                newEmail.Display();
            }
            catch (System.Exception ex)
            {
                StaticHelper.LogMessage(MessageType.Error, ex.ToString());
                throw;
            }
        }

        public void MeetingNormal(IRibbonControl control)
        {
            this.Meeting(control, control.Tag, OlImportance.olImportanceNormal);
        }

        public void MeetingHigh(IRibbonControl control)
        {
            this.Meeting(control, control.Tag, OlImportance.olImportanceHigh);
        }

        public void MeetingLow(IRibbonControl control)
        {
            this.Meeting(control, control.Tag, OlImportance.olImportanceLow);
        }

        public void Meeting(IRibbonControl control, string subject, OlImportance importance)
        {
            try
            {
                string[] idParts = control.Id.Split(new[] { StaticHelper.SplitSequence }, StringSplitOptions.RemoveEmptyEntries);
                Application outlookApp = new ApplicationClass();
                AppointmentItem newMeeting = (AppointmentItem)outlookApp.CreateItem(OlItemType.olAppointmentItem);
                newMeeting.MeetingStatus = OlMeetingStatus.olMeeting;

                // Get the recipients
                string[] recipients = StaticHelper.GetRecipients(idParts[0], control.Id);
                if (!string.IsNullOrEmpty(recipients[1]))
                {
                    Recipient recipRequired = newMeeting.Recipients.Add(recipients[0]);
                    recipRequired.Type = (int)OlMeetingRecipientType.olRequired;
                }

                if (!string.IsNullOrEmpty(recipients[1]))
                {
                    Recipient recipOptional = newMeeting.Recipients.Add(recipients[1]);
                    recipOptional.Type = (int)OlMeetingRecipientType.olOptional;
                }

                newMeeting.Subject = subject;
                newMeeting.Importance = importance;

                string from = StaticHelper.GetFromAccount(idParts[0]);
                if (!string.IsNullOrEmpty(from))
                {
                    // Retrieve the account that has the specific SMTP address.
                    Account account = GetAccountForEmailAddress(outlookApp, from);
                    if (account != null)
                    {
                        // Use this account to send the e-mail.
                        newMeeting.SendUsingAccount = account;
                    }
                }

                newMeeting.Display();
            }
            catch (System.Exception ex)
            {
                StaticHelper.LogMessage(MessageType.Error, ex.ToString());
                throw;
            }
        }

        private static Account GetAccountForEmailAddress(Application application, string smtpAddress)
        {
            // Loop over the Accounts collection of the current Outlook session.
            Accounts accounts = application.Session.Accounts;
            foreach (Account account in accounts)
            {
                // When the e-mail address matches, return the account.
                if (account.SmtpAddress == smtpAddress)
                {
                    return account;
                }
            }

            StaticHelper.LogMessage(MessageType.Error, string.Format(CultureInfo.InstalledUICulture, "No Account with SmtpAddress: {0} exists!", smtpAddress));
            return null;
        }
    }
}