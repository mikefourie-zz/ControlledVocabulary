﻿//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="ThisAddIn.cs">(c) Controlled Vocabulary on CodePlex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace Outlook2007CV
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using ControlledVocabulary;
    using Microsoft.Office.Interop.Outlook;
    using Office = Microsoft.Office.Core;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// ThisAddIn
    /// </summary>
    public partial class ThisAddIn
    {
        private static void Send(Microsoft.Office.Core.CommandBarButton clickedControl, string subject, OlImportance importance)
        {
            try
            {
                // First we need to find which Button was clicked
                string[] idParts = clickedControl.DescriptionText.Split(new[] { StaticHelper.SplitSequence }, StringSplitOptions.RemoveEmptyEntries);

                Application outlookApp = new Outlook.Application();
                MailItem newEmail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

                // Get the recipients
                string[] recipients = StaticHelper.GetRecipients(idParts[0], clickedControl.Id.ToString(CultureInfo.InvariantCulture));
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

                newEmail.Display(true);
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

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                this.AddMenu();
            }
            catch (System.Exception ex)
            {
                StaticHelper.LogMessage(MessageType.Error, ex.ToString());
            }
        }

        private void AddMenu()
        {
            // Define the existing Menu Bar
            Office.CommandBar menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;

            // Add the top level new Menu
            StaticHelper.LogMessage(MessageType.Info, "Adding Controlled Vocab menu");
            Office.CommandBarPopup newMenuBar = (Office.CommandBarPopup)menuBar.Controls.Add(Office.MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            newMenuBar.Caption = "Controlled Vocab";

            // get the buttons
            StaticHelper.LogMessage(MessageType.Info, "Getting buttons");
            menu[] buttons = StaticHelper.GetControlledVocabularyMenus();

            // build the buttons
            StaticHelper.LogMessage(MessageType.Info, "Building menu");
            this.BuildMenu(newMenuBar, buttons);

            StaticHelper.LogMessage(MessageType.Info, "Making menu visible");
            newMenuBar.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void BuildMenu(Office.CommandBarPopup menuBar, IEnumerable<menu> buttons)
        {
            int customButtonPosition = 1;
            foreach (menu customButton in buttons)
            {
                Office.CommandBarPopup addInButton = (Office.CommandBarPopup)menuBar.Controls.Add(Office.MsoControlType.msoControlPopup, this.missing, this.missing, customButtonPosition++, true);
                addInButton.Caption = customButton.label;
                for (int i = 0; i < customButton.Items.Length; i++)
                {
                    if (customButton.Items[i] is menu)
                    {
                        menu categoryMenu = (menu)customButton.Items[i];
                        Office.CommandBarPopup categoryButton = (Office.CommandBarPopup)addInButton.Controls.Add(Office.MsoControlType.msoControlPopup, this.missing, this.missing, i + 1, true);
                        categoryButton.Caption = categoryMenu.label;
                        categoryButton.Tag = categoryMenu.id;
                        
                        // the name of the menu is the name of the file
                        if (categoryMenu.Items.Length > 0)
                        {
                            int j = 0;
                            foreach (var item in categoryMenu.Items)
                            {
                                if (item is button)
                                {
                                    button actionMenu = (button)item;
                                    Office.CommandBarButton actionButton = (Office.CommandBarButton)categoryButton.Controls.Add(Office.MsoControlType.msoControlButton, this.missing, this.missing, j + 1, true);
                                    actionButton.Caption = actionMenu.label;
                                    switch (actionMenu.onAction)
                                    {
                                        case "SendNormal":
                                            actionButton.Click += this.HandleMenuClickNormal;
                                            break;
                                        case "SendHigh":
                                            actionButton.Click += this.HandleMenuClickHigh;
                                            break;
                                        case "SendLow":
                                            actionButton.Click += this.HandleMenuClickLow;
                                            break;
                                        case "Guidance":
                                            actionButton.Click += this.HandleMenuClickGuidance;
                                            break;
                                    }

                                    actionButton.Tag = actionMenu.tag;
                                    actionButton.DescriptionText = actionMenu.id;
                                    j++;
                                }
                            }
                        }
                    }
                    else if (customButton.Items[i] is button)
                    {
                        button actionMenu = (button)customButton.Items[i];
                        Office.CommandBarButton actionButton = (Office.CommandBarButton)addInButton.Controls.Add(Office.MsoControlType.msoControlButton, this.missing, this.missing, i + 1, true);
                        actionButton.Caption = actionMenu.label;
                        switch (actionMenu.onAction)
                        {
                            case "SendNormal":
                                actionButton.Click += this.HandleMenuClickNormal;
                                break;
                            case "SendHigh":
                                actionButton.Click += this.HandleMenuClickHigh;
                                break;
                            case "SendLow":
                                actionButton.Click += this.HandleMenuClickLow;
                                break;
                            case "Guidance":
                                actionButton.Click += this.HandleMenuClickGuidance;
                                break;
                        }

                        actionButton.Tag = actionMenu.tag;
                        actionButton.DescriptionText = actionMenu.id;
                    }
                    else if (customButton.Items[i] is menuSeparator)
                    {
                        menuSeparator sep = (menuSeparator)customButton.Items[i];
                        Office.CommandBarButton actionButton = (Office.CommandBarButton)addInButton.Controls.Add(Office.MsoControlType.msoControlButton, this.missing, this.missing, i + 1, true);
                        actionButton.BeginGroup = true;
                        actionButton.Caption = sep.title;
                        if (string.IsNullOrEmpty(sep.title))
                        {
                            actionButton.Visible = false;
                        }

                        actionButton.Enabled = false;
                    }
                }
            }
        }

        private void HandleMenuClickGuidance(Microsoft.Office.Core.CommandBarButton clickedControl, ref bool cancelDefault)
        {
            string[] idParts = clickedControl.DescriptionText.Split(new[] { StaticHelper.SplitSequence }, StringSplitOptions.RemoveEmptyEntries);
            string guidanceUrl = StaticHelper.GetGuidanceUrl(idParts[0]);
            System.Diagnostics.Process.Start(guidanceUrl);
        }

        private void HandleMenuClickNormal(Microsoft.Office.Core.CommandBarButton clickedControl, ref bool cancelDefault)
        {
            Send(clickedControl, clickedControl.Tag, OlImportance.olImportanceNormal);
        }

        private void HandleMenuClickHigh(Microsoft.Office.Core.CommandBarButton clickedControl, ref bool cancelDefault)
        {
            Send(clickedControl, clickedControl.Tag, OlImportance.olImportanceHigh);
        }

        private void HandleMenuClickLow(Microsoft.Office.Core.CommandBarButton clickedControl, ref bool cancelDefault)
        {
            Send(clickedControl, clickedControl.Tag, OlImportance.olImportanceLow);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            try
            {
                this.Startup += this.ThisAddIn_Startup;
                this.Shutdown += this.ThisAddIn_Shutdown;
            }
            catch (System.Exception ex)
            {
                StaticHelper.LogMessage(MessageType.Error, ex.ToString());
            }
        }
        
        #endregion
    }
}