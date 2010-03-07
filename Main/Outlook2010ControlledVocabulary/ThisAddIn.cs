//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="ThisAddIn.cs">(c) Controlled Vocabulary on Codeplex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace Outlook2010ControlledVocabulary
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Xml.Linq;
    using Office = Microsoft.Office.Core;
    using Outlook = Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// ThisAddIn
    /// </summary>
    public partial class ThisAddIn
    {
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Outlook2010ControlledVocabularyAddIn();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += this.ThisAddIn_Startup;
            this.Shutdown += this.ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
