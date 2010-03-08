//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="ThisAddIn.cs">(c) Controlled Vocabulary on Codeplex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace Outlook2010CV
{
    using Office = Microsoft.Office.Core;

    /// <summary>
    /// ThisAddIn
    /// </summary>
    public partial class ThisAddIn
    {
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Outlook2010CVAddIn();
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
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
