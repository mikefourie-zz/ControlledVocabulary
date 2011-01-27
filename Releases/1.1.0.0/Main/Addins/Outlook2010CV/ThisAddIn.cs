//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="ThisAddIn.cs">(c) Controlled Vocabulary on CodePlex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace Outlook2010CV
{
    using ControlledVocabulary;

    /// <summary>
    /// ThisAddIn
    /// </summary>
    public partial class ThisAddIn
    {
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            try
            {
                return new Outlook2010CVAddIn();
            }
            catch (System.Exception ex)
            {
                StaticHelper.LogMessage(MessageType.Error, ex.ToString());
                throw;
            }
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
