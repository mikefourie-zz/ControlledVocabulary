//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Installer1.cs">(c) Controlled Vocabulary on CodePlex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
    using System.Collections;
    using System.ComponentModel;
    using System.Diagnostics;

    /// <summary>
    /// Installer Class
    /// </summary>
    [RunInstaller(true)]
    public partial class Installer1 : System.Configuration.Install.Installer
    {
        public Installer1()
        {
            this.InitializeComponent();
            if (!EventLog.SourceExists("ControlledVocabulary"))
            {
                EventLog.CreateEventSource("ControlledVocabulary", "Application");
            }
        }

        protected override void OnBeforeInstall(IDictionary savedState)
        {
            try
            {
                Process[] processes = Process.GetProcessesByName("OUTLOOK");
                foreach (Process process in processes)
                {
                    process.Kill();
                }
            }
            catch
            {
                // Do Nothing
            }

            base.OnBeforeInstall(savedState);
        }
    }
}
