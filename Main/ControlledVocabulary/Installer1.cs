//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Installer1.cs">(c) Controlled Vocabulary on Codeplex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
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
    }
}
