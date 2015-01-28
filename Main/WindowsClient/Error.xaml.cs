//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Error.xaml.cs">(c) Controlled Vocabulary on GitHub, 2015. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
    using System;
    using System.Diagnostics;
    using System.Windows;
    using System.Windows.Input;

    /// <summary>
    /// Interaction logic for Error
    /// </summary>
    public partial class Error : Window
    {
        private Exception exception;

        /// <summary>
        /// Error Constructor
        /// </summary>
        /// <param name="ex">exception</param>
        public Error(Exception ex)
        {
            this.InitializeComponent();
            this.exception = ex;
            this.textBox1.Text = ex.ToString();
        }

        private void label2_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            string mailto = "mailto:" + StaticHelper.GetApplicationSetting("ErrorEmail");
            mailto += "?subject=" + this.exception.Message;
            string inner = string.Empty;
            if (this.exception.InnerException != null)
            {
                inner = " --- " + this.exception.InnerException;
            }

            mailto += "&body=" + this.exception.Message + " ---- " + this.exception.StackTrace + inner;
            if (Convert.ToBoolean(StaticHelper.GetApplicationSetting("CallMailtoProtocol")))
            {
                Process.Start(mailto);
            }

            this.Close();
        }
    }
}
