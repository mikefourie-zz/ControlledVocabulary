//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="App.xaml.cs">(c) Controlled Vocabulary on GitHub, 2015. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
    using System;
    using System.Windows;

    /// <summary>
    /// Interaction logic for App
    /// </summary>
    public partial class App : Application
    {
        private static void MyHandler(object sender, UnhandledExceptionEventArgs args)
        {
            Exception e = (Exception)args.ExceptionObject;
            Error errorWindow = new Error(e);
            errorWindow.ShowDialog();
            Environment.Exit(-1);
        }

        private void App_StartUp(object sender, StartupEventArgs e)
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += MyHandler;

            if (e.Args.Length > 0)
            {
                Manager managerWindow = new Manager(e.Args[0]);
                managerWindow.ShowDialog();
            }
        }
    }
}
