using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace TelevendFilter
{
    /// <summary>
    /// Logika interakcji dla klasy App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            this.Dispatcher.UnhandledException += App_DispatcherUnhandledException;
        }

        void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            using (StreamWriter FileOut = new StreamWriter("CrashLog.txt"))
            {
                FileOut.Write(e.Exception);
                FileOut.Write(e.Exception.StackTrace.ToString());
                FileOut.Close();
                //log exception and set Handled to true
                e.Handled = true;
                Shutdown(1);
            }
        }
    }
}
