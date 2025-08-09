using System;
using System.Linq;
using System.Windows.Forms;

namespace OlapTest
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // Setup global exception handling
            SetupGlobalExceptionHandling();
            
            Application.Run(new Form1());
        }

        /// <summary>
        /// Setup global unhandled exception handling
        /// </summary>
        private static void SetupGlobalExceptionHandling()
        {
            // For exceptions in main UI thread
            Application.ThreadException += Application_ThreadException;
            
            // For exceptions in other threads
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            
            // Setup exception handling mode
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        }

        /// <summary>
        /// Exception handler for main UI thread
        /// </summary>
        private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            LogGlobalException("UI Thread Exception", e.Exception);
        }

        /// <summary>
        /// Exception handler for other threads
        /// </summary>
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            var terminatingInfo = e.IsTerminating ? " (CRITICAL - Application will terminate)" : "";
            LogGlobalException($"Unhandled Exception{terminatingInfo}", exception);
        }

        /// <summary>
        /// Log global exceptions
        /// </summary>
        private static void LogGlobalException(string source, Exception exception)
        {
            try
            {
                // Try to find active form for logging
                Form1.LogGlobalException(source, exception);
            }
            catch
            {
                // If can't log to form, write to Debug
                System.Diagnostics.Debug.WriteLine($"[GLOBAL EXCEPTION] {source}: {exception}");
            }
        }
    }
}