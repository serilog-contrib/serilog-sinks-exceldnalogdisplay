using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Logging;
using ExcelDna.Registration;
using Serilog;

namespace SampleAddIn
{
    public class AddIn : IExcelAddIn
    {
        private static ILogger _log = Log.Logger;

        public void AutoOpen()
        {
            try
            {
                Application.ThreadException += ApplicationThreadUnhandledException;
                AppDomain.CurrentDomain.UnhandledException += AppDomainUnhandledException;
                TaskScheduler.UnobservedTaskException += TaskSchedulerUnobservedTaskException;
                ExcelIntegration.RegisterUnhandledExceptionHandler(ExcelUnhandledException);

                _log = Log.Logger = ConfigureLogging();

                _log.Information("Loading sample Excel-DNA Add-In with Serilog Sink LogDisplay");

                _log.Verbose("Registering functions");

                ExcelRegistration.GetExcelFunctions()
                    .Select(UpdateFunctionAttributes)
                    .RegisterFunctions();
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, nameof(AutoOpen));
            }
        }

        public void AutoClose()
        {
            try
            {
                Log.CloseAndFlush();
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, nameof(AutoClose));
            }
        }

        public static void ProcessUnhandledException(Exception ex, string message)
        {
            try
            {
                _log.Error(ex, message);
            }
            catch (Exception lex)
            {
                try
                {
                    Serilog.Debugging.SelfLog.WriteLine(lex.ToString());
                }
                catch
                {
                    // Do nothing...
                }
            }

            if (ex.InnerException != null)
            {
                ProcessUnhandledException(ex.InnerException, message);
                return;
            }

#if DEBUG
            MessageBox.Show(ex.ToString(), "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#else
            const string errorMessage = "An unexpected error ocurred. Please try again in a few minutes, and if the error persists, contact support";
            MessageBox.Show(errorMessage, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
        }

        private static ILogger ConfigureLogging()
        {
            AppDomain.CurrentDomain.ProcessExit += (sender, args) => Log.CloseAndFlush();

            return new LoggerConfiguration()
                .MinimumLevel.Verbose()
                .WriteTo.ExcelDnaLogDisplay(displayOrder: DisplayOrder.NewestFirst)
                .CreateLogger();
        }

        private static ExcelFunctionRegistration UpdateFunctionAttributes(ExcelFunctionRegistration excelFunction)
        {
            excelFunction.FunctionAttribute.Name = excelFunction.FunctionAttribute.Name.ToUpperInvariant();
            return excelFunction;
        }

        private static void ApplicationThreadUnhandledException(object sender, ThreadExceptionEventArgs e)
        {
            ProcessUnhandledException(e.Exception, nameof(ApplicationThreadUnhandledException));
        }

        private static void TaskSchedulerUnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
        {
            ProcessUnhandledException(e.Exception, nameof(TaskSchedulerUnobservedTaskException));
            e.SetObserved();
        }

        private static void AppDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            ProcessUnhandledException((Exception)e.ExceptionObject, nameof(AppDomainUnhandledException));
        }

        private static object ExcelUnhandledException(object ex)
        {
            ProcessUnhandledException((Exception)ex, nameof(AppDomainUnhandledException));
            return ex;
        }
    }
}
