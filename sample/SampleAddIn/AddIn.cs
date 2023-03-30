#region Copyright 2018-2023 C. Augusto Proiete & Contributors
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
#endregion

using System;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Logging;
using ExcelDna.Registration;
using Serilog;

namespace SampleAddIn
{
    public class AddIn : ExcelComAddIn, IExcelAddIn
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
                _log.Information("Starting sample Excel-DNA Add-In with Serilog Sink LogDisplay");

                ExcelComAddInHelper.LoadComAddIn(this);

                _log.Verbose("Registering functions");

                ExcelRegistration.GetExcelFunctions()
                    .Select(UpdateFunctionAttributes)
                    .RegisterFunctions();

                _log.Information("Sample Excel-DNA Add-In with Serilog Sink LogDisplay started");
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex);
            }
        }

        public void AutoClose()
        {
            // Do nothing
        }

        public override void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            try
            {
                base.OnDisconnection(disconnectMode, ref custom);

                _log.Information("Stopping sample Excel-DNA Add-In with Serilog Sink LogDisplay");
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex);
            }
            finally
            {
                _log.Information("Sample Excel-DNA Add-In with Serilog Sink LogDisplay stopped");

                Log.CloseAndFlush();
            }
        }

        public static void ProcessUnhandledException(Exception ex, string message = null, [CallerMemberName] string caller = null)
        {
            try
            {
                _log.Error(ex, message ?? $"Unhandled exception on {caller}");
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

            if (ex is TargetInvocationException && !(ex.InnerException is null))
            {
                ProcessUnhandledException(ex.InnerException, message, caller);
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
            ProcessUnhandledException(e.Exception);
        }

        private static void TaskSchedulerUnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
        {
            ProcessUnhandledException(e.Exception);
            e.SetObserved();
        }

        private static void AppDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            ProcessUnhandledException((Exception)e.ExceptionObject);
        }

        private static object ExcelUnhandledException(object ex)
        {
            ProcessUnhandledException((Exception)ex);
            return ex;
        }
    }
}
