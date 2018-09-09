using System;
using System.IO;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;
using Serilog;
using static SampleAddIn.AddIn;

namespace SampleAddIn
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        private ILogger _log = Log.Logger;
        private IRibbonUI _ribbonUi;

        public override string GetCustomUI(string ribbonId)
        {
            try
            {
                _log = Log.ForContext<Ribbon>();
                _log.Debug("Loading ribbon {ribbonId} via GetCustomUI", ribbonId);

                var xllFolderPath = Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? string.Empty;
                var ribbonXmlFilePath = Path.Combine(xllFolderPath, $"{nameof(Ribbon)}.xml");

                if (!File.Exists(ribbonXmlFilePath))
                {
                    throw new FileNotFoundException($"File not found: {ribbonXmlFilePath}", ribbonXmlFilePath);
                }

                var ribbonXml = File.ReadAllText(ribbonXmlFilePath);

                _log.Debug("Ribbon XML {ribbonXml}", ribbonXml);

                return ribbonXml;
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on {nameof(GetCustomUI)}");
                return null;
            }
        }

        public void OnLoad(IRibbonUI ribbonUi)
        {
            try
            {
                _log.Information("Loading Ribbon...");
                _ribbonUi = ribbonUi;
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(OnLoad)}");
            }
        }

        public string SampleTab_GetLabel(IRibbonControl control)
        {
            try
            {
                var ribbonLabel = "Serilog -> LogDisplay";

                // Excel 2013
                if (Math.Abs(ExcelDnaUtil.ExcelVersion - 15.0) < double.Epsilon)
                {
                    ribbonLabel = ribbonLabel.ToUpperInvariant();
                }

                return ribbonLabel;
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(SampleTab_GetLabel)}");
                return control.Id;
            }
        }

        public void VerboseButton_OnAction(IRibbonControl control)
        {
            try
            {
                _log.Verbose("This is a **Verbose** message");
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(VerboseButton_OnAction)}");
            }
        }

        public void DebugButton_OnAction(IRibbonControl control)
        {
            try
            {
                _log.Debug("This is a **Debug** message");
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(DebugButton_OnAction)}");
            }
        }

        public void InformationButton_OnAction(IRibbonControl control)
        {
            try
            {
                _log.Information("This is an **Information** message");
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(InformationButton_OnAction)}");
            }
        }

        public void WarningButton_OnAction(IRibbonControl control)
        {
            try
            {
                _log.Warning("This is a **Warning** message");
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(WarningButton_OnAction)}");
            }
        }

        public void ErrorButton_OnAction(IRibbonControl control)
        {
            try
            {
                _log.Error("This is an **Error** message");
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(ErrorButton_OnAction)}");
            }
        }

        public void ViewLogDisplayButton_OnAction(IRibbonControl control)
        {
            try
            {
                LogDisplay.Show();
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(ViewLogDisplayButton_OnAction)}");
            }
        }

        public void ClearLogDisplayButton_OnAction(IRibbonControl control)
        {
            try
            {
                LogDisplay.Clear();
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(ClearLogDisplayButton_OnAction)}");
            }
        }

        public string DisplayOrderMenu_GetImage(IRibbonControl control)
        {
            try
            {
                switch (LogDisplay.DisplayOrder)
                {
                    case DisplayOrder.NewestLast:
                        return "EndOfDocument";

                    case DisplayOrder.NewestFirst:
                        return "StartOfDocument";

                    default:
                        throw new NotImplementedException(LogDisplay.DisplayOrder.ToString());
                }
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(DisplayOrderMenu_GetImage)}");
                return null;
            }
        }

        public void LogNewestLastButton_OnAction(IRibbonControl control)
        {
            try
            {
                LogDisplay.DisplayOrder = DisplayOrder.NewestLast;
                _ribbonUi.InvalidateControl("DisplayOrderMenu");
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(LogNewestLastButton_OnAction)}");
            }
        }

        public void LogNewestFirstButton_OnAction(IRibbonControl control)
        {
            try
            {
                LogDisplay.DisplayOrder = DisplayOrder.NewestFirst;
                _ribbonUi.InvalidateControl("DisplayOrderMenu");
            }
            catch (Exception ex)
            {
                ProcessUnhandledException(ex, $"Unhandled exception on ${nameof(LogNewestFirstButton_OnAction)}");
            }
        }
    }
}
