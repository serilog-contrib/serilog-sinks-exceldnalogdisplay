using ExcelDna.Integration;
using Serilog;

namespace SampleAddIn
{
    public static class Functions
    {
        private static readonly ILogger _log = Log.Logger.ForContext(typeof(Functions));

        [ExcelFunction(Category = "Serilog", Description = "Writes a `Verbose` message to the LogDisplay via Serilog")]
        public static string LogVerbose(string message)
        {
            _log.Verbose(message);
            return $"'[VRB] {message}' written to the log";
        }

        [ExcelFunction(Category = "Serilog", Description = "Writes a `Debug` message to the LogDisplay via Serilog")]
        public static string LogDebug(string message)
        {
            _log.Debug(message);
            return $"'[DBG] {message}' written to the log";
        }

        [ExcelFunction(Category = "Serilog", Description = "Writes an `Information` message to the LogDisplay via Serilog")]
        public static string LogInformation(string message)
        {
            _log.Information(message);
            return $"'[INF] {message}' written to the log";
        }

        [ExcelFunction(Category = "Serilog", Description = "Writes a `Warning` message to the LogDisplay via Serilog")]
        public static string LogWarning(string message)
        {
            _log.Warning(message);
            return $"'[WRN] {message}' written to the log";
        }

        [ExcelFunction(Category = "Serilog", Description = "Writes an `Error` message to the LogDisplay via Serilog")]
        public static string LogError(string message)
        {
            _log.Error(message);
            return $"'[ERR] {message}' written to the log";
        }
    }
}
