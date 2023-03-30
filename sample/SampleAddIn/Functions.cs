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
