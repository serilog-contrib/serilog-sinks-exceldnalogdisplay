using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelDna.Logging;

namespace Serilog.Sinks.ExcelDnaLogDisplay.Tests.Support
{
    internal static class LogDisplayEx
    {
        public static List<string> LogStrings
        {
            get
            {
                var logStringsField = typeof(LogDisplay).GetField("LogStrings", BindingFlags.NonPublic | BindingFlags.Static);
                if (logStringsField == null)
                {
                    throw new InvalidOperationException("`LogStrings` doesn't exist anymore");
                }

                if (!(logStringsField.GetValue(null) is List<string> logStrings))
                {
                    throw new InvalidOperationException($"`LogStrings` is no longer a `List<string>`... It's {logStringsField.FieldType}");
                }

                return logStrings;
            }
        }
    }
}
