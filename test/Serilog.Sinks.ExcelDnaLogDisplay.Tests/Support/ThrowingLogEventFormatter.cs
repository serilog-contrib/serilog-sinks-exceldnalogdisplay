using System;
using System.IO;
using Serilog.Events;
using Serilog.Formatting;

namespace Serilog.Sinks.ExcelDnaLogDisplay.Tests.Support
{
    internal class ThrowingLogEventFormatter : ITextFormatter
    {
        public void Format(LogEvent logEvent, TextWriter output)
        {
            throw new NotImplementedException();
        }
    }
}
