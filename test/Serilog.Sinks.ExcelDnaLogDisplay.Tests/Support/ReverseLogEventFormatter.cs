using System.IO;
using System.Linq;
using Serilog.Events;
using Serilog.Formatting;

namespace Serilog.Sinks.ExcelDnaLogDisplay.Tests.Support
{
    internal class ReverseLogEventFormatter : ITextFormatter
    {
        public void Format(LogEvent logEvent, TextWriter output)
        {
            var reversed = new string(logEvent.RenderMessage().Reverse().ToArray());
            output.Write(reversed);
        }
    }
}
