using System;
using Serilog.Sinks.ExcelDnaLogDisplay.Tests.Support;
using Xunit;

namespace Serilog.Sinks.ExcelDnaLogDisplay.Tests
{
    public class LogDisplayConfigurationExtensionsTests
    {
        [Fact]
        public void When_writing_logging_exceptions_are_suppressed()
        {
            using (var log = new LoggerConfiguration()
                .WriteTo.ExcelDnaLogDisplay(new ThrowingLogEventFormatter())
                .CreateLogger())
            {
                log.Information("Hello");
            }
        }

        [Fact]
        public void When_auditing_logging_exceptions_propagate()
        {
            using (var log = new LoggerConfiguration()
                .AuditTo.ExcelDnaLogDisplay(new ThrowingLogEventFormatter())
                .CreateLogger())
            {
                var ex = Assert.Throws<AggregateException>(() => log.Information("Hello"));
                Assert.IsType<NotImplementedException>(ex.GetBaseException());
            }
        }
    }
}
