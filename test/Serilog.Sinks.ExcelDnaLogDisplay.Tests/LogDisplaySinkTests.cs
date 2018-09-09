using System;
using ExcelDna.Logging;
using Serilog.Events;
using Serilog.Sinks.ExcelDnaLogDisplay.Tests.Support;
using Xunit;

namespace Serilog.Sinks.ExcelDnaLogDisplay.Tests
{
    public class LogDisplaySinkTests : IDisposable
    {
        public void Dispose()
        {
            LogDisplay.Clear();
        }

        [Fact]
        public void Events_are_written_to_LogDisplay()
        {
            var evt = Some.LogEvent("Hello, world!");

            using (var sink = new ExcelDnaLogDisplaySink(Some.TextFormatter()))
            {
                sink.Emit(evt);
            }

            var logEntries = LogDisplayEx.LogStrings;

            Assert.Single(logEntries);
            Assert.Equal("Hello, world!", logEntries[0]);
        }

        [Fact]
        public void By_default_events_are_written_to_LogDisplay_in_NewestLast_DisplayOrder()
        {
            var evt1 = Some.LogEvent("Event1");
            var evt2 = Some.LogEvent("Event2");

            using (var sink = new ExcelDnaLogDisplaySink(Some.TextFormatter()))
            {
                sink.Emit(evt1);
                sink.Emit(evt2);
            }

            var logEntries = LogDisplayEx.LogStrings;

            Assert.True(logEntries.Count == 2);
            Assert.Equal("Event1", logEntries[0]);
            Assert.Equal("Event2", logEntries[1]);
        }

        [Fact]
        public void Events_can_be_written_to_LogDisplay_in_NewestFirst_DisplayOrder()
        {
            var evt1 = Some.LogEvent("Event1");
            var evt2 = Some.LogEvent("Event2");

            using (var sink = new ExcelDnaLogDisplaySink(Some.TextFormatter(), DisplayOrder.NewestFirst))
            {
                sink.Emit(evt1);
                sink.Emit(evt2);
            }

            var logEntries = LogDisplayEx.LogStrings;

            Assert.True(logEntries.Count == 2);
            Assert.Equal("Event2", logEntries[0]);
            Assert.Equal("Event1", logEntries[1]);
        }

        [Fact]
        public void LogDisplay_can_be_cleared_up_if_we_want_when_the_sink_is_disposed()
        {
            var evt = Some.LogEvent("Hello, world!");
            var logEntries = LogDisplayEx.LogStrings;

            using (var sink = new ExcelDnaLogDisplaySink(Some.TextFormatter(), clearLogOnDispose: true))
            {
                sink.Emit(evt);

                Assert.Single(logEntries);
                Assert.Equal("Hello, world!", logEntries[0]);
            }

            Assert.Empty(logEntries);
        }

        [Fact]
        public void LogDisplay_uses_a_circular_list_with_maximum_of_10_000_entries()
        {
            var logEntries = LogDisplayEx.LogStrings;

            using (var sink = new ExcelDnaLogDisplaySink(Some.TextFormatter()))
            {
                var counter = 0;
                LogEvent evt;

                do
                {
                    counter++;

                    evt = Some.LogEvent($"Event {counter:##,###}");
                    sink.Emit(evt);

                } while (counter < 10_000);

                Assert.True(logEntries.Count == 10_000, $"logEntries.Count = {logEntries.Count}, but should be 10,000");

                Assert.Equal("Event 1", logEntries[0]);
                Assert.Equal("Event 10,000", logEntries[9999]);

                evt = Some.LogEvent("Event 10,001");
                sink.Emit(evt);

                Assert.Equal("Event 2", logEntries[0]);
                Assert.Equal("Event 10,000", logEntries[9998]);
                Assert.Equal("Event 10,001", logEntries[9999]);
            }
        }

        [Fact]
        public void LogDisplaySink_uses_ITextFormatter_given_to_it_when_emitting_messages()
        {
            var evt = Some.LogEvent("Hello");

            using (var sink = new ExcelDnaLogDisplaySink(new ReverseLogEventFormatter()))
            {
                sink.Emit(evt);
            }

            var logEntries = LogDisplayEx.LogStrings;

            Assert.Single(logEntries);
            Assert.Equal("olleH", logEntries[0]);
        }
    }
}
