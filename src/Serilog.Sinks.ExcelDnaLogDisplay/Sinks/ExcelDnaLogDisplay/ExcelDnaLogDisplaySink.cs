using System;
using System.ComponentModel;
using System.IO;
using System.Text;
using ExcelDna.Logging;
using Serilog.Core;
using Serilog.Events;
using Serilog.Formatting;

namespace Serilog.Sinks.ExcelDnaLogDisplay
{
    internal class ExcelDnaLogDisplaySink : ILogEventSink, IDisposable
    {
        private readonly ITextFormatter _formatter;
        private readonly DisplayOrder _displayOrder;
        private readonly bool _clearLogOnDispose;

        private readonly object _syncRoot = new object();
        private bool _disposed;

        public ExcelDnaLogDisplaySink(ITextFormatter formatter, DisplayOrder displayOrder = DisplayOrder.NewestLast, bool clearLogOnDispose = false)
        {
            if (!Enum.IsDefined(typeof(DisplayOrder), displayOrder))
                throw new InvalidEnumArgumentException(nameof(displayOrder), (int) displayOrder, typeof(DisplayOrder));

            _formatter = formatter ?? throw new ArgumentNullException(nameof(formatter));
            _displayOrder = displayOrder;
            _clearLogOnDispose = clearLogOnDispose;
        }

        public void Emit(LogEvent logEvent)
        {
            EnsureNotDisposed();

            var stringWriter = new StringWriter(new StringBuilder(256));
            _formatter.Format(logEvent, stringWriter);

            lock (_syncRoot)
            {
                LogDisplay.DisplayOrder = _displayOrder;
                LogDisplay.RecordLine(stringWriter.ToString());
            }
        }

        public void Dispose()
        {
            if (_disposed) return;

            if (_clearLogOnDispose)
            {
                lock (_syncRoot)
                {
                    LogDisplay.Clear();
                }
            }

            _disposed = true;
        }

        private void EnsureNotDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ExcelDnaLogDisplaySink));
            }
        }
    }
}
