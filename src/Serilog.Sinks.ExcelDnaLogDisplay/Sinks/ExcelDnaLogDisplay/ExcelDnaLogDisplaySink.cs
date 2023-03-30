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

            using (var stringWriter = new StringWriter(new StringBuilder(256)))
            {
                _formatter.Format(logEvent, stringWriter);

                lock (_syncRoot)
                {
                    LogDisplay.DisplayOrder = _displayOrder;
                    LogDisplay.RecordLine(stringWriter.ToString());
                }
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
