// Copyright 2018-2020 C. Augusto Proiete & Contributors
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

using System;
using System.ComponentModel;
using Serilog.Configuration;
using Serilog.Core;
using Serilog.Events;
using Serilog.Formatting;
using Serilog.Formatting.Display;
using ExcelDna.Logging;
using Serilog.Sinks.ExcelDnaLogDisplay;

// ReSharper disable once CheckNamespace
namespace Serilog
{
    /// <summary>
    /// Adds the WriteTo.ExcelDnaLogDisplay() extension method to <see cref="T:Serilog.LoggerConfiguration" />.
    /// </summary>
    public static class ExcelDnaLogDisplayConfigurationExtensions
    {
        private const string _defaultOutputTemplate =
            "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}";

        /// <summary>
        /// Writes log events to <see cref="T:ExcelDna.Logging.LogDisplay" />.
        /// </summary>
        /// <param name="sinkConfiguration">Logger sink configuration.</param>
        /// <param name="restrictedToMinimumLevel">The minimum level for
        /// events passed through the sink. Ignored when <paramref name="levelSwitch" /> is specified.</param>
        /// <param name="levelSwitch">A switch allowing the pass-through minimum level
        /// to be changed at runtime.</param>
        /// <param name="outputTemplate">A message template describing the format used to write to the sink.
        /// the default is <code>"{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}"</code>.</param>
        /// <param name="formatProvider">Supplies culture-specific formatting information, or null.</param>
        /// <param name="displayOrder">The order in which messages are displayed in the log viewer.
        /// The default is <see cref="F:ExcelDna.Logging.NewestLast" />.</param>
        /// <param name="clearLogOnDispose">Calls <see cref="LogDisplay.Clear()" /> when the sink is disposed.
        /// The default is <see langword="false" />.</param>
        /// <returns>Configuration object allowing method chaining.</returns>
        public static LoggerConfiguration ExcelDnaLogDisplay(this LoggerSinkConfiguration sinkConfiguration,
            LogEventLevel restrictedToMinimumLevel = LogEventLevel.Verbose,
            string outputTemplate = _defaultOutputTemplate,
            IFormatProvider formatProvider = null, LoggingLevelSwitch levelSwitch = null,
            DisplayOrder displayOrder = DisplayOrder.NewestLast, bool clearLogOnDispose = false)
        {
            if (sinkConfiguration == null) throw new ArgumentNullException(nameof(sinkConfiguration));

            if (!Enum.IsDefined(typeof(LogEventLevel), restrictedToMinimumLevel))
                throw new InvalidEnumArgumentException(nameof(restrictedToMinimumLevel), (int) restrictedToMinimumLevel,
                    typeof(LogEventLevel));

            if (!Enum.IsDefined(typeof(DisplayOrder), displayOrder))
                throw new InvalidEnumArgumentException(nameof(displayOrder), (int) displayOrder, typeof(DisplayOrder));

            var templateTextFormatter = new MessageTemplateTextFormatter(outputTemplate, formatProvider);

            return sinkConfiguration.Sink(new ExcelDnaLogDisplaySink(templateTextFormatter, displayOrder, clearLogOnDispose), restrictedToMinimumLevel, levelSwitch);
        }

        /// <summary>
        /// Writes log events to <see cref="T:ExcelDna.Logging.LogDisplay" />.
        /// </summary>
        /// <param name="sinkConfiguration">Logger sink configuration.</param>
        /// <param name="formatter">Controls the rendering of log events into text, for example to log JSON. To
        /// control plain text formatting, use the overload that accepts an output template.</param>
        /// <param name="restrictedToMinimumLevel">The minimum level for
        /// events passed through the sink. Ignored when <paramref name="levelSwitch" /> is specified.</param>
        /// <param name="levelSwitch">A switch allowing the pass-through minimum level
        /// to be changed at runtime.</param>
        /// <param name="displayOrder">The order in which messages are displayed in the log viewer.
        /// The default is <see cref="F:ExcelDna.Logging.NewestLast" />.</param>
        /// <param name="clearLogOnDispose">Calls <see cref="LogDisplay.Clear()" /> when the sink is disposed.
        /// The default is <see langword="false" />.</param>
        /// <returns>Configuration object allowing method chaining.</returns>
        public static LoggerConfiguration ExcelDnaLogDisplay(this LoggerSinkConfiguration sinkConfiguration,
            ITextFormatter formatter, LogEventLevel restrictedToMinimumLevel = LogEventLevel.Verbose,
            LoggingLevelSwitch levelSwitch = null, DisplayOrder displayOrder = DisplayOrder.NewestLast, bool clearLogOnDispose = false)
        {
            if (sinkConfiguration == null) throw new ArgumentNullException(nameof(sinkConfiguration));
            if (formatter == null) throw new ArgumentNullException(nameof(formatter));

            if (!Enum.IsDefined(typeof(LogEventLevel), restrictedToMinimumLevel))
                throw new InvalidEnumArgumentException(nameof(restrictedToMinimumLevel), (int) restrictedToMinimumLevel,
                    typeof(LogEventLevel));

            if (!Enum.IsDefined(typeof(DisplayOrder), displayOrder))
                throw new InvalidEnumArgumentException(nameof(displayOrder), (int) displayOrder, typeof(DisplayOrder));

            return sinkConfiguration.Sink(new ExcelDnaLogDisplaySink(formatter, displayOrder, clearLogOnDispose), restrictedToMinimumLevel, levelSwitch);
        }

        /// <summary>
        /// Audits log events to <see cref="T:ExcelDna.Logging.LogDisplay" />.
        /// </summary>
        /// <param name="sinkConfiguration">Audit Logger sink configuration.</param>
        /// <param name="restrictedToMinimumLevel">The minimum level for
        /// events passed through the sink. Ignored when <paramref name="levelSwitch" /> is specified.</param>
        /// <param name="levelSwitch">A switch allowing the pass-through minimum level
        /// to be changed at runtime.</param>
        /// <param name="outputTemplate">A message template describing the format used to write to the sink.
        /// the default is <code>"{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}"</code>.</param>
        /// <param name="formatProvider">Supplies culture-specific formatting information, or null.</param>
        /// <param name="displayOrder">The order in which messages are displayed in the log viewer.
        /// The default is <see cref="F:ExcelDna.Logging.NewestLast" />.</param>
        /// <param name="clearLogOnDispose">Calls <see cref="LogDisplay.Clear()" /> when the sink is disposed.
        /// The default is <see langword="false" />.</param>
        /// <returns>Configuration object allowing method chaining.</returns>
        public static LoggerConfiguration ExcelDnaLogDisplay(this LoggerAuditSinkConfiguration sinkConfiguration,
            LogEventLevel restrictedToMinimumLevel = LogEventLevel.Verbose,
            string outputTemplate = _defaultOutputTemplate,
            IFormatProvider formatProvider = null, LoggingLevelSwitch levelSwitch = null,
            DisplayOrder displayOrder = DisplayOrder.NewestLast, bool clearLogOnDispose = false)
        {
            if (sinkConfiguration == null) throw new ArgumentNullException(nameof(sinkConfiguration));

            if (!Enum.IsDefined(typeof(LogEventLevel), restrictedToMinimumLevel))
                throw new InvalidEnumArgumentException(nameof(restrictedToMinimumLevel), (int) restrictedToMinimumLevel,
                    typeof(LogEventLevel));

            if (!Enum.IsDefined(typeof(DisplayOrder), displayOrder))
                throw new InvalidEnumArgumentException(nameof(displayOrder), (int) displayOrder, typeof(DisplayOrder));

            var templateTextFormatter = new MessageTemplateTextFormatter(outputTemplate, formatProvider);

            return sinkConfiguration.Sink(new ExcelDnaLogDisplaySink(templateTextFormatter, displayOrder, clearLogOnDispose), restrictedToMinimumLevel, levelSwitch);
        }

        /// <summary>
        /// Audits log events to <see cref="T:ExcelDna.Logging.LogDisplay" />.
        /// </summary>
        /// <param name="sinkConfiguration">Audit Logger sink configuration.</param>
        /// <param name="formatter">Controls the rendering of log events into text, for example to log JSON. To
        /// control plain text formatting, use the overload that accepts an output template.</param>
        /// <param name="restrictedToMinimumLevel">The minimum level for
        /// events passed through the sink. Ignored when <paramref name="levelSwitch" /> is specified.</param>
        /// <param name="levelSwitch">A switch allowing the pass-through minimum level
        /// to be changed at runtime.</param>
        /// <param name="displayOrder">The order in which messages are displayed in the log viewer.
        /// The default is <see cref="F:ExcelDna.Logging.NewestLast" />.</param>
        /// <param name="clearLogOnDispose">Calls <see cref="LogDisplay.Clear()" /> when the sink is disposed.
        /// The default is <see langword="false" />.</param>
        /// <returns>Configuration object allowing method chaining.</returns>
        public static LoggerConfiguration ExcelDnaLogDisplay(this LoggerAuditSinkConfiguration sinkConfiguration,
            ITextFormatter formatter, LogEventLevel restrictedToMinimumLevel = LogEventLevel.Verbose,
            LoggingLevelSwitch levelSwitch = null, DisplayOrder displayOrder = DisplayOrder.NewestLast, bool clearLogOnDispose = false)
        {
            if (sinkConfiguration == null) throw new ArgumentNullException(nameof(sinkConfiguration));
            if (formatter == null) throw new ArgumentNullException(nameof(formatter));

            if (!Enum.IsDefined(typeof(LogEventLevel), restrictedToMinimumLevel))
                throw new InvalidEnumArgumentException(nameof(restrictedToMinimumLevel), (int) restrictedToMinimumLevel,
                    typeof(LogEventLevel));

            if (!Enum.IsDefined(typeof(DisplayOrder), displayOrder))
                throw new InvalidEnumArgumentException(nameof(displayOrder), (int) displayOrder, typeof(DisplayOrder));

            return sinkConfiguration.Sink(new ExcelDnaLogDisplaySink(formatter, displayOrder, clearLogOnDispose), restrictedToMinimumLevel, levelSwitch);
        }
    }
}
