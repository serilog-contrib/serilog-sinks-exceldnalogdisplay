// Copyright 2018-2020 Caio Proiete & Contributors
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
using System.Threading;
using Serilog.Events;
using Serilog.Formatting.Display;
using Xunit.Sdk;

namespace Serilog.Sinks.ExcelDnaLogDisplay.Tests.Support
{
    internal static class Some
    {
        private static int _counter;

        public static int Int()
        {
            return Interlocked.Increment(ref _counter);
        }

        public static LogEvent LogEvent(string messageTemplate, params object[] propertyValues)
        {
            var log = new LoggerConfiguration().CreateLogger();

#pragma warning disable Serilog004 // Constant MessageTemplate verifier
            if (!log.BindMessageTemplate(messageTemplate, propertyValues, out var template, out var properties))
#pragma warning restore Serilog004 // Constant MessageTemplate verifier
            {
                throw new XunitException("Template could not be bound.");
            }

            return new LogEvent(DateTimeOffset.Now, LogEventLevel.Information, null, template, properties);
        }

        public static MessageTemplateTextFormatter TextFormatter(string outputTemplate = "{Message:lj}",
            IFormatProvider formatProvider = null)
        {
            return new MessageTemplateTextFormatter(outputTemplate, formatProvider);
        }
    }
}
