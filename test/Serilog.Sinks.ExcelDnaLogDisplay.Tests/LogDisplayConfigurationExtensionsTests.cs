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
