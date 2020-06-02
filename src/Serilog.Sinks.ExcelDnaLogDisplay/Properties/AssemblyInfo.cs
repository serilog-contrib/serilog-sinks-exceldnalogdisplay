using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("Serilog.Sinks.ExcelDnaLogDisplay")]
[assembly: AssemblyDescription("A Serilog sink that writes events to Excel-DNA LogDisplay")]

#if DEBUG
[assembly: AssemblyConfiguration("Debug")]
#else
[assembly: AssemblyConfiguration("Release")]
#endif

[assembly: AssemblyCompany("augustoproiete.net")]
[assembly: AssemblyProduct("Serilog.Sinks.ExcelDnaLogDisplay")]
[assembly: AssemblyCopyright("Copyright 2018-2020 C. Augusto Proiete & Contributors")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible
// to COM components.  If you need to access a type in this assembly from
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("20e3af43-2c31-47a6-8270-c1b8b4b6c665")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.1.1-beta1")]

[assembly: InternalsVisibleTo("Serilog.Sinks.ExcelDnaLogDisplay.Tests, PublicKey=" +
                              "00240000048000009400000006020000002400005253413100040000010001005db330d3ef1083" +
                              "1fe51df3809c8e717ae5658de73f3a51dd72d7a7b30b49344818c2bc55fde0bfb017f907e7af2b" +
                              "2f507e08707800dca8341ca83722cc79503a5e8449132fce7d81bfa1302fb7f000cd58837ae337" +
                              "b00b9940ec3e433a78c2f04f816843a772f098b667b42e3df91aae44f17b8574892f49576a256b" +
                              "bb13bcd5")]
