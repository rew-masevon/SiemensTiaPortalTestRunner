using Siemens.Engineering;
using Siemens.Engineering.Connection;
using Siemens.Engineering.Download;
using Siemens.Engineering.Download.Configurations;
using Siemens.Engineering.HW;
using Siemens.Engineering.Online;
using Siemens.Engineering.TestSuite;
using Siemens.Engineering.TestSuite.ApplicationTest;
using System;
using System.Linq;

string projectFile = args.Length == 1 ? args[0] : @"<project>";

var tiaPortal = GetOrCreateTiaPortal(projectFile);
var project = GetProject(tiaPortal, projectFile);

DownloadToDevice(project.Devices.First());
RunTest(project);

static void DownloadToDevice(
    Device device,
    string mode = "PN/IE",
    string adapterName = "Siemens PLCSIM Virtual Ethernet Adapter",
    string slotName = "1 X1")
{
    var cpu = device
        .DeviceItems
        .Single(t => t.Classification == DeviceItemClassifications.CPU);

    Download(cpu, mode, adapterName, slotName);
}

static void Download(DeviceItem cpu, string modeName, string adapterName, string slotName)
{
    // Known issues:
    // - Safety must be disabled in order to download.
    // - There's a prompt for a certificate which must be acknowledged once.

    var provider = cpu.GetService<DownloadProvider>();
    var mode = provider.Configuration.Modes.Find(modeName);
    var @interface = mode.PcInterfaces.Single(x => x.Name.Equals(adapterName, StringComparison.OrdinalIgnoreCase));
    var target = @interface.TargetInterfaces.Single(x => x.Name.Equals(slotName, StringComparison.OrdinalIgnoreCase));

    // false should disable communication via tls.
    SetLegacyCommunication(cpu, false, target);

    var result = provider.Download(target, PreDownload, (_) => { }, DownloadOptions.Software);

    if (result.ErrorCount > 0)
    {
        foreach (var error in result.Messages)
        {
            Console.WriteLine(error.Message);
        }

        throw new InvalidOperationException($"Download resulted in {result.ErrorCount} errors.");
    }
}

static void SetLegacyCommunication(DeviceItem cpu, bool enable, ConfigurationTargetInterface target)
{
    var onlineProvider = cpu.GetService<OnlineProvider>();

    Console.WriteLine($"OnlineProvider.IsConfigured: {onlineProvider.Configuration.IsConfigured}");
    Console.WriteLine($"OnlineProvider.EnableLegacyCommunication: {onlineProvider.Configuration.EnableLegacyCommunication}");

    onlineProvider.Configuration.EnableLegacyCommunication = enable;

    var applied = onlineProvider.Configuration.ApplyConfiguration(target);
    Console.WriteLine($"Applied: {applied}");
}

static void PreDownload(DownloadConfiguration downloadConfiguration)
{
}

static void RunTest(Project project)
{
    var testSuiteService = project.GetService<TestSuiteService>();
    var testCaseExecutor = testSuiteService.ApplicationTestGroup.GetService<TestCaseExecutor>();

    var result = testCaseExecutor.Run(testSuiteService.ApplicationTestGroup.TestCases);

    Console.WriteLine($"State: {result.State}, ErrorCount: {result.ErrorCount}");
    foreach (var message in result.Messages)
    {
        Console.WriteLine($"{message.DateTime:O}: {message.Path}: {message.Description} ({message.State})");
    }

}

static TiaPortal GetOrCreateTiaPortal(string projectFile)
    => TiaPortal
           .GetProcesses()
           .FirstOrDefault(t => t.ProjectPath is not null && t.ProjectPath.FullName.Equals(projectFile, StringComparison.OrdinalIgnoreCase))?
           .Attach()
       ?? new TiaPortal(TiaPortalMode.WithUserInterface);

static Project GetProject(TiaPortal portal, string projectFile)
    => portal.Projects.FirstOrDefault(t => t.Path.FullName.Equals(projectFile, StringComparison.OrdinalIgnoreCase))
       ?? portal.Projects.Open(new System.IO.FileInfo(projectFile));