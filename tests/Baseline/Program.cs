using System.Diagnostics;
using StatisticaCli;

if (args is not [var path])
{
    Console.Error.WriteLine("Usage: dotnet run --project tests/Baseline -- <published statistica-cli.exe>");
    return 2;
}

string executable = Path.GetFullPath(path);
int[] initial = StaProcesses.Baseline();
var probe = await WorkerRunner.RunWithFallbackAsync(executable,
    new(Backend.Auto, true, initial, null, null, WorkerOperation.Probe), 30, null, verbose: false);
Require(probe.Failure is null, probe.Failure?.Format() ?? "");
var result = probe.Result?.Probe ?? throw new InvalidDataException("Missing probe response.");
Require(result is { Available: true, Com: "Statistica.Application", Host: { Pwsh: null, X64: true, Apartment: "STA" } },
    "Unexpected probe response.");
Require(result.Components is not null &&
    result.Components["Application"] is { Available: true, Pid: int appPid } && appPid != probe.Exit.WorkerPid &&
    result.Components["Library"] is { Available: true, Pid: int libPid } && libPid == probe.Exit.WorkerPid &&
    result.Components.Values.All(component => component.Registration.Any(item => item.View == "Registry64" && item.Server is not null)),
    "Expected working out-of-process and in-process backends with x64 registrations.");
Require(result.Components!["Application"].Pid is int normalPid && WaitForExit(normalPid), "Application survived normal cleanup.");
Console.WriteLine("Detailed probe, backend PIDs and normal cleanup passed.");

foreach (Backend backend in new[] { Backend.Library, Backend.Application })
{
    var request = new WorkerRequest(backend, false, StaProcesses.Baseline(), null, null, WorkerOperation.VerifyAutomation);
    var run = await WorkerRunner.RunAsync(executable, request, 30);
    Require(run.Failure is null, run.Failure?.Format() ?? "");
    Require(run.Result?.Probe is { Available: true }, "Missing verification result.");
    Console.WriteLine($"{backend} SAFEARRAY, indexed properties, Plot.DisplayLine and dispatch errors passed.");
}

var hold = new WorkerRequest(Backend.Application, false, StaProcesses.Baseline(), null, null, WorkerOperation.HoldServer);
var timed = await WorkerRunner.RunAsync(executable, hold, 10);
Require(timed.Failure is { Status: "timed_out", Phase: "Verify", Operation: "HoldServer" },
    $"Unexpected timeout status: {timed.Failure?.Format()}.");
var server = timed.Events.Single(entry => entry is { Event: "server", Backend: "Application" });
Require(server.Pid is int pid && timed.Exit.EndedServerPids.SequenceEqual([pid]) && WaitForExit(pid),
    "Timed-out worker did not terminate its verified Application server.");
Require(WaitForExit(timed.Exit.WorkerPid), "The timed-out worker survived cleanup.");
Require(StaProcesses.Baseline().Except(initial).Count() == 0, "Unexpected STATISTICA process remains after the baseline tests.");
Console.WriteLine("Worker timeout and verified Application cleanup passed.");
return 0;

static void Require(bool condition, string message)
{
    if (!condition)
        throw new InvalidDataException(message);
}

static bool WaitForExit(int pid)
{
    try
    {
        using var process = Process.GetProcessById(pid);
        return process.WaitForExit(5000);
    }
    catch (ArgumentException) { return true; }
}
