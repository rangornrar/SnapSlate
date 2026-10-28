using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SnapSlate.Tests;

[TestClass]
public sealed class ProjectPerformanceTests
{
    [TestMethod]
    [TestCategory("Performance")]
    public void Profile_SavedProjectChecks_RecordsTimeAndAllocations()
    {
        var imagePath = Environment.GetEnvironmentVariable("SNAPSLATE_PERF_IMAGE");
        var reportPath = Environment.GetEnvironmentVariable("SNAPSLATE_PERF_REPORT");
        if (imagePath is null || reportPath is null) Assert.Inconclusive("Set SNAPSLATE_PERF_IMAGE and SNAPSLATE_PERF_REPORT for an explicit profiling run.");
        var bytes = File.ReadAllBytes(imagePath!);
        using var image = System.Drawing.Image.FromFile(imagePath!);
        var rows = new[] { 1, 10, 50 }.Select(count =>
        {
            var state = new SnapSlateProjectState { Version = 2, Metadata = new(), Documents = Enumerable.Range(0, count).Select(index => new ScreenshotDocument
            {
                BaseTitle = "Step " + index, ImageBytes = bytes.ToArray(), SourcePixelWidth = image.Width, SourcePixelHeight = image.Height
            }).ToList() };
            var tracker = new ProjectSaveTracker();
            tracker.Accept(state);
            Assert.IsFalse(tracker.HasChanges(state));
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            var allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
            var watch = Stopwatch.StartNew();
            for (var iteration = 0; iteration < 5; iteration++)
            {
                Assert.IsFalse(tracker.HasChanges(state));
                tracker.RefreshIndicators(state);
            }
            watch.Stop();
            return new { steps = count, checks = 5, elapsedMs = watch.Elapsed.TotalMilliseconds, allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore };
        }).ToArray();
        File.WriteAllText(reportPath!, JsonSerializer.Serialize(new { image.Width, image.Height, imageBytes = bytes.Length, rows }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
