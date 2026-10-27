using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Windows.Foundation;

namespace SnapSlate.Tests;

[TestClass]
public sealed class ProjectSafetyTests
{
    [TestMethod]
    [DataRow("{\"version\":999,\"documents\":[]}")]
    [DataRow("{\"documents\":null}")]
    [DataRow("{\"documents\":[null]}")]
    [DataRow("{\"documents\":[{\"annotations\":null}]}")]
    [DataRow("{\"documents\":[{\"annotations\":[null]}]}")]
    [DataRow("{\"documents\":[{\"selectedPaletteKey\":\"unknown\"}]}")]
    [DataRow("{\"documents\":[{\"baseTitle\":null}]}")]
    public void Deserialize_InvalidProject_RejectsBeforeReturningState(string json)
    {
        Assert.ThrowsExactly<InvalidDataException>(() => ProjectPersistence.Deserialize(json));
    }

    [TestMethod]
    public void Validate_DuplicateDocumentIds_RejectsAmbiguousState()
    {
        var document = new ScreenshotDocument();
        var state = new SnapSlateProjectState { Documents = [document, document.CloneForHistory()] };
        Assert.ThrowsExactly<InvalidDataException>(() => ProjectValidator.Validate(state));
    }

    [TestMethod]
    public void Validate_CropOutsideImage_RejectsGeometry()
    {
        var state = new SnapSlateProjectState { Documents = [new ScreenshotDocument { AppliedCropRect = new Rect(1300, 0, 100, 100) }] };
        Assert.ThrowsExactly<InvalidDataException>(() => ProjectValidator.Validate(state));
    }

    [TestMethod]
    public void Validate_NonFiniteAnnotation_RejectsGeometry()
    {
        var document = new ScreenshotDocument();
        document.Annotations.Add(new AnnotationModel { Opacity = double.NaN });
        Assert.ThrowsExactly<InvalidDataException>(() => ProjectValidator.Validate(new SnapSlateProjectState { Documents = [document] }));
    }

    [TestMethod]
    public void Serialize_ValidLegacyProject_RoundTripsAnnotations()
    {
        var document = new ScreenshotDocument { BaseTitle = "Projet" };
        document.Annotations.Add(new AnnotationModel { Kind = AnnotationKind.Mask, Bounds = new Rect(10, 20, 100, 80), Opacity = 0.2 });
        var restored = ProjectPersistence.Deserialize(ProjectPersistence.Serialize(new SnapSlateProjectState { Documents = [document] }))!;
        Assert.AreEqual(document.Id, restored.Documents[0].Id);
        Assert.AreEqual(document.Annotations[0].Bounds, restored.Documents[0].Annotations[0].Bounds);
    }

    [TestMethod]
    public void HasChanges_ExportPngAfterEdit_DoesNotAcquitProjectChanges()
    {
        var document = new ScreenshotDocument();
        var state = new SnapSlateProjectState { Documents = [document] };
        var tracker = new ProjectSaveTracker();
        tracker.Accept(state);
        document.StepNote = "New editable note";
        document.SavedPath = "export.png";
        document.IsDirty = false;
        Assert.IsTrue(tracker.HasChanges(state));
        tracker.RefreshIndicators(state);
        Assert.IsTrue(document.IsDirty);
    }

    [TestMethod]
    public void HasChanges_UndoAfterSave_DetectsDifferenceFromDisk()
    {
        var before = new ScreenshotDocument();
        var saved = before.CloneForHistory();
        saved.StepNote = "Saved revision";
        var tracker = new ProjectSaveTracker();
        tracker.Accept(new SnapSlateProjectState { Documents = [saved] });
        var undone = new SnapSlateProjectState { Documents = [before] };
        Assert.IsTrue(tracker.HasChanges(undone));
        tracker.RefreshIndicators(undone);
        Assert.IsTrue(before.IsDirty);
        Assert.IsFalse(tracker.HasChanges(new SnapSlateProjectState { Documents = [saved] }));
    }

    [TestMethod]
    public void Accept_EditWhileSaveIsRunning_KeepsLaterRevisionDirty()
    {
        var current = new ScreenshotDocument();
        var savedSnapshot = new SnapSlateProjectState { Documents = [current.CloneForHistory()] };
        current.StepNote = "Typed while disk write was awaiting";
        var tracker = new ProjectSaveTracker();
        tracker.Accept(savedSnapshot);
        var state = new SnapSlateProjectState { Documents = [current] };
        Assert.IsTrue(tracker.HasChanges(state));
        tracker.RefreshIndicators(state);
        Assert.IsTrue(current.IsDirty);
    }

    [TestMethod]
    public void HasChanges_SelectionAndExportPathOnly_RemainsClean()
    {
        var document = new ScreenshotDocument();
        var state = new SnapSlateProjectState { Documents = [document] };
        var tracker = new ProjectSaveTracker();
        tracker.Accept(state);
        state.SelectedDocumentId = document.Id;
        state.SelectedAnnotationId = Guid.NewGuid();
        document.SavedPath = "different.png";
        document.IsDirty = true;
        Assert.IsFalse(tracker.HasChanges(state));
    }

    [TestMethod]
    public void HasChanges_DocumentRemovedOrReordered_DetectsProjectChange()
    {
        var first = new ScreenshotDocument();
        var second = new ScreenshotDocument();
        var tracker = new ProjectSaveTracker();
        tracker.Accept(new SnapSlateProjectState { Documents = [first, second] });
        Assert.IsTrue(tracker.HasChanges(new SnapSlateProjectState { Documents = [second, first] }));
        Assert.IsTrue(tracker.HasChanges(new SnapSlateProjectState { Documents = [first] }));
    }
}
