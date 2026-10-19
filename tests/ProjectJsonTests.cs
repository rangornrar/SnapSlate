using System.Text.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Windows.Foundation;

namespace SnapSlate.Tests;

[TestClass]
public sealed class ProjectJsonTests
{
    public sealed class GeometryState
    {
        public Point Position { get; set; }
        public Rect Crop { get; set; }
    }

    [TestMethod]
    public void Serialize_WithGeometry_RoundTripsCoordinates()
    {
        var state = new GeometryState { Position = new Point(-12.5, 8), Crop = new Rect(2, 3, 120.5, 80) };
        var json = ProjectJson.Serialize(state);
        var restored = ProjectJson.Deserialize<GeometryState>(json)!;
        Assert.AreEqual(state.Position, restored.Position);
        Assert.AreEqual(state.Crop, restored.Crop);
        StringAssert.Contains(json, "\"position\"");
    }

    [TestMethod]
    public void Deserialize_WithUnknownNestedGeometryFields_PreservesKnownCoordinates()
    {
        const string json = "{\"position\":{\"metadata\":{\"x\":999},\"x\":4,\"y\":5},\"crop\":{\"extra\":[1,2],\"x\":1,\"y\":2,\"width\":30,\"height\":40}}";
        var restored = ProjectJson.Deserialize<GeometryState>(json)!;
        Assert.AreEqual(new Point(4, 5), restored.Position);
        Assert.AreEqual(new Rect(1, 2, 30, 40), restored.Crop);
    }

    [TestMethod]
    public void Deserialize_WithInvalidGeometry_ThrowsJsonException()
    {
        Assert.ThrowsExactly<JsonException>(() => ProjectJson.Deserialize<GeometryState>("{\"position\":42}"));
        Assert.ThrowsExactly<JsonException>(() => ProjectJson.Deserialize<GeometryState>("{\"crop\":[]}"));
    }
}
