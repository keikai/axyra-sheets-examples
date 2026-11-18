namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Reads the input workbooks the feature demos start from, embedded as <c>feature-demo/*.xlsx</c>.</summary>
internal static class DemoBooks
{
    private const string Folder = "feature-demo/";

    internal static byte[] Read(string fileName)
    {
        using Stream? input = typeof(DemoBooks).Assembly.GetManifestResourceStream(Folder + fileName)
            ?? throw new InvalidOperationException("book not found among embedded resources: " + Folder + fileName);
        using MemoryStream buffer = new();
        input.CopyTo(buffer);
        return buffer.ToArray();
    }
}
