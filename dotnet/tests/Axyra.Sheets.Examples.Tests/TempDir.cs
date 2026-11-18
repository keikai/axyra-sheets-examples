namespace Axyra.Sheets.Examples.Tests;

/// <summary>A per-test output directory, deleted on dispose (JUnit's <c>@TempDir</c>).</summary>
public sealed class TempDir : IDisposable
{
    public string Path { get; } = Directory.CreateTempSubdirectory("axyra-examples-").FullName;

    public string Resolve(string name) => System.IO.Path.Combine(Path, name);

    public void Dispose()
    {
        try { Directory.Delete(Path, recursive: true); } catch (IOException) { }
    }
}
