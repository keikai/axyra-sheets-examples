using System.Text;

namespace Axyra.Sheets.Examples.Tests;

/// <summary>
/// Runs a group of feature-demo examples and checks every file each one writes.
/// </summary>
/// <remarks>
/// The demos are split into groups, one test class each, because the evaluation build allows
/// 100 opens, saves and renders per process; <c>run-tests.sh</c> runs each test class in its
/// own process (the Java project relies on Surefire forking one JVM per class).
/// </remarks>
internal static class FeatureDemoTestSupport
{
    private static readonly byte[] PdfMagic = "%PDF"u8.ToArray();
    private static readonly byte[] PngMagic = [0x89, (byte)'P', (byte)'N', (byte)'G'];
    private static readonly byte[] ZipMagic = [(byte)'P', (byte)'K', 3, 4];

    internal static void RunAndCheck(IReadOnlyDictionary<string, string[]> expectedOutputs, string output)
    {
        foreach ((string name, string[] files) in expectedOutputs)
        {
            IExample example = ExampleRunner.Examples.FirstOrDefault(c => c.Name == name)
                ?? throw new Xunit.Sdk.XunitException("no example named " + name);
            string exampleOutput = Path.Combine(output, example.Name.Replace('/', '-'));
            Directory.CreateDirectory(exampleOutput);
            example.Run(exampleOutput);
            foreach (string fileName in files)
                CheckFile(example.Name, Path.Combine(exampleOutput, fileName));
        }
    }

    /// <summary>The file exists and starts the way its format requires.</summary>
    private static void CheckFile(string exampleName, string file)
    {
        string name = Path.GetFileName(file);
        Assert.True(File.Exists(file), $"{exampleName} did not produce {name}");
        byte[] bytes = File.ReadAllBytes(file);
        Assert.True(bytes.Length > 0, $"{exampleName} wrote an empty {name}");
        if (name.EndsWith(".xlsx", StringComparison.Ordinal))
            Assert.True(bytes.AsSpan(0, 4).SequenceEqual(ZipMagic), name + " is not a zip package");
        else if (name.EndsWith(".pdf", StringComparison.Ordinal))
            Assert.True(bytes.AsSpan(0, 4).SequenceEqual(PdfMagic), name + " is not a PDF");
        else if (name.EndsWith(".png", StringComparison.Ordinal))
            Assert.True(bytes.AsSpan(0, 4).SequenceEqual(PngMagic), name + " is not a PNG");
        else if (name.EndsWith(".svg", StringComparison.Ordinal))
            Assert.Contains("<svg", Encoding.UTF8.GetString(bytes));
    }
}
