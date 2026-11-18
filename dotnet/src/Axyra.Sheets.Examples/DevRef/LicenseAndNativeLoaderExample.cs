namespace Axyra.Sheets.Examples.DevRef;

/// <summary>
/// The .NET reading of the Java "license and native loader" page. .NET has no JVM-style
/// native loader: the package carries the engine under <c>runtimes/&lt;rid&gt;/native/</c>,
/// and the static constructor of <see cref="Workbook"/> (or of <see cref="License"/>) loads
/// and checks it on first use. Setting <c>AXYRA_NATIVE_PATH</c> loads a specific engine
/// file instead (Java: <c>-Daxyra.native.path</c>). A load failure surfaces as a
/// <see cref="TypeInitializationException"/> whose inner exception is the cause. Licensing
/// is <see cref="License"/>: <c>License.Status</c> here, <c>License.Apply(...)</c> to install
/// a token at startup.
/// </summary>
public sealed class LicenseAndNativeLoaderExample : IExample
{
    public string Name => "dev-ref/license-and-native-loader";

    public string Description => "Verify native loading and inspect the active license";

    public void Run(string outputDirectory)
    {
        using (Workbook ignored = new Workbook())   // loads and checks the native engine
        {
            LicenseInfo info = License.Status;
            Console.WriteLine("Axyra native library loaded successfully");
            Console.WriteLine("License state: " + info.State);
            Console.WriteLine("PDF enabled: " + License.IsFeatureEnabled("pdf_render"));
        }
    }
}
