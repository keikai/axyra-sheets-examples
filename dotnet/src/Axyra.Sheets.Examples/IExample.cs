namespace Axyra.Sheets.Examples;

/// <summary>A runnable example associated with one Axyra Sheets documentation page.</summary>
public interface IExample
{
    string Name { get; }

    string Description { get; }

    void Run(string outputDirectory);
}
