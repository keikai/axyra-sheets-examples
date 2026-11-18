namespace Axyra.Sheets.Examples.Tests;

public class RemainingExamplesIntegrationTest
{
    [Fact]
    public void RunsEveryTutorialAndDeveloperReferenceExample()
    {
        using TempDir output = new();
        List<IExample> examples = ExampleRunner.Examples
            .Where(example => !example.Name.StartsWith("guides/", StringComparison.Ordinal))
            .Where(example => !example.Name.StartsWith("feature-demo/", StringComparison.Ordinal))
            .Where(example => example.Name != "tutorial/quick-start")
            .ToList();

        Assert.NotEmpty(examples);
        foreach (IExample example in examples)
        {
            string exampleOutput = Directory.CreateDirectory(
                output.Resolve(example.Name.Replace('/', '-'))).FullName;
            example.Run(exampleOutput);
        }
    }
}
