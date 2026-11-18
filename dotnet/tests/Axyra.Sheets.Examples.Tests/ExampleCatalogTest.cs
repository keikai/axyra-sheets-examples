namespace Axyra.Sheets.Examples.Tests;

public class ExampleCatalogTest
{
    [Fact]
    public void EveryExampleHasAUniqueNameAndDescription()
    {
        var names = new HashSet<string>();
        foreach (IExample example in ExampleRunner.Examples)
        {
            Assert.False(string.IsNullOrWhiteSpace(example.Name));
            Assert.False(string.IsNullOrWhiteSpace(example.Description));
            names.Add(example.Name);
        }
        Assert.Equal(ExampleRunner.Examples.Count, names.Count);
    }

    [Fact]
    public void EveryFeatureDemoIsCoveredByAnIntegrationTest()
    {
        var registered = new SortedSet<string>(StringComparer.Ordinal);
        foreach (IExample example in ExampleRunner.Examples)
        {
            if (example.Name.StartsWith("feature-demo/", StringComparison.Ordinal))
            {
                registered.Add(example.Name);
            }
        }
        var tested = new SortedSet<string>(StringComparer.Ordinal);
        tested.UnionWith(FeatureDemoBasicsIntegrationTest.ExpectedOutputs.Keys);
        tested.UnionWith(FeatureDemoDataAndOutputIntegrationTest.ExpectedOutputs.Keys);
        Assert.Equal(20, registered.Count);
        Assert.Equal(registered, tested);
    }
}
