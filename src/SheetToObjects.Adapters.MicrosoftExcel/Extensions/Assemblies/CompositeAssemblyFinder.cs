using System.Collections.Generic;
using System.Reflection;

namespace SheetToObjects.Adapters.MicrosoftExcel.Extensions.Assemblies;

class CompositeAssemblyFinder : AssemblyFinder
{
    readonly AssemblyFinder[] _assemblyFinders;

    public CompositeAssemblyFinder(params AssemblyFinder[] assemblyFinders)
    {
        _assemblyFinders = assemblyFinders;
    }

    public override IReadOnlyList<AssemblyName> FindAssembliesContainingName(string nameToFind)
    {
        var assemblyNames = new List<AssemblyName>();
        foreach (var assemblyFinder in _assemblyFinders)
        {
            assemblyNames.AddRange(assemblyFinder.FindAssembliesContainingName(nameToFind));
        }
        return assemblyNames;
    }

    protected internal override IEnumerable<Assembly> GetAssemblies()
    {
        foreach (var assemblyFinder in _assemblyFinders)
            foreach (var assembly in assemblyFinder.GetAssemblies())
                yield return assembly;
    }
}
