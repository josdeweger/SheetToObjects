using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SheetToObjects.Adapters.MicrosoftExcel.Extensions.Assemblies;

sealed class DllScanningAssemblyFinder : AssemblyFinder
{
    public override IReadOnlyList<AssemblyName> FindAssembliesContainingName(string nameToFind)
    {
        var probeDirs = new List<string>();

        if (!string.IsNullOrEmpty(AppDomain.CurrentDomain.BaseDirectory))
        {
            probeDirs.Add(AppDomain.CurrentDomain.BaseDirectory);

#if NETFRAMEWORK
            var privateBinPath = AppDomain.CurrentDomain.SetupInformation.PrivateBinPath;
            if (!string.IsNullOrEmpty(privateBinPath))
            {
                foreach (var path in privateBinPath.Split(';'))
                {
                    if (Path.IsPathRooted(path))
                    {
                        probeDirs.Add(path);
                    }
                    else
                    {
                        probeDirs.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path));
                    }
                }
            }
#endif
        }
        else
        {
            var assemblyLocation = Path.GetDirectoryName(typeof(AssemblyFinder).Assembly.Location);
            if (assemblyLocation != null)
            {
                probeDirs.Add(assemblyLocation);
            }
        }

        var query = from probeDir in probeDirs
                    where Directory.Exists(probeDir)
                    from outputAssemblyPath in Directory.GetFiles(probeDir, "*.dll")
                    let assemblyFileName = Path.GetFileNameWithoutExtension(outputAssemblyPath)
                    where IsCaseInsensitiveMatch(assemblyFileName, nameToFind)
                    let assemblyName = TryGetAssemblyNameFrom(outputAssemblyPath)
                    where assemblyName != null
                    select assemblyName;

        return query.ToList();

        static AssemblyName? TryGetAssemblyNameFrom(string path)
        {
            try
            {
                return AssemblyName.GetAssemblyName(path);
            }
            catch (BadImageFormatException)
            {
                return null;
            }
        }
    }

    protected internal override IEnumerable<Assembly> GetAssemblies()
    {
        throw new NotImplementedException();
    }
}
