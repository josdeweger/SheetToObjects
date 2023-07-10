namespace SheetToObjects.Configuration.Settings.Configuration;

/// <summary>
/// Defines how the package will identify the assemblies which are scanned for sinks and other Type information.
/// </summary>
public enum ConfigurationAssemblySource
{
    /// <summary>
    /// Try to scan the assemblies already in memory. This is the default. If GetEntryAssembly is null, fallback to DLL scanning.
    /// </summary>
    UseLoadedAssemblies,

    /// <summary>
    /// Scan for assemblies in DLLs from the working directory. This is the fallback when GetEntryAssembly is null.
    /// </summary>
    AlwaysScanDllFiles
}
