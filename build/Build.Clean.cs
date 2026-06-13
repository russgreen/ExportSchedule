using Fallout.Common;
using Fallout.Common.IO;
using Fallout.Common.Tools.DotNet;
using Serilog;
using System.Linq;
using static Fallout.Common.Tools.DotNet.DotNetTasks;

partial class Build
{
    Target Clean => _ => _
        .Executes(() =>
        {
            foreach (var project in Solution.AllProjects.Where(project => project != Solution._build))
            {
                CleanDirectory(project.Directory / "bin");
                CleanDirectory(project.Directory / "obj");
            }

            foreach (var configuration in GlobBuildConfigurations())
            {
                DotNetClean(settings => settings
                    .SetConfiguration(configuration)
                    .SetVerbosity(DotNetVerbosity.minimal));
            }


            CleanDirectory(ArtifactsDirectory);

        });

    static void CleanDirectory(AbsolutePath path)
    {
        Log.Information("Cleaning directory: {Directory}", path);
        path.CreateOrCleanDirectory();
    }
}
