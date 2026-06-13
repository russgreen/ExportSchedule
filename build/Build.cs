using Fallout.Common;
using Fallout.Common.Git;
using Fallout.Common.IO;
using Fallout.Solutions;

partial class Build : FalloutBuild
{
    readonly AbsolutePath ArtifactsDirectory = RootDirectory / "output";

    readonly string[] CompiledAssemblies = { "ExportSchedule.dll" };

    [GitRepository]
    readonly GitRepository GitRepository;

    [Solution(GenerateProjects = true)]
    Solution Solution;

    public static int Main() => Execute<Build>(x => x.Clean);

    //[Parameter("Configuration to build - Default is 'Debug' (local) or 'Release' (server)")]
    //readonly Configuration Configuration = IsLocalBuild ? Configuration.Debug : Configuration.Release;


}
