using Nuke.Common;
using Nuke.Common.IO;
using Nuke.Common.Tools.DotNet;
using Nuke.Common.Tools.SignTool;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using static Nuke.Common.Tools.SignTool.SignToolTasks;

partial class Build
{
    Target Sign => _ => _
    .TriggeredBy(Compile)
    .Executes(() => 
    {
        var compiledAssemblies = new List<string>();

        foreach (var project in Solution.AllProjects.Where(project => project != Solution._build))
        {
            AbsolutePath projectDirectory = project.Directory;
            Log.Information(projectDirectory);

            var files = projectDirectory
                .GlobDirectories(@"**\bin\**")
                .SelectMany(x => x.GlobFiles(CompiledAssemblies));

            foreach (var file in files)
            {
                Log.Information("File : {file}", file);
                compiledAssemblies.Add(file);
            }
        }

        SignFiles(compiledAssemblies);

    });

    static void SignFiles(List<string> compiledAssemblies) => SignTool(s => s
            .SetFileDigestAlgorithm("sha256")
            .SetTimestampServerUrl(@$"http://time.certum.pl")
            .SetSigningSubjectName("Open Source Developer, Russell Green")
            .SetFiles(compiledAssemblies.ToArray()));
}
