﻿using Nuke.Common;
using Nuke.Common.Tools.DotNet;
using Serilog;
using System.Collections.Generic;
using System.Linq;
using static Nuke.Common.Tools.DotNet.DotNetTasks;

partial class Build
{
    Target Compile => _ => _
    .TriggeredBy(Clean)
    .Executes(() =>
    {       
        foreach (var configuration in GlobBuildConfigurations())
        {
            Log.Information("Configuration name: {configuration}", configuration);

            if (configuration.StartsWith("Release"))
            {
                DotNetBuild(settings => settings
                    .SetConfiguration(configuration)
                    .SetVerbosity(DotNetVerbosity.quiet));
            }
        }

    });

    IEnumerable<string> GlobBuildConfigurations()
    {
        var configurations = Solution.Configurations
            .Select(pair => pair.Key)
            .Select(config => config.Remove(config.LastIndexOf('|')))
            .Distinct()
            .ToList();

        return configurations;
    }
}
