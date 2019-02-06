using SharePointPnP.PowerShell.Documentation;
using SharePointPnP.PowerShell.Documentation.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using CmdletInfo = SharePointPnP.PowerShell.Documentation.Model.CmdletInfo;

namespace Generate
{
    [Cmdlet(VerbsData.Publish, "PnPPSDocumentation")]
    public class PublishPnPPSDocumentation : PSCmdlet
    {
        [Parameter(Mandatory = true)]
        public string RepoRoot;

        [Parameter(Mandatory = true)]
        public string OutputFolder;

        [Parameter(Mandatory = false)]
        public SwitchParameter Book;

        protected override void ProcessRecord()
        {
            if (!Directory.Exists(RepoRoot))
            {
                throw new DirectoryNotFoundException($"{RepoRoot} does not exist");
            }

            WriteObject("Generating Module Files");

            List<CmdletInfo> allCmdlets = new List<CmdletInfo>();
            List<CmdletInfo> cmdlets = new List<CmdletInfo>();

            var assemblies = new Dictionary<string, string>();
            assemblies.Add("SharePoint Online", @"Commands\bin\debug\SharePointPnP.PowerShell.Online.Commands.dll");
            assemblies.Add("SharePoint Server 2013", @"Commands\bin\debug15\SharePointPnP.PowerShell.2013.Commands.dll");
            assemblies.Add("SharePoint Server 2016", @"Commands\bin\debug16\SharePointPnP.PowerShell.2016.Commands.dll");
            assemblies.Add("SharePoint Server 2019", @"Commands\bin\debug16\SharePointPnP.PowerShell.2019.Commands.dll");

            foreach (var assembly in assemblies)
            {
                WriteObject($"Processing {assembly.Key}");
                var assemblyPath = assembly.Value;

                Assembly cmdletAssembly = Assembly.LoadFrom(Path.Combine(RepoRoot, assemblyPath));

                var analyzer = new CmdletsAnalyzer(cmdletAssembly, assembly.Key);

                allCmdlets.AddRange(analyzer.Analyze());
            }
            //reorganize them
            foreach (var cmdletGroups in allCmdlets.GroupBy(c => c.FullCommand))
            {
                var first = cmdletGroups.First();
                var cmdletInfo = new CmdletInfo()
                {
                    AdditionalParameters = first.AdditionalParameters,
                    Aliases = first.Aliases,
                    Category = first.Category,
                    CmdletType = first.CmdletType,
                    Copyright = first.Copyright,
                    Description = first.Description,
                    DetailedDescription = first.DetailedDescription,
                    Examples = new List<SharePointPnP.PowerShell.CmdletHelpAttributes.CmdletExampleAttribute>(),
                    Noun = first.Noun,
                    OutputType = first.OutputType,
                    OutputTypeDescription = first.OutputTypeDescription,
                    OutputTypeLink = first.OutputTypeLink,
                    Parameters = new List<CmdletParameterInfo>(),
                    Platform = first.Platform,
                    RelatedLinks = first.RelatedLinks,
                    Syntaxes = new List<CmdletSyntax>(),
                    Verb = first.Verb,
                    Version = first.Version
                };

                foreach (var additionalCmdlet in cmdletGroups)
                {
                    foreach (var parameter in additionalCmdlet.Parameters)
                    {
                        var existingParameter = cmdletInfo.Parameters.FirstOrDefault(p => p.Name == parameter.Name);
                        if (existingParameter == null)
                        {
                            cmdletInfo.Parameters.Add(parameter);
                        }
                        else
                        {
                            var platforms = existingParameter.Platform.Split(new char[] { ',' }).ToList();
                            var parameterPlatforms = parameter.Platform?.Split(new char[] { ',' }).ToList();
                            if (parameterPlatforms != null && parameterPlatforms.Except(platforms).Any())
                            {
                                platforms.AddRange(parameterPlatforms.Except(platforms));
                                existingParameter.Platform = string.Join(",", platforms);
                            }
                        }
                    }
                    foreach (var example in additionalCmdlet.Examples)
                    {
                        if (cmdletInfo.Examples.FirstOrDefault(e => e.Code == example.Code) == null)
                        {
                            cmdletInfo.Examples.Add(example);
                        }
                    }
                    foreach (var syntax in additionalCmdlet.Syntaxes)
                    {
                        if (cmdletInfo.Syntaxes.FirstOrDefault(s => s.ParameterSetName == syntax.ParameterSetName) == null)
                        {
                            cmdletInfo.Syntaxes.Add(syntax);
                        }
                    }
                }
                cmdlets.Add(cmdletInfo);
            }

            WriteObject("Generate documentation");
            var markDownGenerator = new MarkDownGenerator(cmdlets, OutputFolder, Book);
            markDownGenerator.Generate();
        }
    }
}
