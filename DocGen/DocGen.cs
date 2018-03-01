using DocGen.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace DocGen
{
    public class GenerateModuleFiles
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Generating Module Files");
            var repoRoot = @"C:\repos\PnP-PowerShell\Commands\bin";
          
            List<CmdletInfo> allCmdlets = new List<CmdletInfo>();
            List<CmdletInfo> cmdlets = new List<CmdletInfo>();
            List<AssemblyInfo> assemblies = new List<AssemblyInfo>();
            assemblies.Add(new AssemblyInfo() { AssemblyPath = Path.Combine(repoRoot, @"debug\SharePointPnP.PowerShell.Online.Commands.dll"), Platform = "SharePoint Online" });
            assemblies.Add(new AssemblyInfo() { AssemblyPath = Path.Combine(repoRoot, @"debug15\SharePointPnP.PowerShell.2013.Commands.dll"), Platform = "SharePoint 2013" });
            assemblies.Add(new AssemblyInfo() { AssemblyPath = Path.Combine(repoRoot, @"debug16\SharePointPnP.PowerShell.2016.Commands.dll"), Platform = "SharePoint 2016" });
            //assemblies.Add(new AssemblyInfo() { AssemblyPath = Path.Combine(repoRoot, @"debug\netstandard2.0\SharePointPnP.PowerShell.Core.dll"), Platform = "SharePoint Online Cross-Platform" });
            foreach (var assemblyInfo in assemblies)
            {
                //var cmdletAssemblyDefinition = AssemblyDefinition.ReadAssembly(Path.Combine(repoRoot, assemblyPath), new ReaderParameters() { ReadingMode = ReadingMode.Immediate, ReadSymbols = true });
                Assembly cmdletAssembly = Assembly.LoadFrom(assemblyInfo.AssemblyPath);

                var analyzer = new CmdletsAnalyzer(cmdletAssembly, assemblyInfo.Platform);

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
                    Examples = new List<CmdletExampleAttributeEx>(),
                    Noun = first.Noun,
                    OutputType = first.OutputType,
                    OutputTypeDescription = first.OutputTypeDescription,
                    OutputTypeLink = first.OutputTypeLink,
                    Parameters = new List<CmdletParameterInfo>(),
                    Platforms = first.Platforms,
                    RelatedLinks = first.RelatedLinks,
                    Syntaxes = new List<CmdletSyntax>(),
                    Verb = first.Verb,
                    Version = first.Version
                };

                foreach (var additionalCmdlet in cmdletGroups)
                {
                    foreach (var parameter in additionalCmdlet.Parameters)
                    {
                        if (cmdletInfo.Parameters.FirstOrDefault(p => p.Name == parameter.Name) == null)
                        {
                            cmdletInfo.Parameters.Add(parameter);
                        }
                    }
                    foreach (var example in additionalCmdlet.Examples)
                    {
                        if (cmdletInfo.Examples.FirstOrDefault(e => e.Code == example.Code) == null)
                        {
                            cmdletInfo.Examples.Add(example);
                        } else
                        {
                            var existingExample = cmdletInfo.Examples.FirstOrDefault(e => e.Code == example.Code);
                            if(!existingExample.Platforms.Contains(example.Platforms[0]))
                            {
                                existingExample.Platforms.Add(example.Platforms[0]);
                            }
                        }
                    }
                    if(cmdletInfo.Examples.Count == assemblies.Count)
                    {
                        foreach(var example in cmdletInfo.Examples)
                        {
                            example.Platforms = new List<string>() { };
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

            var markDownGenerator = new MarkDownGenerator(cmdlets, @"c:\temp");
            markDownGenerator.Generate();

        }
    }
}
