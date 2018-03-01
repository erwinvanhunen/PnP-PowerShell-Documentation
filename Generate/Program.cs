﻿using Mono.Cecil;
using SharePointPnP.PowerShell.ModuleFilesGenerator;
using SharePointPnP.PowerShell.ModuleFilesGenerator.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Generate
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Generating Module Files");
            var repoRoot = @"C:\repos\PnP-PowerShell\Commands\bin";
            //var assemblyPath = args[0];
            //var configurationName = args[1];
            //var solutionDir = args[2];

            //try
            //{
            List<CmdletInfo> allCmdlets = new List<CmdletInfo>();
            List<CmdletInfo> cmdlets = new List<CmdletInfo>();

            var assemblies = new Dictionary<string, string>();
            assemblies.Add("SharePoint Online", @"debug\SharePointPnP.PowerShell.Online.Commands.dll");
            assemblies.Add("SharePoint Server 2013", @"debug15\SharePointPnP.PowerShell.2013.Commands.dll");
            assemblies.Add("SharePoint Server 2016", @"debug16\SharePointPnP.PowerShell.2016.Commands.dll");

            foreach (var assembly in assemblies)
            //foreach(var assemblyPath in new string[] { @"debug\netstandard2.0\SharePointPnP.PowerShell.Core.dll" })
            {
                var assemblyPath = assembly.Value;
                //var cmdletAssemblyDefinition = AssemblyDefinition.ReadAssembly(Path.Combine(repoRoot, assemblyPath), new ReaderParameters() { ReadingMode = ReadingMode.Immediate, ReadSymbols = true });
                Assembly cmdletAssembly = Assembly.LoadFrom(Path.Combine(repoRoot, assemblyPath));

                var analyzer = new CmdletsAnalyzer(cmdletAssembly,assembly.Key);

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
                                existingParameter.Platform = string.Join(',', platforms);
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

            //var helpFileGenerator = new HelpFileGenerator(cmdlets, cmdletAssembly, $"{assemblyPath}-help.xml");
            //helpFileGenerator.Generate();

            var markDownGenerator = new MarkDownGenerator(cmdlets, @"c:\temp");
            markDownGenerator.Generate();

            //var moduleManifestGenerator = new ModuleManifestGenerator(cmdlets, assemblyPath, configurationName, cmdletAssembly.GetName().Version);
            //moduleManifestGenerator.Generate();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"ERROR: {ex.Message}");
            //    //return 1;
            //}
            //return 0;
        }
    }
}
