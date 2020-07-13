using Newtonsoft.Json;
using SharePointPnP.PowerShell.CmdletHelpAttributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace SharePointPnP.PowerShell.Documentation
{
    internal class MarkDownGenerator
    {
        private List<Model.CmdletInfo> _cmdlets;
        private string _outputDirectory;
        private const string extension = "md";
        private bool _book;
        internal MarkDownGenerator(List<Model.CmdletInfo> cmdlets, string outputDirectory, bool book)
        {
            _cmdlets = cmdlets;
            _outputDirectory = outputDirectory;
            _book = book;
        }

        internal void Generate()
        {
            GenerateCmdletDocs();
            GenerateMappingJson();
            GenerateTOC();
            GenerateMSDNTOC();

            DirectoryInfo di = new DirectoryInfo($"{_outputDirectory}");
            var mdFiles = di.GetFiles("*.md");

            // Clean up old MD files
            foreach (var mdFile in mdFiles)
            {
                if (mdFile.Name.ToLowerInvariant() != $"readme.{extension}")
                {
                    var index = _cmdlets.FindIndex(t => $"{t.Verb}-{t.Noun}.{extension}" == mdFile.Name);
                    if (index == -1)
                    {
                        mdFile.Delete();
                    }
                }
            }
        }

        private void GenerateCmdletDocs()
        {

            foreach (var cmdletInfo in _cmdlets)
            {
                var originalMd = string.Empty;
                var newMd = string.Empty;


                if (!string.IsNullOrEmpty(cmdletInfo.Verb) && !string.IsNullOrEmpty(cmdletInfo.Noun))
                {
                    string mdFilePath = Path.Combine(_outputDirectory, $"{cmdletInfo.Verb}-{cmdletInfo.Noun}.{extension}");

                    if (System.IO.File.Exists(mdFilePath))
                    {
                        originalMd = System.IO.File.ReadAllText(mdFilePath);

                    }
                    var docBuilder = new StringBuilder();

                    if (_book)
                    {
                        docBuilder = WriteBookMD(docBuilder, cmdletInfo);
                    }
                    else
                    {
                        docBuilder.Append($@"---{Environment.NewLine}external help file:{Environment.NewLine}online version: https://docs.microsoft.com/powershell/module/sharepoint-pnp/{cmdletInfo.FullCommand.ToLower()}{Environment.NewLine}applicable: {cmdletInfo.Platform}{Environment.NewLine}schema: 2.0.0{Environment.NewLine}title: {cmdletInfo.FullCommand}{Environment.NewLine}---{Environment.NewLine}");

                        docBuilder.Append($"{Environment.NewLine}# {cmdletInfo.FullCommand}{Environment.NewLine}{Environment.NewLine}");

                        docBuilder.Append($"## SYNOPSIS{Environment.NewLine}");

                        var permissionHeaderAdded = false;
                        // API Permissions
                        if (cmdletInfo.ApiPermissions.Count > 0)
                        {
                            docBuilder.AppendLine($"{Environment.NewLine}**Required Permissions**{Environment.NewLine}");
                            permissionHeaderAdded = true;
                            // Go through each of the API permissions on this cmdlet and build a dictionary with the API type as the key and the requested scope(s) on it as the value
                            var requiredOrPermissionByApiDictionary = new Dictionary<string, List<string>>(cmdletInfo.ApiPermissions.Count);
                            var requiredAndPermissionByApiDictionary = new Dictionary<string, List<string>>(cmdletInfo.ApiPermissions.Count);
                            foreach (var requiredApiPermission in cmdletInfo.ApiPermissions)
                            {

                                requiredOrPermissionByApiDictionary = ParseApiPermissions("OrApiPermissions", requiredApiPermission, requiredOrPermissionByApiDictionary);
                                requiredAndPermissionByApiDictionary = ParseApiPermissions("AndApiPermissions", requiredApiPermission, requiredAndPermissionByApiDictionary);
                            }

                            // Loop through all the items in our dictionary with API types and requested scopes on the API so we can add it to the help text output
                            foreach (var requiredPermissionDictionaryItem in requiredOrPermissionByApiDictionary)
                            {
                                if (requiredPermissionDictionaryItem.Value.Count > 0)
                                {
                                    if (requiredPermissionDictionaryItem.Value.Count() == 1)
                                    {
                                        docBuilder.AppendLine($"  * {requiredPermissionDictionaryItem.Key}: {requiredPermissionDictionaryItem.Value[0]}");
                                    }
                                    else
                                    {
                                        docBuilder.AppendLine($"  * {requiredPermissionDictionaryItem.Key} : One of {string.Join(", ", requiredPermissionDictionaryItem.Value.OrderBy(s => s))}");
                                    }
                                }
                            }
                            foreach (var requiredPermissionDictionaryItem in requiredAndPermissionByApiDictionary)
                            {
                                if (requiredPermissionDictionaryItem.Value.Count > 0)
                                {
                                    if (requiredPermissionDictionaryItem.Value.Count() == 1)
                                    {
                                        docBuilder.AppendLine($"  * {requiredPermissionDictionaryItem.Key}: {requiredPermissionDictionaryItem.Value[0]}");
                                    }
                                    else
                                    {
                                        docBuilder.AppendLine($"  *  {requiredPermissionDictionaryItem.Key}: All of {string.Join(", ", requiredPermissionDictionaryItem.Value.OrderBy(s => s))}");
                                    }
                                }
                            }

                            docBuilder.AppendLine();
                        }

                        // Notice if the cmdlet is a PnPAdminCmdlet that access to the SharePoint Tenant Admin site is needed
                        if (cmdletInfo.CmdletType.BaseType.Name.Equals("PnPAdminCmdlet"))
                        {
                            if (!permissionHeaderAdded)
                            {
                                docBuilder.AppendLine($"{Environment.NewLine}**Required Permissions**{Environment.NewLine}");
                            }
                            docBuilder.AppendLine("* SharePoint: Access to the SharePoint Tenant Administration site");
                            docBuilder.AppendLine();
                        }

                        docBuilder.AppendLine($"{cmdletInfo.Description}{Environment.NewLine}");

                        if (cmdletInfo.Syntaxes.Any())
                        {
                            docBuilder.Append($"## SYNTAX {Environment.NewLine}{Environment.NewLine}");
                            foreach (var cmdletSyntax in cmdletInfo.Syntaxes.OrderBy(s => s.Parameters.Count(p => p.Required)))
                            {
                                if (cmdletSyntax.ParameterSetName != "__AllParameterSets")
                                {
                                    docBuilder.Append($"### {cmdletSyntax.ParameterSetName}{Environment.NewLine}");
                                }
                                var syntaxText = new StringBuilder();
                                syntaxText.AppendFormat("```powershell\r\n{0} ", cmdletInfo.FullCommand);
                                var cmdletLength = cmdletInfo.FullCommand.Length;
                                var first = true;
                                foreach (var par in cmdletSyntax.Parameters.Distinct(new ParameterComparer()).OrderBy(p => p.Order).ThenBy(p => !p.Required).ThenBy(p => p.Position))
                                {
                                    if (first)
                                    {
                                        first = false;
                                    }
                                    else
                                    {
                                        syntaxText.Append(new string(' ', cmdletLength + 1));
                                    }
                                    if (!par.Required)
                                    {
                                        syntaxText.Append("[");
                                    }
                                    if (par.Type.StartsWith("Int"))
                                    {
                                        par.Type = "Int";
                                    }
                                    if (par.Type == "SwitchParameter")
                                    {
                                        syntaxText.AppendFormat("-{0} [<{1}>]", par.Name, par.Type);
                                    }
                                    else
                                    {
                                        syntaxText.AppendFormat("-{0} <{1}>", par.Name, par.Type);
                                    }
                                    if (!par.Required)
                                    {
                                        syntaxText.Append("]");
                                    }
                                    syntaxText.Append("\r\n");
                                }
                                // Add All ParameterSet ones
                                docBuilder.Append(syntaxText);
                                docBuilder.Append($"```{Environment.NewLine}{Environment.NewLine}");
                            }
                        }

                        
                        if (!string.IsNullOrEmpty(cmdletInfo.DetailedDescription))
                        {
                            docBuilder.Append($"## DESCRIPTION{Environment.NewLine}");
                            docBuilder.Append($"{cmdletInfo.DetailedDescription}{Environment.NewLine}{Environment.NewLine}");
                        }

                        if (cmdletInfo.Examples.Any())
                        {
                            docBuilder.Append($"## EXAMPLES{Environment.NewLine}{Environment.NewLine}");
                            var examplesCount = 1;
                            foreach (var example in cmdletInfo.Examples.OrderBy(e => e.SortOrder))
                            {

                                docBuilder.Append($"### ------------------EXAMPLE {examplesCount}------------------{Environment.NewLine}");
                                if (!string.IsNullOrEmpty(example.Introduction))
                                {
                                    docBuilder.Append($"{example.Introduction}{Environment.NewLine}");
                                }
                                docBuilder.Append($"```powershell{Environment.NewLine}{example.Code.Replace("PS:> ", "")}{Environment.NewLine}```{Environment.NewLine}{Environment.NewLine}");
                                docBuilder.Append($"{example.Remarks}{Environment.NewLine}{Environment.NewLine}");
                                examplesCount++;
                            }
                        }

                        if (cmdletInfo.Parameters.Any())
                        {
                            docBuilder.Append($"## PARAMETERS{Environment.NewLine}{Environment.NewLine}");

                            foreach (var parameter in cmdletInfo.Parameters.OrderBy(x => x.Order).ThenBy(x => x.Name).Distinct(new ParameterComparer()))
                            {
                                if (parameter.Type.StartsWith("Int"))
                                {
                                    parameter.Type = "Int";
                                }
                                docBuilder.Append($"### -{parameter.Name}{Environment.NewLine}");
                                docBuilder.Append($"{parameter.Description}");
                                if (!string.IsNullOrEmpty(parameter.Platform))
                                {
                                    var cmdletPlatforms = cmdletInfo.Platform.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
                                    var parameterPlatforms = parameter.Platform.Split(new char[] { ',' });
                                    if (cmdletPlatforms.Except(parameterPlatforms).Any())
                                    {
                                        var rewrittenPlatform = string.Join(", ", parameter.Platform.Split(new char[] { ',' }));
                                        docBuilder.Append($"{Environment.NewLine}{Environment.NewLine}Only applicable to: {rewrittenPlatform}");
                                    }
                                }
                                docBuilder.Append($"{Environment.NewLine}{Environment.NewLine}");
                                docBuilder.Append($"```yaml{Environment.NewLine}");
                                docBuilder.Append($"Type: {parameter.Type}{Environment.NewLine}");
                                if (string.IsNullOrEmpty(parameter.ParameterSetName))
                                {
                                    parameter.ParameterSetName = "(All)";
                                }
                                docBuilder.Append($"Parameter Sets: { parameter.ParameterSetName}{Environment.NewLine}");
                                if (parameter.Aliases.Any())
                                {
                                    docBuilder.Append($"Aliases: {string.Join(",", parameter.Aliases)}{Environment.NewLine}");
                                }
                                docBuilder.Append(Environment.NewLine);
                                docBuilder.Append($"Required: {parameter.Required}{Environment.NewLine}");
                                docBuilder.Append($"Position: {(parameter.Position == int.MinValue ? "Named" : parameter.Position.ToString())}{Environment.NewLine}");
                                docBuilder.Append($"Accept pipeline input: {parameter.ValueFromPipeline}{Environment.NewLine}");
                                docBuilder.Append($"```{Environment.NewLine}{Environment.NewLine}");
                            }
                        }

                        if (cmdletInfo.OutputType != null)
                        {
                            docBuilder.Append($"## OUTPUTS{Environment.NewLine}{Environment.NewLine}");
                            var outputType = "";
                            if (cmdletInfo.OutputType != null)
                            {
                                if (cmdletInfo.OutputType.IsGenericType)
                                {
                                    if (cmdletInfo.OutputType.GetGenericTypeDefinition() == typeof(List<>) || cmdletInfo.OutputType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                                    {
                                        if (cmdletInfo.OutputType.GenericTypeArguments.Any())
                                        {
                                            outputType = $"List<{cmdletInfo.OutputType.GenericTypeArguments[0].FullName}>";
                                        }
                                        else
                                        {
                                            outputType = cmdletInfo.OutputType.FullName;
                                        }
                                    }
                                    else
                                    {
                                        outputType = cmdletInfo.OutputType.FullName;
                                    }
                                }
                                else
                                {
                                    outputType = cmdletInfo.OutputType.FullName;
                                }
                            }
                            if (!string.IsNullOrEmpty(cmdletInfo.OutputTypeLink))
                            {
                                docBuilder.Append($"### {outputType}");
                            }
                            else
                            {
                                docBuilder.Append($"### {outputType}");
                            }
                            if (!string.IsNullOrEmpty(cmdletInfo.OutputTypeDescription))
                            {
                                docBuilder.Append($"\n\n{cmdletInfo.OutputTypeDescription}");
                            }
                            docBuilder.Append("\n\n");
                        }

                        if (cmdletInfo.RelatedLinks.Any())
                        {
                            docBuilder.Append($"## RELATED LINKS{Environment.NewLine}{Environment.NewLine}");
                            foreach (var link in cmdletInfo.RelatedLinks)
                            {
                                docBuilder.Append($"[{link.Text}]({link.Url})");
                            }
                        }
                    }

                    newMd = docBuilder.ToString();

                    var dmp = new DiffMatchPatch.diff_match_patch();

                    var diffResults = dmp.diff_main(newMd, originalMd);

                    foreach (var result in diffResults)
                    {
                        if (result.operation != DiffMatchPatch.Operation.EQUAL)
                        {
                            System.IO.File.WriteAllText(mdFilePath, docBuilder.ToString());
                            break;
                        }
                    }
                }
            }
        }

        private StringBuilder WriteBookMD(StringBuilder docBuilder, Model.CmdletInfo cmdletInfo)
        {
            if (cmdletInfo.Syntaxes.Any())
            {
                docBuilder.Append($"# {cmdletInfo.FullCommand}{Environment.NewLine}{Environment.NewLine}");

                docBuilder.Append($"{cmdletInfo.Description}{Environment.NewLine}{Environment.NewLine}");

                if (!string.IsNullOrEmpty(cmdletInfo.DetailedDescription))
                {
                    docBuilder.Append($"{cmdletInfo.DetailedDescription}{Environment.NewLine}{Environment.NewLine}");
                }

                docBuilder.Append($"## Syntaxes {Environment.NewLine}{Environment.NewLine}");
                foreach (var cmdletSyntax in cmdletInfo.Syntaxes.OrderBy(s => s.Parameters.Count(p => p.Required)))
                {
                    if (cmdletSyntax.ParameterSetName != "__AllParameterSets")
                    {
                        docBuilder.Append($"### {cmdletSyntax.ParameterSetName}{Environment.NewLine}");
                    }
                    var syntaxText = new StringBuilder();
                    syntaxText.AppendFormat("```\r\n{0} ", cmdletInfo.FullCommand);
                    var cmdletLength = cmdletInfo.FullCommand.Length;
                    var first = true;
                    foreach (var par in cmdletSyntax.Parameters.Distinct(new ParameterComparer()).OrderBy(p => p.Order).ThenBy(p => !p.Required).ThenBy(p => p.Position))
                    {
                        if (first)
                        {
                            first = false;
                        }
                        else
                        {
                            syntaxText.Append(new string(' ', cmdletLength + 1));
                        }
                        if (!par.Required)
                        {
                            syntaxText.Append("[");
                        }
                        if (par.Type.StartsWith("Int"))
                        {
                            par.Type = "Int";
                        }
                        if (par.Type == "SwitchParameter")
                        {
                            syntaxText.AppendFormat("-{0} [<{1}>]", par.Name, par.Type);
                        }
                        else
                        {
                            syntaxText.AppendFormat("-{0} <{1}>", par.Name, par.Type);
                        }
                        if (!par.Required)
                        {
                            syntaxText.Append("]");
                        }
                        syntaxText.Append("\r\n");
                    }
                    // Add All ParameterSet ones
                    docBuilder.Append(syntaxText);
                    docBuilder.Append($"```{Environment.NewLine}{Environment.NewLine}");
                }
            }

            if (cmdletInfo.Parameters.Any())
            {
                docBuilder.Append($"## Parameters{Environment.NewLine}{Environment.NewLine}");

                docBuilder.Append($"|Name|Type|Description|Mandatory|Remarks|{Environment.NewLine}");
                docBuilder.Append($"|----|----|-----------|---------|-------|{Environment.NewLine}");

                foreach (var parameter in cmdletInfo.Parameters.OrderBy(x => x.Order).ThenBy(x => x.Name).Distinct(new ParameterComparer()))
                {
                    if (parameter.Type.StartsWith("Int"))
                    {
                        parameter.Type = "Int";
                    }
                    docBuilder.Append($"|{parameter.Name}|{parameter.Type}|{parameter.Description}|{parameter.Required}||{Environment.NewLine}");

                    //if (!string.IsNullOrEmpty(parameter.Platform))
                    //{
                    //    var cmdletPlatforms = cmdletInfo.Platform.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
                    //    var parameterPlatforms = parameter.Platform.Split(new char[] { ',' });
                    //    if (cmdletPlatforms.Except(parameterPlatforms).Any())
                    //    {
                    //        var rewrittenPlatform = string.Join(", ", parameter.Platform.Split(new char[] { ',' }));
                    //        docBuilder.Append($"{Environment.NewLine}{Environment.NewLine}Only applicable to: {rewrittenPlatform}");
                    //    }
                    //}
                    //docBuilder.Append($"{Environment.NewLine}{Environment.NewLine}");
                    //docBuilder.Append($"```yaml{Environment.NewLine}");
                    //docBuilder.Append($"Type: {parameter.Type}{Environment.NewLine}");
                    //if (string.IsNullOrEmpty(parameter.ParameterSetName))
                    //{
                    //    parameter.ParameterSetName = "(All)";
                    //}
                    //docBuilder.Append($"Parameter Sets: { parameter.ParameterSetName}{Environment.NewLine}");
                    //if (parameter.Aliases.Any())
                    //{
                    //    docBuilder.Append($"Aliases: {string.Join(",", parameter.Aliases)}{Environment.NewLine}");
                    //}
                    //docBuilder.Append(Environment.NewLine);
                    //docBuilder.Append($"Required: {parameter.Required}{Environment.NewLine}");
                    //docBuilder.Append($"Position: {(parameter.Position == int.MinValue ? "Named" : parameter.Position.ToString())}{Environment.NewLine}");
                    //docBuilder.Append($"Accept pipeline input: {parameter.ValueFromPipeline}{Environment.NewLine}");
                    //docBuilder.Append($"```{Environment.NewLine}{Environment.NewLine}");
                }
            }

            if (cmdletInfo.Examples.Any())
            {
                docBuilder.Append($"## Examples{Environment.NewLine}{Environment.NewLine}");
                var examplesCount = 1;
                foreach (var example in cmdletInfo.Examples.OrderBy(e => e.SortOrder))
                {

                    docBuilder.Append($"__Example {examplesCount}__{Environment.NewLine}");
                    if (!string.IsNullOrEmpty(example.Introduction))
                    {
                        docBuilder.Append($"{example.Introduction}{Environment.NewLine}");
                    }
                    docBuilder.Append($"```{Environment.NewLine}{example.Code.Replace("PS:> ", "")}{Environment.NewLine}```{Environment.NewLine}{Environment.NewLine}");
                    docBuilder.Append($"{example.Remarks}{Environment.NewLine}{Environment.NewLine}");
                    examplesCount++;
                }
            }

            //if (cmdletInfo.OutputType != null)
            //{
            //    docBuilder.Append($"## OUTPUTS{Environment.NewLine}{Environment.NewLine}");
            //    var outputType = "";
            //    if (cmdletInfo.OutputType != null)
            //    {
            //        if (cmdletInfo.OutputType.IsGenericType)
            //        {
            //            if (cmdletInfo.OutputType.GetGenericTypeDefinition() == typeof(List<>) || cmdletInfo.OutputType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            //            {
            //                if (cmdletInfo.OutputType.GenericTypeArguments.Any())
            //                {
            //                    outputType = $"List<{cmdletInfo.OutputType.GenericTypeArguments[0].FullName}>";
            //                }
            //                else
            //                {
            //                    outputType = cmdletInfo.OutputType.FullName;
            //                }
            //            }
            //            else
            //            {
            //                outputType = cmdletInfo.OutputType.FullName;
            //            }
            //        }
            //        else
            //        {
            //            outputType = cmdletInfo.OutputType.FullName;
            //        }
            //    }
            //    if (!string.IsNullOrEmpty(cmdletInfo.OutputTypeLink))
            //    {
            //        docBuilder.Append($"### {outputType}");
            //    }
            //    else
            //    {
            //        docBuilder.Append($"### {outputType}");
            //    }
            //    if (!string.IsNullOrEmpty(cmdletInfo.OutputTypeDescription))
            //    {
            //        docBuilder.Append($"\n\n{cmdletInfo.OutputTypeDescription}");
            //    }
            //    docBuilder.Append("\n\n");
            //}

            //if (cmdletInfo.RelatedLinks.Any())
            //{
            //    docBuilder.Append($"## RELATED LINKS{Environment.NewLine}{Environment.NewLine}");
            //    foreach (var link in cmdletInfo.RelatedLinks)
            //    {
            //        docBuilder.Append($"[{link.Text}]({link.Url})");
            //    }
            //}
            return docBuilder;

        }
        private void GenerateMappingJson()
        {
            var groups = new Dictionary<string, string>();
            foreach (var cmdletInfo in _cmdlets)
            {
                groups.Add($"{cmdletInfo.FullCommand}", cmdletInfo.Category);
            }

            var json = JsonConvert.SerializeObject(groups);

            var mappingFolder = $"{_outputDirectory}\\Documentation\\Mapping";
            if (!System.IO.Directory.Exists(mappingFolder))
            {
                System.IO.Directory.CreateDirectory(mappingFolder);
            }

            var mappingPath = $"{_outputDirectory}\\Documentation\\Mapping\\groupMapping.json";
            System.IO.File.WriteAllText(mappingPath, json);
        }

        private Dictionary<string, List<string>> ParseApiPermissions(string propertyName, CmdletApiPermissionBase requiredApiPermission, Dictionary<string, List<string>> requiredPermissionByApiDictionary)
        {
            // Through reflection, get the friendly name of the API and the permission scopes. We must use reflection here as generic attribute types do not exist in C# yet. Multiple scopes are returned as a comma-space separated string.
            var scopePermission = requiredApiPermission.GetType().GetProperty(propertyName)?.GetValue(requiredApiPermission, null);

            // If we were unable to retrieve the scopes and/or api through reflection, continue with the next permission attribute
            if (scopePermission != null && requiredApiPermission.ApiName != null)
            {

                // Ensure our dictionary with APIs already has an entry for the current API being processed
                if (!requiredPermissionByApiDictionary.ContainsKey(requiredApiPermission.ApiName))
                {
                    // Add another entry to the dictionary for this API
                    requiredPermissionByApiDictionary.Add(requiredApiPermission.ApiName, new List<string>());
                }

                // The returned scopes are the enum names which use an underscore instead of a dot for the scope names and are comma-space separated if multiple scopes are possible
                var scopes = scopePermission.ToString().Replace("_", ".").Split(',').Where(s => s != "None");

                // Go through each scope
                foreach (var scope in scopes)
                {
                    // Remove the potential trailing space
                    var trimmedScope = scope.Trim();

                    // If the scope is not known to be required yet, add it to the dictionary with APIs and requested scopes
                    if (!requiredPermissionByApiDictionary[requiredApiPermission.ApiName].Contains(trimmedScope))
                    {
                        requiredPermissionByApiDictionary[requiredApiPermission.ApiName].Add(trimmedScope);
                    }
                }
            }
            return requiredPermissionByApiDictionary;
        }
        private void GenerateTOC()
        {
            var originalMd = string.Empty;
            var newMd = string.Empty;

            // Create the readme.md
            var readmePath = $"{_outputDirectory}\\Documentation\\readme.{extension}";
            if (System.IO.File.Exists(readmePath))
            {
                originalMd = System.IO.File.ReadAllText(readmePath);
            }
            var docBuilder = new StringBuilder();


            docBuilder.AppendFormat("# Cmdlet Documentation #{0}", Environment.NewLine);
            docBuilder.AppendFormat("Below you can find a list of all the available cmdlets. Many commands provide built-in help and examples. Retrieve the detailed help with {0}", Environment.NewLine);
            docBuilder.AppendFormat("{0}```powershell{0}Get-Help Connect-PnPOnline -Detailed{0}```{0}{0}", Environment.NewLine);

            // Get all unique categories
            var categories = _cmdlets.Where(c => !string.IsNullOrEmpty(c.Category)).Select(c => c.Category).Distinct();

            foreach (var category in categories.OrderBy(c => c))
            {
                docBuilder.AppendFormat("## {0}{1}", category, Environment.NewLine);

                docBuilder.AppendFormat("Cmdlet|Description|Platforms{0}", Environment.NewLine);
                docBuilder.AppendFormat(":-----|:----------|:--------{0}", Environment.NewLine);
                foreach (var cmdletInfo in _cmdlets.Where(c => c.Category == category).OrderBy(c => c.Noun))
                {
                    var description = cmdletInfo.Description != null ? cmdletInfo.Description.Replace("\r\n", " ") : "";
                    docBuilder.AppendFormat("**[{0}]({1}-{2}.md)** |{3}|{4}{5}", cmdletInfo.FullCommand.Replace("-", "&#8209;"), cmdletInfo.Verb, cmdletInfo.Noun, description, cmdletInfo.Platform, Environment.NewLine);
                }
            }

            newMd = docBuilder.ToString();
            DiffMatchPatch.diff_match_patch dmp = new DiffMatchPatch.diff_match_patch();

            var diffResults = dmp.diff_main(newMd, originalMd);

            foreach (var result in diffResults)
            {
                if (result.operation != DiffMatchPatch.Operation.EQUAL)
                {
                    System.IO.File.WriteAllText(readmePath, docBuilder.ToString());
                }
            }
        }

        private void GenerateMSDNTOC()
        {
            var originalTocMd = string.Empty;
            var newTocMd = string.Empty;

            var msdnDocPath = $"{_outputDirectory}\\Documentation\\docs-conceptual\\sharepoint-pnp";
            if (!Directory.Exists(msdnDocPath))
            {
                Directory.CreateDirectory(msdnDocPath);
            }

            // Generate the landing page
            var landingPagePath = $"{msdnDocPath}\\sharepoint-pnp-cmdlets.{extension}";
            GenerateMSDNLandingPage(landingPagePath);

            // TOC.md generation
            var tocPath = $"{msdnDocPath}\\TOC.{extension}";
            if (System.IO.File.Exists(tocPath))
            {
                originalTocMd = System.IO.File.ReadAllText(tocPath);
            }
            var docBuilder = new StringBuilder();


            docBuilder.AppendFormat("# [SharePoint PnP PowerShell reference](PnP-PowerShell-Overview.md){0}", Environment.NewLine);

            // Get all unique categories
            var categories = _cmdlets.Where(c => !string.IsNullOrEmpty(c.Category)).Select(c => c.Category).Distinct();

            foreach (var category in categories.OrderBy(c => c))
            {
                var categoryMdPage = $"{category.Replace(" ", "")}-category.{extension}";
                var categoryMdPath = $"{msdnDocPath}\\{categoryMdPage}";

                // Add section reference to TOC
                docBuilder.AppendFormat("## [{0}]({1}){2}", category, categoryMdPage, Environment.NewLine);

                var categoryCmdlets = _cmdlets.Where(c => c.Category == category).OrderBy(c => c.Noun);

                // Generate category MD
                GenerateMSDNCategory(category, categoryMdPath, categoryCmdlets);

                // Link cmdlets to TOC
                foreach (var cmdletInfo in categoryCmdlets)
                {
                    var description = cmdletInfo.Description != null ? cmdletInfo.Description.Replace("\r\n", " ") : "";
                    docBuilder.AppendFormat("### [{0}]({1}-{2}.md){3}", cmdletInfo.FullCommand, cmdletInfo.Verb, cmdletInfo.Noun, Environment.NewLine);
                }
            }

            newTocMd = docBuilder.ToString();
            DiffMatchPatch.diff_match_patch dmp = new DiffMatchPatch.diff_match_patch();

            var diffResults = dmp.diff_main(newTocMd, originalTocMd);

            foreach (var result in diffResults)
            {
                if (result.operation != DiffMatchPatch.Operation.EQUAL)
                {
                    System.IO.File.WriteAllText(tocPath, docBuilder.ToString());
                }
            }
        }

        private void GenerateMSDNLandingPage(string landingPagePath)
        {
            var originalLandingPageMd = string.Empty;
            var newLandingPageMd = string.Empty;

            if (System.IO.File.Exists(landingPagePath))
            {
                originalLandingPageMd = System.IO.File.ReadAllText(landingPagePath);
            }
            var docBuilder = new StringBuilder();

            // read base file from disk
            var assemblyPath = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;

            string baseLandingPage = System.IO.File.ReadAllText(Path.Combine(assemblyPath, "landingpage.md"));

            // Get all unique categories
            var categories = _cmdlets.Where(c => !string.IsNullOrEmpty(c.Category)).Select(c => c.Category).Distinct();

            foreach (var category in categories.OrderBy(c => c))
            {
                docBuilder.Append("\n\n");
                docBuilder.AppendFormat("### {0} {1}", category, Environment.NewLine);
                docBuilder.AppendFormat("Cmdlet|Description|Platform{0}", Environment.NewLine);
                docBuilder.AppendFormat(":-----|:----------|:-------{0}", Environment.NewLine);

                var categoryCmdlets = _cmdlets.Where(c => c.Category == category).OrderBy(c => c.Noun);

                foreach (var cmdletInfo in categoryCmdlets)
                {
                    var description = cmdletInfo.Description != null ? cmdletInfo.Description.Replace("\r\n", " ") : "";
                    docBuilder.AppendFormat("**[{0}](../../sharepoint-ps/sharepoint-pnp/{1}-{2}.md)** |{3}|{4}{5}", cmdletInfo.FullCommand.Replace("-", "&#8209;"), cmdletInfo.Verb, cmdletInfo.Noun, description, cmdletInfo.Platform, Environment.NewLine);
                }
            }

            string dynamicLandingPage = docBuilder.ToString();
            newLandingPageMd = baseLandingPage.Replace("---cmdletdata---", dynamicLandingPage);

            DiffMatchPatch.diff_match_patch dmp = new DiffMatchPatch.diff_match_patch();

            var diffResults = dmp.diff_main(newLandingPageMd, originalLandingPageMd);

            foreach (var result in diffResults)
            {
                if (result.operation != DiffMatchPatch.Operation.EQUAL)
                {
                    System.IO.File.WriteAllText(landingPagePath, newLandingPageMd);
                }
            }
        }

        private void GenerateMSDNCategory(string category, string categoryMdPath, IOrderedEnumerable<Model.CmdletInfo> cmdlets)
        {
            var originalCategoryMd = string.Empty;
            var newCategoryMd = string.Empty;

            if (System.IO.File.Exists(categoryMdPath))
            {
                originalCategoryMd = System.IO.File.ReadAllText(categoryMdPath);
            }
            var docBuilder = new StringBuilder();
            docBuilder.AppendFormat("# {0} {1}", category, Environment.NewLine);
            docBuilder.AppendFormat("Cmdlet|Description|Platform{0}", Environment.NewLine);
            docBuilder.AppendFormat(":-----|:----------|:-------{0}", Environment.NewLine);
            foreach (var cmdletInfo in cmdlets)
            {
                var description = cmdletInfo.Description != null ? cmdletInfo.Description.Replace("\r\n", " ") : "";
                docBuilder.AppendFormat("**[{0}]({1}-{2}.md)** |{3}|{4}{5}", cmdletInfo.FullCommand.Replace("-", "&#8209;"), cmdletInfo.Verb, cmdletInfo.Noun, description, cmdletInfo.Platform, Environment.NewLine);
            }

            newCategoryMd = docBuilder.ToString();
            DiffMatchPatch.diff_match_patch dmp = new DiffMatchPatch.diff_match_patch();

            var diffResults = dmp.diff_main(newCategoryMd, originalCategoryMd);

            foreach (var result in diffResults)
            {
                if (result.operation != DiffMatchPatch.Operation.EQUAL)
                {
                    System.IO.File.WriteAllText(categoryMdPath, docBuilder.ToString());
                }
            }
        }


    }
}
