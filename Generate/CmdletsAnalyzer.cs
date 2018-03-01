﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using System.Runtime.Serialization;
using SharePointPnP.PowerShell.CmdletHelpAttributes;
using SharePointPnP.PowerShell.ModuleFilesGenerator.Model;
using CmdletInfo = SharePointPnP.PowerShell.ModuleFilesGenerator.Model.CmdletInfo;
using System.ComponentModel;
using Generate.Model;

namespace SharePointPnP.PowerShell.ModuleFilesGenerator
{
    internal class CmdletsAnalyzer
    {
        private readonly Assembly _assembly;
        private string _platform;

        internal CmdletsAnalyzer(Assembly assembly, string platform)
        {
            _assembly = assembly;
            _platform = platform;
        }

        internal List<CmdletInfo> Analyze()
        {

            return GetCmdlets();

        }
        private List<CmdletInfo> GetCmdlets()
        {
            List<CmdletInfo> cmdlets = new List<CmdletInfo>();
            var types = _assembly.GetTypes().Where(t => t.BaseType != null && (t.BaseType.Name.StartsWith("SPO") || t.BaseType.Name.StartsWith("PnP") || t.BaseType.Name == "PSCmdlet" || (t.BaseType.BaseType != null && (t.BaseType.BaseType.Name.StartsWith("PnP") || t.BaseType.BaseType.Name == "PSCmdlet")))).OrderBy(t => t.Name).ToArray();

            foreach (var type in types)
            {
                var cmdletInfo = new Model.CmdletInfo();
                cmdletInfo.CmdletType = type;

                var attributes = type.GetCustomAttributes();

                foreach (var attribute in attributes)
                {
                    var cmdletAttribute = attribute as CmdletAttribute;
                    if (cmdletAttribute != null)
                    {
#if !NETCOREAPP2_0
                        var a = cmdletAttribute;
                        cmdletInfo.Verb = a.VerbName;
                        cmdletInfo.Noun = a.NounName;
#else
                        var customAttributesData = type.GetCustomAttributesData();
                        var customAttributeData = customAttributesData.FirstOrDefault(c => c.AttributeType == typeof(CmdletAttribute));
                        if (customAttributeData != null)
                        {
                            cmdletInfo.Verb = customAttributeData.ConstructorArguments[0].Value.ToString();
                            cmdletInfo.Noun = customAttributeData.ConstructorArguments[1].Value.ToString();
                        }
#endif
                    }
                    var aliasAttribute = attribute as AliasAttribute;
                    if (aliasAttribute != null)
                    {
#if !NETCOREAPP2_0
                        foreach (var name in aliasAttribute.AliasNames)
                        {
                            cmdletInfo.Aliases.Add(name);
                        }
#else
                        var customAttributeData = type.GetCustomAttributesData().FirstOrDefault(c => c.AttributeType == typeof(AliasAttribute));
                        if (customAttributeData != null)
                        {
                            foreach (var name in customAttributeData.ConstructorArguments)
                            {
                                cmdletInfo.Aliases.Add(name.Value as string);
                            }
                        }
#endif
                    }

                    var helpAttribute = attribute as CmdletHelpAttribute;
                    if (helpAttribute != null)
                    {
                        var a = helpAttribute;
                        cmdletInfo.Description = a.Description;
                        cmdletInfo.Copyright = a.Copyright;
                        cmdletInfo.Version = a.Version;
                        cmdletInfo.DetailedDescription = a.DetailedDescription;
                        cmdletInfo.Category = ToEnumString(a.Category);
                        cmdletInfo.OutputType = a.OutputType;
                        cmdletInfo.OutputTypeLink = a.OutputTypeLink;
                        cmdletInfo.OutputTypeDescription = a.OutputTypeDescription;

                        List<string> platforms = new List<string>();
                        if (a.SupportedPlatform.HasFlag(CmdletSupportedPlatform.All))
                        {
                            platforms.Add("SharePoint Server 2013");
                            platforms.Add("SharePoint Server 2016");
                            platforms.Add("SharePoint Online");
                        }
                        if (a.SupportedPlatform.HasFlag(CmdletSupportedPlatform.OnPremises))
                        {
                            platforms.Add("SharePoint Server 2013");
                            platforms.Add("SharePoint Server 2016");
                        }
                        if (a.SupportedPlatform.HasFlag(CmdletSupportedPlatform.Online))
                        {
                            platforms.Add("SharePoint Online");
                        }
                        if (a.SupportedPlatform.HasFlag(CmdletSupportedPlatform.SP2013))
                        {
                            platforms.Add("SharePoint 2013");
                        }
                        if (a.SupportedPlatform.HasFlag(CmdletSupportedPlatform.SP2016))
                        {
                            platforms.Add("SharePoint 2016");
                        }
                        cmdletInfo.Platform = string.Join(", ", platforms);
                    }
                    var exampleAttribute = attribute as CmdletExampleAttribute;
                    if (exampleAttribute != null)
                    {
                        cmdletInfo.Examples.Add(exampleAttribute);
                    }
                    var linkAttribute = attribute as CmdletRelatedLinkAttribute;
                    if (linkAttribute != null)
                    {
                        cmdletInfo.RelatedLinks.Add(linkAttribute);
                    }
                    var additionalParameter = attribute as CmdletAdditionalParameter;
                    if (additionalParameter != null)
                    {
                        cmdletInfo.AdditionalParameters.Add(additionalParameter);
                    }
                }
                if (!string.IsNullOrEmpty(cmdletInfo.Verb) && !string.IsNullOrEmpty(cmdletInfo.Noun))
                {
                    cmdletInfo.Syntaxes = GetCmdletSyntaxes(cmdletInfo);
                    cmdletInfo.Parameters = GetCmdletParameters(cmdletInfo);
                    cmdlets.Add(cmdletInfo);
                }
            }

            return cmdlets;
        }

        private List<CmdletSyntax> GetCmdletSyntaxes(Model.CmdletInfo cmdletInfo)
        {
            List<CmdletSyntax> syntaxes = new List<CmdletSyntax>();
            var fields = GetFields(cmdletInfo.CmdletType);
            foreach (var field in fields)
            {
                MemberInfo fieldInfo = field;
                var obsolete = fieldInfo.GetCustomAttributes<ObsoleteAttribute>().Any();

                if (!obsolete)
                {
                    //                 var parameterAttributes = fieldInfo.GetCustomAttributes<ParameterAttribute>(true).Where(a => a.ParameterSetName != ParameterAttribute.AllParameterSets);
                    //               var pnpAttributes = field.GetCustomAttributes<PnPParameterAttribute>(true);
                    //                    foreach (var parameterAttribute in parameterAttributes)
                    //                   {
                    var customAttributesData = field.GetCustomAttributesData();
                    var customAttributeData = customAttributesData.Where(c => c.AttributeType == typeof(ParameterAttribute));
                    var parameterAttributes = customAttributeData.Where(c => c.NamedArguments.Any(n => n.MemberName == "ParameterSetName"));

                    var pnpAttributes = field.GetCustomAttributes<PnPParameterAttribute>(true);
                    foreach (var parameterAttribute in parameterAttributes.Where(c => (string)c.NamedArguments.First(n => n.MemberName == "ParameterSetName").TypedValue.Value != ParameterAttribute.AllParameterSets))
                    {

                        var parameterSetName = parameterAttribute.GetAttributeValue<string>("ParameterSetName");
                        var helpMessage = parameterAttribute.GetAttributeValue<string>("HelpMessage");
                        var position = parameterAttribute.GetAttributeValue<int>("Position");
                        var mandatory = parameterAttribute.GetAttributeValue<bool>("Mandatory");
                        var cmdletSyntax = syntaxes.FirstOrDefault(c => c.ParameterSetName == parameterSetName);
                        if (cmdletSyntax == null)
                        {
                            cmdletSyntax = new CmdletSyntax();
                            cmdletSyntax.ParameterSetName = parameterSetName;
                            syntaxes.Add(cmdletSyntax);
                        }
                        var typeString = field.FieldType.Name;
                        if (field.FieldType.IsGenericType)
                        {
                            typeString = field.FieldType.GenericTypeArguments[0].Name;
                        }
                        var fieldAttribute = field.FieldType.GetCustomAttributes<CmdletPipelineAttribute>(false).FirstOrDefault();
                        if (fieldAttribute != null)
                        {
                            if (fieldAttribute.Type != null)
                            {
                                typeString = string.Format(fieldAttribute.Description, fieldAttribute.Type.Name);
                            }
                            else
                            {
                                typeString = fieldAttribute.Description;
                            }
                        }
                        var order = 0;
                        if (pnpAttributes != null && pnpAttributes.Any())
                        {
                            order = pnpAttributes.First().Order;
                        }
                        cmdletSyntax.Parameters.Add(new CmdletParameterInfo()
                        {
                            Name = field.Name,
                            Description = helpMessage,
                            Position = position,
                            Required = mandatory,
                            Type = typeString,
                            Order = order,
                        });
                    }
                }
            }

            foreach (var additionalParameter in cmdletInfo.AdditionalParameters.Where(a => a.ParameterSetName != ParameterAttribute.AllParameterSets))
            {
                var cmdletSyntax = syntaxes.FirstOrDefault(c => c.ParameterSetName == additionalParameter.ParameterSetName);
                if (cmdletSyntax == null)
                {
                    cmdletSyntax = new CmdletSyntax();
                    cmdletSyntax.ParameterSetName = additionalParameter.ParameterSetName;
                    syntaxes.Add(cmdletSyntax);
                }
                var typeString = additionalParameter.ParameterType.Name;
                if (additionalParameter.ParameterType.IsGenericType)
                {
                    typeString = additionalParameter.ParameterType.GenericTypeArguments[0].Name;
                }
                var fieldAttribute = additionalParameter.ParameterType.GetCustomAttributes<CmdletPipelineAttribute>(false).FirstOrDefault();
                if (fieldAttribute != null)
                {
                    if (fieldAttribute.Type != null)
                    {
                        typeString = string.Format(fieldAttribute.Description, fieldAttribute.Type.Name);
                    }
                    else
                    {
                        typeString = fieldAttribute.Description;
                    }
                }
                cmdletSyntax.Parameters.Add(new CmdletParameterInfo()
                {
                    Name = additionalParameter.ParameterName,
                    Description = additionalParameter.HelpMessage,
                    Position = additionalParameter.Position,
                    Required = additionalParameter.Mandatory,
                    Type = typeString,
                    Order = additionalParameter.Order,
                });
            }

            // AllParameterSets
            foreach (var field in fields)
            {
                var obsolete = field.GetCustomAttributes<ObsoleteAttribute>().Any();

                if (!obsolete)
                {

                    //  var tempType = fieldGetCustomAttributes(typeof(ParameterAttribute));
                    var customAttributesData = field.GetCustomAttributesData();
                    var customAttributeData = customAttributesData.Where(c => c.AttributeType == typeof(ParameterAttribute));
                    var parameterAttributes = customAttributeData.Where(c => c.NamedArguments.Any(n => n.MemberName == "ParameterSetName")).Where(p => (string)p.NamedArguments.First(n => n.MemberName == "ParameterSetName").TypedValue.Value == ParameterAttribute.AllParameterSets).ToList();
                    var parameterAttributes1 = customAttributeData.Where(c => c.NamedArguments.Count(n => n.MemberName == "ParameterSetName") == 0).ToList();
                    parameterAttributes.AddRange(parameterAttributes1);
                    var pnpAttributes = field.GetCustomAttributes<PnPParameterAttribute>(true);
                    foreach (var parameterAttribute in parameterAttributes)
                    {
                        var helpMessage = parameterAttribute.GetAttributeValue<string>("HelpMessage");
                        var position = parameterAttribute.GetAttributeValue<int>("Position");
                        var mandatory = parameterAttribute.GetAttributeValue<bool>("Mandatory");

                        if (!syntaxes.Any())
                        {
                            syntaxes.Add(new CmdletSyntax { ParameterSetName = ParameterAttribute.AllParameterSets });
                        }

                        foreach (var syntax in syntaxes)
                        {
                            var typeString = field.FieldType.Name;
                            if (field.FieldType.IsGenericType)
                            {
                                typeString = field.FieldType.GenericTypeArguments[0].Name;
                            }
                            var fieldAttribute = field.FieldType.GetCustomAttributes<CmdletPipelineAttribute>(false).FirstOrDefault();
                            if (fieldAttribute != null)
                            {
                                if (fieldAttribute.Type != null)
                                {
                                    typeString = string.Format(fieldAttribute.Description, fieldAttribute.Type.Name);
                                }
                                else
                                {
                                    typeString = fieldAttribute.Description;
                                }
                            }
                            var order = 0;
                            if (pnpAttributes != null && pnpAttributes.Any())
                            {
                                order = pnpAttributes.First().Order;
                            }
                            syntax.Parameters.Add(new CmdletParameterInfo()
                            {
                                Name = field.Name,
                                Description = helpMessage,
                                Position = position,
                                Required = mandatory,
                                Type = typeString,
                                Order = order,
                                Platform = cmdletInfo.Platform
                            });
                        }
                    }
                }
            }
            return syntaxes;
        }

        private List<CmdletParameterInfo> GetCmdletParameters(Model.CmdletInfo cmdletInfo)
        {
            List<CmdletParameterInfo> parameters = new List<CmdletParameterInfo>();
            var fields = GetFields(cmdletInfo.CmdletType);
            foreach (var field in fields)
            {
                MemberInfo fieldInfo = field;
                var obsolete = fieldInfo.GetCustomAttributes<ObsoleteAttribute>().Any();

                if (!obsolete)
                {
                    var aliases = fieldInfo.GetCustomAttributes<AliasAttribute>(true);
                    //var parameterAttributes = fieldInfo.GetCustomAttributes<ParameterAttribute>(true);
                    var pnpParameterAttributes = fieldInfo.GetCustomAttributes<PnPParameterAttribute>(true);

                    var parameterAttributes = fieldInfo.GetCustomAttributesData().Where(c => c.AttributeType == typeof(ParameterAttribute));
                    foreach (var parameterAttribute in parameterAttributes)
                    {

                        var description = parameterAttribute.GetAttributeValue<string>("HelpMessage");
                        if (string.IsNullOrEmpty(description))
                        {
                            // Maybe a generic one? Find the one with only a helpmessage set
                            var helpParameterAttribute = parameterAttributes.FirstOrDefault(p => !string.IsNullOrEmpty(p.GetAttributeValue<string>("HelpMessage")));
                            if (helpParameterAttribute != null)
                            {
                                description = helpParameterAttribute.GetAttributeValue<string>("HelpMessage");
                            }
                        }
                        var typeString = field.FieldType.Name;
                        if (field.FieldType.IsGenericType)
                        {
                            typeString = field.FieldType.GenericTypeArguments[0].Name;
                        }
                        var fieldAttribute = field.FieldType.GetCustomAttributes<CmdletPipelineAttribute>(false).FirstOrDefault();
                        if (fieldAttribute != null)
                        {
                            if (fieldAttribute.Type != null)
                            {
                                typeString = string.Format(fieldAttribute.Description, fieldAttribute.Type.Name);
                            }
                            else
                            {
                                typeString = fieldAttribute.Description;
                            }
                        }
                        var order = 0;
                        if (pnpParameterAttributes != null && pnpParameterAttributes.Any())
                        {
                            order = pnpParameterAttributes.First().Order;
                        }
                        var existingParameter = parameters.FirstOrDefault(p => p.Name == field.Name);
                        if (existingParameter == null)
                        {
                            var cmdletParameterInfo = new CmdletParameterInfo()
                            {
                                Description = description,
                                Type = typeString,
                                Name = field.Name,
                                Required = parameterAttribute.GetAttributeValue<bool>("Mandatory"),
                                Position = parameterAttribute.GetAttributeValue<int>("Position"),
                                ValueFromPipeline = parameterAttribute.GetAttributeValue<bool>("ValueFromPipeline"),
                                ParameterSetName = parameterAttribute.GetAttributeValue<string>("ParameterSetName"),
                                Order = order,
                                Platform = _platform
                            };

                            if (aliases != null && aliases.Any())
                            {
#if !NETCOREAPP2_0
                            foreach (var aliasAttribute in aliases)
                            {
                                cmdletParameterInfo.Aliases.AddRange(aliasAttribute.AliasNames);
                            }
#else
                                var customAttributesData = fieldInfo.GetCustomAttributesData();
                                foreach (var aliasAttribute in customAttributesData.Where(c => c.AttributeType == typeof(AliasAttribute)))
                                {
                                    cmdletParameterInfo.Aliases.AddRange(aliasAttribute.ConstructorArguments.Select(a => a.ToString()));
                                }
#endif
                            }
                            parameters.Add(cmdletParameterInfo);
                        }
                        else
                        {
                            if (existingParameter.ParameterSetName != null)
                            {
                                var parameterSetNames = existingParameter.ParameterSetName.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries).ToList();
                                if (!parameterSetNames.Contains(parameterAttribute.GetAttributeValue<string>("ParameterSetName")))
                                {
                                    parameterSetNames.Add(parameterAttribute.GetAttributeValue<string>("ParameterSetName"));
                                }
                                existingParameter.ParameterSetName = string.Join(", ", parameterSetNames);
                            }
                            else
                            {
                                existingParameter.ParameterSetName = parameterAttribute.GetAttributeValue<string>("ParameterSetName");
                            }
                        }
                    }
                }
            }

            foreach (var additionalParameter in cmdletInfo.AdditionalParameters)
            {
                var typeString = additionalParameter.ParameterType.Name;
                if (additionalParameter.ParameterType.IsGenericType)
                {
                    typeString = additionalParameter.ParameterType.GenericTypeArguments[0].Name;
                }
                var fieldAttribute = additionalParameter.ParameterType.GetCustomAttributes<CmdletPipelineAttribute>(false).FirstOrDefault();
                if (fieldAttribute != null)
                {
                    if (fieldAttribute.Type != null)
                    {
                        typeString = string.Format(fieldAttribute.Description, fieldAttribute.Type.Name);
                    }
                    else
                    {
                        typeString = fieldAttribute.Description;
                    }
                }
                parameters.Add(new CmdletParameterInfo()
                {
                    Description = additionalParameter.HelpMessage,
                    Type = typeString,
                    Name = additionalParameter.ParameterName,
                    Required = additionalParameter.Mandatory,
                    Position = additionalParameter.Position,
                    ParameterSetName = additionalParameter.ParameterSetName,
                    Platform = _platform
                });
            }

            return parameters;
        }



        #region Helpers
        private static List<FieldInfo> GetFields(Type t)
        {
            var fieldInfoList = new List<FieldInfo>();
            foreach (var fieldInfo in t.GetFields())
            {
                fieldInfoList.Add(fieldInfo);
            }
            if (t.BaseType != null && t.BaseType.BaseType != null)
            {
                fieldInfoList.AddRange(GetFields(t.BaseType.BaseType));
            }
            return fieldInfoList;
        }

        private static string ToEnumString<T>(T type)
        {
            var enumType = typeof(T);
            var name = Enum.GetName(enumType, type);
            try
            {
                var enumMemberAttribute = ((EnumMemberAttribute[])enumType.GetField(name).GetCustomAttributes(typeof(EnumMemberAttribute), true)).Single();
                return enumMemberAttribute.Value;
            }
            catch
            {
                return name;
            }
        }
        #endregion
    }
}


//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Management.Automation;
//using System.Reflection;
//using System.Runtime.Serialization;
//using SharePointPnP.PowerShell.CmdletHelpAttributes;
//using SharePointPnP.PowerShell.ModuleFilesGenerator.Model;
//using CmdletInfo = SharePointPnP.PowerShell.ModuleFilesGenerator.Model.CmdletInfo;
//using System.ComponentModel;
//using Generate.Model;
//using Mono.Cecil;

//namespace SharePointPnP.PowerShell.ModuleFilesGenerator
//{
//    internal class CmdletsAnalyzer
//    {
//        private readonly AssemblyDefinition _assembly;

//        internal CmdletsAnalyzer(AssemblyDefinition assembly)
//        {
//            _assembly = assembly;
//        }

//        internal List<CmdletInfo> Analyze()
//        {

//            return GetCmdlets();

//        }
//        private List<CmdletInfo> GetCmdlets()
//        {
//            List<CmdletInfo> cmdlets = new List<CmdletInfo>();
//            var types = _assembly.MainModule.GetTypes().Where(t => t.BaseType != null && (t.BaseType.Name.StartsWith("SPO") || t.BaseType.Name.StartsWith("PnP") || t.BaseType.Name == "PSCmdlet")).OrderBy(t => t.Name).ToArray();

//            foreach (var type in types)
//            {
//                var cmdletInfo = new Model.CmdletInfo();
//                cmdletInfo.CmdletType = type.GetType();

//                var attributes = type.CustomAttributes;

//                foreach (var attribute in attributes)
//                {
//                    if (attribute.AttributeType.Name == "CmdletAttribute")
//                    {
//                        cmdletInfo.Verb = attribute.ConstructorArguments[0].Value as string;
//                        cmdletInfo.Noun = attribute.ConstructorArguments[1].Value as string;
//                        //var customAttributesData = type.GetCustomAttributesData();
//                        //var customAttributeData = customAttributesData.FirstOrDefault(c => c.AttributeType == typeof(CmdletAttribute));
//                        //if (customAttributeData != null)
//                        //{
//                        //    cmdletInfo.Verb = customAttributeData.ConstructorArguments[0].Value.ToString();
//                        //    cmdletInfo.Noun = customAttributeData.ConstructorArguments[1].Value.ToString();
//                        //}
//                    }
//                    if (attribute.AttributeType.Name == "CmdletAliasAttribute")
//                    {
//                        cmdletInfo.Aliases.Add(attribute.Properties.FirstOrDefault(p => p.Name == "Alias").Argument.Value as string);
//                    }
//                    if (attribute.AttributeType.Name == "CmdletHelpAttribute")
//                    {
//                        cmdletInfo.Description = attribute.Properties.FirstOrDefault(p => p.Name == "Description").Argument.Value as string;
//                        cmdletInfo.Copyright = attribute.Properties.FirstOrDefault(p => p.Name == "Copyright").Argument.Value as string;
//                        cmdletInfo.Version = attribute.Properties.FirstOrDefault(p => p.Name == "Version").Argument.Value as string;
//                        cmdletInfo.DetailedDescription = attribute.Properties.FirstOrDefault(p => p.Name == "DetailedDescription").Argument.Value as string;
//                        cmdletInfo.Category = ToEnumString((CmdletHelpCategory)attribute.Properties.FirstOrDefault(p => p.Name == "Category").Argument.Value);
//                        cmdletInfo.OutputType = attribute.Properties.FirstOrDefault(p => p.Name == "OutputType").Argument.Value as Type;
//                        cmdletInfo.OutputTypeLink = attribute.Properties.FirstOrDefault(p => p.Name == "OutputTypeLink").Argument.Value as string;
//                        cmdletInfo.OutputTypeDescription = attribute.Properties.FirstOrDefault(p => p.Name == "OutputTypeDescription").Argument.Value as string;

//                        if (attribute.Properties.Any(p => p.Name == "SupportedPlatform"))
//                        {
//                            var supportedPlatform = (CmdletSupportedPlatform)attribute.Properties.FirstOrDefault(p => p.Name == "SupportedPlatform").Argument.Value;
//                            if (supportedPlatform.HasFlag(CmdletSupportedPlatform.All))
//                            {
//                                cmdletInfo.Platform = "All";
//                            }
//                            else
//                            {
//                                List<string> platforms = new List<string>();
//                                if (supportedPlatform.HasFlag(CmdletSupportedPlatform.OnPremises))
//                                {
//                                    platforms.Add("SharePoint On-Premises");
//                                }
//                                if (supportedPlatform.HasFlag(CmdletSupportedPlatform.Online))
//                                {
//                                    platforms.Add("SharePoint Online");
//                                }
//                                if (supportedPlatform.HasFlag(CmdletSupportedPlatform.SP2013))
//                                {
//                                    platforms.Add("SharePoint 2013");
//                                }
//                                if (supportedPlatform.HasFlag(CmdletSupportedPlatform.SP2016))
//                                {
//                                    platforms.Add("SharePoint 2016");
//                                }
//                                cmdletInfo.Platform = string.Join(", ", platforms);
//                            }
//                        }
//                    }
//                    if (attribute.AttributeType.Name == "CmdletExampleAttribute")
//                    {
//                        var cmdletExample = new CmdletExampleAttribute();

//                        var code = attribute.Properties.FirstOrDefault(p => p.Name == "Code").Argument.Value as string;
//                        var introduction = attribute.Properties.FirstOrDefault(p => p.Name == "Introduction").Argument.Value as string;
//                        var remarks = attribute.Properties.FirstOrDefault(p => p.Name == "Remarks").Argument.Value as string;
//                        var sortOrder = (int)attribute.Properties.FirstOrDefault(p => p.Name == "SortOrder").Argument.Value;
//                        cmdletExample.Code = code as string;
//                        cmdletExample.Introduction = introduction as string;
//                        cmdletExample.Remarks = remarks as string;
//                        cmdletExample.SortOrder = (int)sortOrder;
//                        cmdletInfo.Examples.Add(cmdletExample);
//                    }
//                    //var linkAttribute = attribute as CmdletRelatedLinkAttribute;
//                    //if (linkAttribute != null)
//                    //{
//                    //    cmdletInfo.RelatedLinks.Add(linkAttribute);
//                    //}
//                    if (attribute.AttributeType.Name == "CmdletAdditionalParameter")
//                    {
//                        var additionalParameter = new CmdletAdditionalParameter();
//                        additionalParameter.HelpMessage = GetPropertyValue<string>(attribute, "HelpMessage");
//                        additionalParameter.Mandatory = GetPropertyValue<bool>(attribute, "Mandatory");
//                        additionalParameter.Order = GetPropertyValue<int>(attribute, "Order");
//                        additionalParameter.ParameterName = GetPropertyValue<string>(attribute, "ParameterName");
//                        additionalParameter.ParameterSetName = GetPropertyValue<string>(attribute, "ParameterSetName");
//                        additionalParameter.ParameterType = GetPropertyValueAsType(attribute, "ParameterType");
//                        cmdletInfo.AdditionalParameters.Add(additionalParameter);
//                    }
//                }
//                if (!string.IsNullOrEmpty(cmdletInfo.Verb) && !string.IsNullOrEmpty(cmdletInfo.Noun))
//                {
//                    cmdletInfo.Syntaxes = GetCmdletSyntaxes(cmdletInfo);
//                    cmdletInfo.Parameters = GetCmdletParameters(cmdletInfo);
//                    cmdlets.Add(cmdletInfo);
//                }
//            }

//            return cmdlets;
//        }

//        private Type GetPropertyValueAsType(CustomAttribute attribute, string property)
//        {
//            for (var q = 0; q < attribute.Properties.Count; q++)
//            {
//                if (attribute.Properties[q].Name == property)
//                {
//                    return attribute.Properties[q].Argument.Value as Type;
//                }
//            }
//            return null;
//        }

//        private T GetPropertyValue<T>(CustomAttribute attribute, string property)
//        {
//            for (var q = 0; q < attribute.Properties.Count; q++)
//            {
//                if (attribute.Properties[q].Name == property)
//                {
//                    return (T)attribute.Properties[q].Argument.Value;
//                }
//            }
//            return default(T);
//        }

//        private List<CmdletSyntax> GetCmdletSyntaxes(Model.CmdletInfo cmdletInfo)
//        {
//            List<CmdletSyntax> syntaxes = new List<CmdletSyntax>();
//            var fields = GetFields(cmdletInfo.CmdletType);
//            foreach (var field in fields)
//            {
//                MemberInfo fieldInfo = field;
//                var obsolete = fieldInfo.GetCustomAttributes<ObsoleteAttribute>().Any();

//                if (!obsolete)
//                {
//                    var customAttributesData = fieldInfo.GetCustomAttributesData();
//                    var customAttributeData = customAttributesData.Where(c => c.AttributeType == typeof(ParameterAttribute));
//                    var parameterAttributes = customAttributeData.Where(c => c.NamedArguments.Any(n => n.MemberName == "ParameterSetName"));

//                    var pnpAttributes = field.GetCustomAttributes<PnPParameterAttribute>(true);
//                    foreach (var parameterAttribute in parameterAttributes.Where(c => (string)c.NamedArguments.First(n => n.MemberName == "ParameterSetName").TypedValue.Value != ParameterAttribute.AllParameterSets))
//                    {
//                        var parameterSetName = parameterAttribute.GetAttributeValue<string>("ParameterSetName");
//                        var helpMessage = parameterAttribute.GetAttributeValue<string>("HelpMessage");
//                        var position = parameterAttribute.GetAttributeValue<int>("Position");
//                        var mandatory = parameterAttribute.GetAttributeValue<bool>("Mandatory");
//                        var cmdletSyntax = syntaxes.FirstOrDefault(c => c.ParameterSetName == parameterSetName);
//                        if (cmdletSyntax == null)
//                        {
//                            cmdletSyntax = new CmdletSyntax();
//                            cmdletSyntax.ParameterSetName = parameterSetName;
//                            syntaxes.Add(cmdletSyntax);
//                        }
//                        var typeString = field.FieldType.Name;
//                        var fieldAttribute = field.FieldType.GetCustomAttributes<CmdletPipelineAttribute>(false).FirstOrDefault();
//                        if (fieldAttribute != null)
//                        {
//                            if (fieldAttribute.Type != null)
//                            {
//                                typeString = string.Format(fieldAttribute.Description, fieldAttribute.Type.Name);
//                            }
//                            else
//                            {
//                                typeString = fieldAttribute.Description;
//                            }
//                        }
//                        var order = 0;
//                        if (pnpAttributes != null && pnpAttributes.Any())
//                        {
//                            order = pnpAttributes.First().Order;
//                        }
//                        cmdletSyntax.Parameters.Add(new CmdletParameterInfo()
//                        {
//                            Name = field.Name,
//                            Description = helpMessage,
//                            Position = position,
//                            Required = mandatory,
//                            Type = typeString,
//                            Order = order
//                        });
//                    }
//                }
//            }

//            foreach (var additionalParameter in cmdletInfo.AdditionalParameters.Where(a => a.ParameterSetName != ParameterAttribute.AllParameterSets))
//            {
//                var cmdletSyntax = syntaxes.FirstOrDefault(c => c.ParameterSetName == additionalParameter.ParameterSetName);
//                if (cmdletSyntax == null)
//                {
//                    cmdletSyntax = new CmdletSyntax();
//                    cmdletSyntax.ParameterSetName = additionalParameter.ParameterSetName;
//                    syntaxes.Add(cmdletSyntax);
//                }
//                var typeString = additionalParameter.ParameterType.Name;
//                var fieldAttribute = additionalParameter.ParameterType.GetCustomAttributes<CmdletPipelineAttribute>(false).FirstOrDefault();
//                if (fieldAttribute != null)
//                {
//                    if (fieldAttribute.Type != null)
//                    {
//                        typeString = string.Format(fieldAttribute.Description, fieldAttribute.Type.Name);
//                    }
//                    else
//                    {
//                        typeString = fieldAttribute.Description;
//                    }
//                }
//                cmdletSyntax.Parameters.Add(new CmdletParameterInfo()
//                {
//                    Name = additionalParameter.ParameterName,
//                    Description = additionalParameter.HelpMessage,
//                    Position = additionalParameter.Position,
//                    Required = additionalParameter.Mandatory,
//                    Type = typeString,
//                    Order = additionalParameter.Order
//                });
//            }

//            // AllParameterSets
//            foreach (var field in fields)
//            {
//                var obsolete = field.GetCustomAttributes<ObsoleteAttribute>().Any();

//                if (!obsolete)
//                {
//                    var customAttributesData = field.GetCustomAttributesData();
//                    var customAttributeData = customAttributesData.Where(c => c.AttributeType == typeof(ParameterAttribute));
//                    //                    var parameterAttributes = customAttributeData.Where(c => c.NamedArguments.Any(n => n.MemberName == "ParameterSetName") );



//                    //var parameterAttributes = field.GetCustomAttributes<ParameterAttribute>(true).Where(a => a.ParameterSetName == ParameterAttribute.AllParameterSets);
//                    var pnpAttributes = field.GetCustomAttributes<PnPParameterAttribute>(true);
//                    foreach (var parameterAttribute in customAttributeData.Where(c =>
//                        (c.NamedArguments.FirstOrDefault(n => n.MemberName == "ParameterSetName") != null && (string)c.NamedArguments.FirstOrDefault(n => n.MemberName == "ParameterSetName").TypedValue.Value == ParameterAttribute.AllParameterSets) ||
//                        c.NamedArguments.FirstOrDefault(n => n.MemberName == "ParameterSetName") == null))
//                    {
//                        var parameterSetName = parameterAttribute.GetAttributeValue<string>("ParameterSetName");
//                        var helpMessage = parameterAttribute.GetAttributeValue<string>("HelpMessage");
//                        var position = parameterAttribute.GetAttributeValue<int>("Position");
//                        var mandatory = parameterAttribute.GetAttributeValue<bool>("Mandatory");
//                        if (!syntaxes.Any())
//                        {
//                            syntaxes.Add(new CmdletSyntax { ParameterSetName = ParameterAttribute.AllParameterSets });
//                        }

//                        foreach (var syntax in syntaxes)
//                        {
//                            var typeString = field.FieldType.Name;
//                            var fieldAttribute = field.FieldType.GetCustomAttributes<CmdletPipelineAttribute>(false).FirstOrDefault();
//                            if (fieldAttribute != null)
//                            {
//                                if (fieldAttribute.Type != null)
//                                {
//                                    typeString = string.Format(fieldAttribute.Description, fieldAttribute.Type.Name);
//                                }
//                                else
//                                {
//                                    typeString = fieldAttribute.Description;
//                                }
//                            }
//                            var order = 0;
//                            if (pnpAttributes != null && pnpAttributes.Any())
//                            {
//                                order = pnpAttributes.First().Order;
//                            }
//                            syntax.Parameters.Add(new CmdletParameterInfo()
//                            {
//                                Name = field.Name,
//                                Description = helpMessage,
//                                Position = position,
//                                Required = mandatory,
//                                Type = typeString,
//                                Order = order
//                            });
//                        }
//                    }
//                }
//            }
//            return syntaxes;
//        }

//        private List<CmdletParameterInfo> GetCmdletParameters(Model.CmdletInfo cmdletInfo)
//        {
//            List<CmdletParameterInfo> parameters = new List<CmdletParameterInfo>();
//            var fields = GetFields(cmdletInfo.CmdletType);
//            foreach (var field in fields)
//            {
//                MemberInfo fieldInfo = field;
//                var obsolete = fieldInfo.GetCustomAttributes<ObsoleteAttribute>().Any();

//                if (!obsolete)
//                {
//                    var aliases = fieldInfo.GetCustomAttributes<AliasAttribute>(true);
//                    var parameterAttributes = fieldInfo.GetCustomAttributes<ParameterAttribute>(true);
//                    var pnpParameterAttributes = fieldInfo.GetCustomAttributes<PnPParameterAttribute>(true);
//                    foreach (var parameterAttribute in parameterAttributes)
//                    {
//                        var description = parameterAttribute.HelpMessage;
//                        if (string.IsNullOrEmpty(description))
//                        {
//                            // Maybe a generic one? Find the one with only a helpmessage set
//                            var helpParameterAttribute = parameterAttributes.FirstOrDefault(p => !string.IsNullOrEmpty(p.HelpMessage));
//                            if (helpParameterAttribute != null)
//                            {
//                                description = helpParameterAttribute.HelpMessage;
//                            }
//                        }
//                        var typeString = field.FieldType.Name;
//                        var fieldAttribute = field.FieldType.GetCustomAttributes<CmdletPipelineAttribute>(false).FirstOrDefault();
//                        if (fieldAttribute != null)
//                        {
//                            if (fieldAttribute.Type != null)
//                            {
//                                typeString = string.Format(fieldAttribute.Description, fieldAttribute.Type.Name);
//                            }
//                            else
//                            {
//                                typeString = fieldAttribute.Description;
//                            }
//                        }
//                        var order = 0;
//                        if (pnpParameterAttributes != null && pnpParameterAttributes.Any())
//                        {
//                            order = pnpParameterAttributes.First().Order;
//                        }
//                        var cmdletParameterInfo = new CmdletParameterInfo()
//                        {
//                            Description = description,
//                            Type = typeString,
//                            Name = field.Name,
//                            Required = parameterAttribute.Mandatory,
//                            Position = parameterAttribute.Position,
//                            ValueFromPipeline = parameterAttribute.ValueFromPipeline,
//                            ParameterSetName = parameterAttribute.ParameterSetName,
//                            Order = order
//                        };

//                        if (aliases != null && aliases.Any())
//                        {
//#if !NETCOREAPP2_0
//                            foreach (var aliasAttribute in aliases)
//                            {
//                                cmdletParameterInfo.Aliases.AddRange(aliasAttribute.AliasNames);
//                            }
//#else
//                            var customAttributesData = fieldInfo.GetCustomAttributesData();
//                            foreach (var aliasAttribute in customAttributesData.Where(c => c.AttributeType == typeof(AliasAttribute)))
//                            {
//                                cmdletParameterInfo.Aliases.AddRange(aliasAttribute.ConstructorArguments.Select(a => a.ToString()));
//                            }
//#endif
//                        }
//                        parameters.Add(cmdletParameterInfo);

//                    }
//                }
//            }

//            foreach (var additionalParameter in cmdletInfo.AdditionalParameters)
//            {
//                var typeString = additionalParameter.ParameterType.Name;
//                var fieldAttribute = additionalParameter.ParameterType.GetCustomAttributes<CmdletPipelineAttribute>(false).FirstOrDefault();
//                if (fieldAttribute != null)
//                {
//                    if (fieldAttribute.Type != null)
//                    {
//                        typeString = string.Format(fieldAttribute.Description, fieldAttribute.Type.Name);
//                    }
//                    else
//                    {
//                        typeString = fieldAttribute.Description;
//                    }
//                }
//                parameters.Add(new CmdletParameterInfo()
//                {
//                    Description = additionalParameter.HelpMessage,
//                    Type = typeString,
//                    Name = additionalParameter.ParameterName,
//                    Required = additionalParameter.Mandatory,
//                    Position = additionalParameter.Position,
//                    ParameterSetName = additionalParameter.ParameterSetName
//                });
//            }

//            return parameters;
//        }



//        #region Helpers
//        private static List<FieldInfo> GetFields(Type t)
//        {
//            var fieldInfoList = new List<FieldInfo>();
//            foreach (var fieldInfo in t.GetFields())
//            {
//                fieldInfoList.Add(fieldInfo);
//            }
//            if (t.BaseType != null && t.BaseType.BaseType != null)
//            {
//                fieldInfoList.AddRange(GetFields(t.BaseType.BaseType));
//            }
//            return fieldInfoList;
//        }

//        private static string ToEnumString<T>(T type)
//        {
//            var enumType = typeof(T);
//            var name = Enum.GetName(enumType, type);
//            try
//            {
//                var enumMemberAttribute = ((EnumMemberAttribute[])enumType.GetField(name).GetCustomAttributes(typeof(EnumMemberAttribute), true)).Single();
//                return enumMemberAttribute.Value;
//            }
//            catch
//            {
//                return name;
//            }
//        }
//        #endregion
//    }
//}
