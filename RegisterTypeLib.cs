// Copyright (c) 2025 Roger Brown.
// Licensed under the MIT License.

using System;
using System.ComponentModel;
using System.Management.Automation;
using System.Runtime.InteropServices;

namespace RhubarbGeekNz.TypeLib
{
    [Cmdlet(VerbsLifecycle.Register, "TypeLib")]
    sealed public class RegisterTypeLib : PSCmdlet
    {
        [Parameter(Position = 0, ParameterSetName = "path", Mandatory = true)]
        public string[] Path;
        [Parameter(ParameterSetName = "literal", Mandatory = true, ValueFromPipeline = true)]
        public string[] LiteralPath;
        [Parameter(Mandatory = false)]
        [ValidateSet("AllUsers", "CurrentUser")]
        public string Scope;
        [Parameter(Mandatory = false)]
        public string HelpDirectory;

        protected override void ProcessRecord()
        {
            switch (ParameterSetName)
            {
                case "path":
                    foreach (string path in Path)
                    {
                        try
                        {
                            var paths = GetResolvedProviderPathFromPSPath(path, out var providerPath);

                            if ("FileSystem".Equals(providerPath.Name))
                            {
                                if (paths.Count == 0)
                                {
                                    Exception ex = new Win32Exception(OleAut32.FILE_NOT_FOUND);
                                    WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.ResourceUnavailable, path));
                                }
                                else
                                {
                                    foreach (string item in paths)
                                    {
                                        RegisterPath(item);
                                    }
                                }
                            }
                            else
                            {
                                WriteError(new ErrorRecord(new Exception($"Provider {providerPath.Name} not handled"), "ProviderError", ErrorCategory.NotImplemented, providerPath));
                            }
                        }
                        catch (ItemNotFoundException ex)
                        {
                            WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.ResourceUnavailable, path));
                        }
                    }
                    break;

                case "literal":
                    foreach (string literalPath in LiteralPath)
                    {
                        try
                        {
                            RegisterPath(GetUnresolvedProviderPathFromPSPath(literalPath));
                        }
                        catch (ItemNotFoundException ex)
                        {
                            WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.ResourceUnavailable, literalPath));
                        }
                    }
                    break;

                default:
                    NotImplementedException ni = new NotImplementedException();
                    WriteError(new ErrorRecord(ni, ni.GetType().Name, ErrorCategory.NotImplemented, ParameterSetName));
                    break;
            }
        }

        private void RegisterPath(string path)
        {
            try
            {
                OleAut32.LoadTypeLibEx(path, REGKIND.REGKIND_NONE, out var typeLib);

                if (Scope == null || 0 == String.Compare(Scope, "CurrentUser", true))
                {
                    OleAut32.RegisterTypeLibForUser(typeLib, path, HelpDirectory);
                }
                else
                {
                    OleAut32.RegisterTypeLib(typeLib, path, HelpDirectory);
                }
            }
            catch (COMException ex)
            {
                WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidResult, path));
            }
        }
    }
}
