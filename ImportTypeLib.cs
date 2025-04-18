// Copyright (c) 2025 Roger Brown.
// Licensed under the MIT License.

using System;
using System.ComponentModel;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace RhubarbGeekNz.TypeLib
{
    [Cmdlet(VerbsData.Import, "TypeLib")]
    [OutputType(typeof(MarshalByRefObject))]

    sealed public class ImportTypeLib : PSCmdlet
    {
        [Parameter(Position = 0, ParameterSetName = "path", Mandatory = true)]
        public string[] Path;
        [Parameter(ParameterSetName = "literal", Mandatory = true, ValueFromPipeline = true)]
        public string[] LiteralPath;
        [Parameter(ParameterSetName = "guid", Mandatory = true)]
        public Guid Guid;
        [Parameter(ParameterSetName = "guid", Mandatory = true)]
        public Version Version;
        [Parameter(ParameterSetName = "guid", Mandatory = true)]
        public UInt32 LCID;

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
                                        GetPath(item);
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
                            GetPath(GetUnresolvedProviderPathFromPSPath(literalPath));
                        }
                        catch (ItemNotFoundException ex)
                        {
                            WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.ResourceUnavailable, literalPath));
                        }
                    }
                    break;

                case "guid":
                    try
                    {
                        OleAut32.LoadRegTypeLib(ref Guid, (UInt16)Version.Major, (UInt16)Version.Minor, LCID, out var lib);
                        WriteObject(lib);
                    }
                    catch (COMException ex)
                    {
                        WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidResult, Guid));
                    }
                    break;

                default:
                    NotImplementedException ni = new NotImplementedException();
                    WriteError(new ErrorRecord(ni, ni.GetType().Name, ErrorCategory.NotImplemented, ParameterSetName));
                    break;
            }
        }

        private void GetPath(string path)
        {
            try
            {
                OleAut32.LoadTypeLibEx(path, REGKIND.REGKIND_NONE, out var lib);
                WriteObject(lib);
            }
            catch (COMException ex)
            {
                WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidResult, path));
            }
        }
    }
}
