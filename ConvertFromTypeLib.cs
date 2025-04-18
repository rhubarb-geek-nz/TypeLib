// Copyright (c) 2025 Roger Brown.
// Licensed under the MIT License.

using System;
using System.ComponentModel;
using System.IO;
using System.Management.Automation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace RhubarbGeekNz.TypeLib
{
    [Cmdlet(VerbsData.ConvertFrom, "TypeLib")]
    [OutputType(typeof(Assembly))]
    sealed public class ConvertFromTypeLib : PSCmdlet
    {
        [Parameter(Position = 0, ParameterSetName = "path", Mandatory = true)]
        public string Path;
        [Parameter(ParameterSetName = "literal", Mandatory = true)]
        public string LiteralPath;
        [Parameter(ParameterSetName = "guid", Mandatory = true)]
        public Guid Guid;
        [Parameter(ParameterSetName = "guid", Mandatory = true)]
        public Version Version;
        [Parameter(ParameterSetName = "guid", Mandatory = true)]
        public UInt32 LCID;
        [Parameter(ParameterSetName = "typelib", Mandatory = true)]
        public MarshalByRefObject TypeLib;
        [Parameter(Mandatory = true)]
        public string Name;

        protected override void ProcessRecord()
        {
            switch (ParameterSetName)
            {
                case "path":
                    try
                    {
                        var paths = GetResolvedProviderPathFromPSPath(Path, out var providerPath);

                        if ("FileSystem".Equals(providerPath.Name))
                        {
                            if (paths.Count == 0)
                            {
                                Exception ex = new Win32Exception(OleAut32.FILE_NOT_FOUND);
                                WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.ResourceUnavailable, Path));
                            }
                            else
                            {
                                foreach (string item in paths)
                                {
                                    ImportTypeLibFromPath(item);
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
                        WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.ResourceUnavailable, Path));
                    }
                    break;

                case "guid":
                    try
                    {
                        OleAut32.LoadRegTypeLib(ref Guid, (UInt16)Version.Major, (UInt16)Version.Minor, LCID, out var typeLib);
                        WriteObject(Import(typeLib, Name));
                    }
                    catch (COMException ex)
                    {
                        WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidResult, Guid));
                    }
                    break;

                case "literal":
                    try
                    {
                        ImportTypeLibFromPath(GetUnresolvedProviderPathFromPSPath(LiteralPath));
                    }
                    catch (ItemNotFoundException ex)
                    {
                        WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.ResourceUnavailable, LiteralPath));
                    }
                    break;

                case "typelib":
                    try
                    {
                        WriteObject(Import((ITypeLib)TypeLib, Name));
                    }
                    catch (COMException ex)
                    {
                        WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidResult, this.TypeLib));
                    }
                    catch (TypeLoadException ex)
                    {
                        WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidType, this.TypeLib));
                    }
                    break;

                default:
                    NotImplementedException ni = new NotImplementedException();
                    WriteError(new ErrorRecord(ni, ni.GetType().Name, ErrorCategory.NotImplemented, ParameterSetName));
                    break;
            }
        }

        private void ImportTypeLibFromPath(string path)
        {
            try
            {
                OleAut32.LoadTypeLibEx(path, REGKIND.REGKIND_NONE, out var typeLib);
                WriteObject(Import(typeLib, Name));
            }
            catch (COMException ex)
            {
                WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidResult, path));
            }
            catch (TypeLoadException ex)
            {
                WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidType, path));
            }
        }

        public Assembly Import(ITypeLib typeLib, string asmFileName)
        {
            return new TypeLibConverter().ConvertTypeLibToAssembly(typeLib, asmFileName, 0, new Sink(this), null, null, false);
        }

        sealed internal class Sink : ITypeLibImporterNotifySink
        {
            private readonly PSCmdlet cmdlet;

            public Sink(PSCmdlet cmdlet)
            {
                this.cmdlet = cmdlet;
            }

            public void ReportEvent(ImporterEventKind eventKind, int eventCode, string eventMsg)
            {
                cmdlet.WriteVerbose(eventMsg);
            }

            public Assembly ResolveRef(object obj)
            {
                return null;
            }
        }
    }
}
