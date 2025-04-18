// Copyright (c) 2025 Roger Brown.
// Licensed under the MIT License.

using System;
using System.ComponentModel;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace RhubarbGeekNz.TypeLib
{
    [Cmdlet(VerbsCommon.Get, "TypeLib")]
    sealed public class GetTypeLib : PSCmdlet
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
        [Parameter(ParameterSetName = "typelib", Mandatory = true, ValueFromPipeline = true)]
        public MarshalByRefObject TypeLib;

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

                case "typelib":
                    try
                    {
                        WriteTypeLib((ITypeLib)TypeLib, null);
                    }
                    catch (COMException ex)
                    {
                        WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidResult, this.TypeLib));
                    }
                    break;

                case "guid":
                    try
                    {
                        OleAut32.QueryPathOfRegTypeLib(ref Guid, (UInt16)Version.Major, (UInt16)Version.Minor, LCID, out string path);
                        OleAut32.LoadTypeLibEx(path, REGKIND.REGKIND_NONE, out var typeLib);
                        WriteTypeLib(typeLib, path);
                    }
                    catch (COMException ex)
                    {
                        WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidResult, this.TypeLib));
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
                WriteTypeLib(lib, path);
            }
            catch (COMException ex)
            {
                WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidResult, path));
            }
        }

        private void WriteTypeLib(ITypeLib lib, string file)
        {
            lib.GetLibAttr(out var attr);

            try
            {
                lib.GetDocumentation(-1, out var name, out var doc, out var context, out var help);

                TLIBATTR tlibAttr = (TLIBATTR)Marshal.PtrToStructure(attr, typeof(TLIBATTR));

                PSObject obj = new PSObject();

                if (file != null)
                {
                    obj.Members.Add(new PSNoteProperty("File", file));
                }

                obj.Members.Add(new PSNoteProperty("Guid", tlibAttr.guid));
                obj.Members.Add(new PSNoteProperty("Version", new Version(tlibAttr.wMajorVerNum, tlibAttr.wMinorVerNum)));
                obj.Members.Add(new PSNoteProperty("LCID", tlibAttr.lcid));
                obj.Members.Add(new PSNoteProperty("SysKind", tlibAttr.syskind));
                obj.Members.Add(new PSNoteProperty("Flags", tlibAttr.wLibFlags));
                obj.Members.Add(new PSNoteProperty("Name", name));
                obj.Members.Add(new PSNoteProperty("Documentation", doc));
                obj.Members.Add(new PSNoteProperty("HelpContext", context));
                obj.Members.Add(new PSNoteProperty("Help", help));

                WriteObject(obj);
            }
            finally
            {
                lib.ReleaseTLibAttr(attr);
            }
        }
    }
}
