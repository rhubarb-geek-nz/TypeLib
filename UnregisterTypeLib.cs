// Copyright (c) 2025 Roger Brown.
// Licensed under the MIT License.

using System;
using System.Management.Automation;
using System.Runtime.InteropServices;
using SYSKIND = System.Runtime.InteropServices.ComTypes.SYSKIND;

namespace RhubarbGeekNz.TypeLib
{
    [Cmdlet(VerbsLifecycle.Unregister, "TypeLib")]
    sealed public class UnregisterTypeLib : PSCmdlet
    {
        [Parameter(Mandatory = true)]
        public Guid Guid;
        [Parameter(Mandatory = true)]
        public Version Version;
        [Parameter(Mandatory = true)]
        public UInt32 LCID;
        [Parameter(Mandatory = true)]
        public SYSKIND SysKind;
        [Parameter(Mandatory = false)]
        [ValidateSet("AllUsers", "CurrentUser")]
        public string Scope;

        protected override void ProcessRecord()
        {
            try
            {
                if (Scope == null || 0 == String.Compare(Scope, "CurrentUser", true))
                {
                    OleAut32.UnRegisterTypeLibForUser(ref Guid, (UInt16)Version.Major, (UInt16)Version.Minor, LCID, SysKind);
                }
                else
                {
                    OleAut32.UnRegisterTypeLib(ref Guid, (UInt16)Version.Major, (UInt16)Version.Minor, LCID, SysKind);
                }
            }
            catch (COMException ex)
            {
                WriteError(new ErrorRecord(ex, ex.GetType().Name, ErrorCategory.InvalidResult, this.Guid));
            }
        }
    }
}
