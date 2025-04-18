// Copyright (c) 2025 Roger Brown.
// Licensed under the MIT License.

using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using SYSKIND = System.Runtime.InteropServices.ComTypes.SYSKIND;

namespace RhubarbGeekNz.TypeLib
{
    internal struct TLIBATTR
    {
        internal Guid guid;
        internal UInt32 lcid;
        internal System.Runtime.InteropServices.ComTypes.SYSKIND syskind;
        internal UInt16 wMajorVerNum, wMinorVerNum, wLibFlags;
    }

    internal enum REGKIND
    {
        REGKIND_NONE = 2
    }

    static internal class OleAut32
    {
        [DllImport("oleaut32", PreserveSig = false, CharSet = CharSet.Unicode)]
        internal static extern void LoadRegTypeLib(ref Guid libid, UInt16 majorVer, UInt16 minorVer, UInt32 lcid, out ITypeLib typeLib);

        [DllImport("oleaut32", PreserveSig = false, CharSet = CharSet.Unicode)]
        internal static extern void LoadTypeLibEx(string szFullPath, REGKIND regKind, out ITypeLib typeLib);

        [DllImport("oleaut32", PreserveSig = false, CharSet = CharSet.Unicode)]
        internal static extern void RegisterTypeLib(ITypeLib ptlib, string szFullPath, string szHelpDir);

        [DllImport("oleaut32", PreserveSig = false, CharSet = CharSet.Unicode)]
        internal static extern void RegisterTypeLibForUser(ITypeLib ptlib, string szFullPath, string szHelpDir);

        [DllImport("oleaut32", PreserveSig = false, CharSet = CharSet.Unicode)]
        internal static extern void UnRegisterTypeLibForUser(ref Guid libid, UInt16 majorVer, UInt16 minorVer, UInt32 lcid, SYSKIND syskind);

        [DllImport("oleaut32", PreserveSig = false, CharSet = CharSet.Unicode)]
        internal static extern void UnRegisterTypeLib(ref Guid libid, UInt16 majorVer, UInt16 minorVer, UInt32 lcid, SYSKIND syskind);

        [DllImport("oleaut32", PreserveSig = false, CharSet = CharSet.Unicode)]
        internal static extern void QueryPathOfRegTypeLib(ref Guid libid, UInt16 majorVer, UInt16 minorVer, UInt32 lcid, [MarshalAs(UnmanagedType.BStr)] out string path);

        internal const int FILE_NOT_FOUND = -2147024894;
    }
}
