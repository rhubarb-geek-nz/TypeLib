# rhubarb-geek-nz.TypeLib
TypeLib tools for PowerShell

## Summary

There may be a time when you really need to deal with type libraries (`*.tlb`) with PowerShell.

## Notes

The Scope for registering and unregistering is either `AllUsers` or the default `CurrentUser`.

## Syntax

Convert from a type library to an assembly. Uses [ConvertTypeLibToAssembly](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.typelibconverter.converttypelibtoassembly?view=netframework-4.8.1). This requires PowerShell 5.1.

```
ConvertFrom-TypeLib [-Path] <string> -Name <string>

ConvertFrom-TypeLib -LiteralPath <string> -Name <string>

ConvertFrom-TypeLib -Guid <guid> -Version <version> -LCID <uint32> -Name <string>

ConvertFrom-TypeLib -TypeLib <MarshalByRefObject> -Name <string>
```

Get details from a type library file. Uses [LoadTypeLibEx](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-loadtypelibex) with `REGKIND_NONE`.

```
Get-TypeLib [-Path] <string[]>

Get-TypeLib -LiteralPath <string[]>

Get-TypeLib -Guid <guid> -Version <version> -LCID <uint>

Get-TypeLib -TypeLib <MarshalByRefObject>
```

Import returns an [ITypeLib](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.comtypes.itypelib).

```
Import-TypeLib [-Path] <string[]>

Import-TypeLib -LiteralPath <string[]>

Import-TypeLib -Guid <guid> -Version <version> -LCID <uint>
```

Register a library file. Uses [RegisterTypeLib](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-registertypelib) or [RegisterTypeLibForUser](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-registertypelibforuser)

```
Register-TypeLib [-Path] <string[]> [-Scope <string>] [-HelpDirectory <string>]

Register-TypeLib -LiteralPath <string[]> [-Scope <string>] [-HelpDirectory <string>]
```

Unregister an existing library. Uses [UnRegisterTypeLib](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-unregistertypelib) or [UnRegisterTypeLibForUser](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-unregistertypelibforuser).

```
Unregister-TypeLib -Guid <guid> -Version <version> -LCID <uint32> -SysKind <SYSKIND> [-Scope <string>]
```
