Option Explicit

Dim locator, srv, reg, rootset, rootkey, subset, subkey, clsid
Const HKEY_CLASSES_ROOT = &H80000000

Set locator = CreateObject("WbemScripting.SWbemLocator")
Set srv = locator.ConnectServer(vbNullString, "root\default")
Set reg = srv.Get("StdRegProv")

reg.EnumKey HKEY_CLASSES_ROOT, "", rootset

For Each rootkey In rootset
    clsid = ""
    reg.EnumKey HKEY_CLASSES_ROOT, CStr(rootkey), subset
    If IsArray(subset) Then
        For Each subkey In subset
            If CStr(subkey) = "CLSID" Then
                reg.GetStringValue HKEY_CLASSES_ROOT, CStr(rootkey) & "\" & CStr(subkey), "", clsid
                Exit For
            End If
        Next
    End If
    If clsid <> "" Then
        WScript.StdOut.WriteLine "ProgID/" & CStr(rootkey) & "/" & clsid
    End If
Next


