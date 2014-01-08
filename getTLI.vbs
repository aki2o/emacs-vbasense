Option Explicit

Dim filepath
For Each filepath In WScript.Arguments
    printLibInfo( filepath )
Next
WScript.Quit


Function escapeLinefeed( text )
    escapeLinefeed = Replace(Replace(text, vbCr, ""), vbLf, "\n")
End Function

Sub printLibInfo( filepath )
    Dim tli, c, i, m, tlitype
    Dim regexInterface

    Set tli = CreateObject("TLI.TLIApplication").TypeLibInfoFromFile(Replace(filepath, "/", "\"))
    
    Set regexInterface = CreateObject("VBScript.RegExp")
    regexInterface.Pattern = "^I[A-Z]"
    regexInterface.Global = False

    WScript.StdOut.WriteLine "App/" & tli.Name & "/" & tli.GUID
    For Each c In tli.CoClasses
        If instr(c.Name, "_") <> 1 Then
            For Each i In c.Interfaces
                For Each m In i.Members
                    If instr(m.Name, "_") <> 1 Then
                        WScript.StdOut.WriteLine "CoClass/" & c.Name & "/" & c.GUID & "/" & getMemberType(m, c) & "/" & m.Name & "/" & getMemberReturn(m) & "/" & getMemberParam(m) & "###" & m.HelpString & "###" & m.HelpFile
                    End if
                Next
            Next
        End If
    Next
    For Each i In tli.Declarations
        If instr(i.Name, "_") <> 1 Then
            For Each m In i.Members
                If instr(m.Name, "_") <> 1 Then
                    WScript.StdOut.WriteLine "Module/" & i.Name & "/" & i.GUID & "/" & getMemberType(m, i) & "/" & m.Name & "/" & getMemberReturn(m) & "/" & getMemberParam(m) & "###" & m.HelpString & "###" & m.HelpFile
                End if
            Next
        End If
    Next
    For Each i In tli.Interfaces
        If instr(i.Name, "_") <> 1 Then
            If i.TypeKind = 3 Then
                tlitype = "Interface"
            Else
                tlitype = "Class"
            End If
            For Each m In i.Members
                If instr(m.Name, "_") <> 1 And instr(m.Name, "Dummy") <> 1 Then
                    WScript.StdOut.WriteLine tlitype & "/" & i.Name & "/" & i.GUID & "/" & getMemberType(m, i) & "/" & m.Name & "/" & getMemberReturn(m) & "/" & getMemberParam(m) & "###" & m.HelpString & "###" & m.HelpFile
                End If
            Next
        End If
    Next
    For Each c In tli.Constants
        For Each m In c.Members
            WScript.StdOut.WriteLine "Const/" & c.Name & "/" & c.GUID & "/" & m.Name & "###" & m.HelpString & "###" & m.HelpFile
        Next
    Next
End Sub

Function getMemberType( member, interfacenm )
    Dim ret
    On Error Resume Next
    Select Case member.InvokeKind
        Case 0
            ret = "Unknown"
        Case 1
            ret = "Function"
        Case 2
            ret = "Property Get"
        Case 4
            ret = "Property Put"
        Case 8
            ret = "Property Let"
        Case 16
            ret = "Event"
        Case 32
            ret = "Const"
        Case Else
    End Select
    If getMemberReturn( member ) = "" And ret = "Function" Then ret = "Sub"
    If instr(interfacenm, "Events") = len(interfacenm) - 5 Then ret = "Event"
    getMemberType = ret
End Function

Function getMemberReturn( member )
    Dim ret
    On Error Resume Next
    If member.ReturnType Is Nothing Then
        ret = ""
    ElseIf Not IsEmpty(member.ReturnType.TypedVariant) Then
        ret = TypeName(member.ReturnType.TypedVariant)
    ElseIf Not member.ReturnType.TypeInfo Is Nothing Then
        ret = member.ReturnType.TypeInfo.Name
    Else
        ret = ""
    End If
    getMemberReturn = ret
End Function

Function getMemberParam( member )
    Dim ret, i, p
    On Error Resume Next
    For i = 1 To member.Parameters.Count
        Set p = member.Parameters.Item(i)
        If i > 1 Then ret = ret & "/"
        ret = ret & p
        ret = ret & ","
        If Not p.VarTypeInfo Is Nothing And Not IsEmpty(p.VarTypeInfo.TypedVariant) Then
            ret = ret & TypeName(p.VarTypeInfo.TypedVariant)
        ElseIf Not p.VarTypeInfo.TypeInfo is Nothing Then
            ret = ret & p.VarTypeInfo.TypeInfo.Name
        End If
        ret = ret & ","
        If p.Optional Then ret = ret & "Optional"
        ret = ret & ","
        If p.Default Then ret = ret & p.DefaultValue.ToString
    Next
    getMemberParam = ret
End Function

