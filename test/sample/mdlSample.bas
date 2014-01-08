Attribute VB_Name = "mdlSample"

Option Explicit

'CTX: DefinePublicVar 1 7 definition
Public gSmp As ISample

'CTX: DefineEnum 1 12 definition
Public Enum gTag
    SMALL = 1
    MIDDLE = 2
    BIG = 3
    'CTX: EndPublicEnum 1 4 end Enum
End Enum

'CTX: DefineType 1 12 definition
Public Type gNW
    Port As Integer
    CableType As gTag
    'CTX: EndPublicType 1 4 end Type
End Type

Public Sub gMethodS(ByRef ws As Worksheet)
    Static rr As Range
    Dim cSmp As clsSample
    
    'CTX: New 1 19 class
    Set gSmp = New clsSample

    'CTX: PublicInstance 1 9 interface-member ISample
    gSmp.testMethodS

    'CTX: ArgInstance 1 16 class-member Worksheet
    Set rr = ws.Cells(1, 1)

    Set cSmp = gSmp

    'CTX: MultilineDefineParam 2 32 param Integer
    'CTX: MultilineDefineRet 1 36 class-member Worksheet
    Set rr = cSmp.smpMethod(rr, 10).Range
    
    'CTX: EndPublicSub 1 4 end Sub
End Sub

Public Function gMethodF(tType As gTag) As String
    Dim nodes As Object

    'CTX: UserMethodRet 1 33 class-member MSXML2.DOMDocument
    Set nodes = gSmp.testMethodF.childNodes
    
    'CTX: EndPublicFunc 1 4 end Function
End Function

