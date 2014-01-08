VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "clsSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'CTX: Implements 1 11 implement
Implements ISample

'CTX: As 2 14 class
'CTX: DefinePrivateVar 1 8 definition
Private ws As Worksheet

Private Sub ISample_testMethodS()
    'CTX: DefineDim 1 8 definition
    Dim c As Range

    'CTX: Set 1 8 variable
    Set ws = ActiveSheet

    'CTX: PrivateInstance 1 15 class-member Worksheet
    Set c = ws.Cells(1, 1)

    'CTX: EndPrivateSub 1 4 end Sub
End Sub

Private Function ISample_testMethodF() As MSXML2.DOMDocument
    'CTX: AsNew 2 19 class
    'CTX: BareMember 1 26 app-member MSXML2
    Dim req As New MSXML2.XMLHTTP
    Dim doc As Object

    'CTX: DimInstance 3 8 class-member MSXML2.XMLHTTP
    'CTX: ParamSpace1 2 13 param String
    'CTX: ParamSpace2 1 47 param Variant
    req.Open "GET", "http://www.google.co.jp/", False
    
    Set doc = CreateObject("MSXML2.DOMDocument")
    
    'CTX: CreateObject 2 13 class-member DOMDocument
    'CTX: ParamBrace 1 21 param String
    Call doc.loadXML(req.responseText)
    
    Set ISample_testMethodF = doc

    'CTX: EndPrivateFunc 1 4 end Function
End Function

Public Function smpMethod(ByRef r As Range, _
                          Optional ByVal optidx As Integer) As Worksheet

    With r.Interior

        MsgBox "hoge"

        'CTX: With 1 9 class-member Interior
        .ColorIndex = optidx
        
    End With

    'CTX: NestMember 1 59 class-member Worksheet
    r.SpecialCells(XlCellType.xlCellTypeVisible).Worksheet.Name

    Set smpMethod = r.Worksheet
    
End Function

