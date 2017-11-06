VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function LoadJsonFile(ByVal fileName As String) As String
    Dim buf As String
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile fileName
        buf = .ReadText
        .Close
    End With
    
    LoadJsonFile = buf
    
End Function
    

