Attribute VB_Name = "taxi_logger"
Sub taxi_logger()
    '読み込むjsonファイルを選択する
    Dim JsonFileName As String
    Dim SheetName As String
    Dim SheetCount As Integer
    Dim NewWorkSheet As Worksheet
    Dim i, j As Integer
    JsonFileName = Application.GetOpenFilename("jsonファイル,*.json")
    If (JsonFileName = "False") Then
        MsgBox ("キャンセルしました")
        End
    End If
        
    SheetName = Right(JsonFileName, InStr(StrReverse(JsonFileName), Application.PathSeparator) - 1)
    SheetName = Left(SheetName, InStr(SheetName, ".") - 1)
    'ファイル名のシートを作る
    SheetCount = Worksheets.Count
    For i = 1 To SheetCount
        If (Worksheets(i).Name = SheetName) Then
            MsgBox ("ワークシート:" & SheetName & "は既に存在します。" & vbCrLf & "処理を中断します。")
            End
        End If
    Next
    Set NewWorkSheet = Worksheets.Add(After:=Worksheets(1))
    NewWorkSheet.Name = SheetName
    
    'jsonファイルを読み込む
    Dim jl As JsonLoader
    Set jl = New JsonLoader
    Dim JsonString, viaString As String
    Dim JsonObj, viaObj As Object
    Dim histories, vias As Integer
    Dim writePos As Integer
    Dim routingUrl, mapUrl As String
    JsonString = jl.LoadJsonFile(JsonFileName)
    Set JsonObj = JsonConverter.ParseJson(JsonString)
    
    histories = JsonObj.Count
    
    'データを書き込む
    writePos = 5
    Cells(2, 7) = "乗務日"
    Cells(2, 8) = JsonObj(1)("Date")
    Cells(4, 3) = "乗車時刻"
    Cells(4, 4) = "乗車地"
    Cells(4, 5) = "経由時刻"
    Cells(4, 6) = "経由地"
    Cells(4, 7) = "降車時刻"
    Cells(4, 8) = "降車地"
    For i = 1 To histories
        Cells(writePos, 2) = i
        Cells(writePos, 3) = JsonObj(i)("GetInTime")
        mapUrl = "http://maps.google.com/maps?q=" & JsonObj(i)("GetInLat") & "," & JsonObj(i)("GetInLng")
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(writePos, 3), Address:=mapUrl, TextToDisplay:="*", ScreenTip:="乗車地をGoogle Mapで表示します"
        Cells(writePos, 4) = JsonObj(i)("GetInAddress")
        Cells(writePos, 7) = JsonObj(i)("GetOutTime")
        mapUrl = "http://maps.google.com/maps?q=" & JsonObj(i)("GetOutLat") & "," & JsonObj(i)("GetOutLng")
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(writePos, 7), Address:=mapUrl, TextToDisplay:="*", ScreenTip:="降車地をGoogle Mapで表示します"
        Cells(writePos, 8) = JsonObj(i)("GetOutAddress")
        routingUrl = "http://www.google.co.jp/maps/dir/" & JsonObj(i)("GetInLat") & "," & JsonObj(i)("GetInLng")
        If (JsonObj(i)("GetInMemo") <> "") Then
            Cells(writePos, 4).AddComment
            Cells(writePos, 4).Comment.Visible = False
            Cells(writePos, 4).Comment.Text Text:=JsonObj(i)("GetInMemo")
        End If
        If (JsonObj(i)("GetOutMemo") <> "") Then
            Cells(writePos, 8).AddComment
            Cells(writePos, 8).Comment.Visible = False
            Cells(writePos, 8).Comment.Text Text:=JsonObj(i)("GetOutMemo")
        End If
        If (JsonObj(i)("ViaMemo") <> "") Then
            Cells(writePos, 1).AddComment
            Cells(writePos, 1).Comment.Visible = False
            Cells(writePos, 1).Comment.Text Text:=JsonObj(i)("ViaMemo")
        End If
        viaString = JsonObj(i)("ViaData")
        If (viaString <> "[]") Then
            Set viaObj = JsonConverter.ParseJson(viaString)
            vias = viaObj.Count
            For j = 1 To vias
                Cells(writePos + j - 1, 5) = viaObj(j)("time")
                mapUrl = "http://maps.google.com/maps?q=" & viaObj(j)("lat") & "," & viaObj(j)("lng")
                ActiveSheet.Hyperlinks.Add Anchor:=Cells(writePos + j - 1, 5), Address:=mapUrl, TextToDisplay:="*", ScreenTip:="経由地をGoogle Mapで表示します"
                Cells(writePos + j - 1, 6) = viaObj(j)("address")
                routingUrl = routingUrl & "/" & viaObj(j)("lat") & "," & viaObj(j)("lng")
                If (viaObj(j)("memo") <> "") Then
                    Cells(writePos + j - 1, 6).AddComment
                    Cells(writePos + j - 1, 6).Comment.Visible = False
                    Cells(writePos + j - 1, 6).Comment.Text Text:=viaObj(j)("memo")
                End If
            Next
            writePos = writePos + vias - 1
            If (vias > 1) Then
                Range(Cells(writePos - vias + 1, 2), Cells(writePos, 2)).MergeCells = True
                Range(Cells(writePos - vias + 1, 3), Cells(writePos, 3)).MergeCells = True
                Range(Cells(writePos - vias + 1, 4), Cells(writePos, 4)).MergeCells = True
                Range(Cells(writePos - vias + 1, 7), Cells(writePos, 7)).MergeCells = True
                Range(Cells(writePos - vias + 1, 8), Cells(writePos, 8)).MergeCells = True
            End If
        End If
        routingUrl = routingUrl & "/" & JsonObj(i)("GetOutLat") & "," & JsonObj(i)("GetOutLng") & "/@/data=!4m2!4m1!3e0"
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(writePos, 2), Address:=routingUrl, ScreenTip:="道のりをGoogle Mapで表示します"
        writePos = writePos + 1
    Next
    
    '体裁を整える
    Call beautify(writePos - 1)
    Cells(1, 1).Select
End Sub
