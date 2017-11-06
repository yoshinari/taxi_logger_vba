Attribute VB_Name = "taxi_logger"
Sub taxi_logger()
    '�ǂݍ���json�t�@�C����I������
    Dim JsonFileName As String
    Dim SheetName As String
    Dim SheetCount As Integer
    Dim NewWorkSheet As Worksheet
    Dim i, j As Integer
    JsonFileName = Application.GetOpenFilename("json�t�@�C��,*.json")
    If (JsonFileName = "False") Then
        MsgBox ("�L�����Z�����܂���")
        End
    End If
        
    SheetName = Right(JsonFileName, InStr(StrReverse(JsonFileName), Application.PathSeparator) - 1)
    SheetName = Left(SheetName, InStr(SheetName, ".") - 1)
    '�t�@�C�����̃V�[�g�����
    SheetCount = Worksheets.Count
    For i = 1 To SheetCount
        If (Worksheets(i).Name = SheetName) Then
            MsgBox ("���[�N�V�[�g:" & SheetName & "�͊��ɑ��݂��܂��B" & vbCrLf & "�����𒆒f���܂��B")
            End
        End If
    Next
    Set NewWorkSheet = Worksheets.Add(After:=Worksheets(1))
    NewWorkSheet.Name = SheetName
    
    'json�t�@�C����ǂݍ���
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
    
    '�f�[�^����������
    writePos = 5
    Cells(2, 7) = "�斱��"
    Cells(2, 8) = JsonObj(1)("Date")
    Cells(4, 3) = "��Ԏ���"
    Cells(4, 4) = "��Ԓn"
    Cells(4, 5) = "�o�R����"
    Cells(4, 6) = "�o�R�n"
    Cells(4, 7) = "�~�Ԏ���"
    Cells(4, 8) = "�~�Ԓn"
    For i = 1 To histories
        Cells(writePos, 2) = i
        Cells(writePos, 3) = JsonObj(i)("GetInTime")
        mapUrl = "http://maps.google.com/maps?q=" & JsonObj(i)("GetInLat") & "," & JsonObj(i)("GetInLng")
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(writePos, 3), Address:=mapUrl, TextToDisplay:="*", ScreenTip:="��Ԓn��Google Map�ŕ\�����܂�"
        Cells(writePos, 4) = JsonObj(i)("GetInAddress")
        Cells(writePos, 7) = JsonObj(i)("GetOutTime")
        mapUrl = "http://maps.google.com/maps?q=" & JsonObj(i)("GetOutLat") & "," & JsonObj(i)("GetOutLng")
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(writePos, 7), Address:=mapUrl, TextToDisplay:="*", ScreenTip:="�~�Ԓn��Google Map�ŕ\�����܂�"
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
                ActiveSheet.Hyperlinks.Add Anchor:=Cells(writePos + j - 1, 5), Address:=mapUrl, TextToDisplay:="*", ScreenTip:="�o�R�n��Google Map�ŕ\�����܂�"
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
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(writePos, 2), Address:=routingUrl, ScreenTip:="���̂��Google Map�ŕ\�����܂�"
        writePos = writePos + 1
    Next
    
    '�̍ق𐮂���
    Call beautify(writePos - 1)
    Cells(1, 1).Select
End Sub
