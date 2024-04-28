Sub 결석일수정(WshtRow, TargetValue)

Dim Wsht As Worksheet: Set Wsht = Worksheets("출결 누가기록")
Dim Osht As Worksheet: Set Osht = Worksheets("설정")
Dim Dsht As Worksheet: Set Dsht = Worksheets("DataBase")

Application.ScreenUpdating = False
Application.EnableEvents = False

dim OriginalData()



If Right(Wsht.Range("Y" & WshtRow).Value, 2) <> "결석" Then
    MsgBox Prompt:="결석일 수정은 결석에 대해서만 가능합니다.", Title:="확인하세요!"
    Wsht.Range("AA" & WshtRow) = "-"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
ElseIf IsNumeric(TargetValue) Then

    Dim 번호 As Integer
    Dim 날짜 As Date
    Dim 이름 As String
    
    If Wsht.Range("N" & WshtRow).Value <> "" Then
        번호 = Wsht.Range("N" & WshtRow).Value
    Else
        번호 = Wsht.Range("N" & WshtRow).End(xlUp).Value
    End If
    If Wsht.Range("O" & WshtRow).Value <> "" Then
        이름 = Wsht.Range("O" & WshtRow).Value
    Else
        이름 = Wsht.Range("O" & WshtRow).End(xlUp).Value
    End If
    If Wsht.Range("M" & WshtRow).Value <> "" Then
        날짜 = Wsht.Range("M" & WshtRow).Value
    Else
        날짜 = Wsht.Range("M" & WshtRow).End(xlUp).Value
    End If

    If MsgBox(번호 & "번 " & 이름 & " " & 날짜 & "(" & TargetValue & "일간)의 출결 정보가 수정됩니다.", vbOKCancel, "확인하세요!") = vbOK Then
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    End If

    n = 0
    For j = 1 To TargetValue
        If Weekday(날짜 + n) <= 1 Or Weekday(날짜 + n) >= 7 Or Not Osht.Range("A17:L24").Find(날짜 + n, LookIn:=xlFormulas, Lookat:=xlWhole) Is Nothing Then
            j = j - 1
        End If
            n = n + 1
    Next
    끝날짜 = 날짜 + (n - 1)
    
    Dlstrw = Dsht.Range("P2").End(xlDown).Row
    For i = 4 To Dlstrw
        If Dsht.Range("B" & i).Value = 번호 And Dsht.Range("A" & i).Value >= 날짜 Then
            If Dsht.Range("A" & i).Value = 날짜 And Dsht.Range("O" & i).Value = "-" Then
                MsgBox Prompt:="결석일 수정은 결석 시작일에 대해서만 가능합니다.", Title:="확인하세요!"
                TargetValue = "-"
                Application.ScreenUpdating = True
                Application.EnableEvents = True
                Exit Sub
            End If
            
            If Dsht.Range("T" & i).Value = "-" Then
                Dsht.Range("T" & i).Value = Date
            ElseIf Dsht.Range("A" & i).Value > 끝날짜 And Dsht.Range("T" & i).Value = "-" Then
                If Dsht.Range("M" & i).Value <> Wsht.Range("Z" & WshtRow).Value Then
                    Exit For
                ElseIf Dsht.Range("O" & i).Value <> "-" Then
                    Exit For
                End If
            End If
            
        End If
    Next
    
    Dim 신규데이터()
    ReDim 신규데이터(19, TargetValue - 1)
    n = 0
    For i = 0 To TargetValue - 1

        If Weekday(날짜 + n) > 1 And Weekday(날짜 + n) < 7 And Osht.Range("A17:L24").Find(날짜 + n, LookIn:=xlFormulas, Lookat:=xlWhole) Is Nothing Then
            신규데이터(0, i) = 날짜 + n
            신규데이터(1, i) = 번호
            신규데이터(2, i) = 이름
            For j = 0 To 8
                신규데이터(3 + j, i) = "/"
            Next
            신규데이터(12, i) = Wsht.Range("Y" & WshtRow).Value
            신규데이터(13, i) = Wsht.Range("Z" & WshtRow).Value
            If i = 0 Then
                신규데이터(14, i) = TargetValue
                신규데이터(15, i) = Wsht.Range("AB" & WshtRow).Value
                신규데이터(16, i) = Wsht.Range("AC" & WshtRow).Value
                신규데이터(17, i) = Wsht.Range("AD" & WshtRow).Value
            Else
                신규데이터(14, i) = "-"
                신규데이터(15, i) = "-"
                신규데이터(16, i) = "-"
                신규데이터(17, i) = "-"
            End If
            신규데이터(18, i) = Date
            신규데이터(19, i) = "-"
        Else
            i = i - 1
        End If
        n = n + 1
    Next
    Dsht.Range("A" & Dlstrw + 1).Resize(UBound(신규데이터, 2) + 1, 20).Value = Application.Transpose(신규데이터)
    Call 누가기록조회
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
Else
    MsgBox Prompt:="결석일(숫자)을 입력해 주세요!", Title:="확인하세요!"
    Call 누가기록조회
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
End If

End Sub
