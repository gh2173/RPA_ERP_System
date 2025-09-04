Sub GroupBy_I_Z_And_Process()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, groupNum As Long
    Dim key As String
    Dim groupMap As Object, groupSums As Object, groupDesc As Object
    Dim maturityDate As Date, adjustedSum As Double
    Dim maturityCol As Long, descCol As Long, taxDateCol As Long
    Dim gDate As Variant
    Dim currentGroup As Long, nextGroup As Long
    Dim lastCol As Long
    Dim pText As String, jText As String

    Set ws = ThisWorkbook.Sheets("Sheet1") ' 시트 이름 필요 시 변경
    Set groupMap = CreateObject("Scripting.Dictionary")
    Set groupSums = CreateObject("Scripting.Dictionary")
    Set groupDesc = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    ' I열과 Z열 기준 정렬
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("I2:I" & lastRow), Order:=xlAscending
    ws.Sort.SortFields.Add Key:=ws.Range("Z2:Z" & lastRow), Order:=xlAscending
    With ws.Sort
        .SetRange ws.Range("A1:AG" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' A열 앞에 '그룹번호' 열 추가
    ws.Columns("A").Insert Shift:=xlToRight
    ws.Cells(1, 1).Value = "그룹번호"

    groupNum = 1

    ' 그룹번호 부여 및 AG 합계, 송장 설명 생성
    For i = 2 To lastRow
        key = ws.Cells(i, "I").Value & "|" & ws.Cells(i, "Z").Value

        If Not groupMap.exists(key) Then
            groupMap(key) = groupNum
            groupSums(groupNum) = 0

            ' 송장 설명: Month(G) & "월 " & P & "_" & J
            gDate = ws.Cells(i, "G").Value
            pText = ws.Cells(i, "P").Value
            jText = ws.Cells(i, "J").Value

            If IsDate(gDate) Then
                groupDesc(groupNum) = Month(gDate) & "월 " & pText & "_" & jText
            Else
                groupDesc(groupNum) = "날짜오류 " & pText & "_" & jText
            End If

            groupNum = groupNum + 1
        End If

        ws.Cells(i, 1).Value = groupMap(key)
        groupSums(groupMap(key)) = groupSums(groupMap(key)) + Val(ws.Cells(i, "AG").Value)
    Next i

    ' 열 추가: 만기일, 송장 설명, 세금계산서일자
    maturityCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, maturityCol).Value = "만기일"

    descCol = maturityCol + 1
    ws.Cells(1, descCol).Value = "송장 설명"

    taxDateCol = descCol + 1
    ws.Cells(1, taxDateCol).Value = "세금계산서일자"

    ' 만기일, 송장 설명, 세금계산서일자 입력
    For i = 2 To lastRow
        Dim gNum As Long
        gNum = ws.Cells(i, 1).Value
        gDate = ws.Cells(i, "G").Value

        If IsDate(gDate) Then
            adjustedSum = groupSums(gNum) * 1.1
            If adjustedSum < 10000000 Then
                maturityDate = WorksheetFunction.EoMonth(gDate, 1) ' 익월 말일
            Else
                maturityDate = WorksheetFunction.EoMonth(gDate, 2) ' 익익월 말일
            End If
            ws.Cells(i, maturityCol).Value = maturityDate

            ' 세금계산서일자: G열의 월 말일
            ws.Cells(i, taxDateCol).Value = WorksheetFunction.EoMonth(gDate, 0)
        Else
            ws.Cells(i, maturityCol).Value = "날짜 오류"
            ws.Cells(i, taxDateCol).Value = "날짜 오류"
        End If

        ' 송장 설명 입력
        ws.Cells(i, descCol).Value = groupDesc(gNum)
    Next i

    ' 날짜 서식 적용 (세금계산서일자 열)
    ws.Range(ws.Cells(2, taxDateCol), ws.Cells(lastRow, taxDateCol)).NumberFormat = "yyyy-mm-dd"

    ' 그룹별 구분선 추가
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 2 To lastRow - 1
        currentGroup = ws.Cells(i, 1).Value
        nextGroup = ws.Cells(i + 1, 1).Value

        If currentGroup <> nextGroup Then
            With ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
        End If
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "그룹화, 만기일, 송장 설명, 세금계산서일자(날짜 형식), 구분선 완료!", vbInformation
End Sub