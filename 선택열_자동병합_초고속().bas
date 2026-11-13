Option Explicit

Sub 선택열_자동병합_초고속()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxLastRow As Long

    Dim targetCols As Variant
    Dim col As Variant

    Dim i As Long
    Dim groupCount As Long
    Dim processCount As Long

    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    Dim inputCols As String
    Dim colNames As Variant
    Dim colName As Variant
    Dim colNumber As Long
    Dim validCount As Long

    On Error GoTo ErrorHandler

    '=========================================================
    ' 현재 시트
    '=========================================================

    Set ws = ActiveSheet

    '=========================================================
    ' 병합할 열 입력
    ' 예: A, B, D, F
    '=========================================================

    inputCols = InputBox( _
        "병합할 열을 입력하세요." & vbCrLf & vbCrLf & _
        "예: A, B, D, F" & vbCrLf & _
        "여러 열은 쉼표(,)로 구분하세요.", _
        "병합할 열 입력")

    If Trim(inputCols) = "" Then
        Exit Sub
    End If

    colNames = Split(inputCols, ",")

    '=========================================================
    ' 입력된 열이 올바른지 확인
    '=========================================================

    validCount = 0

    For Each colName In colNames

        colName = Trim(UCase(CStr(colName)))

        If colName <> "" Then

            On Error Resume Next

            Err.Clear
            colNumber = ws.Range(colName & "1").Column

            If Err.Number <> 0 Then
                colNumber = 0
                Err.Clear
            End If

            On Error GoTo ErrorHandler

            If colNumber > 0 Then
                validCount = validCount + 1
            End If

        End If

    Next colName

    If validCount = 0 Then

        MsgBox "올바른 열을 입력하세요." & vbCrLf & vbCrLf & _
               "예: A, B, D, F", _
               vbExclamation

        Exit Sub

    End If

    '=========================================================
    ' 입력된 열을 배열로 구성
    '=========================================================

    ReDim targetCols(0 To validCount - 1)

    i = 0

    For Each colName In colNames

        colName = Trim(UCase(CStr(colName)))

        If colName <> "" Then

            On Error Resume Next

            Err.Clear
            colNumber = ws.Range(colName & "1").Column

            If Err.Number <> 0 Then
                colNumber = 0
                Err.Clear
            End If

            On Error GoTo ErrorHandler

            If colNumber > 0 Then

                targetCols(i) = colNumber
                i = i + 1

            End If

        End If

    Next colName

    '=========================================================
    ' Excel 현재 설정 저장
    '=========================================================

    oldCalculation = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts

    '=========================================================
    ' Excel 최적화
    '=========================================================

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Application.StatusBar = "데이터를 읽는 중..."

    '=========================================================
    ' 마지막 행 확인
    ' 선택된 열 중 가장 긴 열을 기준으로 계산
    '=========================================================

    maxLastRow = 1

    For Each col In targetCols

        lastRow = ws.Cells(ws.Rows.Count, CLng(col)).End(xlUp).Row

        If lastRow > maxLastRow Then
            maxLastRow = lastRow
        End If

    Next col

    '=========================================================
    ' 선택된 열 처리
    '=========================================================

    groupCount = 0
    processCount = 0

    For Each col In targetCols

        Application.StatusBar = _
            ColumnLetter(ws, CLng(col)) & "열 처리 중..."

        Call ProcessColumnFast( _
            ws, _
            CLng(col), _
            maxLastRow, _
            groupCount, _
            processCount)

    Next col

    '=========================================================
    ' Excel 설정 복구
    '=========================================================

    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating

    Application.StatusBar = False

    MsgBox "완료되었습니다." & vbCrLf & vbCrLf & _
           "선택한 열 : " & GetSelectedColumnsText(ws, targetCols) & vbCrLf & _
           "처리 그룹 수 : " & groupCount, _
           vbInformation

    Exit Sub

ErrorHandler:

    On Error Resume Next

    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.StatusBar = False

    MsgBox "오류가 발생했습니다." & vbCrLf & vbCrLf & _
           "오류 번호 : " & Err.Number & vbCrLf & _
           "오류 내용 : " & Err.Description, _
           vbCritical

End Sub


'=============================================================
' 실제 열 처리
'=============================================================

Private Sub ProcessColumnFast( _
    ByVal ws As Worksheet, _
    ByVal colNumber As Long, _
    ByVal lastRow As Long, _
    ByRef groupCount As Long, _
    ByRef processCount As Long)

    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long

    Dim mergedText As String

    Dim r As Range
    Dim dataArr As Variant

    '=========================================================
    ' 해당 열 데이터를 한 번에 배열로 읽음
    '=========================================================

    dataArr = ws.Range( _
                ws.Cells(1, colNumber), _
                ws.Cells(lastRow, colNumber)).Value2

    '=========================================================
    ' 기존 병합 해제
    '=========================================================

    ws.Range( _
        ws.Cells(1, colNumber), _
        ws.Cells(lastRow, colNumber)).UnMerge

    i = 1

    Do While i <= lastRow

        '=====================================================
        ' 대표값 발견
        '=====================================================

        If Trim(CStr(dataArr(i, 1))) <> "" Then

            startRow = i
            endRow = i

            '=================================================
            ' 다음 대표값 직전까지 검색
            '=================================================

            Do While endRow < lastRow

                If Trim(CStr(dataArr(endRow + 1, 1))) = "" Then

                    endRow = endRow + 1

                Else

                    Exit Do

                End If

            Loop

            '=================================================
            ' 현재 열 내용 합치기
            '=================================================

            mergedText = _
                CombineArrayText( _
                    dataArr, _
                    startRow, _
                    endRow)

            '=================================================
            ' 범위
            '=================================================

            Set r = ws.Range( _
                    ws.Cells(startRow, colNumber), _
                    ws.Cells(endRow, colNumber))

            '=================================================
            ' 기존 내용 삭제
            '=================================================

            r.ClearContents

            '=================================================
            ' 대표 셀에 내용 입력
            '=================================================

            ws.Cells(startRow, colNumber).Value = _
                mergedText

            '=================================================
            ' 병합
            '=================================================

            r.Merge

            '=================================================
            ' 정렬
            '=================================================

            With r

                .WrapText = True
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter

            End With

            groupCount = groupCount + 1
            processCount = processCount + 1

            '=================================================
            ' Excel이 완전히 멈추지 않도록 가끔만 처리
            '=================================================

            If processCount Mod 50 = 0 Then

                Application.StatusBar = _
                    ColumnLetter(ws, colNumber) & _
                    "열 처리 중... " & _
                    processCount & "개 그룹 완료"

                DoEvents

            End If

            i = endRow + 1

        Else

            i = i + 1

        End If

    Loop

End Sub


'=============================================================
' 배열 데이터 줄바꿈 결합
'=============================================================

Private Function CombineArrayText( _
    ByRef arr As Variant, _
    ByVal startRow As Long, _
    ByVal endRow As Long) As String

    Dim i As Long
    Dim result As String
    Dim v As String

    result = ""

    For i = startRow To endRow

        v = Trim(CStr(arr(i, 1)))

        If v <> "" Then

            If result <> "" Then
                result = result & Chr(10)
            End If

            result = result & CStr(arr(i, 1))

        End If

    Next i

    CombineArrayText = result

End Function


'=============================================================
' 열 번호 → 열 문자
'=============================================================

Private Function ColumnLetter( _
    ByVal ws As Worksheet, _
    ByVal colNumber As Long) As String

    ColumnLetter = _
        Split(ws.Cells(1, colNumber).Address(False, False), "1")(0)

End Function


'=============================================================
' 선택한 열 표시
'=============================================================

Private Function GetSelectedColumnsText( _
    ByVal ws As Worksheet, _
    ByVal targetCols As Variant) As String

    Dim i As Long
    Dim result As String

    result = ""

    For i = LBound(targetCols) To UBound(targetCols)

        If result <> "" Then
            result = result & " / "
        End If

        result = result & _
            ColumnLetter(ws, CLng(targetCols(i)))

    Next i

    GetSelectedColumnsText = result

End Function