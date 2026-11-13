Option Explicit

Sub 선택열_병합셀_자동높이_최적화()

    Dim ws As Worksheet
    Dim tempWs As Worksheet

    Dim lastRow As Long
    Dim r As Long

    Dim cell As Range
    Dim area As Range

    Dim txt As String

    Dim totalWidth As Double
    Dim currentHeight As Double
    Dim neededHeight As Double

    Dim rowCount As Long
    Dim addHeight As Double

    Dim c As Range

    Dim count As Long

    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldCalc As XlCalculation
    Dim oldAlerts As Boolean

    '=========================================================
    ' 열 선택 관련 변수
    '=========================================================

    Dim inputCols As String
    Dim colNames As Variant
    Dim colName As Variant
    Dim colNumber As Long

    Dim targetCols As Collection
    Dim selectedCol As Variant

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    '=========================================================
    ' 병합 셀 높이를 조정할 열 선택
    '=========================================================

    inputCols = InputBox( _
        "병합 셀 높이를 조정할 열을 입력하세요." & vbCrLf & vbCrLf & _
        "예: A, B, D, F" & vbCrLf & _
        "여러 열은 쉼표(,)로 구분하세요.", _
        "병합할 열 입력")

    ' 취소 또는 빈칸
    If Trim(inputCols) = "" Then
        Exit Sub
    End If

    '=========================================================
    ' 입력된 열을 확인
    '=========================================================

    Set targetCols = New Collection

    colNames = Split(inputCols, ",")

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

                ' 중복 열은 한 번만 추가
                On Error Resume Next
                targetCols.Add colNumber, CStr(colNumber)
                On Error GoTo ErrorHandler

            End If

        End If

    Next colName

    '=========================================================
    ' 잘못된 입력
    '=========================================================

    If targetCols.Count = 0 Then

        MsgBox "올바른 열을 입력하세요." & vbCrLf & vbCrLf & _
               "예: A, B, D, F", _
               vbExclamation

        Exit Sub

    End If

    '=========================================================
    ' Excel 설정 저장
    '=========================================================

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    oldAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Application.StatusBar = "선택한 열의 병합 셀을 확인하는 중..."

    '=========================================================
    ' 선택된 열을 하나씩 처리
    '=========================================================

    For Each selectedCol In targetCols

        '=====================================================
        ' 마지막 행
        '=====================================================

        lastRow = ws.Cells( _
                    ws.Rows.Count, _
                    CLng(selectedCol)).End(xlUp).Row

        '=====================================================
        ' 임시 시트 생성
        '=====================================================

        Set tempWs = ws.Parent.Worksheets.Add

        With tempWs

            .Visible = xlSheetVisible

            ' 임시 시트 기본 설정
            .Cells.Clear

        End With

        '=====================================================
        ' 선택한 열 검색
        '=====================================================

        r = 1

        Do While r <= lastRow

            Set cell = ws.Cells(r, CLng(selectedCol))

            '=================================================
            ' 병합 셀
            '=================================================

            If cell.MergeCells Then

                Set area = cell.MergeArea

                '=================================================
                ' 병합 영역의 첫 셀에서만 처리
                '=================================================

                If cell.Address = area.Cells(1, 1).Address Then

                    txt = CStr(area.Cells(1, 1).Value)

                    If Len(Trim(txt)) > 0 Then

                        '=========================================
                        ' 실제 병합 영역의 전체 너비 계산
                        '=========================================

                        totalWidth = 0

                        For Each c In area.Columns

                            totalWidth = _
                                totalWidth + c.Width

                        Next c

                        '=========================================
                        ' 현재 병합 영역의 총 높이
                        '=========================================

                        currentHeight = 0

                        For Each c In area.Rows

                            currentHeight = _
                                currentHeight + c.RowHeight

                        Next c

                        rowCount = area.Rows.Count

                        '=========================================
                        ' 임시 셀 초기화
                        '=========================================

                        With tempWs.Range("A1")

                            .Clear

                            .Value = txt

                            .WrapText = True

                            ' 원본 글꼴 그대로 복사
                            .Font.Name = _
                                area.Cells(1, 1).Font.Name

                            .Font.Size = _
                                area.Cells(1, 1).Font.Size

                            .Font.Bold = _
                                area.Cells(1, 1).Font.Bold

                            .Font.Italic = _
                                area.Cells(1, 1).Font.Italic

                            .Font.Underline = _
                                area.Cells(1, 1).Font.Underline

                            ' 원본 병합 영역의 너비를
                            ' 임시 셀의 너비로 맞춤
                            .ColumnWidth = _
                                ColumnWidthFromPoints(totalWidth)

                            ' 임시 셀 높이 초기화
                            .RowHeight = 15

                        End With

                        '=========================================
                        ' 임시 셀에서만 AutoFit
                        '=========================================

                        tempWs.Range("A1").EntireRow.AutoFit

                        neededHeight = _
                            tempWs.Range("A1").RowHeight

                        '=========================================
                        ' 여유 공간
                        '=========================================

                        neededHeight = neededHeight * 1.15

                        '=========================================
                        ' 최소 높이
                        '=========================================

                        If neededHeight < rowCount * 18 Then

                            neededHeight = rowCount * 18

                        End If

                        '=========================================
                        ' Excel 최대 높이
                        '=========================================

                        If neededHeight > 409.5 Then

                            neededHeight = 409.5

                        End If

                        '=========================================
                        ' 현재 높이보다 부족할 때만 변경
                        '=========================================

                        If currentHeight < neededHeight Then

                            addHeight = _
                                (neededHeight - currentHeight) / rowCount

                            For Each c In area.Rows

                                c.RowHeight = _
                                    c.RowHeight + addHeight

                            Next c

                        End If

                        count = count + 1

                    End If

                End If

                '=================================================
                ' 병합 영역을 한 번에 건너뜀
                '=================================================

                r = area.Row + area.Rows.Count

            Else

                r = r + 1

            End If

            '=====================================================
            ' 진행 상황
            '=====================================================

            If count > 0 Then

                If count Mod 50 = 0 Then

                    Application.StatusBar = _
                        ColumnLetter(ws, CLng(selectedCol)) & _
                        "열 병합 셀 처리 중... " & _
                        count & "개"

                    DoEvents

                End If

            End If

        Loop

        '=====================================================
        ' 임시 시트 삭제
        '=====================================================

        tempWs.Delete

        Set tempWs = Nothing

    Next selectedCol

    '=========================================================
    ' Excel 복구
    '=========================================================

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = False

    MsgBox "완료되었습니다." & vbCrLf & vbCrLf & _
           "선택한 열 : " & _
           GetSelectedColumnsText(ws, targetCols) & vbCrLf & _
           "병합 셀 : " & count & "개 확인" & vbCrLf & _
           "내용이 잘리지 않도록 높이를 조정했습니다.", _
           vbInformation

    Exit Sub


ErrorHandler:

    On Error Resume Next

    If Not tempWs Is Nothing Then

        tempWs.Delete

    End If

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = False

    MsgBox "오류가 발생했습니다." & vbCrLf & _
           "오류 번호 : " & Err.Number & vbCrLf & _
           "오류 내용 : " & Err.Description, _
           vbCritical

End Sub


'=============================================================
' Points → Excel ColumnWidth 근사 변환
'=============================================================

Private Function ColumnWidthFromPoints( _
    ByVal points As Double) As Double

    ' Excel 기본 글꼴 기준 근사값
    ColumnWidthFromPoints = points / 7

    If ColumnWidthFromPoints < 0.1 Then

        ColumnWidthFromPoints = 0.1

    End If

    If ColumnWidthFromPoints > 255 Then

        ColumnWidthFromPoints = 255

    End If

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
    ByVal targetCols As Collection) As String

    Dim i As Long
    Dim result As String

    result = ""

    For i = 1 To targetCols.Count

        If result <> "" Then
            result = result & " / "
        End If

        result = result & _
            ColumnLetter(ws, CLng(targetCols(i)))

    Next i

    GetSelectedColumnsText = result

End Function