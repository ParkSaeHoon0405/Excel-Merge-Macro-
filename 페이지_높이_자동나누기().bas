Sub 페이지_높이_자동나누기()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim nextRow As Long
    Dim pageHeight As Double
    Dim usedHeight As Double
    Dim blockHeight As Double
    
    Dim c As Long
    Dim lastCol As Long
    Dim mergeArea As Range
    Dim testArea As Range
    
    Dim rr As Long
    Dim foundMerge As Boolean

    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' 페이지 높이
    pageHeight = 800

    ' 기존 페이지 나누기 삭제
    ws.ResetAllPageBreaks

    ' 마지막 사용 행
    lastRow = ws.Cells.Find( _
                What:="*", _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious).Row

    ' 마지막 사용 열
    lastCol = ws.Cells.Find( _
                What:="*", _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious).Column

    usedHeight = 0
    r = 1

    Do While r <= lastRow

        nextRow = r
        blockHeight = ws.Rows(r).Height
        foundMerge = False

        ' ----------------------------------------
        ' 현재 행에 병합 셀이 있는지 전체 열 검사
        ' ----------------------------------------
        For c = 1 To lastCol

            If ws.Cells(r, c).MergeCells Then

                Set mergeArea = ws.Cells(r, c).MergeArea

                ' 병합 범위가 현재 행에서 시작하는 경우
                If mergeArea.Row = r Then

                    ' 가장 아래 행까지 하나의 블록으로 처리
                    If mergeArea.Row + mergeArea.Rows.Count - 1 > nextRow Then
                        nextRow = mergeArea.Row + mergeArea.Rows.Count - 1
                    End If

                    foundMerge = True

                Else

                    ' 현재 행이 이미 이전 병합범위 안에 있다면
                    ' 그 병합범위 전체를 하나의 블록으로 처리
                    If mergeArea.Row < r Then

                        If mergeArea.Row + mergeArea.Rows.Count - 1 >= r Then

                            ' 이미 지나온 행과 연결된 병합이면
                            ' 페이지 나눔이 여기 들어가지 않도록 처리
                            nextRow = mergeArea.Row + mergeArea.Rows.Count - 1

                            foundMerge = True

                        End If

                    End If

                End If

            End If

        Next c

        ' ----------------------------------------
        ' 블록 전체 높이 계산
        ' ----------------------------------------
        blockHeight = 0

        For rr = r To nextRow
            blockHeight = blockHeight + ws.Rows(rr).Height
        Next rr

        ' ----------------------------------------
        ' 페이지 초과 여부
        ' ----------------------------------------
        If usedHeight + blockHeight > pageHeight Then

            ' 현재 페이지에 이미 내용이 있으면
            ' 현재 블록 시작 전에 페이지 나누기
            If usedHeight > 0 Then

                ws.HPageBreaks.Add Before:=ws.Rows(r)

                usedHeight = blockHeight

            Else

                ' 블록 하나가 페이지보다 큰 경우
                usedHeight = blockHeight

            End If

        Else

            usedHeight = usedHeight + blockHeight

        End If

        ' 다음 블록으로 이동
        r = nextRow + 1

    Loop

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "완료" & vbCrLf & _
           "페이지 높이 기준 : 800" & vbCrLf & _
           "병합 셀 중간 페이지 나눔 방지", vbInformation

End Sub