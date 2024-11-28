Attribute VB_Name = "BarDraw"
Option Explicit

' Bar 새로고침 - 전체
Sub bar_Refresh_full()
    On Error GoTo errrtn
   
    Dim screenUpdateState, statusBarState, calcState, eventsState, displayPageBreakState
    '현재 상태 ======================================
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    displayPageBreakState = ActiveSheet.DisplayPageBreaks 'note this is a sheet-level setting

    '이벤트 제거 ====================================
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False 'note this is a sheet-level setting
   
    Call bar_Refresh(0, 0)

errrtn:
    '작업 완료 후 원상복귀 ==========================
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
    ActiveSheet.DisplayPageBreaks = displayPageBreakState 'note this is a sheet-level setting
   
End Sub

' Bar 새로고침 - 영역
Sub bar_Refresh_range()

    Dim row_top, row_end
   
    On Error GoTo errrtn
   
    Dim screenUpdateState, statusBarState, calcState, eventsState, displayPageBreakState
    '현재 상태 ======================================
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    displayPageBreakState = ActiveSheet.DisplayPageBreaks 'note this is a sheet-level setting

    '이벤트 제거 ====================================
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False 'note this is a sheet-level setting
   
    If TypeName(Selection) <> "Range" Then
      MsgBox "Select Range to Refresh"
      Exit Sub
    End If
   
    row_top = Selection.Cells(1, 1).row
    row_end = row_top + Selection.Rows.Count - 1

   
    Call bar_Refresh(row_top, row_end)
    Cells(row_end + 1, 1).Select

errrtn:
    '작업 완료 후 원상복귀 ==========================
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
    ActiveSheet.DisplayPageBreaks = displayPageBreakState 'note this is a sheet-level setting

End Sub



Private Sub bar_Refresh(row_top, row_end)
    Dim sh As Excel.Worksheet
    Dim formStartDate
    Dim i
   
    Dim stDate    ' 시작일
    Dim endDate   ' 종료일
    Dim duration  ' 기간
   
    Dim barX1, barX2, barX3, barY1, barY2, barY3
    Dim barWidth, barHeight
    Dim actType
   
    Dim blankCnt As Long
    Dim old_row As Long
   
    old_row = Selection.Cells(1, 1).row
   
    configLoad
    Set sh = Application.ActiveSheet
      
    If row_top < PHBAR_ROW_DataTop Then row_top = PHBAR_ROW_DataTop
   
    ' 그려진 도형 지우기
    For i = sh.Shapes.Count To 1 Step -1
      If (sh.Shapes(i).Type = msoLine) Or (sh.Shapes(i).Type = msoAutoShape) Then '20101105 필터 등의 처리시 오류 방지
        If sh.Shapes(i).TopLeftCell.row >= row_top And _
           sh.Shapes(i).TopLeftCell.Column >= PHBAR_COL_BarLeft - 1 Then  '20090722 Office 2007버젼 오류 해결
           If sh.Shapes(i).TopLeftCell.row <= row_end Or _
              row_end = 0 Then sh.Shapes(i).Delete
        End If
      End If
    Next
   
    formStartDate = sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft).Value  ' 공정표 시작일
    
    If Not IsDate(formStartDate) Then
      MsgBox "Invalid [Chart Start Date] at row=" & CStr(PHBAR_ROW_TitleTop + 1) & ", col=" & CStr(PHBAR_COL_BarLeft)
      Exit Sub
    End If
    
    formStartDate = Int(formStartDate)
    Height_of_Row = ActiveSheet.Cells(PHBAR_ROW_DataTop, 1).Height       ' 행 높이
   
    i = row_top
    blankCnt = 0

    On Error GoTo err_bar
    Do Until blankCnt > 5
   
      If PHBAR_ActCnt <= i - row_top Then Exit Do
      If i > row_end And row_end <> 0 Then Exit Do

      If checkRowBlank(sh, i) Then
        blankCnt = blankCnt + 1
      Else
        actType = sh.Cells(i, PHBAR_COL_ActType).Value
        If actType = "" Then actType = "A"
        actType = Left(actType, 1)
       
        ' 계획선 그리기
        stDate = validDate(sh.Cells(i, PHBAR_COL_PLANST).Value)  ' 시작일
        endDate = validDate(sh.Cells(i, PHBAR_COL_PlanEnd).Value)  ' 종료일
       
        If stDate > endDate Then endDate = stDate
   
        If stDate <> "" Then
          drawBarchart formStartDate, i, stDate, endDate, "Plan", actType
          duration = endDate - stDate + 1 ' 기간
        End If
       
        ' 실적선 그리기
        If PHBAR_USEActual Then
          stDate = sh.Cells(i, PHBAR_COL_ActST).Value  ' 시작일
          endDate = sh.Cells(i, PHBAR_COL_ActEnd).Value  ' 종료일
       
          If stDate <> "" Then
            If endDate = "" Then endDate = stDate + duration - 1
            If stDate > endDate Then endDate = stDate
            drawBarchart formStartDate, i, stDate, endDate, "Result", actType
          End If
        End If
       
      End If
      i = i + 1
    Loop
   
exit_bar:
    Cells(old_row, 1).Select
    Exit Sub
    
err_bar:
    MsgBox "Error : Please check data at Line No =" & CStr(i) & Chr(10) & Err.Description
    On Error GoTo 0
    Resume exit_bar
End Sub



Private Sub drawBarchart(formStartDate, i, stDate, endDate, barType, actType)
    Dim C_COLOR_PLAN, C_COLOR_RESULT
   
    Dim progRate, progDate     ' 진도율
   
    Dim stOffset      ' 시작옵셧
    Dim endOffset     '끝지점옵셋
   
    Dim HeightOfRow
    Dim stOut, endOut, outOfRange  ' 시작지점, 끝지점이 범위를 벗어낫는지 여부
   
   
    Dim barX1, barX2, barXProg, barY1, barY2, barY3
    Dim barWidth, barHeight, barWidthProg
   
    Dim resultColor
   
    stOut = False
    endOut = False
    outOfRange = False ' 시작 및 끝점이 공정표 내에 표시되지 않음
   
    If ActiveSheet.Cells(i, 1).Height < Height_of_Row Then
      HeightOfRow = ActiveSheet.Cells(i, 1).Height
    Else
      HeightOfRow = Height_of_Row
    End If

    If barType = "Result" Then
      progRate = ActiveSheet.Cells(i, PHBAR_COL_Progress).Value  ' 진도율
      progDate = endDate
    End If

    
    ' 공정표 기간 초과시 처리
    If get_Property("PHBAR_ChartType") = "Mon" Then
      If (endDate - formStartDate) / 30 > (C_LIMIT_COL - PHBAR_COL_BarLeft - 3) Then
        endDate = formStartDate + (C_LIMIT_COL - PHBAR_COL_BarLeft - 3) * 30
        stOut = True
      End If
     
      If (stDate - formStartDate) >= (PHBAR_ChartDur * 30) Then
        'stDate = formStartDate + PHBAR_ChartDur * 30
        outOfRange = True
      End If
     
      If (endDate - formStartDate) >= (PHBAR_ChartDur * 30) Then
        endDate = formStartDate + PHBAR_ChartDur * 30
        endOut = True
      End If
    Else  '주,일간
      If (endDate - formStartDate) > (C_LIMIT_COL - PHBAR_COL_BarLeft - 3) Then
        endDate = formStartDate + C_LIMIT_COL - PHBAR_COL_BarLeft - 3
        endOut = True
      End If
       
      If (stDate - formStartDate) >= (PHBAR_ChartDur * 7) Then
        'stDate = formStartDate + PHBAR_ChartDur * 7 - 1
        'stOut = True
        outOfRange = True
      End If
     
      If (endDate - formStartDate) >= (PHBAR_ChartDur * 7) Then
        endDate = formStartDate + PHBAR_ChartDur * 7 - 1
        endOut = True
      End If

    End If
   
    If PHBAR_ChartType = "Mon" Then
      stOffset = (Year(stDate) - Year(formStartDate)) * 12 + Month(stDate) - Month(formStartDate)
      endOffset = (Year(endDate) - Year(formStartDate)) * 12 + Month(endDate) - Month(formStartDate)
    Else
      stOffset = stDate - formStartDate       ' 시작옵셧
      endOffset = endDate - formStartDate   '끝지점옵셋
    End If
    
    If outOfRange Then Exit Sub
    
    If get_Property("PHBAR_ChartType") = "Mon" Then
      If stOffset < 0 Then
        barX1 = Cells(i, PHBAR_COL_BarLeft).Left
        stOffset = True
      Else
        barX1 = Cells(i, PHBAR_COL_BarLeft + stOffset).Left + Day(stDate) / 30 * Cells(i, PHBAR_COL_BarLeft + stOffset).Width
      End If
      
      If endOffset < 0 Then
        outOfRange = True
        barX2 = Cells(i, PHBAR_COL_BarLeft).Left
      Else
        barX2 = Cells(i, PHBAR_COL_BarLeft + endOffset).Left + Day(endDate) / 30 * Cells(i, PHBAR_COL_BarLeft + endOffset).Width
      End If
    Else
    
      If stOffset < 0 Then
        stOut = True
        stOffset = 0
      End If
      If endOffset < 0 Then
        endOffset = 0
        outOfRange = True
      End If
      
      barX1 = Cells(i, PHBAR_COL_BarLeft + stOffset).Left  ' 시작점
      barX2 = Cells(i, PHBAR_COL_BarLeft + endOffset).Left + Cells(i, PHBAR_COL_BarLeft + endOffset).Width  ' 종료점
    End If
    
    If outOfRange Then Exit Sub
   
   
    If barType = "Plan" Then
        If PHBAR_USEActual Or PHBAR_USEResource Then
          barY1 = Cells(i, 1).Top + 1 / 8 * HeightOfRow
          barY2 = Cells(i, 1).Top + 4 / 8 * HeightOfRow
        Else
          barY1 = Cells(i, 1).Top + 2 / 8 * HeightOfRow
          barY2 = Cells(i, 1).Top + 6 / 8 * HeightOfRow
        End If
       
        barWidth = barX2 - barX1
        barHeight = barY2 - barY1
       
        If barWidth < 1 Then barWidth = 1
        If barHeight < 1 Then barHeight = 1
       
        If actType = "G" Then  ' Act.Group
            barY3 = Cells(i, 1).Top + 3 / 7 * HeightOfRow
            With ActiveSheet.Shapes.AddLine(barX1, barY3, barX2, barY3)
              .Line.Weight = 4
              .Line.DashStyle = msoLineSolid
              .Line.Style = msoLineSingle
              .Line.Visible = msoTrue
              .Line.ForeColor.RGB = COLOR_GROUPPLAN
              .Line.BackColor.RGB = RGB(255, 255, 255)
              .Line.BeginArrowheadLength = msoArrowheadLengthMedium
              .Line.BeginArrowheadWidth = msoArrowheadWidthMedium
              .Line.BeginArrowheadStyle = msoArrowheadDiamond
              .Line.EndArrowheadLength = msoArrowheadLengthMedium
              .Line.EndArrowheadWidth = msoArrowheadWidthMedium
              .Line.EndArrowheadStyle = msoArrowheadDiamond
            End With
        ElseIf actType = "M" Then  'Milestone
            With ActiveSheet.Shapes.AddShape(msoShapeDiamond, barX1, barY1, 12#, 15#)
              .Fill.Visible = msoTrue
              .Fill.Solid
              .Fill.ForeColor.RGB = COLOR_MSPLAN
              .Fill.Transparency = 0#
              .Line.Weight = 0.5
              .Line.DashStyle = msoLineSolid
              .Line.Style = msoLineSingle
              .Line.Transparency = 0#
              .Line.Visible = msoTrue
              .Line.ForeColor.SchemeColor = 64
              .Line.BackColor.RGB = RGB(255, 255, 255)
            End With
       
        Else
            With ActiveSheet.Shapes.AddShape(msoShapeRectangle, barX1, barY1, barWidth, barHeight)
              .Name = "pBar" & CStr(i)
              .Fill.ForeColor.RGB = COLOR_ACTPLAN
              .Fill.Visible = msoTrue
              .Fill.Solid
              .Line.Weight = 0.5
              .Line.ForeColor.RGB = RGB(200, 200, 200)
              .Line.Visible = msoTrue
            End With
           
        End If
       
    ElseIf barType = "Result" Then
        barY1 = Cells(i, 1).Top + 4 / 8 * HeightOfRow
        barY2 = Cells(i, 1).Top + 7 / 8 * HeightOfRow
        ' 계획표
        barWidth = barX2 - barX1
        barHeight = barY2 - barY1
        If barWidth < 1 Then barWidth = 1
        If barHeight < 1 Then barHeight = 1
       
        If actType = "G" Then  ' Act.Group
            barY3 = Cells(i, 1).Top + 5 / 7 * HeightOfRow
            With ActiveSheet.Shapes.AddLine(barX1, barY3, barX2, barY3)
              .Line.Weight = 4
              .Line.DashStyle = msoLineSolid
              .Line.Style = msoLineSingle
              .Line.Visible = msoTrue
              .Line.ForeColor.RGB = COLOR_GROUPACTUAL
              .Line.BackColor.RGB = RGB(255, 255, 255)
              .Line.BeginArrowheadLength = msoArrowheadLengthMedium
              .Line.BeginArrowheadWidth = msoArrowheadWidthMedium
              .Line.BeginArrowheadStyle = msoArrowheadDiamond
              .Line.EndArrowheadLength = msoArrowheadLengthMedium
              .Line.EndArrowheadWidth = msoArrowheadWidthMedium
              .Line.EndArrowheadStyle = msoArrowheadDiamond
            End With
        ElseIf actType = "M" Then  'Milestone
            barY3 = Cells(i, 1).Top + 1 / 7 * HeightOfRow
            With ActiveSheet.Shapes.AddShape(msoShapeDiamond, barX1, barY3, 12#, 15#)
              .Fill.Visible = msoTrue
              .Fill.Solid
              .Fill.ForeColor.RGB = COLOR_MSACTUAL
              .Fill.Transparency = 0#
              .Line.Weight = 0.5
              .Line.DashStyle = msoLineSolid
              .Line.Style = msoLineSingle
              .Line.Visible = msoTrue
              .Line.ForeColor.RGB = RGB(200, 200, 200)
              .Line.BackColor.RGB = RGB(255, 255, 255)
            End With
       
        Else
            With ActiveSheet.Shapes.AddShape(msoShapeRectangle, barX1, barY1, barWidth, barHeight)
              .Name = "rBar" & CStr(i)
              .Fill.ForeColor.RGB = RGB(250, 250, 50)
              .Fill.Visible = msoTrue
              .Fill.Solid
              .Line.Visible = msoTrue
              .Line.Weight = 0.5
              .Line.ForeColor.RGB = RGB(200, 200, 200)
            End With
           
            ' 진도 표시
            If IsNumeric(progRate) = False Then
                'MsgBox "Error:  Please check PROGRESS-data of row " & i, 48, "BAR-CHART"
                'Exit Sub
            Else
              If progRate > 0 Then
   
                If get_Property("PHBAR_ChartType") = "Mon" Then
                  endOffset = (Year(progDate) - Year(formStartDate)) * 12 + Month(progDate) - Month(formStartDate)
                  If endOffset < 0 Then endOffset = 0
                  barXProg = Cells(i, PHBAR_COL_BarLeft + endOffset).Left + Day(progDate) / 30 * Cells(i, PHBAR_COL_BarLeft + endOffset).Width
                Else
                  endOffset = progDate - formStartDate   '끝지점옵셋
                  If endOffset < 0 Then endOffset = 0
                  barXProg = Cells(i, PHBAR_COL_BarLeft + endOffset).Left + Cells(i, PHBAR_COL_BarLeft + endOffset).Width  ' 종료점
                End If
               
                barWidthProg = (barXProg - barX1) * progRate
                If barX2 < barX1 + barWidthProg Then barWidthProg = barX2 - barX1
                
                If barWidthProg > 0 Then
                  With ActiveSheet.Shapes.AddShape(msoShapeRectangle, barX1, barY1, barWidthProg, barHeight)
                    .Name = "zBar" & CStr(i)
                    .Fill.ForeColor.RGB = COLOR_ACTACTUAL
                    .Fill.Visible = msoTrue
                    .Fill.Solid
                    .Line.Visible = msoFalse
                  End With
                End If
              End If
            End If

        End If
    End If

End Sub


Sub bar_clearChart()
    Dim i, sh
  
    configLoad
    Set sh = Application.ActiveSheet
    
    For i = sh.Shapes.Count To 1 Step -1
      If (sh.Shapes(i).Type = msoLine) Or (sh.Shapes(i).Type = msoAutoShape) Then '20101105 필터 등의 처리시 오류 방지
        If sh.Shapes(i).TopLeftCell.row >= PHBAR_ROW_DataTop And _
           sh.Shapes(i).TopLeftCell.Column >= PHBAR_COL_BarLeft - 1 Then  '20090722 Office 2007버젼 오류 해결
           sh.Shapes(i).Delete
        End If
      End If
    Next

End Sub

Sub bar_clearAll()
    Dim i, sh
    Set sh = Application.ActiveSheet
    
    For i = sh.Shapes.Count To 1 Step -1
      If (sh.Shapes(i).Type = msoLine) Or (sh.Shapes(i).Type = msoAutoShape) Then
           sh.Shapes(i).Delete
      End If
    Next
End Sub




