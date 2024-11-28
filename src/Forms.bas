Attribute VB_Name = "Forms"
Option Explicit


' 새 공정표 양식
Sub makeNewChartSheet()
  On Error GoTo errrtn
  Dim sh As Excel.Worksheet
  Dim rng As Excel.Range
  Dim iTmp
  
  Call configLoad
    
  Dim maxAct As Long
  Dim re As Variant
  
  re = InputBox("Maximum Activity Drawing Limit?", "New Barchart Sheet", "300")
  If re = "" Then
    Exit Sub
  Else
    maxAct = re
  End If
  
  If ActiveWorkbook Is Nothing Then
    Workbooks.Add
  Else
    ActiveWorkbook.Sheets.Add
  End If
 
  
  Set sh = Application.ActiveSheet
  setVersion
  
  set_Property "PHBAR_ActCnt", CStr(maxAct)
  
  sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_ActID).Value = "ID"
  sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_ActDesc).Value = "Description"
  sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_ActType).Value = "T"
  
  sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_PLANST).Value = "Plan"
  sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_PLANST).Value = "Start"
  sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_PlanEnd).Value = "Finish"
  sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_PlanDur).Value = "Dur"
  
  sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_ActST).Value = "Actual"
  sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_ActST).Value = "Start"
  sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_ActEnd).Value = "Finish"
  sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_ActDur).Value = "Dur"
  sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_Progress).Value = "Prog."
  sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_Difference).Value = "Diff."
  
  Set rng = sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_ActType), sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_ActType))
  rng.AddComment
  rng.Comment.Text Text:= _
        "M : Milestone" & Chr(10) & "G : Group of Activity" & Chr(10) & "A : Activity (default)"
  
  Set rng = sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_PLANST), sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_PlanDur))
  rng.Merge
  rng.HorizontalAlignment = xlCenter
  
  Set rng = sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_ActST), sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_Progress))
  rng.Merge
  rng.HorizontalAlignment = xlCenter
  
  ' 셀 표시형식 지정
  sh.Range(sh.Columns(PHBAR_COL_ActDesc), sh.Columns(PHBAR_COL_ActDesc)).NumberFormatLocal = "@"
  
  sh.Range(sh.Columns(PHBAR_COL_PLANST), sh.Columns(PHBAR_COL_PlanEnd)).NumberFormatLocal = "yyyy-mm-dd"
  sh.Range(sh.Columns(PHBAR_COL_ActST), sh.Columns(PHBAR_COL_ActEnd)).NumberFormatLocal = "yyyy-mm-dd"
  
  sh.Range(sh.Columns(PHBAR_COL_PlanDur), sh.Columns(PHBAR_COL_PlanDur)).NumberFormatLocal = "0_ "
  sh.Range(sh.Columns(PHBAR_COL_ActDur), sh.Columns(PHBAR_COL_ActDur)).NumberFormatLocal = "0_ "
  
  sh.Range(sh.Columns(PHBAR_COL_Progress), sh.Columns(PHBAR_COL_Progress)).NumberFormatLocal = "0%"
  
  
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop, 1), sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft - 1))
    .Interior.ColorIndex = 34
    .HorizontalAlignment = xlCenter
  End With
  sh.Columns(PHBAR_COL_ActID).ColumnWidth = 5
  sh.Columns(PHBAR_COL_ActDesc).ColumnWidth = 26.6
  sh.Columns(PHBAR_COL_ActType).ColumnWidth = 4
  
  sh.Columns(PHBAR_COL_PLANST).ColumnWidth = 12
  sh.Columns(PHBAR_COL_PlanEnd).ColumnWidth = 12
  sh.Columns(PHBAR_COL_ActST).ColumnWidth = 12
  sh.Columns(PHBAR_COL_ActEnd).ColumnWidth = 12
  
  sh.Columns(PHBAR_COL_PlanDur).ColumnWidth = 4.3
  sh.Columns(PHBAR_COL_ActDur).ColumnWidth = 4.3
  sh.Columns(PHBAR_COL_Progress).ColumnWidth = 4.8
  sh.Columns(PHBAR_COL_Difference).ColumnWidth = 4.8
  
  ' 표제 영역 서식
  Set rng = sh.Range(sh.Cells(PHBAR_ROW_TitleTop, 1), sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft - 1))
  rng.Borders(xlDiagonalDown).LineStyle = xlNone
  rng.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
  Set rng = sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_PLANST), sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_PlanDur))
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
  Set rng = sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_ActST), sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_Progress))
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
  ' 데이터 영역 서식
  Set rng = sh.Range(sh.Cells(PHBAR_ROW_DataTop, 1), sh.Cells(PHBAR_ROW_DataTop + maxAct - 1, PHBAR_COL_BarLeft - 1))
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    rng.RowHeight = 21
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    
  Application.ActiveWindow.Zoom = 75
  
  sh.Cells(1, 1).Select
  
  
errrtn:
  If Err Then MsgBox "BarChart Error-" & Err.Description
End Sub


' 양식 다시 그리기
Sub formRedraw()
  Dim dt, curDate
  
  Call configLoad
  On Error GoTo errrtn
  
  Dim formNewForm As New formNewForm
  If formNewForm Is Nothing Then
    Set formNewForm = New formNewForm
  End If
  
  dt = ActiveSheet.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft).Value  ' 공정표 시작일
  If IsDate(dt) Then
    curDate = CDate(dt)
  Else
    curDate = Now
  End If
  
  formNewForm.txStDtc.Text = Format(curDate, "yyyymmdd")
  formNewForm.optWeek.SetFocus
  
  formNewForm.Show 1
  
  Set formNewForm = Nothing
  
  Exit Sub
errrtn:
  If Err Then MsgBox "BarChart Error-" & Err.Description
End Sub


Sub formClearDrawArea()
  Dim iColOffset, sh
  
  Call configLoad
  Set sh = Application.ActiveSheet
  
  'If PHBAR_ChartType = "Mon" Then
  '    iColOffset = PHBAR_COL_BarLeft + PHBAR_ChartDur - 1
  'Else
  '    iColOffset = PHBAR_COL_BarLeft + PHBAR_ChartDur * 7 - 1
  'End If
  
  ' 시트 내용지우기
  With sh.Range(sh.Columns(PHBAR_COL_BarLeft), sh.Columns(16300))
    .UnMerge
    .Clear
  End With
End Sub
