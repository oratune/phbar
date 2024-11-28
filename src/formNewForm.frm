VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formNewForm 
   Caption         =   "Barchart Style"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   OleObjectBlob   =   "formNewForm.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "formNewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub optDay_Change()
  Call optTypeChangeProc
End Sub

Private Sub optMonth_Change()
  Call optTypeChangeProc
End Sub
Private Sub optWeek_Change()
  Call optTypeChangeProc
End Sub

Private Sub optTypeChangeProc()
  On Error GoTo errrtn
  If optMonth.Value = True Then
      lblUnit.Caption = "Months"
  Else
      lblUnit.Caption = "Weeks"
  End If
  Exit Sub
errrtn:
  If Err Then MsgBox "BarChart Error-" & Err.Description
End Sub

Private Sub btnOK_Click()
  Dim screenUpdateState
  
  Call configLoad
  Call setVersion
  
  '���� ���� ======================================
  screenUpdateState = Application.ScreenUpdating
  
  On Error GoTo errrtn
  
  ' �׷��� ���� �����
  Call bar_clearChart
  ' ���� �����
  Call formClearDrawArea
  
  set_Property "PHBAR_ActCount", txMaxAct.Text
  set_Property "PHBAR_ChartDur", txDuration.Text
  
  If Me.optDay.Value = "True" Then
    set_Property "PHBAR_ChartType", "Day"
    Call set_DailyForm
  ElseIf Me.optWeek.Value = "True" Then
    set_Property "PHBAR_ChartType", "Week"
    Call set_WeeklyForm
  Else
    set_Property "PHBAR_ChartType", "Mon"
    Call set_MonthForm
  End If
     
errrtn:
  Application.ScreenUpdating = screenUpdateState
  If Err Then MsgBox "BarChart Error-" & Err.Description
  Me.Hide
End Sub




Private Sub UserForm_Activate()
  Dim cnt
  cnt = get_Property("PHBAR_ActCnt")
  If Not IsNumeric(cnt) Then
    cnt = 100
  ElseIf cnt < 10 Then
    cnt = 10
  End If
  
  txMaxAct.Text = cnt
End Sub


' ���� ��� ����
Private Sub set_MonthForm()
  On Error GoTo errrtn
  Dim sh As Excel.Worksheet
  Dim strDate, iDuration
  Dim startDate
  Dim strWeek(7) As String
  Dim rowCnt As Long
  Dim i As Integer
  Dim curSH As String
  
  Set sh = Application.ActiveSheet
  
  curSH = sh.CodeName
  If curSH = "" Then curSH = sh.Name
  
   
  ' ������ ����
  strDate = txStDtc.Text
  
  startDate = DateSerial(Mid(strDate, 1, 4), Mid(strDate, 5, 2), 1)

  ' �Ⱓ������ ���� ��� ����
  iDuration = txDuration.Value
  
  If Not IsNumeric(iDuration) Then   ' ���� üũ
    MsgBox iDuration & " Is invalid Duration Data. Please Enter Numeric Value"
    Exit Sub
  End If
  
  If iDuration > (16300 - PHBAR_COL_BarLeft) Then ' �Ⱓ üũ
    MsgBox "The [Limit of period] is exceed Excel Limit"
    Exit Sub
  End If
  
  rowCnt = txMaxAct.Value
  If Not IsNumeric(rowCnt) Then   ' ���� üũ
    MsgBox rowCnt & " Is invalid Row Count Data. Please Enter Numeric Value"
    Exit Sub
  End If
  
  '��¥ǥ��
  For i = 1 To iDuration
    sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft + i - 1).Value = i
    sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft + i - 1).Value = DateAdd("m", i - 1, startDate)
  Next
    
  ' ǥ�� ����
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft - 1 + iDuration))
     .Interior.ColorIndex = 36
  End With
  
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft - 1 + iDuration))
    .Interior.ColorIndex = 35
    .NumberFormatLocal = "yy/mm"  ' ��� ǥ��
  End With
  
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_TitleTop + 1, 16300))
    .HorizontalAlignment = xlCenter
  End With
    
    ' ���� ǥ��
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_TitleTop + rowCnt - 1, PHBAR_COL_BarLeft + iDuration - 1))
      .Borders(xlDiagonalDown).LineStyle = xlNone
      .Borders(xlDiagonalUp).LineStyle = xlNone
      With .Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .Weight = xlHairline
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .Weight = xlHairline
          .ColorIndex = xlAutomatic
      End With
  End With
  
  sh.Range(sh.Columns(PHBAR_COL_BarLeft), sh.Columns(PHBAR_COL_BarLeft + iDuration)).ColumnWidth = 20
  sh.Cells(1, 1).Select
errrtn:
  If Err Then MsgBox "BarChart Error-" & Err.Description
End Sub

Private Sub set_WeeklyForm()
  On Error GoTo errrtn
  Dim sh As Excel.Worksheet
  Dim strDate, iDuration
  Dim startDate
  Dim strWeek(7) As String
  Dim rowCnt As Long
  Dim i As Integer
  Dim curSH As String
  
  Set sh = Application.ActiveSheet
  curSH = sh.CodeName
  If curSH = "" Then curSH = sh.Name
  
  strWeek(0) = "X"
  strWeek(1) = "S"
  strWeek(2) = "M"
  strWeek(3) = "T"
  strWeek(4) = "W"
  strWeek(5) = "T"
  strWeek(6) = "F"
  strWeek(7) = "S"

  ' ������ ����
  strDate = txStDtc.Text
  
  startDate = DateSerial(Mid(strDate, 1, 4), Mid(strDate, 5, 2), Mid(strDate, 7, 2))
  For i = 0 To 7
    If Weekday(startDate) = 1 Then Exit For
    startDate = startDate - 1
  Next
  sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft).Value = startDate

  ' �Ⱓ������ ���� ��� ����
  iDuration = txDuration.Value
  If Not IsNumeric(iDuration) Then   ' ���� üũ
    MsgBox iDuration & " Is invalid Duration Data. Please Enter Numeric Value"
    Exit Sub
  End If
  
  If (iDuration * 7) > (16300 - PHBAR_COL_BarLeft) Then ' �Ⱓ üũ
    MsgBox "The [Limit of period] is exceed Excel Limit"
    Exit Sub
  End If

  
  rowCnt = txMaxAct.Value
  If Not IsNumeric(rowCnt) Then   ' ���� üũ
    MsgBox rowCnt & " Is invalid Row Count Data. Please Enter Numeric Value"
    Exit Sub
  End If
  
  '��¥ �� ���� ǥ��
  For i = 0 To iDuration * 7 - 1
    sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft + i).Value = strWeek(Weekday(startDate + i))
  Next
  
  For i = 0 To iDuration - 1
    sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft + i * 7).Value = startDate + i * 7
    sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft + i * 7 + 3).Value = "~"
    sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft + i * 7 + 4).Value = startDate + i * 7 + 6
    sh.Range(sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft + i * 7), sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft + i * 7 + 2)).Merge
    sh.Range(sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft + i * 7 + 4), sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft + i * 7 + 6)).Merge
  Next
    
  ' ����ǥ��
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft - 1 + 7 * iDuration))
    .Interior.ColorIndex = 36
  End With
  
  ' ����ǥ��
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft - 1 + 7 * iDuration))
    .Interior.ColorIndex = 35
    .NumberFormatLocal = "mm/dd"
  End With
  
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_TitleTop + 1, 16300))
    .HorizontalAlignment = xlCenter
  End With
  
  ' �ֺ� ����
  For i = 1 To iDuration
    ' Week ǥ�ú�
    sh.Cells(PHBAR_ROW_TitleTop - 1, PHBAR_COL_BarLeft + 7 * (i - 1)).Value = CStr(i) & " Week"
    With sh.Range(sh.Cells(PHBAR_ROW_TitleTop - 1, PHBAR_COL_BarLeft + 7 * (i - 1)), sh.Cells(PHBAR_ROW_TitleTop - 1, PHBAR_COL_BarLeft - 1 + 7 * i))
      .Merge
      .Font.Bold = True
    End With
  
    ' ���� ǥ��
    With sh.Range(sh.Cells(PHBAR_ROW_TitleTop - 1, PHBAR_COL_BarLeft + 7 * (i - 1)), sh.Cells(PHBAR_ROW_DataTop + rowCnt - 1, PHBAR_COL_BarLeft - 1 + 7 * i))
      .Borders(xlDiagonalDown).LineStyle = xlNone
      .Borders(xlDiagonalUp).LineStyle = xlNone
      With .Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .Weight = xlHairline
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .Weight = xlHairline
          .ColorIndex = xlAutomatic
      End With
    End With
  Next
  
  sh.Range(sh.Columns(PHBAR_COL_BarLeft), sh.Columns(PHBAR_COL_BarLeft + 7 * iDuration)).ColumnWidth = 2

  sh.Cells(1, 1).Select

errrtn:
  If Err Then MsgBox "BarChart Error-" & Err.Description
End Sub

' �ϰ� ��� ����
Private Sub set_DailyForm()
  On Error GoTo errrtn
  Dim sh As Excel.Worksheet
  Dim strDate, iDuration
  Dim startDate
  Dim strWeek(7) As String
  Dim rowCnt As Long
  Dim i As Integer
  Dim curSH As String
  
  Set sh = Application.ActiveSheet
  curSH = sh.CodeName
  If curSH = "" Then curSH = sh.Name
  
  strWeek(0) = "N/A"
  strWeek(1) = "Sun"
  strWeek(2) = "Mon"
  strWeek(3) = "Tue"
  strWeek(4) = "Wed"
  strWeek(5) = "Thu"
  strWeek(6) = "Fri"
  strWeek(7) = "Sat"
  
  ' ������ ����
  strDate = txStDtc.Text
  
  startDate = DateSerial(Mid(strDate, 1, 4), Mid(strDate, 5, 2), Mid(strDate, 7, 2))
  For i = 0 To 7
    If Weekday(startDate) = 1 Then Exit For
    startDate = startDate - 1
  Next
  sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft).Value = startDate

  ' �Ⱓ������ ���� ��� ����
  iDuration = txDuration.Value
  If Not IsNumeric(iDuration) Then   ' ���� üũ
    MsgBox iDuration & " Is invalid Duration Data. Please Enter Numeric Value"
    Exit Sub
  End If
  
  If (iDuration * 7) > (16300 - PHBAR_COL_BarLeft) Then ' �Ⱓ üũ
    MsgBox "The [Limit of period] is exceed Excel Limit"
    Exit Sub
  End If
  
  rowCnt = txMaxAct.Value
  If Not IsNumeric(rowCnt) Then   ' ���� üũ
    MsgBox rowCnt & " Is invalid Row Count Data. Please Enter Numeric Value"
    Exit Sub
  End If
  
  '��¥ �� ���� ǥ��
  For i = 0 To iDuration * 7 - 1
    sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft + i).Value = strWeek(Weekday(startDate + i))
    sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft + i).Value = startDate + i
  Next
    
  ' ����ǥ��
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft - 1 + 7 * iDuration))
    .Interior.ColorIndex = 36
  End With
  
  ' ����ǥ��
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft - 1 + 7 * iDuration))
    .Interior.ColorIndex = 35
    .NumberFormatLocal = "mm/dd"
  End With
  
  With sh.Range(sh.Cells(PHBAR_ROW_TitleTop, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_TitleTop + 1, 16300))
    .HorizontalAlignment = xlCenter
  End With
  
  ' �ֺ� ����
  For i = 1 To iDuration
    ' Week ǥ�ú�
    sh.Cells(PHBAR_ROW_TitleTop - 1, PHBAR_COL_BarLeft + 7 * (i - 1)).Value = CStr(i) & " Week"
    With sh.Range(sh.Cells(PHBAR_ROW_TitleTop - 1, PHBAR_COL_BarLeft + 7 * (i - 1)), sh.Cells(PHBAR_ROW_TitleTop - 1, PHBAR_COL_BarLeft - 1 + 7 * i))
      .Merge
      .Font.Bold = True
      .HorizontalAlignment = xlCenter
    End With
  
    ' ���� ǥ��
    With sh.Range(sh.Cells(PHBAR_ROW_TitleTop - 1, PHBAR_COL_BarLeft + 7 * (i - 1)), sh.Cells(PHBAR_ROW_DataTop + rowCnt - 1, PHBAR_COL_BarLeft - 1 + 7 * i))
      .Borders(xlDiagonalDown).LineStyle = xlNone
      .Borders(xlDiagonalUp).LineStyle = xlNone
      With .Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .Weight = xlHairline
          .ColorIndex = xlAutomatic
      End With
      With .Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .Weight = xlHairline
          .ColorIndex = xlAutomatic
      End With
    End With
  Next
  
  sh.Cells(1, PHBAR_COL_BarLeft).Select
  sh.Range(sh.Columns(PHBAR_COL_BarLeft), sh.Columns(PHBAR_COL_BarLeft + 7 * iDuration)).ColumnWidth = 5
  sh.Cells(1, 1).Select
  
errrtn:
  If Err Then MsgBox "BarChart Error-" & Err.Description
End Sub

