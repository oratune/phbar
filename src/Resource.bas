Attribute VB_Name = "Resource"
Option Explicit


' 리소스 배부 - 전체
Sub rsc_distrib_full()
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
   
    Call rsc_distrib(0, 0)

errrtn:
    '작업 완료 후 원상복귀 ==========================
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
    ActiveSheet.DisplayPageBreaks = displayPageBreakState 'note this is a sheet-level setting
   
End Sub

' 리소스 배부 - 영역
Sub rsc_distrib_range()

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

   
    Call rsc_distrib(row_top, row_end)
    Cells(row_end + 1, 1).Select

errrtn:
    '작업 완료 후 원상복귀 ==========================
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
    ActiveSheet.DisplayPageBreaks = displayPageBreakState 'note this is a sheet-level setting

End Sub

'리소스 배부
Private Sub rsc_distrib(row_top, row_end)
    Dim chkWeekday, stDate, endDate, aDay, blankCnt
    Dim formStartDate As Date
    Dim monStDate As Date, monEndDate As Date
    Dim monVal As Double
    Dim stOffset, endOffset As Long
    Dim i, iDay, iMon, iDur, iColOffset As Long
    Dim iRscBase As Double  ' 배부대상 모수가 되는 자원
    Dim iRscDist As Double  ' 배부되는 자원
    Dim sh As Excel.Worksheet
    
    configLoad  '설정사항 읽기
    
    Set sh = Application.ActiveSheet
    
    i = PHBAR_ROW_DataTop
    blankCnt = 0
    If PHBAR_HolidayType = "6" Then
      chkWeekday = 7
    ElseIf PHBAR_HolidayType = "5" Then
      chkWeekday = 6
    Else
      chkWeekday = 0
    End If
    
    If row_top < PHBAR_ROW_DataTop Then row_top = PHBAR_ROW_DataTop
    If row_end = 0 Then
      row_end = row_top + PHBAR_ActCnt
    ElseIf row_end < row_top Then
      row_end = row_top
    End If
    
    ' 내용지우기
    On Error Resume Next
    If PHBAR_ChartType = "Mon" Then
      iColOffset = PHBAR_COL_BarLeft + PHBAR_ChartDur - 1
    Else
      iColOffset = PHBAR_COL_BarLeft + PHBAR_ChartDur * 7 - 1
    End If
    If Not IsNumeric(iColOffset) Or iColOffset > C_LIMIT_COL Then iColOffset = C_LIMIT_COL
    sh.Range(sh.Cells(row_top, PHBAR_COL_BarLeft), sh.Cells(row_end, iColOffset)).ClearContents


    On Error GoTo err_bar
    i = row_top
    blankCnt = 0
    
    formStartDate = sh.Cells(PHBAR_ROW_TitleTop + 1, PHBAR_COL_BarLeft).Value  ' 공정표 시작일
    
    If Not IsDate(formStartDate) Then
      MsgBox "Invalid [Chart Start Date] at row=" & CStr(PHBAR_ROW_TitleTop + 1) & ", col=" & CStr(PHBAR_COL_BarLeft)
      Exit Sub
    End If
    
    Do Until blankCnt > 5
    
      If PHBAR_ActCnt <= i - row_top Then Exit Do
      If i > row_end And row_end <> 0 Then Exit Do

      iRscBase = sh.Cells(i, PHBAR_COL_Resource).Value '배분할 리소스
      
      stDate = validDate(sh.Cells(i, PHBAR_COL_PLANST).Value)
      endDate = validDate(sh.Cells(i, PHBAR_COL_PLANST + 1).Value)
      
      If Not IsNumeric(iRscBase) Then iRscBase = 0
      
      If checkRowBlank(sh, i) Then
        blankCnt = blankCnt + 1
      ElseIf iRscBase <> 0 And stDate <= endDate Then
        ' Duration 계산
        iDur = 0
        If PHBAR_HolidayType <> "7" Then
          For iDay = stDate To endDate
            If Weekday(iDay, vbMonday) < chkWeekday Then iDur = iDur + 1
          Next iDay
        Else
          iDur = endDate - stDate + 1
        End If
             
        ' 일별 분할 값 계산
        If Not IsNumeric(iDur) Or iDur = 0 Then
          iRscDist = 0
        Else
          iRscDist = iRscBase / iDur  ' 일별 분할값
        End If
        
            
        If PHBAR_ChartType = "Mon" Then
          stOffset = (Year(stDate) - Year(formStartDate)) * 12 + Month(stDate) - Month(formStartDate)
          endOffset = (Year(endDate) - Year(formStartDate)) * 12 + Month(endDate) - Month(formStartDate)
        Else
          stOffset = stDate - formStartDate       ' 시작옵셧
          endOffset = endDate - formStartDate   '끝지점옵셋
        End If
   
        If stOffset < 0 Then stOffset = 0
        If endOffset < 0 Then endOffset = 0
        
        If PHBAR_ChartType = "Mon" Then
          If endOffset >= PHBAR_ChartDur Then endOffset = PHBAR_ChartDur - 1
        Else
          If endOffset >= PHBAR_ChartDur * 7 Then endOffset = PHBAR_ChartDur * 7 - 1
        End If
        
        ' 리소스 표시하기
        If PHBAR_ChartType = "Mon" Then  ' 월별 리소스 표시
          formStartDate = DateSerial(Year(formStartDate), Month(formStartDate), 1)
          
          For iMon = stOffset To endOffset '월별 리소스 계산
            monVal = 0
            monStDate = DateAdd("m", iMon, formStartDate) ' 당월 첫날
            monEndDate = DateAdd("m", iMon + 1, formStartDate) - 1 ' 당월 마지막날
            If monStDate < stDate Then monStDate = stDate
            If monEndDate > endDate Then monEndDate = endDate

            If PHBAR_HolidayType <> "7" Then
              For iDay = monStDate To monEndDate
                If Weekday(iDay, vbMonday) < chkWeekday Then monVal = monVal + iRscDist
              Next
            Else
              For iDay = monStDate To monEndDate
                monVal = monVal + iRscDist
              Next
            End If
              
            sh.Cells(i, PHBAR_COL_BarLeft + iMon).Value = monVal
            
          Next
        Else                             ' 일별 리소스 표시
            If PHBAR_HolidayType <> "7" Then
              For iDay = stOffset To endOffset
                If Weekday(formStartDate + iDay, vbMonday) < chkWeekday Then sh.Cells(i, PHBAR_COL_BarLeft + iDay).Value = iRscDist
              Next
            Else
              For iDay = stOffset To endOffset
                sh.Cells(i, PHBAR_COL_BarLeft + iDay).Value = iRscDist
              Next
            End If
        End If
      End If
        
      i = i + 1
    Loop
       
exit_bar:
    Cells(1, 1).Select
    Exit Sub
err_bar:
    Application.ScreenUpdating = True
    
    MsgBox "Error : Please check data at Line No =" & CStr(i) & Chr(10) & Err.Description
    On Error GoTo 0
    Resume exit_bar
End Sub

' 리소스 삭제
Sub rsc_clear()
    Dim iColOffset As Long
    Dim sh As Excel.Worksheet
    
    configLoad
    Set sh = Application.ActiveSheet
    
    On Error Resume Next
    
    If PHBAR_ChartType = "Mon" Then
      iColOffset = PHBAR_COL_BarLeft + PHBAR_ChartDur - 1
    Else
      iColOffset = PHBAR_COL_BarLeft + PHBAR_ChartDur * 7 - 1
    End If
    
    If Not IsNumeric(iColOffset) Or iColOffset > 16300 Then iColOffset = 16300

    sh.Range(sh.Cells(PHBAR_ROW_DataTop, PHBAR_COL_BarLeft), sh.Cells(PHBAR_ROW_DataTop + PHBAR_ActCnt - 1, iColOffset)).ClearContents

End Sub
