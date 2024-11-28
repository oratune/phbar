Attribute VB_Name = "Calc"
Option Explicit

' Duration 자동 계산- 수식 입력
Sub formula_duration()
    On Error GoTo err_bar
    Dim i, stDate, endDate, blankCnt
    Dim sh As Excel.Worksheet
    
    Application.ScreenUpdating = False

    configLoad
    Set sh = Application.ActiveSheet
    
    i = PHBAR_ROW_DataTop
    blankCnt = 0

    Do Until blankCnt > 5
    
      If PHBAR_ActCnt <= i - PHBAR_ROW_DataTop Then Exit Do

      If checkRowBlank(sh, i) Then
        blankCnt = blankCnt + 1
        
      Else
        sh.Range(sh.Cells(i, PHBAR_COL_PlanDur), sh.Cells(i, PHBAR_COL_PlanDur)).Formula = _
                                       "=if(and(rc[-1]<>"""",rc[-2]<>""""),RC[-1]-RC[-2]+1,0)"
        
        If PHBAR_USEActual Then
          sh.Range(sh.Cells(i, PHBAR_COL_ActDur), sh.Cells(i, PHBAR_COL_ActDur)).Formula = _
                                       "=if(and(rc[-1]<>"""",rc[-2]<>""""),RC[-1]-RC[-2]+1,0)"
        End If
      End If
        
      i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
exit_bar:
    Cells(1, 1).Select
    Exit Sub
err_bar:
    Application.ScreenUpdating = True

    MsgBox "Error : Please check data at Line No =" & CStr(i) & Chr(10) & Err.Description
    On Error GoTo 0
    Resume exit_bar
End Sub

' Duration 자동 계산- 휴무일 반영
Sub calc_duration()
    On Error GoTo err_bar
    Dim i, chkWeekday, stDate, endDate, aDay, blankCnt
    Dim iDur As Long
    Dim sh As Excel.Worksheet
    
    Application.ScreenUpdating = False

    configLoad
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
    
    Do Until blankCnt > 5
    
      If PHBAR_ActCnt <= i - PHBAR_ROW_DataTop Then Exit Do

      If checkRowBlank(sh, i) Then
        blankCnt = blankCnt + 1
      Else
        ' 계획 기간 계산
        stDate = sh.Cells(i, PHBAR_COL_PLANST).Value
        endDate = sh.Cells(i, PHBAR_COL_PLANST + 1).Value
        iDur = 0
        
        If stDate = "" Or endDate = "" Or IsEmpty(stDate) Or IsEmpty(endDate) Or endDate < stDate Then
          iDur = 0
        ElseIf PHBAR_HolidayType <> "7" Then
          For aDay = stDate To endDate
            If Weekday(aDay, vbMonday) < chkWeekday Then iDur = iDur + 1
          Next aDay
        Else
          iDur = endDate - stDate + 1
        End If
        
        sh.Cells(i, PHBAR_COL_PLANST + 2).Value = iDur
        
        ' 실적 기간 계산
        If PHBAR_USEActual Then
          stDate = sh.Cells(i, PHBAR_COL_ActST).Value
          endDate = sh.Cells(i, PHBAR_COL_ActST + 1).Value
          iDur = 0
          
          If stDate = "" Or endDate = "" Or IsEmpty(stDate) Or IsEmpty(endDate) Or endDate < stDate Then
            iDur = 0
          ElseIf PHBAR_HolidayType <> "7" Then
            For aDay = stDate To endDate
              If Weekday(aDay, vbMonday) < chkWeekday Then iDur = iDur + 1
            Next aDay
          Else
            iDur = endDate - stDate + 1
          End If
          sh.Cells(i, PHBAR_COL_ActST + 2).Value = iDur
        End If
      End If
        
      i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
exit_bar:
    Cells(1, 1).Select
    Exit Sub
err_bar:
    Application.ScreenUpdating = True
    
MsgBox "Error : Please check data at Line No =" & CStr(i) & Chr(10) & Err.Description
    On Error GoTo 0
    Resume exit_bar
  
  
End Sub

Sub formula_actualFinish()
    On Error GoTo err_bar
    Dim i, stDate, endDate, AFDate, blankCnt
    Dim c1, c2, c3, f
    Dim sh As Excel.Worksheet
    
    Application.ScreenUpdating = False

    configLoad
    Set sh = Application.ActiveSheet
    
    i = PHBAR_ROW_DataTop
    blankCnt = 0

    
    c1 = PHBAR_COL_ActST - PHBAR_COL_ActEnd   'AS
    c2 = PHBAR_COL_PLANST - PHBAR_COL_ActEnd    'plan start
    c3 = PHBAR_COL_PLANST + 1 - PHBAR_COL_ActEnd   'plan end

    
    Do Until blankCnt > 5
    
      If PHBAR_ActCnt <= i - PHBAR_ROW_DataTop Then Exit Do

      stDate = sh.Cells(i, PHBAR_COL_PLANST).Value
      endDate = sh.Cells(i, PHBAR_COL_PLANST + 1).Value
      AFDate = sh.Cells(i, PHBAR_COL_ActST + 1).Value
      
      If checkRowBlank(sh, i) Then
        blankCnt = blankCnt + 1
      ElseIf AFDate = "" Or IsEmpty(AFDate) Then
        f = "=if(RC[" & c1 & "]<> """",   RC[" & c1 & "] + RC[" & c3 & "] - RC[" & c2 & "] ,"""")"
        sh.Range(sh.Cells(i, PHBAR_COL_ActEnd), sh.Cells(i, PHBAR_COL_ActEnd)).Formula = f
      End If
        
      i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
exit_bar:
    Cells(1, 1).Select
    Exit Sub
err_bar:
    Application.ScreenUpdating = True
    MsgBox "Error : Please check data at Line No =" & CStr(i) & Chr(10) & Err.Description
    On Error GoTo 0
    Resume exit_bar

End Sub

' 완료예정일 자동 계산- 휴무일 반영
Sub calc_actualFinish()
    On Error GoTo err_bar
    Dim i, iDay, chkWeekday, stDate, plandur, endDate, AFDate, blankCnt
    Dim sh As Excel.Worksheet
    
    Application.ScreenUpdating = False

    configLoad
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
    
    Do Until blankCnt > 5
    
      If PHBAR_ActCnt <= i - PHBAR_ROW_DataTop Then Exit Do

      If checkRowBlank(sh, i) Then
        blankCnt = blankCnt + 1
      Else
        stDate = sh.Cells(i, PHBAR_COL_ActST).Value
        endDate = sh.Cells(i, PHBAR_COL_ActST + 1).Value
        plandur = sh.Cells(i, PHBAR_COL_PLANST + 2).Value
        
        If endDate = "" Or IsEmpty(endDate) Then
          If stDate <> "" And plandur <> "" And Not IsEmpty(stDate) And Not IsEmpty(plandur) Then
            If PHBAR_HolidayType <> "7" Then
              iDay = 1
              AFDate = stDate - 1
              Do
                AFDate = AFDate + 1
                If Weekday(AFDate, vbMonday) < chkWeekday Then iDay = iDay + 1
              Loop Until iDay > plandur
            Else
              AFDate = stDate
            End If
            sh.Cells(i, PHBAR_COL_ActST + 1).Value = AFDate
          End If
                    
        End If
      
      End If
        
      i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
exit_bar:
    Cells(1, 1).Select
    Exit Sub
err_bar:
    Application.ScreenUpdating = True
    
    MsgBox "Error : Please check data at Line No =" & CStr(i) & Chr(10) & Err.Description
    On Error GoTo 0
    Resume exit_bar
  
  
End Sub

Sub formula_difference()
    On Error GoTo err_bar
    Dim i, stDate, endDate, blankCnt
    Dim c1, c2, c3
    Dim sh As Excel.Worksheet
    
    Application.ScreenUpdating = False

    configLoad
    Set sh = Application.ActiveSheet
    
    i = PHBAR_ROW_DataTop
    blankCnt = 0
    
    c1 = PHBAR_COL_PlanEnd - PHBAR_COL_Difference 'EF
    c2 = PHBAR_COL_ActEnd - PHBAR_COL_Difference  'AF
    

    Do Until blankCnt > 5
    
      If PHBAR_ActCnt <= i - PHBAR_ROW_DataTop Then Exit Do

      If checkRowBlank(sh, i) Then
        blankCnt = blankCnt + 1
        
      Else
        sh.Range(sh.Cells(i, PHBAR_COL_Difference), sh.Cells(i, PHBAR_COL_Difference)).Formula = _
            "=if(and(RC[" & c1 & "] <> """", RC[" & c2 & "] <> """"),RC[" & c1 & "] -RC[" & c2 & "] ,0)"
      End If
        
      i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
exit_bar:
    Cells(1, 1).Select
    Exit Sub
err_bar:
    Application.ScreenUpdating = True
    
    MsgBox "Error : Please check data at Line No =" & CStr(i) & Chr(10) & Err.Description
    On Error GoTo 0
    Resume exit_bar

End Sub


' 공정차이 계산- 휴무일 반영
Sub calc_difference()
    On Error GoTo err_bar
    Dim i, chkWeekday, plandate, actualdate, aDay, blankCnt
    Dim iDiff As Long
    Dim sh As Excel.Worksheet
    
    Application.ScreenUpdating = False

    configLoad
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
    
    Do Until blankCnt > 5
    
      If PHBAR_ActCnt <= i - PHBAR_ROW_DataTop Then Exit Do

      If checkRowBlank(sh, i) Then
        blankCnt = blankCnt + 1
      Else
        
        plandate = sh.Cells(i, PHBAR_COL_PLANST + 1).Value
        actualdate = sh.Cells(i, PHBAR_COL_ActST + 1).Value
        
        iDiff = 0
        
        If plandate = "" Or actualdate = "" Or Not IsDate(plandate) Or Not IsDate(actualdate) Then
          iDiff = 0
        ElseIf PHBAR_HolidayType <> "7" Then
          If CDate(plandate) < CDate(actualdate) Then
            For aDay = plandate To actualdate
              If Weekday(aDay, vbMonday) < chkWeekday Then iDiff = iDiff + 1
            Next aDay
          Else
            For aDay = actualdate To plandate
              If Weekday(aDay, vbMonday) < chkWeekday Then iDiff = iDiff - 1
            Next aDay
          End If
        Else
          iDiff = plandate - actualdate
        End If
        
        sh.Cells(i, PHBAR_COL_Difference).Value = iDiff
        
      End If
        
      i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    
exit_bar:
    Cells(1, 1).Select
    Exit Sub
err_bar:
    Application.ScreenUpdating = True
    
    MsgBox "Error : Please check data at Line No =" & CStr(i) & Chr(10) & Err.Description
    On Error GoTo 0
    Resume exit_bar
End Sub

