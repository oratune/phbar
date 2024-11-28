Attribute VB_Name = "callback"
Option Explicit

'===================================================
Sub phNewChart(ByVal control As IRibbonControl)
  Call http_CheckServer
  Call makeNewChartSheet
End Sub

Sub phFormRedraw(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  Call formRedraw
End Sub

Sub phFormSetting(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  Set ThisWorkbook.chartSheet = ActiveSheet
  formConfig.Show 0
  
  Call http_CheckServer
End Sub

'===================================================
Sub phFuncAF(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  If nvl(get_Property("PHBAR_USEActual"), "1") = "0" Then
    MsgBox "Fibidden!" & Chr(10) & "Check Form Setting!"
    Exit Sub
  End If
  Call formula_actualFinish
End Sub

Sub phFuncDur(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  Call formula_duration
End Sub

Sub phFuncDiff(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  If nvl(get_Property("PHBAR_USEActual"), "1") = "0" Then
    MsgBox "Fibidden!" & Chr(10) & "Check Form Setting!"
    Exit Sub
  End If
  Call formula_difference
End Sub


'===================================================
Sub phCalcAF(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  If nvl(get_Property("PHBAR_USEActual"), "1") = "0" Then
    MsgBox "Fibidden!" & Chr(10) & "Check Form Setting!"
    Exit Sub
  End If
  
  Call calc_actualFinish

End Sub

Sub phCalcDur(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  Call calc_duration
End Sub

Sub phCalcDiff(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  If nvl(get_Property("PHBAR_USEDifference"), "1") = "0" Then
    MsgBox "Fibidden!" & Chr(10) & "Check Form Setting!"
    Exit Sub
  End If
  Call calc_difference
End Sub


'===================================================
Sub phGroupRange(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  MsgBox "Under Developing"

End Sub

Sub phGroupAll(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  MsgBox "Under Developing"

End Sub


'===================================================
Sub phDrawFull(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  Call bar_Refresh_full
  Call http_CheckServer
End Sub

Sub phDrawRange(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  Call bar_Refresh_range
End Sub

Sub phDrawClear(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  Call bar_clearChart
End Sub

Sub phDrawClearAll(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  Call bar_clearAll
End Sub

'===================================================
Sub phRscFull(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  If nvl(get_Property("PHBAR_USEResource"), "0") = "0" Then
    MsgBox "Fibidden!" & Chr(10) & "Check Form Setting!"
    Exit Sub
  End If
  
  Call rsc_distrib_full
End Sub

Sub phRscRange(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  If nvl(get_Property("PHBAR_USEResource"), "0") = "0" Then
    MsgBox "Fibidden!" & Chr(10) & "Check Form Setting!"
    Exit Sub
  End If

  Call rsc_distrib_range

End Sub

Sub phRscClear(ByVal control As IRibbonControl)
  If Not checkPhBarMsg Then Exit Sub
  If nvl(get_Property("PHBAR_USEResource"), "0") = "0" Then
    MsgBox "Fibidden!" & Chr(10) & "Check Form Setting!"
    Exit Sub
  End If
   
  Call rsc_clear
End Sub

'===================================================
Sub phAbout(ByVal control As IRibbonControl)
  formAbout.Show 1
End Sub
