VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formConfig 
   Caption         =   "Config"
   ClientHeight    =   6936
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   OleObjectBlob   =   "formConfig.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "formConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim activeText As String

Private Sub btnColorActActual_Click()
  txtColorActActual.BackColor = PickNewColor(txtColorActActual.BackColor)
End Sub

Private Sub btnColorActPlan_Click()
  txtColorActPlan.BackColor = PickNewColor(txtColorActPlan.BackColor)
End Sub

Private Sub btnColorGroupActual_Click()
  txtColorGroupActual.BackColor = PickNewColor(txtColorGroupActual.BackColor)
End Sub

Private Sub btnColorGroupPlan_Click()
  txtColorGroupPlan.BackColor = PickNewColor(txtColorGroupPlan.BackColor)
End Sub

Private Sub btnColorMSActual_Click()
  txtColorMSActual.BackColor = PickNewColor(txtColorMSActual.BackColor)
End Sub

Private Sub btnColorMSPlan_Click()
  txtColorMSPlan.BackColor = PickNewColor(txtColorMSPlan.BackColor)
End Sub

Private Sub btnDefault_Click()
  optWeek.Value = True
  txtChartTypeP.Text = "Weekly"
  lblUnit.Caption = "Weeks"
  
  optWork6.Value = True
  
  txtChartDur.Text = txtChartDurP.Text
  txtActCnt.Text = txtActCntP.Text
  
  '색상
  txtColorMSPlan.BackColor = C_COLOR_MSPLAN
  txtColorMSActual.BackColor = C_COLOR_MSACTUAL
  txtColorGroupPlan.BackColor = C_COLOR_GROUPPLAN
  txtColorGroupActual.BackColor = C_COLOR_GROUPACTUAL
  txtColorActPlan.BackColor = C_COLOR_ACTPLAN
  txtColorActActual.BackColor = C_COLOR_ACTACTUAL

  ' Columns / Rows
  txtColActID.Text = C_COL_ActID
  txtColActDesc.Text = C_COL_ActDesc
  txtColActType.Text = C_COL_ActType
  txtColPlanSt.Text = C_COL_PLANST
  txtColActSt.Text = C_COL_ActST
  txtColBarLeft.Text = C_COL_BarLeft
  txtRowTitleTop.Text = C_ROW_TitleTop
  txtRowDataTop.Text = C_ROW_DataTop
  txtColProgress.Text = C_COL_Progress
  txtColDifferance.Text = C_COL_Difference
  txtColResourceP.Text = C_COL_Resource

  chkUseActual.Value = True
  chkUseDiffence.Value = True
  chkUseResource.Value = False

  txtColResource.Enabled = PHBAR_USEActual
  txtColDifferance.Enabled = PHBAR_USEDifference
  txtColResource.Enabled = PHBAR_USEResource
  
End Sub

Private Sub chkUseActual_Change()
  txtColActSt.Enabled = chkUseActual.Value
End Sub

Private Sub chkUseDiffence_Change()
  txtColDifferance.Enabled = chkUseDiffence.Value
End Sub

Private Sub chkUseResource_Change()
  txtColResource.Enabled = chkUseResource.Value
End Sub


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


Private Sub MultiPage1_Change()
  activeText = ""
  If MultiPage1.Value = 1 Then
    activeText = "txtColActID"
  ElseIf MultiPage1.Value = 2 Then
    activeText = "txtRowTitleTop"
  End If
End Sub

Private Sub txtColActDur_Enter()
  activeText = txtColActDur.Name
End Sub

Private Sub txtColActEnd_Enter()
  activeText = txtColActEnd.Name
End Sub

Private Sub txtColActID_Enter()
  activeText = txtColActID.Name
End Sub

Private Sub txtColActDesc_Enter()
  activeText = txtColActDesc.Name
End Sub

Private Sub txtColActType_Enter()
  activeText = txtColActType.Name
End Sub

Private Sub txtColPlanDur_Enter()
  activeText = txtColPlanDur.Name
End Sub

Private Sub txtColPlanEnd_Enter()
  activeText = txtColPlanEnd.Name
End Sub

Private Sub txtColPlanSt_Enter()
  activeText = txtColPlanSt.Name
End Sub

Private Sub txtColActSt_Enter()
  activeText = txtColActSt.Name
End Sub

Private Sub txtColBarLeft_Enter()
  activeText = txtColBarLeft.Name
End Sub

Private Sub txtColProgress_Enter()
  activeText = txtColProgress.Name
End Sub

Private Sub txtColDifferance_Enter()
  activeText = txtColDifferance.Name
End Sub

Private Sub txtColResource_Enter()
  activeText = txtColResource.Name
End Sub

Private Sub txtRowTitleTop_Enter()
  activeText = txtRowTitleTop.Name
End Sub

Private Sub txtRowDataTop_Enter()
  activeText = txtRowDataTop.Name
End Sub

Public Sub setVals(row As Long, col As Long)
  If Not (Me.Visible) Then Exit Sub
  
  Select Case activeText
    Case "txtColActID"
        txtColActID.Text = col
    Case "txtColActDesc"
        txtColActDesc.Text = col
    Case "txtColActType"
        txtColActType.Text = col
        
    Case "txtColPlanSt"
        txtColPlanSt.Text = col
    Case "txtColPlanEnd"
        txtColPlanEnd.Text = col
    Case "txtColPlanDur"
        txtColPlanDur.Text = col
    Case "txtColResource"
        txtColResource.Text = col
    
    Case "txtColActSt"
        txtColActSt.Text = col
    Case "txtColActEnd"
        txtColActEnd.Text = col
    Case "txtColActDur"
        txtColActDur.Text = col
        
    Case "txtColBarLeft"
        txtColBarLeft.Text = col
    Case "txtColProgress"
        txtColProgress.Text = col
    Case "txtColDifferance"
        txtColDifferance.Text = col
        
    Case "txtRowTitleTop"
        txtRowTitleTop.Text = row
    Case "txtRowDataTop"
        txtRowDataTop.Text = row
  End Select
End Sub


Private Sub UserForm_Activate()
  Dim chartType, holidayType
  Set ThisWorkbook.chartSheet = ActiveSheet
  activeText = 0
  
  MultiPage1.Value = 0

  configLoad
    
  chartType = PHBAR_ChartType
  If chartType = "Day" Then
    optDay.Value = True
    txtChartTypeP.Text = "Daily"
    lblUnit.Caption = "Weeks"
  ElseIf chartType = "Mon" Then
    optMonth.Value = True
    txtChartTypeP.Text = "Monthly"
    lblUnit.Caption = "Months"
  Else
    optWeek.Value = True
    txtChartTypeP.Text = "Weekly"
    lblUnit.Caption = "Weeks"
  End If
  
  If PHBAR_HolidayType = 7 Then
    optWork7.Value = True
  ElseIf PHBAR_HolidayType = 5 Then
    optWork5.Value = True
  Else
    optWork6.Value = True
  End If

  txtChartDurP.Text = PHBAR_ChartDur
  txtActCntP.Text = PHBAR_ActCnt
  txtChartDur.Text = txtChartDurP.Text
  txtActCnt.Text = txtActCntP.Text
  
  '색상
  txtColorMSPlan.BackColor = COLOR_MSPLAN
  txtColorMSActual.BackColor = COLOR_MSACTUAL
  txtColorGroupPlan.BackColor = COLOR_GROUPPLAN
  txtColorGroupActual.BackColor = COLOR_GROUPACTUAL
  txtColorActPlan.BackColor = COLOR_ACTPLAN
  txtColorActActual.BackColor = COLOR_ACTACTUAL

  ' Columns / Rows
  txtColActIDP.Text = PHBAR_COL_ActID
  txtColActDescP.Text = PHBAR_COL_ActDesc
  txtColActTypeP.Text = PHBAR_COL_ActType
  txtColPlanStP.Text = PHBAR_COL_PLANST
  txtColPlanEndP.Text = PHBAR_COL_PlanEnd
  txtColPlanDurP.Text = PHBAR_COL_PlanDur
  txtColResourceP.Text = PHBAR_COL_Resource
  
  txtColActStP.Text = PHBAR_COL_ActST
  txtColActEndP.Text = PHBAR_COL_ActEnd
  txtColActDurP.Text = PHBAR_COL_ActDur
  
  txtColProgressP.Text = PHBAR_COL_Progress
  txtColDifferanceP.Text = PHBAR_COL_Difference
  
  txtColBarLeftP.Text = PHBAR_COL_BarLeft
  
  txtRowTitleTopP.Text = PHBAR_ROW_TitleTop
  txtRowDataTopP.Text = PHBAR_ROW_DataTop
  
  txtColActID.Text = txtColActIDP.Text
  txtColActDesc.Text = txtColActDescP.Text
  txtColActType.Text = txtColActTypeP.Text
  
  txtColPlanSt.Text = txtColPlanStP.Text
  txtColPlanEnd.Text = txtColPlanEndP.Text
  txtColPlanDur.Text = txtColPlanDurP.Text
  txtColResource.Text = txtColResourceP.Text
  
  txtColActSt.Text = txtColActStP.Text
  txtColActEnd.Text = txtColActEndP.Text
  txtColActDur.Text = txtColActDurP.Text
  
  txtColBarLeft.Text = txtColBarLeftP.Text
  txtColProgress.Text = txtColProgressP.Text
  txtColDifferance.Text = txtColDifferanceP.Text
  
  txtRowTitleTop.Text = txtRowTitleTopP.Text
  txtRowDataTop.Text = txtRowDataTopP.Text

  chkUseActual.Value = PHBAR_USEActual
  chkUseDiffence.Value = PHBAR_USEDifference
  chkUseResource.Value = PHBAR_USEResource

  txtColResource.Enabled = PHBAR_USEActual
  txtColDifferance.Enabled = PHBAR_USEDifference
  txtColResource.Enabled = PHBAR_USEResource
  
End Sub


Private Sub btnOK_Click()
  Dim chartType As String
  Dim holidayType As String
  
  setVersion
  
  ' TAB1 - General =============================
  ' Barchar Type
  If optDay.Value = "True" Then
    chartType = "Day"
  ElseIf optMonth.Value = "True" Then
    chartType = "Mon"
  Else
    chartType = "Week"
  End If
  set_Property "PHBAR_ChartType", chartType
  
  If optWork7.Value = "True" Then
    holidayType = "7"
  ElseIf optWork5.Value = "True" Then
    holidayType = "5"
  Else
    holidayType = "6"
  End If
  set_Property "PHBAR_HolidayType", holidayType
  
  ' Data Limit
  set_Property "PHBAR_ActCnt", txtActCnt.Text
  set_Property "PHBAR_ChartDur", txtChartDur.Text
  
  '색상
  set_Property "PHBAR_COLOR_MSPLAN", CStr(txtColorMSPlan.BackColor)
  set_Property "PHBAR_COLOR_MSACTUAL", CStr(txtColorMSActual.BackColor)
  set_Property "PHBAR_COLOR_GROUPPLAN", CStr(txtColorGroupPlan.BackColor)
  set_Property "PHBAR_COLOR_GROUPACTUAL", CStr(txtColorGroupActual.BackColor)
  set_Property "PHBAR_COLOR_ACTPLAN", CStr(txtColorActPlan.BackColor)
  set_Property "PHBAR_COLOR_ACTACTUAL", CStr(txtColorActActual.BackColor)
  
  
  ' TAB2 - Columns =============================
  set_Property "PHBAR_COL_ActID", txtColActID.Text
  set_Property "PHBAR_COL_ActDesc", txtColActDesc.Text
  set_Property "PHBAR_COL_ActType", txtColActType.Text
  
  set_Property "PHBAR_COL_PLANST", txtColPlanSt.Text
  set_Property "PHBAR_COL_PLANEND", txtColPlanEnd.Text
  set_Property "PHBAR_COL_PLANDUR", txtColPlanDur.Text
  
  set_Property "PHBAR_COL_ActST", txtColActSt.Text
  set_Property "PHBAR_COL_ActEND", txtColActEnd.Text
  set_Property "PHBAR_COL_ActDUR", txtColActDur.Text
  
  If chkUseActual.Value Then
    set_Property "PHBAR_USEActual", "1"
  Else
    set_Property "PHBAR_USEActual", "0"
  End If
  
  If chkUseDiffence.Value Then
    set_Property "PHBAR_USEDifference", "1"
  Else
    set_Property "PHBAR_USEDifference", "0"
  End If
  
  If chkUseResource.Value Then
    set_Property "PHBAR_USEResource", "1"
  Else
    set_Property "PHBAR_USEResource", "0"
  End If
  set_Property "PHBAR_COL_Resource", txtColResource.Text
    
  set_Property "PHBAR_COL_Progress", txtColProgress.Text
  set_Property "PHBAR_COL_Difference", txtColDifferance.Text
  
  set_Property "PHBAR_COL_BarLeft", txtColBarLeft.Text
  
  ' TAB3 - Rows =============================
  set_Property "PHBAR_ROW_TitleTop", txtRowTitleTop.Text
  set_Property "PHBAR_ROW_DataTop", txtRowDataTop.Text
    
  configLoad
  
  Me.Hide
End Sub

Private Sub btnCancel_Click()
  Me.Hide
End Sub

