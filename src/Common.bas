Attribute VB_Name = "Common"
Option Explicit

'================================================
Global Const C_Ver_No = "7.21"
Global Const C_Ver = "7.21"
Global Const C_VerDate = "2016.11.06"
'================================================

Global Const C_LIMIT_COL = 16300
Global Const C_HEIGHT_OF_ROW = 21

Public Height_of_Row As Long

Global Const C_ROW_TitleTop = 4
Global Const C_ROW_DataTop = 6

Global Const C_COL_ActID = 1
Global Const C_COL_ActDesc = 2
Global Const C_COL_ActType = 3
Global Const C_COL_PLANST = 4
Global Const C_COL_PLANEND = 5
Global Const C_COL_PLANDUR = 6
Global Const C_COL_Resource = 11
Global Const C_COL_ActST = 7
Global Const C_COL_ActEND = 8
Global Const C_COL_ActDUR = 9
Global Const C_COL_Progress = 10
Global Const C_COL_Difference = 11
Global Const C_COL_BarLeft = 12

Public PHBAR_ChartType As String
Public PHBAR_HolidayType As String
Public PHBAR_ActCnt As Long
Public PHBAR_ChartDur As Long

Public PHBAR_ROW_TitleTop As Long
Public PHBAR_ROW_DataTop As Long

Public PHBAR_COL_ActID As Long
Public PHBAR_COL_ActDesc As Long
Public PHBAR_COL_ActType As Long
Public PHBAR_COL_PLANST As Long     ' Start
Public PHBAR_COL_PlanEnd As Long     ' Finish
Public PHBAR_COL_PlanDur As Long     ' Duration
Public PHBAR_COL_ActST As Long     ' Start
Public PHBAR_COL_ActEnd As Long     ' Finish
Public PHBAR_COL_ActDur As Long     ' Duration
Public PHBAR_COL_Progress As Long     ' 진도율
Public PHBAR_COL_Difference As Long     ' 차이
Public PHBAR_COL_Resource As Long     ' 차이

Public PHBAR_COL_BarLeft As Long



Public PHBAR_USEActual As Boolean
Public PHBAR_USEDifference As Boolean
Public PHBAR_USEResource As Boolean

Global Const C_COLOR_MSPLAN = 10027008
Global Const C_COLOR_MSACTUAL = 222
Global Const C_COLOR_GROUPPLAN = 10027008
Global Const C_COLOR_GROUPACTUAL = 222
Global Const C_COLOR_ACTPLAN = 14070636
Global Const C_COLOR_ACTACTUAL = 11318000

Public COLOR_MSPLAN As Double
Public COLOR_MSACTUAL As Double
Public COLOR_GROUPPLAN As Double
Public COLOR_GROUPACTUAL As Double
Public COLOR_ACTPLAN As Double
Public COLOR_ACTACTUAL As Double


Sub http_CheckServer(Optional bMsg As Boolean = False)
   Dim objHTTP
   Dim url, re
   Dim verNo, verName, verUrl
   Dim browser
   
   On Error GoTo errrtn
   
   ' HTTP Call
   Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
   url = "http://leepong.cafe24.com/phbar/phbar.php"
   objHTTP.Open "HEAD", url, False
   objHTTP.setRequestHeader "User-Agent", "PhBarchart"
   objHTTP.send ("")
  
   'MsgBox objHTTP.getAllResponseHeaders()
     
   ' 버전 번호 비교
   verNo = objHTTP.getResponseHeader("phbar_ver")
   
   If Not IsNumeric(verNo) Then Exit Sub
   If CDbl(verNo) <= CDbl(C_Ver_No) Then
     If bMsg Then MsgBox "You are using latest version of PhBarchart."
     Exit Sub
   End If
    
   verName = objHTTP.getResponseHeader("phbar_vernm")
    
   re = MsgBox("There is new version of PhBarchart" & Chr(10) & verName & Chr(10) & _
               "Do you want to move to Upate Site ?", _
          vbYesNo, _
          "Phbarchart Update")
   If re <> vbYes Then Exit Sub
      
   ' 업데이트 사이트 이동
   verUrl = objHTTP.getResponseHeader("phbar_verurl")
   Set browser = CreateObject("InternetExplorer.Application")
   browser.Navigate (verUrl)
   browser.StatusBar = True
   browser.Toolbar = True
   browser.Visible = True
   browser.Resizable = True
   browser.AddressBar = True
   
   Exit Sub
errrtn:
  If bMsg Then MsgBox "BarChart Version check Error-" & Err.Description
End Sub


Function checkPhBarMsg() As Boolean
    If getVersion() = "" Then
      MsgBox "This Worksheet is not a [PHBarchart]"
      checkPhBarMsg = False
    Else
      checkPhBarMsg = True
    End If
End Function

Sub configLoad()
  ' General
  PHBAR_ChartType = nvl(get_Property("PHBAR_ChartType"), "week")
  PHBAR_HolidayType = nvl(get_Property("PHBAR_HolidayType"), "6")
  PHBAR_ChartDur = nvl(get_Property("PHBAR_ChartDur"), 0)
  PHBAR_ActCnt = nvl(get_Property("PHBAR_ActCnt"), 500)
  
  ' 색상 =====================
  COLOR_MSPLAN = nvl(get_Property("PHBAR_COLOR_MSPLAN"), C_COLOR_MSPLAN)
  COLOR_MSACTUAL = nvl(get_Property("PHBAR_COLOR_MSACTUAL"), C_COLOR_MSACTUAL)
  COLOR_GROUPPLAN = nvl(get_Property("PHBAR_COLOR_GROUPPLAN"), C_COLOR_GROUPPLAN)
  COLOR_GROUPACTUAL = nvl(get_Property("PHBAR_COLOR_GROUPACTUAL"), C_COLOR_GROUPACTUAL)
  COLOR_ACTPLAN = nvl(get_Property("PHBAR_COLOR_ACTPLAN"), C_COLOR_ACTPLAN)
  COLOR_ACTACTUAL = nvl(get_Property("PHBAR_COLOR_ACTACTUAL"), C_COLOR_ACTACTUAL)
  
  ' Column ==================
  PHBAR_COL_ActID = nvl(get_Property("PHBAR_COL_ActID"), C_COL_ActID)
  PHBAR_COL_ActDesc = nvl(get_Property("PHBAR_COL_ActDesc"), C_COL_ActDesc)
  PHBAR_COL_ActType = nvl(get_Property("PHBAR_COL_ActType"), C_COL_ActType)
  PHBAR_COL_PLANST = nvl(get_Property("PHBAR_COL_PLANST"), C_COL_PLANST)
  PHBAR_COL_PlanEnd = nvl(get_Property("PHBAR_COL_PLANEND"), C_COL_PLANEND)
  PHBAR_COL_PlanDur = nvl(get_Property("PHBAR_COL_PLANDUR"), C_COL_PLANDUR)
  
  PHBAR_COL_ActST = nvl(get_Property("PHBAR_COL_ActST"), C_COL_ActST)
  PHBAR_COL_ActEnd = nvl(get_Property("PHBAR_COL_ActEND"), C_COL_ActEND)
  PHBAR_COL_ActDur = nvl(get_Property("PHBAR_COL_ActDUR"), C_COL_ActDUR)
  
  PHBAR_COL_Progress = nvl(get_Property("PHBAR_COL_Progress"), C_COL_Progress)
  PHBAR_COL_Difference = nvl(get_Property("PHBAR_COL_Difference"), C_COL_Difference)
  PHBAR_COL_Resource = nvl(get_Property("PHBAR_COL_Resource"), C_COL_Resource)
  PHBAR_COL_BarLeft = nvl(get_Property("PHBAR_COL_BarLeft"), C_COL_BarLeft)
  
  If nvl(get_Property("PHBAR_USEActual"), "1") = "1" Then
    PHBAR_USEActual = True
  Else
    PHBAR_USEActual = False
  End If
  If nvl(get_Property("PHBAR_USEDifference"), "1") = "1" Then
    PHBAR_USEDifference = True
  Else
    PHBAR_USEDifference = False
  End If
  If nvl(get_Property("PHBAR_USEResource"), "0") = "1" Then
    PHBAR_USEResource = True
  Else
    PHBAR_USEResource = False
  End If
  
  ' Row =====================
  PHBAR_ROW_TitleTop = nvl(get_Property("PHBAR_ROW_TitleTop"), C_ROW_TitleTop)
  PHBAR_ROW_DataTop = nvl(get_Property("PHBAR_ROW_DataTop"), C_ROW_DataTop)
End Sub


Sub set_Property(pname As String, val As String)
  On Error GoTo errrtn
  Dim prop As CustomProperty
  Dim sh As Worksheet
  
  Set sh = ActiveSheet
  
  If Not get_PropertyExists(pname) Then
    sh.CustomProperties.Add pname, val
    Exit Sub
  End If
  
  For Each prop In sh.CustomProperties
    If prop.Name = pname Then
      prop.Value = val
      Exit For
    End If
  Next
errrtn:
  If Err Then MsgBox "BarChart Set Property Error-" & Err.Description
End Sub

Function get_PropertyExists(pname As String) As Boolean
  Dim prop As CustomProperty
  
  get_PropertyExists = False
  
  For Each prop In ActiveSheet.CustomProperties
    If prop.Name = pname Then
      get_PropertyExists = True
      Exit For
    End If
  Next
  Exit Function
End Function

Function get_Property(pname As String) As Variant
  On Error GoTo errrtn
  Dim prop As CustomProperty
  
  get_Property = ""
  
  For Each prop In ActiveSheet.CustomProperties
    If prop.Name = pname Then
      get_Property = prop.Value
      Exit For
    End If
  Next
  Exit Function
errrtn:
  get_Property = "0"
End Function

Function nvl(val As Variant, newVal As Variant) As Variant
  If IsNull(val) Then
    nvl = newVal
  ElseIf val = "" Then
    nvl = newVal
  Else
    nvl = val
  End If
End Function

Sub setVersion()
  set_Property "PHBar_Version", C_Ver
End Sub

Function getVersion() As Variant
  getVersion = get_Property("PHBar_Version")
End Function

' 색상 선택
Function PickNewColor(Optional i_OldColor As Double = xlNone) As Double
  Const BGColor As Long = 13160660  'background color of dialogue
  Const ColorIndexLast As Long = 32 'index of last custom color in palette

  Dim myOrgColor As Double          'original color of color index 32
  Dim myNewColor As Double          'color that was picked in the dialogue
  Dim myRGB_R As Integer            'RGB values of the color that will be
  Dim myRGB_G As Integer            'displayed in the dialogue as
  Dim myRGB_B As Integer            '"Current" color
  
  'save original palette color, because we don't really want to change it
  myOrgColor = ActiveWorkbook.Colors(ColorIndexLast)
  
  If i_OldColor = xlNone Then
    'get RGB values of background color, so the "Current" color looks empty
    Color2RGB BGColor, myRGB_R, myRGB_G, myRGB_B
  Else
    'get RGB values of i_OldColor
    Color2RGB i_OldColor, myRGB_R, myRGB_G, myRGB_B
  End If
  
  'call the color picker dialogue
  If Application.Dialogs(xlDialogEditColor).Show(ColorIndexLast, _
     myRGB_R, myRGB_G, myRGB_B) = True Then
    '"OK" was pressed, so Excel automatically changed the palette
    'read the new color from the palette
    PickNewColor = ActiveWorkbook.Colors(ColorIndexLast)
    'reset palette color to its original value
    ActiveWorkbook.Colors(ColorIndexLast) = myOrgColor
  Else
    '"Cancel" was pressed, palette wasn't changed
    'return old color (or xlNone if no color was passed to the function)
    PickNewColor = i_OldColor
  End If

End Function

Sub Color2RGB(ByVal i_Color As Long, _
              o_R As Integer, o_G As Integer, o_B As Integer)
  o_R = i_Color Mod 256
  i_Color = i_Color \ 256
  o_G = i_Color Mod 256
  i_Color = i_Color \ 256
  o_B = i_Color Mod 256
End Sub


Function checkRowBlank(sh As Worksheet, iRow) As Boolean
  If (sh.Cells(iRow, PHBAR_COL_ActID).Value = "" Or IsEmpty(sh.Cells(iRow, PHBAR_COL_ActID).Value)) And _
     (sh.Cells(iRow, PHBAR_COL_ActDesc).Value = "" Or IsEmpty(sh.Cells(iRow, PHBAR_COL_ActDesc).Value)) And _
     (sh.Cells(iRow, PHBAR_COL_PLANST).Value = "" Or IsEmpty(sh.Cells(iRow, PHBAR_COL_PLANST).Value)) Then
    checkRowBlank = True
  Else
    checkRowBlank = False
  End If
End Function
  

'날짜형식 체크
Function validDate(dt)
    If IsDate(dt) Then
      validDate = Int(dt)
    Else
      validDate = 0
    End If
End Function
