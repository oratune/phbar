VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formAbout 
   Caption         =   "About PhBarchart?"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   OleObjectBlob   =   "formAbout.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "formAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
  Me.Hide
End Sub

Private Sub CommandButton2_Click()
  Call http_CheckServer(True)
End Sub

Private Sub UserForm_Activate()
  lblVersion.Caption = "* Version : " & C_Ver
  lblLastUpdated.Caption = "* Last Updated : " & C_VerDate
End Sub
