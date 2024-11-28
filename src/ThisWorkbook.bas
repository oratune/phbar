Option Explicit

Public WithEvents chartSheet As Worksheet

Private Sub chartSheet_SelectionChange(ByVal Target As Range)
  Call formConfig.setVals(Target.row, Target.Column)
End Sub
