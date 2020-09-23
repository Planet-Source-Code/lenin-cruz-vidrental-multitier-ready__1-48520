Attribute VB_Name = "modSystem_Function"
Option Explicit

Public Function FormToolEnabled(tlbAny As MSComctlLib.Toolbar, ByVal strKey As String) As Boolean
   On Error Resume Next
   FormToolEnabled = tlbAny.Buttons(strKey).Enabled
End Function

Public Sub FormToolBarToggle(tlbAny As MSComctlLib.Toolbar)
   ' Do this to avoid errors with any buttons
   ' that have been taken away
   On Error Resume Next
   
   With tlbAny
      .Buttons("NewRecord").Enabled = Not .Buttons("NewRecord").Enabled
      .Buttons("DeleteRecord").Enabled = Not .Buttons("DeleteRecord").Enabled
      .Buttons("SaveRecord").Enabled = Not .Buttons("SaveRecord").Enabled
      .Buttons("UndoRecord").Enabled = Not .Buttons("UndoRecord").Enabled
   End With
End Sub

Public Sub ToolbarSetup(tlbAny As MSComctlLib.Toolbar)
   Dim btn As MSComctlLib.Button

   With tlbAny
      Set btn = .Buttons.Add(, , , tbrSeparator)
      Set btn = .Buttons.Add(, "NewRecord", , , "NewRecord")
      btn.ToolTipText = "New"
      Set btn = .Buttons.Add(, "DeleteRecord", , , "DeleteRecord")
      btn.ToolTipText = "Delete"
      Set btn = .Buttons.Add(, "SaveRecord", , , "SaveRecord")
      btn.ToolTipText = "Save Changes"
      btn.Enabled = False
      Set btn = .Buttons.Add(, "UndoRecord", , , "UndoRecord")
      btn.ToolTipText = "Undo Changes"
      btn.Enabled = False
      Set btn = .Buttons.Add(, "StdSeparator", , tbrSeparator)
   End With
End Sub

Public Function DeleteAsk(ByVal strMsg As String) As Integer
   If MsgBox(strMsg, vbQuestion + vbYesNo, Screen.ActiveForm.Caption) = vbYes Then
      DeleteAsk = True
   Else
      DeleteAsk = False
   End If
End Function

