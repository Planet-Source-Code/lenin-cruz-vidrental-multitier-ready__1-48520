Attribute VB_Name = "ComboBoxHelper"
Option Explicit

Public Function ListSearch(lstCtrl As ComboBox, lngSearchValue As Long) As Integer

  Dim lngIndex As Long
  MsgBox (lstCtrl.ListCount)
  If lstCtrl.ListCount <> -1 Then
      For lngIndex = 0 To lstCtrl.ListCount - 1
        If lstCtrl.ItemData(lngIndex) = lngSearchValue Then
          ListSearch = lngIndex
          Exit Function
        End If
      Next
  End If
  
  ListSearch = -1

End Function

'************************************************************
'* Function Name: ListReposition()
'* Copyright    : Copyright 1995-1998 PDSA, Inc.
'************************************************************
Public Function ListReposition(lstCtrl As Control, intIndex As Integer) As Integer
   If lstCtrl.ListCount = 0 Then
      ListReposition = -1
   Else
      intIndex = intIndex + 1
      
      If intIndex >= lstCtrl.ListCount - 1 Then
         lstCtrl.ListIndex = lstCtrl.ListCount - 1
      Else
         intIndex = intIndex - 1
         If intIndex <= 0 Then
            lstCtrl.ListIndex = 0
         Else
            lstCtrl.ListIndex = intIndex
         End If
      End If
      ListReposition = lstCtrl.ListIndex
   End If
End Function
