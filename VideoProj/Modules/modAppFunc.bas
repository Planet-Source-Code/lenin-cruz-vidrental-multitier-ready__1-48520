Attribute VB_Name = "ModAppFunctions"
Option Explicit

Public Function ConcurUpdate(intConcurID As Integer) As String
   Dim strSQL As String
   
   If intConcurID = -1 Then
      strSQL = strSQL & "iConcurrency_id = Null "
   Else
      If intConcurID > 32766 Then
         strSQL = strSQL & "iConcurrency_id = 1 "
      Else
         strSQL = strSQL & "iConcurrency_id = iConcurrency_id + 1 "
      End If
   End If
End Function


Public Function ItemData2Field(ctlList As Control) As String
   If ctlList.ListIndex = -1 Then
      ItemData2Field = "Null"
   Else
      ItemData2Field = ctlList.ItemData(ctlList.ListIndex)
   End If
End Function

Public Function Str2Field(strValue As String) As String
   If strValue = "" Then
      Str2Field = "Null"
   Else
      Str2Field = "'" & QuoteConvert(strValue) & "'"
   End If
End Function
Public Function ID2Field(ByVal lngValue As Long) As String
   If lngValue = -1 Then
      ID2Field = "Null"
   Else
      ID2Field = CStr(lngValue)
   End If
End Function
Public Function Num2Field(strValue As Variant) As String
   If IsNumeric(strValue) Then
      strValue = CStr(strValue)
      If strValue = "" Then
         Num2Field = "Null"
      Else
         Num2Field = strValue
      End If
   Else
      Num2Field = "Null"
   End If
End Function

Public Function CheckBox2Field(intValue As Integer) As Integer
   If intValue = vbChecked Then
      CheckBox2Field = 1
   Else
      CheckBox2Field = 0
   End If
End Function

Public Function Date2Field(strValue As String) As String
   If strValue = "" Then
      Date2Field = "Null"
   Else
      If IsDate(strValue) Then
         Date2Field = "'" & CDate(strValue) & "'"
      Else
         Date2Field = "Null"
      End If
   End If
End Function

Public Function Field2Str(vntField As Variant) As String
   If IsNull(vntField) Then
      Field2Str = ""
   Else
      Field2Str = Trim$(CStr(vntField))
   End If
End Function

Public Function Field2Long(vntField As Variant) As Long
   If IsNull(vntField) Then
      Field2Long = -1
   Else
      Field2Long = CLng(vntField)
   End If
End Function

Public Function Field2Int(vntField As Variant) As Integer
   If IsNull(vntField) Then
      Field2Int = -1
   Else
      Field2Int = CInt(vntField)
   End If
End Function

Public Function Field2CheckBox(vntField As Variant) As Integer
   If IsNull(vntField) Then
      Field2CheckBox = vbUnchecked
   Else
      Field2CheckBox = IIf(vntField, vbChecked, vbUnchecked)
   End If
End Function

Public Function QuoteConvert(strValue As String)
   QuoteConvert = Replace(strValue, "'", "''")
End Function



