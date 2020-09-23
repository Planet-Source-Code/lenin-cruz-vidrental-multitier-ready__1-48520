Attribute VB_Name = "modStartup"
Public goDataconn As DataConnection

Private Sub main()
Set goDataconn = New DataConnection

With goDataconn
    .DataSource = App.Path & "\Database\VideoRentals.mdb"
    .ProviderConst = pdsaJet
End With

If goDataconn.DataOpen() Then
    frmMain.Show
Else
    MsgBox (goDataconn.ErrorMsg)
End If
 

    
End Sub
