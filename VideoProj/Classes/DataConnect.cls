VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

Private moConn As Connection

' ADO Provider Information
Private mstrProvider As String         ' Provider String
Private mintProviderConst As Integer   ' Provider Constant
Private mstrInitialCatalog As String   ' Initial Catalog
Private mstrDataSource As String       ' Data Source

Private mstrLoginID As String          ' Login ID
Private mstrPassword As String         ' User Password

' ODBC Provider Information
Private mstrDSN As String              ' DSN
Private mstrDatabase As String         ' Database to use
Private mboolUseODBC As Boolean        ' Use ODBC Provider

Private mstrErrMsg As String           ' Error Messages

Public Enum pdsaProvider
   pdsaJet = 1
   pdsaSQLServer = 2
End Enum

Property Get ProviderConst() As pdsaProvider
   ProviderConst = mintProviderConst
End Property

Property Let ProviderConst(intProvider As pdsaProvider)
   mintProviderConst = intProvider
   Select Case intProvider
      Case 1
         mstrProvider = "Microsoft.Jet.OLEDB.4.0"
      Case 2
         mstrProvider = "Sqloledb"
      Case Else
         mstrProvider = "Microsoft.Jet.OLEDB.4.0"
   End Select
End Property

Property Get Connection() As Connection
   Set Connection = moConn
End Property

Property Get Provider() As String
   Provider = mstrProvider
End Property

Property Let Provider(ByVal strTemp As String)
   mstrProvider = strTemp
End Property

Property Let LoginId(ByVal strValue As String)
   mstrLoginID = strValue
End Property

Property Get LoginId() As String
   LoginId = mstrLoginID
End Property

Property Get Password() As String
   Password = mstrPassword
End Property

Property Let Password(ByVal strPassword As String)
   mstrPassword = strPassword
End Property

Property Get InitialCatalog() As String
   InitialCatalog = mstrInitialCatalog
End Property

Property Let InitialCatalog(ByVal strTemp As String)
   mstrInitialCatalog = strTemp
End Property

Property Get DataSource() As String
   DataSource = mstrDataSource
End Property

Property Let DataSource(ByVal strTemp As String)
   mstrDataSource = strTemp
End Property

Property Get ErrorMsg() As String
   ErrorMsg = mstrErrMsg
End Property

Property Let UseODBC(ByVal boolUseODBC As Boolean)
   mboolUseODBC = boolUseODBC
End Property

Property Get UseODBC() As Boolean
   UseODBC = mboolUseODBC
End Property

Property Let DatabaseName(ByVal strDatabase As String)
   mstrDatabase = strDatabase
End Property

Property Get DatabaseName() As String
   DatabaseName = mstrDatabase
End Property

Property Let DSN(ByVal strDSN As String)
   mstrDSN = strDSN
End Property

Property Get DSN() As String
   DSN = mstrDSN
End Property

Public Function ConnectString() As String
   Dim strReturn As String
   
   If mboolUseODBC Then
      ' Build ODBC String
      strReturn = "DSN=" & mstrDSN
      If Len(mstrLoginID) Then
         strReturn = strReturn & ";UID=" & mstrLoginID
      End If
      If Len(mstrLoginID) Then
         strReturn = strReturn & ";PWD=" & mstrPassword
      End If
      If Len(mstrDatabase) Then
         strReturn = strReturn & ";DATABASE = " & mstrDatabase
      End If
      ConnectString = strReturn
   Else
      If mintProviderConst = pdsaJet Then
         ' Build OLE DB Provider String
         ConnectString = "Provider        = " & mstrProvider & ";" & _
                         "Data Source     = " & mstrDataSource
      Else
         ' Build OLE DB Provider String
         ConnectString = "Provider        = " & mstrProvider & ";" & _
                         "Data Source     = " & mstrDataSource & ";" & _
                         "Initial Catalog = " & mstrInitialCatalog & ";" & _
                         "User Id         = " & mstrLoginID & ";" & _
                         "Password        = " & mstrPassword & "; "
      End If
   End If
End Function

Public Function NetworkUserID() As String
    On Error GoTo ErrorHandler
    
    Dim lngReturn As Long
    Dim lngLen As Long
    Dim strUser As String * 256
    
    lngLen = 255
    lngReturn = WNetGetUser(vbNullString, strUser, lngLen)
    
    If ((lngReturn <> 0) Or (0 = InStr(strUser, Chr$(0)))) Then
        NetworkUserID = ""
    Else
        NetworkUserID = Left$(strUser, InStr(strUser, Chr$(0)) - 1)
    End If

    Exit Function

ErrorHandler:
   mstrErrMsg = Err.Description
   Exit Function
End Function

Private Sub Class_Initialize()
   ' Call Property Let to set provider string
   ProviderConst = pdsaJet
   ' Default Data source to Employees.mdb
   mstrDataSource = "VideoRentals.mdb"
   ' Create a New Connection
   Set moConn = New Connection
End Sub

Public Function DataOpen() As Boolean
   On Error GoTo Open_EH
   
   If moConn.State = adStateOpen Then
      moConn.Close
   End If
   
   moConn.CursorLocation = adUseClient
   
   ' Set the connection string by calling
   ' a function.
   moConn.ConnectionString = Me.ConnectString()
      
   ' Set the mode of the connection
   moConn.Mode = adModeReadWrite
      
   ' Open the Connection
   moConn.Open
   
   DataOpen = True
   
   Exit Function
   
Open_EH:
   Call ErrorHandler
   DataOpen = False
   Exit Function
End Function

Public Function ErrorHandler(Optional oConn As Connection) As String
   Dim oErr As ADODB.Error
   Dim strMsg As String
   
   If oConn Is Nothing Then
      Set oConn = moConn
   Else
      If oConn.State = adStateClosed Then
         Set oConn = moConn
      End If
   End If
   
   For Each oErr In oConn.Errors
      strMsg = strMsg & "Error #: " & oErr.Number & vbCrLf
      strMsg = strMsg & "Description: " & oErr.Description & vbCrLf
   Next
   
   mstrErrMsg = strMsg
   ErrorHandler = strMsg
End Function

Public Sub DataClose()
   On Error Resume Next

   If moConn.State = adStateOpen Then
      moConn.Close
   End If
   
   Set moConn = Nothing
End Sub
