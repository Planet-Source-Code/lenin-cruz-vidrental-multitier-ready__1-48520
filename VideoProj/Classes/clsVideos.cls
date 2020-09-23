VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVideos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************
'I like the way PDSA make Multi Tier Programs thats why
'I just follow there CODING STANDARDS its very neat and
'can easily be trasported into MULTI-TIER
'*******************************************************

Option Explicit
'******************************************************
' Private Data To Match the Table Definition
'******************************************************
Private mlngVideoID As Long
Private mstrVideoName As String
Private mstrCategory As String
Private mstrMedia As String
Private mstrDirector As String
Private mstrOverview As String
Private mstrGenre As String
Private mstrCastings As String
Private mstrDateRented As String
Private mstrDue As String
Private mstrPrice As String
Private mlngRenterID As Long
Private mstrLastUpdate As String

'******************************************************
'* The following enumerations will change for each
'* data access class
'******************************************************
Public Enum dacSelectFilterclsVideos
   dacSelectclsVideoAll = 0
   dacSelectclsVideoListBox = 1
End Enum

Public Enum dacWhereFilterclsVideos
   dacWhereclsVideoName = 1
   dacWhereclsVideoLikeName = 2
   dacWhereclsVideoID = 3
End Enum

Public Enum dacOrderByFilterclsVideos
    dacOrderByclsVideoName = 1
    dacOrderByclsVideoID = 2
    dacOrderByclsVideoCategory = 3
End Enum

'******************************************************
'* These are all standard properties below this line
'******************************************************
' ADO Recordset
Private moRS As ADODB.Recordset
' Data Connection Information
Private moDataConn As DataConnection

' ADO Recordset Types/Options/Lock Types
Private mintRSType As CursorTypeEnum
Private mintOptions As CursorOptionEnum
Private mintLockType As LockTypeEnum

' Transaction Processing Flag
Private mstrTransType As String

' Filters
Private mlngWhereFilter As Long
Private mlngOrderByFilter As Long
Private mlngSelectFilter As Long
Private mlngInsertFilter As Long
Private mlngUpdateFilter As Long

' Misc. Private Data
Private mstrSQL As String
Private mlngRowsAffected As Long
Private mboolEOF As Boolean
Private mstrClassName As String

' For Info & Error Handling Messages
Private mstrInfoMsg As String
Private mlngInfoCode As Long
Private mstrFieldInErr As String

'******************************************************
'* Properties to Match Table Definition
'******************************************************
Property Let VideoID(ByVal lngVideoID As Long)
   mlngVideoID = lngVideoID
End Property

Property Get VideoID() As Long
   VideoID = mlngVideoID
End Property

Property Let VideoName(ByVal strVideoName As String)
    mstrVideoName = strVideoName
End Property

Property Get VideoName() As String
    VideoName = mstrVideoName
End Property

Property Let VideoCategory(ByVal strCategory As String)
    mstrCategory = strCategory
End Property

Property Get VideoCategory() As String
    VideoCategory = mstrCategory
End Property

Property Let VideoOverView(ByVal strOverView As String)
    mstrOverview = strOverView
End Property

Property Get VideoOverView() As String
    VideoOverView = mstrOverview
End Property

Property Let VideoMedia(ByVal strMedia As String)
    mstrMedia = strMedia
End Property

Property Get VideoMedia() As String
    VideoMedia = mstrMedia
End Property

Property Let VideoDirector(ByVal strDirector As String)
    mstrDirector = strDirector
End Property

Property Get VideoDirector() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VideoDirector
    VideoDirector = mstrDirector
End Property

Public Property Let VideoGenre(ByVal strGenre As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VideoGenre = 5
    mstrGenre = strGenre
End Property

Public Property Get VideoGenre() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VideoGenre
    VideoGenre = mstrGenre
End Property

Public Property Let VideoCastings(ByVal strCastings As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VideoCastings = 5
    mstrCastings = strCastings
End Property

Public Property Get VideoCastings() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VideoCastings
    VideoCastings = mstrCastings
End Property
'================================================================================
Public Property Let VideoDateRented(ByVal strDateRented As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VideoDateRented = 5
    mstrDateRented = strDateRented
End Property

Public Property Get VideoDateRented() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VideoDateRented
    VideoDateRented = mstrDateRented
End Property
'================================================================================

'================================================================================
Property Let VideoDateDue(ByVal strDateDue As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VideoDateDue = 5
    mstrDue = strDateDue
End Property

Public Property Get VideoDateDue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VideoDateDue
    VideoDateDue = mstrDue
End Property
'================================================================================

Property Let VideoPrice(ByVal strPrice As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VideoPrice = 5
    mstrPrice = strPrice
End Property

Property Get VideoPrice() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VideoPrice
    VideoPrice = mstrPrice
End Property

Public Property Let VideoRenterId(ByVal lngRenterID As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VideoRenterId = 5
    mlngRenterID = lngRenterID
End Property

Public Property Get VideoRenterId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VideoRenterId
    VideoRenterId = mlngRenterID
End Property

Property Let LastUpdate(strLastUpdate As String)
    mstrLastUpdate = strLastUpdate
End Property

Property Get LastUpdate() As String
    LastUpdate = mstrLastUpdate
End Property

'*********************************************************
'* The following filters will change for each
'* data access class
'*********************************************************
Property Get SelectFilter() As dacSelectFilterclsVideos
   SelectFilter = mlngSelectFilter
End Property

Property Let SelectFilter(ByVal lngValue As dacSelectFilterclsVideos)
   mlngSelectFilter = lngValue
End Property

Property Get WhereFilter() As dacWhereFilterclsVideos
   WhereFilter = mlngWhereFilter
End Property

Property Let WhereFilter(ByVal lngValue As dacWhereFilterclsVideos)
   mlngWhereFilter = lngValue
End Property

Property Get OrderByFilter() As dacOrderByFilterclsVideos
   OrderByFilter = mlngOrderByFilter
End Property

Property Let OrderByFilter(ByVal lngValue As dacOrderByFilterclsVideos)
   mlngOrderByFilter = lngValue
End Property

'*********************************************************
'* These are all standard properties below this line
'*********************************************************
Property Let Options(ByVal intValue As CursorOptionEnum)
   mintOptions = intValue
End Property

Property Get Options() As CursorOptionEnum
   Options = mintOptions
End Property

Property Let RSType(ByVal intValue As CursorTypeEnum)
   mintRSType = intValue
End Property

Property Get RSType() As CursorTypeEnum
   RSType = mintRSType
End Property

Property Let LockType(ByVal intValue As LockTypeEnum)
   mintLockType = intValue
End Property

Property Get LockType() As LockTypeEnum
   LockType = mintLockType
End Property

Property Get InfoMsg() As String
   InfoMsg = mstrInfoMsg
End Property

Property Let InfoMsg(ByVal strValue As String)
   mstrInfoMsg = strValue
End Property

Property Get FieldInError() As String
   FieldInError = mstrFieldInErr
End Property

Property Get DataConnection() As DataConnection
   Set DataConnection = moDataConn
End Property

Property Set DataConnection(oData As DataConnection)
   Set moDataConn = oData
End Property

Property Let EOF(boolValue As Boolean)
   mboolEOF = boolValue
End Property

Property Get EOF() As Boolean
   EOF = mboolEOF
End Property

Property Let SQL(strValue As String)
   mstrSQL = strValue
End Property

Property Get SQL() As String
   SQL = mstrSQL
End Property
'************************************************************
'* Method Name  : Class_Initialize
'* Syntax       : <Object>.Class_Initialize()
'* Description  : This event is called when the first property or method is accessed from
'*                the object.  It is used to initialize some private data variables.
'************************************************************
Private Sub Class_Initialize()
    Set moRS = New ADODB.Recordset
    mstrClassName = "clsVideos"
    mintRSType = adOpenForwardOnly
    mintOptions = adLockReadOnly
End Sub
'************************************************************
'* Method Name  : Class_Terminate
'* Copyright    : Lenin M. Cruz
'* Syntax       : <Object>.Class_Initialize()
'* Parameters   : None
'*
'* Description  : This event is called when the object goes out of scope
'*
'* Example:
'*
'************************************************************
Private Sub Class_Terminate()
    On Error Resume Next
    Call CloseRecordset
    Set moRS = Nothing
End Sub

Public Sub CloseRecordset()
   On Error Resume Next

   mlngSelectFilter = 0
   mlngWhereFilter = 0
   mlngOrderByFilter = 0
   mlngInsertFilter = 0
   mlngUpdateFilter = 0
   mstrSQL = vbNullString
   mboolEOF = True
   moRS.Close
   
End Sub

Public Sub CopyObject(oCopy As clsVideos)
    With oCopy
        .VideoID = mlngVideoID
        .VideoName = mstrVideoName
        .VideoCategory = mstrCategory
        .VideoCastings = mstrCastings
        .VideoDateDue = mstrDue
        .VideoDateRented = mstrDateRented
        .VideoOverView = mstrOverview
        .VideoGenre = mstrGenre
        .VideoMedia = mstrMedia
        .VideoPrice = mstrPrice
        .VideoRenterId = mlngRenterID
    End With
End Sub


'************************************************************
'* Method Name  : Find
'* Syntax       : Boolean = <Object>.Find()
'* Parameters   : None
'*
'* Description  : This method will find any row(s) in the table based on the FilterName property being filled in, and the corresmonding properties.
'*
'* Example:
'*
'************************************************************
Public Function Find() As Boolean
   ' Build Recordset & Retrieve Data
   If OpenRecordset() Then
      Find = True
   Else
      Find = False
   End If
End Function

'************************************************************
'* Method Name  : InfoMsgClear
'* Syntax       : Call <Object>.InfoMsgClear()
'* Parameters   : None
'*
'* Description  : This Private method will clear the information message properties.
'*
'* Example:
'*
'************************************************************
Private Sub InfoMsgClear()
   mstrFieldInErr = vbNullString
   mstrInfoMsg = vbNullString
   mlngInfoCode = 0
End Sub

'************************************************************
'* Method Name  : MoveNext
'* Syntax       : Boolean = <Object>.MoveNext()
'* Parameters   : None
'*
'* Description  : This method will Move to the next row in the record set.
'*
'* Example:
'
'************************************************************
Public Function MoveNext() As Boolean
   On Error GoTo Error_EH

   ' Use Recordset Class To Move
   If Not moRS.EOF Then
      moRS.MoveNext
      If Not moRS.EOF Then
         ' Load Properties
         Call Retrieve
         MoveNext = True
      Else
         mboolEOF = True
         MoveNext = False
         Call CloseRecordset
      End If
   Else
      mboolEOF = True
      MoveNext = False
      Call CloseRecordset
   End If

   Exit Function
   
Error_EH:
   mstrInfoMsg = mstrClassName & ".MoveNext()" & _
                 vbCrLf & moDataConn.ErrorHandler() & _
                 vbCrLf & Err.Description
   mlngInfoCode = Err.Number
   Exit Function
End Function

'************************************************************
'* Method Name  : MovePrevious
'* Syntax       : Boolean = <Object>.MovePrevious()
'* Parameters   : None
'*
'* Description  : This method will Move to the previous row in the record set.
'*
'* Example:
'*
'************************************************************
Public Function MovePrevious() As Boolean
   On Error GoTo Error_EH

   ' Use Recordset Class To Move
   If Not moRS.BOF Then
      moRS.MovePrevious
      If Not moRS.BOF Then
         ' Load Properties
         Call Retrieve
         MovePrevious = True
         mboolEOF = False
      Else
         MovePrevious = True
         moRS.MoveNext
      End If
   Else
      MovePrevious = False
   End If

   Exit Function
   
Error_EH:
   mstrInfoMsg = mstrClassName & ".MovePrevious()" & _
                 vbCrLf & moDataConn.ErrorHandler() & _
                 vbCrLf & Err.Description
   mlngInfoCode = Err.Number
   Exit Function
End Function

'************************************************************
'* Method Name  : MoveFirst
'* Syntax       : Boolean = <Object>.MoveFirst()
'* Parameters   : None
'*
'* Description  : This method will Move to the first row in the record set.
'*
'* Example:
'*
'************************************************************
Public Function MoveFirst() As Boolean
   On Error GoTo Error_EH

   moRS.MoveFirst
   Call Retrieve
   MoveFirst = True
   mboolEOF = False

   Exit Function
   
Error_EH:
   mstrInfoMsg = mstrClassName & ".MoveFirst()" & _
                 vbCrLf & moDataConn.ErrorHandler() & _
                 vbCrLf & Err.Description
   mlngInfoCode = Err.Number
   Exit Function
End Function

'************************************************************
'* Method Name  : MoveLast
'* Syntax       : Boolean = <Object>.MoveLast()
'* Parameters   : None
'*
'* Description  : This method will Move to the last row in the record set.
'*
'* Example:
'*
'************************************************************
Public Function MoveLast() As Boolean
   On Error GoTo Error_EH

   moRS.MoveLast
   Call Retrieve
   MoveLast = True
   mboolEOF = True

   Exit Function
   
Error_EH:
   mstrInfoMsg = mstrClassName & ".MoveLast()" & _
                 vbCrLf & moDataConn.ErrorHandler() & _
                 vbCrLf & Err.Description
   mlngInfoCode = Err.Number
   Exit Function
End Function

'************************************************************
'* Method Name  : OpenRecordset
'* Syntax       : Boolean = <Object>.OpenRecordset()
'* Parameters   : None
'*
'* Description  : This method will open a new recordset.  It uses any FilterName set prior to calling this method.
'*
'* Example:
'*
'************************************************************
Public Function OpenRecordset() As Boolean
   On Error GoTo Error_EH

   ' Initialize Variables
   OpenRecordset = False

   ' Build Select Statement
   mstrSQL = SelectSQL()

   ' Clear Informational Message
   Call InfoMsgClear

   ' Use the Recordset Class to Open result set for this class
   If moRS.State = adStateOpen Then
      moRS.Close
   End If
   mboolEOF = False
   moRS.Open mstrSQL, _
             moDataConn.Connection, _
             mintRSType

   If Not moRS.EOF Then
      ' Retrieve the info
      If Retrieve() Then
         mboolEOF = False
         OpenRecordset = True
      Else
         mboolEOF = True
         OpenRecordset = False
      End If
   Else
      mboolEOF = True
      Call CloseRecordset
      OpenRecordset = False
   End If
   
   Exit Function
   
Error_EH:
   mstrInfoMsg = mstrClassName & ".OpenRecordset()" & _
                 vbCrLf & moDataConn.ErrorHandler() & _
                 vbCrLf & Err.Description & _
                 vbCrLf & "Did you set the DataConnection property on this Class?"
   mlngInfoCode = Err.Number
   Exit Function
End Function


'************************************************************
'* Method Name  : Retrieve
'* Syntax       : Boolean = <Object>.Retrieve()
'* Parameters   : None
'*
'* Description  : This Private method will retrieve the data from the result set and populate the appropriate properties.
'*
'* Example:
'*
'************************************************************
Private Function Retrieve() As Boolean
   ' Set Error Handling On
   On Error GoTo Error_EH

   Select Case mlngSelectFilter
      Case dacSelectclsVideoListBox
        mlngVideoID = Field2Long(moRS("Video_ID"))
        mstrVideoName = Field2Str(moRS("VideoName"))

      Case Else
        mlngVideoID = Field2Long(moRS("Video_ID"))
        mstrVideoName = Field2Str(moRS("VideoName"))
        mstrCategory = Field2Str(moRS("Category"))
        mstrMedia = Field2Str(moRS("Media"))
        mstrDirector = Field2Str(moRS("Director"))
        mstrOverview = Field2Str(moRS("Overview"))
        mstrGenre = Field2Str(moRS("Genre"))
        mstrCastings = Field2Str(moRS("Castings"))
        mstrDateRented = Field2Str(moRS("DateRented"))
        mstrDue = Field2Str(moRS("Due"))
        mstrPrice = Field2Str(moRS("Price"))
        mlngRenterID = Field2Long(moRS("Renter_ID"))
        
   End Select

   Retrieve = True
   
   Exit Function

Error_EH:
   Retrieve = False
   mstrInfoMsg = mstrClassName & ".Retrieve()" & _
                 vbCrLf & moDataConn.ErrorHandler() & _
                 vbCrLf & Err.Description & _
                 vbCrLf & "most likely a column name is misspelled in the Retrieve() method"
   mlngInfoCode = Err.Number

   Exit Function
End Function

'************************************************************
'* Method Name  : SelectSQL
'* Syntax       : String = <Object>.SelectSQL()
'* Parameters   : None
'*
'* Description  : This Private method is used to build the SELECT string to access the base table for this class
'*
'* Example:
'*
'************************************************************
Private Function SelectSQL() As String
   Dim strSQL As String

   Select Case mlngSelectFilter
      Case dacSelectclsVideoListBox
         strSQL = "SELECT "
         strSQL = strSQL & "Video_ID, "
         strSQL = strSQL & "VideoName "
         
         strSQL = strSQL & "FROM tblVideo"

      Case Else
         strSQL = "SELECT "
         strSQL = strSQL & "Video_ID, "
         strSQL = strSQL & "VideoName, "
         strSQL = strSQL & "Category, "
         strSQL = strSQL & "Media, "
         strSQL = strSQL & "Director, "
         strSQL = strSQL & "Overview, "
         strSQL = strSQL & "Genre, "
         strSQL = strSQL & "Castings, "
         strSQL = strSQL & "DateRented, "
         strSQL = strSQL & "Due, "
         strSQL = strSQL & "Price, "
         strSQL = strSQL & "Renter_ID "
         
         strSQL = strSQL & "FROM tblVideo"
         
   End Select

   ' Add Filter To This SELECT Statement
   strSQL = strSQL & WhereFilterGet()

   ' Add Sort To This SELECT Statement
   strSQL = strSQL & OrderByFilterGet()
   

   SelectSQL = strSQL

End Function

'************************************************************
'* Method Name  : WhereFilterGet
'* Syntax       : String = <Object>.WhereFilterGet()
'* Parameters   : None
'*
'* Description  : This Private method will return a WHERE clause that is added to the SELECT statement for filtering records for selection or deleting.
'************************************************************
Private Function WhereFilterGet() As String
   Dim strSQL As String

   ' Select Records using Requested Filter
   Select Case mlngWhereFilter
   
      Case dacWhereclsVideoName
      
        strSQL = strSQL & " WHERE szAddress = " & Str2Field(mstrVideoName)
         
      Case dacWhereclsVideoLikeName
        strSQL = strSQL & " WHERE VideoName LIKE " & Str2Field(mstrVideoName)
      
      Case dacWhereclsVideoID
        strSQL = strSQL & " WHERE Video_ID = " & mlngVideoID
      
   End Select

   WhereFilterGet = strSQL
End Function

'************************************************************
'* Method Name  : OrderByFilterGet
'* Syntax       : String = <Object>.OrderByFilterGet()
'* Parameters   : None
'* Description  : This Private method is used to retrieve the ORDER BY clause that will be added to the SELECT statement.  Choose a sort order by setting the SortOrder property prior to opening a recordset.
'*
'************************************************************
Private Function OrderByFilterGet() As String
   Dim strSQL As String

   ' Add ORDER BY
   Select Case mlngOrderByFilter
      
      Case dacOrderByclsVideoName
         strSQL = strSQL & " ORDER BY VideoName"
         
      Case dacOrderByclsVideoCategory
         strSQL = strSQL & " ORDER BY Category"
      
      Case dacOrderByclsVideoID
         strSQL = strSQL & " ORDER BY Video_ID"

   End Select

   OrderByFilterGet = strSQL
End Function

'************************************************************
'* Method Name  : AddNew
'* Syntax       : Boolean = <Object>.AddNew()
'* Parameters   : None
'*
'* Description  : Adds a new row to the base table for this class
'************************************************************
Public Function AddNew() As Boolean
   Dim strSQL As String

   On Error GoTo Error_EH

   AddNew = False

   ' Check Business Rules and Duplicate Values First
   If Check() Then
  '    If NoDuplicates(False) Then
        Select Case mlngInsertFilter
               Case Else
                  strSQL = "INSERT INTO tblVideo ("
                  strSQL = strSQL & "Video_ID, "
                  strSQL = strSQL & "VideoName, "
                  strSQL = strSQL & "Category, "
                  strSQL = strSQL & "Media, "
                  strSQL = strSQL & "Director, "
                  strSQL = strSQL & "Overview, "
                  strSQL = strSQL & "Genre, "
                  strSQL = strSQL & "Castings, "
                  strSQL = strSQL & "DateRented, "
                  strSQL = strSQL & "Due, "
                  strSQL = strSQL & "Price, "
                  strSQL = strSQL & "Renter_ID "
                  strSQL = strSQL & ") VALUES ("
                  strSQL = strSQL & mlngVideoID & ", "
                  strSQL = strSQL & Str2Field(mstrVideoName) & ", "
                  strSQL = strSQL & Str2Field(mstrCategory) & ", "
                  strSQL = strSQL & Str2Field(mstrMedia) & ", "
                  strSQL = strSQL & Str2Field(mstrDirector) & ", "
                  strSQL = strSQL & Str2Field(mstrOverview) & ", "
                  strSQL = strSQL & Str2Field(mstrGenre) & ", "
                  strSQL = strSQL & Str2Field(mstrCastings) & ", "
                  strSQL = strSQL & Str2Field(mstrDateRented) & ", "
                  strSQL = strSQL & Str2Field(mstrDue) & ", "
                  strSQL = strSQL & Str2Field(mstrPrice) & ", "
                  strSQL = strSQL & mlngRenterID & ")"
               End Select

            Call InfoMsgClear
            ' Submit the SQL Statement
            moDataConn.Connection.Execute strSQL, mlngRowsAffected
            If mlngRowsAffected = 1 Then
               mlngInsertFilter = 0
               AddNew = True
            Else
               AddNew = False
               mstrInfoMsg = mstrClassName & ".AddNew()" & vbCrLf & _
                             "Could not insert into this table: tblEmployees"
            End If
   '      End If
      End If
   
   Exit Function

Error_EH:
   AddNew = False
   mstrInfoMsg = mstrClassName & ".AddNew()" & _
                 vbCrLf & moDataConn.ErrorHandler() & _
                 vbCrLf & Err.Description
   mlngInfoCode = Err.Number
   Exit Function
End Function
'************************************************************
'* Method Name  : Replace
'* Syntax       : Boolean = <Object>.Replace()
'* Parameters   : None
'*
'* Description  : This method will replace all the fields in the base table with the information in the properties.
'*
'* Example:
'*
'************************************************************
Public Function Replace() As Boolean
   Dim strSQL As String

   On Error GoTo Error_EH

   If Check() Then
         Select Case mlngInsertFilter
               Case Else
                        
               strSQL = "UPDATE tblVideo SET "
               strSQL = strSQL & "VideoName = " & Str2Field(mstrVideoName) & ", "
               strSQL = strSQL & "Category = " & Str2Field(mstrCategory) & ", "
               strSQL = strSQL & "Media = " & Str2Field(mstrMedia) & ", "
               strSQL = strSQL & "Director = " & Str2Field(mstrDirector) & ", "
               strSQL = strSQL & "Overview = " & Str2Field(mstrOverview) & ", "
               strSQL = strSQL & "Genre = " & Str2Field(mstrGenre) & ", "
               strSQL = strSQL & "Castings = " & Str2Field(mstrCastings) & ", "
               strSQL = strSQL & "DateRented = " & Date2Field(mstrDateRented) & ", "
               strSQL = strSQL & "Due = " & Date2Field(mstrDue) & ", "
               strSQL = strSQL & "Price = " & Str2Field(mstrPrice) & ", "
               strSQL = strSQL & "Renter_ID = " & ID2Field(mlngRenterID)
               
               strSQL = strSQL & " WHERE "
               strSQL = strSQL & "Video_ID = " & VideoID & " "
               Debug.Print strSQL
               
         End Select

         ' Clear Informational Message
         Call InfoMsgClear

         ' Submit the SQL Statement
         moDataConn.Connection.Execute strSQL, mlngRowsAffected
         If mlngRowsAffected = 1 Then
            mlngInsertFilter = 0
            Replace = True
         Else
            Replace = False
            mstrInfoMsg = "Could Not Update This Table: tblVideo"
         End If
   End If

   Exit Function

Error_EH:
  mstrInfoMsg = mstrClassName & ".Replace()" & _
                vbCrLf & moDataConn.ErrorHandler() & _
                vbCrLf & Err.Description
   mlngInfoCode = Err.Number
   Exit Function
End Function
'************************************************************
'* Method Name  : Delete
'* Syntax       : Boolean = <Object>.AddNew()
'* Parameters   : None
'*
'* Description  : Adds a new row to the base table for this class
'************************************************************
Public Function Delete() As Boolean
Dim strSQL As String

On Error GoTo Error_EH

strSQL = "DELETE * FROM tblVideo " & WhereFilterGet()

moDataConn.Connection.Execute strSQL, mlngRowsAffected
 If mlngRowsAffected > 0 Then
   Delete = True
   Exit Function
 Else
   Delete = False
   mstrInfoMsg = mstrClassName & ".Delete()" & vbCrLf & _
                 "Could not delete into this table: tblVideo"
 End If
 
Error_EH:
   Delete = False
   mstrInfoMsg = mstrClassName & ".Delete()" & _
                 vbCrLf & moDataConn.ErrorHandler() & _
                 vbCrLf & Err.Description
   mlngInfoCode = Err.Number
   Exit Function
End Function
'************************************************************
'* Method Name  : Check
'* Syntax       : Boolean = <Object>.Check()
'* Parameters   : None
'*
'* Description  : This method performs business rule checking
'*
'* Example:
'*
'************************************************************
Public Function Check() As Boolean
   Check = False

   Call InfoMsgClear
   
   If Trim$(mstrVideoName) = "" Then
      mstrInfoMsg = mstrInfoMsg & "Video Name" & vbCrLf
      mstrFieldInErr = "VideoName"
   End If
   
   
   If mstrInfoMsg = "" Then
      Check = True
   End If
End Function
'************************************************************
'* Method Name  : NoDuplicates
'* Syntax       : Boolean = <Object>.NoDuplicates(<Boolean>)
'* Parameters   : boolReplace => True if UPDATE, False if INSERT
'*
'* Description  : This method will check for any duplicate values prior to inserting a new row into the base table.
'*
'* Example:
'*
'************************************************************
Public Function NoDuplicates(ByVal boolReplace As Boolean) As Boolean
   Dim strSQL As String
   Dim oRS As ADODB.Recordset

   On Error GoTo Error_EH
   ' Insert your Duplicate Checking SQL Here
   strSQL = "Select * FROM tblVideos WHERE VideoName = " & "'" & mstrVideoName & "'"

   If strSQL <> "" Then
      Set oRS = New Recordset
      oRS.Open strSQL, moDataConn, adOpenForwardOnly, adLockReadOnly
      If oRS.EOF Then
         NoDuplicates = True
      Else
         NoDuplicates = False
      End If
      oRS.Close
   Else
      NoDuplicates = True
   End If
   Set oRS = Nothing

   Exit Function

Error_EH:
   mstrInfoMsg = mstrClassName & ".NoDuplicates()" & _
                 vbCrLf & moDataConn.ErrorHandler() & _
                 vbCrLf & Err.Description
   mlngInfoCode = Err.Number
   NoDuplicates = False

   Exit Function
End Function

'************************************************************
'* Method Name  : RowCount
'* Syntax       : Long = <Object>.RowCount()
'* Parameters   : None
'*
'* Description  : This method returns the number of rows in a result set.  Uses the current Filter in the object
'*
'* Example:
'*
'************************************************************
Public Function RowCount() As Long
   Dim strSQL As String
   Dim lngReturn As Long

   On Error GoTo Error_EH

   Call InfoMsgClear

   strSQL = "SELECT Count(*) As NumRecs FROM tblVideo" & vbCrLf
   ' Add Filter To This SELECT Statement
   strSQL = strSQL & WhereFilterGet()

   moRS.Open strSQL, moDataConn.Connection, _
             adOpenForwardOnly, adLockReadOnly
   If moRS.EOF Then
      lngReturn = 0
   Else
      lngReturn = moRS("NumRecs")
   End If
   moRS.Close

   RowCount = lngReturn
   
   Exit Function

Error_EH:
   mstrInfoMsg = mstrClassName & ".RowCount()" & _
                 vbCrLf & moDataConn.ErrorHandler() & _
                 vbCrLf & Err.Description
   mlngInfoCode = Err.Number
   RowCount = 0

   Exit Function
End Function


