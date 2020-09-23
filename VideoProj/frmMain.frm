VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDateDue 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtDateRented 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d""/""MMMM""/""yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtGenre 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtVideoID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtMediaType 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtCast 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "frmMain.frx":0000
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtOverView 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmMain.frx":0006
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtDirector 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txtVideoCategory 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtVideoName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   2895
   End
   Begin VB.ListBox lstVideos 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Overview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6480
      TabIndex        =   18
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cast"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6480
      TabIndex        =   19
      Top             =   3840
      Width           =   1815
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1200
      Top             =   4560
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   43
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":000C
            Key             =   "Downsize"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":045E
            Key             =   "Upsize"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08B0
            Key             =   "SP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D02
            Key             =   "ASP"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1154
            Key             =   "Camera"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1266
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16B8
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B0A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C1C
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":206E
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2180
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2292
            Key             =   "SaveAll"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26E4
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B36
            Key             =   "Database"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F88
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":309A
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":393E
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41E2
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4634
            Key             =   "Generate"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A86
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B98
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CAA
            Key             =   "NewRecord"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4DBC
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4ECE
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5320
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5432
            Key             =   "Pencil"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5884
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5996
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6118
            Key             =   "Property"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":622A
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":633C
            Key             =   "SaveRecord"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":644E
            Key             =   "Ruler"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68A0
            Key             =   "Security"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CF2
            Key             =   "SortA"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E04
            Key             =   "SortD"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F16
            Key             =   "Spelling"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7028
            Key             =   "DeleteRecord"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":747A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":78CC
            Key             =   "UndoRecord"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79DE
            Key             =   "Waste"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7E30
            Key             =   "Tool"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Label Label10 
      Caption         =   "Rented Price"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   23
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   22
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Due Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   21
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Date Rented"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   20
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Director"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Genre"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Media Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Video Categoy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Video Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Video ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' Less coding at the front means other members of the team
' who's not that advance can work on the Form Design while the advanced
' members can work on the classes. He/she didn't need to know the whole
' working of the class! he/she needs only to know the exposed
' properties of the class - AddNew - Delete - Update
'
'**************************************************************

Option Explicit

' ADO Recordset
Private moVideo As clsVideos

Private Sub ListLoad()
  lstVideos.Clear
  With moVideo
    .SelectFilter = dacSelectclsVideoListBox
    .OrderByFilter = dacOrderByclsVideoID
      
    Set .DataConnection = goDataconn
    If .OpenRecordset() Then
      Do Until .EOF
        lstVideos.AddItem .VideoName
        lstVideos.ItemData(lstVideos.NewIndex) = .VideoID
        .MoveNext
      Loop
    Else
     If Len(.InfoMsg) Then
      MsgBox .InfoMsg
     End If
    End If
    .CloseRecordset
    If lstVideos.ListCount > 0 Then
      lstVideos.ListIndex = 0
    End If
  End With
End Sub

Private Sub lstVideos_Click()

  If lstVideos.ListIndex <> -1 Then
    With moVideo
      .SelectFilter = dacSelectclsVideoAll
      .WhereFilter = dacWhereclsVideoID
      .VideoID = lstVideos.ItemData(lstVideos.ListIndex)
       If .Find() Then
        Call FormShow
       Else
        If Len(.InfoMsg) Then
          MsgBox .InfoMsg
        End If
       End If
        .CloseRecordset
    End With
  End If
    
End Sub

Private Sub cmdExit_Click()
 'AppQuit
    Unload Me
End Sub

Private Function Update() As Boolean
    
  If moVideo.Replace() Then
    MsgBox "Update Successfull"
    ' Re-Read The Data
    lstVideos.List(lstVideos.ListIndex) = moVideo.VideoName
    Call FormShow
    Call ToggleButtons
    Update = True
    
  Else
    If Len(moVideo.InfoMsg) Then
      MsgBox moVideo.InfoMsg
      Update = False
    End If
  End If
 
End Function

Private Function AddNew() As Boolean
    
  If moVideo.AddNew() Then
    tbrMain.Tag = ""
    Call ToggleButtons
    AddNew = True
    lstVideos.AddItem (moVideo.VideoName)
    lstVideos.ItemData(lstVideos.NewIndex) = moVideo.VideoID
  Else
    If Len(moVideo.InfoMsg) Then
      MsgBox moVideo.InfoMsg
      AddNew = False
    End If
  End If
      
End Function

Private Sub FormNew()
   tbrMain.Tag = "Add"
   Call ToggleButtons
   Call ClearTextBoxes
   txtDateRented = Format(Now(), "short date")
   txtVideoID.Enabled = True
   txtVideoID.SetFocus
End Sub

Private Function FormSave() As Boolean
   ' Move Form Data Into Properties
   Call CopyObjects

   If tbrMain.Tag = "Add" Then
      FormSave = AddNew()
   Else
      FormSave = Update()
   End If
End Function

Private Sub FormCancel()
   tbrMain.Tag = ""
   Call ToggleButtons
   If lstVideos.ListCount > 0 Then
      moVideo.VideoID = lstVideos.ItemData(lstVideos.ListIndex)
      If moVideo.Find() Then
         Call FormShow
      End If
   End If
End Sub

Private Sub FormDelete()
Dim lngIndex As Long
  
  If DeleteAsk("Do you wish to delete the current record?") Then
    lngIndex = lstVideos.ListIndex
    With moVideo
      .WhereFilter = dacWhereclsVideoID
      .VideoID = lstVideos.ItemData(lngIndex)
       If .Delete() Then
        lstVideos.RemoveItem (lngIndex)
        Call ClearTextBoxes
        Call ToggleButtons
        If lstVideos.ListCount > 0 Then
          lstVideos.ListIndex = 0
        End If
      Else
        If Len(.InfoMsg) Then
          MsgBox .InfoMsg
        End If
      End If
    End With
  End If
  
End Sub

Private Sub Form_Load()
  Screen.MousePointer = vbHourglass
   
   Set moVideo = New clsVideos
   
  ' Initialize Form
   Call FormInit
  ' Initialize Toolbar
   Call FormToolbar
   
   Call ListLoad
    
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
    Set moVideo = Nothing
End Sub

Private Sub FormInit()

   ' Set Toolbar's ImageList property
   Set tbrMain.ImageList = imgList
  
End Sub

Private Sub ClearTextBoxes()
  Dim Actrls As Object
  For Each Actrls In Controls
    If TypeOf Actrls Is TextBox Then
      Actrls.Text = ""
    End If
  Next
    
End Sub

Private Sub TextChanged()
   If Not FormToolEnabled(tbrMain, "SaveRecord") Then
      If tbrMain.Tag <> "Show" Then
         Call ToggleButtons
      End If
   End If
End Sub

Private Sub ToggleButtons()
   lstVideos.Enabled = Not lstVideos.Enabled
   Call FormToolBarToggle(tbrMain)
End Sub

Private Sub FormToolbar()
   Dim btn As MSComctlLib.Button
   
   Call ToolbarSetup(tbrMain)
   
   '**************************************
   ' Add your own Toolbar buttons here
   '**************************************
   Set btn = tbrMain.Buttons.Add(, "Find", , , "Find")
   btn.ToolTipText = "Find"
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
      Case "NewRecord"
         Me.Tag = "New"
         Call FormNew
      Case "DeleteRecord"
         Call FormDelete
         Me.Tag = ""
      Case "SaveRecord"
         Call FormSave
      Case "UndoRecord"
         Call FormCancel
   End Select
End Sub

Private Sub FormShow()
   Dim strOldMsg As String

   Screen.MousePointer = vbHourglass

   tbrMain.Tag = "Show"
   ' Fill in all fields from Object
   With moVideo
      txtVideoID = .VideoID
      txtVideoName = .VideoName
      txtVideoCategory = .VideoCategory
      txtMediaType = .VideoMedia
      txtDirector = .VideoDirector
      txtOverView = .VideoOverView
      txtGenre = .VideoGenre
      txtCast = .VideoCastings
      txtDateRented = Format(.VideoDateRented, "Short Date")
      txtDateDue = Format(.VideoDateDue, "Short Date")
      txtPrice = .VideoPrice
   End With
   
   tbrMain.Tag = ""
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub CopyObjects()
  With moVideo
    .VideoID = txtVideoID
    .VideoName = txtVideoName
    .VideoCategory = txtVideoCategory
    .VideoMedia = txtMediaType
    .VideoDirector = txtDirector
    .VideoOverView = txtOverView
    .VideoGenre = txtGenre
    .VideoCastings = txtCast
    .VideoDateRented = txtDateRented
    .VideoDateDue = txtDateDue
    .VideoPrice = txtPrice
  End With
End Sub


Private Sub txtCast_Change()
  Call TextChanged
End Sub

Private Sub txtDateDue_Change()
  Call TextChanged
End Sub

Private Sub txtDateRented_Change()
  Call TextChanged
End Sub

Private Sub txtDirector_Change()
  Call TextChanged
End Sub

Private Sub txtGenre_Change()
  Call TextChanged
End Sub

Private Sub txtMediaType_Change()
  Call TextChanged
End Sub

Private Sub txtOverView_Change()
  Call TextChanged
End Sub

Private Sub txtPrice_Change()
  Call TextChanged
End Sub

Private Sub txtVideoCategory_Change()
  Call TextChanged
End Sub

Private Sub txtVideoID_Change()
  Call TextChanged
End Sub

Private Sub txtVideoName_Change()
  Call TextChanged
End Sub
