VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDogs 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6060
   ClientLeft      =   810
   ClientTop       =   1305
   ClientWidth     =   8160
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   8160
   Begin VB.TextBox txt 
      DataField       =   "dtBirth_dt"
      DataSource      =   "tblDogs"
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtDirector 
      DataField       =   "szColor_nm"
      DataSource      =   "tblDogs"
      Height          =   375
      Left            =   5400
      TabIndex        =   22
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtMediatype 
      DataField       =   "szColor_nm"
      DataSource      =   "tblDogs"
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      Top             =   1920
      Width           =   2655
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.TextBox txtOverView 
      DataField       =   "sBark_type"
      DataSource      =   "tblDogs"
      Height          =   975
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox txtColor 
      DataField       =   "szColor_nm"
      DataSource      =   "tblDogs"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      DataField       =   "VideoName"
      DataSource      =   "tblVideo"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtDateRented 
      DataField       =   "dtBirth_dt"
      DataSource      =   "tblDogs"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "cPrice_amt"
      DataSource      =   "tblDogs"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtCost 
      DataField       =   "cCost_amt"
      DataSource      =   "tblDogs"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame fraSex 
      Caption         =   "Sex"
      Height          =   1155
      Left            =   6810
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
      Begin VB.OptionButton optMale 
         Alignment       =   1  'Right Justify
         Caption         =   "Male"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Tag             =   "sSex_nm"
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optFemale 
         Alignment       =   1  'Right Justify
         Caption         =   "Female"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Tag             =   "sSex_nm"
         Top             =   660
         Width           =   915
      End
   End
   Begin VB.ListBox lstNames 
      DataField       =   "szDog_nm"
      DataSource      =   "tblDogs"
      Height          =   2595
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   510
      Width           =   3675
   End
   Begin MSComctlLib.ImageList ilsList 
      Left            =   6240
      Top             =   5520
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
            Picture         =   "frmMains.frx":0000
            Key             =   "Downsize"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":0452
            Key             =   "Upsize"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":08A4
            Key             =   "SP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":0CF6
            Key             =   "ASP"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":1148
            Key             =   "Camera"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":125A
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":16AC
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":1AFE
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":1C10
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":2062
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":2174
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":2286
            Key             =   "SaveAll"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":26D8
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":2B2A
            Key             =   "Database"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":2F7C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":308E
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":34E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":3932
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":3D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":41D6
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":4628
            Key             =   "Generate"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":4A7A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":4B8C
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":4C9E
            Key             =   "NewRecord"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":4DB0
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":4EC2
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":5314
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":5426
            Key             =   "Pencil"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":5878
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":598A
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":610C
            Key             =   "Property"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":621E
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":6330
            Key             =   "SaveRecord"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":6442
            Key             =   "Ruler"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":6894
            Key             =   "Security"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":6CE6
            Key             =   "SortA"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":6DF8
            Key             =   "SortD"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":6F0A
            Key             =   "Spelling"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":701C
            Key             =   "DeleteRecord"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":746E
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":78C0
            Key             =   "UndoRecord"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":79D2
            Key             =   "Waste"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMains.frx":7E24
            Key             =   "Tool"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLabel 
      Caption         =   "Date Due"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   3840
      TabIndex        =   24
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   3840
      TabIndex        =   19
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblLabel 
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
      Height          =   300
      Index           =   4
      Left            =   3840
      TabIndex        =   18
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblLabel 
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
      Height          =   300
      Index           =   1
      Left            =   3870
      TabIndex        =   17
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      Caption         =   "Video Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   3870
      TabIndex        =   16
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblLabel 
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
      Height          =   300
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblLabel 
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
      Height          =   300
      Index           =   0
      Left            =   3870
      TabIndex        =   14
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label lblLabel 
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
      Height          =   300
      Index           =   6
      Left            =   3870
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblLabel 
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
      Height          =   300
      Index           =   7
      Left            =   3870
      TabIndex        =   12
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      Caption         =   "Cost"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   3960
      TabIndex        =   11
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblID 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "lDog_id"
      DataSource      =   "tblDogs"
      Height          =   315
      Left            =   5430
      TabIndex        =   1
      Top             =   510
      Width           =   915
   End
End
Attribute VB_Name = "frmDogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim moDogs As clsDogs
Dim mstrOldMsg As String
Dim mintConcur As Integer

Private Sub ComboLoad()
   Dim oBreeds As clsBreed
   Dim oBreeders As clsBreeders
   Dim boolPerform As Boolean

   ' Load Breeds
   Set oBreeds = New clsBreed
   With oBreeds
      Set .LoginInfo = goLogin
      Set .Connection = goConnect
      .SelectFilter = "LISTBOX"
      boolPerform = .OpenRecordset
      Do While boolPerform
         cboBreed.AddItem .BreedName
         cboBreed.ItemData(cboBreed.NewIndex) = .BreedId
         boolPerform = .MoveNext()
      Loop
      .CloseRecordset
   End With

   ' Load Breeders
   Set oBreeders = New clsBreeders
   With oBreeders
      Set .LoginInfo = goLogin
      Set .Connection = goConnect
      .SelectFilter = "LISTBOX"
      boolPerform = .OpenRecordset
      Do While boolPerform
         cboBreeder.AddItem .BreederName
         cboBreeder.ItemData(cboBreeder.NewIndex) = .BreederId
         boolPerform = .MoveNext()
      Loop
      .CloseRecordset
   End With
End Sub

Private Sub Form_Activate()
   Dim intResponse As Integer

   If lstNames.ListCount = 0 Then
      ' If No Records Prompt To Add Some
      intResponse = MsgBox("There are no records on file. Do you wish to add some ?", vbQuestion + vbYesNo)
      If intResponse = vbYes Then
         Call FormNew
      Else
         Unload Me
      End If
   End If
End Sub

Private Sub ListLoad()
   Dim boolPerform As Integer

   ' Clear the List Box
   lstNames.Clear

   ' Retrieve first record
   With moDogs
      .SelectFilter = "LISTBOX"
      boolPerform = .OpenRecordset()
      If boolPerform Then
         Do While boolPerform
            lstNames.AddItem .DogName
            lstNames.ItemData(lstNames.NewIndex) = CLng(.DogId)
            ' Retrieve next row
            boolPerform = .MoveNext()
         Loop
         .CloseRecordset

         ' Trigger a call to FormShow if data found
         If lstNames.ListCount > 0 Then
            lstNames.ListIndex = 0
         End If
      Else
         Call FormClear
      End If
      .SelectFilter = ""
   End With
End Sub


Private Sub lstNames_Click()
   Dim intIndex As Integer

   intIndex = lstNames.ListIndex
   If intIndex <> -1 Then
      moDogs.DogId = lstNames.ItemData(intIndex)
      If moDogs.Find() Then
         Call FormShow
      Else
         MsgBox "This record has been deleted from the table by another user."
         lstNames.RemoveItem intIndex
         intIndex = ListReposition(lstNames, intIndex)
         If intIndex = -1 Then
            Call FormClear
            Call Form_Activate
         Else
            ' Force Click Event
            If lstNames.ListIndex = intIndex Then
               Call FormShow
            Else
               lstNames.ListIndex = intIndex
            End If
         End If
         moDogs.CloseRecordset
      End If
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass

   ' Create the Data Access Object
   Call ObjectCreate

   ' Initialize Form
   Call FormInit

   ' Initialize Toolbar
   Call FormToolbar

   ' Load Breeds and Breeders
   Call ComboLoad

   Call ListLoad
   
   Screen.MousePointer = vbDefault
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Not tbrMain.Buttons("SaveRecord").Enabled Then
      ' If Window is minimized, cancel the keystroke
      If Me.WindowState <> vbMinimized Then
         ' See if the Keystroke will be Destructive
         If FormCheckKey(KeyCode, Shift) Then
            Call TextChanged
         End If
      Else
         ' Cancel the KeyStroke
         KeyCode = 0
      End If
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If Me.WindowState = vbMinimized Then
      KeyAscii = 0
   End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim intResponse As Integer

   If tbrMain.Buttons("SaveRecord").Enabled Then
      intResponse = FormCheckUnload(Me)
      Select Case intResponse
         Case vbYes
            Cancel = Not FormSave()

         Case vbCancel
            Cancel = True

      End Select
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next

   Set moDogs = Nothing
   Set frmDogs = Nothing
End Sub


Private Sub FormDelete()
   Dim intIndex As Integer

   If DeleteAsk("Do you wish to delete the current record?") Then
      intIndex = lstNames.ListIndex
      moDogs.DogId = lstNames.ItemData(intIndex)
      moDogs.ConcurrencyId = mintConcur
      If moDogs.Delete() Then
         lstNames.RemoveItem intIndex
         intIndex = ListReposition(lstNames, intIndex)
         If intIndex = -1 Then
            Call FormClear
            Call Form_Activate
         Else
            ' Force Click Event
            lstNames.ListIndex = intIndex
         End If
      Else
         MsgBox moDogs.InfoMsg
      End If
   End If
End Sub

Private Sub FormInit()

   ' Set Toolbar's ImageList property
   Set tbrMain.ImageList = ilsList

   txtBark.MaxLength = 10
   txtColor.MaxLength = 20
   txtName.MaxLength = 50
   txtBirthDate.MaxLength = 8
   txtPrice.MaxLength = 10
   txtCost.MaxLength = 10
End Sub

Private Function AddNew() As Boolean
   If moDogs.AddNew() Then
      Call ListAddItem(lstNames, moDogs.DogId, moDogs.DogName)
      lstNames.ListIndex = ListReposition(lstNames, lstNames.NewIndex)
      tbrMain.Tag = ""
      Call ToggleButtons
      AddNew = True
   Else
      MsgBox moDogs.InfoMsg, vbExclamation, Me.Caption
      AddNew = False
   End If
End Function

Private Sub FormToolbar()
   Dim btn As MSComctlLib.Button
   
   Call ToolbarSetup(tbrMain)
   
   '**************************************
   ' Add your own Toolbar buttons here
   '**************************************
   Set btn = tbrMain.Buttons.Add(, "Find", , , "Find")
   btn.ToolTipText = "Find"
End Sub

Private Sub TextChanged()
   If Not FormToolEnabled(tbrMain, "SaveRecord") Then
      If tbrMain.Tag <> "Show" Then
         Call ToggleButtons
      End If
   End If
End Sub

Private Function Replace() As Boolean
   If moDogs.Replace() Then
      lstNames.List(lstNames.ListIndex) = moDogs.DogName
      ' Re-Read The Data
      Call FormShow
      Call ToggleButtons
      Replace = True
   Else
      MsgBox moDogs.InfoMsg, vbExclamation, Me.Caption
      Replace = False
   End If
End Function

Private Function FormSave() As Boolean
   ' Move Form Data Into Properties
   Call FormMove

   If tbrMain.Tag = "Add" Then
      FormSave = AddNew()
   Else
      FormSave = Replace()
   End If
End Function



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

Private Sub ObjectCreate()
   Set moDogs = New clsDogs
   
   With moDogs
      Set .LoginInfo = goLogin
      Set .Connection = goConnect
   End With
End Sub

Private Sub FormClear()
   txtBark = ""
   txtColor = ""
   txtName = ""
   txtBirthDate = ""
   txtPrice = ""
   txtCost = ""
End Sub

Private Sub FormMove()
   With moDogs
      .Barktype = txtBark
      .ColorName = txtColor
      .DogName = txtName
      .BirthDate = txtBirthDate
      .PriceAmount = CCur(Val(txtPrice))
      .CostAmount = CCur(Val(txtCost))
      .DogId = lstNames.ItemData(lstNames.ListIndex)
      .BreederId = cboBreeder.ItemData(cboBreeder.ListIndex)
      .BreedId = cboBreed.ItemData(cboBreed.ListIndex)
      .ConcurrencyId = mintConcur
   End With
End Sub

Private Sub FormShow()
   Dim strOldMsg As String

   Screen.MousePointer = vbHourglass

   tbrMain.Tag = "Show"
   ' Fill in all fields from Object
   With moDogs
      lblID.Caption = .DogId
      txtBark = .Barktype
      txtColor = .ColorName
      txtName = .DogName
      txtBirthDate = DateShow(.BirthDate)
      txtPrice = .PriceAmount
      txtCost = .CostAmount
      mintConcur = .ConcurrencyId
      cboBreed.ListIndex = ListFindItem(cboBreed, .BreedId)
      cboBreeder.ListIndex = ListFindItem(cboBreeder, .BreederId)
   End With
   tbrMain.Tag = ""

   Screen.MousePointer = vbDefault
End Sub

Private Sub ToggleButtons()
   lstNames.Enabled = Not lstNames.Enabled
   Call FormToolBarToggle(tbrMain)
End Sub

Private Sub txtBark_Change()
   Call TextChanged
End Sub

Private Sub txtColor_Change()
   Call TextChanged
End Sub

Private Sub txtName_Change()
   Call TextChanged
End Sub

Private Sub txtBirthDate_Change()
   Call TextChanged
End Sub

Private Sub txtPrice_Change()
   Call TextChanged
End Sub

Private Sub txtCost_Change()
   Call TextChanged
End Sub

Private Sub FormNew()
   tbrMain.Tag = "Add"
   Call ToggleButtons
   Call FormClear
   txtName.SetFocus
End Sub


Private Sub FormCancel()
   tbrMain.Tag = ""
   Call ToggleButtons
   If lstNames.ListCount > 0 Then
      moDogs.DogId = lstNames.ItemData(lstNames.ListIndex)
      If moDogs.Find() Then
         Call FormShow
      End If
   Else
      Call Form_Activate
   End If
End Sub






