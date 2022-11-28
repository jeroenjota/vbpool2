VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPointTypes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Punten types"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo cmbCategories 
      Height          =   315
      Left            =   4080
      TabIndex        =   8
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc dtcPointTypes 
      Height          =   330
      Left            =   1080
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "Nieuw"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Opslaan"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   7560
      Width           =   3645
   End
   Begin MSComCtl2.UpDown UpDnOrder 
      Height          =   375
      Left            =   6300
      TabIndex        =   4
      Top             =   7560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtOrder"
      BuddyDispid     =   196614
      OrigLeft        =   4800
      OrigTop         =   6840
      OrigRight       =   5055
      OrigBottom      =   7215
      Max             =   500
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtOrder 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5460
      TabIndex        =   3
      Top             =   7560
      Width           =   840
   End
   Begin MSDataGridLib.DataGrid grdPointTypes 
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11880
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Omschrijving"
         Caption         =   "Omschrijving"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Categorie"
         Caption         =   "Categorie"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "volgorde"
         Caption         =   "Volgorde"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   3404,977
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   1454,74
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   900,284
         EndProperty
      EndProperty
   End
   Begin VB.Label lblHead 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Punten omschrijving en volgorde"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmPointTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim newState As Boolean
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Private Sub btnClose_Click()
    Unload Me
    Debug.Print Err, Error
End Sub

Private Sub btnNew_Click()
    Me.txtDescription = ""
    Me.cmbCategories = ""
    Me.UpDnOrder = 1
    newState = True
End Sub

Private Sub btnSave_Click()
    'save record
    Set rs = New ADODB.Recordset
    
    Dim sqlstr As String
    Dim saveId As Long
    saveId = Me.dtcPointTypes.Recordset!id
    rs.Open "Select * from tblPointTypes", cn, adOpenKeyset, adLockOptimistic
    
    With rs
        If newState Then
           .AddNew
        End If
        !pointTypeDescription = Me.txtDescription
        !pointtypeCategory = val(Me.cmbCategories.BoundText)
        !pointtypelistorder = val(Me.txtOrder)
        .Update
    End With
    Me.dtcPointTypes.Refresh
    Me.grdPointTypes.Refresh
    Me.dtcPointTypes.Recordset.Find "id = " & saveId
    newState = False
    
End Sub

Private Sub Form_Load()
Dim sqlstr As String
    Set rs = New ADODB.Recordset
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With

    sqlstr = "Select pointtypeid as id, pointtypedescription as Omschrijving, pointcategoryId as categoryId, "
    sqlstr = sqlstr & "pointcategorydescription as categorie, pointtypelistorder as volgorde from tblPointTypes a"
    sqlstr = sqlstr & " inner join tblpointcategories b on a.pointtypecategory = b.pointcategoryid order by a.pointtypelistorder"
    Me.dtcPointTypes.ConnectionString = cn.ConnectionString
    Me.dtcPointTypes.RecordSource = sqlstr
    Me.dtcPointTypes.Refresh
    Set Me.grdPointTypes.DataSource = Me.dtcPointTypes
    
    sqlstr = "Select pointCategoryId as id, pointCategoryDescription as omschrijving from tblPointCategories order by pointCategoryID"
'    FillCombo Me.cmbCategories, sqlstr, cn, "omschrijving", "id"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    With Me.cmbCategories
        Set .RowSource = rs
        .BoundColumn = "id"
        .ListField = "omschrijving"
    End With
    
    centerForm Me
    UnifyForm Me
    UpdateEditFields
    rs.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'tidy up
    'Clean-up procedure
    If Not rs Is Nothing Then
        'first, check if the state is open, if yes then close it
        If (rs.State And adStateOpen) = adStateOpen Then
            rs.Close
        End If
        'set them to nothing
        Set rs = Nothing
    End If
    'same comment with rs
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
End Sub

Private Sub grdPointTypes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    UpdateEditFields
End Sub

Sub UpdateEditFields()
      
    With Me.dtcPointTypes.Recordset
        Me.txtDescription = .Fields("omschrijving")
'        Do While Not Me.cmbCategories.ItemData
        Me.cmbCategories.BoundText = .Fields("categoryId")
        Me.UpDnOrder.value = .Fields("volgorde")
    End With
End Sub


