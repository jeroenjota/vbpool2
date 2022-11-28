VERSION 5.00
Object = "{3E5D9624-07F7-4D22-90F8-1314327F7BAC}#1.0#0"; "VBFLXGRD14.OCX"
Begin VB.Form frmPlayers 
   Caption         =   "Spelers aanpassen"
   ClientHeight    =   6180
   ClientLeft      =   16335
   ClientTop       =   6420
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   9675
   Begin VB.CommandButton Command1 
      Caption         =   "&Alles"
      Height          =   375
      Left            =   5760
      TabIndex        =   16
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "&Zoek"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Top             =   5640
      Width           =   3135
   End
   Begin VB.PictureBox picDummy 
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   2715
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox cmbPlayers 
      Height          =   360
      Left            =   3000
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VBFLXGRD14.VBFlexGrid grdPlayers 
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7223
      ScrollBars      =   2
      HighLight       =   0
      FocusRect       =   2
   End
   Begin VB.CheckBox chkActive 
      Alignment       =   1  'Right Justify
      Caption         =   "Actief"
      Height          =   240
      Left            =   8040
      TabIndex        =   7
      Top             =   600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Sluiten"
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Opslaan"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtfield 
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox txtfield 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   3000
   End
   Begin VB.TextBox txtfield 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voornaam"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voornaam"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecteer"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Voetballers"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Tag             =   "kop"
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voornaam"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private currentCountry As Long
Dim currentID As Integer

Public Property Get country() As Long
    country = currentCountry
End Property

Public Property Let country(ByVal NewValue As Long)
    currentCountry = NewValue
End Property

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
Dim i As Integer
  With rs
    If currentID = 0 Then Exit Sub
    .Find "id = " & currentID
    If Not .EOF Then
      For i = 1 To rs.Fields.Count - 1
        .Fields(i) = Me.txtField(i - 1)
        Me.txtField(i - 1).Enabled = False
        Me.txtField(i - 1).Text = ""
      Next
      .Update
    End If
   
  End With
  
  DoEvents
  'editMode = False
  Me.btnSave.Enabled = False
  Me.grdPlayers.Enabled = True
  initForm
  With Me.grdPlayers
    For i = 0 To Me.grdPlayers.rows - 1
      If .TextMatrix(i + 1, 0) = currentID Then
        Exit For
      End If
    Next
    .row = i + 1
    .TopRow = i + 1
  End With

End Sub

Private Sub btnSearch_Click()
Dim srchTxt As String
Dim filterTxt As String
Dim i As Integer
srchTxt = Trim(Me.txtSearch)
If srchTxt > "" Then
  'edit velden leegmaken
  For i = 0 To 2
    Me.txtField(i) = ""
    Me.txtField(i).Enabled = False
  Next
  Me.btnSave.Enabled = False
  filterTxt = "bijnaam LIKE *" & srchTxt & "*"
  filterTxt = filterTxt & " OR voornaam LIKE *" & srchTxt & "*"
  filterTxt = filterTxt & " OR achternaam LIKE *" & srchTxt & "*"
  On Error GoTo errHandler
  rs.Filter = filterTxt
Else
  rs.Filter = ""
End If
endSub:
  initForm
  Exit Sub
errHandler:
  MsgBox "Zoektekst niet correct"
  rs.Filter = ""
  Resume endSub
End Sub

Private Sub Command1_Click()
  Me.txtSearch = ""
  rs.Filter = ""
  initForm
End Sub

Private Sub Form_Load()
Dim sqlstr As String
Dim i As Integer
Dim thisCountry
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    sqlstr = "Select peopleid as ID, firstname as Voornaam, "
    sqlstr = sqlstr & "lastname as Achternaam, "
    sqlstr = sqlstr & "nickname as bijnaam "
    sqlstr = sqlstr & "from tblPeople"
    sqlstr = sqlstr & " ORDER BY nickname"
    Set rs = New ADODB.Recordset
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    If rs.RecordCount Then
      thisPlayer = rs!id
    End If
    initForm
    UnifyForm Me
    centerForm Me
End Sub

Sub initForm()
Dim sqlstr As String
Dim i As Integer
Dim j As Integer
Dim grdWidth As Integer
Dim savColWidth()

'    sqlstr = "Select * from tblCountries ORDER BY countryName"
'    FillCombo Me.cmbCountry, sqlstr, cn, "countryname", "countryid"
'    sqlstr = "Select * from tblPeople order by nickname"
'    FillCombo Me.cmbPlayers, sqlstr, cn, "nickname", "peopleid"
    'fill the grid
    Me.btnSave.Enabled = False
    Me.grdPlayers.Enabled = True
    i = 0
    With Me.grdPlayers
      .Redraw = False
      .Clear
      .cols = rs.Fields.Count
      ReDim savColWidth(.cols)
      'kolom namen
      For j = 0 To rs.Fields.Count - 1
        If Not IsNull(rs.Fields(j).Name) Then
          .TextMatrix(i, j) = rs.Fields(j).Name
          .ColAlignment(j) = FlexAlignmentLeftCenter
          If j > 0 Then Me.Label1(j - 1).Caption = rs.Fields(j).Name
        End If
      Next
      If rs.RecordCount = 0 Then Exit Sub
      .rows = rs.RecordCount + 1
      rs.MoveFirst
      Me.picDummy.Font.Name = Me.grdPlayers.Font.Name
      Me.picDummy.Font.Size = Me.grdPlayers.Font.Size
      Me.picDummy.ScaleMode = vbTwips
      Do While Not rs.EOF
        i = i + 1
        For j = 0 To rs.Fields.Count - 1
          If Not IsNull(rs.Fields(j).value) Then
            .TextMatrix(i, j) = rs.Fields(j).value
          Else
            .TextMatrix(i, j) = ""
          End If
          If j = 0 Then
            .colWidth(j) = 0
          Else
            If Me.picDummy.TextWidth(.TextMatrix(i, j) & "XXX") > savColWidth(j) Then
              savColWidth(j) = Me.picDummy.TextWidth(.TextMatrix(i, j) & "XXX")
            End If
          End If
          DoEvents
        Next
        DoEvents
        rs.MoveNext
      Loop
      Me.Label1(0).Left = Me.grdPlayers.Left
      Me.txtField(0).Left = Me.grdPlayers.Left
      For j = 1 To rs.Fields.Count - 1
        .colWidth(j) = Me.txtField(j - 1).width
        Me.Label1(j - 1).width = Me.txtField(j - 1).width
'        Me.Label1(j - 1).width = savColWidth(j)
'        Me.txtfield(j - 1).width = savColWidth(j)
        If j = .cols - 1 Then
'          Me.cmbCountry.width = savColWidth(j)
        End If
        If j > 1 Then
          Me.Label1(j - 1).Left = Me.txtField(j - 2).Left + Me.txtField(j - 2).width
          Me.txtField(j - 1).Left = Me.txtField(j - 2).Left + Me.txtField(j - 2).width
        End If
      Next
      Me.grdPlayers.width = Me.txtField(rs.Fields.Count - 2).Left + Me.txtField(rs.Fields.Count - 2).width + 240
      Me.btnClose.Left = Me.grdPlayers.width - Me.btnClose.width - 260
      'Me.btnNew.Left = Me.grdPlayers.width - Me.btnNew.width - 240
      'Me.btnDelete.Left = Me.btnClose.Left - Me.btnDelete.width - 20
      Me.btnSave.Left = Me.grdPlayers.width - Me.btnSave.width - 240
      'Me.lblTitle.width = Me.grdPlayers.width
      Me.width = Me.grdPlayers.width + 240
      .Redraw = True
  End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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

Private Sub grdPlayers_RowColChange()
  Dim i As Integer
  rs.MoveFirst
  currentID = val(Me.grdPlayers.TextMatrix(Me.grdPlayers.row, 0))
  rs.Find "id = " & currentID
  With rs
    If Not .EOF Then
      For i = 1 To .Fields.Count - 1
        Me.txtField(i - 1) = nz(.Fields(i), "")
        Me.txtField(i - 1).Enabled = True
      Next
    End If
  End With
  Me.btnSave.Enabled = True
'  Me.grdPlayers.Enabled = False
  'editMode = True
  'thisAddress = currentID

End Sub

Private Sub lblTitle_Click()
initForm
End Sub
