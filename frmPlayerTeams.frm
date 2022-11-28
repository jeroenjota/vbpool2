VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPlayerTeams 
   Caption         =   "Spelers"
   ClientHeight    =   9345
   ClientLeft      =   12540
   ClientTop       =   3435
   ClientWidth     =   3570
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
   ScaleHeight     =   9345
   ScaleWidth      =   3570
   Begin VB.ComboBox cmbTeams 
      Height          =   360
      Left            =   960
      TabIndex        =   5
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "Nieuw"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   8040
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstPlayers 
      Height          =   6855
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   12091
      View            =   2
      Arrange         =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblSelectCount 
      Caption         =   "Geselecteerd: "
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   7800
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spelers"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Tag             =   "kop"
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Team"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmPlayerTeams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'to preserve the tournamentTeamCode
Dim cn As ADODB.Connection

Dim rs As ADODB.Recordset

Private Sub btnNew_Click()
    'add player to database
    frmPlayersNew.Country = getTeamInfo(Me.cmbTeams.ItemData(Me.cmbTeams.ListIndex), "teamCountryId", cn)
    frmPlayersNew.Show 1
    DoEvents
    'rs.Requery
    updateListview
End Sub

Private Sub btnOk_Click()
Unload Me
End Sub


Private Sub cmbTeams_Click()
    updateListview
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "Select * from tblTeamNames Where teamtype <> 0  and teamNameId IN "
    sqlstr = sqlstr & " (Select teamid from tblTournamentTeamCodes where tournamentid = " & thisTournament
    sqlstr = sqlstr & " ) Order by teamName"
    
    'fill teams combo
    FillCombo Me.cmbTeams, sqlstr, cn, "teamName", "teamNameid"
    Me.cmbTeams.ListIndex = 0
    UnifyForm Me
'    centerForm Me
    updateListview
    
End Sub

Sub updateListview()
    Dim rsPlayers As ADODB.Recordset
    Dim selCount As Integer
    selCount = 0
    Dim lItem As ListItem
    Dim sqlstr As String
    
    Set rsPlayers = New ADODB.Recordset
    
    'get the tournament teamcode for this team
    thisTeam = Me.cmbTeams.ItemData(Me.cmbTeams.ListIndex)
    
    sqlstr = "Select nickname, peopleID from tblPeople "
    sqlstr = sqlstr & " Where countryCode = " & nz(getTeamInfo(thisTeam, "teamCountryId", cn), 0)
    sqlstr = sqlstr & " and functionID > 1 and functionID <> 6"
    'sqlstr = sqlstr & " and active = -1"
    sqlstr = sqlstr & " Order by nickname"
    rsPlayers.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    
    With Me.lstPlayers
      .ListItems.Clear
      DoEvents
      .ColumnHeaders.Clear
      .ColumnHeaders.Add , , "Bijnaam", Me.lstPlayers.width - 60
      .ColumnHeaders.Add , , "ID", 0
      .View = lvwReport
      .Checkboxes = True
      .Sorted = True
      .SortKey = 0
      Do While Not rsPlayers.EOF
        Set lItem = .ListItems.Add(.ListItems.Count + 1)
        lItem.Text = rsPlayers!nickName
        lItem.Checked = playerInTournamentTeam(rsPlayers!peopleID, thisTeam, cn)
        If lItem.Checked Then selCount = selCount + 1
        lItem.SubItems(1) = nz(rsPlayers!peopleID, "")
        rsPlayers.MoveNext
      Loop
      .Refresh
    End With
    Me.lblSelectCount = "Geselecteerd: " & selCount
    If (rsPlayers.State And adStateOpen) = adStateOpen Then rsPlayers.Close
    Set rsPlayers = Nothing

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

Private Sub lstPlayers_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'add / remove player from tournament team
    Dim sqlstr As String
    If Item.Checked Then
        sqlstr = "Insert into tblTeamPlayers (tournamentId, teamId, playerId) "
        sqlstr = sqlstr & "VALUES (" & thisTournament
        sqlstr = sqlstr & ", " & thisTeam
        sqlstr = sqlstr & ", " & val(Item.SubItems(1))
        sqlstr = sqlstr & ")"
    Else
        sqlstr = "Delete from tblTeamPlayers where tournamentId = " & thisTournament
        sqlstr = sqlstr & " AND teamID = " & thisTeam
        sqlstr = sqlstr & " AND playerId = " & val(Item.SubItems(1))
    End If
    cn.Execute sqlstr
End Sub
