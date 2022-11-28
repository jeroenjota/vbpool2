VERSION 5.00
Begin VB.Form frmTournamentGroups 
   Caption         =   "Ploegen en spelers"
   ClientHeight    =   1845
   ClientLeft      =   12030
   ClientTop       =   3765
   ClientWidth     =   5055
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
   ScaleHeight     =   1845
   ScaleWidth      =   5055
   Begin VB.ComboBox cmbTeams 
      Height          =   360
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton btnPlayers 
      Caption         =   "Spelers"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblPoolName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pool A"
      Height          =   360
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Tag             =   "kop"
      Top             =   240
      Width           =   2100
   End
   Begin VB.Label lblPoolNr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   600
      Width           =   200
   End
End
Attribute VB_Name = "frmTournamentGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection

Sub makeForm()
Dim adoCmd As ADODB.Command

Dim rsTeamNames As ADODB.Recordset
Dim rsTeamCodes As ADODB.Recordset

Dim i As Integer
Dim sqlstr As String
Dim groups As Integer
Dim teams As Integer
Dim row As Integer, col As Integer
Dim grpCounter As Integer
Dim counter As Integer
Dim groupSize As Integer
Dim grp As Integer
'Set rsTeamNames = ADODB.Recordset

    Set adoCmd = New ADODB.Command
    
    'fill combobox with teamnames
    sqlstr = "Select teamNameId, TeamName, teamShortname, teamType from tblTeamNames"
    If getTournamentInfo("tournamentType", cn) = "EK" Then
        sqlstr = sqlstr & " Where teamtype <= 1"
    End If
    If getTournamentInfo("tournamentType", cn) = "CL" Then
        sqlstr = sqlstr & " Where teamtype > 2"
    End If
'    rsTeamNames.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        Set rsTeamNames = .Execute
        .CommandText = "Select * from tblTournamentTeamCodes where tournamentid = " & thisTournament
        .ActiveConnection.CursorLocation = adUseClient
        
        Set rsTeamCodes = .Execute
    End With
    If rsTeamNames.EOF Then Exit Sub
    
    groups = nz(getTournamentInfo("tournamentGroupCount", cn), 0)
    teams = getTournamentInfo("tournamentTeamCount", cn)
    groupSize = teams / groups
    counter = 0
    grp = 0
    For row = 1 To groups / 2
        For col = 1 To 2
            grp = grp + 1
            If lblPoolName.Count < grp Then
                Load lblPoolName(grp - 1)
                lblPoolName(grp - 1).Visible = True
            End If
            With Me.lblPoolName(grp - 1)
                .Caption = "Pool " & Chr(64 + grp)
                .width = 2100
                .Height = 360
                .Top = 240 + (row - 1) * (.Height + 100 + (groupSize * 360))
                .Left = 300 + (col - 1) * 2200
            End With
            For grpCounter = 1 To groupSize
                counter = counter + 1
                If lblPoolNr.Count < counter Then
                    Load Me.lblPoolNr(counter - 1)
                    Load Me.cmbTeams(counter - 1)
                    Me.lblPoolNr(counter - 1).Visible = True
                    Me.cmbTeams(counter - 1).Visible = True
                    Me.cmbTeams(counter - 1).TabIndex = counter
                End If
                With Me.lblPoolNr(counter - 1)
                    .Caption = grpCounter
                    .Left = 300 + (col - 1) * 2200
                    .Top = 600 + (grpCounter - 1) * 360 + (row - 1) * (.Height + 100 + (groupSize * 360))
                End With
                FillCombo Me.cmbTeams(counter - 1), sqlstr, cn, "teamname", "teamNameId"
                With Me.cmbTeams(counter - 1)
                    .Left = 600 + (col - 1) * 2200
                    .width = 1800
                    .Top = Me.lblPoolNr(counter - 1).Top
                    'Add tag to find record later in tblTournamentTeamCodes
                    .Tag = Chr(64 + grp) & Format(grpCounter, "0")
                    'if teamId exists in table, show team
                    rsTeamCodes.MoveFirst
                    rsTeamCodes.Find "teamcode = '" & .Tag & "'"
                    setCombo Me.cmbTeams(counter - 1), rsTeamCodes!teamID
'                    For i = 0 To .ListCount - 1
'                        If .ItemData(i) = rsTeamCodes!teamId Then
'                            .ListIndex = i
'                            Exit For
'                        End If
'                    Next
                End With
            Next
        Next
        Me.Height = (Me.Height - Me.ScaleHeight) + 640 + row * (groupSize + 1) * 360
    Next
    'ruimte voor knoppen
    Me.Height = Me.Height + Me.btnClose.Height + 240
    Me.btnClose.Top = Me.ScaleHeight - Me.btnClose.Height - 180
    Me.btnPlayers.Top = Me.btnClose.Top
    Me.btnPlayers.Left = Me.btnClose.Left - Me.btnPlayers.width - 50
End Sub

Private Sub btnClose_Click()
    Dim ctl As Control
    For Each ctl In lblPoolName
        If ctl.Index <> 0 Then
            Unload ctl
        End If
    Next
    For Each ctl In cmbTeams
        If ctl.Index <> 0 Then
            Unload ctl
        End If
    Next
    For Each ctl In lblPoolNr
        If ctl.Index <> 0 Then
            Unload ctl
        End If
    Next
    Set ctl = Nothing
    Unload Me
End Sub

Private Sub btnPlayers_Click()
    frmPlayerTeams.Show 1
End Sub

Private Sub cmbTeams_LostFocus(Index As Integer)
Dim sqlstr As String
Dim cmd As New ADODB.Command
    If Me.cmbTeams(Index).Text = "" Then Exit Sub
    'find and update the record based on the tag of the control
    sqlstr = "Update tblTournamentTeamCodes Set teamId = " & Me.cmbTeams(Index).ItemData(Me.cmbTeams(Index).ListIndex)
    sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND teamcode = '" & Me.cmbTeams(Index).Tag & "'"
    On Error GoTo dataerror
    cn.BeginTrans
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Execute
    End With
    cn.CommitTrans
    Exit Sub
dataerror:
    cn.RollbackTrans
    
End Sub

Private Sub Form_Activate()
Dim i As Integer
    For i = Me.cmbTeams.Count - 1 To 0 Step -1
        Me.cmbTeams(i).SetFocus
    Next
    Me.btnPlayers.SetFocus
End Sub

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    
    makeForm
'    Me.cmbTeams().SetFocus
    centerForm Me
    UnifyForm Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
End Sub
