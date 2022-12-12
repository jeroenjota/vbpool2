VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3E5D9624-07F7-4D22-90F8-1314327F7BAC}#1.0#0"; "VBFLXGRD14.OCX"
Begin VB.Form frmMatchEvents 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wedstrijd data"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6780
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.UpDown updnMinuut 
      Height          =   360
      Left            =   1020
      TabIndex        =   21
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   635
      _Version        =   393216
      BuddyControl    =   "txtMinuut"
      BuddyDispid     =   196619
      OrigLeft        =   1200
      OrigTop         =   600
      OrigRight       =   1455
      OrigBottom      =   975
      Max             =   121
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.PictureBox picPenalties 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00008000&
      ForeColor       =   &H0000FFFF&
      Height          =   1860
      Left            =   1800
      ScaleHeight     =   1800
      ScaleWidth      =   3210
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   3270
      Begin MSComCtl2.UpDown UpDnPenals 
         Height          =   360
         Index           =   1
         Left            =   2535
         TabIndex        =   20
         Top             =   900
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   635
         _Version        =   393216
         BuddyControl    =   "txtPenals(1)"
         BuddyDispid     =   196610
         BuddyIndex      =   1
         OrigLeft        =   2535
         OrigTop         =   900
         OrigRight       =   2790
         OrigBottom      =   1260
         Max             =   50
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDnPenals 
         Height          =   360
         Index           =   0
         Left            =   2535
         TabIndex        =   19
         Top             =   510
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   635
         _Version        =   393216
         BuddyControl    =   "txtPenals(0)"
         BuddyDispid     =   196610
         BuddyIndex      =   0
         OrigLeft        =   2535
         OrigTop         =   510
         OrigRight       =   2790
         OrigBottom      =   870
         Max             =   50
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtPenals 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   0
         Left            =   2235
         TabIndex        =   15
         Top             =   480
         Width           =   300
      End
      Begin VB.TextBox txtPenals 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   1
         Left            =   2235
         TabIndex        =   14
         Top             =   900
         Width           =   300
      End
      Begin VB.CommandButton btnClosePenalties 
         Caption         =   "Klaar"
         Height          =   300
         Left            =   2040
         TabIndex        =   13
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Strafschoppen"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   450
         TabIndex        =   18
         Top             =   75
         Width           =   2220
      End
      Begin VB.Label lblPenalTeam 
         BackStyle       =   0  'Transparent
         Caption         =   "Team1"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   17
         Top             =   555
         Width           =   2130
      End
      Begin VB.Label lblPenalTeam 
         BackStyle       =   0  'Transparent
         Caption         =   "Team1"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   16
         Top             =   945
         Width           =   2130
      End
   End
   Begin VB.CommandButton btnMatchEnd 
      Caption         =   "Einde wedstrijd"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton btnEventsClear 
      Caption         =   "Wissen"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4800
      Width           =   1020
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Annuleren"
      Height          =   375
      Left            =   3510
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VBFLXGRD14.VBFlexGrid grdMatchEvents 
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5318
      FixedCols       =   0
      AllowBigSelection=   0   'False
      RowSizingMode   =   1
      SelectionMode   =   1
      ScrollBars      =   2
      FocusRect       =   2
      SingleLine      =   -1  'True
   End
   Begin VB.ComboBox cmbPlayer 
      Height          =   360
      Left            =   3120
      TabIndex        =   5
      Top             =   600
      Width           =   2895
   End
   Begin VB.ComboBox cmbEvent 
      Height          =   360
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtMinuut 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   420
   End
   Begin VB.CommandButton btnEventAdd 
      Default         =   -1  'True
      Height          =   330
      Left            =   6120
      Picture         =   "frmMatchEvents.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Voeg gebeurtenis toe"
      Top             =   600
      Width           =   450
   End
   Begin VB.CommandButton btnEventDelete 
      Height          =   330
      Left            =   120
      Picture         =   "frmMatchEvents.frx":00EA
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Wis gebeurtenis"
      Top             =   600
      Width           =   450
   End
   Begin VB.Label lblStand 
      BackStyle       =   0  'Transparent
      Caption         =   "Stand:"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   120
      TabIndex        =   8
      Top             =   4215
      Width           =   2505
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblOther 
      BackStyle       =   0  'Transparent
      Caption         =   "Penalties:"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   2805
      TabIndex        =   7
      Top             =   4200
      Width           =   2925
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Wedstrijd"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmMatchEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Dim matchDescr As String

Dim tc(1) As Integer 'teamCodes
Dim team(1) As String 'teamNames

Dim ht(1) As Integer  'goals at halftime
Dim ft(1) As Integer  'goals at fulltime
Dim xt(1) As Integer  'goals in extra time
Dim pns(1) As Integer 'penalties in shootout

Private Sub btnCancel_Click()
  write2Log "Wedstrijd bijhouden gesloten " & getMatchDescription(thisMatch, cn, , , True), True
  Unload Me
End Sub

Private Sub btnClosePenalties_Click()
  Dim msg As String
  'must be different numbers
  If Me.UpDnPenals(0) = Me.UpDnPenals(1) Then
    msg = "Er moet een winnaar zijn!"
    msg = msg & vbNewLine & "Mocht er geloot zijn of zo, geef dan het winnende team een extra penalty"
    MsgBox msg, vbInformation + vbOKOnly, "Geen winnaar?"
    Exit Sub
  Else
    pns(0) = Me.UpDnPenals(0)
    pns(1) = Me.UpDnPenals(1)
  End If
  
  Me.picPenalties.Visible = False
  fillGrid
End Sub

Private Sub btnEventAdd_Click()
  Dim sqlstr As String
  Dim teamID As Integer
  If Me.updnMinuut.value > 0 And Me.cmbEvent.ListIndex <> -1 And (Me.cmbPlayer.ListIndex <> -1 Or Me.cmbEvent.ItemData(Me.cmbEvent.ListIndex) = 7) Then
    'delete possible exiting event
    sqlstr = "Delete from tblMatchEvents WHERE tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND matchOrder = " & thisMatch
    sqlstr = sqlstr & " AND minuut = " & Me.updnMinuut
    cn.Execute sqlstr
    'add new event
    sqlstr = "INSERT INTO tblMatchEvents (tournamentID, matchOrder, minuut, eventid, playerid )"
    sqlstr = sqlstr & " VALUES ( " & thisTournament
    sqlstr = sqlstr & ", " & thisMatch
    sqlstr = sqlstr & ", " & Me.updnMinuut
    sqlstr = sqlstr & ", " & Me.cmbEvent.ItemData(Me.cmbEvent.ListIndex)
 '   If Me.cmbEvent.ItemData(Me.cmbEvent.ListIndex) <> 7 Then
      sqlstr = sqlstr & ", " & Me.cmbPlayer.ItemData(Me.cmbPlayer.ListIndex)
 '   Else
 '     sqlstr = sqlstr & ", 5360"  '5360 is een fictieve speler
 '   End If
    sqlstr = sqlstr & ") "
    cn.Execute sqlstr
    fillGrid
    If Me.cmbEvent.ItemData(Me.cmbEvent.ListIndex) = 7 Then 'penalty shootout
      Me.picPenalties.Visible = True
    End If
    write2Log "Gebeurtenis vastgelegd: wedstrijd " & thisMatch & ", minuut " & Me.updnMinuut
  End If
  Me.txtMinuut.SetFocus
End Sub

Private Sub btnEventDelete_Click()
Dim sqlstr As String
  sqlstr = "Delete from tblMatchEvents WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder = " & thisMatch
  sqlstr = sqlstr & " AND minuut = " & Me.updnMinuut
  cn.Execute sqlstr
  If Me.cmbEvent.ItemData(Me.cmbEvent.ListIndex) = 7 Then  'clear the penalty shootout values
    Me.UpDnPenals(0) = 0
    Me.UpDnPenals(1) = 0
  End If
  write2Log "Gebeurtenis weggehaald: wedstrijd " & thisMatch & ", minuut " & Me.updnMinuut
  fillGrid
End Sub

Private Sub btnEventsClear_Click()
Dim sqlstr As String
  sqlstr = "Delete from tblMatchEvents WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder = " & thisMatch
  cn.Execute sqlstr
  sqlstr = "Delete from tblMatchResults WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder = " & thisMatch
  cn.Execute sqlstr
  setMatchPlayed thisMatch, False, cn
  Me.updnMinuut = 0
  Me.UpDnPenals(0) = 0
  Me.UpDnPenals(1) = 0
  Me.cmbEvent.ListIndex = -1
  Me.cmbPlayer.ListIndex = -1
  write2Log "Alle gebeurtenissen verwijderd: wedstrijd " & thisMatch & ", minuut " & Me.updnMinuut

  fillGrid
  updateSituation
End Sub

Private Sub btnMatchEnd_Click()
Dim processMatch As Boolean
Dim msg As String
Dim sqlstr As String
Dim answ As Integer
Dim i As Integer
Dim grp As String
Dim teams() As Long
Dim winner As Long
Dim thisTeam As Long
'check if match was already played
  processMatch = Not matchPlayed(thisMatch, cn)
  If Not processMatch Then 'ask if match should be re-processed
    msg = "Uitslag opnieuw doorrekenen?"
    answ = MsgBox(msg, vbYesNoCancel + vbQuestion, "Wedstrijd al ingevoerd")
    Select Case answ
    Case vbYes
      processMatch = True
    Case vbNo
      processMatch = False
    Case Else
    End Select
  End If
  If processMatch Then
    If Me.grdMatchEvents.rows <= 1 Then
        msg = "Niks gebeurd in deze wedstrijd? " & vbNewLine & "Toch markeren als gespeeld? (Ja) of Wissen (Nee)"
        processMatch = MsgBox(msg, vbYesNo + vbQuestion, "Wedstrijd sluiten")
        For i = 0 To 1
          ht(i) = 0
          ft(i) = 0
          xt(i) = 0
          pns(i) = 0
        Next
    End If
    If getMatchInfo(thisMatch, "matchtype", cn) <> 1 And Me.grdMatchEvents.rows <= 1 Then
        'match has to have a winner, not a group match
        processMatch = False
        msg = "Er moet een winnaar uit deze wedstrijd komen"
        msg = msg & vbNewLine & "Wedstrijd wordt niet verwerkt, invoerscherm sluiten?"
        answ = MsgBox(msg, vbYesNo, "Uitslag kan niet")
        
        'remove old result
        sqlstr = "DELETE from tblMatchResults where matchOrder = " & thisMatch
        sqlstr = sqlstr & " AND tournamentid = " & thisTournament
        cn.Execute sqlstr
        
    End If
  End If
  If processMatch Then
    'process the scores and set the team for the next round
    setMatchPlayed thisMatch, True, cn
    If Not processMatchScores Then Exit Sub
    'set the groupstandings if matchtype == 1
    If getMatchInfo(thisMatch, "matchType", cn) = 1 Then
      'set the group standings
      calcGroupStandings cn
      grp = getMatchGroup(thisMatch, cn)
      If grpPlayedCount(grp, cn) = 6 Then
        msg = "Dit was de laatste wedstrijd van groep " & grp
        msg = msg & vbNewLine & "Controleer in het volgende scherm of de posities kloppen"
        msg = msg & vbNewLine & "en pas die posities eventueel aan om te bepalen wie er naar de volgende ronde gaat."
        msg = msg & vbNewLine & "De regels zijn nogal ingewikkeld bij gelijke stand"
        msg = msg & vbNewLine & vbNewLine & "(Welke derde plaatsen doorgaan bepalen we na de laatste groepswedstrijd)"
        
        MsgBox msg, vbOKOnly + vbInformation, "Groep uitgespeeld"
        'show the form with the group standings to be able to adjust positions
        
        frmGroupStands.Show 1
        'set the teams on positions 1 and 2 through to the finals
        Set8Finals cn
        'if all groupmatches are played and there are 6 groups then ask for best 3rd places
        If getLastMatchPlayed(cn) = getMatchCount(1, cn) And nz(getTournamentInfo("tournamentGroupCount", cn), 0) = 6 Then
          frm8Finals.Show 1
        End If
      End If
      
    Else
      'set teams in next round
      winner = getMatchresult(thisMatch, 6, cn)
      teams = getMatchTeamIDs(thisMatch, cn)
      'if there is a third place AND this is a semi final set the loser
      If getTournamentInfo("tournamentThirdPlace", cn) And getMatchInfo(thisMatch, "matchtype", cn) = 3 Then  'halve finale verliezers bepalen voor derde plaats (als aanwezig)
        'thisTeam is een van de teams in de kleine finale
        If teams(0) = winner Then
          thisTeam = teams(1)
        Else
          thisTeam = teams(0)
        End If
        sqlstr = "UPDATE tblTournamentTeamCodes SET teamID = " & thisTeam
        sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
        sqlstr = sqlstr & " AND teamCode = 'V" & Format(thisMatch, "0") & "'"
        cn.Execute sqlstr
      End If
      'set the winners through to the next round
      sqlstr = "UPDATE tblTournamentTeamCodes SET teamID = " & winner
      sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
      sqlstr = sqlstr & " AND teamCode = 'W" & Format(getMatchNumber(thisMatch, cn), "0") & "'"
      cn.Execute sqlstr
    End If
    'calculate poolform scores
    Me.Visible = False
    showInfo True, "Ogenblikje ..."
    write2Log "Wedstrijd ogeslagen " & getMatchDescription(thisMatch, cn, , , True), True
    'bereken de punten voor de deelnemers
    updatePoolFormPoints thisMatch, cn
    'en zet de posities in de stand
    updatePoolPositions thisMatch, cn
    showInfo False
  End If
  
'close form
  Unload Me
End Sub

Private Sub Form_Load()
  'open the database
  Set cn = New ADODB.Connection
  With cn
      .ConnectionString = lclConn()
      .Open
  End With
  pns(0) = nz(getMatchresult(thisMatch, 11, cn), 0)
  pns(1) = nz(getMatchresult(thisMatch, 12, cn), 0)
  Me.UpDnPenals(0) = pns(0)
  Me.UpDnPenals(1) = pns(1)
  write2Log "Wedstrijd bijhouden " & getMatchDescription(thisMatch, cn, , , True), True
  initForm
  UnifyForm Me
  centerForm Me
End Sub

Sub initForm()
'fill the combo boxes
  Dim i As Integer
  Dim sqlstr As String
  Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblTournamentSchedule  where tournamentID = " & thisTournament
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  rs.Find "matchOrder = " & thisMatch
  If rs.EOF Then
    MsgBox "ERRRORRR", vbOKOnly + vbExclamation, "ERROR"
    Exit Sub
  End If
  matchDescr = rs!matchOrder & "e wed (nr " & rs!matchNumber
  matchDescr = matchDescr & "): " & getScheduleTeamName(rs!matchteamA, cn)
  matchDescr = matchDescr & " - " & getScheduleTeamName(rs!matchteamB, cn)
  matchDescr = matchDescr & "  " & Format(rs!matchDate, "d MMM") & " om " & Format(rs!matchtime, "HH:NN")
  Me.lblTitle = matchDescr
  rs.Close
  sqlstr = "Select * from tblEvents"
  FillCombo Me.cmbEvent, sqlstr, cn, "eventDescription", "eventID"
  'select the players from both teams
  tc(0) = getTeamIdFromCode(getMatchInfo(thisMatch, "matchTeamA", cn), cn)
  tc(1) = getTeamIdFromCode(getMatchInfo(thisMatch, "matchTeamB", cn), cn)
  team(0) = getTeamInfo(tc(0), "teamName", cn)
  team(1) = getTeamInfo(tc(1), "teamName", cn)
  sqlstr = "Select tp.playerID as ID, p.nickName as speler"
  sqlstr = sqlstr & " FROM tblTeamPlayers tp INNER JOIN tblPeople p ON tp.playerID = p.peopleID"
  sqlstr = sqlstr & " WHERE tp.Tournamentid = " & thisTournament
  sqlstr = sqlstr & " AND (tp.TeamID = " & tc(0)
  sqlstr = sqlstr & " OR tp.TeamID = " & tc(1) & ")"
  sqlstr = sqlstr & " ORDER by p.nickName"
  FillCombo Me.cmbPlayer, sqlstr, cn, "speler", "ID"
  Me.updnMinuut = 0
  'set labels in (hidden)  penalty shootout frame
  For i = 0 To 1
    Me.lblPenalTeam(i) = team(i)
  Next
  fillGrid
  
End Sub

Sub fillGrid()

  Dim sqlstr As String
  Dim i As Integer
  Dim J As Integer
  Dim goals(1) As Integer
  Dim penalties As Integer
  Dim cards(1) As Integer
  Set rs = New ADODB.Recordset
  sqlstr = "Select m.minuut as mn, e.eventDescription as Gebeurtenis, p.nickName as Speler, t.teamShortname as Team "
  sqlstr = sqlstr & " from (((tblMatchEvents m INNER JOIN tblEvents e ON m.eventID = e.eventID)"
  sqlstr = sqlstr & " INNER JOIN tblPeople p ON m.playerID = p.peopleID)"
  sqlstr = sqlstr & " INNER JOIN tblTeamPlayers tp ON (m.tournamentID = tp.TournamentID) AND (p.peopleID = tp.playerid))"
  sqlstr = sqlstr & " INNER JOIN tblTeamNames t ON tp.teamID = t.teamNameID"
  sqlstr = sqlstr & " WHERE m.tournamentid = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder = " & thisMatch
  sqlstr = sqlstr & " ORDER BY m.minuut"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  With Me.grdMatchEvents
    .Clear
    .rows = rs.RecordCount + 1
    .cols = rs.Fields.Count
    .colWidth(0) = 450
    .colWidth(1) = 2000
    .colWidth(2) = 2900
    .colWidth(3) = 900
    i = 0
    For J = 0 To rs.Fields.Count - 1
      If Not IsNull(rs.Fields(J).Name) Then
        .TextMatrix(i, J) = rs.Fields(J).Name
      End If
    Next
    If rs.EOF Then 'match not played yet or not active
      Exit Sub
    End If
    rs.MoveFirst
    Do While Not rs.EOF
      i = i + 1
      For J = 0 To rs.Fields.Count - 1
        If Not IsNull(rs.Fields(J).value) Then
          .TextMatrix(i, J) = rs.Fields(J).value
        Else
          .TextMatrix(i, J) = ""
        End If
      Next
      If rs!mn > 120 Then
        updateShootOut
      End If
      rs.MoveNext
      
    Loop
  End With
  Me.grdMatchEvents.row = Me.grdMatchEvents.rows - 1
  grdMatchEvents_Click
  rs.Close
  Set rs = Nothing
  updateSituation
End Sub

Sub updateShootOut()
'get the penaty shootout results
  'Me.UpDnPenals(0) = nz(getMatchresult(thisMatch, 11, cn), 0)
  'Me.UpDnPenals(1) = nz(getMatchresult(thisMatch, 12, cn), 0)
  'pns(0) = Me.UpDnPenals(0)
  'pns(1) = Me.UpDnPenals(1)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not rs Is Nothing Then
        If (rs.State And adStateOpen) = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Sub

Private Sub grdMatchEvents_Click()
'update the edit line
Dim i As Integer
  If Me.grdMatchEvents.row > 0 Then
    Me.updnMinuut = val(Me.grdMatchEvents.TextMatrix(Me.grdMatchEvents.row, 0))
    'update textvalue of comboboxes
    setCombo Me.cmbEvent, Me.grdMatchEvents.TextMatrix(Me.grdMatchEvents.row, 1)
    setCombo Me.cmbPlayer, Me.grdMatchEvents.TextMatrix(Me.grdMatchEvents.row, 2)
  End If

End Sub

Private Sub txtMinuut_Change()
  If val(Me.txtMinuut) > 121 Then Exit Sub
  Me.updnMinuut = val(Me.txtMinuut)
End Sub

Private Sub txtMinuut_GotFocus()
  SelectAllText Me.txtMinuut

End Sub

Private Sub txtMinuut_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
      KeyAscii = 0
  End If
End Sub

Private Sub txtPenals_KeyPress(Index As Integer, KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
      KeyAscii = 0
  End If
End Sub

Private Sub updnMinuut_Change()
  Me.cmbEvent.Enabled = Me.updnMinuut > 0
  Me.cmbPlayer.Enabled = Me.updnMinuut > 0
  Me.btnEventAdd.Enabled = Me.updnMinuut > 0
  Me.btnEventDelete.Enabled = Me.updnMinuut > 0
  Me.txtMinuut = Me.updnMinuut
  If Me.updnMinuut = 121 Then
    Me.cmbEvent.ListIndex = Me.cmbEvent.ListCount - 1
    Me.cmbPlayer.ListIndex = Me.cmbPlayer.ListCount - 1
    Me.cmbPlayer.Enabled = False
  End If
End Sub

Sub updateSituation()
Dim rs As ADODB.Recordset
Dim rsMatch As ADODB.Recordset
Dim teamNr As Integer
Dim pen As Integer 'penalties
Dim cards(1) As Integer '(0 = yelow, 1 = red)
Dim txtStr As String
Dim sqlstr As String
'  t1 = GetOrigToernooiTeam(t1)
'  t2 = GetOrigToernooiTeam(t2)

  Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblMatchEvents WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder = " & thisMatch
  rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  ft(0) = 0
  ft(1) = 0
  ht(0) = 0
  ht(1) = 0
  xt(0) = 0
  xt(1) = 0
  'pns(0) = 0
  'pns(1) = 0
  With rs
    Me.lblStand = ""
    Me.lblOther = ""
    If .EOF Then
      Me.updnMinuut = 0
      Me.cmbEvent.ListIndex = -1
      Me.cmbPlayer.ListIndex = -1
      Exit Sub
    End If
    Do While Not .EOF
      If !eventID = 1 Then pen = pen + 1
      If getPlayerTeam(!playerID, cn) = tc(0) Then
        teamNr = 0
      Else
        teamNr = 1
      End If
      Select Case !eventID
      Case 1, 2 'doelpunt, penalty
        If !Minuut <= 45 Then
            ht(teamNr) = ht(teamNr) + 1
        End If
        If !Minuut <= 90 Then
            ft(teamNr) = ft(teamNr) + 1
        Else
            xt(teamNr) = xt(teamNr) + 1
        End If
      Case 3 'eigen doelpunt
        If teamNr = 1 Then
          teamNr = 0
        Else
          teamNr = 1
        End If
        If !Minuut <= 45 Then
            ht(teamNr) = ht(teamNr) + 1
        End If
        If !Minuut <= 90 Then
            ft(teamNr) = ft(teamNr) + 1
        Else
            xt(teamNr) = xt(teamNr) + 1
        End If
      Case 4, 5
        cards(!eventID - 4) = cards(!eventID - 4) + 1
      Case 6
        pen = pen + 1
      Case 7
      End Select
      .MoveNext
    Loop
    .Close
    
    txtStr = "Stand: " & ft(0) & "-" & ft(1)
    If val(Me.grdMatchEvents.TextMatrix(Me.grdMatchEvents.rows - 1, 0)) > 45 Or matchPlayed(thisMatch, cn) Then
        txtStr = txtStr & " (" & ht(0) & "-" & ht(1) & ")"
    End If
    If val(Me.grdMatchEvents.TextMatrix(Me.grdMatchEvents.rows - 1, 0)) > 90 Then
        txtStr = txtStr & " na verl: " & ft(0) + xt(0) & "-" & ft(1) + xt(1)
    End If
    
    If pns(0) <> pns(1) Then
      If pns(0) > pns(1) Then
        txtStr = txtStr & vbNewLine & team(0)
      Else
        txtStr = txtStr & vbNewLine & team(1)
      End If
      txtStr = txtStr & " WNS (" & pns(0) & "-" & pns(1) & ")"
    End If
    Me.lblStand.Caption = txtStr
  End With
  Me.lblOther = "Penalties: " & pen
  Me.lblOther = Me.lblOther & vbNewLine & "Kaarten Geel: " & cards(0)
  Me.lblOther = Me.lblOther & "  Rood: " & cards(1)
End Sub

Function processMatchScores()
Dim sqlstr As String
Dim answ As Boolean
Dim winner As Integer, toto As Integer, wns As Boolean
  'remove old result
  'add new matchresult
  'get the winner
  If ft(0) > ft(1) Then
    toto = 1
    winner = tc(0)
  ElseIf ft(1) > ft(0) Then
    toto = 2
    winner = tc(1)
  Else
    toto = 3
    winner = 0
  End If
  If getMatchInfo(thisMatch, "matchType", cn) <> 1 Then
    'there has to be a winner
    If toto = 3 Then 'equal after 90 minutes
      If xt(0) > xt(1) Then 'goals in extra time
        winner = tc(0)
      ElseIf xt(1) > xt(0) Then
        winner = tc(1)
      Else 'equal after extra time, so penalty shootout
        wns = True
        If pns(0) > pns(1) Then
          winner = tc(0)
        ElseIf pns(1) > pns(0) Then
          winner = tc(1)
        Else
          winner = 0  ' cannot be!!!
          MsgBox "Kies strafschoppen serie bij gebeurtenissen. er moet een winnaar komen"
          Exit Function
        End If
      End If
    End If
  End If
  answ = True
  If answ Then
    'add match to the results table
    sqlstr = "DELETE from tblMatchResults where matchOrder = " & thisMatch
    sqlstr = sqlstr & " AND tournamentid = " & thisTournament
    cn.Execute sqlstr
    sqlstr = "INSERT INTO tblMatchResults VALUES(" & thisTournament
    sqlstr = sqlstr & ", " & thisMatch
    sqlstr = sqlstr & ", " & ht(0)
    sqlstr = sqlstr & ", " & ht(1)
    sqlstr = sqlstr & ", " & ft(0)
    sqlstr = sqlstr & ", " & ft(1)
    sqlstr = sqlstr & ", " & winner
    sqlstr = sqlstr & ", " & toto
    sqlstr = sqlstr & ", " & xt(0)
    sqlstr = sqlstr & ", " & xt(1)
    sqlstr = sqlstr & ", " & IIf(wns, -1, 0)
    sqlstr = sqlstr & ", " & pns(0)
    sqlstr = sqlstr & ", " & pns(1)
    sqlstr = sqlstr & ", " & tc(0)
    sqlstr = sqlstr & ", " & tc(1)
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
  End If
  processMatchScores = True
End Function

Private Sub UpDnPenals_Change(Index As Integer)
  Me.txtPenals(Index) = Me.UpDnPenals(Index)
  pns(Index) = Me.UpDnPenals(Index)
End Sub
