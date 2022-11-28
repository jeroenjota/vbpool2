VERSION 5.00
Object = "{3E5D9624-07F7-4D22-90F8-1314327F7BAC}#1.0#0"; "VBFLXGRD14.OCX"
Begin VB.Form frmMatchesOverview 
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMatchesOverview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleMode       =   0  'User
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnFollowMatch 
      Caption         =   "Volg de wedstrijd "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   9120
      Width           =   3975
   End
   Begin VBFLXGRD14.VBFlexGrid grdMatches 
      Height          =   2055
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3625
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      AllowSelection  =   0   'False
      SelectionMode   =   1
      ScrollBars      =   2
      HighLight       =   2
      FocusRect       =   2
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wedstrijden"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   12255
   End
End
Attribute VB_Name = "frmMatchesOverview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dontMove As Boolean 'prevent editBar from updateting

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Dim colWidths(9) As Double

Sub setMatchGrid()
Dim sqlstr As String
Dim dCol As Object
Dim col As Column
Dim i As Integer, J As Integer
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim lastMatch As Integer
lastMatch = getLastMatchPlayed(cn)
Dim formatStr As String
  sqlstr = "SELECT m.matchOrder as ID, format(m.matchDate,'dd-MM') as Datum, format(matchTime ,'HH:NN') as Tijd, "
  sqlstr = sqlstr & " a.teamcode as A, ta.teamName as Team1, b.teamcode as B, tb.teamName as Team2, "
  sqlstr = sqlstr & " t.matchTypeDescription as Type, s.stadiumName & '/' & s.stadiumLocation as Locatie, m.matchnumber as nr, m.matchplayed as pl "
  sqlstr = sqlstr & " FROM ((((tblTournamentSchedule m LEFT JOIN tblStadiums s ON m.matchStadiumID = s.stadiumID) "
  sqlstr = sqlstr & " LEFT JOIN tblTournamentTeamCodes AS b ON m.matchTeamB = b.teamCode) "
  sqlstr = sqlstr & " LEFT JOIN tblTeamNames AS tb ON b.teamID = tb.teamNameID) "
  sqlstr = sqlstr & " LEFT JOIN (tblTournamentTeamCodes a "
  sqlstr = sqlstr & " LEFT JOIN tblTeamNames ta ON a.teamID = ta.teamNameID) ON m.matchTeamA = a.teamCode) "
  sqlstr = sqlstr & " LEFT JOIN tblMatchTypes t ON m.matchType = t.matchTypeID"
  sqlstr = sqlstr & " WHERE m.tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND a.tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND b.tournamentID = " & thisTournament
  sqlstr = sqlstr & " ORDER BY m.matchOrder"
  rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
'fill the grid
  With Me.grdMatches
   .Left = 240
   .Top = Me.Label1.Height + Me.Label1.Top + 140
   .ZOrder
   .SelectionMode = flexSelectionByRow
   .Clear
    .rows = rs.RecordCount + 1
   .cols = rs.Fields.Count
   i = 0
   For J = 0 To rs.Fields.Count - 1
     If Not IsNull(rs.Fields(J).Name) Then
       .TextMatrix(i, J) = rs.Fields(J).Name
     End If
   Next
   rs.MoveFirst
   Do While Not rs.EOF
     i = i + 1
     For J = 0 To rs.Fields.Count - 1
       If Not IsNull(rs.Fields(J).value) Then
          If J <> 10 Then
          .TextMatrix(i, J) = rs.Fields(J).value
          ElseIf rs.Fields(J).value Then
            .TextMatrix(i, J) = "x"
          Else
            .TextMatrix(i, J) = ""
          End If
       Else
         .TextMatrix(i, J) = ""
       End If
     Next
     rs.MoveNext
   Loop
   .colWidth(0) = 400
   .ColAlignment(0) = flexAlignCenterCenter
   .colWidth(1) = 700
   .ColAlignment(1) = flexAlignCenterCenter
   .colWidth(2) = 600
   .ColAlignment(2) = flexAlignCenterCenter
   .colWidth(3) = 700
   .ColAlignment(3) = flexAlignCenterCenter
   .colWidth(4) = 1650
   .colWidth(5) = 700
   .ColAlignment(5) = flexAlignCenterCenter
   .colWidth(6) = 1650
   .colWidth(7) = 1900
   .colWidth(8) = 3050
   .colWidth(9) = 500
   .colWidth(10) = 300
   .ColAlignment(9) = flexAlignCenterCenter
   
  End With
  If lastMatch < getMatchCount(0, cn) Then
     Me.grdMatches.row = lastMatch + 1
  Else
     Me.grdMatches.row = Me.grdMatches.rows - 1
  End If
  grdMatches_Click
  'select entire row
  ' Me.grdMatches.MarqueeStyle = dbgHighlightRow
  'force update of editBar controls
  'grdMatches_RowColChange 1, 1
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnFollowMatch_Click()
  'addMatchResultRecord thisMatch, cn 'add this match to the tblMatchResults if not yet present
  frmMatchEvents.Show 1
  Unload Me
 ' Me.grdMatches.Refresh
  'Me.grdMatches.row = Me.grdMatches.row + 1
End Sub

Private Sub Form_Load()
    'open the database
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    
    setMatchGrid
'    Me.grdMatches.row = 2
'    Me.grdMatches.row = 1
'    setState 'only if admin is logged in is editting possible
    
    UnifyForm Me
    centerForm Me

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

Private Sub Form_Resize()
'attempt to resize everything with the form
'if to complicated we will fix the forms borders
Dim initHeigth As Integer, initWidth As Integer
Dim gridHeight As Integer, gridWidth As Integer
Dim col As Column
Dim i As Integer
Const leftPos = 240
'set the initial Height/Width
initHeigth = 10215
initWidth = 13000
With Me
  If .Height < initHeigth Then .Height = initHeigth
  If .width <> initWidth Then .width = initWidth
  gridHeight = .Height - Me.grdMatches.Top - Me.btnFollowMatch.Height - 800
  gridWidth = .width - 585
  .grdMatches.Height = gridHeight
  .grdMatches.width = gridWidth
  .btnClose.Top = .Height - 1195
  .btnClose.Left = .width - 2265
  .btnClose.Height = 520
  .btnFollowMatch.Height = .btnClose.Height
  .btnFollowMatch.Top = .btnClose.Top
  .btnFollowMatch.Left = .grdMatches.Left
  .btnFollowMatch.width = .grdMatches.width - .btnClose.width - 500
  i = 0
End With

End Sub


Private Sub grdMatches_Click()
  Dim sqlstr As String
  Dim i As Integer
  Dim searchTeam As String
  Dim matchDescr As String
  Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblTournamentSchedule  where tournamentID = " & thisTournament
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  rs.Find "matchOrder = " & Me.grdMatches.TextMatrix(Me.grdMatches.row, 0)
  thisMatch = rs!matchOrder
  matchDescr = "Volg de " & thisMatch & "e wedstrijd (nr: " & rs!matchNumber & ")"
  matchDescr = matchDescr & " op " & Format(rs!matchDate, "d MMMM") & " om " & Format(rs!matchtime, "HH:NN")
  With Me.grdMatches
    matchDescr = matchDescr & " tussen " & .TextMatrix(.row, 4)
    matchDescr = matchDescr & " en " & .TextMatrix(.row, 6)
    If .TopRow < .TextMatrix(.row, 0) - 25 Then
      .TopRow = .TextMatrix(.row, 0) - 25
    End If
    If .TopRow > .TextMatrix(.row, 0) Then
      .TopRow = .TextMatrix(.row, 0)
    End If
    Me.btnFollowMatch.Enabled = True
    If .row > 1 Then 'check if previous match was played and enable/disable button
      Me.btnFollowMatch.Enabled = matchPlayed(.TextMatrix(.row - 1, 0), cn)
    End If
  End With
  Me.btnFollowMatch.Caption = matchDescr
  

End Sub

