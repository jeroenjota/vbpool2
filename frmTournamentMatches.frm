VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3E5D9624-07F7-4D22-90F8-1314327F7BAC}#1.0#0"; "VBFLXGRD14.OCX"
Begin VB.Form frmTournamentMatches 
   Caption         =   "Wedstrijden"
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
   Icon            =   "frmTournamentMatches.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleMode       =   0  'User
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbLocation 
      Height          =   360
      Left            =   9360
      TabIndex        =   24
      Top             =   720
      Width           =   2055
   End
   Begin VB.ComboBox cmbTypes 
      Height          =   360
      Left            =   7560
      TabIndex        =   23
      Top             =   720
      Width           =   1815
   End
   Begin VB.ComboBox cmbTeamB 
      Height          =   360
      Left            =   5640
      TabIndex        =   22
      Top             =   720
      Width           =   1950
   End
   Begin VB.ComboBox cmbTeamA 
      Height          =   360
      Left            =   3720
      TabIndex        =   21
      Top             =   720
      Width           =   1950
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   2805
      TabIndex        =   7
      Text            =   "00:00"
      Top             =   720
      Width           =   630
   End
   Begin VB.TextBox txtNr 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   420
   End
   Begin VB.TextBox txtOrder 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   420
   End
   Begin VB.CommandButton btnDelete 
      Height          =   495
      Left            =   12120
      Picture         =   "frmTournamentMatches.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VBFLXGRD14.VBFlexGrid grdMatches 
      Height          =   2055
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3625
      BackColorBkg    =   -2147483633
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   10680
      TabIndex        =   2
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton btnSave 
      Height          =   495
      Left            =   11520
      Picture         =   "frmTournamentMatches.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin MSComCtl2.UpDown UpDnMinutes 
      Height          =   375
      Left            =   3420
      TabIndex        =   8
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Increment       =   15
      Max             =   45
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDnHours 
      Height          =   375
      Left            =   2550
      TabIndex        =   9
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Max             =   23
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   375
      Left            =   1575
      TabIndex        =   10
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MM"
      Format          =   147193859
      CurrentDate     =   43939
   End
   Begin MSComCtl2.UpDown upDnNr 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtNr"
      BuddyDispid     =   196614
      OrigLeft        =   840
      OrigTop         =   480
      OrigRight       =   1095
      OrigBottom      =   855
      Max             =   144
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDnOrder 
      Height          =   375
      Left            =   540
      TabIndex        =   12
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtOrder"
      BuddyDispid     =   196615
      OrigLeft        =   840
      OrigTop         =   480
      OrigRight       =   1095
      OrigBottom      =   855
      Max             =   144
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ordr"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Datum"
      Height          =   255
      Index           =   1
      Left            =   1575
      TabIndex        =   19
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tijd"
      Height          =   255
      Index           =   2
      Left            =   2805
      TabIndex        =   18
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Team A"
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   17
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Team B"
      Height          =   255
      Index           =   4
      Left            =   5655
      TabIndex        =   16
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Soort wedstrijd"
      Height          =   255
      Index           =   5
      Left            =   7590
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Locatie"
      Height          =   255
      Index           =   6
      Left            =   9405
      TabIndex        =   14
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nr"
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   13
      Top             =   480
      Width           =   375
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
Attribute VB_Name = "frmTournamentMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dontMove As Boolean 'prevent editBar from updateting

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rsTeams As ADODB.Recordset
Dim rsTypes As ADODB.Recordset
Dim rsLocation As ADODB.Recordset

Dim colWidths(9) As Double

Sub setMatchGrid()
Dim sqlstr As String
Dim dCol As Object
Dim col As Column
Dim i As Integer, J As Integer
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
  sqlstr = "SELECT m.matchOrder as Ord, m.matchNumber as nr, format(m.matchDate,'dd-MM') as Datum, format(matchTime ,'HH:NN') as Tijd, "
  sqlstr = sqlstr & " a.teamcode as A, ta.teamName as Team1, b.teamcode as B, tb.teamName as Team2, "
  sqlstr = sqlstr & " t.matchTypeDescription as Type, s.stadiumName & '/' & s.stadiumLocation as Locatie"
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
    .Top = 1200
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
    If rs.RecordCount = 0 Then Exit Sub
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
      rs.MoveNext
    Loop
    .colWidth(0) = 400
    .ColAlignment(0) = flexAlignCenterCenter
    .colWidth(2) = 700
    .ColAlignment(2) = flexAlignCenterCenter
    .colWidth(3) = 600
    .ColAlignment(3) = flexAlignCenterCenter
    .colWidth(4) = 700
    .ColAlignment(4) = flexAlignCenterCenter
    .colWidth(5) = 1650
    .colWidth(6) = 700
    .ColAlignment(6) = flexAlignCenterCenter
    .colWidth(7) = 1650
    .colWidth(8) = 1900
    .colWidth(9) = 3250
    .colWidth(1) = 500
    .ColAlignment(1) = flexAlignCenterCenter
    
   End With
   'select entire row
  ' Me.grdMatches.MarqueeStyle = dbgHighlightRow
   'force update of editBar controls
   'grdMatches_RowColChange 1, 1
    
End Sub

Sub setEditBar()

    Set rsTeams = New ADODB.Recordset
    Set rsTypes = New ADODB.Recordset
    Set rsLocation = New ADODB.Recordset

    Dim sqlstr As String

    ' Using DataCombo boxes for a change. Is so much easties in this case
    ' Normal ComboBox.ItemData can only be long data type
    'besides it is doing strange thing when filling and getting the actual value

    sqlstr = "Select teamID, teamCode & ': ' & teamName as team "
    sqlstr = sqlstr & "from tblTournamentTeamCodes c LEFT JOIN tblTeamNames n on c.teamId = n.teamnameid"
    sqlstr = sqlstr & " Where c.tournamentid = " & thisTournament
    FillCombo Me.cmbTeamA, sqlstr, cn, "team", "teamID"
    
    FillCombo Me.cmbTeamB, sqlstr, cn, "team", "teamID"
    
    sqlstr = "Select matchtypeId as id , matchtypedescription as descr from tblMatchTypes"
    FillCombo Me.cmbTypes, sqlstr, cn, "descr", "id"
    
    sqlstr = "Select stadiumId as id, stadiumName & '/' & stadiumLocation as name from tblStadiums order by stadiumName"
    FillCombo Me.cmbLocation, sqlstr, cn, "name", "id"
    rsLocation.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    Me.dtDate = getTournamentInfo("tournamentStartDate", cn)
    Me.UpDnHours = 20
    
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
'delete curtrent selected match
  Dim sqlstr As String
  Dim curMatch As Integer
  curMatch = Me.UpDnOrder
  If Me.upDnNr > 0 Then
    If MsgBox("Wedstrijd nr: " & Me.upDnNr & " verwijderen?", vbQuestion + vbYesNo, "Wedstrijden") = vbYes Then
      sqlstr = "Delete from tblTournamentSchedule "
      sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
      sqlstr = sqlstr & " AND matchOrder = " & curMatch
      cn.Execute sqlstr
      If Me.UpDnOrder >= Me.grdMatches.rows - 2 Then curMatch = curMatch - 1
      setMatchGrid
      findInGrid CStr(curMatch)
    End If
  End If
End Sub

Private Sub btnSave_Click()
    Dim sqlstr As String
    Dim curMatch As Integer
    Set rs = New ADODB.Recordset
    'validate
    Dim msg As String
    msg = ""
    If Not IsNumeric(Me.txtNr.Text) Then msg = msg & "Geen wedstrijdnummer" & vbNewLine
    If Me.cmbTeamA.Text = "" Then msg = msg & "Geen Team A" & vbNewLine
    If Me.cmbTeamB.Text = "" Then msg = msg & "Geen Team B" & vbNewLine
    If Me.cmbTypes.Text = "" Then msg = msg & "Geen soort wedstrijd" & vbNewLine
    If Me.cmbLocation.Text = "" Then msg = msg & "Geen locatie" & vbNewLine
    If Not IsNumeric(Me.txtOrder.Text) Then Me.UpDnOrder = Me.upDnNr
    If msg > "" Then
        msg = "FOUT: " & vbNewLine & msg
        MsgBox msg, vbOKOnly + vbCritical, "Wedstrijd toevoegen"
        Exit Sub
    End If
    
    sqlstr = "Select * from tblTournamentSchedule Where tournamentId = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    curMatch = Me.UpDnOrder
    rs.Find "matchOrder = " & curMatch
    If rs.EOF Then 'add new match
        rs.AddNew
    End If
    With rs
        !tournamentID = thisTournament
        !matchOrder = curMatch
        !matchNumber = val(Me.txtNr)
        !matchDate = CDbl(Me.dtDate)
        !matchtime = IIf(Me.txtTime = "24:00", "23:59", Me.txtTime)
        !matchteamA = Left(Me.cmbTeamA.Text, InStr(Me.cmbTeamA.Text, ":") - 1)
        !matchteamB = Left(Me.cmbTeamB.Text, InStr(Me.cmbTeamB.Text, ":") - 1)
        !matchType = Me.cmbTypes.ItemData(Me.cmbTypes.ListIndex)
        !matchStadiumID = Me.cmbLocation.ItemData(Me.cmbLocation.ListIndex)
    End With
    rs.Update
    
    setMatchGrid
    findInGrid CStr(curMatch)
    DoEvents
    
End Sub

Sub findInGrid(txt As String)
'find the txt in the grdMatches
  Dim i As Integer
  Dim found As Boolean
  For i = 1 To Me.grdMatches.rows - 1
    If Me.grdMatches.TextMatrix(i, 0) = txt Then
      found = True
      Exit For
    End If
  Next
  Me.btnDelete.Enabled = found
  If found Then
    Me.grdMatches.row = i
    grdMatches_RowColChange
  Else
    Me.cmbTeamA.Text = ""
    Me.cmbTeamB.Text = ""
    Me.cmbLocation.Text = ""
    Me.dtDate.SetFocus
  End If
End Sub


Private Sub Form_Load()
    'open the database
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    setEditBar
    setMatchGrid
    Me.grdMatches.row = 2
    Me.grdMatches.row = 1
    setState 'only if admin is logged in is editting possible
    
    UnifyForm Me
    centerForm Me

End Sub

Sub setState()
Dim ctl As Control
Dim col As Object
    For Each ctl In Me.Controls
        If TypeOf ctl Is UpDown _
            Or TypeOf ctl Is ComboBox _
            Or TypeOf ctl Is DTPicker _
            Or TypeOf ctl Is TextBox _
            Or ctl.Name = "btnSave" Then
            ctl.Enabled = adminLogin
        End If
'                TypeOf ctl Is DataCombo Or _
'        Me.grdMatches.AllowAddNew = adminLogin
'        Me.grdMatches.AllowDelete = adminLogin
'        Me.grdMatches.AllowUpdate = adminLogin
'        For Each col In Me.grdMatches.Columns
'            col.Locked = Not adminLogin
'        Next
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not rs Is Nothing Then
        If (rs.State And adStateOpen) = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    If Not rsTeams Is Nothing Then
        If (rsTeams.State And adStateOpen) = adStateOpen Then rsTeams.Close
        Set rsTeams = Nothing
    End If
    
    If Not rsLocation Is Nothing Then
        If (rsLocation.State And adStateOpen) = adStateOpen Then rsLocation.Close
        Set rsLocation = Nothing
    End If
    If Not rsTypes Is Nothing Then
        If (rsTypes.State And adStateOpen) = adStateOpen Then rsTypes.Close
        Set rsTypes = Nothing
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
initWidth = 12945
gridHeight = 7815
gridWidth = 12360
With Me
  If .Height < initHeigth Then .Height = initHeigth
  If .width < initWidth Then .width = initWidth
  .grdMatches.Height = .Height - 2400
  .grdMatches.width = .width - 585
  .btnClose.Top = .Height - 1095
  .btnClose.Left = .width - 2265
  .btnDelete.Left = .width - 1000
  .btnSave.Left = .btnDelete.Left - .btnSave.width - 20
  i = 0
'  For Each col In .grdMatches.Columns
'    col.width = .grdMatches.width * colWidths(i)
'    i = i + 1
'  Next
  
End With

End Sub


Private Sub grdMatches_RowColChange()
  Dim sqlstr As String
  Dim i As Integer
  Dim searchTeam As String
  Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblTournamentSchedule  where tournamentID = " & thisTournament
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  rs.Find "matchOrder = " & Me.grdMatches.TextMatrix(Me.grdMatches.row, 0)
  With rs
    If Not .EOF Then
      Me.upDnNr = !matchNumber
'      Me.txtNr = !matchnumber
      Me.dtDate = !matchDate
      Me.txtTime = !matchtime
'      Me.cmbTypes.BoundText = !matchType
      setCombo Me.cmbTypes, !matchType
'      For i = 0 To Me.cmbTypes.ListCount - 1
'        If Me.cmbTypes.ItemData(i) = !matchType Then
'          Me.cmbTypes.ListIndex = i
'          Exit For
'        End If
'      Next
'      Me.cmbLocation.BoundText = !matchStadiumId
      setCombo Me.cmbLocation, CInt(!matchStadiumID)
'      For i = 0 To Me.cmbLocation.ListCount - 1
'        If Me.cmbLocation.ItemData(i) = !matchStadiumID Then
'          Me.cmbLocation.ListIndex = i
'          Exit For
'        End If
'      Next
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''CANNOT USE setCombo because of INSTR function
'      setCombo Me.cmbTeamA, !matchteamA
''''''''''''''''''''''''''''''''''''''''''''''''''
      For i = 0 To Me.cmbTeamA.ListCount - 1
        If InStr(Me.cmbTeamA.List(i), !matchteamA) Then
          Me.cmbTeamA.ListIndex = i
          Exit For
        End If
      Next
'      Me.cmbTeamB.BoundText = !matchTeamB
'      setCombo Me.cmbTeamB, !matchteamB
      For i = 0 To Me.cmbTeamB.ListCount - 1
        If InStr(Me.cmbTeamB.List(i), !matchteamB) Then
          Me.cmbTeamB.ListIndex = i
          Exit For
        End If
      Next
      Me.UpDnOrder = !matchOrder
      'Me.txtOrder = !matchorder
    End If
  End With
  With Me.grdMatches
    If .TopRow < val(.TextMatrix(.row, 0)) - 25 Then
      .TopRow = .TextMatrix(.row, 0) - 25
    End If
    If .TopRow > val(.TextMatrix(.row, 0)) Then
      .TopRow = val(.TextMatrix(.row, 0))
    End If
    
  End With
End Sub


Private Sub updnMinutes_Change()
  Me.txtTime = Format(Me.UpDnHours, "00") & ":" & Format(Me.UpDnMinutes, "00")
End Sub

Private Sub updnHours_Change()
  Me.txtTime = Format(Me.UpDnHours, "00") & ":" & Format(Me.UpDnMinutes, "00")
End Sub


Private Sub upDnNr_DownClick()
  findInGrid Me.txtNr
End Sub

Private Sub upDnNr_UpClick()
  findInGrid Me.txtNr
End Sub
