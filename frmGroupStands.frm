VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGroupStands 
   Caption         =   "Groepstanden"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame frmGroup 
      Caption         =   "Frame1"
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5415
      Begin VB.TextBox txtPos 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   4
         Top             =   1800
         Width           =   360
      End
      Begin VB.CommandButton btnSavePos 
         Caption         =   "Opslaan"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin MSComCtl2.UpDown updnPos 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   2
         Top             =   1800
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtPos(0)"
         BuddyDispid     =   196611
         BuddyIndex      =   0
         OrigLeft        =   4560
         OrigTop         =   1800
         OrigRight       =   4815
         OrigBottom      =   2175
         Max             =   4
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid grdGroup 
         Height          =   1575
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   5
         Cols            =   9
         FixedCols       =   0
      End
      Begin VB.Label lblTeamName 
         Alignment       =   1  'Right Justify
         Caption         =   "Positie"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   2265
      End
   End
   Begin VB.Label lblSub 
      Alignment       =   2  'Center
      Caption         =   "Klik op de team naam en pas de positie aan"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   10695
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Controleer de groepsstand en pas eventueel aan (per groep op 'Opslaan' klikken)"
      Height          =   375
      Left            =   -240
      TabIndex        =   6
      Tag             =   "kop"
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "frmGroupStands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Private Sub btnClose_Click()
  Unload Me
End Sub

Private Sub btnSavePos_Click(Index As Integer)
  savePositions Index
End Sub

Sub savePositions(Index As Integer)
'save the positions in the final group standing
'check if the positin are all different
  Dim J As Integer
  Dim i As Integer
  Dim msg As String
  Dim sqlstr As String
  Dim savePos As Boolean
  savePos = True
  For i = 1 To 3
    For J = i + 1 To 4
      With Me.grdGroup(Index)
        If .TextMatrix(i, 0) = .TextMatrix(J, 0) Then
          msg = "Teams kunnen niet gelijk eindigen!!"
          MsgBox msg, vbOKOnly + vbCritical, "Fout in de posities"
          savePos = False
          Exit For
        End If
      End With
    Next
    If Not savePos Then Exit For
  Next
  If savePos Then
    For i = 1 To 4
      sqlstr = "UPDATE tblGroupLayout SET teamPosition = " & val(Me.grdGroup(Index).TextMatrix(i, 0))
      sqlstr = sqlstr & " where teamID = " & getTeamIdFromName(Me.grdGroup(Index).TextMatrix(i, 1), cn)
      sqlstr = sqlstr & " AND tournamentID = " & thisTournament
      cn.Execute sqlstr
    Next
  End If
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Set cn = New ADODB.Connection
  With cn
      .ConnectionString = lclConn()
      .Open
  End With
  centerForm Me
  UnifyForm Me
  initForm
End Sub

Sub initForm()
  'fillGroupGrds
Dim i As Integer
  For i = 0 To getTournamentInfo("tournamentGroupCount", cn) - 1
    fillGroupGrd i
  Next
  Me.btnClose.Top = Me.frmGroup(i - 1).Top + Me.frmGroup(i - 1).Height + 120
  Me.btnClose.Left = Me.frmGroup(i - 1).Left + Me.frmGroup(i - 1).width - (Me.btnClose.width)
  Me.Height = Me.Height + Me.btnClose.Height
  Me.width = Me.width - 500
End Sub

Sub fillGroupGrd(Index As Integer)
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim cl As Integer
Dim rw As Integer
Dim i As Integer
Dim xPos As Integer
Dim yPos As Integer
  Set rs = New ADODB.Recordset
  xPos = Me.frmGroup(0).Left
  yPos = Me.frmGroup(0).Top
  Me.frmGroup(0).Caption = "Groep A"
  
  i = 1
  If Index > 0 Then
    Load Me.frmGroup(Index)
    Load Me.grdGroup(Index)
    Load Me.lblTeamName(Index)
    Load Me.txtPos(Index)
    Load Me.updnPos(Index)
    
    Me.frmGroup(Index).Visible = True
    Me.frmGroup(Index).Caption = "Groep " & Chr(65 + Index)
    Set Me.grdGroup(Index).Container = Me.frmGroup(Index)
    Set Me.lblTeamName(Index).Container = Me.frmGroup(Index)
    Set Me.txtPos(Index).Container = Me.frmGroup(Index)
    Set Me.updnPos(Index).Container = Me.frmGroup(Index)
    Me.lblTeamName(Index).Left = Me.lblTeamName(Index - 1).Left
    Me.lblTeamName(Index).Top = Me.lblTeamName(Index - 1).Top
'    Me.updnPos(Index).BuddyProperty = "Text"
    Me.txtPos(Index).Left = Me.txtPos(Index - 1).Left
    Me.txtPos(Index).Top = Me.txtPos(Index - 1).Top
    Me.txtPos(Index).width = Me.txtPos(Index - 1).width
    Me.updnPos(Index).Left = Me.updnPos(Index - 1).Left
    Me.updnPos(Index).Top = Me.updnPos(Index - 1).Top
    Me.lblTeamName(Index).Visible = True
    Me.txtPos(Index).Visible = True
    Me.updnPos(Index).Visible = True
    Load Me.btnSavePos(Index)
    Set Me.btnSavePos(Index).Container = Me.frmGroup(Index)
    Me.btnSavePos(Index).Visible = True
'    With Me.grdGroup(index)
'      .Left = Me.grdGroup(index - 1).Left
'      .Top = Me.grdGroup(index - 1).Top
'    End With
    
    yPos = Me.frmGroup(Index - 1).Top + Me.frmGroup(Index - 1).Height + 20
    If Index = getTournamentInfo("tournamentGroupCount", cn) / 2 Then
      yPos = Me.frmGroup(0).Top
    End If
    If Index >= getTournamentInfo("tournamentGroupCount", cn) / 2 Then
      xPos = xPos + Me.frmGroup(0).width + 30
      i = i + 1
    End If
    Me.frmGroup(Index).Left = xPos
    Me.frmGroup(Index).Top = yPos
    Me.updnPos(Index).BuddyControl = Me.txtPos(Index)
    Me.txtPos(0).width = Me.txtPos(Index).width
    Me.updnPos(0).BuddyControl = Me.txtPos(0)
  End If
  sqlstr = "Select teamposition as ps, teamName as Team, mPl as pl, mWon as wn, mLost as vl, "
  sqlstr = sqlstr & "mDraw as gl, mScored as V, mAgainst as T, teampoints as pt "
  sqlstr = sqlstr & " from tblGroupLayout gl INNER JOIN tblTeamNames tn on gl.teamID = tn.teamNameId"
  sqlstr = sqlstr & " where tournamentID = " & thisTournament
  With Me.grdGroup(Index)
    .Visible = True
    sqlstr = sqlstr & " AND groupLetter = '" & Chr(65 + Index) & "'"
    sqlstr = sqlstr & " ORDER BY teamposition, groupPlace"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    .width = 1500
    For cl = 0 To rs.Fields.Count - 1
      .TextMatrix(0, cl) = rs.Fields(cl).Name
      .colWidth(cl) = 360
      .width = .width + .colWidth(cl)
      .ColAlignment(cl) = flexAlignCenterCenter
    Next
    If Index > 0 Then
      .width = Me.grdGroup(0).width
    End If
    .colWidth(1) = 1500
    .ColAlignment(1) = flexAlignLeftCenter
    rw = 0
    Do While Not rs.EOF
      rw = rw + 1
      For cl = 0 To rs.Fields.Count - 1
        .TextMatrix(rw, cl) = rs.Fields(cl)
      Next
      rs.MoveNext
    Loop
    Me.lblTeamName(Index).Caption = "Positie " & .TextMatrix(1, 1)
    Me.updnPos(Index) = 1
    'check if all matches in this are played and enable postition controls
    Me.btnSavePos(Index).Enabled = grpPlayedCount(Chr(65 + Index), cn) = 6
    Me.txtPos(Index).Enabled = Me.btnSavePos(Index).Enabled
    Me.updnPos(Index).Enabled = Me.btnSavePos(Index).Enabled
  End With
  Me.frmGroup(Index).width = Me.grdGroup(Index).width + Me.grdGroup(Index).Left + 30
  Me.Height = Me.frmGroup(Index).Top + Me.frmGroup(Index).Height + 1000
  Me.width = (Me.frmGroup(0).Left + Me.frmGroup(0).width + 400) * i
  rs.Close
  Set rs = Nothing
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

Private Sub grdGroup_Click(Index As Integer)
  With Me.grdGroup(Index)
    Me.lblTeamName(Index).Caption = "Positie team " & .TextMatrix(.row, 1)
    Me.updnPos(Index) = val(.TextMatrix(.row, 0))
  End With
End Sub


Private Sub updnPos_Change(Index As Integer)
  Dim r As Integer
  r = Me.grdGroup(Index).row
  Me.grdGroup(Index).TextMatrix(r, 0) = Me.updnPos(Index)
End Sub
