VERSION 5.00
Begin VB.Form frm8Finals 
   Caption         =   "8e finales"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Annuleren"
      Height          =   375
      Left            =   8640
      TabIndex        =   43
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton btnSaveClose 
      Caption         =   "Opslaan"
      Height          =   375
      Left            =   10440
      TabIndex        =   31
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame frmMatch 
      Caption         =   "Wedstrijd nr 42"
      Height          =   1335
      Index           =   7
      Left            =   9120
      TabIndex        =   24
      Top             =   2880
      Width           =   2895
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2F"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   840
         TabIndex        =   42
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   840
         TabIndex        =   41
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1D:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   390
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "2F:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   285
      End
   End
   Begin VB.Frame frmMatch 
      Caption         =   "Wedstrijd nr 43"
      Height          =   1335
      Index           =   6
      Left            =   6120
      TabIndex        =   21
      Top             =   2880
      Width           =   2895
      Begin VB.ComboBox cmb3rdPlace 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   840
         TabIndex        =   30
         Top             =   780
         Width           =   1935
      End
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1E"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   840
         TabIndex        =   40
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1E:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   390
         Width           =   285
      End
      Begin VB.Label lblTeamCmb 
         AutoSize        =   -1  'True
         Caption         =   "3ABCD:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   645
      End
   End
   Begin VB.Frame frmMatch 
      Caption         =   "Wedstrijd nr 42"
      Height          =   1335
      Index           =   5
      Left            =   3120
      TabIndex        =   18
      Top             =   2880
      Width           =   2895
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2E"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   840
         TabIndex        =   39
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   840
         TabIndex        =   38
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "2E:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2D:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   390
         Width           =   300
      End
   End
   Begin VB.Frame frmMatch 
      Caption         =   "Wedstrijd nr 41"
      Height          =   1335
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   2895
      Begin VB.ComboBox cmb3rdPlace 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   840
         TabIndex        =   29
         Top             =   780
         Width           =   1935
      End
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1F"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   840
         TabIndex        =   37
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1F:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   390
         Width           =   285
      End
      Begin VB.Label lblTeamCmb 
         AutoSize        =   -1  'True
         Caption         =   "3ABC:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   525
      End
   End
   Begin VB.Frame frmMatch 
      Caption         =   "Wedstrijd nr 40"
      Height          =   1335
      Index           =   3
      Left            =   9120
      TabIndex        =   12
      Top             =   1320
      Width           =   2895
      Begin VB.ComboBox cmb3rdPlace 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   840
         TabIndex        =   28
         Top             =   780
         Width           =   1935
      End
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   840
         TabIndex        =   36
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblTeamCmb 
         AutoSize        =   -1  'True
         Caption         =   "3DEF:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1C:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   390
         Width           =   300
      End
   End
   Begin VB.Frame frmMatch 
      Caption         =   "Wedstrijd nr 39"
      Height          =   1335
      Index           =   2
      Left            =   6120
      TabIndex        =   9
      Top             =   1320
      Width           =   2895
      Begin VB.ComboBox cmb3rdPlace 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   840
         TabIndex        =   27
         Top             =   780
         Width           =   1935
      End
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   840
         TabIndex        =   35
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblTeamCmb 
         AutoSize        =   -1  'True
         Caption         =   "3ADEF:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1B:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   390
         Width           =   285
      End
   End
   Begin VB.Frame frmMatch 
      Caption         =   "Wedstrijd nr 38"
      Height          =   1335
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   1320
      Width           =   2895
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   720
         TabIndex        =   34
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   33
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2A:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   390
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "2B:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   285
      End
   End
   Begin VB.Frame frmMatch 
      Caption         =   "Wedstrijd nr 37"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2895
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   720
         TabIndex        =   32
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblTeamName 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "2C:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1A:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   300
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Achtste finales"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Tag             =   "subkop"
      Top             =   0
      Width           =   11775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Selecteer de overgebleven teams in wedstrijden 39, 40, 41 en 43 (De beste derde plaatsen uit de groepswedstrijden)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   12015
   End
End
Attribute VB_Name = "frm8Finals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Dim allPlayed As Boolean

Private Sub btnCancel_Click()
  Unload Me
End Sub

Private Sub btnSaveClose_Click()
  Dim matchNr As Integer
  matchNr = getFirstFinalMatchNumber(cn) - 1
  If savPlaces Then
    update3rdPlacePoints matchNr, cn
    updatePoolPositions matchNr, cn
    showInfo False
    Unload Me
  Else
    MsgBox "Je moet alle 4 de derde plaatsen invullen", vbOKOnly + vbCritical, "Niet alles ingevuld"
  End If
End Sub

Function savPlaces()
  Dim i As Integer
  Dim sqlstr As String
  Dim whereStr As String
  For i = 0 To Me.cmb3rdPlace.UBound
    If Me.cmb3rdPlace(i) = "" Then
      savPlaces = False
    Else
      sqlstr = "UPDATE tblTournamentTeamCodes SET teamID = " & Me.cmb3rdPlace(i).ItemData(Me.cmb3rdPlace(i).ListIndex)
      sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
      Select Case i
        Case 0 'macth 39
          whereStr = sqlstr & " AND teamcode = '3ADEF'"
          cn.Execute whereStr
        Case 1 'match 40
          whereStr = sqlstr & " AND teamcode = '3DEF'"
          cn.Execute whereStr
        Case 2 'macth 41
          whereStr = sqlstr & " AND teamcode = '3ABC'"
          cn.Execute whereStr
        Case 3 'match 43
          whereStr = sqlstr & " AND teamcode = '3ABCD'"
          cn.Execute whereStr
      End Select
      savPlaces = True
    End If
  Next
End Function

Private Sub Form_Load()
  Dim i As Integer
  Set cn = New ADODB.Connection
  With cn
      .ConnectionString = lclConn()
      .Open
  End With
  If getMatchCount(1, cn) > getLastMatchPlayed(cn) Then
    MsgBox "Dit formulier pas gebruiken nadat alle groepswedstrijden zijn gespeeld", vbOKOnly + vbExclamation, "Achtste finales"
    allPlayed = False
    Exit Sub
  Else
    allPlayed = True
    centerForm Me
    UnifyForm Me
    initForm
  End If

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

Sub initForm()
  Dim sqlstr As String
  Dim i As Integer
  Set rs = New ADODB.Recordset
  Dim teamGroups As String
  Dim cmbSelstr As String
  sqlstr = "Select l.TeamID as id, t.teamName as team from tblTeamNames t"
  sqlstr = sqlstr & " INNER JOIN tblGroupLayout l ON t.teamNameID = l.teamID "
  sqlstr = sqlstr & " WHERE (l.tournamentId = " & thisTournament
  sqlstr = sqlstr & " AND l.teamposition = 3"
  sqlstr = sqlstr & ") AND ("
  teamGroups = "l.groupletter = 'A' or l.groupletter = 'D' or l.groupletter = 'E' or l.groupletter = 'F')"
  cmbSelstr = sqlstr & teamGroups & " ORDER BY t.teamname"
  FillCombo Me.cmb3rdPlace(0), cmbSelstr, cn, "team", "id"
  teamGroups = "l.groupletter = 'D' or l.groupletter = 'E' or l.groupletter = 'F')"
  cmbSelstr = sqlstr & teamGroups & " ORDER BY t.teamname"
  FillCombo Me.cmb3rdPlace(1), cmbSelstr, cn, "team", "id"
  teamGroups = "l.groupletter = 'A' or l.groupletter = 'B' or l.groupletter = 'C')"
  cmbSelstr = sqlstr & teamGroups & " ORDER BY t.teamname"
  FillCombo Me.cmb3rdPlace(2), cmbSelstr, cn, "team", "id"
  teamGroups = "l.groupletter = 'A' or l.groupletter = 'B' or l.groupletter = 'C' or l.groupletter = 'D' )"
  cmbSelstr = sqlstr & teamGroups & " ORDER BY t.teamname"
  FillCombo Me.cmb3rdPlace(3), cmbSelstr, cn, "team", "id"
  
  Me.btnSaveClose.Enabled = allPlayed
  
  sqlstr = "Select * from tblTournamentTeamCodes where tournamentid = " & thisTournament
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  'get the teams that are allready known
  If Not rs.EOF Then
    For i = 0 To Me.lblTeamName.UBound
      rs.Find "teamcode = '" & Me.lblTeamName(i).Caption & "'"
      If Not rs.EOF Then
        Me.lblTeamName(i).Caption = getTeamInfo(rs!teamID, "teamName", cn)
      End If
      rs.MoveFirst
    Next
  End If
  'fill the combo's  with aleady chosen values
  For i = 0 To Me.lblTeamCmb.UBound
    With Me.lblTeamCmb(i)
      rs.Find "teamcode ='" & Left(.Caption, Len(.Caption) - 1) & "'"
      If Not rs.EOF Then
        setCombo Me.cmb3rdPlace(i), rs!teamID
      Else
        Me.cmb3rdPlace(i).ListIndex = -1
      End If
      rs.MoveFirst
    End With
  Next
End Sub

