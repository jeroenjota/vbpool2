VERSION 5.00
Begin VB.Form frmPlayersNew 
   Caption         =   "Speler toevoegen"
   ClientHeight    =   2760
   ClientLeft      =   16335
   ClientTop       =   6420
   ClientWidth     =   4470
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
   ScaleHeight     =   2760
   ScaleWidth      =   4470
   Begin VB.CheckBox chkActive 
      Alignment       =   1  'Right Justify
      Caption         =   "Actief"
      Height          =   240
      Left            =   3360
      TabIndex        =   12
      Top             =   1387
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   360
      Left            =   960
      TabIndex        =   11
      Top             =   1867
      Width           =   2175
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Opslaan"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1860
      Width           =   1095
   End
   Begin VB.TextBox txtNickName 
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtAName 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtTname 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtVnaam 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Land"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nieuwe speler"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Tag             =   "kop"
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Bijnaam"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Achternaam"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voornaam"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmPlayersNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private currentCountry As Long

Public Property Get Country() As Long
    Country = currentCountry
End Property

Public Property Let Country(ByVal NewValue As Long)
    currentCountry = NewValue
End Property

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Dim sqlstr As String
    Dim teamID As Long
    Dim teamCode As Integer
    Dim peopleID As Long
    Dim rs As New ADODB.Recordset
    If Me.txtNickName = "" Then Me.txtNickName = buildNickName
    
    If Not playerExists(Me.txtVnaam, Me.txtTname, Me.txtAName, Me.txtNickName, cn) Then
        teamCode = val(Me.cmbCountry.ItemData(Me.cmbCountry.ListIndex))
        sqlstr = "Insert into tblPeople (firstName, middleName, lastName, nickName, functionID, countryCode, active"
        sqlstr = sqlstr & ") VALUES ('" & Me.txtVnaam
        sqlstr = sqlstr & "','" & Me.txtTname
        sqlstr = sqlstr & "','" & Me.txtAName
        sqlstr = sqlstr & "','" & Me.txtNickName
        sqlstr = sqlstr & "', 3" 'make it a midfielder, for now
        sqlstr = sqlstr & ", " & teamCode
        sqlstr = sqlstr & ", -1)"
        cn.Execute sqlstr
        
        sqlstr = "Select peopleID from tblPeople"
        rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
        rs.MoveLast
        peopleID = rs!peopleID
        rs.Close
        Set rs = Nothing
        teamID = getTeamIDFromCountryCode(teamCode, cn)
        sqlstr = "Insert into tblTeamPlayers( tournamentID, teamID, PlayerID"
        sqlstr = sqlstr & ") VALUES (" & thisTournament
        sqlstr = sqlstr & ", " & teamID
        sqlstr = sqlstr & ", " & peopleID
        sqlstr = sqlstr & ")"
        cn.Execute sqlstr
    Else
        MsgBox "Speler bestaat al", vbOKOnly + vbInformation, "Speler toevoegen"
    End If
    Unload Me
End Sub

Private Sub Form_Load()
Dim sqlstr As String
Dim i As Integer
Dim thisCountry
thisCountry = Country
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    sqlstr = "Select * from tblCountries ORDER BY countryName"
    FillCombo Me.cmbCountry, sqlstr, cn, "countryname", "countryid"
    'Me.cmbCountry.ListIndex = -1
    setCombo Me.cmbCountry, thisCountry
'    If thisCountry > 0 Then
'      Do While Me.cmbCountry.ItemData(i) <> thisCountry
'        i = i + 1
'      Loop
'      Me.cmbCountry.ListIndex = i
'    End If
    UnifyForm Me
End Sub

Function buildNickName()
Dim nickName As String
    nickName = Me.txtAName
    If Me.txtVnaam > "" Or Me.txtTname > "" Then
        If nickName > "" Then nickName = nickName & ","
    End If
    If Me.txtVnaam > "" Then nickName = nickName & " " & Me.txtVnaam
    If Me.txtTname > "" Then nickName = nickName & " " & Me.txtTname
    buildNickName = Trim(nickName)
End Function

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

Private Sub Label4_Click()
    Me.txtNickName = buildNickName
End Sub

Private Sub txtNickName_GotFocus()
  Me.txtNickName = buildNickName()
End Sub
