VERSION 5.00
Begin VB.Form frmOrganisation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Organisatie"
   ClientHeight    =   3015
   ClientLeft      =   5175
   ClientTop       =   3090
   ClientWidth     =   6420
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Tag             =   "adressen"
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1530
      PasswordChar    =   "*"
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Opslaan"
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtCity 
      DataField       =   "plaats"
      Height          =   360
      Left            =   2535
      TabIndex        =   8
      Top             =   1575
      Width           =   3795
   End
   Begin VB.TextBox txtEmail 
      Height          =   360
      Left            =   3795
      TabIndex        =   12
      Top             =   2055
      Width           =   2535
   End
   Begin VB.TextBox txtAnaam 
      DataField       =   "aNaam"
      Height          =   360
      Left            =   3360
      TabIndex        =   3
      Top             =   615
      Width           =   2955
   End
   Begin VB.TextBox txtTnaam 
      DataField       =   "tNaam"
      Height          =   360
      Left            =   2595
      TabIndex        =   2
      Top             =   615
      Width           =   690
   End
   Begin VB.TextBox txtVnaam 
      DataField       =   "vnaam"
      Height          =   360
      Left            =   1530
      TabIndex        =   1
      Top             =   615
      Width           =   1005
   End
   Begin VB.TextBox txtStreet 
      DataField       =   "Adres"
      Height          =   360
      Left            =   1530
      TabIndex        =   5
      Top             =   1095
      Width           =   4785
   End
   Begin VB.TextBox txtPostcode 
      DataField       =   "poco"
      Height          =   360
      Left            =   1530
      TabIndex        =   7
      Top             =   1575
      Width           =   900
   End
   Begin VB.TextBox txtPhone 
      Height          =   360
      Left            =   1530
      TabIndex        =   10
      Top             =   2055
      Width           =   1515
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nieuw admin wachtwoord:"
      Height          =   495
      Left            =   45
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Organisatie gegevens"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Tag             =   "kop"
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Email"
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   3120
      TabIndex        =   11
      Top             =   2115
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Telefoon"
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   45
      TabIndex        =   9
      Top             =   2115
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Postcode/Plaats"
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   45
      TabIndex        =   6
      Top             =   1635
      Width           =   1380
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Adres"
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   45
      TabIndex        =   4
      Top             =   1155
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Naam"
      ForeColor       =   &H00004000&
      Height          =   240
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   1095
   End
End
Attribute VB_Name = "frmOrganisation"
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

Private Sub btnSave_Click()
    Dim MD5 As New clsMD5
    Dim sqlstr As String
    On Error GoTo dberror
    Dim newPassword As String
    'organisation data
'    newPassword = UCase(MD5.DigestStrToHexStr(Me.txtPassword))
    cn.Execute "Delete from tblOrganisation"
    sqlstr = "INSERT INTO tblOrganisation ("
    sqlstr = sqlstr & "firstname, middlename, lastname, "
    sqlstr = sqlstr & "address, postalcode, city, "
    sqlstr = sqlstr & "telephone, email, passwd) VALUES ("
    sqlstr = sqlstr & "'" & Me.txtVnaam & "'"
    sqlstr = sqlstr & ", '" & Me.txtTnaam & "'"
    sqlstr = sqlstr & ", '" & Me.txtAnaam & "'"
    sqlstr = sqlstr & ", '" & Me.txtStreet & "'"
    sqlstr = sqlstr & ", '" & Me.txtPostcode & "'"
    sqlstr = sqlstr & ", '" & Me.txtCity & "'"
    sqlstr = sqlstr & ", '" & Me.txtPhone & "'"
    sqlstr = sqlstr & ", '" & Me.txtEmail & "'"
    sqlstr = sqlstr & ", '" & newPassword & "')"
    cn.Execute sqlstr
    Unload Me
    Exit Sub
dberror:
    MsgBox Err & ":  " & Error & vbNewLine & "Contact Jota Services", vbOKOnly + vbCritical, "Unrecoverable error"
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sqlstr As String
    
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn
        .Open
    End With
        
    Set rs = New ADODB.Recordset
    
    sqlstr = "Select * from tblOrganisation"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveFirst 'should be only one record
        Me.txtVnaam = rs!firstname
        Me.txtTnaam = rs!middlename
        Me.txtAnaam = rs!lastname
        Me.txtStreet = rs!Address
        Me.txtPostcode = rs!postalcode
        Me.txtCity = rs!city
        Me.txtPhone = rs!telephone
        Me.txtEmail = rs!email
    End If
    UnifyForm Me
    centerForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Clean-up procedure
    'rs recordset
    If Not rs Is Nothing Then
        If (rs.State And adStateOpen) = adStateOpen Then
            rs.Close
        End If
        Set rs = Nothing
    End If
    'same comment with cn
    If Not cn Is Nothing Then
        'first, check if the state is open, if yes then close it
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        'set them to nothing
        Set cn = Nothing
    End If
End Sub
