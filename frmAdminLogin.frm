VERSION 5.00
Begin VB.Form frmAdminLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Admin login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
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
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1410
      TabIndex        =   2
      Top             =   135
      Width           =   2205
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1410
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2205
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Naam:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Wachtwoord:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   540
      Width           =   1200
   End
End
Attribute VB_Name = "frmAdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private mUsername As String
Private mPassword As String
Private mCancel As Boolean

Public Function GetLogIn(ByRef UserName As String, ByRef Password As String) As Boolean
    Me.txtUserName.Text = UserName
    
    Me.Show vbModal
    
    UserName = mUsername
    Password = mPassword
    
    GetLogIn = Not mCancel
End Function

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    adminLogin = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mCancel = False
    mUsername = Me.txtUserName.Text
    mPassword = Me.txtPassword.Text
    
    Unload Me
End Sub

