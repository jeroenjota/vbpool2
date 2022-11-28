VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Ogenblikje"
   ClientHeight    =   2475
   ClientLeft      =   7860
   ClientTop       =   6420
   ClientWidth     =   4680
   ControlBox      =   0   'False
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4680
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   1950
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   4410
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   1290
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   4410
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   675
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   4410
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   4410
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    UnifyForm Me
    centerForm Me
End Sub
