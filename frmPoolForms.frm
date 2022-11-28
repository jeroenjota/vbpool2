VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPoolForms 
   Caption         =   "Pool formulieren"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14115
   Icon            =   "frmPoolForms.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   14115
   StartUpPosition =   3  'Windows Default
   Tag             =   "xsmall"
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   12345
      TabIndex        =   137
      Top             =   9840
      Width           =   1695
   End
   Begin VB.Frame Frame9 
      Caption         =   "Topscorers"
      Height          =   1170
      Left            =   3240
      TabIndex        =   60
      Top             =   8640
      Width           =   2895
      Begin VB.CommandButton btnNewPlayer 
         Caption         =   "Nieuwe speler"
         Height          =   255
         Left            =   120
         TabIndex        =   140
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtTopScorerDP 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2160
         TabIndex        =   41
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox cmbTopScorer 
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
         Left            =   140
         TabIndex        =   40
         Top             =   480
         Width           =   1935
      End
      Begin MSComCtl2.UpDown updnTopScorerDP 
         Height          =   375
         Left            =   2520
         TabIndex        =   43
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtTopScorerDP"
         BuddyDispid     =   196612
         OrigLeft        =   2520
         OrigTop         =   480
         OrigRight       =   2775
         OrigBottom      =   855
         Max             =   20
         Enabled         =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "Doelp"
         Height          =   255
         Left            =   2280
         TabIndex        =   136
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Speler"
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTopScorer 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   20
         TabIndex        =   134
         Tag             =   "xsmall"
         Top             =   570
         Width           =   90
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Aantallen"
      Height          =   2535
      Left            =   3360
      TabIndex        =   59
      Top             =   1920
      Width           =   2775
      Begin VB.TextBox txtAantal 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   0
         Left            =   1920
         TabIndex        =   42
         Text            =   "1"
         Top             =   200
         Width           =   615
      End
      Begin VB.Label lblAantal 
         AutoSize        =   -1  'True
         Caption         =   "Gele kaarten"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   70
         Top             =   260
         Width           =   915
      End
   End
   Begin VB.Frame frmRanking 
      Caption         =   "Eindstand"
      Height          =   1170
      Left            =   120
      TabIndex        =   58
      Top             =   8640
      Width           =   3015
      Begin VB.ComboBox cmbRanking 
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
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbRanking 
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
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblRanking 
         AutoSize        =   -1  'True
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   133
         Tag             =   "xsmall"
         Top             =   803
         Width           =   90
      End
      Begin VB.Label lblRanking 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   132
         Tag             =   "xsmall"
         Top             =   323
         Width           =   90
      End
   End
   Begin VB.Frame frm1Finales 
      Caption         =   "Finalisten"
      Height          =   1095
      Left            =   4680
      TabIndex        =   57
      Top             =   7440
      Width           =   1460
      Begin VB.Frame frm1FinWed 
         Height          =   765
         Left            =   40
         TabIndex        =   68
         Top             =   230
         Width           =   1340
         Begin VB.Label lbl1FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   72
            Tag             =   "small"
            Top             =   450
            Width           =   195
         End
         Begin VB.Label lbl1FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   200
            TabIndex        =   79
            Tag             =   "small"
            Top             =   180
            Width           =   195
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   15
            Left            =   40
            TabIndex        =   81
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
      End
   End
   Begin VB.Frame frm3Finales 
      Caption         =   "Derde plaats"
      Height          =   1095
      Left            =   3240
      TabIndex        =   56
      Top             =   7440
      Width           =   1400
      Begin VB.Frame frm3FinWed 
         Height          =   765
         Left            =   40
         TabIndex        =   85
         Top             =   230
         Width           =   1320
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   14
            Left            =   40
            TabIndex        =   87
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
         Begin VB.Label lbl3FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   200
            TabIndex        =   89
            Tag             =   "small"
            Top             =   145
            Width           =   195
         End
         Begin VB.Label lbl3FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   200
            TabIndex        =   96
            Tag             =   "small"
            Top             =   435
            Width           =   195
         End
      End
   End
   Begin VB.Frame frm2Finales 
      Caption         =   "Halve finalisten"
      Height          =   1095
      Left            =   120
      TabIndex        =   55
      Top             =   7440
      Width           =   3075
      Begin VB.Frame frm2FinWed 
         Height          =   765
         Index           =   0
         Left            =   1560
         TabIndex        =   130
         Tag             =   "small"
         Top             =   230
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   27
            Left            =   600
            TabIndex        =   37
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   26
            Left            =   600
            TabIndex        =   36
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   13
            Left            =   40
            TabIndex        =   99
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
         Begin VB.Label lbl2FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   102
            Tag             =   "small"
            Top             =   145
            Width           =   195
         End
         Begin VB.Label lbl2FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   107
            Tag             =   "small"
            Top             =   435
            Width           =   195
         End
      End
      Begin VB.Frame frm2FinWed 
         Height          =   765
         Index           =   4
         Left            =   120
         TabIndex        =   126
         Top             =   230
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   600
            TabIndex        =   35
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   24
            Left            =   600
            TabIndex        =   34
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl2FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   129
            Tag             =   "small"
            Top             =   435
            Width           =   195
         End
         Begin VB.Label lbl2FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   128
            Tag             =   "small"
            Top             =   145
            Width           =   195
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   12
            Left            =   40
            TabIndex        =   127
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
      End
   End
   Begin VB.Frame frm4Finales 
      Caption         =   "Kwart finalisten"
      Height          =   1095
      Left            =   120
      TabIndex        =   54
      Top             =   6360
      Width           =   6015
      Begin VB.Frame frm4FinWed 
         Height          =   765
         Index           =   3
         Left            =   4440
         TabIndex        =   122
         Top             =   230
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   600
            TabIndex        =   33
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   22
            Left            =   600
            TabIndex        =   32
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl4FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   125
            Tag             =   "small"
            Top             =   413
            Width           =   195
         End
         Begin VB.Label lbl4FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   124
            Tag             =   "small"
            Top             =   143
            Width           =   195
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   11
            Left            =   40
            TabIndex        =   123
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
      End
      Begin VB.Frame frm4FinWed 
         Height          =   765
         Index           =   2
         Left            =   3000
         TabIndex        =   118
         Top             =   230
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   21
            Left            =   600
            TabIndex        =   31
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   20
            Left            =   600
            TabIndex        =   30
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl4FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   121
            Tag             =   "small"
            Top             =   413
            Width           =   195
         End
         Begin VB.Label lbl4FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   120
            Tag             =   "small"
            Top             =   143
            Width           =   195
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   10
            Left            =   40
            TabIndex        =   119
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
      End
      Begin VB.Frame frm4FinWed 
         Height          =   765
         Index           =   1
         Left            =   1560
         TabIndex        =   114
         Top             =   230
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   19
            Left            =   600
            TabIndex        =   29
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   18
            Left            =   600
            TabIndex        =   28
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl4FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   117
            Tag             =   "small"
            Top             =   413
            Width           =   195
         End
         Begin VB.Label lbl4FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   116
            Tag             =   "small"
            Top             =   143
            Width           =   195
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   9
            Left            =   40
            TabIndex        =   115
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
      End
      Begin VB.Frame frm4FinWed 
         Height          =   765
         Index           =   0
         Left            =   120
         TabIndex        =   110
         Top             =   230
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   17
            Left            =   600
            TabIndex        =   27
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   600
            TabIndex        =   26
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   8
            Left            =   40
            TabIndex        =   113
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
         Begin VB.Label lbl4FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   112
            Tag             =   "small"
            Top             =   143
            Width           =   195
         End
         Begin VB.Label lbl4FinCode 
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   111
            Tag             =   "small"
            Top             =   413
            Width           =   195
         End
      End
   End
   Begin VB.Frame frm8Finales 
      Caption         =   "Achtste finalisten"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   53
      Top             =   4440
      Width           =   6015
      Begin VB.Frame frm8FinWed 
         Height          =   765
         Index           =   7
         Left            =   4440
         TabIndex        =   105
         Top             =   1000
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   600
            TabIndex        =   25
            Tag             =   "cmbsmall"
            Top             =   415
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   600
            TabIndex        =   24
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   7
            Left            =   40
            TabIndex        =   109
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   108
            Tag             =   "small"
            Top             =   150
            Width           =   315
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   106
            Tag             =   "small"
            Top             =   420
            Width           =   315
         End
      End
      Begin VB.Frame frm8FinWed 
         Height          =   765
         Index           =   6
         Left            =   3000
         TabIndex        =   98
         Top             =   1000
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   600
            TabIndex        =   23
            Tag             =   "cmbsmall"
            Top             =   415
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   600
            TabIndex        =   22
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   6
            Left            =   40
            TabIndex        =   104
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   103
            Tag             =   "small"
            Top             =   150
            Width           =   315
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   100
            Tag             =   "small"
            Top             =   420
            Width           =   315
         End
      End
      Begin VB.Frame frm8FinWed 
         Height          =   765
         Index           =   5
         Left            =   1560
         TabIndex        =   93
         Top             =   1000
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   600
            TabIndex        =   21
            Tag             =   "cmbsmall"
            Top             =   415
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   600
            TabIndex        =   20
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   97
            Tag             =   "small"
            Top             =   150
            Width           =   315
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   95
            Tag             =   "small"
            Top             =   420
            Width           =   315
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   5
            Left            =   40
            TabIndex        =   94
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
      End
      Begin VB.Frame frm8FinWed 
         Height          =   765
         Index           =   4
         Left            =   120
         TabIndex        =   88
         Top             =   1000
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   600
            TabIndex        =   19
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   600
            TabIndex        =   18
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   4
            Left            =   40
            TabIndex        =   92
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   91
            Tag             =   "small"
            Top             =   150
            Width           =   315
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   90
            Tag             =   "small"
            Top             =   420
            Width           =   315
         End
      End
      Begin VB.Frame frm8FinWed 
         Height          =   765
         Index           =   3
         Left            =   4440
         TabIndex        =   82
         Top             =   260
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   600
            TabIndex        =   17
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   600
            TabIndex        =   16
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   86
            Tag             =   "small"
            Top             =   420
            Width           =   315
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   84
            Tag             =   "small"
            Top             =   150
            Width           =   315
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   3
            Left            =   40
            TabIndex        =   83
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
      End
      Begin VB.Frame frm8FinWed 
         Height          =   765
         Index           =   2
         Left            =   3000
         TabIndex        =   76
         Top             =   260
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   600
            TabIndex        =   15
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   600
            TabIndex        =   14
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   80
            Tag             =   "small"
            Top             =   420
            Width           =   315
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   78
            Tag             =   "small"
            Top             =   150
            Width           =   315
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   2
            Left            =   40
            TabIndex        =   77
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
      End
      Begin VB.Frame frm8FinWed 
         Height          =   765
         Index           =   1
         Left            =   1560
         TabIndex        =   71
         Top             =   260
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   600
            TabIndex        =   13
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   600
            TabIndex        =   12
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   1
            Left            =   40
            TabIndex        =   75
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   74
            Tag             =   "small"
            Top             =   420
            Width           =   315
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   73
            Tag             =   "small"
            Top             =   150
            Width           =   315
         End
      End
      Begin VB.Frame frm8FinWed 
         Height          =   765
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   260
         Width           =   1440
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   600
            TabIndex        =   11
            Tag             =   "cmbsmall"
            Top             =   435
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFinTeam 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   600
            TabIndex        =   10
            Tag             =   "cmbsmall"
            Top             =   145
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   69
            Tag             =   "small"
            Top             =   450
            Width           =   315
         End
         Begin VB.Label lbl8FinCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1A"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Tag             =   "small"
            Top             =   180
            Width           =   400
         End
         Begin VB.Label lblWedNum 
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   0
            Left            =   40
            TabIndex        =   66
            Tag             =   "xsmall"
            Top             =   300
            Width           =   135
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Uitslagen"
      Height          =   9330
      Left            =   6240
      TabIndex        =   51
      Top             =   480
      Width           =   7800
      Begin MSDataGridLib.DataGrid grdMatches 
         Height          =   9015
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   15901
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         WrapCellPointer =   -1  'True
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "matchOrder"
            Caption         =   "ID"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "matchdate"
            Caption         =   "datum"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "matchnumber"
            Caption         =   "nr"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "matchDesc"
            Caption         =   "Wedstrijd"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "r1"
            Caption         =   "rust"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "r2"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "e1"
            Caption         =   "eind"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "e2"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "tt"
            Caption         =   "toto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1043
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   2
         BeginProperty Split0 
            MarqueeStyle    =   2
            ScrollBars      =   0
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            Size            =   2
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1604,976
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   360
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   3000,189
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   390,047
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   390,047
            EndProperty
         EndProperty
         BeginProperty Split1 
            MarqueeStyle    =   2
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               DividerStyle    =   0
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   450,142
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               DividerStyle    =   0
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   450,142
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   345,26
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   404,787
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frmGroups 
      Caption         =   "Groepsfase"
      Height          =   2535
      Left            =   120
      TabIndex        =   50
      Top             =   1920
      Width           =   3135
      Begin VB.TextBox txtGroupTeamPos 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   6
         Text            =   "1"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtGroupTeamPos 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   2280
         TabIndex        =   9
         Text            =   "1"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtGroupTeamPos 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   8
         Text            =   "1"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtGroupTeamPos 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Text            =   "1"
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox lstGroupCodes 
         Height          =   2010
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Shape Shape1 
         Height          =   1575
         Left            =   960
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblGroupTeam 
         AutoSize        =   -1  'True
         Caption         =   "GroupTeam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   64
         Top             =   660
         Width           =   825
      End
      Begin VB.Label lblGroupTeam 
         AutoSize        =   -1  'True
         Caption         =   "GroupTeam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   63
         Top             =   1740
         Width           =   825
      End
      Begin VB.Label lblGroupTeam 
         AutoSize        =   -1  'True
         Caption         =   "GroupTeam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   62
         Top             =   1380
         Width           =   825
      End
      Begin VB.Label lblGroupTeam 
         AutoSize        =   -1  'True
         Caption         =   "GroupTeam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   61
         Top             =   1020
         Width           =   825
      End
   End
   Begin VB.Frame frmCompetitior 
      Caption         =   "Deelnemer"
      Height          =   1335
      Left            =   120
      TabIndex        =   47
      Top             =   480
      Width           =   6015
      Begin VB.TextBox txtNickName 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Betaald"
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblPhone 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4560
         TabIndex        =   131
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Poolnaam"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2280
         TabIndex        =   48
         Top             =   360
         Width           =   2205
      End
   End
   Begin VB.CommandButton btnDeleteForm 
      Caption         =   "Verwijder formuler"
      Height          =   375
      Left            =   4320
      TabIndex        =   101
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox cmbAddresses 
      Height          =   315
      Left            =   9000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton btnOpenAddressForm 
      Caption         =   "Adressenlijst"
      Height          =   375
      Left            =   11760
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox cmbCompetitors 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Doelpunten"
      Height          =   255
      Left            =   6360
      TabIndex        =   139
      Top             =   9960
      Width           =   4215
   End
   Begin VB.Label lblPoolInfo 
      Alignment       =   2  'Center
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   195
      TabIndex        =   138
      Top             =   9840
      Width           =   5865
   End
   Begin VB.Label Label2 
      Caption         =   "Formulier toevoegen van"
      Height          =   255
      Left            =   7080
      TabIndex        =   46
      Top             =   150
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Open formulier"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmPoolForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Private grdRs As ADODB.Recordset

Dim cn As ADODB.Connection

Dim thisNickname As String
Dim txtBoxValueSaved As Integer
Dim txtBoxTextSaved As String

Dim firstFinalMatchNumber As Integer

Dim thirdPlace As Boolean

Private Sub btnClose_Click()
  Unload Me
End Sub

Private Sub btnDeleteForm_Click()
Dim sqlstr As String
  Dim msg As String
  msg = "DIT IS DEFINITIEF!"
  msg = msg & vbNewLine & "Weet je zeker dat pool " & Me.txtNickName
  msg = msg & " van " & Me.lblName
  msg = msg & vbNewLine & "moet worden verwijderd?"
  msg = msg & vbNewLine & "(kan dus niet worden hersteld...)"
  If MsgBox(msg, vbYesNo + vbQuestion, "Formulier verwijderen") = vbYes Then
    sqlstr = "Delete from tblPrediction_MatchResults WHERE competitorPoolID = " & thisPoolForm
    cn.Execute sqlstr
    sqlstr = "Delete from tblPredictionTopscorers WHERE competitorPoolID = " & thisPoolForm
    cn.Execute sqlstr
    sqlstr = "Delete from tblPrediction_Finals WHERE competitorPoolID = " & thisPoolForm
    cn.Execute sqlstr
    sqlstr = "Delete from tblPrediction_Numbers WHERE competitorPoolID = " & thisPoolForm
    cn.Execute sqlstr
    sqlstr = "Delete from tblPredictionGroupresults WHERE competitorPoolID = " & thisPoolForm
    cn.Execute sqlstr
    sqlstr = "Delete from tblCompetitorPools  WHERE competitorPoolID = " & thisPoolForm
    cn.Execute sqlstr
    
    thisPoolForm = 1
    fillCompetitorCmb
    If Me.cmbCompetitors.ListCount > 0 Then
      Me.cmbCompetitors.ListIndex = 0
    End If
  End If
End Sub

Private Sub btnNewPlayer_Click()
    frmPlayersNew.Show 1
    DoEvents
    'refill the players combo
    
    fillPlayersCombo
End Sub

Sub fillPlayersCombo()
  Dim sqlstr As String
  sqlstr = "SELECT peopleID, nickName from tblPeople WHERE peopleID IN "
  sqlstr = sqlstr & "(SELECT playerID from tblTeamPlayers WHERE tournamentID = " & thisTournament & ")"
  sqlstr = sqlstr & " ORDER BY nickName"
  FillCombo Me.cmbTopScorer, sqlstr, cn, "nickName", "peopleID"
  DoEvents
End Sub

Private Sub btnOpenAddressForm_Click()
  Dim i As Integer
  frmAddress.Show 1
  fillAddressCmb
  For i = 0 To Me.cmbAddresses.ListCount - 1
    If Me.cmbAddresses.ItemData(i) = thisAddress Then
      Exit For
    End If
  Next
  If i < Me.cmbAddresses.ListCount Then
    Me.cmbAddresses = Me.cmbAddresses.List(i)
  End If
End Sub

Private Sub cmbAddresses_Click()
  'select an address to add a form for
  Dim adoCmd As ADODB.Command
  Set adoCmd = New ADODB.Command
  Dim sqlstr As String
  
  Dim msg As String
  Dim nickName As String
  Dim adrID As Long
  Dim i As Integer
  
  With Me.cmbAddresses
    adrID = .ItemData(.ListIndex)
    sqlstr = "Select addressID, nickName from tblCompetitorPools "
    sqlstr = sqlstr & "WHERE addressID = ?" '& adrID
    sqlstr = sqlstr & " AND poolid = ?" '& thisPool
    With adoCmd
      .ActiveConnection = cn
      .CommandType = adCmdText
      .CommandText = sqlstr
      .Prepared = True
      .Parameters.Append .CreateParameter("adrID", adInteger, adParamInput)
      .Parameters.Append .CreateParameter("pool", adInteger, adParamInput)
      .Parameters("adrID").value = adrID
      .Parameters("pool").value = thisPool
      Set rs = .Execute
    End With
    'check if we really want to add as new poolform
    If Not rs.EOF Then
      msg = .Text & " heeft al " & rs.RecordCount & " pool" & IIf(rs.RecordCount > 1, "s", "")
      msg = msg & vbNewLine & "Nog een pool toevoegen?"
    Else
      msg = "Pool toevoegen voor " & .Text & "?"
    End If
    nickName = getAddressInfo(adrID, "firstName", cn) & "." & Left(getAddressInfo(adrID, "lastName", cn), 4)
    If rs.RecordCount > 0 Then
      nickName = nickName & Trim(Str(rs.RecordCount)) + 1
    End If
    rs.Close
    If MsgBox(msg, vbOKCancel + vbQuestion, "Nieuw formulier") = vbOK Then
      sqlstr = "INSERT INTO tblCompetitorPools (addressID, poolID, nickName, predictionTeam1, predictionTeam2, predictionTeam3, predictionTeam4)"
      sqlstr = sqlstr & " VALUES (" & .ItemData(.ListIndex) & ", " & thisPool & ", '" & nickName & "', 0, 0, 0, 0)"
      cn.Execute sqlstr
      'retrieve new competitorPool ID
      thisPoolForm = cn.Execute("Select @@identity from tblCompetitorPools")(0)
      'fill in default cq emtpy values in poolCompetitor tables
      
      'update competitor combo
      fillCompetitorCmb
      
      fillDefaultPredictions cn
    
    Else
    End If
  End With
End Sub

Private Sub cmbCompetitors_Click()
  'get address data from this competitor
  Dim adrID As Long
  thisPoolForm = Me.cmbCompetitors.ItemData(Me.cmbCompetitors.ListIndex)
  adrID = getCompetitorPoolInfo(thisPoolForm, "addressID", cn)
  Me.lblName = getAddressInfo(adrID, "fullName", cn)
  Me.lblEmail = getAddressInfo(adrID, "email", cn)
  Me.lblPhone = nz(getAddressInfo(adrID, "telephone", cn), "")
  Me.txtNickName = Me.cmbCompetitors.Text
  thisNickname = Me.txtNickName
  'get the form data
  getCompetitorGroups Me.lstGroupCodes
  getCompetitorFinals
  getCompetitorNumbers
  getCompetitorRanking
  getCompetitorTopScorer
  txtBoxValueSaved = 0
  txtBoxTextSaved = ""
  'fill match grid
  fillMatchGrid
  
  updatePoolInfo
  
  Me.lstGroupCodes = "A"
  'set focus on first field: nickname
  On Error Resume Next 'skip this if form is being loaded
    Me.txtNickName.SetFocus
  On Error GoTo 0
  
End Sub

Sub updatePoolInfo()
  Me.lblPoolInfo.Caption = "Pool: " & Me.cmbCompetitors.ListIndex + 1 & " van " & Me.cmbCompetitors.ListCount
End Sub

Sub getCompetitorTopScorer()
  'read the topscorer(s) on this poolform
  Dim sqlstr As String
  Dim i As Integer
  Dim J As Integer
  sqlstr = "Select topscorerPlayerID, topscorerGoals from tblPredictionTopscorers "
  sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
  rs.Open sqlstr, cn, adOpenForwardOnly, adLockReadOnly
  Me.cmbTopScorer.Text = ""
  Me.updnTopScorerDP = 0
  Me.txtTopScorerDP = 0
  If Not rs.EOF Then
    If rs!topScorerPlayerID <> 0 Then
      For J = 0 To Me.cmbTopScorer.ListCount - 1
        If Me.cmbTopScorer.ItemData(J) = rs!topScorerPlayerID Then
          Me.cmbTopScorer.ListIndex = J
          Me.updnTopScorerDP = rs!topScorergoals
          Me.txtTopScorerDP = Me.updnTopScorerDP
          Exit For
        End If
      Next
    End If
    rs.Close
  Else
    rs.Close
  End If
End Sub

Sub getCompetitorGroups(grp As String)
  Dim sqlstr As String
  Dim i As Integer
  sqlstr = "Select * from tblPredictionGroupResults "
  sqlstr = sqlstr & " where competitorPoolID = " & thisPoolForm
  sqlstr = sqlstr & " and ucase(groupLetter) = '" & grp & "'"
  rs.Open sqlstr, cn
  For i = 0 To 3
    If Not rs.EOF Then
      Me.txtGroupTeamPos(i).Text = rs("predictionGroupPosition" & Format(Str(i + 1), "0"))
    Else
      Me.txtGroupTeamPos(i).Text = ""
    End If
  Next
  rs.Close
End Sub

Sub getCompetitorRanking()
Dim sqlstr As String
Dim i As Integer
Dim J As Integer
Dim idx As Integer
Dim ctl As ComboBox
Dim team(3) As String
  
  sqlstr = "Select p.predictionTeam1 AS T1, p.predictionTeam2 AS T2, p.predictionTeam3 AS T3, p.predictionTeam3 AS T4"
  sqlstr = sqlstr & " from tblCompetitorPools P WHERE poolid = " & thisPool
  sqlstr = sqlstr & " AND competitorPoolID = " & thisPoolForm
  rs.Open sqlstr, cn
  
  If Not rs.EOF Then
    For i = 0 To Me.cmbRanking.UBound 'for each ranking combobox
      Set ctl = Me.cmbRanking(i)
      ctl.ListIndex = -1
      setCombo ctl, rs.Fields("T" & Format(i + 1, "0"))
'      For j = 0 To ctl.ListCount - 1
'        If ctl.ItemData(j) = rs.Fields("T" & Format(i + 1, "0")) Then
'          ctl.ListIndex = j
'          Exit For
'        End If
'      Next
    Next
  End If
  rs.Close
End Sub

Sub getCompetitorFinals()
'read the final teams on this form for the competitor
Dim sqlstr As String
Dim i As Integer
Dim J As Integer
Dim matchAantal As Integer
Dim teamA As String
Dim teamB As String

  sqlstr = "Select matchOrder, teamNameA as idA, teamNameB as idB "
  sqlstr = sqlstr & " from tblPrediction_Finals"
  sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
  rs.Open sqlstr, cn
  
  If rs.EOF Then
    rs.Close
    Exit Sub
  End If
  matchAantal = getMatchCount(0, cn)
  For i = 0 To matchAantal - firstFinalMatchNumber
    rs.MoveFirst
    rs.Find "matchOrder = " & getMatchNumber(i + firstFinalMatchNumber, cn)
    'If i = 8 Then Stop
    If Not rs.EOF And Not rs.BOF Then
      'found it!
'      Me.cmbFinTeam(i * 2).ListIndex = -1
      setCombo Me.cmbFinTeam(i * 2), rs!idA
      setCombo Me.cmbFinTeam(i * 2 + 1), rs!idB
    Else
      Debug.Print "No record " & i + firstFinalMatchNumber
    End If
  Next
  rs.Close
End Sub

Sub getCompetitorNumbers()
  Dim sqlstr As String
  Dim i As Integer
  sqlstr = "Select predictionTypeID as ID, predictionNumber as num from tblPrediction_Numbers Where competitorPoolID = " & thisPoolForm
  rs.Open sqlstr, cn
  Do While Not rs.EOF
    For i = 0 To Me.txtAantal.UBound
      If Me.txtAantal(i).Tag = rs!id Then
        Me.txtAantal(i).Text = rs!Num
        If i > 0 Then
          Me.txtAantal(i).TabIndex = Me.txtAantal(i - 1).TabIndex + 1
        End If
        Exit For
      End If
    Next
    rs.MoveNext
  Loop
  rs.Close
End Sub

Private Sub cmbFinTeam_GotFocus(Index As Integer)
'  SelectAllText Me.cmbFinTeam(Index)
  If Me.cmbFinTeam(Index).ListIndex > -1 Then
    txtBoxValueSaved = Me.cmbFinTeam(Index).ItemData(Me.cmbFinTeam(Index).ListIndex)
  End If
End Sub

Private Sub cmbFinTeam_LostFocus(Index As Integer)
Dim sqlstr As String
Dim teamFld As String
Dim matchNr As Integer
  If Me.cmbFinTeam(Index).ListIndex = -1 Then Exit Sub
  If txtBoxValueSaved = Me.cmbFinTeam(Index).ItemData(Me.cmbFinTeam(Index).ListIndex) Then
    Exit Sub
  Else
    'save data
    'left or right team
    If Int(Index / 2) = Index / 2 Then 'even number -> left team
      teamFld = "teamNameA"
      matchNr = Index / 2 + firstFinalMatchNumber
    Else
      teamFld = "teamNameB"
      matchNr = (Index - 1) / 2 + firstFinalMatchNumber
    End If
    sqlstr = "UPDATE tblPrediction_Finals SET " & teamFld
    sqlstr = sqlstr & " = " & Me.cmbFinTeam(Index).ItemData(Me.cmbFinTeam(Index).ListIndex)
    sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
    sqlstr = sqlstr & " AND matchOrder = " & getMatchOrder(matchNr, cn)
    cn.Execute sqlstr

  End If
  
End Sub

Private Sub cmbRanking_GotFocus(Index As Integer)
  With Me.cmbRanking(Index)
  If .ListIndex > -1 Then
    txtBoxValueSaved = .ItemData(.ListIndex)
  End If
  End With
End Sub

Private Sub cmbRanking_LostFocus(Index As Integer)
  With Me.cmbRanking(Index)
  If .ListIndex = -1 Then Exit Sub
  saveRanking Index, .ItemData(.ListIndex)
'  If txtBoxValueSaved <> .ItemData(.ListIndex) Then
'    'save data
'  End If
  End With
End Sub

Sub saveRanking(Index As Integer, teamCode As Long)
'save the ranking
Dim sqlstr As String
Dim rankingFld As String
    rankingFld = "predictionTeam" & Format(Index + 1, "0")
    sqlstr = "UPDATE tblCompetitorPools SET " & rankingFld
    sqlstr = sqlstr & " = " & teamCode
    sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
    cn.Execute sqlstr
End Sub

Private Sub cmbTopScorer_GotFocus()
  If Me.cmbTopScorer.ListIndex > -1 Then
    txtBoxValueSaved = Me.cmbTopScorer.ItemData(Me.cmbTopScorer.ListIndex)
  End If
End Sub

Private Sub cmbTopScorer_LostFocus()
'save topscorer
  If Me.cmbTopScorer.ListIndex = -1 Then Exit Sub
  If txtBoxValueSaved <> Me.cmbTopScorer.ItemData(Me.cmbTopScorer.ListIndex) Then
    updateTopScorer
  End If
End Sub

Sub updateTopScorer()
  Dim sqlstr As String
  'If Not Me.txtTopScorerDP Then Me.txtTopScorerDP = "0"
  If Me.cmbTopScorer.ListIndex <> -1 Then
    sqlstr = "UPDATE tblPredictionTopScorers SET topscorerPlayerID = " & Me.cmbTopScorer.ItemData(Me.cmbTopScorer.ListIndex)
    sqlstr = sqlstr & ", topScorerGoals = " & Me.updnTopScorerDP
    sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
    sqlstr = sqlstr & " AND topscorerPosition = 1"
    cn.Execute sqlstr
  End If
End Sub

Private Sub Form_Activate()
 'set the focus to nickname field
  Me.txtNickName.SetFocus
End Sub

Private Sub Form_Load()
  Set cn = New ADODB.Connection
  With cn
      .ConnectionString = lclConn
      .CursorLocation = adUseClient
      .Open
  End With
  
  Set rs = New ADODB.Recordset
  
  thirdPlace = getTournamentInfo("tournamentThirdPlace", cn)
  firstFinalMatchNumber = getFirstFinalMatchNumber(cn)
  
  UnifyForm Me
  initForm
  centerForm Me
  'if there are forms then open first participant
  If Me.cmbCompetitors.ListCount > 0 Then
    Me.cmbCompetitors.ListIndex = 0
  End If
  
End Sub

Sub initForm()
  Dim i As Integer
  Dim J As Integer
  Dim groupCount As Integer
  Dim sqlstr As String
  Dim wedNumFinal As Integer
  
  Me.lblPoolInfo.Caption = ""
  
  fillAddressCmb
  fillCompetitorCmb
  
  'fill groups listbox
  groupCount = getTournamentInfo("tournamentGroupCount", cn)
  For i = 0 To groupCount - 1
    Me.lstGroupCodes.AddItem Chr(65 + i)
  Next
  'fill in first group teamnames
  Me.lstGroupCodes.ListIndex = 0 'will trigger lstGroupCodes_Click
  'read all the matches in a recordset
  sqlstr = "Select * from tblTournamentSchedule WHERE tournamentID = " & thisTournament
  rs.Open sqlstr, cn
  
  For i = 0 To 13
    Me.lblWedNum(i).Caption = firstFinalMatchNumber + i
    'update finals teamcodes
    rs.MoveFirst
    rs.Find "matchNumber = " & firstFinalMatchNumber + i
    If Not rs.EOF And Not rs.BOF Then
      If i <= 7 Then
        Me.lbl8FinCode(i * 2).Caption = rs!matchteamA & ":"
        Me.lbl8FinCode((i * 2) + 1).Caption = rs!matchteamB & ":"
        Me.lbl8FinCode(i * 2).Left = 100
        Me.lbl8FinCode(i * 2).width = 520
        Me.lbl8FinCode(i * 2).Top = 180
        Me.lbl8FinCode((i * 2) + 1).Left = Me.lbl8FinCode(i * 2).Left
        Me.lbl8FinCode((i * 2) + 1).width = Me.lbl8FinCode(i * 2).width
        Me.lbl8FinCode((i * 2) + 1).Top = 450
      ElseIf i <= 11 Then
        Me.lbl4FinCode((i - 8) * 2).Caption = rs!matchteamA & ":"
        Me.lbl4FinCode(((i - 8) * 2) + 1).Caption = rs!matchteamB & ":"
        Me.lbl4FinCode((i - 8) * 2).Top = 180
        Me.lbl4FinCode(((i - 8) * 2) + 1).Top = 450
      Else
        Me.lbl2FinCode((i - 12) * 2).Caption = rs!matchteamA & ":"
        Me.lbl2FinCode((i - 12) * 2).Top = 180
        Me.lbl2FinCode(((i - 12) * 2) + 1).Caption = rs!matchteamB & ":"
        Me.lbl2FinCode(((i - 12) * 2) + 1).Top = 450
      End If
    End If
  Next
  For i = 0 To Me.cmbFinTeam.UBound
    Me.cmbFinTeam(i).Left = 640
  Next
  wedNumFinal = 14
  If thirdPlace Then 'load frame fields for 3rd place match
    Me.lblWedNum(14) = firstFinalMatchNumber + wedNumFinal
    rs.MoveFirst
    rs.Find "matchNumber = " & firstFinalMatchNumber + wedNumFinal
    If Not rs.EOF And Not rs.BOF Then
      Me.lbl3FinCode(0).Caption = rs!matchteamA & ":"
      Me.lbl3FinCode(0).Top = 180
      Me.lbl3FinCode(1).Caption = rs!matchteamB & ":"
      Me.lbl3FinCode(1).Top = 450
    End If
    'combo's for 3rd plce matches
    i = Me.cmbFinTeam.UBound + 1
    Load Me.cmbFinTeam(i)
    With Me.cmbFinTeam(i)
      Set .Container = Me.frm3FinWed
      .Left = 520
      .Top = 145
      .TabIndex = Me.cmbFinTeam(i - 1).TabIndex + 1
    End With
    i = i + 1
    Load Me.cmbFinTeam(i)
    With Me.cmbFinTeam(i)
      Set .Container = Me.frm3FinWed
      .Left = 520
      .Top = 435
      .TabIndex = Me.cmbFinTeam(i - 1).TabIndex + 1
    End With
    
    'also load ranking cmb's
    For J = 0 To 1
      i = Me.cmbRanking.UBound + 1
      Load Me.lblRanking(i)
      With Me.lblRanking(i)
        Set .Container = Me.frmRanking
        .Top = Me.lblRanking(i - 2).Top
        .Left = Me.cmbRanking(i - 2).Left + Me.cmbRanking(i - 2).width + 105
        .Caption = val(Me.lblRanking(i - 1).Caption) + 1
        .Visible = True
      End With
      Load Me.cmbRanking(i)
      With Me.cmbRanking(i)
        Set .Container = Me.frmRanking
        .Top = Me.cmbRanking(i - 2).Top
        .Left = Me.lblRanking(i).Left + 120
        .TabIndex = Me.cmbRanking(i - 1).TabIndex + 1
        .Visible = True
      End With
    Next
    wedNumFinal = wedNumFinal + 1
  Else
    'wider ranking cmb's , why not
    Me.cmbRanking(0).width = 2600
    Me.cmbRanking(1).width = 2600
    Me.frm3Finales.Enabled = False
    Me.lbl3FinCode(0) = ""
    Me.lbl3FinCode(1) = ""
    Me.lblWedNum(14) = ""
    ' 3rd place situation done
  End If
  Me.lblWedNum(15) = firstFinalMatchNumber + wedNumFinal
  rs.MoveFirst
  rs.Find "matchNumber = " & firstFinalMatchNumber + wedNumFinal
  If Not rs.EOF And Not rs.BOF Then
    Me.lbl1FinCode(0).Caption = rs!matchteamA & ":"
    Me.lbl1FinCode(0).Top = 180
    Me.lbl1FinCode(1).Caption = rs!matchteamB & ":"
    Me.lbl1FinCode(1).Top = 450
  End If
  rs.Close
  'combo's for final
  i = Me.cmbFinTeam.UBound + 1
  Load Me.cmbFinTeam(i)
  With Me.cmbFinTeam(i)
    Set .Container = Me.frm1FinWed
    .Left = 540
    .Top = 145
    .TabIndex = Me.cmbFinTeam(i - 1).TabIndex + 1
  End With
  i = i + 1
  Load Me.cmbFinTeam(i)
  With Me.cmbFinTeam(i)
    Set .Container = Me.frm1FinWed
    .Left = 540
    .Top = 435
    .TabIndex = Me.cmbFinTeam(i - 1).TabIndex + 1
  End With
  
  'fill the players combo
  sqlstr = "SELECT peopleID, nickName from tblPeople WHERE peopleID IN "
  sqlstr = sqlstr & "(SELECT playerID from tblTeamPlayers WHERE tournamentID = " & thisTournament & ")"
  sqlstr = sqlstr & " ORDER BY nickName"
  FillCombo Me.cmbTopScorer, sqlstr, cn, "nickName", "peopleID"
  
  'Update the prediction numbers labels
  getPoolPredictionNumbers
  
  ' update final team combo's
  For i = 0 To Me.cmbFinTeam.UBound
    sqlstr = "Select teamNameID as ID, teamShortName as name from tblTeamNames WHERE"
    sqlstr = sqlstr & " teamNameId in (select teamID from tblTournamentTeamCodes where tournamentID = " & thisTournament & ")"
    sqlstr = sqlstr & " ORDER BY teamShortName"
    FillCombo Me.cmbFinTeam(i), sqlstr, cn, "name", "ID"
    Me.cmbFinTeam(i).Visible = True
  Next
  
  'fill teamnames for ranking cmb's
  For i = 0 To Me.cmbRanking.UBound
    sqlstr = "Select teamNameID as ID, teamName as name from tblTeamNames WHERE"
    sqlstr = sqlstr & " teamNameId in (select teamID from tblTournamentTeamCodes where tournamentID = " & thisTournament & ")"
    sqlstr = sqlstr & " ORDER BY teamName"
    FillCombo Me.cmbRanking(i), sqlstr, cn, "name", "ID"
    Me.cmbRanking(i).Visible = True
  Next
  
  DoEvents
End Sub

Sub getPoolPredictionNumbers()
'get the 'numbers' predictions
  Dim sqlstr As String
  Dim i As Integer
  Dim txt As String
  sqlstr = "Select pointTypeID as ID, pointDescrShort as descr from tblPointTypes Where pointTypeCategory = 6"
  sqlstr = sqlstr & " AND pointTypeID IN (Select pointTypeId from tblPoolPoints WHERE poolID = " & thisPool & ")"
  sqlstr = sqlstr & " ORDER BY pointTypeListOrder"
  rs.Open sqlstr, cn
  i = 0
  Do While Not rs.EOF
    If i > Me.lblAantal.UBound Then
      Load Me.lblAantal(i)
      Load Me.txtAantal(i)
      Me.lblAantal(i).Top = Me.lblAantal(i - 1).Top + Me.txtAantal(0).Height + 10
      Me.txtAantal(i).Top = Me.txtAantal(i - 1).Top + Me.txtAantal(0).Height + 10
      Me.txtAantal(i).TabIndex = Me.txtAantal(i - 1).TabIndex + 1
      Me.txtAantal(i).Height = Me.txtAantal(0).Height
    End If
    Me.txtAantal(i).Tag = rs!id 'save the ID in the tag for future reference
    txt = rs!descr
    If Left(rs!descr, 6) = "aantal" Then
      txt = Mid(rs!descr, 8)
    End If
    Me.lblAantal(i).Caption = txt
    Me.lblAantal(i).Visible = True
    Me.txtAantal(i).Visible = True
    Me.txtAantal(i).Tag = rs!id
    i = i + 1
    rs.MoveNext
  Loop
  rs.Close
End Sub

Sub fillMatchGrid()
Dim sqlstr As String
Dim i As Integer
Dim J As Integer
'build the recordset sql Updating the records is done in grdMatches_AfterCoEdit
  cn.Execute "Delete from tblPrediction_MatchResults_TMP"
  
  sqlstr = "INSERT INTO tblPrediction_MatchResults_TMP "
  'get the sqlStr for the poolForm matches
  sqlstr = sqlstr & sqlCompetitorMatches(thisPoolForm)
  cn.Execute sqlstr
  
  DoEvents
 ' Debug.Print rs!wedstrijd
  sqlstr = "Select * from tblPrediction_MatchResults_TMP order by matchOrder"
  Set grdRs = New ADODB.Recordset
  With grdRs
    .CursorLocation = adUseClient
    .Open sqlstr, cn, adOpenDynamic, adLockOptimistic
  End With
  Set Me.grdMatches.DataSource = grdRs
  calcInfo
  DoEvents
  'grdRs.Close
End Sub


Sub fillAddressCmb()
  Dim sqlstr As String
  Me.cmbAddresses.Clear
  sqlstr = "Select addressID as ID, trim( "
  sqlstr = sqlstr & "iif(firstname >' ', trim(firstname) & ' ','') & "
  sqlstr = sqlstr & "iif(middlename >' ', trim(middlename) & ' ','') & "
  sqlstr = sqlstr & "iif(lastname >' ', trim(lastname),'') "
  sqlstr = sqlstr & ") as name "
  sqlstr = sqlstr & "from tblAddresses ORDER BY firstname, lastname"
  FillCombo Me.cmbAddresses, sqlstr, cn, "name", "ID"
  If thisAddress <> 0 Then
    'me.cmbAddresses.list
  End If
End Sub

Sub fillCompetitorCmb()
  Dim sqlstr As String
  Dim i As Integer
  Me.cmbCompetitors.Clear
  sqlstr = "Select competitorPoolID as ID, trim(nickName) as nickName "
  sqlstr = sqlstr & "from tblCompetitorPools WHERE poolID = " & thisPool
  sqlstr = sqlstr & " ORDER BY nickName"
  FillCombo Me.cmbCompetitors, sqlstr, cn, "nickName", "ID"
  If thisPoolForm > 0 Then
    For i = 0 To Me.cmbCompetitors.ListCount - 1
      If Me.cmbCompetitors.ItemData(i) = thisPoolForm Then
        Exit For
      End If
    Next
  End If
End Sub

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
    If Not grdRs Is Nothing Then
        'first, check if the state is open, if yes then close it
        On Error Resume Next  'sometimes some weird error happens
        If (grdRs.State And adStateOpen) = adStateOpen Then
            grdRs.Close
        End If
        'set them to nothing
        Set grdRs = Nothing
    End If
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub grdMatches_AfterColEdit(ByVal ColIndex As Integer)
''after a column is editted (left) update the field in the database
Dim warning As Boolean
Dim answer As Boolean
Dim sqlstr As String
Dim i As Integer
  'is it a number?
  With Me.grdMatches
    If Not IsNumeric(.Columns(ColIndex)) Then
      .Columns(ColIndex).value = 0
    End If
    warning = nz(.Columns(4), 0) > 10
    For i = 5 To 7
      warning = warning Or nz(.Columns(i)) > 10
    Next
    warning = warning Or nz(.Columns(8)) > 3
    warning = warning Or nz(.Columns(8)) = 0
    answer = True
    If warning Then
      answer = MsgBox("Klopt deze invoer wel?", vbYesNo + vbQuestion, "Vreemde getallen") = vbYes
    End If
    If answer Then
      sqlstr = "UPDATE tblPrediction_MatchResults SET "
      sqlstr = sqlstr & " htA = " & .Columns(4)
      sqlstr = sqlstr & ", htB = " & .Columns(5)
      sqlstr = sqlstr & ", ftA = " & .Columns(6)
      sqlstr = sqlstr & ", ftB = " & .Columns(7)
      sqlstr = sqlstr & ", tt = " & .Columns(8)
      sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
      sqlstr = sqlstr & " AND matchOrder = " & .Columns(0)
      cn.Execute sqlstr
      calcInfo
    Else  'reset the row
      sqlstr = "Select * from tblPrediction_MatchResults "
      sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
      sqlstr = sqlstr & " AND matchOrder = " & .Columns(0)
      rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
      If Not rs.EOF Then
        .Columns(ColIndex) = rs.Fields(ColIndex - 6)
      End If
      
      rs.Close
    End If
  End With

End Sub

Sub calcInfo()
Dim selstr As String
Dim draws As Integer
Dim goals As Integer
    selstr = "select ftA, ftB from tblPrediction_MatchResults WHERE competitorPoolid= " & thisPoolForm
    rs.Open selstr, cn, adOpenStatic, adLockReadOnly
    Do While Not rs.EOF
        goals = goals + rs!ftA + rs!ftB
        If rs!ftA = rs!ftB Then draws = draws + 1
        rs.MoveNext
    Loop
    Me.lblInfo = "Doelpunten: " & goals & "; Gelijkspel: " & draws
    Me.lblInfo.Visible = True
    rs.Close
End Sub


Private Sub grdMatches_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  calcInfo
  
End Sub

Private Sub lstGroupCodes_Click()
'user clicked on list to get teams for this group
  If Me.lstGroupCodes > "" Then
    fillGroupTeams Me.lstGroupCodes
  End If
  On Error Resume Next 'skip this if form is being loaded
  getCompetitorGroups Me.lstGroupCodes
  Me.txtGroupTeamPos(0).SetFocus
  On Error GoTo 0
  
End Sub

Sub fillGroupTeams(GroupCodes As String)
'get the teams for this group
  Dim sqlstr As String
  Dim i As Integer
  sqlstr = "Select t.teamName as team from tblGroupLayout l INNER JOIN tblTeamNames t on l.teamID = t.teamNameID"
  sqlstr = sqlstr & " WHERE l.tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND l.groupLetter = '" & GroupCodes & "'"
  rs.Open sqlstr, cn
  i = 0
  Do While Not rs.EOF
    Me.lblGroupTeam(i) = rs!team
    rs.MoveNext
    i = i + 1
  Loop
  rs.Close
End Sub

Private Sub txtAantal_GotFocus(Index As Integer)
  SelectAllText Me.txtAantal(Index)
  txtBoxValueSaved = val(Me.txtAantal(Index))
End Sub

Private Sub txtAantal_LostFocus(Index As Integer)
  Dim sqlstr As String
  If txtBoxValueSaved <> val(Me.txtAantal(Index)) Then
    sqlstr = "UPDATE tblPrediction_Numbers SET predictionNumber = "
    sqlstr = sqlstr & val(Me.txtAantal(Index))
    sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
    sqlstr = sqlstr & " AND predictionTypeID = " & val(Me.txtAantal(Index).Tag)
    cn.Execute sqlstr
  End If
End Sub

Private Sub txtGroupTeamPos_GotFocus(Index As Integer)
  SelectAllText Me.txtGroupTeamPos(Index)
  txtBoxValueSaved = val(Me.txtGroupTeamPos(Index))
End Sub

Sub saveGroupTeamPos(Index As Integer)
'save the value
  Dim sqlstr As String
  sqlstr = "UPDATE tblPredictionGroupResults SET predictionGroupPosition"
  sqlstr = sqlstr & Format(Index + 1, "0") & " = " & val(Me.txtGroupTeamPos(Index))
  sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
  sqlstr = sqlstr & " AND groupLetter = '" & Me.lstGroupCodes & "'"
  cn.Execute sqlstr
End Sub

Private Sub txtGroupTeamPos_KeyPress(Index As Integer, KeyAscii As Integer)
  If Not KeyAscii = 8 And Not (KeyAscii >= 49 And KeyAscii <= 52) Then
    KeyAscii = 0
  End If
  
End Sub

Private Sub txtGroupTeamPos_LostFocus(Index As Integer)
    If txtBoxValueSaved <> val(Me.txtGroupTeamPos(Index)) Then
      saveGroupTeamPos Index
    End If
    If Index >= 3 Then
      If Me.lstGroupCodes = Me.lstGroupCodes.List(Me.lstGroupCodes.ListCount - 1) Then
        'if last group is entered jumnp to finals block
        Me.cmbFinTeam(0).SetFocus
      Else
        Me.lstGroupCodes.ListIndex = Me.lstGroupCodes.ListIndex + 1
      End If
    End If
End Sub

Private Sub txtNickName_GotFocus()
  SelectAllText Me.txtNickName
  txtBoxTextSaved = Me.txtNickName.Text
End Sub

Private Sub txtNickName_LostFocus()
  'save the nickname
  Dim sqlstr As String
  If Me.txtNickName <> thisNickname Then
    sqlstr = "UPDATE tblCompetitorPools SET nickName = '" & Trim(Me.txtNickName.Text) & "'"
    sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
    cn.Execute sqlstr
    
    'update the combo
    fillCompetitorCmb
    thisNickname = Me.txtNickName
  End If
  
End Sub

Private Sub txtTopScorerDP_Change()
  Me.updnTopScorerDP = val(Me.txtTopScorerDP)
End Sub

Private Sub txtTopScorerDP_GotFocus()
  SelectAllText Me.txtTopScorerDP
  txtBoxValueSaved = Me.updnTopScorerDP
End Sub

Private Sub txtTopScorerDP_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
      KeyAscii = 0
  End If
End Sub

Private Sub updnTopScorerDP_Change()
  Me.txtTopScorerDP = Me.updnTopScorerDP
  updateTopScorer
End Sub
