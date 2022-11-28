VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPoolEdit 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pools"
   ClientHeight    =   5400
   ClientLeft      =   12630
   ClientTop       =   6360
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5790
   Begin VB.Frame frmPrizes 
      Caption         =   "Prijzen"
      Height          =   2415
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   5775
      Begin MSComCtl2.UpDown UpDnPerc 
         Height          =   375
         Index           =   0
         Left            =   3825
         TabIndex        =   30
         Top             =   660
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtPercentage(0)"
         BuddyDispid     =   196610
         BuddyIndex      =   0
         OrigLeft        =   3840
         OrigTop         =   660
         OrigRight       =   4095
         OrigBottom      =   1035
         Increment       =   5
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox txtHighestDayscore 
         DataField       =   "prizeMostDayPoints"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   660
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "€ #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHighestPosition 
         DataField       =   "prizeBestDayPosition"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   1132
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "€ #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtLowestPosition 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   1650
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "€ #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrizeLastOverall 
         DataField       =   "prizeLastOverallPosition"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   2
         EndProperty
         DataSource      =   "dtcPools"
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "€ #,##0.00"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDnPerc 
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   31
         Top             =   660
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtPercentage(1)"
         BuddyDispid     =   196610
         BuddyIndex      =   1
         OrigLeft        =   5400
         OrigTop         =   720
         OrigRight       =   5655
         OrigBottom      =   1095
         Increment       =   5
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDnPerc 
         Height          =   375
         Index           =   2
         Left            =   3825
         TabIndex        =   32
         Top             =   1125
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtPercentage(2)"
         BuddyDispid     =   196610
         BuddyIndex      =   2
         OrigLeft        =   3840
         OrigTop         =   1125
         OrigRight       =   4095
         OrigBottom      =   1500
         Increment       =   5
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDnPerc 
         Height          =   375
         Index           =   3
         Left            =   5280
         TabIndex        =   33
         Top             =   1125
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtPercentage(3)"
         BuddyDispid     =   196610
         BuddyIndex      =   3
         OrigLeft        =   5400
         OrigTop         =   1200
         OrigRight       =   5655
         OrigBottom      =   1575
         Increment       =   5
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.Label txtPercentage 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   37
         Top             =   1125
         Width           =   600
      End
      Begin VB.Label txtPercentage 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   36
         Top             =   1125
         Width           =   585
      End
      Begin VB.Label txtPercentage 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   35
         Top             =   660
         Width           =   600
      End
      Begin VB.Label txtPercentage 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   34
         Top             =   660
         Width           =   585
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Totaal: 100%"
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Laatste plaats"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2760
         TabIndex        =   27
         Top             =   1950
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Onderaan"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1710
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bovenaan"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1192
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Meeste punten"
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   0
         TabIndex        =   24
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "2e"
         Height          =   375
         Left            =   4320
         TabIndex        =   23
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "1e"
         Height          =   375
         Left            =   2880
         TabIndex        =   22
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "4e"
         Height          =   255
         Left            =   4320
         TabIndex        =   21
         Top             =   1185
         Width           =   375
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "3e"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   1192
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Eindstand (% inleg - dagprijzen)"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   3015
      End
      Begin VB.Line Line2 
         X1              =   2640
         X2              =   2640
         Y1              =   360
         Y2              =   2040
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Dagprijzen"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSMask.MaskEdBox txtCosts 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1740
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "€ #,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtPoolName 
      DataSource      =   "dtcPools"
      Height          =   360
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5730
      TabIndex        =   17
      Top             =   4665
      Width           =   5790
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuleren"
         Height          =   495
         Left            =   2880
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Opslaan"
         Height          =   495
         Left            =   1440
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "Sluiten"
         Default         =   -1  'True
         Height          =   495
         Left            =   4320
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker dtpStart 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1740
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   146866177
      CurrentDate     =   43932
   End
   Begin MSComCtl2.DTPicker dtpEind 
      DataSource      =   "dtcPools"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   1740
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   146866177
      CurrentDate     =   43932
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pool gegevens aanpassen"
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Tag             =   "kop"
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inleg "
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pool naam"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   -120
      TabIndex        =   0
      Top             =   780
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5640
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "tot"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inleveren"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblTournament 
      BackStyle       =   0  'Transparent
      Caption         =   "Toernooi"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
End
Attribute VB_Name = "frmPoolEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim editState As Boolean
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnCancel_Click()
    setState False
    'updateForm
End Sub

Private Sub btnSave_Click()
    Dim sqlstr As String
    Dim cmd As ADODB.Command
    sqlstr = "UPDATE tblPools SET poolName = ? "
        sqlstr = sqlstr & ", poolFormsFrom = ?"
        sqlstr = sqlstr & ", poolFormsTill= ?"
        sqlstr = sqlstr & ", poolCost= ?"
        sqlstr = sqlstr & ", prizeHighDayScore= ?"
        sqlstr = sqlstr & ", prizeHighDayPosition=?"
        sqlstr = sqlstr & ", prizeLowDayposition=?"
        sqlstr = sqlstr & ", prizePercentage1=?"
        sqlstr = sqlstr & ", prizePercentage2=?"
        sqlstr = sqlstr & ", prizePercentage3=?"
        sqlstr = sqlstr & ", prizePercentage4=?"
        sqlstr = sqlstr & ", prizeLowFinalPosition=?"
        sqlstr = sqlstr & " WHERE poolID =  " & thisPool
    Set cmd = New ADODB.Command
    If editState Then
        With cmd
            .CommandType = adCmdText
            .CommandText = sqlstr
            Set .ActiveConnection = cn
            .Execute , Array( _
                Me.txtPoolName, CDbl(Me.dtpStart), CDbl(Me.dtpEind), CDbl(Me.txtCosts), CDbl(Me.txtHighestDayscore), _
                CDbl(Me.txtHighestPosition), CDbl(Me.txtLowestPosition), CDbl(Me.txtPercentage(0)), CDbl(Me.txtPercentage(1)), _
                CDbl(Me.txtPercentage(2)), CDbl(Me.txtPercentage(3)), CDbl(Me.txtPrizeLastOverall) _
                ), adCmdText Or adExecuteNoRecords
        End With
    'save the record
        sqlstr = "Select * from tblPools where poolId = " & thisPool
        If thisPool = 0 Then
            MsgBox "Geen poolID actief", vbOKOnly + vbCritical, "Kan niet opslaan"
            Exit Sub
        End If
        
        
'        sqlstr = "UPDATE tblPools SET tournamentId = " & thisTournament
'        sqlstr = sqlstr & ", poolName='" & Me.txtPoolName & "'"
'        sqlstr = sqlstr & ", poolFormsFrom=" & CDbl(Me.dtpStart)
'        sqlstr = sqlstr & ", poolFormsTill=" & CDbl(Me.dtpEind)
'        sqlstr = sqlstr & ", poolCost=" & val(float(Me.txtCosts.Text)) * 100
'        sqlstr = sqlstr & ", prizeHighDayScore=" & val(float(Me.txtHighestDayscore)) * 100
'        sqlstr = sqlstr & ", prizeHighDayPosition=" & val(float(Me.txtHighestPosition)) * 100
'        sqlstr = sqlstr & ", prizeLowDayposition=" & val(float(Me.txtLowestPosition)) * 100
'        sqlstr = sqlstr & ", prizePercentage1=" & val(float(Me.txtPercentage(0))) * 100
'        sqlstr = sqlstr & ", prizePercentage2=" & val(float(Me.txtPercentage(1))) * 100
'        sqlstr = sqlstr & ", prizePercentage3=" & val(float(Me.txtPercentage(2))) * 100
'        sqlstr = sqlstr & ", prizePercentage4=" & val(float(Me.txtPercentage(3))) * 100
'        sqlstr = sqlstr & ", prizeLowFinalPosition=" & val(float(Me.txtPrizeLastOverall)) * 100
'        sqlstr = sqlstr & " WHERE poolID =  " & thisPool
'        cn.Execute sqlstr
        setState False
        DoEvents
    Else
        'set edit mode
        setState True
    End If
End Sub

Private Sub Form_Load()
Dim sqlstr As String
Dim i As Integer
Set rs = New ADODB.Recordset
Dim tnInfo As String
'basis tabel
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblPools where poolid=" & thisPool
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
'fill the fields

    Me.txtPoolName.Text = rs!poolName
    
    tnInfo = getTournamentInfo("description", cn)
    tnInfo = tnInfo & " van " & Format(getTournamentInfo("tournamentStartDate", cn), "ddd d mmm")
    tnInfo = tnInfo & " - " & Format(getTournamentInfo("tournamentEndDate", cn), "ddd d mmm")
    Me.lblTournament = "Toernooi: " & tnInfo
    Me.txtCosts.Text = rs!poolCost
    Me.dtpStart = rs!poolFormsFrom
    Me.dtpEind = rs!poolFormsTill
'    'prizes
    Me.txtHighestDayscore = rs!prizeHighDayScore
    Me.txtHighestPosition = rs!prizeHighDayPosition
    Me.txtLowestPosition = rs!prizeLowDayPosition
    Me.UpDnPerc(0) = rs!prizePercentage1
    Me.UpDnPerc(1) = rs!prizePercentage2
    Me.UpDnPerc(2) = rs!prizePercentage3
    Me.UpDnPerc(3) = rs!prizePercentage4
    Me.txtPrizeLastOverall = rs!prizeLowFinalPosition
'
    
    Me.btnSave.Enabled = Not chkTournamentStarted(cn)
'set Form defaults
    UnifyForm Me

'back color of frame
    Me.frmPrizes.BackColor = Me.BackColor
    Me.lblTotal.BackColor = Me.frmPrizes.BackColor
    'set form state
    setState False
    rs.Close
    Set rs = Nothing
End Sub

Sub setState(edit As Boolean)
Dim ctl As Control
    editState = edit
    With Me
        For Each ctl In .Controls
            If TypeOf ctl Is DTPicker Or _
                TypeOf ctl Is TextBox Or _
                TypeOf ctl Is ComboBox Or _
                TypeOf ctl Is MaskEdBox Or _
                TypeOf ctl Is UpDown Then
                ctl.Enabled = edit
            End If
        Next
'                TypeOf ctl Is DataCombo Or _
        .btnCancel.Visible = edit
        If edit Then
            .btnSave.Caption = "Opslaan"
        Else
            .btnSave.Caption = "Bewerken"
        End If
        .btnClose.Enabled = Not edit
    End With
End Sub

Sub calcTotalPercentage()
'calculate the total of the percentage prizes
    Dim totalPerc As Double
    Dim i As Integer
    totalPerc = 0
    For i = 0 To 3
        totalPerc = totalPerc + Me.UpDnPerc(i).value
'        totalPerc = totalPerc + val(float(Me.txtPercentage(i).Text))
    Next
    Me.lblTotal.Caption = "Totaal " & totalPerc & "%"
    If totalPerc <> 100 Then
        Me.lblTotal.BackColor = &H8080FF
    Else
        Me.lblTotal.BackColor = Me.frmPrizes.BackColor
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
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If

End Sub

Private Sub UpDnPerc_Change(Index As Integer)
    calcTotalPercentage
End Sub
