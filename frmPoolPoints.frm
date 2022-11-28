VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPoolPoints 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Punten toekenning"
   ClientHeight    =   9060
   ClientLeft      =   12975
   ClientTop       =   4365
   ClientWidth     =   5895
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
   ScaleHeight     =   9060
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMarge 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   600
      Width           =   300
   End
   Begin VB.TextBox txtPnt 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid grdPoolpoints 
      Height          =   7455
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   13150
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.CommandButton btnCopyDefaultPoints 
      Caption         =   "Beginwaarden"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   8520
      Width           =   1815
   End
   Begin VB.ComboBox cmbPointTypes 
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   8520
      Width           =   1455
   End
   Begin MSComCtl2.UpDown upDnPnt 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtPnt"
      BuddyDispid     =   196610
      OrigLeft        =   840
      OrigTop         =   480
      OrigRight       =   1095
      OrigBottom      =   855
      Max             =   150
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown uipDnMarge 
      Height          =   375
      Left            =   5460
      TabIndex        =   8
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtMarge"
      BuddyDispid     =   196609
      OrigLeft        =   840
      OrigTop         =   480
      OrigRight       =   1095
      OrigBottom      =   855
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "mrg"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "pnt"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Punten en marges"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   5670
   End
End
Attribute VB_Name = "frmPoolPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Dim cmbGridSelect As Boolean

Dim thisPointType As Long

'if true then don't update combo box after datagrid.refresh
'to prevent recordsource jump back to first record

Private Sub btnClose_Click()
'check if any row has 0 point and delete this record
    Dim msg As String
    Dim qry As ADODB.Command
    Dim sqlstr As String
    Dim rows As Long
    Set qry = New ADODB.Command
    sqlstr = "Delete from tblPoolPoints where poolid = " & thisPool
    sqlstr = sqlstr & " AND pointPointsAward = 0 "
    qry.ActiveConnection = cn
    qry.CommandText = sqlstr
    qry.CommandType = adCmdText
    qry.Execute rows
    If rows Then
        If rows = 1 Then
            msg = "1 rij zonder punten is verwijderd"
        Else
            msg = rows & " rijen zonder punten, zijn verwijderd"
        End If
        MsgBox msg, vbOKOnly + vbInformation, "Pool instellingen"
        'cn.CommitTrans
    End If
    Unload Me
    Set qry = Nothing
End Sub

Sub insertRecord()
Dim qry As ADODB.Command
Dim sqlstr As String
Dim rows As Long

    Set qry = New ADODB.Command
    sqlstr = "insert into tblPoolPoints (poolID, pointTypeId, pointPointsAward, pointPointsMargin) "
    sqlstr = sqlstr & "VALUES ( " & thisPool
    sqlstr = sqlstr & ", " & Me.cmbPointTypes.ItemData(Me.cmbPointTypes.ListIndex)
    sqlstr = sqlstr & ", 0, 0)"
    qry.CommandType = adCmdText
    qry.CommandText = sqlstr
    qry.ActiveConnection = cn
    cn.BeginTrans
    qry.Execute rows
'    MsgBox rows & "  record toegevoegd", vbOKOnly + vbInformation, "Voorspelling opgeslagen"
    cn.CommitTrans
    Set qry = Nothing
End Sub

Private Sub btnCopyDefaultPoints_Click()
  'copy default points table
  copyDefaultPoints
  fillPointsGrid
End Sub

Private Sub saveValues()
 'save the values
  Dim qry As ADODB.Command
  Dim sqlstr As String
  Dim rows As Long
  'parameter uodate qyery
  sqlstr = "UPDATE tblPoolPoints "
  sqlstr = sqlstr & "SET pointPointsAward = ? "
  sqlstr = sqlstr & ", pointPointsMargin = ? "
  sqlstr = sqlstr & " WHERE poolid = " & thisPool
  sqlstr = sqlstr & " AND pointtypeid = " & thisPointType
  
  Set qry = New ADODB.Command
  qry.CommandType = adCmdText
  qry.CommandText = sqlstr
  qry.ActiveConnection = cn
  cn.BeginTrans
  qry.Execute , Array(Me.upDnPnt, Me.uipDnMarge), adCmdText Or adExecuteNoRecords
  
  cn.CommitTrans
  Set qry = Nothing
  fillPointsGrid
End Sub

Private Sub cmbPointTypes_Click()
'select the row in the grid
  Dim i As Integer
  With Me.grdPoolpoints
    For i = 1 To .rows - 1
      If .TextMatrix(i, 2) = Me.cmbPointTypes.List(Me.cmbPointTypes.ListIndex) Then
        .row = i
        .col = 0
        .ColSel = .cols - 1
        Exit Sub
      End If
    Next
    'if we are here then not found, so add new row
    insertRecord
    fillPointsGrid
  End With
End Sub

Private Sub Form_Load()
    Dim sqlstr As String
    
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    
    sqlstr = "Select pointtypeId as id, pointTypeDescription as omschrijving from tblPointtypes "
    If Not getTournamentInfo("tournamentThirdPlace", cn) Then
        'exclude pointtype catgory "Kleine Finale"
        sqlstr = sqlstr & " where pointTypeCategory <> 7"
        'sqlstr = sqlstr & " AND pointtypeid NOT IN (Select pointtypeID from tblPoolPoints WHERE poolID = " & thisPool & ")"
    End If
    
    sqlstr = sqlstr & " order by pointtypecategory, pointtypelistorder"
    
    FillCombo Me.cmbPointTypes, sqlstr, cn, "omschrijving", "id"
'    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
'    With Me.cmbPointTypes
'        Set .RowSource = rs
'        .BoundColumn = "id"
'        .ListField = "omschrijving"
'    End With
    fillPointsGrid
    UnifyForm Me
    centerForm Me
End Sub

Sub fillPointsGrid()
Dim sqlstr As String
Dim i As Integer, J As Integer
  Set rs = New ADODB.Recordset
  sqlstr = "Select a.pointTypeID as id, a.poolID as poolId, pointTypeDescription as Omschrijving,"
  sqlstr = sqlstr & "pointPointsAward as Punten,"
  sqlstr = sqlstr & "pointPointsMargin as Marge "
  sqlstr = sqlstr & "from tblPoolPoints a inner join tblPointTypes b "
  sqlstr = sqlstr & "on a.pointtypeid = b.pointtypeid"
  sqlstr = sqlstr & " where a.poolID = " & thisPool
  sqlstr = sqlstr & " order by b.pointtypecategory, b.pointtypelistorder"
  
  With rs
      .CursorLocation = adUseClient
      .Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  End With
  With Me.grdPoolpoints
    .Clear
    .cols = rs.Fields.Count
    .rows = rs.RecordCount + 1
    i = 0
    For J = 0 To rs.Fields.Count - 1
      .TextMatrix(i, J) = rs.Fields(J).Name
    Next
    Do While Not rs.EOF
      i = i + 1
      For J = 0 To rs.Fields.Count - 1
        .TextMatrix(i, J) = rs.Fields(J).value
      Next
      rs.MoveNext
    Loop
    .colWidth(0) = 0
    .colWidth(1) = 0
    .colWidth(2) = 4000
    .ColAlignment(2) = flexAlignLeftCenter
    .colWidth(3) = 800
    .ColAlignment(3) = flexAlignCenterCenter
    .colWidth(4) = 800
    .ColAlignment(4) = flexAlignCenterCenter
    .Height = (.rows + 1) * .RowHeight(0)
    Me.btnClose.Top = .Top + .Height + 20
    Me.btnCopyDefaultPoints.Top = Me.btnClose.Top
    Me.Height = Me.btnClose.Top + Me.btnClose.Height + 600
  End With
  rs.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    Set rs = Nothing
    If (cn.State And adStateOpen) = adStateOpen Then cn.Close
    Set cn = Nothing
End Sub

Private Sub grdPoolpoints_RowColChange()
Dim id As Integer, i As Integer

  id = Me.grdPoolpoints.TextMatrix(Me.grdPoolpoints.row, 0)
  thisPointType = id
  
  For i = 0 To Me.cmbPointTypes.ListCount - 1
    If Me.cmbPointTypes.ItemData(i) = id Then
      Exit For
    End If
  Next
  If i < Me.cmbPointTypes.ListCount Then
    Me.cmbPointTypes = Me.cmbPointTypes.List(i)
    Me.upDnPnt = Me.grdPoolpoints.TextMatrix(Me.grdPoolpoints.row, 3)
    Me.uipDnMarge = val(Me.grdPoolpoints.TextMatrix(Me.grdPoolpoints.row, 4))
  End If
End Sub

Private Sub uipDnMarge_DownClick()
  saveValues
End Sub

Private Sub uipDnMarge_UpClick()
  saveValues
End Sub

Private Sub upDnPnt_DownClick()
  saveValues

End Sub

Private Sub upDnPnt_UpClick()
  saveValues

End Sub
