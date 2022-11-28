VERSION 5.00
Begin VB.Form frmCopyData 
   Caption         =   "Get Data"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6045
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
   ScaleHeight     =   3165
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbTournament 
      Height          =   360
      Left            =   4080
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.CheckBox chkNewDb 
      Alignment       =   1  'Right Justify
      Caption         =   "Nieuwe database aanmaken"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Sluiten"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Start"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblTournamentInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   1155
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Van"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   8
      Top             =   1155
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Selecteer toernooi voor deze pool"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   3135
   End
   Begin VB.Shape shpFill 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   120
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape shpBorder 
      Height          =   495
      Left            =   120
      Top             =   2520
      Width           =   4000
   End
   Begin VB.Label lblRecord 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Tag             =   "kop"
      Top             =   2040
      Width           =   3885
   End
   Begin VB.Label lblTblName 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Tag             =   "kop"
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Haal de basisgegevens op van het internet"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmCopyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim myConn As ADODB.Connection

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
Dim msg As String

    msg = "Nieuwe Pool database aanmaken?"
    If Me.chkNewDb Then
    'OBSOLETE
        If MsgBox(msg, vbYesNo + vbQuestion, "Nieuwe database") = vbYes Then
        '(re) create an .mdb Access database file, for the local tables
            createDb
        End If
    End If
    If Me.cmbTournament.Text > "" Then
        msg = "Tournooigegevens inlezen van " & Me.cmbTournament.Text
    Else
        msg = "Tournooigegevens inlezen van alle toernooien in de database"
    End If
    If MsgBox(msg, vbYesNo + vbQuestion, "Toernooi gegevens") = vbYes Then
    'add tournament tables to the database from the server
        copyTournamentTables
        'copyData
    End If
End Sub

Private Sub cmbTournament_Click()
Dim periodText As String
    thisTournament = val(Me.cmbTournament.ItemData(Me.cmbTournament.ListIndex))
    periodText = Format(getTournamentInfo("tournamentStartDate", myConn), "ddd d MMM")
    periodText = periodText & " - " & Format(getTournamentInfo("tournamentEndDate", myConn), "ddd d MMM")
    Me.lblTournamentInfo.Caption = periodText
End Sub

Private Sub Form_Load()
Dim sqlstr As String
    Set myConn = New ADODB.Connection
    With myConn
      .ConnectionString = mySqlConn
      .Open
    End With
    Set cn = New ADODB.Connection
    With cn
      .ConnectionString = lclConn
      .Open
    End With
    'fill combobox
    fillCmbTournaments Me.cmbTournament, False
    
    UnifyForm Me
    centerForm Me
End Sub

Sub copyTournamentTables()
Dim srcTable As String
Dim newDb As String
Dim rsTables As ADODB.Recordset
Dim rsCols As ADODB.Recordset
Dim sqlstr As String
Dim tournTable As Boolean
    'get the tables from the mySql table collection
    Set rsTables = New ADODB.Recordset
    sqlstr = "SHOW TABLES in " & dbName
    rsTables.Open sqlstr, myConn, adOpenStatic, adLockReadOnly
    If rsTables.EOF Then
        MsgBox "Geen MySQL tabellen gevonden!", vbOKOnly, "FOUT"
        Exit Sub
    End If
    rsTables.MoveFirst
    Do While Not rsTables.EOF
        Set rsCols = New ADODB.Recordset
        srcTable = rsTables.Fields(0)
        If Left(srcTable, 6) <> "local_" Then
            'open connection to mySql
            Me.lblTblName.Caption = "Tabel: " & srcTable
            rsCols.Open "SHOW COLUMNS from " & srcTable, myConn, adOpenForwardOnly, adLockReadOnly
            tournTable = False
            Do While Not rsCols.EOF 'check if there is a field for tournamentID, if so copy only data for this tournament
                If UCase(rsCols.Fields(0)) = "TOURNAMENTID" Then
                    tournTable = True
                    Exit Do
                End If
                rsCols.MoveNext
            Loop
            copyData srcTable, tournTable
        End If
        rsTables.MoveNext
    Loop
    
    rsTables.Close
    Set rsTables = Nothing
    
    Me.lblTblName.Caption = "Klaar! Alles ingelezen"
    Me.lblRecord.Caption = ""
End Sub

Sub copyData(tblName As String, tournTable As Boolean)
    'tournData indicates if only specific tournament data will copied

Dim cmnd As ADODB.Command
Dim rsFrom As ADODB.Recordset
Dim rsTo As ADODB.Recordset
Dim sqlstr As String
Dim dellstr As String
Dim delStr As String
Dim valStr As String
Dim fld As field
    
    Set cmnd = New ADODB.Command
    'open the fromTable
    With cmnd
        .ActiveConnection = myConn
        .CommandType = adCmdText
        sqlstr = "Select * from " & tblName
        delStr = "Delete from " & tblName
        If tournTable And Me.cmbTournament.ListIndex > 0 Then
            'only copy records for seleted tournament
            sqlstr = sqlstr & " WHERE tournamentID = " & Me.cmbTournament.ItemData(Me.cmbTournament.ListIndex)
            delStr = delStr & " WHERE tournamentID = " & Me.cmbTournament.ItemData(Me.cmbTournament.ListIndex)
        End If
        .CommandText = sqlstr
        
        Set rsFrom = .Execute
    End With
    'delete records from local table
    cn.Execute delStr
    'add to the toTable
    Set rsTo = New ADODB.Recordset
    rsTo.Open "Select * from " & tblName, cn, adOpenKeyset, adLockOptimistic
    Do While Not rsFrom.EOF  'loop through records
        rsTo.AddNew
        'show info on form
        Me.shpFill.width = rsFrom.AbsolutePosition * (Me.shpBorder.width / rsFrom.RecordCount)
        Me.lblRecord.Caption = "Record " & rsFrom.AbsolutePosition & "/" & rsFrom.RecordCount
        DoEvents
        For Each fld In rsFrom.Fields  'loop through fields
            If Not IsNull(fld.value) Then
                rsTo(fld.Name) = fld.value
            Else
                If rsTo(fld.Name).Attributes = 70 Or rsTo(fld.Name).Attributes = 86 Then
                'if the field can not be NULL / just in case
                    If rsTo(fld.Name).Type = adVarWChar Then
                        rsTo(fld.Name) = "" 'set it to empty string
                    Else
                        rsTo(fld.Name) = 0 'set it to 0
                    End If
                End If
            End If
        Next
        rsTo.Update
        rsFrom.MoveNext 'next record
    Loop
    'tidy up
    rsFrom.Close
    Set rsFrom = Nothing
    
    Set cmnd = Nothing
    
    rsTo.Close
    Set rsTo = Nothing
    
End Sub

'Sub duplicateFields(toTable As ADOX.Table, fromTbl As String)
'    'copy tbl fields to Access database
'    Dim rs As ADODB.Recordset  'to store the columns
'    Dim col As ADOX.Column
'    Dim sqlstr As String
'    Dim ln As Integer
'    openMySql
'    'get all tables from the server
'    Set rs = New ADODB.Recordset
'    sqlstr = "SHOW COLUMNS in " & fromTbl & " in " & dbName
'    rs.Open sqlstr, myConn, adOpenStatic, adLockReadOnly
'    'copy the field defintion
'
'    With toTable
'        Do While Not rs.EOF
'            fldName = rs.Fields(0).value
'            Set col = New ADOX.Column
'            col.Name = fldName
'            col.Type = cFieldType(rs.Fields("Type"))
'            .Columns.Append col
'            If InStr(LCase(rs.Fields("Type")), "varchar") Then
'                ln = val(Mid(rs.Fields("Type"), 9, Len(rs.Fields("Type")) - 9))
'                .Columns(fldName).DefinedSize = ln
'            End If
'            If LCase(rs.Fields("Extra")) = "auto_increment" And rs.Fields("Type") = "int(11)" Then
'                .Columns(fldName).Properties("AutoIncrement").value = True
'                .Keys.Append "PrimaryKey", adKeyPrimary, fldName
'            Else
'                If rs.Fields("Type") = "tinyint(3)" Then
'                    .Columns(fldName).Attributes = adColNullable
'                End If
'            End If
'            '''' TEST
'            Dim prop As ADOX.Property
'            For Each prop In col.Properties
'                Debug.Print fromTbl, fldName, prop.Name, prop.value
'            Next
'            rs.MoveNext
'        Loop
'    End With
'
'    'release from memory
'    rs.Close
'    Set rs = Nothing
'    Set col = Nothing
'    myConn.Close
'    Set myConn = Nothing
'End Sub
'
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Clean-up procedure
    If Not cn Is Nothing Then
        'first, check if the state is open, if yes then close it
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        'set them to nothing
        Set cn = Nothing
    End If
    'same comment with cn
    If Not myConn Is Nothing Then
        If (myConn.State And adStateOpen) = adStateOpen Then
            myConn.Close
        End If
        Set myConn = Nothing
    End If
End Sub
