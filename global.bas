Attribute VB_Name = "global"
Option Explicit

'currentPool is read and stored in dbFunctions module

'save active entities
Public thisPool As Long
Public thisTournament As Long
Public thisAddress As Long
Public thisPlayer As Long
Public thisPoolForm As Long
Public thisMatch As Integer
Public thisTeam As Integer
'variable to preserve the current active country
Public currentCountry As Long  'used to pass information between forms
Public adminLogin As Boolean

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const penalty = 1
Public Const doelp = 2
Public Const eigdoelp = 3
Public Const geel = 4
Public Const rood = 5
Public Const penaltyMis = 6

Public Const dpAant = 16
Public Const gelijkSp = 17
Public Const geelKrt = 18
Public Const roodKrt = 19
Public Const pensAant = 20
Public Const eigdp = 32

Sub Main()
    
    'commandline arguments
    Dim i As Integer
    Dim strArgs() As String
    ' check if we started the app as admin
    strArgs = Split(Command$, " ")
    For i = 0 To UBound(strArgs)
        If strArgs(i) = "admin" Then
            adminLogin = True
            Exit For
        End If
    Next
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    write2Log "App started", True
    'check other instance of app
    If App.PrevInstance = True Then
        MsgBox "VBPool2.0 draait al...."
        Exit Sub
    End If
    'set and open the database
    If Dir(App.Path & "\" & dbName & ".mdb") = "" Then
        createDb
        write2Log "No vbpool2.mdb, dbcreated"
    End If
    'now that the database is created we can open the connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    
    'if there is a pools table with at least one record
    If recordsExist("tblPools", cn) Then
        ' get last poolID
        thisPool = val(GetSetting(App.EXEName, "global", "lastpool", 0))
    End If
    If thisPool <= 0 Then thisPool = 1
    If thisPool Then
        thisTournament = getThisPoolTournamentId(cn)
    End If
    cn.Close
    Set cn = Nothing
    'open main form
    frmMain.Show
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Sub

Sub UnifyForm(frm As Form, Optional center As Boolean)
'basic format for all forms
    Dim ctl As Control
    For Each ctl In frm.Controls
        On Error Resume Next 'if property does not exist
        ctl.Font.Name = "Tahoma"
        If ctl.Tag <> "cmbsmall" Then ctl.Font.Size = 10
        
        If InStr(ctl.Tag, "kop") Then 'small heading
            ctl.Font.Name = "Times New Roman"
            ctl.Font.Size = 14
            If InStr(ctl.Tag, "kop2") Then 'larger heading
                ctl.Font.Size = 20
            End If
            If InStr(ctl.Tag, "kop1") Then  'large heading
                ctl.Font.Size = 32
            End If
            If InStr(ctl.Tag, "subkop") Then 'larger heading
                ctl.Font.Size = 16
            End If
            
        End If
        
        If TypeOf ctl Is Label Then
            ctl.ForeColor = &H4000&  'dark green
            ctl.BackStyle = vbTransparent
            If ctl.Tag = "small" Then
              ctl.AutoSize = True
              ctl.Font.Size = 8
            End If
            If ctl.Tag = "xsmall" Then
              ctl.AutoSize = True
              ctl.Font.Size = 6
            End If
            
        End If
        If ctl.Tag = "cmbsmall" Then
          ctl.Font.Size = 7
        End If
        If TypeOf ctl Is CheckBox Then
            ctl.BackColor = frm.BackColor
        End If
        If InStr(ctl.Tag, "copyright") Then  'used for ©copyright message
 '           ctl.ForeColor = vbBlue
            ctl.Font.Size = 11
            ctl.Font.Name = "Garamond"
        End If
    Next
End Sub

Sub centerForm(frm As Object)
   frm.Move (Screen.width - frm.width) / 2, (Screen.Height - frm.Height) / 2
End Sub

Function float(strNumber As String) As String
'convert formatted dutch float number to dot seperated decimal
    Dim number As String
    If InStr(strNumber, "%") Then
        strNumber = val(Left(strNumber, Len(strNumber) - 1)) / 100
    End If
    
    If Not IsNumeric(strNumber) Then
        Exit Function
    Else
        float = Replace(strNumber, ",", ".")
    End If
End Function

Public Sub setCombo(objCmb As ComboBox, val As Variant)
    'set the combo listitem based on val in the listindex
    Dim i As Integer
    With objCmb
      .ListIndex = -1
      For i = 0 To .ListCount - 1
        If nz(val, 0) > 0 Then
          If .ItemData(i) = val Or .List(i) = val Then
            .ListIndex = i
            Exit For
          End If
        End If
      Next
    End With
End Sub

Public Sub FillCombo(objComboBox As Object, _
                     strSQL As String, _
                     cn As ADODB.Connection, _
                     strFieldToShow As String, _
                     Optional strFieldForItemData As String)

'Fills a combobox with values from a database

'code from VBforums

    Dim oRS As ADODB.Recordset  'Load the data
    Set oRS = New ADODB.Recordset
    oRS.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If oRS.EOF Then
        'MsgBox "Geen records in recordset", vbCritical + vbOKOnly, "FillCombo"
        Exit Sub
    End If
    With objComboBox          'Fill the combo box
        .Clear
        If strFieldForItemData = "" Then
            Do While Not oRS.EOF      '(without ItemData)
                .AddItem oRS.Fields(strFieldToShow).value
                oRS.MoveNext
            Loop
        Else
            Do While Not oRS.EOF      '(with ItemData)
                .AddItem oRS.Fields(strFieldToShow).value
                .ItemData(.NewIndex) = nz(oRS.Fields(strFieldForItemData).value, 0)
                oRS.MoveNext
            Loop
            DoEvents
        End If
    End With
    
    oRS.Close                 'Tidy up
    Set oRS = Nothing

End Sub

Sub fillList(objListBox As ListBox, _
              strSQL As String, _
              cn As ADODB.Connection, _
              strFieldToShow As String, _
              Optional strFieldForItemData As String)

    Dim oRS As ADODB.Recordset  'Load the data
    Set oRS = New ADODB.Recordset
    oRS.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If oRS.EOF Then
        Exit Sub
    End If
    With objListBox          'Fill the list box
        .Clear
        If strFieldForItemData = "" Then
            Do While Not oRS.EOF      '(without ItemData)
                .AddItem oRS.Fields(strFieldToShow).value
                oRS.MoveNext
            Loop
        Else
            Do While Not oRS.EOF      '(with ItemData)
                .AddItem oRS.Fields(strFieldToShow).value
                .ItemData(.NewIndex) = oRS.Fields(strFieldForItemData).value
                oRS.MoveNext
            Loop
        End If
    End With
    
    oRS.Close                 'Tidy up
    Set oRS = Nothing


End Sub

Public Function DoLogin() As Boolean

'login system originally from Michael Ciurescu (CVMichael from vbforums.com)
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn()
        .Open
    End With
    

    Dim UserName As String, Password As String, ret As Boolean
    Dim LoginSuccessful As Boolean, rsData As ADODB.Recordset
    Dim MD5 As New clsMD5
    
    Randomize
    
    ' Get the user that last logged in from the registry
    UserName = getOrganisation(cn, "lastname")
        
    ' prompt user to enter username and password
    ret = frmAdminLogin.GetLogIn(UserName, Password)
    
    Do While ret
        Set rsData = cn.Execute("SELECT Passwd FROM tblOrganisation WHERE lastname = '" & Replace(UserName, "'", "''") & "'")
        
        ' if a record was found, it means the user exists
        If Not rsData.EOF Then
            ' check if the password is correct
            If UCase(MD5.DigestStrToHexStr(Password)) = UCase(rsData("Passwd").value) Then
                
                LoginSuccessful = True
                Exit Do
            End If
        End If
        
        If Not LoginSuccessful Then
            ret = False
            
            If MsgBox("Wachtwoord onjuist, nog eens proberen?", vbQuestion + vbYesNo, "Login mislukt") = vbYes Then
                ' to prevent brute force password cracking from the application
                Sleep 200 + 300 * Rnd
                
                ' if login was not successfull, prompt again until Cancel is clicked
                ret = frmAdminLogin.GetLogIn(UserName, Password)
            End If
        End If
    Loop
    If Not LoginSuccessful Then
        write2Log "Login failed", True
    Else
        write2Log "Login successfull", True
    End If
    DoLogin = LoginSuccessful
    
    cn.Close
    Set cn = Nothing
End Function

'add the nz function
Public Function nz(strValue As Variant, Optional alternative As String = "") As Variant
    If Not IsNull(strValue) Then
        nz = strValue
    Else
        nz = alternative
    End If
End Function

Public Sub write2Log(txt, Optional timekolom As Boolean)
Dim iFileNr As Integer
Dim filenaam As String
Dim timestamp  As String

    iFileNr = FreeFile
    filenaam = App.Path & "\vbpool20.log"
    If timekolom Then
        timestamp = Format(Now(), "YYYY-MM-DD hh:nn:ss")
    Else
        timestamp = Space(20)
    End If
    
    Open filenaam For Append As #iFileNr
        Print #iFileNr, timestamp, txt
    Close #iFileNr

End Sub

Sub getTournamentTables()
Dim srcTable As String
Dim rsTables As ADODB.Recordset
Dim rsCols As ADODB.Recordset
Dim sqlstr As String
Dim tournTable As Boolean
Dim myConn As ADODB.Connection
    
    'get the tables from the mySql table collection
    Set rsTables = New ADODB.Recordset
    Set myConn = New ADODB.Connection
    With myConn
        .CursorLocation = adUseClient
        .ConnectionString = mySqlConn
        .Open
    End With
    sqlstr = "Select tournamentID from tblTournaments order by tournamentStartDate"
    rsTables.Open sqlstr, myConn, adOpenKeyset, adLockReadOnly
    If rsTables.EOF Then
        MsgBox "Geen verbinding gemaakt of geen gegevens gevonden!" & vbNewLine & "Kan niet verder gaan", vbOKOnly + vbCritical, "Database probleem"
        Exit Sub
    End If
    rsTables.MoveLast
    thisTournament = rsTables!tournamentID
    rsTables.Close
    'Use different sql in rsTablses now
    sqlstr = "SHOW TABLES in " & dbName
    rsTables.Open sqlstr, myConn, adOpenStatic, adLockReadOnly
    If rsTables.EOF Then
        MsgBox "Geen MySQL tabellen gevonden!", vbOKOnly, "FOUT"
        Exit Sub
    End If
    'get the id of the last tournament
    rsTables.MoveFirst
    Do While Not rsTables.EOF
        Set rsCols = New ADODB.Recordset
        srcTable = rsTables.Fields(0)
        If Left(srcTable, 6) <> "local_" Then
'            Me.lblTblName.Caption = "Tabel: " & srcTable
            'open connection to mySql
            rsCols.Open "SHOW COLUMNS from " & srcTable, myConn, adOpenForwardOnly, adLockReadOnly
            tournTable = False
            Do While Not rsCols.EOF 'check if there is a field for tournamentID, if so copy only data for this tournament
                If UCase(rsCols.Fields(0)) = "TOURNAMENTID" Then
                    tournTable = True
                    Exit Do
                End If
                rsCols.MoveNext
            Loop
            rsCols.Close
            copyTournamentData srcTable, tournTable, myConn
        End If
        rsTables.MoveNext
    Loop
    
    If Not rsTables Is Nothing Then
        If (rsTables.State And adStateOpen) = adStateOpen Then rsTables.Close
        Set rsTables = Nothing
    End If
    If Not rsCols Is Nothing Then
        If (rsCols.State And adStateOpen) = adStateOpen Then rsCols.Close
        Set rsCols = Nothing
    End If
    If Not myConn Is Nothing Then
        If (myConn.State And adStateOpen) = adStateOpen Then myConn.Close
        Set myConn = Nothing
    End If
'    Me.lblTblName.Caption = "Klaar! Alles ingelezen"
'    Me.lblRecord.Caption = ""
End Sub

Sub copyTournamentData(tblName As String, tournTable As Boolean, myConn As ADODB.Connection)
    'tournTable indicates if only specific tournament data will copied

Dim cmnd As ADODB.Command
Dim rsFrom As ADODB.Recordset
Dim rsTo As ADODB.Recordset
Dim sqlstr As String
Dim dellstr As String
Dim delStr As String
Dim valStr As String
Dim fld As field
Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn
        .Open
    End With
    
    Set cmnd = New ADODB.Command
    'open the fromTable
    With cmnd
        .ActiveConnection = myConn
        .CommandType = adCmdText
        sqlstr = "Select * from " & tblName
        delStr = "Delete from " & tblName
        If tournTable Then
            'only copy records for seleted tournament
            sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
            delStr = delStr & " WHERE tournamentID = " & thisTournament
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
        'Me.shpFill.Width = rsFrom.AbsolutePosition * (Me.shpBorder.Width / rsFrom.RecordCount)
        'Me.lblRecord.Caption = "Record " & rsFrom.AbsolutePosition & "/" & rsFrom.RecordCount
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
    
    If Not rsFrom Is Nothing Then
        If (rsFrom.State And adStateOpen) Then rsFrom.Close
        Set rsFrom = Nothing
    End If
    If Not cmnd Is Nothing Then
        Set cmnd = Nothing
    End If
    
    If Not rsTo Is Nothing Then
        If (rsTo.State And adStateOpen) = adStateOpen Then rsTo.Close
        Set rsTo = Nothing
    End If
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Sub


Sub fillCmbTournaments(cmb As ComboBox, _
                      lcl As Boolean)
'fill a combobox with tournaments from local database (lcl = true) or from server (lcl = false)
Dim cn As ADODB.Connection
Dim connStr As String
Dim sqlstr As String
  Set cn = New ADODB.Connection
  sqlstr = "Select tournamentId, "
  If lcl Then
    connStr = lclConn
    sqlstr = sqlstr & "tournamentYear & ' - '  & tournamentType "
  Else
    connStr = mySqlConn
    sqlstr = sqlstr & " concat(tournamentYear, ' - ', tournamentType) "
  End If
  sqlstr = sqlstr & " as tournament from tblTournaments order by tournamentYear"
  With cn
    .ConnectionString = connStr
    .Open
  End With
  FillCombo cmb, sqlstr, cn, "tournament", "tournamentID"
  cn.Close
  Set cn = Nothing
End Sub

Sub SelectAllText(tb As TextBox)
'select all text in a textbox. Use whet box gets focus
  tb.SelStart = 0
  tb.SelLength = Len(tb.Text)

End Sub


Sub showInfo(shw As Boolean, Optional kop As String, Optional t1 As String, Optional t2 As String, Optional t3 As String)
'open/close the information form
Dim i As Integer
    If Not shw Then
        Unload frmInfo
        Exit Sub
    End If
    With frmInfo
        .lblInfo(0).Caption = kop
        .lblInfo(1).Caption = t1
        .lblInfo(2).Caption = t2
        .lblInfo(3).Caption = t3
        For i = 0 To 3
            .lblInfo(i).Visible = True
        Next
        On Error Resume Next
        .Show
        DoEvents
        On Error GoTo 0
    End With
    
End Sub

Sub bindApp(ByRef appObj As Object, appNaam As String)
'globale routine voor het binden van een applicatie
  On Error Resume Next
  Set appObj = CreateObject(appNaam & ".application")
  appObj.Visible = False
  On Error GoTo 0
End Sub

Sub resetTestData(fromMatch As Integer)
'remove de matchResults vanaf fromMatch
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
  
  Set cn = New ADODB.Connection
  Set rs = New ADODB.Recordset
  
  With cn
    .ConnectionString = lclConn()
    .Open
  End With
  
  sqlstr = "Select * from tblTournamentSchedule"
  sqlstr = sqlstr & " where tournamentid = " & thisTournament
  sqlstr = sqlstr & " and matchorder >= " & fromMatch
  
  rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  Do While Not rs.EOF
    rs!matchPlayed = 0
    rs.MoveNext
  Loop
  sqlstr = "Delete from tblMatchresults"
  sqlstr = " WHERE tournamentID = " & thisTournament
  sqlstr = " AND matchorder >= " & fromMatch
  cn.Execute sqlstr
'en nu weer opnieuw doorrekeken
  '......
  
End Sub

Sub fillTestData(fromMatch As Integer, tillMatch As Integer)
'routine to fill test match results
Dim sqlstr As String
Dim htA As Integer
Dim htB As Integer
Dim ftA As Integer
Dim ftB As Integer
Dim xtA As Integer
Dim xtB As Integer
Dim toto As Integer
Dim wap As Boolean 'won after penalties
Dim teamID_A As Long
Dim teamID_B As Long
Dim pensA As Integer
Dim pensB As Integer
Dim winner As Long

Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

'Stop

With cn
  .ConnectionString = lclConn()
  .Open
End With

'eerst maar even matchresults leeg maken
sqlstr = "DELETE FROM tblMatchresults WHERE tournamentID = " & thisTournament
sqlstr = sqlstr & " AND matchorder >= " & fromMatch

cn.Execute sqlstr

'matchplayed op 0 zetten
sqlstr = "UPDATE tblTournamentSchedule SET matchPlayed = 0 "
sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
sqlstr = sqlstr & " AND matchorder >= " & fromMatch

cn.Execute sqlstr

sqlstr = "Select * from tblTournamentSchedule"
sqlstr = sqlstr & " where tournamentid = " & thisTournament
sqlstr = sqlstr & " AND matchorder >= " & fromMatch
sqlstr = sqlstr & " AND matchorder <= " & tillMatch
sqlstr = sqlstr & " ORDER BY matchOrder"

rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
Randomize
If rs.RecordCount Then
  Do While Not rs.EOF
    
    teamID_A = getTeamIdFromCode(rs!matchteamA, cn)
    teamID_B = getTeamIdFromCode(rs!matchteamB, cn)
    htA = Int(2 * Rnd())
    htB = Int(2 * Rnd())
    ftA = htA + Int(2 * Rnd())
    ftB = htB + Int(2 * Rnd())
    If teamID_A = 204 Then
      ftB = ftA + 5
    End If
    If teamID_B = 204 Then
      ftA = ftB + 3
    End If
    If teamID_A = 36 Then
      ftA = ftB + 1
    End If
    If teamID_B = 36 Then
      ftB = ftA + 1
    End If
    
    If ftA > ftB Then
      toto = 1
      winner = teamID_A
    End If
    If ftA < ftB Then
      toto = 2
      winner = teamID_B
    End If
    If ftA = ftB Then
      toto = 3
      winner = 0
    End If
    
    setMatchPlayed rs!matchOrder, True, cn
    
    If rs!matchOrder >= getFirstFinalMatchNumber(cn) And winner = 0 Then
      'ook extra tijd doelpunten
      xtA = Int(2 * Rnd())
      xtB = Int(2 * Rnd())
      If xtA = xtB Then
        wap = True
        If htB > htA Then
          pensA = Int(3 * Rnd()) + 2
          pensB = pensA - 1
          winner = teamID_A
        Else
          pensB = Int(3 * Rnd()) + 2
          pensA = pensB - 1
          winner = teamID_B
        End If
      Else
        If xtA > xtB Then
          winner = teamID_A
        Else
          winner = teamID_B
        End If
      End If
    End If
    
    sqlstr = "DELETE from tblMatchResults where matchOrder = " & rs!matchOrder
    sqlstr = sqlstr & " AND tournamentid = " & thisTournament
    cn.Execute sqlstr
    
    sqlstr = "INSERT INTO tblMatchResults VALUES(" & thisTournament
    sqlstr = sqlstr & ", " & rs!matchOrder
    sqlstr = sqlstr & ", " & htA
    sqlstr = sqlstr & ", " & htB
    sqlstr = sqlstr & ", " & ftA
    sqlstr = sqlstr & ", " & ftB
    sqlstr = sqlstr & ", " & winner
    sqlstr = sqlstr & ", " & toto
    sqlstr = sqlstr & ", " & xtA
    sqlstr = sqlstr & ", " & xtB
    sqlstr = sqlstr & ", " & IIf(wap, -1, 0)
    sqlstr = sqlstr & ", " & pensA
    sqlstr = sqlstr & ", " & pensB
    sqlstr = sqlstr & ", " & teamID_A
    sqlstr = sqlstr & ", " & teamID_B
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
    processMatch rs!matchOrder, cn
    
    rs.MoveNext
  Loop
End If

rs.Close
Set rs = Nothing
cn.Close
Set cn = Nothing
End Sub

Sub processMatch(matchOrder As Integer, cn As ADODB.Connection)
  Dim matchNr As Integer
  Dim grp As String
  Dim msg As String
  Dim winner As Long
  Dim sqlstr As String
    matchNr = getMatchNumber(matchOrder, cn)
    If getMatchInfo(matchOrder, "matchType", cn) = 1 Then
      'set the group standings
      calcGroupStandings cn
      grp = getMatchGroup(matchOrder, cn)
      If grpPlayedCount(grp, cn) = 6 Then
        msg = "Dit was de laatste wedstrijd van groep " & grp
        msg = msg & vbNewLine & "Controleer in het volgende scherm of de posities kloppen"
        msg = msg & vbNewLine & "en pas die posities eventueel aan om te bepalen wie er naar de volgende ronde gaat."
        msg = msg & vbNewLine & "De regels zijn nogal ingewikkeld bij gelijke stand"
'        msg = msg & vbNewLine & vbNewLine & "(Welke derde plaatsen doorgaan bepalen we na de laatste groepswedstrijd)"
        
        MsgBox msg, vbOKOnly + vbInformation, "Groep uitgespeeld"
        'show the form with the group standings to be able to adjust positions
        
        frmGroupStands.Show 1
        'set the teams on positions 1 and 2 through to the finals
        Set8Finals cn
      End If
      
    Else
      'set the winners through to the next round
      winner = nz(getMatchresult(matchOrder, 6, cn), 0)
      sqlstr = "UPDATE tblTournamentTeamCodes SET teamID = " & winner
      sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
      sqlstr = sqlstr & " AND teamCode = 'W" & Format(matchNr, "0") & "'"
      cn.Execute sqlstr
    End If

End Sub
