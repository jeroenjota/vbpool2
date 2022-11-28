Attribute VB_Name = "dbDefinition"
''''''''''
'Routines and functions for the definition and copying the database

Option Explicit

Public Const dbName = "vbpool2"

Function lclConn()
Dim fullPath As String
    fullPath = App.Path & "\" & dbName & ".mdb"
    lclConn = "PROVIDER='Microsoft.Jet.OLEDB.4.0';Data Source=" & fullPath & ";"
End Function

Function mySqlConn()
    ' opzetje om met mySql verbinding te maken.
    'toekomst??
    Dim server As String
    Dim driver As String
    Dim cnStr As String
    Dim passwd As String
    passwd = "!xjer56!"
    server = "websrv"
    'server = "jotaservices.duckdns.org"
    driver = "{MariaDB ODBC 3.1 Driver}"
    mySqlConn = "DRIVER=" & driver & ";TCPIP=1;SERVER=" & server & ";DATABASE=" & dbName & ";UID=jeroen;PWD=" & passwd & ";port=3306;"

End Function

Function tableExists(srcTable As String, cn As ADODB.Connection)
'check if table exists in local database
Dim rs As ADODB.Recordset
    Set rs = cn.OpenSchema(adSchemaColumns, Array(Empty, Empty, srcTable, Empty))
    tableExists = Not (rs.BOF And rs.EOF)
    rs.Close
    Set rs = Nothing
End Function

Function recordsExist(tblName As String, cn As ADODB.Connection)
'check if record exists in tblName
    Dim rs As ADODB.Recordset
    If tableExists(tblName, cn) Then
        Set rs = New ADODB.Recordset
        rs.Open "Select * from " & tblName, cn, adOpenKeyset, adLockReadOnly
        recordsExist = Not rs.EOF
    Else
        recordsExist = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Function cFieldType(fldType As String) As Integer
'convert mySQL fldType to ADODB type
    Dim returnType As Integer
    If Left(fldType, 7) = "varchar" Then
        returnType = adVarWChar  'default type
    Else
        Select Case LCase(fldType)
        Case "Date", "time", "datetime", "timestamp"
            returnType = adDate
        Case "int(11)"
            returnType = adInteger
        Case "double"
            returnType = adDouble
        Case "decimal(19,4)"
            returnType = adCurrency
        Case "tinyint(3)", "tinyint(3) unsigned"
            returnType = adUnsignedTinyInt
        Case "tinyint(1)"
            returnType = adBoolean
        Case Else
            Stop
        End Select
    End If
    cFieldType = returnType
End Function

'create the database
Sub createDb()
'this makes in fact a copy of the setup database
    Dim setupDb As String
    Dim newDb As String
    Dim msg As String
    Dim fileName As String
    
    ' MDB to be created. In app.path
    newDb = App.Path & "\" & dbName & ".mdb"
    ' Drop the existing database, if any.
    If Dir(newDb) > "" Then
        msg = "Er is al een database " & newDb & vbNewLine
        msg = msg & "Wil je een kopie hiervan bewaren?" & vbNewLine
        If MsgBox(msg, vbYesNo, "Nieuwe database aanmaken") = vbYes Then
            FileCopy newDb, newDb & ".bak"
        End If
        Kill newDb 'remove the old db
    End If
    setupDb = App.Path & "\vbpSetup.mdb"
    If Dir(setupDb) = "" Then
      MsgBox "Sorry, kan niet, de setup database is niet gevonden", vbOKOnly + vbExclamation, "Setup"
      Exit Sub
    End If
    FileCopy setupDb, newDb

End Sub

Sub fillDefaultValues()
'fill some tables with default values
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sqlstr As String
    Dim cmd As ADODB.Command
    Dim tournament As String
    Dim orgID As Long
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn
        .Open
    End With
    If thisTournament = 0 Then
        MsgBox "Geen toernooi geselcteerd ??, Heel vreemd!", vbOKOnly + vbCritical, "Database Fout in fillDefaultValues"
        Exit Sub
    End If
    tournament = getTournamentInfo("description", cn)
    'fill the first pool record
    'get the OrganisationID - should be only one organisation
    orgID = 1 'no longer need the organisation field, but just in case...
    'create a first record in tblPools
    sqlstr = "INSERT INTO tblPools (tournamentId, OrganisationId, poolName, poolFormsFrom, poolFormsTill, "
    sqlstr = sqlstr & "poolCost, prizeHighDayScore, prizeHighDayPosition, prizeLowDayposition, "
    sqlstr = sqlstr & "prizePercentage1, prizePercentage2, prizePercentage3, prizePercentage4, "
    sqlstr = sqlstr & "prizeLowFinalPosition ) VALUES ("
    sqlstr = sqlstr & thisTournament & ", " & orgID & ", '" & tournament & " pool" & "', " & CDbl(Date) & ", " & CDbl(getTournamentInfo("tournamentStartDate", cn) - 7) & ", "
    sqlstr = sqlstr & "10, 2.5, 1, 0.1, 50, 30, 20, 0, 10)"
    cn.Execute sqlstr
    
    'set the thisPool global variable
    sqlstr = "Select * from tblPools order by poolID"
    Set rs = New ADODB.Recordset
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    rs.MoveLast
    thisPool = rs!poolID
    
    'default data for points
    sqlstr = "INSERT into tblPoolPoints ( poolID, pointTypeId, pointPointsAward, pointPointsMargin )"
    sqlstr = sqlstr & " Select " & thisPool & ", pointtypeId , pointDefaultPoints, pointDefaultMargin from tblpointTypes"
    If Not getTournamentInfo("tournamentThirdPlace", cn) Then 'do not import 3rd and 4th place categories
        sqlstr = sqlstr & " WHERE pointTypeCategory <> 7"
        If getTournamentInfo("tournamentTeamCount", cn) <= 16 Then 'do not import 8th final categories
            sqlstr = sqlstr & " AND pointTypeCategory <> 2"
        End If
    End If
    cn.Execute sqlstr
    If Not rs Is Nothing Then
        If (rs.State And adStateOpen) = adStateOpen Then
            rs.Close
        End If
        Set rs = Nothing
    End If
    
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
End Sub

Sub makeTables(cn As ADODB.Connection)
    Dim sqlstr As String
    'OBSOLETE, just leave it here
    'new local tables just in case
    
    'address table
    sqlstr = "CREATE TABLE tblAddress ( "
    sqlstr = sqlstr & "addressID INTEGER NOT NULL, "
    sqlstr = sqlstr & "firstname VARCHAR(50), "
    sqlstr = sqlstr & "middlename VARCHAR(32), "
    sqlstr = sqlstr & "lastname VARCHAR(50), "
    sqlstr = sqlstr & "shortname VARCHAR(24), "
    sqlstr = sqlstr & "address VARCHAR(50), "
    sqlstr = sqlstr & "postalcode VARCHAR(10), "
    sqlstr = sqlstr & "city VARCHAR(50), "
    sqlstr = sqlstr & "telephone VARCHAR(20), "
    sqlstr = sqlstr & "email VarChar(255) "
    sqlstr = sqlstr & ") "
    cn.Execute sqlstr
    sqlstr = "CREATE INDEX PrimaryKey on tblAddress (addressID) WITH PRIMARY"
    cn.Execute sqlstr
    
    'competitorpoints
    sqlstr = "CREATE TABLE tblCompetitorPoints ("
    sqlstr = sqlstr & "addressID INTEGER NOT NULL,"
    sqlstr = sqlstr & "matchOrder INTEGER NOT NULL,"
    sqlstr = sqlstr & "pointsMatchTeams INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGroupStanding INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsFinals_8 INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsFinals_4 INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsFinals_2 INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsFinals_34 INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsFinal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsMatchResults INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTopscorers INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsOther INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsDayTotal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrandTotal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "poisitionDay INTEGER,"
    sqlstr = sqlstr & "positionTotal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "moneyDay DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "moneyDayPosition DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "moneyDayLast DECIMAL(19,4),"
    sqlstr = sqlstr & "moneyTotal DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "moneyDayTotal DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "pointsDay INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "positionMatches INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "poinstDayGoals INTEGER,"
    sqlstr = sqlstr & "pointsHalfTime INTEGER,"
    sqlstr = sqlstr & "pointsFulltime INTEGER,"
    sqlstr = sqlstr & "pointsToto INTEGER,"
    sqlstr = sqlstr & "pointsGrpA INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpB INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpC INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpD INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpE INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpF INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpG INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrpH INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8A INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8B INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8C INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8D INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8E INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8F INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8G INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals8H INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals4A INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals4B INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals4C INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals4D INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals2A INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals2B INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinals34 INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTeamsFinal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsTotalAfterFinal34 INTEGER DEFAULT 0"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
    'deelnemers
    sqlstr = "CREATE TABLE tblCompetitorPools ("
    sqlstr = sqlstr & "competitorPoolID INTEGER NOT NULL, "
    sqlstr = sqlstr & "poolid INTEGER NOT NULL,"
    sqlstr = sqlstr & "addressID INTEGER NOT NULL,"
    sqlstr = sqlstr & "nickName VARCHAR(50) NOT NULL,"
    sqlstr = sqlstr & "payed YESNO DEFAULT 0,"
    sqlstr = sqlstr & "predictionTeam1 INTEGER,"
    sqlstr = sqlstr & "predictionTeam2 INTEGER,"
    sqlstr = sqlstr & "predictionTeam3 INTEGER,"
    sqlstr = sqlstr & "predictionTeam4 INTEGER"
    sqlstr = sqlstr & ") "
    cn.Execute sqlstr
    sqlstr = "CREATE INDEX PrimaryKey on tblPoolCompetitors (competitorPoolId) WITH PRIMARY"
    cn.Execute sqlstr

    'pool points
    sqlstr = "CREATE TABLE tblPoolPoints ("
    sqlstr = sqlstr & "poolid INTEGER NOT NULL,"
    sqlstr = sqlstr & "pointTypeID INTEGER NOT NULL,"
    sqlstr = sqlstr & "pointPointsAward INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointPointsMargin byte DEFAULT 0 )"
    cn.Execute sqlstr

    'pools
    sqlstr = "CREATE TABLE tblPools ("
    sqlstr = sqlstr & "poolID INTEGER NOT NULL DEFAULT 0,"
    sqlstr = sqlstr & "tournamentID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "organisationID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "poolName varchar(50) DEFAULT NULL,"
    sqlstr = sqlstr & "poolStartAcceptForms datetime DEFAULT NULL,"
    sqlstr = sqlstr & "poolEndAcceptForms datetime DEFAULT NULL,"
    sqlstr = sqlstr & "poolCost decimal(19,4) DEFAULT 10.0000,"
    sqlstr = sqlstr & "prizeHighDayScore decimal(19,4) DEFAULT 0.0000,"
    sqlstr = sqlstr & "prizeHighDayPosition decimal(19,4) DEFAULT 0.0000,"
    sqlstr = sqlstr & "prizeLowDayPosition decimal(19,4) DEFAULT 0.0000,"
    sqlstr = sqlstr & "prizePercentage1 double DEFAULT 0,"
    sqlstr = sqlstr & "prizePercentage2 double DEFAULT 0,"
    sqlstr = sqlstr & "prizePercentage3 double DEFAULT 0,"
    sqlstr = sqlstr & "prizePercentage4 double DEFAULT 0,"
    sqlstr = sqlstr & "prizeLowFinalPosition decimal(19,4) DEFAULT 0.0000"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    sqlstr = "CREATE INDEX PrimaryKey on tblPools (poolID) WITH PRIMARY"
    cn.Execute sqlstr
    
    'predictions - Groups
    sqlstr = "CREATE TABLE tblPredictionGroupResults ("
    sqlstr = sqlstr & "competitorPoolId INTEGER NOT NULL,"
    sqlstr = sqlstr & "groupLetter varchar(1) DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGroupPosition1 varchar(255) DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGroupPosition2 varchar(255) DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGroupPosition3 varchar(255) DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGroupPosition4 varchar(255) DEFAULT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr


    sqlstr = "CREATE TABLE tblPredictionTopscorers ("
    sqlstr = sqlstr & "competitorPoolId INTEGER NOT NULL,"
    sqlstr = sqlstr & "predictionTopscorerPosittion INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "predictionTopscorePlayerID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "predictionTopscoreGoals INTEGER DEFAULT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr

    sqlstr = "CREATE TABLE tblPrediction_Finals ("
    sqlstr = sqlstr & "competitorPoolId INTEGER NOT NULL,"
    sqlstr = sqlstr & "matchOrder INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "teamNameA INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "teamNameB INTEGER DEFAULT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
    sqlstr = "CREATE TABLE tblPrediction_MatchResults ("
    sqlstr = sqlstr & "competitorPoolId INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "matchOrder INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGoalsHalftimeA BYTE DEFAULT NULL,"
    sqlstr = sqlstr & "predictionGoalsHalftimeB BYTE DEFAULT 0,"
    sqlstr = sqlstr & "predictionGoalsFulltimeA BYTE DEFAULT 0,"
    sqlstr = sqlstr & "predictionGoalsFulltimeB BYTE DEFAULT 0,"
    sqlstr = sqlstr & "predictionResultToto BYTE DEFAULT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
    sqlstr = "CREATE TABLE tblPrediction_Numbers ("
    sqlstr = sqlstr & "competitorPoolId INTEGER NOT NULL,"
    sqlstr = sqlstr & "predictionTypeID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "predictionNumber INTEGER DEFAULT NULL"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
    sqlstr = "CREATE TABLE tblUsers ("
    sqlstr = sqlstr & "userID INTEGER NOT NULL, "
    sqlstr = sqlstr & "username VARCHAR(50), "
    sqlstr = sqlstr & "Passwd VARCHAR(50) NOT NULL"
    sqlstr = sqlstr & ")"
    
    cn.Execute sqlstr
    
End Sub

Sub copyDefaultPoints()
'copy the default points table to current pool
Dim cnFrom As ADODB.Connection
Dim cnTo As ADODB.Connection
Dim cnStr As String
Dim fileName As String
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim msg As String
Dim fld As field
  
  Set cnTo = New ADODB.Connection
  Set cnFrom = New ADODB.Connection
  Set rs = New ADODB.Recordset
  Set rs2 = New ADODB.Recordset
  
  fileName = App.Path & "\vbpSetup.mdb"
  If Dir(fileName) = "" Then
    msg = "Bestand 'vbpSetup.mdb' niet gevonden in installatie map"
    msg = msg & vbNewLine & "Kan standaard puntentabel niet kopieëren"
    MsgBox msg, vbOKOnly + vbCritical, "Puntentabel"
    Exit Sub
  End If
  
  cnStr = "PROVIDER='Microsoft.Jet.OLEDB.4.0';Data Source=" & fileName & ";"
  With cnFrom
    .ConnectionString = cnStr
    .Open
  End With
  
  With cnTo
    .ConnectionString = lclConn
    .Open
    .CursorLocation = adUseClient
  End With
  'selete current records
  sqlstr = "Delete from tblPoolPoints WHERE poolid = " & thisPool
  cnTo.Execute sqlstr
  
  sqlstr = "Select * from tblPoolPoints"
  rs2.Open sqlstr, cnTo, adOpenKeyset, adLockOptimistic
  
  sqlstr = "Select * from tblPoolPoints where poolID = 0"
  rs.Open sqlstr, cnFrom, adOpenKeyset, adLockReadOnly
  'copy the records
  Do While Not rs.EOF
    rs2.AddNew
      For Each fld In rs.Fields
        rs2(fld.Name) = rs(fld.Name)
        If fld.Name = "poolid" Then rs2(fld.Name) = thisPool
      Next
    rs2.Update
    rs.MoveNext
  Loop
  'tidy up
  If (rs.State And adStateOpen) = adStateOpen Then rs.Close
  Set rs = Nothing
  If (rs2.State And adStateOpen) = adStateOpen Then rs2.Close
  Set rs2 = Nothing
  If (cnTo.State And adStateOpen) = adStateOpen Then cnTo.Close
  Set cnTo = Nothing
  If (cnFrom.State And adStateOpen) = adStateOpen Then cnFrom.Close
  Set cnFrom = Nothing
End Sub


Function sqlCompetitorMatches(competitorID As Long)
Dim sqlstr As String
  sqlstr = "SELECT ts.matchOrder as matchOrder, format(ts.matchDate, 'ddd dd') & ' om ' & format(ts.matchTime, 'HH:NN') as matchDate, "
  sqlstr = sqlstr & "format(ts.matchDate, 'd-m') as shortMatchDate, "
  sqlstr = sqlstr & "tc1.teamCode & ': ' & tn1.teamName & ' - ' & tc2.teamCode & ': ' & tn2.teamName as matchDesc, "
  sqlstr = sqlstr & "tn1.teamShortName & ' - ' & tn2.teamShortName as shortMatchDesc, "
  sqlstr = sqlstr & "tc1.teamCode & ' - ' & tc2.teamCode as shortCodeDesc, "
  sqlstr = sqlstr & "pr.htA as r1, pr.htB as r2, pr.ftA as e1, pr.ftB as e2, pr.tt as tt, matchnumber "
  sqlstr = sqlstr & "FROM ((((tblTournamentSchedule ts LEFT JOIN tblTournamentTeamCodes tc1 ON "
  sqlstr = sqlstr & "ts.tournamentID = tc1.tournamentID AND ts.matchTeamA = tc1.teamCode) "
  sqlstr = sqlstr & "LEFT JOIN tblTeamNames tn1 ON tc1.teamID = tn1.teamNameID) "
  sqlstr = sqlstr & "LEFT JOIN tblTournamentTeamCodes tc2 ON "
  sqlstr = sqlstr & "ts.tournamentID = tc2.tournamentID AND ts.matchTeamB = tc2.teamCode) "
  sqlstr = sqlstr & "LEFT JOIN tblTeamNames tn2 ON tc2.teamID = tn2.teamNameID) "
  sqlstr = sqlstr & "LEFT JOIN tblPrediction_MatchResults pr ON ts.matchOrder = pr.matchOrder"
  sqlstr = sqlstr & " WHERE ts.tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND pr.competitorpoolid = " & competitorID
  sqlstr = sqlstr & " ORDER BY ts.matchOrder"

  sqlCompetitorMatches = sqlstr
End Function
