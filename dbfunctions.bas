Attribute VB_Name = "dbfunctions"
Option Explicit

Dim adoCmd As ADODB.Command

Function getMatchCount(matchType As Integer, cn As ADODB.Connection)
'return number of matches from type matchType of thisTournament
'if matchType = 0 the return all matches

  Dim sqlstr As String
  Dim rs As ADODB.Recordset
  Dim cmd As ADODB.Command
  Set cmd = New ADODB.Command
  Dim result As Integer
  result = 0
  sqlstr = "Select COUNT(matchType) as cnt from tblTournamentSchedule "
  sqlstr = sqlstr & " WHERE tournamentID = ?"
  If matchType > 0 Then
    sqlstr = sqlstr & " AND matchType like ?"
  End If
  With cmd
    .ActiveConnection = cn
    .CommandType = adCmdText
    .CommandText = sqlstr
    .Prepared = True
    .Parameters.Append .CreateParameter("id", adInteger, adParamInput)
    .Parameters("id") = thisTournament
    If matchType > 0 Then
      .Parameters.Append .CreateParameter("type", adInteger, adParamInput)
      .Parameters("type") = matchType
    End If
    Set rs = .Execute
  End With
  If Not rs.EOF Then
    result = rs!cnt
  End If
  getMatchCount = result
  rs.Close
  Set rs = Nothing
  Set cmd = Nothing
End Function

Function getAddressInfo(id As Long, fldName As String, cn As ADODB.Connection)
'return the value of fieldnmame in tblAddresses
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command
    Dim sqlstr As String
    Dim result As Variant
    Dim rs As ADODB.Recordset
    
    sqlstr = "Select * from tblAddresses Where addressID = ? "
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Prepared = True
        .Parameters.Append .CreateParameter("id", adInteger, adParamInput)
        .Parameters("id").value = id
        Set rs = .Execute
    End With
    If Not rs.EOF Then
        ' add fullname as extra - Access doesn't understand concat
        If fldName = "fullName" Then
            result = Trim(IIf(rs!firstname > " ", Trim(rs!firstname) & " ", "") & IIf(rs!middlename > " ", Trim(rs!middlename) & " ", "") & rs!lastname)
        Else
            If rs(fldName).Type = adBoolean Then
                result = CBool(rs(fldName)) * 1
            Else
                result = nz(rs(fldName), "")
            End If
        End If
    Else
        result = Null
    End If
    getAddressInfo = result
    rs.Close
    Set rs = Nothing
    Set adoCmd = Nothing


End Function

Function getCompetitorPoolInfo(id As Long, fldName As String, cn As ADODB.Connection)
'return the value of fieldnmame in tblAddresses
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command
    Dim sqlstr As String
    Dim result As Variant
    Dim rs As ADODB.Recordset
    
    sqlstr = "Select * from tblCompetitorPools Where competitorPoolID = ? "
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Prepared = True
        .Parameters.Append .CreateParameter("id", adInteger, adParamInput)
        .Parameters("id").value = id
        Set rs = .Execute
    End With
    If Not rs.EOF Then
        ' add fullname as extra - Access doesn't understand concat
      If rs(fldName).Type = adBoolean Then
          result = CBool(rs(fldName)) * 1
      Else
          result = rs(fldName)
      End If
    Else
        result = Null
    End If
    getCompetitorPoolInfo = result
    rs.Close
    Set rs = Nothing
    Set adoCmd = Nothing

End Function

Function getOrganisation(cn As ADODB.Connection, Optional field As String) As String
'get the name for the organisation of this pool / or just the content of field
Dim adoCmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim sqlstr As String
Dim result As String
    sqlstr = "Select * from tblOrganisation"
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        Set rs = .Execute
    End With

    If Not rs.EOF Then
        If field = "" Then
            result = Trim(rs!firstname)
            If rs!middlename > "" Then
                result = result & " " & Trim(rs!middlename)
            End If
            If rs!lastname > "" Then
                result = result & " " & Trim(rs!lastname)
            End If
'            result = result & vbNewLine & Trim(rs!address) & vbNewLine & Trim(rs!postalcode) & " " & Trim(rs!city)
        Else
            result = rs(field)
        End If
    End If
    getOrganisation = result
    rs.Close
    Set rs = Nothing
End Function

Function getPoolInfo(fldName As String, cn As ADODB.Connection)
'return the value of fieldnmame in tblPools
Dim adoCmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim sqlstr As String

    Set adoCmd = New ADODB.Command
    sqlstr = "Select " & fldName & " from tblPools where poolid = ?"
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Prepared = True
        .Parameters.Append .CreateParameter("id", adInteger, adParamInput)
        .Parameters("id").value = thisPool
        Set rs = .Execute
    End With
    If Not rs.EOF Then
        getPoolInfo = rs(fldName)
    Else
        getPoolInfo = Null
    End If
    
    rs.Close
    Set rs = Nothing
    Set adoCmd = Nothing

End Function

Function getTopscorerCount(cn As ADODB.Connection)
'get the number of topscorers on the poolForm
  Dim rs As ADODB.Recordset
  
  Dim result As Integer
  Dim sqlstr As String
  Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblPoolPoints WHERE poolid = " & thisPool
  sqlstr = sqlstr & " AND pointtypeID in (Select pointTypeID from tblPointTypes"
  sqlstr = sqlstr & " WHERE left(pointTypeDescription,9) = 'topscorer')"
  rs.Open sqlstr, cn
  If Not rs.EOF Then
    getTopscorerCount = rs.RecordCount
  Else
    getTopscorerCount = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getTournamentInfo(fldName As String, cn As ADODB.Connection)
'return the value of fieldnmame in tblTournaments
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command
    Dim sqlstr As String
    Dim result As Variant
    Dim rs As ADODB.Recordset
    
    sqlstr = "Select * from tblTournaments Where tournamentID = ? "
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Prepared = True
        .Parameters.Append .CreateParameter("id", adInteger, adParamInput)
        .Parameters("id").value = thisTournament
        Set rs = .Execute
    End With
    If Not rs.EOF Then
        ' add description as extra - Access doesn't understand concat
        If fldName = "description" Then
            result = rs!tournamenttype & " voetbal"
        Else
            If rs(fldName).Type = adBoolean Then
                result = CBool(rs(fldName)) * 1
            Else
                result = rs(fldName)
            End If
        End If
    Else
        result = Null
    End If
    getTournamentInfo = result
    rs.Close
    Set rs = Nothing
    Set adoCmd = Nothing
End Function

Function chkPoolHasCompetitors(pool As Long, cn As ADODB.Connection)
'are there competitors for this pool
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
        
        sqlstr = "Select  poolID from tblPoolCompetitors Where poolid = " & pool
        rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
        chkPoolHasCompetitors = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
End Function

Function chkTournamentHasPools(tournament As Long, cn As ADODB.Connection)
'are there pools for this tournament?
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
        sqlstr = "Select tournamentID from tblPools Where tournamentid = " & tournament
        rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
        chkTournamentHasPools = Not rs.EOF
    rs.Close
    Set rs = Nothing
End Function

Function getThisPoolTournamentId(cn As ADODB.Connection) As Long
'return the tournament for the current pool
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    getThisPoolTournamentId = 0
    Dim sqlstr As String
    sqlstr = "Select tournamentID from tblPools Where poolid = " & thisPool
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getThisPoolTournamentId = rs!tournamentID
    End If
    rs.Close
    Set rs = Nothing
End Function

Function chkTournamentStarted(cn As ADODB.Connection)
'check to see if tournament already started

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    chkTournamentStarted = False
    sqlstr = "Select * from tblTournaments Where tournamentid = " & getThisPoolTournamentId(cn)
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        chkTournamentStarted = CDbl(rs!tournamentStartDate) < CDate(Now())
    End If
    rs.Close
    Set rs = Nothing
End Function

Function supportsTransactions(cn As ADODB.Connection) As Boolean
'check if connection supports transactions
    On Error GoTo err_supportsTransactions:
        Dim lValue As Long
        lValue = cn.Properties("Transaction DDL").value
        supportsTransactions = True
    Exit Function
err_supportsTransactions:
    Select Case Err.number
    Case adErrItemNotFound:
        supportsTransactions = False
    Case Else
        MsgBox Err.Description
    End Select
End Function

Function tournamentHasSchedule(cn As ADODB.Connection) As Boolean
'check if there is already a schedule made
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "select * from tblTournamentSchedule where tournamentid = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    tournamentHasSchedule = Not rs.EOF
    rs.Close
    Set rs = Nothing
End Function

Function tournamentBaseSchedule() As Boolean
'check if there is already a base schedule made
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn
        .Open
    End With
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "select * from tblTournamentTeamCodes where tournamentid = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    tournamentBaseSchedule = Not rs.EOF
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Function

Sub generateSchedule()
'this routine builds the teams codes table for later use in Schedule. There we will add teamnames to these codes

Dim rsSchedule As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim sqlstr As String
Dim msg As String
Dim qry As ADODB.Command
Dim makeSchedule As Boolean
Dim letter As Integer
Dim matches As Integer
Dim groupSize  As Integer
Dim i As Integer, J As Integer
Dim teamCode As String

    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn
        .Open
    End With
    
    Set rsSchedule = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set qry = New ADODB.Command
    
    'we will exit
    'if there are is already a base schedule (in tblTournamentTeamCodses) for this tournament
     If tournamentBaseSchedule Then Exit Sub
    ''!!!!!!!!!!!!!!!!!!!
    'this routine gereates all the teamcodes necessary for this tournament. It will OVERWRITE the existing tblTournamentTeamCodes
    '!!!!!!!!!!!!!!!!!!!!
    sqlstr = "Select tournamentTeamCount as teams, tournamentGroupCount as groups from tblTournaments where tournamentId = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then Exit Sub
    groupSize = rs!teams / rs!groups
    matches = (groupSize - 1) * 2 * rs!groups 'total matches during groupfase
    'empty the codes table for this tournament
    cn.Execute "Delete from tblTournamentTeamCodes where tournamentid = " & thisTournament
    cn.Execute "Delete from tblTournamentSchedule where tournamentID = " & thisTournament
    sqlstr = "Select * from tblTournamentTeamCodes"
    rsSchedule.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    With rsSchedule
        For i = 1 To rs!groups
            For J = 1 To groupSize
                .AddNew
                !tournamentID = thisTournament
                teamCode = Chr(i + 64) & Format(J, "0")
                !teamCode = teamCode
                .Update
            Next
        Next
        If rs!groups > 4 Then
        '8th finales (normally I hope), should be 16 teams
            For i = 1 To rs!groups
                .AddNew
                !tournamentID = thisTournament
                !teamCode = "1" & Chr(i + 64)
                .Update
                .AddNew
                !tournamentID = thisTournament
                !teamCode = "2" & Chr(i + 64)
                .Update
            Next
            'if there are 6 groups then we need to add the best 3rd places to gt to 16
            If rs!groups = 6 Then  'add best 3rd places
                .AddNew
                !tournamentID = thisTournament
                !teamCode = "3ABC"
                .Update
                .AddNew
                !tournamentID = thisTournament
                !teamCode = "3ABCD"
                .Update
                .AddNew
                !tournamentID = thisTournament
                !teamCode = "3DEF"
                .Update
                .AddNew
                !tournamentID = thisTournament
                !teamCode = "3ADEF"
                .Update
            End If
        End If
        'other finals just the W(inner) of the matchnumber
        For i = matches + 1 To matches + 15
            .AddNew
            !tournamentID = thisTournament
            !teamCode = "W" & Format(i, "00")
            .Update
        Next
        If getTournamentInfo("tournamentThirdPlace", cn) Then 'add match for third place
            .AddNew
            !tournamentID = thisTournament
            !teamCode = "V" & Format(matches + 14, "00")
            .Update
        End If
    End With
    If (rs.State And adStateOpen) = adStateOpen Then rs.Close
    If (rsSchedule.State And adStateOpen) = adStateOpen Then rsSchedule.Close
    Set rs = Nothing
    Set rsSchedule = Nothing
    cn.Close
    Set cn = Nothing
End Sub

Sub addPlayers(cn As ADODB.Connection)
'add all players in the tblPeople table from a country in this tournament
    Dim sqlstr As String
    Dim rsTeams As ADODB.Recordset
    Dim rsPlayers As ADODB.Recordset
    
    Set rsTeams = New ADODB.Recordset
    Set rsPlayers = New ADODB.Recordset
    
    'remove all players in thistournament first
    sqlstr = "Delete from tblTeamPlayers where tournamentid = " & thisTournament
    cn.Execute sqlstr
    ' now build sqlstr to add players to teams
    sqlstr = "SELECT tournamentID, teamID, a.teamcodeID, teamName, b.teamCountryID, teamType "
    sqlstr = sqlstr & " FROM tblTeamNames b INNER JOIN tblTournamentTeamCodes a ON b.teamNameID = a.teamID"
    sqlstr = sqlstr & " where tournamentID = " & thisTournament
    rsTeams.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    If rsTeams.EOF Then Exit Sub 'if there are no teams what are we doing here
    
    rsTeams.MoveFirst
    Do While Not rsTeams.EOF
        'get all football players (tblPeople.function betweeen 2 and 5) from the same country as the team (NOT for clubteams)
        sqlstr = "Select * from tblPeople where function > 1 and function < 6 and countryCode = " & rsTeams!teamCountryId
        rsPlayers.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
        
        Do While Not rsPlayers.EOF
            sqlstr = "Insert into tblTeamPlayers (tournamentId, teamId, PlayerId) VALUES (" & thisTournament & "," & rsTeams!teamcodeID & ", " & rsPlayers!peopleID & ")"
            cn.Execute sqlstr
            rsPlayers.MoveNext
        Loop
        rsPlayers.Close
        rsTeams.MoveNext
    Loop
    If (rsTeams.State And adStateOpen) = adStateOpen Then rsTeams.Close
    Set rsTeams = Nothing
    If (rsPlayers.State And adStateOpen) = adStateOpen Then rsPlayers.Close
    Set rsPlayers = Nothing
End Sub

Function getPredictionInfo(id As Long, fldName As String, cn As ADODB.Connection)
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblPointTypes where pointTypeId = " & id
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getPredictionInfo = rs(fldName)
    Else
        getPredictionInfo = Null
    End If
    rs.Close
    Set rs = Nothing
End Function


Function getPlayerInfo(playerID As Long, fld As String, cn As ADODB.Connection)
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblPeople where peopleId = " & playerID
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getPlayerInfo = rs(fld)
    Else
        getPlayerInfo = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getPlayerTeam(playerID As Long, cn As ADODB.Connection)
  Dim sqlstr As String
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select teamID from tblTeamPlayers where playerId = " & playerID
  sqlstr = sqlstr & " AND tournamentID = " & thisTournament
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getPlayerTeam = rs!teamID
  Else
    getPlayerTeam = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getTeamInfo(teamID As Integer, fld As String, cn As ADODB.Connection)
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblTeamNames where teamNameId = " & teamID
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getTeamInfo = rs(fld)
    Else
        getTeamInfo = ""
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getTeamId(tournamentTeamCode As Long, cn As ADODB.Connection)
'get the basic id  of a tournament teamcode
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlstr = "Select * from tblTournamentTeamCodes where teamCodeId = " & tournamentTeamCode
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getTeamId = rs(rs!teamID)
    Else
        getTeamId = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getCountryID(Country As String, cn)
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlstr = "Select * from tblCountries where countryName = '" & Country & "'"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getCountryID = rs!countryID
    Else
        getCountryID = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getTeamIdFromCode(tournamentTeamCode As String, cn As ADODB.Connection)
'get the basic id  of a tournament teamcode
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlstr = "Select * from tblTournamentTeamCodes where teamCode = '" & tournamentTeamCode & "'"
    sqlstr = sqlstr & " AND tournamentID = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getTeamIdFromCode = rs!teamID
    Else
        getTeamIdFromCode = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getTeamIDFromCountryCode(countryCode As Integer, cn As ADODB.Connection)
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlstr = "Select * from tblTeamNames where teamCOuntryID = " & countryCode
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getTeamIDFromCountryCode = rs!teamNameID
    Else
        getTeamIDFromCountryCode = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getTeamIdFromName(teamName As String, cn As ADODB.Connection)
'get the basic id  of a tournament teamcode
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlstr = "Select * from tblTeamNames where teamName = '" & teamName & "'"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getTeamIdFromName = rs!teamNameID
    Else
        getTeamIdFromName = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getFinalsTeamCode(teamID As Long, cn As ADODB.Connection)
'get the teamId from a tounamentTeamCode
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblTournamentTeamCodes where tournamentId = " & thisTournament
    sqlstr = sqlstr & " AND teamId = " & teamID
    sqlstr = sqlstr & " AND left(teamCode,1) >= 'V' AND left(teamcode,1) <= 'W'"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getFinalsTeamCode = rs!teamCode
    Else
        getFinalsTeamCode = Null
    End If
    rs.Close
    Set rs = Nothing

End Function


Function getTournamentTeamCode(teamID As Long, cn As ADODB.Connection)
'get the teamId from a tounamentTeamCode
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblTournamentTeamCodes where tournamentId = " & thisTournament
    sqlstr = sqlstr & " AND teamId = " & teamID
    sqlstr = sqlstr & " AND left(teamCode,1) >= 'A' AND left(teamcode,1) <= 'H'"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getTournamentTeamCode = rs!teamCode
    Else
        getTournamentTeamCode = Null
    End If
    rs.Close
    Set rs = Nothing

End Function

Function playerInTournamentTeam(playerID As Long, teamID As Integer, cn As ADODB.Connection)
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblTeamPlayers where teamId = " & teamID
    sqlstr = sqlstr & " AND playerId = " & playerID
    sqlstr = sqlstr & " AND tournamentId = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    playerInTournamentTeam = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
End Function

Function playerExists(fName As String, mName As String, lName As String, nickName As String, cn As ADODB.Connection)
    'check double entries
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblPeople where (firstname = '" & fName
    sqlstr = sqlstr & "' AND middleName = '" & mName
    sqlstr = sqlstr & "' AND lastName = '" & lName
    sqlstr = sqlstr & "') OR nickName = '" & nickName & "'"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    playerExists = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
End Function


Function convertTournamentScheduleTable()
'change the reference in the tables from teamCodeID(Former primary Key from tblTournamentTeamCodes) to teamCode(string, A1, B2 etc)
'
'this makes the relation between schedule and teamcodes more intuitive, allbeit more complex (on two fields: tournamentID AND teamCode)
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = lclConn
        .Open
    End With
    

    Dim rsTn As ADODB.Recordset
    Dim rsCodes As ADODB.Recordset
    Set rsTn = New ADODB.Recordset
    Set rsCodes = New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "select * from  tblTournamentTeamCodes where teamCodeID > 0"
    rsCodes.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    Do While Not rsCodes.EOF
        sqlstr = "UPDATE tblGroupLayout SET teamID = " & rsCodes!teamID
        sqlstr = sqlstr & " WHERE teamId = " & rsCodes!teamcodeID
        If Not IsNull(rsCodes!teamID) Then cn.Execute sqlstr
        
        rsCodes.MoveNext
    Loop
    If Not rsTn Is Nothing Then
        rsTn.Close
        Set rsTn = Nothing
    End If
    If Not rsCodes Is Nothing Then
        rsCodes.Close
        Set rsCodes = Nothing
    End If
    
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
End Function

Function getMatchInfo(matchOrder As Integer, fldName As String, cn As ADODB.Connection)
'nr is the match ORDER number
'return the value of fieldnmame in tblTournamentSchedule
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command
    Dim sqlstr As String
    Dim result As Variant
    Dim rs As ADODB.Recordset
    
    
    sqlstr = "Select * from tblTournamentSchedule Where tournamentID = ? AND matchOrder = ? "
    
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Prepared = True
        .Parameters.Append .CreateParameter("id", adInteger, adParamInput)
        .Parameters("id").value = thisTournament
        .Parameters.Append .CreateParameter("match", adInteger, adParamInput)
        .Parameters("match").value = matchOrder
        Set rs = .Execute
    End With
    If Not rs.EOF Then
      If rs(fldName).Type = adBoolean Then
          result = CBool(rs(fldName)) * 1
      Else
          result = rs(fldName)
      End If
    Else
      result = Null
    End If
    getMatchInfo = result
    rs.Close
    Set rs = Nothing
    Set adoCmd = Nothing
End Function

Function getLastMatchPlayed(cn As ADODB.Connection)
'return the matchOrder number of the last match played
'!!!!!!!!!!!!!!  DO NOT Use MatchNUMBER becasue it can be different then the order of play
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select matchOrder from tblTournamentSchedule where tournamentId = " & thisTournament
  sqlstr = sqlstr & " AND matchPlayed = True Order by matchOrder"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveLast
    getLastMatchPlayed = rs!matchOrder
  Else
    getLastMatchPlayed = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getAllMatchesPlayedOnDay(thisDay As Date, cn As ADODB.Connection) As Boolean
'returns true if all the matches on date thisDay have a final result
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim matchesToPlay As Integer
  Set rs = New ADODB.Recordset
  sqlstr = "Select count(matchDate) as NumberOfMatches from tblTournamentSchedule where tournamentId = " & thisTournament
  sqlstr = sqlstr & " AND cdbl(matchDate) = " & CDbl(thisDay)
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  matchesToPlay = rs!NumberofMatches
  
  rs.Close
  sqlstr = sqlstr & " AND matchPlayed = true"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If matchesToPlay > 0 Then
    getAllMatchesPlayedOnDay = rs!NumberofMatches = matchesToPlay
  Else
    getAllMatchesPlayedOnDay = False
  End If

  rs.Close
  Set rs = Nothing

End Function

Function getCount(strSQL As String, cn As ADODB.Connection)
  'return number of records in fromTbl
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open strSQL, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveLast
    getCount = rs.RecordCount
  Else
    getCount = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getPointsFor(Description As String, cn As ADODB.Connection) As Integer
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim ret As Integer
  Set rs = New ADODB.Recordset
  sqlstr = "Select pointPointsAward, pointPointsMargin from tblPoolpoints "
  sqlstr = sqlstr & " WHERE poolid = " & thisPool
  sqlstr = sqlstr & " AND pointTypeID IN ("
  sqlstr = sqlstr & "Select pointTypeID from tblPointtypes WHERE "
  sqlstr = sqlstr & " pointTypeDescription = '" & Description & "'"
  sqlstr = sqlstr & " OR pointDescrShort = '" & Description & "')"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
'    ret(1) = rs!pointpointsAward
    ret = nz(rs!pointpointsAward, 0)
    getPointsFor = ret
  Else
    getPointsFor = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getPointsForID(pointTypeID As Long, cn As ADODB.Connection) As Integer
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim ret As Integer
  Set rs = New ADODB.Recordset
  sqlstr = "Select pointPointsAward, pointPointsMargin from tblPoolpoints "
  sqlstr = sqlstr & " WHERE poolid = " & thisPool
  sqlstr = sqlstr & " AND pointTypeID IN ("
  sqlstr = sqlstr & "Select pointTypeID from tblPointtypes WHERE "
  sqlstr = sqlstr & " pointTypeID  = " & pointTypeID & ")"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    ret = nz(rs!pointpointsAward, 0)
    
    getPointsForID = ret
  Else
    getPointsForID = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getMarginForID(pointTypeID As Long, cn As ADODB.Connection) As Integer
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim ret As Integer
  Set rs = New ADODB.Recordset
  sqlstr = "Select pointPointsMargin from tblPoolpoints "
  sqlstr = sqlstr & " WHERE poolid = " & thisPool
  sqlstr = sqlstr & " AND pointTypeID IN ("
  sqlstr = sqlstr & "Select pointTypeID from tblPointtypes WHERE "
  sqlstr = sqlstr & " pointTypeID  = " & pointTypeID & ")"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    ret = nz(rs!pointPointsMargin, 0)
    getMarginForID = ret
  Else
    getMarginForID = 0
  End If
  rs.Close
  Set rs = Nothing
End Function



Function getLastPoolID(cn As ADODB.Connection)
'get the ID of the last pool that was added
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblPools ORDER by poolid"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveLast
    getLastPoolID = rs!poolID
  Else
    getLastPoolID = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getFirstFinalMatchNumber(cn As ADODB.Connection)
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select matchOrder from tblTournamentSchedule WHERE matchType > 1 "
  sqlstr = sqlstr & " AND tournamentID = " & thisTournament
  sqlstr = sqlstr & " ORDER BY matchOrder"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveFirst
    getFirstFinalMatchNumber = rs!matchOrder
  Else
    getFirstFinalMatchNumber = 0
  End If
  rs.Close
  Set rs = Nothing
  
End Function

Sub fillDefaultPredictions(cn As ADODB.Connection)
'new competitor prepare empty tables for the current poolForm
'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'CAREFULL this resets the poolForm for thisCompetitor!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim matchCount As Integer
Dim i As Integer
Dim sqlstr As String
Dim rs As ADODB.Recordset
  
  matchCount = getMatchCount(0, cn)
  sqlstr = "Delete from tblPrediction_Finals  WHERE competitorPoolID = " & thisPoolForm
  cn.Execute sqlstr
  For i = getFirstFinalMatchNumber(cn) To matchCount
    sqlstr = "INSERT into tblPrediction_Finals (competitorPoolID, matchOrder) "
    sqlstr = sqlstr & " VALUES (" & thisPoolForm & ", " & i & ")"
    cn.Execute sqlstr
  Next
  'groupresults
  sqlstr = "Delete from tblPredictionGroupResults WHERE competitorPoolID = " & thisPoolForm
  cn.Execute sqlstr
  For i = 0 To getTournamentInfo("tournamentGroupCount", cn) - 1
    sqlstr = "INSERT into tblPredictionGroupResults (competitorPoolID, groupLetter) "
    sqlstr = sqlstr & " VALUES (" & thisPoolForm & ", '" & Chr(65 + i) & "')"
    cn.Execute sqlstr
  Next
'  'TopScorers
'  sqlstr = "Delete from tblPredictionTopScorers WHERE competitorPoolID = " & thisPoolForm
'  cn.Execute sqlstr
'  For i = 0 To getTopscorerCount(cn) - 1
'    sqlstr = "INSERT INTO tblPredictionTopScorers (competitorPoolID, topscorerPosition) "
'    sqlstr = sqlstr & " VALUES (" & thisPoolForm & ", " & i + 1 & ")"
'    cn.Execute sqlstr, cn
'  Next
  'Numbers
  sqlstr = "Delete from tblPrediction_Numbers WHERE competitorPoolID = " & thisPoolForm
  cn.Execute sqlstr
  sqlstr = "Select pointTypeID from tblPointTypes WHERE pointTypeCategory = 6 "
  sqlstr = sqlstr & " AND pointTypeID in (Select pointTypeID from tblPoolPoints where poolID = " & thisPool & ")"
  Set rs = New ADODB.Recordset
  rs.Open sqlstr, cn
  Do While Not rs.EOF
    sqlstr = "INSERT INTO tblPrediction_Numbers (competitorPoolID, predictionTypeID) "
    sqlstr = sqlstr & " VALUES (" & thisPoolForm & ", " & rs!pointTypeID & ")"
    cn.Execute sqlstr
    rs.MoveNext
  Loop
  rs.Close
  'match predictions
  sqlstr = "Delete from tblPrediction_Matchresults WHERE competitorPoolID = " & thisPoolForm
  cn.Execute sqlstr
  For i = 1 To matchCount
    sqlstr = "INSERT into tblPrediction_Matchresults (competitorPoolID, matchOrder, htA, htB"
    sqlstr = sqlstr & ", ftA,  ftB, tt) "
    sqlstr = sqlstr & " VALUES (" & thisPoolForm & ", " & i & ", 0, 0, 0, 0, 3)"
    cn.Execute sqlstr
  Next
  'topscorers
  For i = 1 To getTopscorerCount(cn)
    sqlstr = "INSERT INTO tblPredictionTopScorers (competitorPoolID, topscorerposition, topScorerPlayerID, topscorerGoals)"
    sqlstr = sqlstr & "VALUES (" & thisPoolForm & ", " & i & ", 0, 0)"
    cn.Execute sqlstr
  Next
  Set rs = Nothing
End Sub

Function getGroupTeamName(group As String, pos As Integer, cn As ADODB.Connection)
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "select teamName from tblTeamNames WHERE teamNameID IN ("
  sqlstr = sqlstr & "SELECT teamID from tblTournamentTeamCodes  "
  sqlstr = sqlstr & "WHERE tournamentID = " & thisTournament ' & ")"
  sqlstr = sqlstr & " AND teamCode = '" & group & Format(pos, "0") & "')"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getGroupTeamName = rs!teamName
  End If
  rs.Close
  Set rs = Nothing
End Function

Function matchPlayed(matchOrder As Integer, cn As ADODB.Connection)
Dim sqlstr As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    sqlstr = "select matchPlayed from tblTournamentSchedule WHERE tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND matchOrder = " & matchOrder
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
      matchPlayed = rs!matchPlayed
    Else
      matchPlayed = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getLastPlayDay(cn As ADODB.Connection)
  'return the last date that all matches where played
Dim sqlstr As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
  sqlstr = "Select matchdate , matchorder"
  sqlstr = sqlstr & " from tblTournamentSchedule "
  sqlstr = sqlstr & " where matchplayed = True"
  sqlstr = sqlstr & " AND tournamentID = " & thisTournament
  sqlstr = sqlstr & " ORDER BY matchorder"
  rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  If Not rs.EOF Then
    
  End If
End Function

Function getScheduleTeamName(scheduleCode As String, cn As ADODB.Connection, Optional shortname As Boolean)
' if known get the team name form the code in the schedule
'some codes doen't have a team yet if stages are to early

Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim fld As String
Dim zkID As Integer
Set rs = New ADODB.Recordset
'select the teamname format
  fld = "teamName"
  If shortname Then fld = "teamShortName"
'get the id from the schedule code
  sqlstr = "SELECT teamID from tblTournamentTeamCodes WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND teamCode = '" & scheduleCode & "'"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  'get the name from the table
  If Not rs.EOF Then
    zkID = nz(rs!teamID, 0)
    getScheduleTeamName = getTeamInfo(zkID, fld, cn)
  Else
    getScheduleTeamName = ""
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getScheduleTeamNames(matchOrdrNr As Integer, cn As ADODB.Connection, Optional shortname As Boolean) As String()
  'return an array of two teamnames from the matchNr
  Dim teams(1) As String
  Dim teamCode As String
  teamCode = getMatchInfo(matchOrdrNr, "matchTeamA", cn)
  teams(0) = getScheduleTeamName(teamCode, cn, shortname)
  If teams(0) = "" Then teams(0) = teamCode
  teamCode = getMatchInfo(matchOrdrNr, "matchTeamB", cn)
  teams(1) = getScheduleTeamName(teamCode, cn, shortname)
  If teams(1) = "" Then teams(1) = teamCode
  getScheduleTeamNames = teams
End Function

Function getMatchTeamCodes(matchOrdrNr As Integer, cn As ADODB.Connection) As String()
  Dim teamCodes(1) As String
  teamCodes(0) = getMatchInfo(matchOrdrNr, "matchTeamA", cn)
  teamCodes(1) = getMatchInfo(matchOrdrNr, "matchTeamB", cn)
  getMatchTeamCodes = teamCodes
End Function

Function getMatchTeamIDs(matchOrder As Integer, cn As ADODB.Connection) As Long()
  Dim teams(1) As Long
  teams(0) = getMatchresult(matchOrder, 13, cn) 'get teamA_ID
  teams(1) = getMatchresult(matchOrder, 14, cn) 'get teamB_ID
  getMatchTeamIDs = teams
End Function

Function getMatchresult(matchOrder As Integer, wat As Integer, cn As ADODB.Connection)
'return information from the matchresult table
  Dim sqlstr As String
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblMatchResults WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getMatchresult = rs.Fields(wat)
  Else
    getMatchresult = Null
  End If
  rs.Close
  Set rs = Nothing
End Function

Sub addMatchResultRecord(matchNumber As Integer, cn As ADODB.Connection)
'Dim sqlstr As String
'Dim matchExists As Boolean
'  Dim rs As ADODB.Recordset
'  Set rs = New ADODB.Recordset
'
'  sqlstr = "Select * from tblMatchResults where matchNumber = " & matchNumber
'  sqlstr = sqlstr & " AND tournamentID = " & thisTournament
'  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
'  matchExists = Not rs.EOF
'  rs.Close
'  If Not matchExists Then
'    sqlstr = "INSERT INTO tblMatchResults (tournamentID, matchNumber)"
'    sqlstr = sqlstr & " VALUES (" & thisTournament
'    sqlstr = sqlstr & ", " & matchNumber & ")"
'    cn.Execute sqlstr
'  End If
End Sub

Sub setMatchPlayed(matchOrder As Integer, played As Boolean, cn As ADODB.Connection)
  Dim sqlstr As String
  'Update the tblTourtnamentSchedule
  sqlstr = "UPDATE tblTournamentSchedule set matchPlayed = " & IIf(played, -1, 0)
  sqlstr = sqlstr & " WHERE tournamentId = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  cn.Execute sqlstr

End Sub

Function IsLastMatchOfDay(matchOrder As Integer, cn As ADODB.Connection)
'is deze wedstrijd de laatste van de dag
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim savdat As Date
Set rs = New ADODB.Recordset
    savdat = getMatchInfo(matchOrder, "MatchDate", cn)
    sqlstr = "Select * from tblTournamentSchedule where tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND format(matchDate,'d-m-yyyy') = '" & Format(savdat, "d-m-yyyy") & "'"
    sqlstr = sqlstr & " ORDER BY matchOrder"
    rs.Open sqlstr, cn, adOpenStatic, adLockOptimistic
    rs.MoveLast
    IsLastMatchOfDay = rs!matchOrder = matchOrder
    rs.Close
    Set rs = Nothing
End Function

Function getPredictionDayGoals(poolFormID As Long, matchDate As Date, cn As ADODB.Connection)
'find the total of predicted goals for this poolform on macthDate
  Dim rs As ADODB.Recordset
  Dim sqlstr As String
  Set rs = New ADODB.Recordset
    sqlstr = "SELECT SUM(ftA+ftB) as ttlGoals from tblPrediction_MatchResults"
    sqlstr = sqlstr & " WHERE competitorPoolID = " & poolFormID
    sqlstr = sqlstr & " AND matchNumber IN ("
    sqlstr = sqlstr & " SELECT matchNumber from tblTournamentSchedule "
    sqlstr = sqlstr & " WHERE clng(matchDate) = " & CLng(matchDate) & ")"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
      getPredictionDayGoals = rs!ttlGoals
    Else
      getPredictionDayGoals = 0
    End If
  rs.Close
  Set rs = Nothing
End Function

Function getGoalsPerDay(matchDate As Date, cn As ADODB.Connection)
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select sum(ftA) as A, sum(ftB) as B from tblMatchresults"
  sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder IN "
  sqlstr = sqlstr & " (Select matchOrder from tblTournamentSchedule WHERE "
  sqlstr = sqlstr & " tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND clng(matchDate) = " & CLng(matchDate)
  sqlstr = sqlstr & ")"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getGoalsPerDay = rs!A + rs!b
  Else
    getGoalsPerDay = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getPredictionGoalsPerDay(poolFormID As Long, matchDate As Date, cn)
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select sum(ftA) as A, sum(ftB) as B from tblPrediction_Matchresults"
  sqlstr = sqlstr & " WHERE competitorPoolId = " & poolFormID
  sqlstr = sqlstr & " AND matchOrder IN "
  sqlstr = sqlstr & " (Select matchOrder from tblTournamentSchedule WHERE "
  sqlstr = sqlstr & " tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND clng(matchDate) = " & CLng(matchDate)
  sqlstr = sqlstr & ")"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getPredictionGoalsPerDay = rs!A + rs!b
  Else
    getPredictionGoalsPerDay = 0
  End If
  rs.Close
  Set rs = Nothing

End Function

Function getGroup(teamID As Long, cn As ADODB.Connection)
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select groupLetter from tblGroupLayout "
  sqlstr = sqlstr & " WHERE tournamentid = " & thisTournament
  sqlstr = sqlstr & " AND teamID = " & teamID
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getGroup = rs!groupletter
  Else
    getGroup = "?"
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getGroupPlace(teamID As Long, cn As ADODB.Connection)
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select groupPlace from tblGroupLayout "
  sqlstr = sqlstr & " WHERE tournamentid = " & thisTournament
  sqlstr = sqlstr & " AND teamID = " & teamID
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getGroupPlace = rs!groupPlace
  Else
    getGroupPlace = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getMatchGroup(matchOrder As Integer, cn As ADODB.Connection)
'get the group letter for this match (if matchOrder is a groupstage match
  Dim grp As String
  grp = nz(getMatchInfo(matchOrder, "matchTeamA", cn), "")
  grp = Left(grp, 1)
  getMatchGroup = grp
End Function

Function grpPlayedAll(group As String, cn As ADODB.Connection)
  grpPlayedAll = grpPlayedCount(group, cn) = 6
End Function

Function grpPlayedCount(group As String, cn As ADODB.Connection)
'return matches played in this group
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select sum(mPl) as played from tblGroupLayout "
  sqlstr = sqlstr & " WHERE tournamentid = " & thisTournament
  sqlstr = sqlstr & " AND groupLetter = '" & group & "'"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    grpPlayedCount = rs!played / 2
  Else
    grpPlayedCount = 0
  End If
  rs.Close
  Set rs = Nothing
  
End Function

Function getLastGroupMatch(grp As String, cn As ADODB.Connection)
'RETURNS last matchOrder within the group
Dim sqlstr As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblTournamentSchedule "
  sqlstr = sqlstr & " where tournamentId = " & thisTournament
  sqlstr = sqlstr & " AND matchType = 1"
  sqlstr = sqlstr & " AND left(matchTeamA,1) = '" & grp & "'"
  sqlstr = sqlstr & " ORDER BY matchOrder"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  rs.MoveLast
  If Not rs.EOF Then
    getLastGroupMatch = rs!matchOrder
  Else
    getLastGroupMatch = 0
  End If
  rs.Close
  Set rs = Nothing
  
End Function

Function getpoolFormCount(cn As ADODB.Connection)
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  sqlstr = "Select count(competitorPoolID) as cnt from tblCompetitorPools WHERE poolid = " & thisPool
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getpoolFormCount = rs!cnt
  Else
    getpoolFormCount = 0
  End If
  rs.Close
  Set rs = Nothing

End Function

Function getMatchDescrCombo(matchOrdrNr As Integer, cn As ADODB.Connection)
  Dim matchDescr As String
  Dim matchDate As String
  Dim teams() As String
  teams = getScheduleTeamNames(matchOrdrNr, cn, True)
  matchDescr = Format(matchOrdrNr, "0\:")
  matchDescr = matchDescr & teams(0) & "-" & teams(1)   'remove spaces as well
  matchDate = Format(getMatchInfo(matchOrdrNr, "matchDate", cn), "D MMM")
  matchDate = matchDate & "," & Format(getMatchInfo(matchOrdrNr, "matchTime", cn), "HH\u")
  matchDescr = matchDescr & ": " & matchDate
  getMatchDescrCombo = matchDescr
End Function

Function getMatchDescription(matchOrder As Integer, cn As ADODB.Connection, Optional withDate As Boolean, Optional withTime As Boolean, Optional shortDescr As Boolean, Optional noMatchNr As Boolean)
  Dim matchDescr As String
  Dim matchDate As String
  Dim teams() As String
  If Not noMatchNr Then
    matchDescr = Format(getMatchNumber(matchOrder, cn), "0") & ": "
  End If
  teams = getScheduleTeamNames(matchOrder, cn, shortDescr)
  If shortDescr Then
    matchDescr = matchDescr & teams(0) & "-" & teams(1)   'remove spaces as well
  Else
    matchDescr = matchDescr & teams(0) & " - " & teams(1)
  End If
  If withDate Then
    matchDate = Format(getMatchInfo(matchOrder, "matchDate", cn), "D MMM")
    If withTime Then
      matchDate = matchDate & " om " & Format(getMatchInfo(matchOrder, "matchTime", cn), "HH:NN")
    End If
    matchDescr = matchDescr & " op " & matchDate
  End If
  getMatchDescription = matchDescr
End Function

Function getPrevMatchNr(matchOrder As Integer, cn As ADODB.Connection)
'get the previous match in date/time order before matchnr
''''''''''''''''''''''''''''''''
''because matchOrder is now key and always in correct order
'we can just substract one number for previous match
'''''''''''''''''''
If matchOrder > 0 Then
  getPrevMatchNr = matchOrder - 1
Else
  getPrevMatchNr = 0
End If

'Dim rs As New ADODB.Recordset
'Dim sqlstr As String
'Dim savdat As Date
'  If matchOrder = 1 Then
'    getPrevMatchNr = 0
'  Else
'    sqlstr = "Select matchOrder from tblTournamentSchedule"
'    sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
'    sqlstr = sqlstr & " order by matchOrder"
'    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    rs.Find "matchOrder = " & matchOrder
'    If Not rs.EOF Then
'      rs.MovePrevious
'      getPrevMatchNr = rs!matchOrder
'    Else
'      getPrevMatchNr = 0
'    End If
'    rs.Close
'    Set rs = Nothing
'  End If
'
End Function

Function getMatchresultStr(matchOrder As Integer, inclHalfTime As Boolean, cn As ADODB.Connection) As String
'return a formatted string with the match result
  Dim retStr As String
  If matchPlayed(matchOrder, cn) Then
    retStr = Format(getMatchresult(matchOrder, 4, cn), "0")
    retStr = retStr & "-" & Format(getMatchresult(matchOrder, 5, cn), "0")
    If inclHalfTime Then
      retStr = retStr & "(" & Format(getMatchresult(matchOrder, 2, cn), "0")
      retStr = retStr & "-" & Format(getMatchresult(matchOrder, 3, cn), "0") & ")"
    End If
  Else
    retStr = ""
  End If
  getMatchresultStr = retStr
End Function

Function getMatchResultPartStr(matchOrder As Integer, part As Integer, cn As ADODB.Connection)
'return a formatted string with the match result
  Dim ht As String
  Dim ft As String
  Dim xt As String
  Dim retStr As String
  ht = Format(getMatchresult(matchOrder, 2, cn), "0")
  ht = ht & "-" & Format(getMatchresult(matchOrder, 3, cn), "0")
  ft = Format(getMatchresult(matchOrder, 4, cn), "0")
  ft = ft & "-" & Format(getMatchresult(matchOrder, 5, cn), "0")
  xt = Format(getMatchresult(matchOrder, 8, cn) + getMatchresult(matchOrder, 4, cn), "0")
  xt = xt & "-" & Format(getMatchresult(matchOrder, 9, cn) + getMatchresult(matchOrder, 5, cn), "0")
  Select Case part
  Case 0
    retStr = ht
  Case 1
    retStr = ft
  Case 2
    retStr = ft & "(" & ht & ")"
    If matchOrder >= getFirstFinalMatchNumber(cn) Then
      If getMatchresult(matchOrder, 4, cn) = getMatchresult(matchOrder, 5, cn) Then
        retStr = ft & "(" & ht & ") nv:" & xt
        If getMatchresult(matchOrder, 8, cn) = getMatchresult(matchOrder, 9, cn) Then
          retStr = retStr & "," & getTeamInfo(getMatchresult(matchOrder, 6, cn), "teamShortname", cn) & " wns"
        End If
      End If
    End If
  End Select
  getMatchResultPartStr = retStr

End Function

Function getTotalPointsPrevDay(poolFormID As Long, matchOrder As Integer, cn As ADODB.Connection)
Dim rs As New ADODB.Recordset
Dim sqlstr As String
Dim savdat As Date
  savdat = getMatchInfo(matchOrder, "matchDate", cn)
  sqlstr = "Select SUM(pointsDay) as ttl from tblCompetitorPoints"
  sqlstr = sqlstr & " WHERE competitorPoolID = " & poolFormID
  sqlstr = sqlstr & " AND matchOrder in (Select matchOrder from tblTournamentSchedule"
  sqlstr = sqlstr & " WHERE cdbl(matchDate) < " & CDbl(savdat)
  sqlstr = sqlstr & " AND tournamentID = " & thisTournament & ")"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getTotalPointsPrevDay = nz(rs!ttl, 0)
  Else
    getTotalPointsPrevDay = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getMatchPrevDay(matchOrder As Integer, cn As ADODB.Connection)
Dim rs As New ADODB.Recordset
Dim sqlstr As String
Dim savdat As Date
  If matchOrder = 1 Then
    getMatchPrevDay = 0
  Else
    savdat = getMatchInfo(matchOrder, "matchDate", cn)
    sqlstr = "Select * from tblTournamentSchedule"
    sqlstr = sqlstr & " where cdbl(matchdate) <" & CDbl(savdat)
    sqlstr = sqlstr & " AND tournamentID = " & thisTournament
    sqlstr = sqlstr & " order by matchOrder"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
      rs.MoveLast
      getMatchPrevDay = rs!matchOrder
    Else
      getMatchPrevDay = 0
    End If
    rs.Close
    Set rs = Nothing
  End If
End Function

Function getMoneyTotal(poolForm As Long, matchOrder As Integer, cn As ADODB.Connection)
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  sqlstr = "Select * from tblCompetitorPoints"
  sqlstr = sqlstr & " WHERE competitorPoolID = " & poolForm
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getMoneyTotal = rs!moneyTotal
  Else
    getMoneyTotal = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getFin8Points(poolFormID As Long, tillMatch As Integer, grp As String, cn As ADODB.Connection)
'sum the points for the 8th final teams group by group
'Becauuse of the 3rd places that are calculateted after last group match we have to do this seperately
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  Dim pnts As Integer
  
  sqlstr = "Select SUM(pointsTeamsFinals8" & grp & ") AS ttl from tblCompetitorPoints"
  sqlstr = sqlstr & " WHERE competitorPoolID = " & poolFormID
  sqlstr = sqlstr & " AND matchorder <= " & tillMatch
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getFin8Points = rs!ttl
  Else
    getFin8Points = 0
  End If
  rs.Close
  Set rs = Nothing

End Function

Function getPoolFormPoints(poolFormID As Long, matchOrder As Integer, ptsType As Integer, cn As ADODB.Connection, Optional grp As String)
'0: competitorPoolID; 1: matchNumber ; 2: ptsDayGoals ; 3: ptsHt ; 4: ptsFt ; 5: ptsToto; 6: ptsMatch;
'7: pointsGroupStanding; 8: pointsGrpA; 9: pointsGrpB; 10: pointsGrpC; 11: pointsGrpD;
'12: pointsGrpE; 13: pointsGrpF; 14: pointsGrpG; 15: pointsGrpH;
'16: pointsFinals_8; 17: pointsTeamsFinals8A; 18: pointsTeamsFinals8B; 19: pointsTeamsFinals8C;
'20: pointsTeamsFinals8D; 21: pointsTeamsFinals8E; 22: pointsTeamsFinals8F; 23: pointsTeamsFinals8G; 24: pointsTeamsFinals8H;
'25: pointsFinals_4; 26: pointsTeamsFinals4A; 27: pointsTeamsFinals4B; 28: pointsTeamsFinals4C; 29: pointsTeamsFinals4D;
'30: pointsFinals_2; 31: pointsTeamsFinals2A; 32: pointsTeamsFinals2B;
'33: pointsFinals_34; 34: pointsTeamsFinals34; 35: pointsTotalAfterFinal34;
'36: pointsFinal; 37: pointsTeamsFinal; 38: pointsFinalStanding;
'39: pointsTopscorers; 40: pointsOther;
'41: pointsDay; 42: pointsDayTotal; 43: pointsGrandTotal;
'44: positionMatches; 45: positionDay; 46: positionTotal;
'47: moneyDay; 48: moneyDayPosition; 49: moneyDayLast; 50: moneyTotal; 51: moneyDayTotal;
  
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  Dim pnts As Integer
  
  sqlstr = "Select * from tblCompetitorPoints"
  sqlstr = sqlstr & " WHERE competitorPoolID = " & poolFormID
  If ptsType > 7 And ptsType < 41 And ptsType <> 16 And ptsType <> 25 And ptsType <> 30 Then
  ' then we add the points for the selected field
    sqlstr = sqlstr & " AND matchOrder <= " & matchOrder
  Else
  'we only return the selected field
    sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  End If
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  Do While Not rs.EOF
      pnts = pnts + rs.Fields(ptsType)
      rs.MoveNext
  Loop
  getPoolFormPoints = nz(pnts, 0)
  rs.Close
  Set rs = Nothing
End Function

Function getLongestNickName(cn As ADODB.Connection) As String
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim longTxt As String
  Dim sqlstr As String
  sqlstr = "Select * from tblCompetitorPools"
  sqlstr = sqlstr & " WHERE PoolID = " & thisPool
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  
  Do While Not rs.EOF
    If Len(rs!nickName) > Len(longTxt) Then longTxt = rs!nickName
    rs.MoveNext
  Loop
  getLongestNickName = longTxt
  rs.Close
  Set rs = Nothing

End Function

Function getHighesPts(matchOrder As Integer, cn As ADODB.Connection)
'find the highest overall points after a match
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  sqlstr = "Select MAX(pointsGrandTotal) as hi from tblCompetitorPoints"
  sqlstr = sqlstr & " WHERE matchOrder = " & matchOrder
  sqlstr = sqlstr & " AND competitorPoolID IN "
  sqlstr = sqlstr & " (SELECT competitorPoolID from tblCompetitorPools "
  sqlstr = sqlstr & " WHERE poolId = " & thisPool & ")"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getHighesPts = rs!hi
  Else
    getHighesPts = 0
  End If
  rs.Close
  Set rs = Nothing

End Function

Function getFinalmatchOrder(matchType As Integer, first As Boolean, cn As ADODB.Connection)
'return the last match order nmber for a certain matchtype
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  Dim retNr As Integer
  sqlstr = "Select matchOrder from tblTournamentSchedule WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchType = " & matchType
  sqlstr = sqlstr & " ORDER BY matchOrder"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    If Not first Then rs.MoveLast
    retNr = rs!matchOrder
  End If
  getFinalmatchOrder = retNr

  rs.Close
  Set rs = Nothing
End Function

Function getEventCount(matchNr As Integer, eventID As Long, cn As ADODB.Connection)
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
    sqlstr = "Select eventID from tblMatchEvents "
    sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND matchOrder <= " & matchNr
    sqlstr = sqlstr & " AND eventID = " & eventID
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
      rs.MoveLast
      getEventCount = rs.RecordCount
    Else
      getEventCount = 0
    End If
    rs.Close
    Set rs = Nothing

End Function

Function getStatistics(matchOrder As Integer, statType As Integer, cn As ADODB.Connection)
'return the stats based on pointsID
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  If statType = 17 Then 'gelijke spelen
    getStatistics = getDrawCount(matchOrder, cn)
  Else
    sqlstr = "Select eventID from tblMatchEvents "
    sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND matchOrder <= " & matchOrder
    Select Case statType
    Case 16 'doelpunten
      sqlstr = sqlstr & " AND eventID <= 3"
    Case 18 'gele kaarten
      sqlstr = sqlstr & " AND eventID = 4"
    Case 19
      sqlstr = sqlstr & " AND eventID = 5"
    Case 20
      sqlstr = sqlstr & " AND (eventID = 1 OR eventID = 6)"
    End Select
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
      rs.MoveLast
      getStatistics = rs.RecordCount
    Else
      getStatistics = 0
    End If
    rs.Close
    Set rs = Nothing
  End If
End Function

Function getDrawCount(matchNr As Integer, cn As ADODB.Connection)
'return the number of draw matches
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  
  sqlstr = "Select matchOrder from tblMatchResults "
  sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder <= " & matchNr
  sqlstr = sqlstr & " AND toto = 3"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveLast
    getDrawCount = rs.RecordCount
  Else
    getDrawCount = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getGrpPredictionCount(pos As Integer, teamID As Long, cn As ADODB.Connection)
'for favourites report
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  ' get the groupPlace for this team
  Dim place As Integer
  Dim grp As String
  place = getGroupPlace(teamID, cn)
  grp = getGroup(teamID, cn)
  sqlstr = "Select count(competitorPoolID) as cnt from tblPredictionGroupResults "
  sqlstr = sqlstr & " WHERE predictionGroupPosition" & Format(place, "0")
  sqlstr = sqlstr & " = " & pos
  sqlstr = sqlstr & " AND groupLetter = '" & grp & "'"
  sqlstr = sqlstr & " AND competitorPoolID IN ("
  sqlstr = sqlstr & " Select competitorPoolId from tblCompetitorPools where poolID = " & thisPool & ")"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveLast
    getGrpPredictionCount = rs!cnt
  Else
    getGrpPredictionCount = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getFavRowCount(matchNr As Integer, teamField As String, cn As ADODB.Connection)
'for favourites report
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  sqlstr = "Select matchOrder, " & teamField
  sqlstr = sqlstr & " FROM tblPrediction_Finals"
  sqlstr = sqlstr & " WHERE competitorPoolID IN ("
  sqlstr = sqlstr & " SELECT competitorPoolID from tblCompetitorPools WHERE poolid = " & thisPool & ")"
  sqlstr = sqlstr & " GROUP BY matchOrder, " & teamField
  sqlstr = sqlstr & " HAVING matchOrder = " & matchNr
  sqlstr = sqlstr & " AND " & teamField & " > 0"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getFavRowCount = rs.RecordCount
  Else
    getFavRowCount = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getTournamentDayCount(cn As ADODB.Connection)
'return number of days in this tournament
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  sqlstr = "Select matchdate "
  sqlstr = sqlstr & " FROM tblTournamentSchedule"
  sqlstr = sqlstr & " GROUP BY matchdate, tournamentID"
  sqlstr = sqlstr & " HAVING tournamentID = " & thisTournament
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getTournamentDayCount = rs.RecordCount
  Else
    getTournamentDayCount = 0
  End If
  rs.Close
  Set rs = Nothing

End Function

Function poolMoneyStr(cn As ADODB.Connection)
'returns a string with the calculated money for thispool
Dim days As Integer
Dim dayMoney As Double
Dim poolCost As Double
Dim moneyLeft As Double
Dim infostr As String
   
  days = getTournamentDayCount(cn)
  dayMoney = getPoolInfo("prizehighdayscore", cn)
  dayMoney = dayMoney + getPoolInfo("prizehighdayPosition", cn)
  dayMoney = dayMoney + getPoolInfo("prizeLowdayPosition", cn)
  poolCost = getPoolInfo("poolcost", cn)
  dayMoney = dayMoney * days
  moneyLeft = poolCost * getpoolFormCount(cn) - dayMoney - getPoolInfo("prizeLowFinalPosition", cn)
  infostr = infostr & "Inleg: " & getpoolFormCount(cn) * getPoolInfo("poolcost", cn)
  infostr = infostr & "; Aantal dagen: " & days
  infostr = infostr & "; Dagprijzen totaal: " & Format(dayMoney, "currency") & vbNewLine
  infostr = infostr & "; 1e pr: " & Format(moneyLeft * (getPoolInfo("prizePercentage1", cn) / 100), "€ 0.00")
  infostr = infostr & "; 2e pr: " & Format(moneyLeft * (getPoolInfo("prizePercentage2", cn) / 100), "€ 0.00")
  infostr = infostr & "; 3e pr: " & Format(moneyLeft * (getPoolInfo("prizePercentage3", cn) / 100), "€ 0.00")
  infostr = infostr & "; 4e pr: " & Format(moneyLeft * (getPoolInfo("prizePercentage4", cn) / 100), "€ 0.00")
  infostr = infostr & "; Laatste: " & Format(getPoolInfo("prizeLowFinalPosition", cn), "€ 0.00")

  poolMoneyStr = infostr
  
End Function

Function getHighestDayMatchNr(savdat As Date, cn As ADODB.Connection)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sqlstr As String
  sqlstr = "Select * from tblTournamentSchedule WHERE clng(matchDate) = " & CLng(savdat)
  sqlstr = sqlstr & " ORDER BY matchnumber"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveLast
    getHighestDayMatchNr = rs!matchNumber
  End If
rs.Close
Set rs = Nothing

End Function

Function getPrizeMoney(pos As Integer, cn As ADODB.Connection)
Dim prizes(4) As Double
Dim days As Integer
Dim dayMoney As Double
Dim poolCost As Double
Dim moneyLeft As Double
Dim poolFormCount As Integer

Dim i As Integer
  If pos <= 4 Then
    days = getTournamentDayCount(cn)
    dayMoney = getPoolInfo("prizehighdayscore", cn)
    dayMoney = dayMoney + getPoolInfo("prizehighdayPosition", cn)
    dayMoney = dayMoney + getPoolInfo("prizeLowdayPosition", cn)
    poolFormCount = getpoolFormCount(cn)
    poolCost = getPoolInfo("poolcost", cn)
  
    prizes(0) = getPoolInfo("prizeLowFinalPosition", cn)
    For i = 1 To 4
      prizes(i) = getPoolInfo("prizePercentage" & i, cn)
    Next
    moneyLeft = poolCost * poolFormCount - (dayMoney * days) - prizes(0)
    
    If pos > 0 Then
      getPrizeMoney = prizes(pos) * moneyLeft / 100
    Else
      getPrizeMoney = prizes(0)
    End If
  Else
    getPrizeMoney = 0
  End If
  
End Function

Function getLastPoolFormPosition(afterMatch As Integer, cn As ADODB.Connection)
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  sqlstr = "Select positionTotal from tblCompetitorPoints WHERE"
  sqlstr = sqlstr & " competitorPoolID IN "
  sqlstr = sqlstr & "(Select competitorPoolID from tblCompetitorPools where poolid = " & thisPool & ")"
  sqlstr = sqlstr & " AND matchOrder = " & afterMatch
  sqlstr = sqlstr & " ORDER BY positionTotal DESC"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getLastPoolFormPosition = rs!positionTotal
  Else
    getLastPoolFormPosition = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getPoolFormEndPoints(poolFormID As Long, pl As Integer, cn As ADODB.Connection)
'haal de punten voor de eindstand op
Dim rs As New ADODB.Recordset
Dim sqlstr As String
Dim pnt As Integer
Dim deelnpnt() As Integer
Dim finalPlaces As Integer
Dim winner As Long
Dim loser As Long
Dim finMatch As Integer
Dim thPlMatch As Integer
Dim i As Integer

  sqlstr = "Select * from tblCompetitorPools WHERE"
  sqlstr = sqlstr & " poolid = " & thisPool
  sqlstr = sqlstr & " AND competitorpoolid = " & poolFormID
  rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
  
  finMatch = getFinalmatchOrder(4, True, cn)
  If pl > 2 Then finMatch = finMatch - 1
  If getTournamentInfo("tournamentThirdPlace", cn) Then
    thPlMatch = getFinalmatchOrder(7, True, cn)
    finalPlaces = 4
  Else
    thPlMatch = 0
    finalPlaces = 2
  End If
  ReDim deelnpnt(finalPlaces)
  
  With rs
    i = 0
    Select Case pl
    Case 1
      pnt = getPointsForID(15, cn) 'tournooi winnaar
    Case 2
      pnt = getPointsForID(14, cn) '2e
    Case 3
      pnt = getPointsForID(13, cn) '3e
    Case 4
      pnt = getPointsForID(29, cn) '4e
    End Select
    If pnt > 0 Then
      If getMatchresult(finMatch, 6, cn) = getMatchresult(finMatch, 13, cn) Then
        winner = getMatchresult(finMatch, 13, cn)
        loser = getMatchresult(finMatch, 14, cn)
      Else
        winner = getMatchresult(finMatch, 14, cn)
        loser = getMatchresult(finMatch, 13, cn)
      End If
    Else
      getPoolFormEndPoints = 0
      Exit Function
    End If
    i = i + 1
    If pl <= 2 Then
      If !predictionteam1 = winner Then
        deelnpnt(0) = deelnpnt(0) + pnt
        deelnpnt(i) = pnt
      End If
      i = i + 1
      If !predictionteam2 = loser Then
        deelnpnt(0) = deelnpnt(0) + pnt
        deelnpnt(i) = pnt
      End If
    End If
'   3rd and 4th places for  later date
    If pl > 2 Then
      i = i + 1
      If !predictionteam3 = winner Then
        deelnpnt(0) = deelnpnt(0) + pnt
        deelnpnt(i) = pnt
      End If
      i = i + 1
      If !predictionteam4 = loser Then
        deelnpnt(0) = deelnpnt(0) + pnt
        deelnpnt(i) = pnt
      End If
    End If
  End With
  
  rs.Close
  Set rs = Nothing
  
  getPoolFormEndPoints = deelnpnt(pl)
End Function

Function getTotalGroupTtl(poolFormID As Long, matchNr As Integer, ptsType As Integer, fldCnt As Integer, cn As ADODB.Connection)
'returns the totalpoints for the subfields up till current match
'to put into the table, makes it easier later
'So input:
'    ptsType = is the starting fieldnumer in the tblCompetitorPointst table
'    fldCnt = is the amount of fields to be sumd
'for instance to get the value of pointsFinals4 : input would be ptstype 26, fldCnt 4
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim ttl As Integer
  Dim i As Integer
  Dim sqlstr As String
  sqlstr = "Select * from tblCompetitorPoints "
  sqlstr = sqlstr & " WHERE competitorPoolId = " & poolFormID
  sqlstr = sqlstr & " AND matchOrder >= " & matchNr  'matchnr and matchorder zijn hier nog verwisseld
  sqlstr = sqlstr & " AND matchOrder < " & getFirstFinalMatchNumber(cn)
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  Do While Not rs.EOF
    For i = 0 To fldCnt - 1
      ttl = ttl + rs.Fields(i + ptsType)
    Next
    rs.MoveNext
  Loop
  getTotalGroupTtl = ttl
  rs.Close
  Set rs = Nothing
End Function

Function getTotalPointsForFieldUntilMatch(poolFormID As Long, tillMatch As Integer, fieldName As String, cn As ADODB.Connection)
'returns the sum of points in fieldName till tillMatch number for poolformID
Dim sqlstr As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
sqlstr = "Select sum(" & fieldName & ") as ttl from tblCompetitorPoints "
sqlstr = sqlstr & " WHERE competitorPoolID = " & poolFormID
sqlstr = sqlstr & " AND matchOrder <= " & tillMatch
rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
  getTotalPointsForFieldUntilMatch = rs!ttl
Else
  getTotalPointsForFieldUntilMatch = 0
End If

End Function

Function getmatchOrderfromCode(teamCode As String, cn As ADODB.Connection)
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
    sqlstr = "Select matchOrder from tblTournamentSchedule "
    sqlstr = sqlstr & " WHERE matchTeamA = '" & teamCode & "' OR"
    sqlstr = sqlstr & " matchTeamB = '" & teamCode & "'"
    sqlstr = sqlstr & " and tournamentID = " & thisTournament
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
      getmatchOrderfromCode = rs!matchOrder
    Else
      getmatchOrderfromCode = 0
    End If
  rs.Close
  Set rs = Nothing
End Function


Function getMatchPts(poolFormID As Long, matchOrder As Integer, scorePart As Integer, cn As ADODB.Connection)
'matchresults
Dim poolPts As Integer
Dim sqlstr As String
Dim rsR As ADODB.Recordset
  Set rsR = New ADODB.Recordset
  sqlstr = "Select * from tblMatchResults "
  sqlstr = sqlstr & " WHERE matchOrder = " & matchOrder
  sqlstr = sqlstr & " AND tournamentid = " & thisTournament
  rsR.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
'participant poolForm
Dim rsP As ADODB.Recordset
  Set rsP = New ADODB.Recordset
  sqlstr = "Select * from tblPrediction_MatchResults "
  sqlstr = sqlstr & "where competitorPoolid = " & poolFormID
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  rsP.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  
  If rsR.EOF Or rsP.EOF Then
    getMatchPts = 0
    Exit Function
  End If
  
  Select Case scorePart
    Case 1 'halftime
      If rsR!htA = rsP!htA And rsR!htB = rsP!htB Then
        poolPts = getPointsFor("ruststand goed", cn)
      End If
    Case 2
      If rsR!ftA = rsP!ftA And rsR!ftB = rsP!ftB Then
        poolPts = getPointsFor("eindstand goed", cn)
      End If
    Case 3
      If rsR!toto = rsP!tt Then
        poolPts = getPointsFor("toto goed", cn)
      End If
  End Select
  getMatchPts = poolPts
  rsP.Close
  rsR.Close
  Set rsP = Nothing
  Set rsR = Nothing
End Function


Function getGroupPoints(poolFormID As Long, grp As String, cn As ADODB.Connection)
'get the points for the group standings
Dim rsForm As ADODB.Recordset
Dim rsGrp As ADODB.Recordset
Dim sqlstr As String
Dim pts As Integer
Dim ttlPts As Integer
Dim lastGroupPoints As Integer
  
Dim i As Integer
  ttlPts = 0
  pts = getPointsFor("groepstand", cn)
  Set rsForm = New ADODB.Recordset
  sqlstr = "Select * from tblPredictionGroupResults WHERE competitorPoolID = " & poolFormID
  sqlstr = sqlstr & " AND groupletter = '" & grp & "'"
  rsForm.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  Set rsGrp = New ADODB.Recordset
  sqlstr = "Select * from tblGroupLayout where tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND groupletter = '" & grp & "'"
  sqlstr = sqlstr & " ORDER by groupPlace"
  rsGrp.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  For i = 1 To 4
    rsGrp.Find ("groupPlace = " & i)
    If Not rsGrp.EOF Then
      If rsGrp!teamposition = rsForm("predictionGroupPosition" & Format(i, "0")) Then
        ttlPts = ttlPts + pts
      End If
    End If
    rsGrp.MoveFirst
  Next
  'return the total amount of points
  getGroupPoints = ttlPts
  
  rsForm.Close
  rsGrp.Close
  Set rsForm = Nothing
  Set rsGrp = Nothing
End Function

Function getPoolFormFinal8Points(poolFormID As Long, grp As String, cn As ADODB.Connection)
Dim sqlstr As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim rsMatch As ADODB.Recordset
Set rsMatch = New ADODB.Recordset
Dim rsForm As ADODB.Recordset
Set rsForm = New ADODB.Recordset
Dim ttlPts As Integer
Dim grpPts As Integer
Dim pts(1) As Integer
  
Dim matchNrFirst As Integer
Dim matchNrLast As Integer
Dim teamID As Long
Dim matchOrder As Integer  'remember the match number

Dim teamLeft As Boolean 'remember teamPosition
Dim onPosition As Boolean
  pts(0) = getPointsFor("8e fin team", cn)
  pts(1) = getPointsFor("8e fin pos", cn)
  matchNrFirst = getFinalmatchOrder(5, True, cn)
  matchNrLast = getFinalmatchOrder(5, False, cn)
  
  'stappen:
  '1. Zoek de teamID bij plaats1 van deze groep uit tblTournamentTeamCodes
  sqlstr = "Select * from tblTournamentTeamCodes WHERE  tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND (left(teamCode,1) = '1' or left(teamCode,1) = '2') "
  sqlstr = sqlstr & " AND right (teamCode,1) = '" & grp & "'"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveFirst
  '2. Zoek de wedstrijd van de gevonden teamCode uit de tblTournamentSchedule tabel
    teamID = rs!teamID
    sqlstr = "Select * from tblTournamentSchedule WHERE tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND (matchTeamA = '" & rs!teamCode & "' OR matchTeamB = '" & rs!teamCode & "')"
    rsMatch.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  '3. Onthoud het wedstrijd nr en de team positie (A of B)
    If Not rsMatch.EOF Then
      matchOrder = rsMatch!matchOrder
      teamLeft = rsMatch!matchteamA = rs!teamCode 'which team do we have home or away
    Else
      MsgBox "BIIG ERRROORRRR"
      End
    End If
  '4 Zoek in de tblPredictions_Finals naar de teamID in kolom teamnameA of B
    sqlstr = "Select * from tblPrediction_Finals WHERE competitorpoolID = " & poolFormID
    sqlstr = sqlstr & " AND matchOrder Between " & matchNrFirst & " AND " & matchNrLast
    sqlstr = sqlstr & " AND (teamNameA = " & teamID & " OR teamNameB = " & teamID & ")"
    rsForm.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  '5 Als gevonden: bepaal wedstrijd nummer en postitie A of B
    If Not rsForm.EOF Then
      grpPts = pts(0)  'in ieder geval de teamnaam punten
      Do While Not rsForm.EOF
      '6. als er meer dan 1 is gevonden zoekdan de beste plek (slechts een keer uitkeren, de hoogste)
        onPosition = matchOrder = rsForm!matchOrder
        If teamLeft Then
          onPosition = onPosition And rsForm!teamnameA = teamID
        Else
          onPosition = onPosition And rsForm!teamnameB = teamID
        End If
        If onPosition Then Exit Do
        rsForm.MoveNext
      Loop
      If onPosition Then grpPts = pts(1)
    End If
    rsMatch.Close
    rsForm.Close
    ttlPts = grpPts
    grpPts = 0
'doe hetzelfde voor de teamID van plaats 2
    rs.MoveLast 'should be the second record
    teamID = rs!teamID
    sqlstr = "Select * from tblTournamentSchedule WHERE tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND (matchTeamA = '" & rs!teamCode & "' OR matchTeamB = '" & rs!teamCode & "')"
    rsMatch.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  '3. Onthoud het wedstrijd nr en de team positie (A of B)
    If Not rsMatch.EOF Then
      matchOrder = rsMatch!matchOrder
      teamLeft = rsMatch!matchteamA = rs!teamCode 'which team do we have home or away
    Else
      MsgBox "BIIG ERRROORRRR"
      End
    End If
  '4 Zoek in de tblPredictions_Finals naar de teamID in kolom teamnameA of B
    sqlstr = "Select * from tblPrediction_Finals WHERE competitorpoolID = " & poolFormID
    sqlstr = sqlstr & " AND matchOrder Between " & matchNrFirst & " AND " & matchNrLast
    sqlstr = sqlstr & " AND (teamNameA = " & teamID & " OR teamNameB = " & teamID & ")"
    rsForm.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  '5 Als gevonden: bepaal wedstrijd nummer en postitie A of B
    If Not rsForm.EOF Then
      grpPts = pts(0)  'in ieder geval de teamnaam punten
      Do While Not rsForm.EOF
      '6. als er meer dan 1 is gevonden zoekdan de beste plek (slechts een keer uitkeren, de hoogste)
        onPosition = matchOrder = rsForm!matchOrder
        If teamLeft Then
          onPosition = onPosition And rsForm!teamnameA = teamID
        Else
          onPosition = onPosition And rsForm!teamnameB = teamID
        End If
        If onPosition Then Exit Do
        rsForm.MoveNext
      Loop
      If onPosition Then grpPts = pts(1)
    End If
  End If
  ttlPts = ttlPts + grpPts
  getPoolFormFinal8Points = ttlPts

End Function



'Function getPoolFormFinal4Points(poolFormID As Long, matchNr As Integer, cn As ADODB.Connection)
'punten voor de kwart finalisten
''!!!!!!!!!!!!!  OBSOLETE  !!!!!!!!!!!!!!!
''''''''''''''''''''''''''''''''''''''''''
'Dim rsSchedule As New ADODB.Recordset
'Dim rsPoolFormFinals As New ADODB.Recordset
'Dim sqlstr As String
'Dim grp As String
'Dim match As Integer
'Dim zkTeam As Long
'Dim zkTeam2 As Long
'Dim vldNaam As String
'Dim pnt As Integer
'Dim tmPos As String
'Dim matchNrFirst As Integer
'Dim matchNrLast As Integer
'  matchNrFirst = getFinalmatchOrder(2, True, cn)
'  matchNrLast = getFinalmatchOrder(2, False, cn)
'  sqlstr = "Select * from tblTournamentSchedule where tournamentID= " & thisTournament
'  sqlstr = sqlstr & " AND matchTeamA = 'W" & matchNr & "'"
'  rsSchedule.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'  vldNaam = "matchTeamA"
'  tmPos = "links"
'  If rsSchedule.EOF And rsSchedule.BOF Then 'moeten we het andere team hebben
'    rsSchedule.Close
'    sqlstr = "Select * from tblTournamentSchedule where tournamentID= " & thisTournament
'    sqlstr = sqlstr & " AND matchTeamB = 'W" & matchNr & "'"
'    rsSchedule.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    vldNaam = "matchTeamB"
'    tmPos = "rechts"
'    If rsSchedule.EOF And rsSchedule.BOF Then 'zou niet motte magge
'      MsgBox "Sorry, er zit een foutje in het programma, kan niet verder gaan", vbOKOnly + vbCritical, "Routine 'getHvFinPnt'"
'      End
'    End If
'  End If
'  match = rsSchedule!matchNumber
'  'grp = Left(rsSchedule!code1, 1)
'
'  pnt = 0
'  'If poolFormID = 7 Then Stop
'  sqlstr = "Select * from tblPrediction_Finals"
'  sqlstr = sqlstr & " Where competitorPoolID = " & poolFormID
'  sqlstr = sqlstr & " AND matchNumber >= " & matchNrFirst
'  sqlstr = sqlstr & " AND matchnumber <= " & matchNrLast
'  rsPoolFormFinals.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'
'  Do While Not rsPoolFormFinals.EOF
'    If tmPos = "links" Then
'      zkTeam = rsPoolFormFinals!teamnameA
'      zkTeam2 = rsPoolFormFinals!teamnameB
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) And match = rsPoolFormFinals!matchNumber Then
'        pnt = pnt + getPointsForID(7, cn) '"kwart finale positie"
'        Exit Do
'      End If
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) Or zkTeam2 = getTeamIdFromCode(rsSchedule(vldNaam), cn) Then
'        pnt = pnt + getPointsForID(6, cn) '"kwart finale team"
'        Exit Do
'      End If
'    Else
'      zkTeam = rsPoolFormFinals!teamnameB
'      zkTeam2 = rsPoolFormFinals!teamnameA
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) And match = rsPoolFormFinals!matchNumber Then
'        pnt = pnt + getPointsForID(7, cn) '"kwart finale positie"
'        Exit Do
'      End If
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) Or zkTeam2 = getTeamIdFromCode(rsSchedule(vldNaam), cn) Then
'        pnt = pnt + getPointsForID(6, cn) '"kwart finale team"
'        Exit Do
'      End If
'    End If
'    rsPoolFormFinals.MoveNext
'  Loop
'  rsPoolFormFinals.Close
'  Set rsPoolFormFinals = Nothing
'
'  rsSchedule.Close
'  Set rsSchedule = Nothing
'
'  getPoolFormFinal4Points = pnt
'
'End Function


'Function getPoolFormFinal2Points(poolFormID As Long, matchNr As Integer, cn As ADODB.Connection)
''punten voor de kwart finalisten
'''!!!!!!!!!!!!!  OBSOLETE  !!!!!!!!!!!!!!!
'''''''''''''''''''''''''''''''''''''''''''
'
'Dim rsSchedule As New ADODB.Recordset
'Dim rsPoolFormFinals As New ADODB.Recordset
'Dim sqlstr As String
'Dim grp As String
'Dim match As Integer
'Dim zkTeam As Long
'Dim zkTeam2 As Long
'Dim vldNaam As String
'Dim pnt As Integer
'Dim tmPos As String
'Dim matchNrFirst As Integer
'Dim matchNrLast As Integer
'  matchNrFirst = getFinalmatchOrder(3, True, cn)
'  matchNrLast = getFinalmatchOrder(3, False, cn)
'  sqlstr = "Select * from tblTournamentSchedule where tournamentID= " & thisTournament
'  sqlstr = sqlstr & " AND macthTeamA = 'W" & matchNr & "'"
'  rsSchedule.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'  vldNaam = "matchTeamA"
'  tmPos = "links"
'  If rsSchedule.EOF And rsSchedule.BOF Then 'moeten we het andere team hebben
'    rsSchedule.Close
'    sqlstr = "Select * from tblTournamentSchedule where tournamentID= " & thisTournament
'    sqlstr = sqlstr & " AND macthTeamB = 'W" & matchNr & "'"
'    rsSchedule.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    vldNaam = "matchTeamB"
'    tmPos = "rechts"
'    If rsSchedule.EOF And rsSchedule.BOF Then 'zou niet motte magge
'      MsgBox "Sorry, er zit een foutje in het programma, kan niet verder gaan", vbOKOnly + vbCritical, "Routine 'getHvFinPnt'"
'      End
'    End If
'  End If
'  match = rsSchedule!matchNumber
'  'grp = Left(rsSchedule!code1, 1)
'
'  pnt = 0
'  sqlstr = "Select * from tblPrediction_Finals"
'  sqlstr = sqlstr & " Where competitorPoolID = " & poolFormID
'  sqlstr = sqlstr & " AND matchNumber >= " & matchNrFirst
'  sqlstr = sqlstr & " AND matchnumber <= " & matchNrLast
'  rsPoolFormFinals.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'
'  Do While Not rsPoolFormFinals.EOF
'    If tmPos = "links" Then
'      zkTeam = rsPoolFormFinals!teamnameA
'      zkTeam2 = rsPoolFormFinals!teamnameB
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) And match = rsPoolFormFinals!matchNumber Then
'        pnt = pnt + getPointsForID(10, cn) '"halve finale team"
'        Exit Do
'      End If
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) Or zkTeam2 = getTeamIdFromCode(rsSchedule(vldNaam), cn) Then
'        pnt = pnt + getPointsForID(9, cn) '"halve finale positie"
'        Exit Do
'      End If
'    Else
'      zkTeam = rsPoolFormFinals!teamnameB
'      zkTeam2 = rsPoolFormFinals!teamnameA
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) And match = rsPoolFormFinals!matchNumber Then
'        pnt = pnt + getPointsForID(10, cn) '"halve finale positie"
'        Exit Do
'      End If
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) Or zkTeam2 = getTeamIdFromCode(rsSchedule(vldNaam), cn) Then
'        pnt = pnt + getPointsForID(9, cn) '"halve finale team"
'        Exit Do
'      End If
'    End If
'    rsPoolFormFinals.MoveNext
'  Loop
'  rsPoolFormFinals.Close
'  Set rsPoolFormFinals = Nothing
'
'  rsSchedule.Close
'  Set rsSchedule = Nothing
'
'  getPoolFormFinal2Points = pnt
'
'End Function

'Function getPoolFormFinalPoints(poolFormID As Long, matchNr As Integer, cn As ADODB.Connection)
''punten voor de kwart finalisten
'
'Dim rsSchedule As New ADODB.Recordset
'Dim rsPoolFormFinals As New ADODB.Recordset
'Dim sqlstr As String
'Dim grp As String
'Dim match As Integer
'Dim zkTeam As Long
'Dim zkTeam2 As Long
'Dim vldNaam As String
'Dim pnt As Integer
'Dim tmPos As String
'Dim matchNrFirst As Integer
'Dim matchNrLast As Integer
'  matchNrFirst = getFinalmatchOrder(4, True, cn)
'  matchNrLast = getFinalmatchOrder(4, False, cn)
'  sqlstr = "Select * from tblTournamentSchedule where tournamentID= " & thisTournament
'  sqlstr = sqlstr & " AND macthTeamA = 'W" & matchNr & "'"
'  rsSchedule.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'  vldNaam = "matchTeamA"
'  tmPos = "links"
'  If rsSchedule.EOF And rsSchedule.BOF Then 'moeten we het andere team hebben
'    rsSchedule.Close
'    sqlstr = "Select * from tblTournamentSchedule where tournamentID= " & thisTournament
'    sqlstr = sqlstr & " AND macthTeamB = 'W" & matchNr & "'"
'    rsSchedule.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    vldNaam = "matchTeamB"
'    tmPos = "rechts"
'    If rsSchedule.EOF And rsSchedule.BOF Then 'zou niet motte magge
'      MsgBox "Sorry, er zit een foutje in het programma, kan niet verder gaan", vbOKOnly + vbCritical, "Routine 'getHvFinPnt'"
'      End
'    End If
'  End If
'  match = rsSchedule!matchNumber
'  'grp = Left(rsSchedule!code1, 1)
'
'  pnt = 0
'  sqlstr = "Select * from tblPrediction_Finals"
'  sqlstr = sqlstr & " Where competitorPoolID = " & poolFormID
'  sqlstr = sqlstr & " AND matchNumber >= " & matchNrFirst
'  sqlstr = sqlstr & " AND matchnumber <= " & matchNrLast
'  rsPoolFormFinals.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'
'  Do While Not rsPoolFormFinals.EOF
'    If tmPos = "links" Then
'      zkTeam = rsPoolFormFinals!teamnameA
'      zkTeam2 = rsPoolFormFinals!teamnameB
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) And match = rsPoolFormFinals!matchNumber Then
'        pnt = pnt + getPointsForID(12, cn) '" finale positie"
'        Exit Do
'      End If
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) Or zkTeam2 = getTeamIdFromCode(rsSchedule(vldNaam), cn) Then
'        pnt = pnt + getPointsForID(11, cn) '" finale team"
'        Exit Do
'      End If
'    Else
'      zkTeam = rsPoolFormFinals!teamnameB
'      zkTeam2 = rsPoolFormFinals!teamnameA
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) And match = rsPoolFormFinals!matchNumber Then
'        pnt = pnt + getPointsForID(12, cn) '" finale positie"
'        Exit Do
'      End If
'      If zkTeam = getTeamIdFromCode(rsSchedule(vldNaam), cn) Or zkTeam2 = getTeamIdFromCode(rsSchedule(vldNaam), cn) Then
'        pnt = pnt + getPointsForID(11, cn) '" finale team"
'        Exit Do
'      End If
'    End If
'    rsPoolFormFinals.MoveNext
'  Loop
'  rsPoolFormFinals.Close
'  Set rsPoolFormFinals = Nothing
'
'  rsSchedule.Close
'  Set rsSchedule = Nothing
'
'  getPoolFormFinalPoints = pnt
'
'End Function

Function getTournamentStandingPoints(poolFormID As Long, cn As ADODB.Connection)
Dim rs As ADODB.Recordset
Dim sqlstr As String
Dim pts As Integer
Dim i As Integer
Dim podiumPlaces As Integer
  If getTournamentInfo("tournamentThirdPlace", cn) Then
    podiumPlaces = 4
  Else
    podiumPlaces = 2
  End If

  sqlstr = "SELECT * FROM competitorPools "
  sqlstr = sqlstr & " WHERE poolID = " & thisPool
  sqlstr = sqlstr & " AND competitorID = " & poolFormID
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  pts = 0
  With rs
    For i = 1 To podiumPlaces
      If rs("predictionTeam" & i) = getTournamentFinalTeam(i, cn) Then
        pts = pts + getPointsFor(Format(i, "0") & "e plaats", cn)
      End If
    Next
    .Close
  End With
  getTournamentStandingPoints = pts
  Set rs = Nothing
End Function


Function getTournamentFinalTeam(place As Integer, cn As ADODB.Connection)
  Dim rs As ADODB.Recordset
  Dim sqlstr As String
  Set rs = New ADODB.Recordset
  Dim retTeam As Long
  sqlstr = "Select r.winner, r.teamA_ID, r.teamB_ID, s.matchType  "
  sqlstr = sqlstr & " from tblMatchResults r "
  sqlstr = sqlstr & " INNER JOIN tblTournamentSchedule s ON (s.matchOrder = r.matchOrder) "
  sqlstr = sqlstr & " AND (s.tournamentID  = r.tournamentID) "
  sqlstr = sqlstr & " WHERE s.tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND s.matchPlayed = True"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  Select Case place
    Case 1
      rs.Find "matchtype = 4"  'finale
      retTeam = rs!winner
    Case 2
      rs.Find "matchtype = 4"  'finale
      If rs!teamA_ID <> rs!winner Then
        retTeam = rs!teamA_ID
      Else
        retTeam = rs!teamB_ID
      End If
    Case 3
      rs.Find "matchtype = 7"  'kl finale
      retTeam = rs!winner
    Case 4
      rs.Find "matchtype = 7"  'kl finale
      If rs!teamA_ID <> rs!winner Then
        retTeam = rs!teamA_ID
      Else
        retTeam = rs!teamB_ID
      End If
  End Select
  getTournamentFinalTeam = retTeam
End Function


Function getPoolFormTopScoresPoints(poolFormID As Long, cn As ADODB.Connection)
Dim rs As ADODB.Recordset
Dim rsTS As ADODB.Recordset

Dim sqlstr As String
Dim goals As Integer
Dim i As Integer
Dim ts(2) As Long
Dim pts(2) As Integer

Set rs = New ADODB.Recordset
Set rsTS = New ADODB.Recordset

  pts(1) = getPointsForID(21, cn)
  pts(2) = getPointsForID(24, cn)
  sqlstr = "Select * from tblPredictionTopscorers "
  sqlstr = sqlstr & " WHERE competitorPololID = " & poolFormID
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
  sqlstr = "SELECT playerID, count(playerID)  as aantal from tblMatchEvents"
  sqlstr = sqlstr & " WHERE eventID <=2"
  sqlstr = sqlstr & " GROUP BY tournamentID, playerID"
  sqlstr = sqlstr & " HAVING tournamentID = " & thisTournament
  sqlstr = sqlstr & " ORDER BY count(playerID) DESC"
  rsTS.Open sqlstr, cn, adOpenStatic, adLockReadOnly
  rsTS.MoveFirst
  goals = rsTS!aantal
  If rsTS!playerID = rs!topscorererPlayerID Then
    pts(0) = pts(0) + pts(1)
  End If
  If goals = rs!topscorererGoals Then
    pts(0) = pts(0) + pts(2)
  End If
  getPoolFormTopScoresPoints = pts(0)
rs.Close
Set rs = Nothing


End Function

Function getStatsPointsFor(statType As Integer, poolFormID As Long, cn As ADODB.Connection)
'get the points for the specific stattype
Dim rs As ADODB.Recordset
Dim sqlstr As String
Dim realStat As Integer
Dim pnt As Integer
Dim marge As Byte
Dim retPts As Integer
Dim matchOrdrNr As Integer
  matchOrdrNr = getLastMatchPlayed(cn)
  realStat = getStatistics(matchOrdrNr, statType, cn)
  marge = getMarginForID(CLng(statType), cn)
  sqlstr = "Select * from tblPrediction_Numbers "
  sqlstr = sqlstr & " WHERE competitorPoolID = " & poolFormID
  sqlstr = sqlstr & " AND predictionTypeID = " & statType
  Set rs = New ADODB.Recordset
  With rs
    .Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    pnt = getPointsForID(CLng(statType), cn)
    If pnt > 0 Then
      If !predictionNumber >= realStat - marge And !predictionNumber <= realStat + marge Then
        retPts = pnt
      End If
    End If
    .Close
  End With
  Set rs = Nothing
  getStatsPointsFor = retPts
End Function

Function getStatsPoints(poolFormID As Long, cn As ADODB.Connection)
Dim rsDeeln As ADODB.Recordset
Dim sqlstr As String
Dim dp As Integer
Dim gelijk As Integer
Dim gele As Integer
Dim rode As Integer
Dim pens As Integer

Dim pnt As Integer
Dim marge As Byte
Dim deelnpnt As Integer
Dim wd As Integer 'stores matchOrder nr
Set rsDeeln = New ADODB.Recordset
Dim statPnt(5) As Integer

sqlstr = "Select * from tblPrediction_Numbers "
sqlstr = sqlstr & " WHERE competitorIPoolID = " & poolFormID
rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
wd = getLastMatchPlayed(cn) 'returns matchOrder
dp = getStatistics(wd, doelp, cn) + getStatistics(wd, penalty, cn) + getStatistics(wd, eigdoelp, cn)
pens = getStatistics(wd, penalty, cn) + getStatistics(wd, penaltyMis, cn)
gele = getStatistics(wd, geel, cn)
rode = getStatistics(wd, rood, cn)
gelijk = getDrawCount(wd, cn)

deelnpnt = 0
If rsDeeln.RecordCount = 0 Then Exit Function
With rsDeeln
  pnt = getPointsForID(dpAant, cn)
  If pnt > 0 Then
      .Find "predictionTypeID = " & dpAant
      marge = getMarginForID(dpAant, cn)
      If Not .EOF Then
          If !predictionNumber >= dp - marge And !predictionNumber <= dp + marge Then
              deelnpnt = deelnpnt + pnt
          End If
      End If
  End If
  pnt = getPointsForID(gelijkSp, cn)
  If pnt > 0 Then
      .Find "predictionTypeID = " & gelijkSp
      marge = getMarginForID(gelijkSp, cn)
      If Not .EOF Then
          If !predictionNumber >= gelijk - marge And !predictionNumber <= gelijk + marge Then
              deelnpnt = deelnpnt + pnt
          End If
      End If
  End If
  pnt = getPointsForID(geelKrt, cn)
  If pnt > 0 Then
      .Find "predictionTypeID = " & geelKrt
      marge = getMarginForID(geelKrt, cn)
      If Not .EOF Then
          If !predictionNumber >= gele - marge And !predictionNumber <= gele + marge Then
              deelnpnt = deelnpnt + pnt
          End If
      End If
  End If
  pnt = getPointsForID(roodKrt, cn)
  If pnt > 0 Then
      .Find "predictionTypeID = " & roodKrt
      marge = getMarginForID(roodKrt, cn)
      If Not .EOF Then
          If !predictionNumber >= rode - marge And !predictionNumber <= rode + marge Then
              deelnpnt = deelnpnt + pnt
          End If
      End If
  End If
  pnt = getPointsForID(pensAant, cn)
  If pnt > 0 Then
      .Find "predictionTypeID = " & pensAant
      marge = getMarginForID(pensAant, cn)
      If Not .EOF Then
          If !predictionNumber >= pens - marge And !predictionNumber <= pens + marge Then
              deelnpnt = deelnpnt + pnt
          End If
      End If
  End If
End With
rsDeeln.Close
Set rsDeeln = Nothing

getStatsPoints = deelnpnt


End Function


Function getTotalThisDayPoints(poolFormID As Long, matchNr As Integer, cn As ADODB.Connection)
Dim rs As New ADODB.Recordset
Dim rsPnt As New ADODB.Recordset
Dim sqlstr As String
Dim pnt As Integer
  If matchNr > 0 Then
    sqlstr = " Select matchOrder from tblTournamentSchedule where "
    sqlstr = sqlstr & " clng(matchDate) = " & CLng(getMatchInfo(matchNr, "matchDate", cn))
    sqlstr = sqlstr & " AND matchPlayed = true"
    
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    pnt = 0
    Do While Not rs.EOF
      sqlstr = "Select pointsDay from tblCompetitorPoints"
      sqlstr = sqlstr & " WHERE matchOrder = " & rs!matchOrder
      sqlstr = sqlstr & " AND competitorPoolID = " & poolFormID
      rsPnt.Open sqlstr, cn, adOpenStatic, adLockReadOnly
      If rsPnt.RecordCount > 0 Then
        pnt = pnt + rsPnt!pointsDay
      End If
      rs.MoveNext
      rsPnt.Close
      Set rsPnt = Nothing
    Loop
    rs.Close
    Set rs = Nothing
  Else
    pnt = 0
  End If
  getTotalThisDayPoints = pnt
  rs.Close
  rsPnt.Close
  Set rs = Nothing
  Set rsPnt = Nothing
End Function

Function getMatchOrder(matchNumber As Integer, cn As ADODB.Connection)
  Dim rs As ADODB.Recordset
  Dim sqlstr As String
  Set rs = New ADODB.Recordset
  sqlstr = "SELECT matchOrder FROM tblTournamentSchedule"
  sqlstr = sqlstr & "  WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchnumber = " & matchNumber
  rs.Open sqlstr, cn, adOpenForwardOnly, adLockReadOnly
  If Not rs.EOF Then
    getMatchOrder = rs!matchOrder
  Else
    getMatchOrder = 0
  End If
  rs.Close
  Set rs = Nothing

End Function

Function getMatchNumber(matchOrder As Integer, cn As ADODB.Connection)
  Dim rs As ADODB.Recordset
  Dim sqlstr As String
  Set rs = New ADODB.Recordset
  sqlstr = "SELECT matchNumber FROM tblTournamentSchedule"
  sqlstr = sqlstr & "  WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  rs.Open sqlstr, cn, adOpenForwardOnly, adLockReadOnly
  If Not rs.EOF Then
    getMatchNumber = rs!matchNumber
  Else
    getMatchNumber = 0
  End If
  rs.Close
  Set rs = Nothing

End Function

Function getWinners(final As Boolean, cn As ADODB.Connection) As Long()
'return an array with the winners (1-2)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sqlstr As String
Dim winners(2) As Long
Dim finalType As Integer
  If final Then
    finalType = 4
  Else
    finalType = 7
  End If
  sqlstr = "Select winner, teamA_ID, teamB_ID from tblMatchresults"
  sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchOrder = " & getFinalmatchOrder(finalType, True, cn)
  rs.Open sqlstr, cn, adOpenForwardOnly, adLockReadOnly
  If Not rs.EOF Then
    winners(1) = rs!winner
    If rs!winner = rs!teamA_ID Then
      winners(2) = rs!teamB_ID
    Else
      winners(2) = rs!teamA_ID
    End If
  End If
  getWinners = winners
rs.Close
Set rs = Nothing
End Function

Sub addAllPlayerstoTournament()
'temporary routine to add the new people table to the tourament
Dim rs As New ADODB.Recordset
Dim cn As ADODB.Connection
Dim sqlstr As String
Dim teamCode As Integer
  Set cn = New ADODB.Connection
  With cn
    .ConnectionString = lclConn
    .Open
  End With

  sqlstr = "Select * from tblPeople"
  rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  
  Do While Not rs.EOF
    'get the tournament teamname/country
    teamCode = getTeamIDFromCountryCode(rs!countryCode, cn)
    sqlstr = "INSERT into tblTeamPlayers(tournamentID, teamID, playerID"
    sqlstr = sqlstr & ") VALUES (" & 17
    sqlstr = sqlstr & ", " & teamCode
    sqlstr = sqlstr & ", " & rs!peopleID
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    rs.MoveNext
  Loop
  MsgBox "Alle namen uit de tabel tblPeople toegevoegd aan toernooi", vbOKOnly, "Namenlijst"
End Sub
