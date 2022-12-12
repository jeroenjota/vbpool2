Attribute VB_Name = "calc"
Option Explicit

Dim goalsPerDay As Integer

''''
'Beslissingscriteria gelijke stand groepsfase EK 2021
'1. Meeste punten in de groepswedstrijden tegen andere ploegen met gelijk aantal punten.
'2. Doelpuntensaldo als resultaat van de groepswedstrijden tegen andere ploegen met gelijk aantal punten.
'3. Meeste doelpunten gescoord in de groepswedstrijden tegen andere ploegen met gelijk aantal punten.
'4. Wanneer teams na het toepassen van de eerste drie criteria nog steeds gelijk staan, zullen criteria 1 tot en met 3 opnieuw worden toegepast, maar nu alleen tussen de teams in kwestie. Wanneer teams nog steeds gelijk staan, worden criteria 5 tot en met 9 toegepast.
'5. Het doelpuntensaldo over alle groepswedstrijden.
'6. Meeste doelpunten gescoord in alle groepswedstrijden.
'7. Als twee teams hetzelfde aantal punten hebben en ook gelijk staan volgens criteria 1 tot en met 6 nadat ze in de laatste wedstrijd van de groepsfase tegen elkaar hebben gespeeld, wordt hun positie bepaald door een strafschoppenserie. Dit criterium wordt niet gebruikt als er meer dan twee teams op hetzelfde aantal punten eindigen.
'8. Fair Play-ranglijst (1 punt voor een losse gele kaart, 3 punten voor een rode kaart als gevolg van twee gele kaarten, 3 punten voor een directe rode kaart en 4 punten voor een gele kaart gevolgd door een directe rode kaart).
'9. Positie op de UEFA-coëfficiëntenranglijst
''''
'

'This module contains various calculations

'*****************************************************
' Purpose: fill the competitorPoolPoints table with the points for various form fields after a match
'
' Inputs: matchOrder
'
' Returns: -
'*****************************************************

Function add3rdPlacePoints(thisPoolForm As Long, matchOrder As Integer, cn As ADODB.Connection)
'search for teams that  where 3rd, but in the 8th finals
  Dim rsTeamCodes As ADODB.Recordset
  Dim rs As ADODB.Recordset
  Dim sqlstr As String
  Set rsTeamCodes = New ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim matchNrFirst As Integer
  Dim matchNrLast As Integer
  Dim pts(2) As Integer
  Dim grp As String
  matchNrFirst = getFinalmatchOrder(5, True, cn)
  matchNrLast = getFinalmatchOrder(5, False, cn)
  pts(1) = getPointsForID(4, cn)
  pts(2) = getPointsForID(5, cn)
  
  sqlstr = "Select * from tblTournamentTeamCodes "
  sqlstr = sqlstr & " WHERE left(teamCode,1) = '3'"
  sqlstr = sqlstr & " AND tournamentID =" & thisTournament
  rsTeamCodes.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  
  sqlstr = "Select * from tblPrediction_Finals "
  sqlstr = sqlstr & " WHERE competitorPoolID =  " & thisPoolForm
  sqlstr = sqlstr & " AND matchOrder BETWEEN " & matchNrFirst & " AND " & matchNrLast
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  With rsTeamCodes
    Do While Not .EOF
      grp = getGroup(rsTeamCodes!teamID, cn)
      rs.Find "teamNameA = " & !teamID
      If Not rs.EOF Then
      'bijna altijd gebruiken we matchOrder als ID, maar in dit geval matchNumber
        If rs!matchNumber = getmatchOrderfromCode(!teamCode, cn) Then
          pts(0) = pts(2)
        Else
          pts(0) = pts(1)
        End If
        Exit Do
      End If
      rs.Find "teamNameB = " & !teamID
      If Not rs.EOF Then
      'bijna altijd gebruiken we matchOrder als ID, maar in dit geval matchNumber
        If rs!matchNumber = getmatchOrderfromCode(!teamCode, cn) Then
          pts(0) = pts(1)
        Else
          pts(0) = pts(2)
        End If
        Exit Do
      End If
      .MoveNext
    Loop
    If pts(0) > 0 Then
      'If thisPoolForm = 7 Then Stop
      sqlstr = "UPDATE tblCompetitorPoints SET "
      sqlstr = sqlstr & " pointsTeamsFinals8" & grp & " = [pointsTeamsFinals8" & grp & "] + " & pts(0)
      sqlstr = sqlstr & ", pointsFinals_8 = [pointsFinals_8] + " & pts(0)
'      sqlstr = sqlstr & ", pointsDay = [pointsDay] + " & pts(0)
'      sqlstr = sqlstr & ", pointsDayTotal = [pointsDayTotal] + " & pts(0)
'      sqlstr = sqlstr & ", pointsGrandTotal = [pointsGrandTotal] + " & pts(0)
      sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
      sqlstr = sqlstr & " AND matchOrder = " & matchOrder
      cn.Execute sqlstr
    End If
    add3rdPlacePoints = pts(0)
  End With
'and add the points to the table

End Function

Sub updatePoolFormPoints(matchOrder As Integer, cn As ADODB.Connection)
  'calculate all the participant poolforms for matchOrder and return the totalpoints for this match
  Dim savdat As Date
  Dim sqlstr As String
  Dim rs As ADODB.Recordset
  Dim poolFormName As String
  Dim thisForm As Long
  Dim teamCode() As String
  Dim teamID(2) As Long
  Dim team() As String
  Dim group As String
  Dim matchPts As Integer
  Dim grpPts As Integer
  Dim points As Integer
  Dim totalPoints As Integer
  Dim dayPoints As Integer
  Dim final8Pts As Integer
  Dim matchType As Integer
  Dim matchNr As Integer
  'we have changed all references to matchnumber to matchOrder, which is also the ordernumber of the match
  'get the date of this match
  savdat = getMatchInfo(matchOrder, "matchDate", cn)
  'the following 2 'get'functions return an array with teams / teamcodes
  teamCode = getMatchTeamCodes(matchOrder, cn)
  team = getScheduleTeamNames(matchOrder, cn)
  teamID(1) = getTeamIdFromCode(teamCode(0), cn)
  teamID(2) = getTeamIdFromCode(teamCode(1), cn)
  
  Set rs = New ADODB.Recordset
  'goalsperDay is a global variable
  goalsPerDay = getGoalsPerDay(savdat, cn)
  sqlstr = "Select * from tblCompetitorPools WHERE poolId = " & thisPool
  sqlstr = sqlstr & " ORDER BY nickname"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  matchNr = getMatchNumber(matchOrder, cn)
  Do While Not rs.EOF
    
    thisForm = rs!competitorPoolID
'    If thisForm = 87 Then Stop
 '   If thisForm = 100 Then Stop
    poolFormName = rs!nickName & " (" & rs.AbsolutePosition & "/" & rs.RecordCount & ")"
    'show some info
    showInfo True, "Punten tellen", "Na " & matchOrder & "e wedstrijd (nr " & matchNr & "): " & team(0) & " - " & team(1), poolFormName
    DoEvents
    'clear all points records from this match onward. This to be sure that calculations are correct
    sqlstr = "Delete from tblCompetitorPoints "
    sqlstr = sqlstr & " WHERE competitorPoolID = " & thisForm
    sqlstr = sqlstr & " AND matchOrder >= " & matchOrder
    cn.Execute sqlstr
    'add a record for thisForm and matchOrder
    sqlstr = "INSERT INTO tblCompetitorPoints (competitorPoolID, matchOrder)"
    sqlstr = sqlstr & "VALUES (" & thisForm & ", " & matchOrder & ")"
    cn.Execute sqlstr
    'get the points for matches (and daily goals)
    points = getPoolFormMatchPoints(thisForm, matchOrder, cn)
    dayPoints = points
    matchType = getMatchInfo(matchOrder, "matchtype", cn)
    If matchType = 1 Then 'still in the group stage
    'check if all the matches in the group are played and calculate the (8th) finalists points
      group = getGroup(teamID(1), cn)
      If grpPlayedCount(group, cn) = 6 And getLastGroupMatch(group, cn) = matchOrder Then
        'all matches in group are played and this is the last match of the group
        'get the points fro the group standings
        points = getGroupStandPoints(thisForm, matchOrder, cn)
        dayPoints = dayPoints + points
        'and check the teams in the next round (only 1st eand 2nd places, best 3rd places not yet known)
       ' If thisForm = 7 Then Stop
        points = get8FinPts(thisForm, group, matchOrder, cn)
        dayPoints = dayPoints + points
      End If
      If getMatchCount(1, cn) = matchOrder Then 'kijk of deze wedstrijd de laatste is van de groepsfase
       'recalcultae all the 8th finalists
        dayPoints = dayPoints + add3rdPlacePoints(thisForm, matchOrder, cn)
      End If
    Else 'we are playing the final rounds
      If matchType = 5 Then '8e finale
       'If thisForm = 7 Then Stop
        points = get4finalistPoints(thisForm, matchOrder, cn)
        dayPoints = dayPoints + points
      End If
      If matchType = 2 Then 'kwart finale
      'check the poolForm semifinalists
        points = getSemifinalistPoints(thisForm, matchOrder, cn)
        dayPoints = dayPoints + points
      End If
      If matchType = 3 Then 'halve finale
      'check the poolForm finalists/3rd place match
        points = getfinalistPoints(thisForm, matchOrder, True, cn) ' 3rd place match
        points = points + getfinalistPoints(thisForm, matchOrder, False, cn)
        dayPoints = dayPoints + points
      End If
      If matchType = 7 Then '3rd place
        
        'If thisForm = 116 Then Stop

        points = getTournamentStandingPoints(thisForm, matchOrder, cn)
        dayPoints = dayPoints + points
      End If
      If matchType = 4 Then 'THE FINAL ;-)
       'calculate match, tournament reslts, and all the statistics
       'If thisForm = 23 Then Stop
       points = getEndWinnerPoints(thisForm, matchOrder, cn)
       dayPoints = dayPoints + points
      End If
    End If
    'If thisForm = 73 Then Stop
    totalPoints = getTotalDayPoints(thisForm, matchOrder, savdat, cn) + dayPoints
    sqlstr = "UPDATE tblCompetitorPoints set pointsDay = " & dayPoints '+ final8Pts
    sqlstr = sqlstr & ", pointsDayTotal = " & totalPoints
    sqlstr = sqlstr & ", pointsGrandTotal = " & calcGrandTotal(thisForm, matchOrder, cn) + dayPoints
    sqlstr = sqlstr & " WHERE competitorpoolID = " & thisForm
    sqlstr = sqlstr & " AND matchOrder = " & matchOrder
    cn.Execute sqlstr
    
    rs.MoveNext
  Loop
End Sub

Function getTournamentStandingPoints(thisForm As Long, matchOrder As Integer, cn As ADODB.Connection)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim i As Integer
Dim final As Boolean
Dim place As Integer
Dim teamPos As Integer
Dim sqlstr As String
Dim pts(2) As Integer
Dim ttlPts As Integer
Dim winners() As Long
  final = matchOrder = getFinalmatchOrder(4, True, cn)
  winners = getWinners(final, cn)
  If final Then
    pts(1) = getPointsForID(15, cn) '3rdplace
    pts(2) = getPointsForID(14, cn) '4th place
    place = 1
  Else  '3rd place
    pts(1) = getPointsForID(13, cn) '3rdplace
    pts(2) = getPointsForID(29, cn) '4th place
    place = 3
  End If
  sqlstr = "Select * from tblCompetitorPools"
  sqlstr = sqlstr & " WHERE competitorPoolID = " & thisForm
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  'get the points for the final standing
  teamPos = 0
  If Not rs.EOF Then
    For i = place To place + 1
      teamPos = teamPos + 1
      If rs("predictionTeam" & i) = winners(teamPos) Then
        pts(0) = pts(0) + pts(teamPos)
      End If
    Next
    ttlPts = pts(0)
  End If
  'If thisForm = 81 Then Stop
  sqlstr = "UPDATE tblCompetitorPoints SET "
  sqlstr = sqlstr & " pointsFinalStanding = " & ttlPts
  sqlstr = sqlstr & ", pointsGroupStanding = " & getPoolFormPoints(thisForm, matchOrder - 1, 7, cn)
  sqlstr = sqlstr & ", pointsFinals_8 = " & getPoolFormPoints(thisForm, matchOrder - 1, 16, cn)
  sqlstr = sqlstr & ", pointsFinals_4 = " & getPoolFormPoints(thisForm, matchOrder - 1, 25, cn)
  sqlstr = sqlstr & ", pointsFinals_2 = " & getPoolFormPoints(thisForm, matchOrder - 1, 30, cn)
  sqlstr = sqlstr & ", pointsFinal = " & getPoolFormPoints(thisForm, matchOrder - 1, 36, cn) ' + ttlPts
  sqlstr = sqlstr & " WHERE competitorPoolID = " & thisForm
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  cn.Execute sqlstr

  getTournamentStandingPoints = ttlPts
  rs.Close
End Function

Function getEndWinnerPoints(thisForm As Long, matchOrder As Integer, cn As ADODB.Connection)
'check the end ranking
''''''''''''''''''''''''
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sqlstr As String
Dim ptsTs As Integer
Dim ptsStats(5) As Integer
Dim ttlPts As Integer
Dim rsTS As ADODB.Recordset
Set rsTS = New ADODB.Recordset
Dim realCnt As Integer
Dim i As Integer

  ttlPts = getTournamentStandingPoints(thisForm, matchOrder, cn)
  'get the topscorer(s) points
  sqlstr = "Select * from tblPredictionTopScorers WHERE competitorPoolID = " & thisForm
  sqlstr = sqlstr & " AND topscorerPosition = 1"  'not realy necessary, as long we only have one topscorer
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  
  sqlstr = "SELECT top 1 playerId, Count(playerId) AS cnt"
  sqlstr = sqlstr & " From tblMatchEvents"
  sqlstr = sqlstr & " Where eventid <= 2"
  sqlstr = sqlstr & " GROUP BY tournamentId,playerId"
  sqlstr = sqlstr & " Having tournamentId = 16"
  sqlstr = sqlstr & " ORDER BY Count(playerId) DESC;"
  rsTS.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  'loop through possibly more than one topscorer
  Do While Not rsTS.EOF
    realCnt = rsTS!cnt
    If rsTS!playerID = rs!topScorerPlayerID Then
      ptsTs = getPointsForID(21, cn)
      Exit Do
    End If
    rsTS.MoveNext
  Loop
  rsTS.Close
  Set rsTS = Nothing
  If rs!topScorergoals = realCnt Then
    ptsTs = ptsTs + getPointsForID(24, cn)
  End If
  'ttlPts = ttlPts + ptsTs
  rs.Close
  Set rs = Nothing
'get the statistics points
  ptsStats(0) = 0
  For i = dpAant To pensAant
    ptsStats(i - dpAant + 1) = getStatsPointsFor(i, thisForm, cn)
    ptsStats(0) = ptsStats(0) + ptsStats(i - dpAant + 1)
    '(that was quick ;-)
  Next
  'ttlPts = ttlPts + ptsStats(0)
    
'now insert the points into the competitorPoints table
'If thisForm = 23 Then Stop
  sqlstr = "UPDATE tblCompetitorPoints SET "
  sqlstr = sqlstr & " pointsFinalStanding = " & ttlPts
  sqlstr = sqlstr & ", pointsTopscorers = " & ptsTs
  sqlstr = sqlstr & ", pointsOther = " & ptsStats(0)
  sqlstr = sqlstr & ", pointsGroupStanding = " & getPoolFormPoints(thisForm, matchOrder - 1, 7, cn)
  sqlstr = sqlstr & ", pointsFinals_8 = " & getPoolFormPoints(thisForm, matchOrder - 1, 16, cn)
  sqlstr = sqlstr & ", pointsFinals_4 = " & getPoolFormPoints(thisForm, matchOrder - 1, 25, cn)
  sqlstr = sqlstr & ", pointsFinals_2 = " & getPoolFormPoints(thisForm, matchOrder - 1, 30, cn)
  sqlstr = sqlstr & ", pointsFinal = " & getPoolFormPoints(thisForm, matchOrder - 1, 36, cn)
  sqlstr = sqlstr & " WHERE competitorPoolID = " & thisForm
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  cn.Execute sqlstr
  'return the total
  getEndWinnerPoints = ttlPts + ptsTs + ptsStats(0)
End Function

Function getfinalistPoints(thisForm As Long, matchOrder As Integer, small As Boolean, cn As ADODB.Connection)
'check the two finalists
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim rsFin As ADODB.Recordset
Set rsFin = New ADODB.Recordset
Dim sqlstr As String
Dim pts(2) As Integer
Dim finMatchNumber As Integer
Dim leftTeam As Boolean
Dim grp As String
Dim ttlPts As Integer
Dim matchNumber As Integer
  matchNumber = getMatchNumber(matchOrder, cn)
  leftTeam = matchNumber = getFinalmatchOrder(3, True, cn)
  sqlstr = "Select * from tblTournamentTeamCodes "
  sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND teamID > 0"
  If small Then
    sqlstr = sqlstr & " AND teamCode = 'V" & matchNumber & "'"
    pts(1) = getPointsForID(30, cn)
    pts(2) = getPointsForID(31, cn)
    finMatchNumber = getFinalmatchOrder(7, True, cn)
  Else
    sqlstr = sqlstr & " AND teamCode = 'W" & matchNumber & "'"
    pts(1) = getPointsForID(11, cn)
    pts(2) = getPointsForID(12, cn)
    finMatchNumber = getFinalmatchOrder(4, True, cn)
  End If
  rsFin.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rsFin.EOF Then
    sqlstr = "Select * from tblPrediction_Finals"
    sqlstr = sqlstr & " WHERE competitorPoolId = " & thisForm
    If small Then
      sqlstr = sqlstr & " AND matchOrder between " & getFinalmatchOrder(7, True, cn)
      sqlstr = sqlstr & " AND " & getFinalmatchOrder(7, False, cn)
    Else
      sqlstr = sqlstr & " AND matchOrder between " & getFinalmatchOrder(4, True, cn)
      sqlstr = sqlstr & " AND " & getFinalmatchOrder(4, False, cn)
    End If
    sqlstr = sqlstr & " AND (teamnameA = " & rsFin!teamID
    sqlstr = sqlstr & " or teamNameB = " & rsFin!teamID & ")"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    Do While Not rs.EOF  'found a record so at least some points
      'If thisForm = 119 Then Stop
       pts(0) = pts(1)
      'found the team check the place
      If rs!matchOrder = finMatchNumber Then
        If rs!teamnameA = rsFin!teamID And leftTeam Then pts(0) = pts(2)
        If rs!teamnameB = rsFin!teamID And Not leftTeam Then pts(0) = pts(2)
        Exit Do
      End If
      rs.MoveNext
    Loop
    ttlPts = ttlPts + pts(0)
    rs.Close
  End If
  rsFin.Close
  Set rsFin = Nothing
  Set rs = Nothing
  If small Then
    sqlstr = "UPDATE tblCompetitorPoints SET pointsTeamsFinals34 = " & ttlPts
  Else
   ' If ttlPts > 0 Then Stop
    sqlstr = "UPDATE tblCompetitorPoints SET pointsTeamsFinal = " & ttlPts
  End If
  'copy values from previous match for reposrts later
  sqlstr = sqlstr & ", pointsGroupStanding = " & getPoolFormPoints(thisForm, matchOrder - 1, 7, cn)
  sqlstr = sqlstr & ", pointsFinals_8 = " & getPoolFormPoints(thisForm, matchOrder - 1, 16, cn)
  sqlstr = sqlstr & ", pointsFinals_4 = " & getPoolFormPoints(thisForm, matchOrder - 1, 25, cn)
  sqlstr = sqlstr & ", pointsFinals_2 = " & getPoolFormPoints(thisForm, matchOrder - 1, 30, cn)
  'total field for final teams
  sqlstr = sqlstr & ", pointsFinal = " & getTotalGroupTtl(thisForm, matchOrder - 1, 36, 1, cn) + ttlPts
  sqlstr = sqlstr & " WHERE competitorPoolID = " & thisForm
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  cn.Execute sqlstr
  getfinalistPoints = ttlPts
End Function

Function get4finalistPoints(thisForm As Long, matchOrder As Integer, cn As ADODB.Connection)
'check the 4 quarter finalists
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim rsFin As ADODB.Recordset
Set rsFin = New ADODB.Recordset
Dim sqlstr As String
Dim pts(2) As Integer
Dim finmatchOrder As Integer
Dim leftTeam As Boolean
Dim grp As String
Dim ttlPts As Integer
Dim matchNr As Integer
  matchNr = getMatchNumber(matchOrder, cn)
  pts(1) = getPointsForID(6, cn)
  pts(2) = getPointsForID(7, cn)
  
  Select Case matchNr
  Case 49
    'the quarter final match number and home or away
    finmatchOrder = 57
    leftTeam = True
    grp = "B"
  Case 50
    finmatchOrder = 57
    leftTeam = False
    grp = "B"
   ' Stop
  Case 52
    'the quarter final match number and home or away
    finmatchOrder = 59
    leftTeam = False
    grp = "D"
  Case 51
    finmatchOrder = 59
    leftTeam = True
    grp = "D"
  Case 53
    'the quarter final match number and home or away
    finmatchOrder = 58
    leftTeam = True
    grp = "A"
  Case 54
    finmatchOrder = 58
    leftTeam = False
    grp = "A"
  Case 55
    'the quarter final match number and home or away
    finmatchOrder = 60
    leftTeam = True
    grp = "C"
  Case 56
    finmatchOrder = 60
    leftTeam = False
    grp = "C"
  End Select
  sqlstr = "Select * from tblTournamentTeamCodes "
  sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND teamID > 0"
  sqlstr = sqlstr & " AND teamCode = 'W" & matchNr & "'"
  rsFin.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rsFin.EOF Then
    sqlstr = "Select * from tblPrediction_Finals"
    sqlstr = sqlstr & " WHERE competitorPoolId = " & thisForm
    sqlstr = sqlstr & " AND matchOrder between " & getFinalmatchOrder(2, True, cn)
    sqlstr = sqlstr & " AND " & getFinalmatchOrder(2, False, cn)
    sqlstr = sqlstr & " AND (teamnameA = " & rsFin!teamID
    sqlstr = sqlstr & " or teamNameB = " & rsFin!teamID & ")"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
'    rs.Find "teamNameA = " & rsFin!teamID '& " OR teamNameB = " & rsFin!teamID
    Do While Not rs.EOF  'found a record so at least some points
       pts(0) = pts(1)
      'found the team check the place
      
      '' OOPS ergens was iets fout in de invoer routine, weer eens Order en Number verwisseld denk ik
      '' eigenli9jk is matchorder dus matchnumber
      '' Hersteld met getMatchorder()
      If getMatchOrder(rs!matchOrder, cn) = finmatchOrder Then
        If rs!teamnameA = rsFin!teamID And leftTeam Then pts(0) = pts(2)
        If rs!teamnameB = rsFin!teamID And Not leftTeam Then pts(0) = pts(2)
        Exit Do
      End If
      rs.MoveNext
    Loop
    ttlPts = ttlPts + pts(0)
    rs.Close
  End If
  rsFin.Close
  Set rsFin = Nothing
  Set rs = Nothing
  sqlstr = "UPDATE tblCompetitorPoints SET pointsTeamsFinals4" & grp
  sqlstr = sqlstr & "= " & ttlPts
  'copy values from previous match for reports later
  sqlstr = sqlstr & ", pointsGroupStanding = " & getPoolFormPoints(thisForm, matchOrder - 1, 7, cn)
  sqlstr = sqlstr & ", pointsFinals_8 = " & getPoolFormPoints(thisForm, matchOrder - 1, 16, cn)
  'set the total field for quarter matches
  sqlstr = sqlstr & ", pointsFinals_4 = " & getTotalGroupTtl(thisForm, matchOrder, 26, 4, cn) + ttlPts
  sqlstr = sqlstr & " WHERE competitorPoolID = " & thisForm
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  cn.Execute sqlstr
  get4finalistPoints = ttlPts
End Function

Function getSemifinalistPoints(thisForm As Long, matchOrder As Integer, cn As ADODB.Connection)
'check the 4 quarter finalists
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim rsFin As ADODB.Recordset
Set rsFin = New ADODB.Recordset
Dim sqlstr As String
Dim pts(2) As Integer
Dim finmatchOrder As Integer
Dim leftTeam As Boolean
Dim grp As String
Dim ttlPts As Integer
Dim matchNr As Integer
matchNr = getMatchNumber(matchOrder, cn)
  pts(1) = getPointsForID(9, cn)
  pts(2) = getPointsForID(10, cn)
  Select Case matchNr
  Case 58
    finmatchOrder = 61
    leftTeam = False
    grp = "A"
  Case 57
    finmatchOrder = 61
    leftTeam = True
    grp = "A"
  Case 60
    finmatchOrder = 62
    leftTeam = False
    grp = "B"
  Case 59
    finmatchOrder = 62
    leftTeam = True
    grp = "B"
  End Select
  
  sqlstr = "Select * from tblTournamentTeamCodes "
  sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND teamID > 0"
  sqlstr = sqlstr & " AND teamCode = 'W" & matchNr & "'"
  rsFin.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rsFin.EOF Then
    sqlstr = "Select * from tblPrediction_Finals"
    sqlstr = sqlstr & " WHERE competitorPoolId = " & thisForm
    sqlstr = sqlstr & " AND matchorder between " & getFinalmatchOrder(3, True, cn)
    sqlstr = sqlstr & " AND " & getFinalmatchOrder(3, False, cn)
    sqlstr = sqlstr & " AND (teamnameA = " & rsFin!teamID
    sqlstr = sqlstr & " or teamNameB = " & rsFin!teamID & ")"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    Do While Not rs.EOF
      pts(0) = pts(1)
      'found the team check the place
      If rs!matchOrder = finmatchOrder Then
        If rs!teamnameA = rsFin!teamID And leftTeam Then pts(0) = pts(2)
        If rs!teamnameB = rsFin!teamID And Not leftTeam Then pts(0) = pts(2)
        Exit Do ' if highes possible posititon is found then finish
      End If
      rs.MoveNext
    Loop
    ttlPts = ttlPts + pts(0)
    rs.Close
  End If
  rsFin.Close
  Set rsFin = Nothing
  Set rs = Nothing
  sqlstr = "UPDATE tblCompetitorPoints SET pointsTeamsFinals2" & grp
  sqlstr = sqlstr & "= " & ttlPts
  'copy values from previous match for reposrts later
  sqlstr = sqlstr & ", pointsGroupStanding = " & getPoolFormPoints(thisForm, matchOrder - 1, 7, cn)
  sqlstr = sqlstr & ", pointsFinals_8 = " & getPoolFormPoints(thisForm, matchOrder - 1, 16, cn)
  sqlstr = sqlstr & ", pointsFinals_4 = " & getPoolFormPoints(thisForm, matchOrder - 1, 25, cn)
  sqlstr = sqlstr & ", pointsFinals_2 = " & getTotalGroupTtl(thisForm, matchOrder, 31, 2, cn) + ttlPts
  sqlstr = sqlstr & " WHERE competitorPoolID = " & thisForm
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  cn.Execute sqlstr
  getSemifinalistPoints = ttlPts
End Function

'Function getFinalPoints(poolForm As Long, matchNr As Integer, matchType As Integer, cn As ADODB.Connection)
''check if the winner from the match is in the next round on the poolform
'Dim sqlstr As String
'Dim rs As ADODB.Recordset
'Set rs = New ADODB.Recordset
'Dim rsMatches As ADODB.Recordset
'Set rsMatches = New ADODB.Recordset
'Dim pts(1) As Integer
'Dim winner As Long
'Dim matchNrs(1) As Integer
'Dim points As Integer
'Dim finMatch As Integer
'Dim teamCode As String
'
'  winner = getMatchresult(matchNr, 6, cn)
'  teamCode = getFinalsTeamCode(winner, cn)
'  finMatch = getmatchOrderfromCode(teamCode, cn)
'  Select Case matchType
'  Case 5 '8th finals -> sdo check 4th
'    matchNrs(0) = getFinalmatchOrder(2, True, cn)
'    matchNrs(1) = getFinalmatchOrder(2, False, cn)
'    pts(0) = getPointsForID(6, cn)
'    pts(1) = getPointsForID(7, cn)
'  Case 2 '4th final -> semi's
'    matchNrs(0) = getFinalmatchOrder(3, True, cn)
'    matchNrs(1) = getFinalmatchOrder(3, False, cn)
'    pts(0) = getPointsForID(9, cn)
'    pts(1) = getPointsForID(10, cn)
'  Case Else
'    If getTournamentInfo("tournamentThirdPlace", cn) Then
'      matchNrs(0) = getFinalmatchOrder(7, True, cn)
'      matchNrs(1) = getFinalmatchOrder(7, False, cn)
'      pts(0) = getPointsForID(30, cn)
'      pts(1) = getPointsForID(31, cn)
'      MsgBox "DIT IS NOG NIET GEPROGRAMMEERD"
'      Stop
'    Else
'      matchNrs(0) = getFinalmatchOrder(4, True, cn)
'      matchNrs(1) = getFinalmatchOrder(4, False, cn)
'      pts(0) = getPointsForID(11, cn)
'      pts(1) = getPointsForID(12, cn)
'    End If
'  End Select
'  sqlstr = "Select * from tblPrediction_Finals "
'  sqlstr = sqlstr & " WHERE competitorPoolID = " & poolForm
'  sqlstr = sqlstr & " AND (teamNameA = " & winner & " OR teamNameB = " & winner
'  sqlstr = sqlstr & ") AND matchnumber BETWEEN " & matchNrs(0) & " AND " & matchNrs(1)
'  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
'  If Not rs.EOF Then
'    'found the team
'    points = pts(0)
'    'is it the right position?
'    Do While Not rs.EOF
'      If rs!matchNumber = finMatch Then
'        points = pts(1)
'        Exit Do
'      End If
'      rs.MoveNext
'    Loop
'  End If
'  'update the pointstable
'
'  Select Case matchType
'  Case 5
'    sqlstr = "UPDATE tblCompetitorPoints SET pointsFinals_4 = " & points
'    'sqlstr = sqlstr & ", points"
'  Case 2
'    sqlstr = "UPDATE tblCompetitorPoints SET pointsFinals_2 = " & points
'    'sqlstr = sqlstr & ", points"
'  Case Else
'    sqlstr = "UPDATE tblCompetitorPoints SET pointsFinals = " & points
'    'sqlstr = sqlstr & ", points"
'  End Select
'  sqlstr = sqlstr & " WHERE competitorPoolID = " & poolForm
'  sqlstr = sqlstr & " AND matchNumber = " & matchNr
'  cn.Execute sqlstr
'  getFinalPoints = points
'End Function

Function getTotalDayPoints(thisForm As Long, beforeMatch As Integer, savdat As Date, cn As ADODB.Connection)
'calculate total daypoints on savdat for poolForm, excluding beforeMatch
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sqlstr As String
'Dim matchOrder As Integer
'matchOrder = getHighestDayMatchNr(savdat, cn)
sqlstr = "Select SUM(pointsDay) as ttl from tblCompetitorPoints where competitorPoolID = " & thisForm
sqlstr = sqlstr & " AND matchOrder <= " & beforeMatch
sqlstr = sqlstr & " AND matchOrder IN (SELECT matchOrder from tblTournamentSchedule "
sqlstr = sqlstr & " where cdbl(matchDate) = " & CDbl(savdat) & ")"
rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
getTotalDayPoints = nz(rs!ttl, 0)
rs.Close
Set rs = Nothing
End Function

Function getPoolFormMatchPoints(poolFormID As Long, matchOrder As Integer, cn As ADODB.Connection)
'calculate poolform score for match
Dim sqlstr As String
Dim savdat As Date
Dim htPts As Integer
Dim ftPts As Integer
Dim totoPts As Integer
Dim dayGoalsPts As Integer
Dim matchNr As Integer
matchNr = getMatchNumber(matchOrder, cn)
'compare the match with the prediction and save the points in variables
  htPts = getMatchPts(poolFormID, matchOrder, 1, cn)
  ftPts = getMatchPts(poolFormID, matchOrder, 2, cn)
  totoPts = getMatchPts(poolFormID, matchOrder, 3, cn)
  dayGoalsPts = 0
  savdat = getMatchInfo(matchOrder, "matchDate", cn)
  
  If IsLastMatchOfDay(matchOrder, cn) Then
    'get the total of goals
    If getPredictionGoalsPerDay(poolFormID, savdat, cn) = goalsPerDay Then
      dayGoalsPts = getPointsFor("doelpunten op een dag", cn)
    Else
      dayGoalsPts = 0
    End If
  End If
  'Update the  competitorPoolPoints table
  sqlstr = "UPDATE tblCompetitorPoints SET"
  sqlstr = sqlstr & " ptsMatch = " & htPts + ftPts + totoPts
  sqlstr = sqlstr & ", ptsHt = " & htPts
  sqlstr = sqlstr & ", ptsFt = " & ftPts
  sqlstr = sqlstr & ", ptsToto = " & totoPts
  sqlstr = sqlstr & ", ptsDayGoals = " & dayGoalsPts
  sqlstr = sqlstr & " WHERE competitorpoolID = " & poolFormID
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  cn.Execute sqlstr
  'return the total amount of points
  getPoolFormMatchPoints = htPts + ftPts + totoPts + dayGoalsPts
  
End Function

Function getGroupStandPoints(poolFormID As Long, matchOrder As Integer, cn As ADODB.Connection)
'get the points for the group standings
Dim rsForm As ADODB.Recordset
Dim rsGrp As ADODB.Recordset
Dim sqlstr As String
Dim pts As Integer
Dim ttlPts As Integer
Dim grp As String
Dim lastGroupPoints As Integer
  
Dim i As Integer
  
  grp = getMatchGroup(matchOrder, cn)
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
  'update table with the points
  sqlstr = "UPDATE tblCompetitorPoints SET"
  sqlstr = sqlstr & " pointsGroupStanding = " & getTotalGroupTtl(poolFormID, matchOrder, 8, 8, cn) + ttlPts
  sqlstr = sqlstr & ",pointsGrp" & grp & " = " & ttlPts
  sqlstr = sqlstr & " WHERE competitorpoolID = " & poolFormID
  sqlstr = sqlstr & " AND matchorder = " & matchOrder
  cn.Execute sqlstr
  
  'return the total amount of points
  getGroupStandPoints = ttlPts
  
  rsForm.Close
  rsGrp.Close
  Set rsForm = Nothing
  Set rsGrp = Nothing
End Function


Function get8FinPts(poolFormID As Long, grp As String, forMatch As Integer, cn As ADODB.Connection)
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
Dim matchNr As Integer  'remember the match number

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
      matchNr = rsMatch!matchOrder
      teamLeft = rsMatch!matchteamA = rs!teamCode 'which team do we have home or away
    Else
      MsgBox "BIIG ERRROORRRR"
      End
    End If
  '4 Zoek in de tblPredictions_Finals naar de teamID in kolom teamnameA of B
    sqlstr = "Select * from tblPrediction_Finals WHERE competitorpoolID = " & poolFormID
    sqlstr = sqlstr & " AND matchorder Between " & matchNrFirst & " AND " & matchNrLast
    sqlstr = sqlstr & " AND (teamNameA = " & teamID & " OR teamNameB = " & teamID & ")"
    rsForm.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  '5 Als gevonden: bepaal wedstrijd nummer en postitie A of B
    If Not rsForm.EOF Then
      grpPts = pts(0)  'in ieder geval de teamnaam punten
      Do While Not rsForm.EOF
      '6. als er meer dan 1 is gevonden zoekdan de beste plek (slechts een keer uitkeren, de hoogste)
        onPosition = matchNr = rsForm!matchOrder
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
      matchNr = rsMatch!matchOrder
      teamLeft = rsMatch!matchteamA = rs!teamCode 'which team do we have home or away
    Else
      MsgBox "BIIG ERRROORRRR"
      End
    End If
  '4 Zoek in de tblPredictions_Finals naar de teamID in kolom teamnameA of B
    sqlstr = "Select * from tblPrediction_Finals WHERE competitorpoolID = " & poolFormID
    sqlstr = sqlstr & " AND matchorder Between " & matchNrFirst & " AND " & matchNrLast
    sqlstr = sqlstr & " AND (teamNameA = " & teamID & " OR teamNameB = " & teamID & ")"
    rsForm.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  '5 Als gevonden: bepaal wedstrijd nummer en postitie A of B
    If Not rsForm.EOF Then
      grpPts = pts(0)  'in ieder geval de teamnaam punten
      Do While Not rsForm.EOF
      '6. als er meer dan 1 is gevonden zoekdan de beste plek (slechts een keer uitkeren, de hoogste)
        onPosition = matchNr = rsForm!matchOrder
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
  get8FinPts = ttlPts
  'update the points table
  sqlstr = "UPDATE tblCompetitorPoints SET pointsTeamsFinals8" & grp & " = " & ttlPts
  sqlstr = sqlstr & ", pointsFinals_8 = " & getTotalGroupTtl(poolFormID, matchNr, 17, 8, cn) + ttlPts
  sqlstr = sqlstr & " WHERE competitorpoolID = " & poolFormID
  sqlstr = sqlstr & " AND matchOrder = " & forMatch
  cn.Execute sqlstr
End Function

Function calcGrandTotal(poolFormID As Long, matchNr As Integer, cn As ADODB.Connection)
'add all the dayTotalpoints for this poolForm up till matchnr
Dim sqlstr As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
  sqlstr = "Select SUM(pointsDay) as ttl from tblCompetitorpoints "
  sqlstr = sqlstr & " WHERE competitorPoolid = " & poolFormID
  sqlstr = sqlstr & " AND matchOrder < " & matchNr
  rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  If Not rs.EOF Then
    calcGrandTotal = nz(rs!ttl, 0)
  Else
    calcGrandTotal = 0
  End If
End Function

Sub updateAllPoolPoints(cn As ADODB.Connection)
'updates all the particpant scores for all matches
'in fact recalculates the entire pool
  Dim match As Integer
  For match = 1 To getMatchCount(0, cn)
    updatePoolpointsForMatch match, cn
  Next
End Sub

Sub updatePoolpointsForMatch(matchOrder As Integer, cn As ADODB.Connection)
'recalculate points for single match
Dim sqlstr As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
  
  sqlstr = "Select * from tblTournamentSchedule WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchPlayed = -1"
  sqlstr = sqlstr & " AND matchorder = " & matchOrder
  
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    'showInfo True, "Punten tellen"
    updatePoolFormPoints rs!matchOrder, cn
    DoEvents
    'if matchnumber is last groupmatch then recalculate 3rd places
    If rs!matchOrder = getFirstFinalMatchNumber(cn) - 1 And getTournamentInfo("tournamentgroupcount", cn) = 6 Then
      update3rdPlacePoints rs!matchOrder, cn
    End If
    updatePoolPositions rs!matchOrder, cn
  End If
  rs.Close
  Set rs = Nothing
End Sub

Sub updatePoolPositions(matchOrder As Integer, cn As ADODB.Connection)
' calculate positions for dayPoints and totalPoints
  Dim pos As Integer
  Dim oldPnt As Integer
  Dim oldPos As Integer
  Dim lastPos As Integer
  
'  Dim pntTopPos As Integer
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim money As Currency
  Dim sqlstr As String
  Dim orderStr As String
  Dim matchNr As Integer
  matchNr = getMatchNumber(matchOrder, cn)
  
  sqlstr = "Select * from tblCompetitorPoints WHERE competitorPoolID IN "
  sqlstr = sqlstr & " (SELECT competitorPoolID from tblCompetitorPools WHERE poolID = " & thisPool & ")"
  sqlstr = sqlstr & " AND matchOrder = " & matchOrder
  
  orderStr = " ORDER BY ptsMatch DESC"
  rs.Open sqlstr & orderStr, cn, adOpenKeyset, adLockOptimistic
  With rs
    oldPnt = 0
    pos = 0
    oldPos = 0
    Do While Not .EOF
      pos = pos + 1
      If oldPnt <> !ptsMatch Then
        oldPos = pos
        oldPnt = !ptsMatch
      End If
      !positionMatches = oldPos
      .Update
      .MoveNext
    Loop
    .Close
  End With
'daily order
  orderStr = " ORDER BY pointsDayTotal DESC"
  rs.Open sqlstr & orderStr, cn, adOpenKeyset, adLockOptimistic
  With rs
    oldPnt = 0
    pos = 0
    oldPos = 0
    Do While Not .EOF
      pos = pos + 1
      If oldPnt <> !pointsDayTotal Then
        oldPos = pos
        oldPnt = !pointsDayTotal
      End If
      !positionDay = oldPos
      .Update
      .MoveNext
    Loop
    .Close
  End With
  'grand total order
  orderStr = " ORDER BY pointsGrandTotal DESC"
  rs.Open sqlstr & orderStr, cn, adOpenKeyset, adLockOptimistic
  With rs
    oldPnt = 0
    pos = 0
    oldPos = 0
    Do While Not .EOF
      pos = pos + 1
      If oldPnt <> !pointsGrandTotal Then
        oldPos = pos
        oldPnt = !pointsGrandTotal
      End If
      !positionTotal = oldPos
      .Update
      .MoveNext
    Loop
    lastPos = oldPos  'save the last positiion for money calculation later
    .Close
  End With
  'calculate money only on last match of the day
  If IsLastMatchOfDay(matchOrder, cn) Then
    'get all the pools with highest day position
    sqlstr = "Select * from tblCompetitorPoints WHERE competitorPoolID IN "
    sqlstr = sqlstr & " (SELECT competitorPoolID from tblCompetitorPools WHERE poolID = " & thisPool & ")"
    sqlstr = sqlstr & " AND matchOrder = " & matchOrder
    orderStr = " AND positionDay = 1"
    rs.Open sqlstr & orderStr, cn, adOpenKeyset, adLockOptimistic
    With rs
      If Not .EOF Then
        .MoveLast
        .MoveFirst
        Do While Not .EOF
          !moneyDay = CCur(getPoolInfo("prizeHighDayScore", cn) / .RecordCount)
          .Update
          .MoveNext
        Loop
      End If
      .Close
    End With
    orderStr = " AND positionTotal = 1"
    rs.Open sqlstr & orderStr, cn, adOpenKeyset, adLockOptimistic
    With rs
      If Not .EOF Then
        .MoveLast
        .MoveFirst
        Do While Not .EOF
          
          !moneydayposition = CCur(getPoolInfo("prizeHighDayPosition", cn) / .RecordCount)
          .Update
          .MoveNext
        Loop
      End If
      .Close
    End With
    orderStr = " AND positionTotal = " & lastPos
    rs.Open sqlstr & orderStr, cn, adOpenKeyset, adLockOptimistic
    With rs
      .MoveLast
      .MoveFirst
      Do While Not .EOF
        !moneyDayLast = CCur(getPoolInfo("prizeLowDayPosition", cn) / .RecordCount)
        .Update
        .MoveNext
      Loop
      .Close
    End With
        'fill some calculated fields just for convenience
    rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    With rs
      Do While Not .EOF
        !moneyDaytotal = !moneyDay + !moneydayposition + !moneyDayLast
        !moneyTotal = getMoneyTotal(!competitorPoolID, getMatchPrevDay(matchNr, cn), cn) + !moneyDaytotal
        .MoveNext
      Loop
      .Close
    End With

  End If
  If matchOrder = getMatchCount(0, cn) Then
  'get the final day money prizes
    updateFinalDayMoney cn
  End If
  
  Set rs = Nothing
End Sub

Sub updateFinalDayMoney(cn As ADODB.Connection)
Dim cnt As Integer
Dim i As Integer
Dim J As Integer
Dim sqlstr As String
Dim posStr As String
Dim prizes(4) As Double
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  prizes(0) = getPrizeMoney(0, cn)
  sqlstr = "Select * from tblCompetitorPoints "
  sqlstr = sqlstr & " WHERE matchorder = " & getMatchCount(0, cn)
  sqlstr = sqlstr & " AND competitorPoolID IN ("
  sqlstr = sqlstr & " SELECT competitorPoolID from tblCompetitorPools WHERE poolID = " & thisPool
  sqlstr = sqlstr & ")"
  For J = 1 To 4
    posStr = " AND positionTotal = " & J
  'first prize
    rs.Open sqlstr & posStr, cn, adOpenKeyset, adLockOptimistic
    cnt = rs.RecordCount
    For i = J To cnt + J - 1
      If i <= 4 Then
        prizes(J) = prizes(J) + getPrizeMoney(i, cn)
      End If
    Next
    prizes(J) = prizes(J) / cnt
    Do While Not rs.EOF
      rs!moneyDaytotal = rs!moneyDaytotal + prizes(J)
      rs!moneyTotal = rs!moneyTotal + prizes(J)
      rs.Update
      rs.MoveNext
    Loop
    rs.Close
  Next
  'last place
  posStr = " AND positionTotal = " & getLastPoolFormPosition(getMatchCount(0, cn), cn)
  rs.Open sqlstr + posStr, cn, adOpenKeyset, adLockOptimistic
  cnt = rs.RecordCount
  Do While Not rs.EOF
    rs!moneyDaytotal = rs!moneyDaytotal + prizes(0) / cnt
    rs!moneyTotal = rs!moneyTotal + prizes(0) / cnt
    rs.Update
    rs.MoveNext
  Loop
  rs.Close

  Set rs = Nothing
End Sub

Sub cleanGroupStandings(cn As ADODB.Connection)
'just (re)calculate all the groupstandings
Dim sqlstr As String
  sqlstr = "Update tblGroupLayout set "
  sqlstr = sqlstr & "mPl = 0, mWon = 0, mLost = 0, mDraw = 0, "
  sqlstr = sqlstr & "mScored = 0, mAgainst = 0, teamPoints = 0"
  sqlstr = sqlstr & " WHERE tournamentID  = " & thisTournament
  cn.Execute sqlstr
End Sub

Sub calcGroupStandings(cn As ADODB.Connection)
'just (re)calculate all the groupstandings
Dim rs As ADODB.Recordset
Dim rsResults As ADODB.Recordset
Dim sqlstr As String

  Set rs = New ADODB.Recordset
  Set rsResults = New ADODB.Recordset
  cleanGroupStandings cn
  sqlstr = "Select * from tblTournamentSchedule where tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchtype = 1 and matchPlayed = -1"
  rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
  If rs.RecordCount <> 0 Then
    
    rs.MoveFirst
    Do While Not rs.EOF
        sqlstr = "Select * from tblMatchResults where tournamentid = " & thisTournament
        sqlstr = sqlstr & " AND matchOrder = " & rs!matchOrder
        
        rsResults.Open sqlstr, cn, adOpenDynamic, adLockOptimistic
        
        If Not rsResults.EOF Then
          rsResults.MoveFirst
          updateGroupLayout rsResults, rsResults!teamA_ID, True, cn
          updateGroupLayout rsResults, rsResults!teamB_ID, False, cn
        End If
        rsResults.Close
        rs.MoveNext
    Loop
    updateGroupPositions cn
  End If
  rs.Close
  Set rs = Nothing
  Set rsResults = Nothing
End Sub

Sub updateGroupLayout(rsResults As ADODB.Recordset, teamID As Long, teamLeft As Boolean, cn As ADODB.Connection)
  'update the groupLayout table (played/won/lost/draw etc) to represent the results from the matches
Dim sqlstr As String
Dim rsGrp As ADODB.Recordset
  Set rsGrp = New ADODB.Recordset
  sqlstr = "Select * from tblGroupLayout WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND teamID = " & teamID
  rsGrp.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  With rsGrp
    !mPl = !mPl + 1
    If !teamID = rsResults!winner Then
      !mWon = !mWon + 1
      !teamPoints = !teamPoints + 3
    ElseIf rsResults!winner = 0 Then
      !mDraw = !mDraw + 1
      !teamPoints = !teamPoints + 1
    Else
      !mLost = !mLost + 1
    End If
    If teamLeft Then
      !mScored = !mScored + rsResults!ftA
      !mAgainst = !mAgainst + rsResults!ftB
    Else
      !mScored = !mScored + rsResults!ftB
      !mAgainst = !mAgainst + rsResults!ftA
    End If
    .Update
  End With
  rsGrp.Close
  Set rsGrp = Nothing
End Sub

Sub updateGroupPositions(cn As ADODB.Connection)
  Dim sqlstr As String
  Dim grp As Integer
  Dim pl As Integer
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
    For grp = 0 To getTournamentInfo("tournamentGroupCount", cn) - 1
      sqlstr = "Select * from tblGroupLayout where tournamentId = " & thisTournament
      sqlstr = sqlstr & " AND groupLetter = '" & Chr(65 + grp) & "'"
      sqlstr = sqlstr & " ORDER BY teamPoints DESC, (mScored-mAgainst) DESC, mScored DESC, groupPlace ASC"
      rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
      If Not rs.EOF Then
        rs.MoveFirst
        pl = 0
        Do While Not rs.EOF
          pl = pl + 1
          rs!teamposition = pl
          If rs!mPl = 0 Then rs!teamposition = 4
          rs.Update
          rs.MoveNext
        Loop
      End If
      rs.Close
    Next
  Set rs = Nothing
End Sub

Sub Set8Finals(cn As ADODB.Connection)
'bepaal de achtste (of kwartfinales als het een 16 teams toernooi is)
Dim fldName As String
Dim rsGrp As New ADODB.Recordset
Dim groupAllPlayed As Boolean
Dim sqlstr As String
    sqlstr = "Select * from tblGroupLayout Where tournamentid=" & thisTournament
    sqlstr = sqlstr & " and teamPosition < 3"
    sqlstr = sqlstr & " ORDER by groupLetter, teamPosition"
    rsGrp.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    'loop through all the groups , position 1 and 2
    Do While Not rsGrp.EOF
      groupAllPlayed = grpPlayedCount(rsGrp!groupletter, cn) = 6
      fldName = "1" & rsGrp!groupletter
      sqlstr = "UPDATE tblTournamentTeamCodes SET teamID = "
      If groupAllPlayed Then
        sqlstr = sqlstr & rsGrp!teamID
      Else
        sqlstr = sqlstr & "0"
      End If
      sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
      sqlstr = sqlstr & " AND teamCode = '" & fldName & "'"
      cn.Execute sqlstr
      rsGrp.MoveNext 'next group posistion (2)
      If Not rsGrp.EOF Then 'just in case
        fldName = "2" & rsGrp!groupletter
        sqlstr = "UPDATE tblTournamentTeamCodes SET teamID = "
        If groupAllPlayed Then
          sqlstr = sqlstr & rsGrp!teamID
        Else
          sqlstr = sqlstr & "0"
        End If
        sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
        sqlstr = sqlstr & " AND teamCode = '" & fldName & "'"
        cn.Execute sqlstr
        rsGrp.MoveNext 'next group and position 1
      End If
    Loop
    rsGrp.Close
    Set rsGrp = Nothing
End Sub

Sub update3rdPlacePoints(matchNr As Integer, cn As ADODB.Connection)
'for tournmnts where a third place can mve to 8th finals
Dim endMatchNr  As Integer
Dim sqlstr As String
Dim pts(1) As Integer
Dim teams(4) As Long
Dim points As Integer
Dim rs As ADODB.Recordset
Dim rsFn As ADODB.Recordset
Dim teamCode As Long
Dim grp As String
Dim i As Integer
Dim finMatch(4) As Integer 'save the matchnumbers where the 3rd places are going to play
  
  '''' GET OUT if match is not played
  If Not getMatchInfo(matchNr, "matchplayed", cn) Then Exit Sub
  ''''''
  
  Set rs = New ADODB.Recordset
  Set rsFn = New ADODB.Recordset
  finMatch(1) = 39
  finMatch(2) = 40
  finMatch(3) = 41
  finMatch(4) = 43
  teams(1) = getTeamIdFromCode("3ADEF", cn)
  teams(2) = getTeamIdFromCode("3DEF", cn)
  teams(3) = getTeamIdFromCode("3ABC", cn)
  teams(4) = getTeamIdFromCode("3ABCD", cn)
  endMatchNr = matchNr + 8
  pts(0) = getPointsFor("8e fin team", cn)
  pts(1) = getPointsFor("8e fin pos", cn)
  'open een recordset met alle deelnemers voor de laatste wedstrijd van de groepsfase
  sqlstr = "Select * from tblCompetitorPoints "
  sqlstr = sqlstr & " WHERE matchNumber = " & matchNr
  sqlstr = sqlstr & " AND competitorPoolId IN "
  sqlstr = sqlstr & " (Select competitorPoolID  from tblCompetitorPools"
  sqlstr = sqlstr & " WHERE poolID = " & thisPool & ")"
  rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  Do While Not rs.EOF
  'OPEN EEN RECORDSET MET DE FINALE TEAMS VAN DEZE DEELNEMER
    sqlstr = "Select * from tblPrediction_finals "
    sqlstr = sqlstr & " WHERE competitorPoolID = " & rs!competitorPoolID
    sqlstr = sqlstr & " AND matchnumber BETWEEN " & matchNr & " AND " & endMatchNr
    rsFn.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rsFn.EOF Then  'just in case
      'zoek de teamcodes voor de derde plaatsen in de recordset
      For i = 1 To 4
        teamCode = teams(i)
        grp = getGroup(teamCode, cn)
        rsFn.Find "teamNameB = " & teamCode
        If Not rsFn.EOF Then
          points = pts(0)
          'check the matchnumber
          If rsFn!matchNumber = finMatch(i) Then
            points = pts(1)
          End If
          'update the points table
          rs("pointsTeamsFinals8" & grp) = rs("pointsTeamsFinals8" & grp) + points
          rs!pointsFinals_8 = rs!pointsFinals_8 + points
          rs!pointsDay = rs!pointsDay + points
          rs!pointsDayTotal = rs!pointsDayTotal + points
          rs!pointsGrandTotal = rs!pointsGrandTotal + points
        End If
      Next
    End If
    rsFn.Close
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  Set rsFn = Nothing

End Sub



'
'Sub updatePoolFormPointsWKPool(matchNr As Integer, cn As ADODB.Connection)
''old system from first wkpool
'
''(re) calculate the points for this poolform for this match
''!!!!!!!!!!!!!  DIRTY PROGRAMMING just copied and adapted FROM old version(WKPool, Sub DeelnemersDoorrekenen)
'
'Dim rsResults As ADODB.Recordset
'Dim rsPoolForm As ADODB.Recordset
'Dim rsPoolFormPoints As ADODB.Recordset
'Dim rsPredicGrp As ADODB.Recordset
'Dim sqlstr As String
''Dim Dagpnt As Integer
'Dim dayPts As Integer
'Dim prevTotal As Integer
'Dim dayGoalsPts As Integer
'Dim thisPoolForm As Long
'Dim savdat As Date
'Dim goalsThisDay  As Integer
'Dim team1 As Long
'Dim team2 As Long
'Dim matchType As Integer
'Dim matchDescr As String
'Dim lastDay As Boolean
''Dim matchNr As Integer
'Dim prevMatchNr As Integer
'Dim group As String
'Dim grpCount As Integer
'Dim grpPnt As Integer
'Dim fin8pnt As Integer
'Dim fin4pnt As Integer
'Dim chkFinals8 As Boolean 'check the 3rd places for the 8th finals (new EK system)
'Dim tournamentFinished As Boolean
'
'  Set rsResults = New ADODB.Recordset
'  Set rsPoolForm = New ADODB.Recordset
'
'  matchType = getMatchInfo(matchNr, "matchType", cn)
'  grpCount = getTournamentInfo("tournamentGroupCount", cn)
'  prevMatchNr = getPrevMatchNr(matchNr, cn)
'  savdat = getMatchInfo(matchNr, "matchdate", cn)
'  lastDay = IsLastMatchOfDay(matchNr, cn)
'  matchDescr = getMatchDescription(matchNr, cn)
'  goalsThisDay = getGoalsPerDay(savdat, cn)
'
'  sqlstr = "SELECT * FROM tblMatchResults "
'  sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
'  sqlstr = sqlstr & " AND matchNumber = " & matchNr
'  rsResults.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'
'  sqlstr = "SELECT competitorPoolID, nickname FROM tblCompetitorPools "
'  sqlstr = sqlstr & " WHERE poolID = " & thisPool
'  sqlstr = sqlstr & " ORDER BY nickName"
'  rsPoolForm.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'
'  Do While Not rsPoolForm.EOF
'    Set rsPredicGrp = New ADODB.Recordset
'    thisPoolForm = rsPoolForm!competitorPoolID
'    dayGoalsPts = 0
'    team1 = rsResults!teamA_ID
'    team2 = rsResults!teamB_ID
'
'    sqlstr = "SELECT * FROM tblPredictionGroupResults"
'    sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm
'    sqlstr = sqlstr & " ORDER BY competitorPoolID, groupletter"
'
'    'open recordset for editing
'    rsPredicGrp.Open sqlstr, cn, adOpenStatic, adLockOptimistic
'
'    sqlstr = "SELECT * FROM tblCompetitorPoints"
'    sqlstr = sqlstr & " WHERE competitorPoolID = " & thisPoolForm & " AND matchNumber = " & matchNr
'    'sqlstr = sqlstr & " order by matchNumber "
'    Set rsPoolFormPoints = New ADODB.Recordset
'    rsPoolFormPoints.Open sqlstr, cn, adOpenDynamic, adLockOptimistic
'    Dim nameStr As String
'    nameStr = rsPoolForm!nickName & "  (" & rsPoolForm.AbsolutePosition & "/" & rsPoolForm.RecordCount & ")"
'    showInfo True, "Punten tellen", "Na wedstrijd " & matchDescr, nameStr
'    With rsPoolFormPoints
'      If Not rsPoolFormPoints.EOF Then
'        .Delete 'voor de zekerheid eerst even wissen
'      End If
'
'      .AddNew
'      !competitorPoolID = thisPoolForm
'      !matchNumber = matchNr
''      .Update
''      .Close
''      .Open sqlstr, cn, adOpenKeyset, adLockOptimistic
'    'is dit de laatste wedstrijd van de dag
'      If lastDay Then 'get the goals-of-day points
'        If getPredictionDayGoals(!competitorPoolID, savdat, cn) = goalsThisDay Then
'          dayGoalsPts = getPointsForID(28, cn)
'        Else
'          dayGoalsPts = 0
'        End If
'      End If
'      !ptsHt = getMatchPts(thisPoolForm, matchNr, 1, cn) 'voor het punten-opbouw rapport ook nog gesplitst
'      !ptsFt = getMatchPts(thisPoolForm, matchNr, 2, cn)
'      !ptsToto = getMatchPts(thisPoolForm, matchNr, 3, cn)
'      !ptsMatch = !ptsHt + !ptsFt + !ptsToto
'      !ptsDayGoals = dayGoalsPts
'      If matchType = 1 Then 'check finale wedstrijden en posities
'        group = getGroup(team1, cn)
'        If grpPlayedCount(group, cn) = 6 And getLastGroupMatch(group, cn) = matchNr Then
'          .Fields("pointsGrp" & group) = getGroupPoints(thisPoolForm, group, cn)
'          If grpCount > 4 Then
'            .Fields("pointsTeamsFinals8" & group) = getPoolFormFinal8Points(thisPoolForm, group, cn)
'          Else
'            .Fields("pointsTeamsFinals4" & group) = 0
'            MsgBox "Kwartfinale berekening nog niet gedaan in code als er slechts 4 groepen zijn", vbAbortRetryIgnore + vbOKOnly, "NIET AF"
'          End If
'          If grpCount = 6 And matchNr = getMatchCount(1, cn) Then
'            'new ek system
'            chkFinals8 = True
'          End If
'        Else
'          If grpCount > 4 Then
'            .Fields("pointsGrp" & group) = 0
'          Else
'            .Fields("pointsTeamsFinals4" & group) = 0
'          End If
'        End If
'        !pointsGroupStanding = getPoolFormPoints(thisPoolForm, prevMatchNr, 7, cn) + nz(.Fields("pointsGrp" & group), 0)
'        If grpCount > 4 Then
'          !pointsFinals_8 = getPoolFormPoints(thisPoolForm, prevMatchNr, 16, cn) + nz(.Fields("pointsTeamsFinals8" & group), 0)
'          fin8pnt = nz(.Fields("pointsTeamsFinals8" & group), 0)
'        Else
'          !pointsFinals_4 = getPoolFormPoints(thisPoolForm, prevMatchNr, 25, cn) + nz(.Fields("pointsTeamsFinals4" & group), 0)
'          fin4pnt = nz(.Fields("pointsTeamsFinals4" & group), 0)
'        End If
'        grpPnt = nz(.Fields("pointsGrp" & group), 0)
'
'      Else
'        !pointsGroupStanding = getPoolFormPoints(thisPoolForm, prevMatchNr, 7, cn)
''        For i = 1 To 8
''          .Fields("pointsGrp" & Chr(64 + i)) = getPoolFormPoints(thisPoolForm, prevMatchNr, 7 + i, cn)
''        Next
'        !pointsFinals_8 = getPoolFormPoints(thisPoolForm, prevMatchNr, 16, cn)
'        !pointsFinals_4 = getPoolFormPoints(thisPoolForm, prevMatchNr, 25, cn)
'        !pointsFinals_2 = getPoolFormPoints(thisPoolForm, prevMatchNr, 30, cn)
'        !pointsFinals_34 = getPoolFormPoints(thisPoolForm, prevMatchNr, 33, cn)
'        !pointsFinal = getPoolFormPoints(thisPoolForm, prevMatchNr, 36, cn)
'        !pointsTotalAfterFinal34 = 0
'        grpPnt = 0
'        fin8pnt = 0
'      End If
'      If matchType = 5 Then 'AchtsteFinale
'        '
'        If grpCount > 6 Then
'          Select Case matchNr
'          Case 49, 50
'              group = "B"
'          Case 51, 52
'              group = "C"
'          Case 53, 54
'              group = "A"
'          Case 55, 56
'              group = "D"
'          End Select
'        Else
'          Select Case matchNr
'          Case 37, 39
'              group = "B"
'          Case 38, 40
'              group = "C"
'          Case 41, 42
'              group = "A"
'          Case 43, 44
'              group = "D"
'          End Select
'        End If
'
'        .Fields("pointsTeamsFinals4" & group) = nz(getPoolFormFinal4Points(thisPoolForm, matchNr, cn), 0)
'        !pointsFinals_4 = getPoolFormPoints(thisPoolForm, prevMatchNr, 25, cn) + .Fields("pointsTeamsFinals4" & group)
'        '.Fields("pntfin4" & group) = getPoolFormPoints(prevwednum, thisPoolForm, 9, "4" & group) + .Fields("pntfin4" & group)
'
'        fin8pnt = .Fields("pointsTeamsFinals4" & group)
'      End If
'      If matchType = 2 Then  'kwartfinale
'        If grpCount = 8 Then
'          Select Case rsResults!matchNumber
'          Case 57, 58
'              group = "A"
'          Case 59, 60
'              group = "B"
'          End Select
'          .Fields("pointsTeamsFinals2" & group) = nz(getPoolFormFinal2Points(thisPoolForm, matchNr, cn), 0)
'          !pointsFinals_2 = getPoolFormPoints(thisPoolForm, prevMatchNr, 30, cn) + .Fields("pointsTeamsFinals2" & group)
'          fin8pnt = .Fields("pointsTeamsFinals2" & group)
'        ElseIf grpCount = 6 Then
'          Select Case rsResults!matchNumber
'          Case 45, 46
'              group = "A"
'          Case 47, 48
'              group = "B"
'          End Select
'          .Fields("pointsTeamsFinals2" & group) = nz(getPoolFormFinal2Points(thisPoolForm, matchNr, cn), 0)
'          !pointsFinals_2 = getPoolFormPoints(thisPoolForm, prevMatchNr, 30, cn) + .Fields("pointsTeamsFinals2" & group)
'          fin8pnt = .Fields("pointsTeamsFinals2" & group)
'        Else
'          If rsResults!wedNum = 25 Or rsResults!wedNum = 26 Then
'              group = "A"
'          Else
'              group = "B"
'          End If
'          .Fields("pointsTeamsFinals2" & group) = nz(getPoolFormFinal2Points(thisPoolForm, matchNr, cn), 0)
'          !pointsFinals_2 = getPoolFormPoints(thisPoolForm, prevMatchNr, 30, cn) + .Fields("pointsTeamsFinals2" & group)
'          fin4pnt = .Fields("pointsTeamsFinals2" & group)
'        End If
'      End If
'      If matchType = 3 Then 'halve finale
'        If getTournamentInfo("tournamentThirdPlace", cn) Then
''              !pntklfindag = getKLfinpnt(rsResults!wedNum, thisPoolForm)
''              !pntklfin = getPoolFormPoints(prevMatchNr, thisPoolForm, 6) + getKLfinpnt(rsResults!wedNum, thisPoolForm)
'          MsgBox "Kleine finale berekening nog niet gedaan in code als die er is", vbAbortRetryIgnore + vbOKOnly, "NIET AF"
'        End If
'        !pointsTeamsFinal = getPoolFormFinalPoints(thisPoolForm, matchNr, cn)
'        !pointFinal = getPoolFormPoints(thisPoolForm, prevMatchNr, 36, cn) + !pointsTeamsFinal
'      End If
'      If matchType = 7 Then 'kleine finale
'        MsgBox "Kleine finale berekening nog niet gedaan in code als die er is", vbAbortRetryIgnore + vbOKOnly, "NIET AF"
''              !pntuitslnaklfin = getEindStandpnt(thisPoolForm, 3) + getEindStandpnt(thisPoolForm, 4)
''              fin8pnt = !pntuitslnaklfin
'      End If
'      If matchType = 4 Then 'finale
'        !pointsTotalAfterFinal34 = 0 ' getEindStandpnt(thisPoolForm, 3) + getEindStandpnt(thisPoolForm, 4)
'        !pointsFinalStanding = getTournamentStandingPoints(thisPoolForm, cn)  'getEindStandpnt(thisPoolForm, 1) + getEindStandpnt(thisPoolForm, 2)
'        fin8pnt = !ptsMatch
'        !pointsTopscorers = getPoolFormTopScoresPoints(thisPoolForm, cn)
'        !pointsOther = getStatsPoints(thisPoolForm, cn)
'        tournamentFinished = True
'      Else
'        tournamentFinished = False
'      End If
'      prevTotal = getPoolFormPoints(thisPoolForm, prevMatchNr, 43, cn)
'      !pointsDay = !ptsMatch + grpPnt + dayGoalsPts
'      If grpCount > 4 Then
'        !pointsDay = !pointsDay + fin8pnt
'      Else
'        !pointsDay = !pointsDay + fin4pnt
'        If tournamentFinished Then !pointsDay = !pointsDay + fin8pnt
'      End If
'      !pointsDay = !pointsDay + nz(!pointsFinals_34, 0) + nz(!pointsFinal, 0)
'      !pointsDay = !pointsDay + nz(!pointsTopscorers, 0) + nz(!pointsOther, 0)
'      !pointsgrandTotal = prevTotal + !pointsDay
'      dayPts = getTotalThisDayPoints(thisPoolForm, matchNr, cn)
'      !pointsDayTotal = dayPts
'      !moneytotal = getMoneyTotal(!competitorPoolID, prevMatchNr, cn)
'      .Update
'      If chkFinals8 Then
'        'search also for 3rd place teams
'        add3rdPlacePoints thisPoolForm, matchNr, cn
'      End If
'    End With
'    rsPredicGrp.Close
'    Set rsPredicGrp = Nothing
'
'    rsPoolFormPoints.Close
'    Set rsPoolFormPoints = Nothing
'    rsPoolForm.MoveNext
'  Loop
'
'rsResults.Close
'Set rsResults = Nothing
'
'rsPoolForm.Close
'Set rsPoolForm = Nothing
'
'If matchNr Then
'    updatePoolPositions matchNr, cn
'End If
'showInfo False
'If tournamentFinished Then
''    MaakEindafrekening
'End If
'
'
'
'End Sub

