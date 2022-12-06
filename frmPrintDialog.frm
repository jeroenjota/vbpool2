VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrintDialog 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Afdrukken"
   ClientHeight    =   7995
   ClientLeft      =   11790
   ClientTop       =   6075
   ClientWidth     =   10485
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   450
   Icon            =   "frmPrintDialog.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7995
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picCompetitorList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   6600
      ScaleHeight     =   2505
      ScaleWidth      =   3315
      TabIndex        =   34
      Top             =   1920
      Width           =   3345
      Begin VB.ListBox lstCompetitorPools 
         Height          =   1980
         Left            =   80
         MultiSelect     =   1  'Simple
         TabIndex        =   37
         Top             =   80
         Width           =   3120
      End
      Begin VB.OptionButton optAll 
         Caption         =   "Allemaal"
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   240
         TabIndex        =   36
         Top             =   2150
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Selectie"
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   1635
         TabIndex        =   35
         Top             =   2150
         Width           =   1230
      End
   End
   Begin VB.PictureBox picVolgorde 
      Appearance      =   0  'Flat
      FillColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3120
      ScaleHeight     =   405
      ScaleWidth      =   3435
      TabIndex        =   31
      Top             =   90
      Width           =   3465
      Begin VB.CheckBox chkCombi 
         Caption         =   "combi"
         Height          =   255
         Left            =   2520
         TabIndex        =   47
         Top             =   68
         Width           =   975
      End
      Begin VB.OptionButton poolFormOrder 
         Appearance      =   0  'Flat
         Caption         =   "Op score"
         ForeColor       =   &H00004000&
         Height          =   390
         Index           =   1
         Left            =   1320
         TabIndex        =   33
         Top             =   0
         Width           =   1080
      End
      Begin VB.OptionButton poolFormOrder 
         Appearance      =   0  'Flat
         Caption         =   "Alfabetisch"
         ForeColor       =   &H00004000&
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.PictureBox picPrnterSettings 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   3120
      ScaleHeight     =   1575
      ScaleWidth      =   3435
      TabIndex        =   14
      Top             =   2880
      Width           =   3465
      Begin MSComCtl2.UpDown upDnCopies 
         Height          =   375
         Left            =   3000
         TabIndex        =   41
         Top             =   1125
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtCopies"
         BuddyDispid     =   196618
         OrigLeft        =   2760
         OrigTop         =   1200
         OrigRight       =   3015
         OrigBottom      =   1575
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkNwePagKop 
         Alignment       =   1  'Right Justify
         Caption         =   "Nwe pag kop"
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   1560
         TabIndex        =   30
         ToolTipText     =   "Print wel/niet de kopregels op een nieuwe pagina"
         Top             =   375
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.TextBox txtCopies 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Text            =   "1"
         Top             =   1125
         Width           =   480
      End
      Begin VB.ComboBox cmbPrinters 
         Height          =   360
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Tag             =   "printer"
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton optLandscape 
         Caption         =   "Liggend"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.OptionButton optPortrait 
         Caption         =   "Staand"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CheckBox chkDblSide 
         Alignment       =   1  'Right Justify
         Caption         =   "Dubbelzijdig"
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1155
         Width           =   1425
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Afdruk opties"
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aantal"
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   1845
         TabIndex        =   20
         Top             =   1192
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   435
         Width           =   570
      End
   End
   Begin VB.PictureBox Picture4 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   10425
      TabIndex        =   12
      Top             =   7005
      Width           =   10485
      Begin VB.CommandButton btnStickers 
         Caption         =   "Stickers"
         Height          =   860
         Left            =   2760
         TabIndex        =   46
         Top             =   60
         Width           =   1245
      End
      Begin VB.CommandButton btnFinalPlayerPrint 
         Caption         =   "Eindstand deelnemers"
         Height          =   860
         Left            =   1440
         TabIndex        =   45
         Top             =   60
         Width           =   1245
      End
      Begin VB.CommandButton btnClose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sluiten"
         Height          =   860
         Left            =   5280
         TabIndex        =   29
         Tag             =   "SluitPrintDial"
         Top             =   60
         Width           =   1125
      End
      Begin VB.CommandButton btnPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Voorbeeld"
         Height          =   420
         Index           =   1
         Left            =   4080
         TabIndex        =   28
         ToolTipText     =   "Bekijk een voorbeeld op het scherm"
         Top             =   60
         Width           =   1125
      End
      Begin VB.CommandButton btnPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Printer"
         Height          =   420
         Index           =   0
         Left            =   4080
         TabIndex        =   27
         ToolTipText     =   "Stuur dit rapport naar de printer"
         Top             =   520
         Width           =   1125
      End
      Begin VB.CommandButton btnPrntAllAfterDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alles afdrukken"
         Height          =   860
         Left            =   75
         TabIndex        =   13
         ToolTipText     =   "Druk alles af na de laatste wedstrijd van de dag"
         Top             =   60
         Width           =   1245
      End
      Begin VB.CheckBox chkEindstand 
         Appearance      =   0  'Flat
         Caption         =   "Eind stand"
         ForeColor       =   &H00004000&
         Height          =   615
         Left            =   720
         TabIndex        =   22
         Tag             =   "chkEinstand"
         Top             =   120
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.PictureBox picVoorWed 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3120
      ScaleHeight     =   450
      ScaleWidth      =   3435
      TabIndex        =   9
      Top             =   1080
      Width           =   3465
      Begin MSComCtl2.UpDown upDnForMatch 
         Height          =   375
         Left            =   1290
         TabIndex        =   39
         Top             =   15
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtForMatch"
         BuddyDispid     =   196634
         OrigLeft        =   2520
         OrigRight       =   2775
         OrigBottom      =   375
         Max             =   80
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtForMatch 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Text            =   "1"
         Top             =   15
         Width           =   450
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   "e wedstrijd"
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   1560
         TabIndex        =   43
         Top             =   80
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Voor de"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   75
         TabIndex        =   11
         Top             =   80
         Width           =   780
      End
   End
   Begin VB.PictureBox picToMatch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3120
      ScaleHeight     =   450
      ScaleWidth      =   3435
      TabIndex        =   6
      Top             =   600
      Width           =   3465
      Begin VB.ComboBox cmbMatchesPlayed 
         Height          =   360
         Left            =   615
         TabIndex        =   44
         Top             =   57
         Width           =   2640
      End
      Begin MSComCtl2.UpDown upDnToMatch 
         Height          =   375
         Left            =   450
         TabIndex        =   38
         Top             =   30
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtToMatch"
         BuddyDispid     =   196639
         OrigLeft        =   2520
         OrigTop         =   30
         OrigRight       =   2775
         OrigBottom      =   405
         Max             =   64
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtToMatch 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Text            =   "1"
         Top             =   30
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "T/m"
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   75
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   75
      ScaleHeight     =   4365
      ScaleWidth      =   2850
      TabIndex        =   0
      Tag             =   "afdruk"
      Top             =   90
      Width           =   2880
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Plaats per dag"
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   10
         Left            =   90
         TabIndex        =   48
         ToolTipText     =   "Druk per deelnemer per dag de positie in de pool af"
         Top             =   3570
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Dagresultaat"
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   9
         Left            =   90
         TabIndex        =   42
         Top             =   1645
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Punten samenstelling"
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   26
         Top             =   2415
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Voorspellingen"
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   25
         Top             =   1260
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Punten per wedstrijd"
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   6
         Left            =   90
         TabIndex        =   24
         Top             =   2800
         Value           =   -1  'True
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Stand in de Pool"
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   23
         Top             =   2030
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Inschrijffomulieren"
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   105
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Ingevulde Pools"
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   4
         Top             =   490
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Stand in toernooi"
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   8
         Left            =   90
         TabIndex        =   3
         Top             =   3960
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Grafiek pool stand"
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   7
         Left            =   90
         TabIndex        =   2
         Top             =   3185
         Width           =   2670
      End
      Begin VB.OptionButton optPrintDoc 
         Appearance      =   0  'Flat
         Caption         =   "Favorieten"
         ForeColor       =   &H00004000&
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   1
         Top             =   875
         Width           =   2670
      End
   End
   Begin MSComDlg.CommonDialog printerDialog 
      Left            =   2760
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Tahoma"
   End
End
Attribute VB_Name = "frmPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'added function to print color on dark background
Private Declare Function SetBkMode Lib "gdi32" _
 (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function GetBkMode Lib "gdi32" _
 (ByVal hdc As Long) As Long

'constants for SetBkMode
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private iBKMode As Long

'global objects
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

'add rotator class to print in different angles
Dim rotater As rotator

'the printPreview form
Dim printPrev As frmPrintPreview

'font definitions
Const headingFont = "Kristen ITC"
Const textFont = "Calibri"
Const smallFont = textFont ' or maybe "Segoe UI"
Const headBGcolor = &HE8244
Const headTxtColor = vbWhite
Const textColor = vbBlack
Const lineColor = 0

'gobals for every print
Dim heading1 As String 'top of the section
Dim toMatch As Integer  'to store the matchOrder number till where we should print
Dim currentMatch As Integer 'the currentMatch ordernumber

'default sizes
Dim lineHeight8 As Integer
Dim lineHeight10 As Integer
Dim lineHeight12 As Integer
Dim lineHeight18 As Integer
Dim thisLineHeight As Integer
'
''remember headerheight anf footerStartPos globally
Dim headerHeight As Integer
Dim footerPos As Integer
'
'OLD STUFF
'Dim KolHeight As Integer
'Dim kolwidth As Integer
'Dim kol As Integer
''voor de printFavourites afdruk
'Dim favYpos As Integer
'Dim favXpos As Integer
'
'Dim x As Integer
'
Dim RegHeight As Integer
Dim printObj As Object

Dim maxFavYpos As Integer 'voor afdrukken van printFavourites

Dim thisColor(64) As Long 'voor grafiek

Private Sub btnStickers_Click()
'open word template with sticker layout
Dim wdApp As Object
Dim wDoc As Object 'Word.Document
Dim etikNaam As String
Dim adjust As Integer
Dim dbNaam As String
Dim mrgConn As String
Dim msg As String
Dim etikAant As Integer
Dim sqlstr As String
Dim sql2str As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
'set the MS Word environment
Me.Hide
  showInfo True, "Stickers aanmaken in MS Word ...", "Even geduld ..."
  etikNaam = "etiketten3x8.docx"
  dbNaam = App.Path & "\" & dbName & ".mdb"
  bindApp wdApp, "Word"
  'get the sql records
  sqlstr = "SELECT * FROM qryStickers"
  sqlstr = sqlstr & " WHERE t.tournamentID=" & thisTournament
  sqlstr = sqlstr & " AND cp.matchOrder=" & getMatchCount(0, cn)

  'set the connection to the data
  mrgConn = "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & dbNaam
  mrgConn = mrgConn & ";Mode=Read;Jet OLEDB:System database="""""
  mrgConn = mrgConn & ";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet"
  'get the amount of records
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  etikAant = rs.RecordCount
  rs.Close
  'open the doc
  Set wDoc = wdApp.Documents.Open(App.Path & "\" & etikNaam)
  'copy the cells of the table
  CopyToAllCells wDoc, etikAant
  With wDoc
    .Activate
    .MailMerge.OpenDataSource Name:=dbNaam, ConfirmConversions:=False, _
    ReadOnly:=False, LinkToSource:=True, AddToRecentFiles:=False, _
    PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
    WritePasswordTemplate:="", Revert:=False, Format:=0, _
    Connection:=mrgConn, SQLStatement:=sqlstr, SQLStatement1:=sql2str, _
    SubType:=1
    .MailMerge.SuppressBlankLines = True
    .MailMerge.MainDocumentType = 1
    .MailMerge.Execute
  End With
  wDoc.Close False
  wdApp.Visible = True
  'wdApp.Windows(wdApp.Windows.Count).Activate
  AppActivate wdApp.Windows(wdApp.Windows.Count)
  showInfo False
  'MsgBox "Stickers zijn aangemaakt. Je kunt ze vanuit Word afdrukken", vbOKOnly + vbInformation, "Stickers"
  'Me.Show

End Sub

Sub CopyToAllCells(wd As Object, aant As Integer)
Dim atable As Object
Dim i As Long
Dim J As Long
Dim Source As Object
Dim target As Object
Dim myrange As Object
Dim etikAant As Integer
Dim telr As Integer
    Set atable = wd.tables(1)
'get content of first cell
    Set Source = atable.Cell(1, 1)
    Set myrange = Source.Range
    myrange.Collapse 1
'add NEXT field
    wd.Fields.Add Range:=myrange, Text:="NEXT", PreserveFormatting:=False
'copy first cell
    Source.Range.Copy
    
    etikAant = aant
    For i = 1 To atable.rows.Count
        For J = 1 To atable.Columns.Count
          telr = telr + 1
'if no more records
            If telr > etikAant Then Exit For
            Set target = atable.Cell(i, J)
'if target cell has a field (typically this is a NEXT field)
            If target.Range.Fields.Count > 0 Then
'paste content of first cell (including NEXT field)
                target.Range.Paste
            End If
        Next J
    Next i
'remove NEXT feld from first cell
    atable.Cell(1, 1).Range.Fields(1).Delete
    
End Sub


Private Sub cmbMatchesPlayed_Click()
If Me.cmbMatchesPlayed.ListIndex = -1 Then Exit Sub
  toMatch = Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'tidy up
If Not rs Is Nothing Then
  If (rs.State And adStateOpen) = adStateOpen Then rs.Close
  Set rs = Nothing
End If
If Not cn Is Nothing Then
  If (cn.State And adStateOpen) = adStateOpen Then cn.Close
  Set cn = Nothing
End If
If Not rotater Is Nothing Then
  Set rs = Nothing
End If
If Not printPrev Is Nothing Then
  Unload printPrev
  Set printPrev = Nothing
End If

End Sub

Private Sub optPrintDoc_Click(Index As Integer)
Dim i As Integer
Dim sqlstr As String

Me.picCompetitorList.Visible = False
Me.optPrintDoc(Index).value = True


Select Case Index
  Case 0 'inschrijf formulieren
    Me.picToMatch.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.picCompetitorList.Visible = False
   ' Me.chkDblSide.Value = 1

  Case 1
   'deelnemers met voorspellingen
    Me.picToMatch.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    sqlstr = "Select competitorpoolID, nickName from tblCompetitorPools where poolid=" & thisPool
    sqlstr = sqlstr & " order by nickName"
    fillList Me.lstCompetitorPools, sqlstr, cn, "nickName", "competitorpoolID"
    Me.picCompetitorList.Visible = True
    Me.optPortrait.value = True
  Case 2
    ' printFavourites
    Me.picToMatch.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.picCompetitorList.Visible = False
  Case 3
    'voorspelling per wedstrijd
    picVolgorde.Visible = False
    picVoorWed.Visible = True
    picToMatch.Visible = False
    Me.optPortrait.value = True
    Me.optLandscape.value = False
    Me.upDnForMatch = getLastMatchPlayed(cn) + 1
    Me.picCompetitorList.Visible = False
  Case 4
    'score/ poolstandings / combi
    picVolgorde.Visible = True 'GetDeelnemAant(thisPool) > 32
'    Me.poolFormOrder(2).Visible = getpoolFormCount(cn) <= 62
    Me.chkCombi.Enabled = getpoolFormCount(cn) <= 58
    picVoorWed.Visible = False
    Me.optPortrait.value = True
    Me.picCompetitorList.Visible = False
    picToMatch.Visible = True
    If toMatch > 0 Then
      'Me.upDnToMatch = getLastMatchPlayed(cn)
      setCombo Me.cmbMatchesPlayed, getLastMatchPlayed(cn)

      'Me.upDnToMatch.SetFocus
      toMatch = Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
    End If
    
 Case 5
    'points overview
    Me.picVolgorde.Visible = False
    Me.picVoorWed.Visible = False
    Me.picToMatch.Visible = True
    Me.optLandscape.value = True
    Me.poolFormOrder(0) = True
    toMatch = Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
    DoEvents
    
  Case 6
    'punten per wedstrijd
    picVolgorde.Visible = True
    Me.chkCombi.Enabled = False
    Me.chkCombi = 0
    picVoorWed.Visible = False
    picToMatch.Visible = True
    toMatch = Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
    DoEvents
    Me.picCompetitorList.Visible = False  'getTournamentInfo("groepen")
    Me.optLandscape.value = getTournamentInfo("tournamentGroupCount", cn) > 4
    Me.optPortrait.value = Not Me.optLandscape.value
    
  Case 7
    'grafiek
    Me.picVolgorde.Visible = False
    Me.picVoorWed.Visible = False
    Me.picToMatch.Visible = True
    Me.optLandscape = True
    Me.poolFormOrder(1) = True
    'Me.vscrlTM = GetMyNum(getLastMatchPlayed(cn)())
    DoEvents
    'Me.upDnToMatch = getLastMatchPlayed(cn)
    toMatch = Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)

  Case 8
    'samenvatting stand
    'Stand in toernooi
    Me.picToMatch.Visible = False
    Me.picVoorWed.Visible = False
    Me.picVolgorde.Visible = False
    Me.optPortrait.value = True
    Me.picCompetitorList.Visible = False
    toMatch = Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
    
  Case 9
    'daily results per match NEW (19-6-2021
    Me.picVolgorde.Visible = False
    Me.picVoorWed.Visible = False
    Me.picToMatch.Visible = True
    Me.optPortrait.value = True
    Me.poolFormOrder(0) = True
    toMatch = Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
    DoEvents
    
  Case 10
    'place per day after match  NEW Dec 2022
    Me.picVolgorde.Visible = False
    Me.chkCombi.Enabled = False
    picVoorWed.Visible = False
    Me.optPortrait.value = True
    Me.picCompetitorList.Visible = False
    picToMatch.Visible = True
    If toMatch > 0 Then
      'Me.upDnToMatch = getLastMatchPlayed(cn)
      setCombo Me.cmbMatchesPlayed, getLastMatchPlayed(cn)

      'Me.upDnToMatch.SetFocus
      toMatch = Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
    End If
    
  End Select
  
End Sub

Sub horline(kleur As Integer)
    printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY), kleur
End Sub

Sub subHeader(txt As String)
Dim savFontSize As Integer
Dim savFontColor As Long
Dim savBold As Boolean, savItalic As Boolean
    savFontSize = printObj.FontSize
    savFontColor = printObj.ForeColor
    savBold = printObj.FontBold
    savItalic = printObj.FontItalic
    
    fontSizing 16
    printObj.FontBold = True
    printObj.ForeColor = headBGcolor
    printObj.Print txt
    
    fontSizing savFontSize
    printObj.ForeColor = savFontColor
    printObj.FontBold = savBold
    printObj.FontItalic = savItalic
End Sub

Sub subHeading(txt As String, Optional large As Boolean)
  'heading
  Dim fontSz As Integer
  fontSz = printObj.FontSize
  printObj.FontBold = True
  If large Then
    fontSizing fontSz + 4
  End If
  printObj.ForeColor = vbBlue
  
  printObj.Print txt
  fontSizing fontSz
  printObj.ForeColor = 1
  printObj.FontBold = False

End Sub

Sub printPoolForm()
Dim xPos As Integer, yPos As Integer

    printObj.FillStyle = vbFSTransparent
    
    heading1 = "Inschrijfformulier     inleg: " & Format(getPoolInfo("poolCost", cn), "currency")
    
    InitPage True, True
    
    xPos = printObj.CurrentX
    yPos = printObj.CurrentY
    printPoolFormInstructions xPos, yPos
    
    xPos = printObj.CurrentX
    yPos = printObj.CurrentY + 50
    printPoolFormGroupSection xPos, yPos
    
    xPos = printObj.CurrentX
    yPos = printObj.CurrentY - 50
    printPoolFormFinalSection xPos, yPos
  
    xPos = 0
    yPos = printObj.CurrentY + 100
    printPoolFormBottomBlock xPos, yPos
    
    'print 2nd page with all the matches
    heading1 = "Wedstrijdvoorspellingen"
    addNewPage True, True
    printPoolFormMatches
    
    'bottom line with final deliverydate
    printPoolFormDeliverDate
    'InvulFormAfdrukken
End Sub

Sub printPoolFormInstructions(xPos As Integer, yPos As Integer)
Dim txt As String
Dim i As Integer
Dim aant As Integer
Dim amount As Integer
Dim topY As Integer
Dim lineYpos As Integer
Dim lineXpos As Integer
Dim cols(3) As Integer
    cols(0) = xPos + 100
    cols(1) = cols(0) + printObj.TextWidth("Naam: ")
    cols(2) = printObj.ScaleWidth / 5 * 3 - 40
    cols(3) = printObj.ScaleWidth / 5 * 3
    topY = yPos
    
    printObj.ForeColor = vbBlack
    printObj.FontBold = False
    printObj.CurrentY = topY
    fontSizing 18
    printObj.Line (0, topY - 50)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft - 20, topY + printObj.TextHeight("WW") * 2 + 250), lineColor, B
    printObj.Print
    
    'dotted line
    printObj.DrawStyle = vbDash
    printObj.DrawWidth = 1
    
    printObj.CurrentY = topY
    lineYpos = printObj.CurrentY + printObj.TextHeight("Naam") - 20
    printObj.CurrentX = cols(0)
    printObj.Print "Naam: ";
    lineXpos = cols(0) + printObj.TextWidth("Naam: ")
    printObj.Line (lineXpos, lineYpos)-(cols(2), lineYpos)
    printObj.CurrentY = topY
    printObj.CurrentX = cols(3)
    printObj.Print "Telefoon:";
    lineXpos = cols(3) + printObj.TextWidth("Telefoon: ")
    printObj.Line (lineXpos, lineYpos)-(printObj.ScaleWidth - 40, lineYpos)
    printObj.CurrentY = topY + printObj.TextHeight("WW") + 200
    printObj.CurrentX = cols(0)
    lineYpos = printObj.CurrentY + printObj.TextHeight("Naam") - 20
    printObj.Print "Email: ";
    lineXpos = xPos + printObj.TextWidth("Email: ")
    printObj.Line (lineXpos, lineYpos)-(cols(3), lineYpos)
    printObj.CurrentY = topY + printObj.TextHeight("WW") + 200
    printObj.Print "Betaald ";
    lineXpos = cols(3) + printObj.TextWidth("Betaald ")
    xPos = printObj.CurrentX
    yPos = printObj.CurrentY
    printObj.DrawWidth = 3
    printObj.Line (xPos, yPos)-(xPos + printObj.TextWidth("W"), yPos + printObj.TextHeight("W")), lineColor, B
    lineXpos = lineXpos + printObj.TextWidth("W")
    printObj.DrawWidth = 1
    printObj.CurrentY = yPos
    printObj.CurrentX = printObj.CurrentX + 30
    printObj.Print " bij:";
    lineYpos = printObj.CurrentY + printObj.TextHeight("Naam") - 20
    lineXpos = lineXpos + printObj.TextWidth("bij: ")
    printObj.Line (lineXpos, lineYpos)-(printObj.ScaleWidth - 40, lineYpos)
    printObj.DrawStyle = vbSolid
    fontSizing 4
    printObj.Print
    'sqlstr = "Select * from poolpnt Where thisPool = " & thisPool
    'rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    fontSizing 11
    subHeader "Instructies"
    printObj.Print "Op dit formulier kun je jouw voorspellingen invoeren voor het "; getTournamentInfo("description", cn);
    printObj.Print " van "; Format(getTournamentInfo("tournamentstartdate", cn), "d MMMM"); " tot "; Format(getTournamentInfo("tournamentEnddate", cn), "d MMMM yyyy")
    printObj.Print "Voor elke juiste voorspelling krijg je punten, bij de verschillende onderdelen staat hoeveel."
    printObj.Print "De voorspellingen hoeven onderling niet te kloppen, je kunt bijvoorbeeld een team dat bij jou in de groepsfase laatste "
    printObj.Print "wordt toch de halve finale laten spelen, of een team dat niet in de finale staat toch kampioen maken."
    If getTournamentInfo("tournamentGroupCount", cn) = 6 And getTournamentInfo("tournamentTeamCount", cn) = 24 Then ' de vier beste derde plaatsen naar kwart finales
      printObj.Print "De beste 4 derde plaatsen kwalificeren zich ook voor de 8e finales."
    End If
    subHeader "Prijzen"
    printObj.ForeColor = vbBlack
    'printObj.Print "Na de finale worden de hoofdprijzen te verdeeld, maar ook per dag zijn er geldprijzen te winnen."
    printObj.FontBold = True
    printObj.Print "-  Per dag";
    printObj.FontBold = False
    printObj.Print " zijn de volgende geldprijzen te verdienen:"
    printObj.Print "  -  ";
    printObj.Print "Degene die op ";
    printObj.FontItalic = True
    printObj.Print "één dag de meeste punten";
    printObj.FontItalic = False
    printObj.Print " heeft verzameld, ";
    printObj.Print " krijgt daarvoor ";
    printObj.FontBold = True
    printObj.Print Format(getPoolInfo("prizeHighDayScore", cn), "currency")
    printObj.FontBold = False
    printObj.Print "  -  ";
    printObj.Print "Degene die na een dag in de ";
    printObj.FontItalic = True
    printObj.Print "totaalstand bovenaan";
    printObj.FontItalic = False
    printObj.Print " staat, ";
    printObj.Print " krijgt daarvoor ";
    printObj.FontBold = True
    printObj.Print Format(getPoolInfo("prizeHighDayPosition", cn), "currency")
    printObj.FontBold = False
    printObj.Print "  -  ";
    printObj.Print "Degene die na een dag in de ";
    printObj.FontItalic = True
    printObj.Print "totaalstand onderaan";
    printObj.FontItalic = False
    printObj.Print " staat, ";
    printObj.Print " krijgt daarvoor als troost ";
    printObj.FontBold = True
    printObj.Print Format(getPoolInfo("prizeLowDayPosition", cn), "currency")
    printObj.FontBold = False
    printObj.Print "  -  ";
    xPos = printObj.CurrentX
    printObj.Print "De punten voor de finalerondes tellen mee voor de dagprijs op de dag dat de teams bekend zijn"
    printObj.CurrentX = xPos
    printObj.Print "De punten voor de eindstand, topscorers en aantallen tellen op de dag van de finale mee voor de dagprijs"
    printObj.Print "-  ";
    printObj.FontBold = True
    printObj.Print "Aan het eind van het toernooi";
    printObj.FontBold = False
    printObj.Print " zijn de volgende geldprijzen te verdienen:"
    amount = getPoolInfo("prizeLowFinalPosition", cn)
    If amount > 0 Then
        printObj.Print "  -  ";
        xPos = printObj.CurrentX
        printObj.Print "De ";
        printObj.FontItalic = True
        printObj.ForeColor = vbRed
        printObj.Print "rode lantaarn";
        printObj.ForeColor = vbBlack
        printObj.FontItalic = False
        printObj.Print " ontvangt als troostprijs "; Format(amount, "currency")
    End If
    
    printObj.Print "  -  ";
    xPos = printObj.CurrentX
    printObj.Print "De ";
    printObj.FontItalic = True
    printObj.Print "hoogste";
    printObj.FontItalic = False
    printObj.Print " deelnemers in de totaalstand krijgen de volgende prijzen:"
    printObj.CurrentX = xPos
    
    printObj.Print "1e pl: ";
    printObj.FontBold = True
    printObj.Print Format(getPoolInfo("prizePercentage1", cn) / 100, "0%");
    printObj.FontBold = False
    amount = getPoolInfo("prizePercentage2", cn)
    If amount > 0 Then
        printObj.Print ", 2e pl: ";
        printObj.FontBold = True
        printObj.Print Format(amount / 100, "0%");
        printObj.FontBold = False
    End If
    amount = getPoolInfo("prizePercentage3", cn)
    If amount > 0 Then
        printObj.Print ", 3e pl: ";
        printObj.FontBold = True
        printObj.Print Format(amount / 100, "0%");
        printObj.FontBold = False
    End If
    amount = getPoolInfo("prizePercentage4", cn)
    If amount > 0 Then
        printObj.Print ", 4e pl: ";
        printObj.FontBold = True
        printObj.Print Format(amount / 100, "0%");
        printObj.FontBold = False
    End If
    printObj.Print " van de totale inleg (minus de dagprijzen en de rode lantaarn)"
    printObj.Print "-  ";
    printObj.FontItalic = True
    printObj.Print "Bij een gelijk aantal punten wordt de betreffende prijs verdeeld"
    printObj.FontItalic = False

End Sub

Sub printSectionHeader(xPos As Integer, yPos As Integer, width As Integer, header As String, Optional subText As String)
    fontSizing 14
    printObj.FontBold = True
    printObj.FillColor = headBGcolor
    printObj.FillStyle = vbFSSolid
    printObj.Line (xPos, yPos - 10)-(xPos + width, yPos + printObj.TextHeight("W") + 10), lineColor, B
    printObj.CurrentY = yPos
    printObj.CurrentX = xPos + 50
    iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
    printObj.ForeColor = headTxtColor
    printObj.Print header;
    fontSizing 10
'    printObj.FontBold = False
    printObj.CurrentY = yPos + 60
    printObj.Print subText;
    printObj.CurrentY = yPos
    fontSizing 14
    printObj.ForeColor = vbBlack
    printObj.FontBold = False
    printObj.Print
End Sub

Sub printPoolFormTopScorers(xPos As Integer, yPos As Integer, headerWidth As Integer)
  Dim xPos2 As Integer, yPos2 As Integer
  Dim i As Integer
  Dim savYpos As Integer
  Dim txt As String
  Dim pnts As Integer
  Dim points(5) As Integer 'get the points for different categories
  
  'topscorers
  pnts = getPointsFor("topscorer 1", cn)
  points(0) = pnts
  If points(0) > 0 Then
    pnts = getPointsFor("doelpunten topscorer 1", cn)
    points(1) = pnts
  Else
    points(1) = 0
  End If
  pnts = getPointsFor("topscorer 2", cn)
  points(2) = pnts
  If points(2) > 0 Then
    pnts = getPointsFor("doelpunten topscorer 2", cn)
    points(3) = pnts
  Else
    points(3) = 0
  End If
''   headerWidth = (printObj.ScaleWidth) / 4 - 100
  If points(2) Then
    txt = "Topscorers & aantal goals"
  Else
    txt = "Topscorer & aantal goals"
  End If
  If points(0) Then
    printSectionHeader xPos, yPos, headerWidth, txt
  End If
  yPos = printObj.CurrentY
  printObj.CurrentY = printObj.CurrentY + 70
  printObj.FillStyle = vbFSTransparent
  For i = 0 To 3 Step 2
    If points(i) Then
      printObj.CurrentX = xPos + 50
      fontSizing 14
      If points(2) Then   'print numbers before topscorers
        printObj.Print Format(i + 1, "0") & ":";
      End If
      printObj.CurrentX = xPos + 50
      xPos2 = xPos + headerWidth / 4 * 3 - 100
      printObj.CurrentX = xPos2 - printObj.TextWidth("(30) ")
      fontSizing 12
      printObj.Print "(" & Format(points(i), "0") & "p)";
      fontSizing 18
      yPos2 = yPos + printObj.TextHeight("99")
      printObj.Line (xPos, yPos)-(xPos2 + 50, yPos2), lineColor, B
      printObj.CurrentY = yPos + 70
      printObj.CurrentX = xPos + headerWidth - TextWidth("(0000)")
      fontSizing 12
      printObj.Print "(" & Format(points(i + 1), "0") & "p)";
      printObj.Line (xPos2 + 50, yPos)-(xPos + headerWidth, yPos2), lineColor, B
    End If
  Next
  yPos = printObj.CurrentY
  fontSizing 18
  yPos2 = yPos + printObj.TextHeight("99") * 1.8
  printObj.Line (xPos, yPos)-(xPos + headerWidth, yPos2), lineColor, B
  fontSizing 12
  printObj.CurrentY = yPos + 70
  txt = "Reserve:"
  printObj.CurrentX = xPos + 50
  printObj.Print txt
  fontSizing 8
  printObj.Print
  txt = "(als je topscorer niet in een van de teams zit na " & Format(CDbl(getTournamentInfo("tournamentStartDate", cn) - 1), "d MMM") & ")"
  printObj.CurrentX = xPos + 50
  printObj.Print txt;
  fontSizing 12
  
End Sub

Sub printPoolFormGroupSection(xPos As Integer, yPos As Integer)
  Dim txt As String
  Dim i As Integer
  Dim pnts As Integer
  Dim columnWidth As Integer
  Dim groupCount As Integer
    groupCount = getTournamentInfo("tournamentGroupCount", cn)
    'groepsstanden
    fontSizing 10
    printObj.Print
    pnts = getPointsFor("groepstand per juist team", cn)
    txt = "   Vul in: 1 t/m 4 (" & CStr(pnts) & " pnt per correcte invoer)"
    
    printSectionHeader xPos, yPos, printObj.ScaleWidth, "Groepsstanden", txt
    
    yPos = printObj.CurrentY
    xPos = printObj.CurrentX
    
    fontSizing 10
    printObj.FillStyle = vbFSTransparent
    'draw square aroung group section
    printObj.Line (xPos, yPos)-(printObj.ScaleWidth - 20, yPos + printObj.TextHeight("W") * 5.5), lineColor, B
    
    columnWidth = printObj.ScaleWidth / groupCount
    printObj.ForeColor = vbBlack
    For i = 1 To groupCount
        fontSizing 12
        xPos = columnWidth * (i - 1) + 50
        printObj.CurrentY = yPos + 10
        printObj.CurrentX = xPos
        printObj.FontBold = True
        printObj.Print "Groep " & Chr(i + 64)
        printObj.FontBold = False
        printPoolFormGroupBlock i
        If i > 1 Then        'for all groups after the first draw a line
          printObj.Line (xPos - 80, yPos)-(xPos - 80, yPos + printObj.TextHeight("W") * 5.5), lineColor
        End If
    Next
    printObj.Font = textFont
    fontSizing 8
    printObj.Print
End Sub

Sub printPoolFormBottomBlock(xPos As Integer, yPos As Integer)
  Dim headerWidth As Integer
  Dim xPos2 As Integer, yPos2 As Integer
  Dim i As Integer
  Dim savYpos As Integer
  Dim txt As String
  Dim pnts As Integer
  Dim points(5) As Integer 'get the points for different categories
  'remember vertical position
  savYpos = yPos
  'Champions and runners up
  headerWidth = printObj.ScaleWidth / 4
  printSectionHeader xPos + 50, yPos, headerWidth, "Eindstand"
  printObj.FillStyle = vbFSTransparent
  For i = 0 To 3
    pnts = getPointsFor(Format(i + 1, "0") & "e plaats", cn)
    points(i) = pnts
    'printObj.CurrentY = printObj.CurrentY + 50
    If points(i) > 0 Then
      yPos = printObj.CurrentY
      printObj.CurrentY = yPos + 50
      printObj.CurrentX = xPos + 70
      fontSizing 12
      printObj.Print Format(i + 1, "0") & "e:";
      printObj.CurrentX = headerWidth - printObj.TextWidth("(50p)")
      printObj.Print "(" & Format(points(i), "0") & "p)"
      fontSizing 18
      xPos2 = xPos + 50 + headerWidth
      yPos2 = yPos + printObj.TextHeight("99")
      printObj.Line (xPos + 50, yPos)-(xPos2, yPos2), lineColor, B
      printObj.CurrentY = yPos2
    End If
  Next
  
  'topscorers
  xPos = xPos + headerWidth + 100
  yPos = savYpos
  headerWidth = printObj.ScaleWidth * 3 / 4 / 2 - 80
  printPoolFormTopScorers xPos, yPos, headerWidth
  
  'aantallen
  xPos = xPos + headerWidth + 50
  yPos = savYpos
  printPoolFormNumberCounts xPos, yPos, headerWidth - 80
  
  'square around bottom block
  yPos = savYpos - 50
  xPos = 0
  printObj.FillStyle = vbFSTransparent
  xPos2 = printObj.ScaleWidth - 10
  yPos2 = printObj.CurrentY + 50
  printObj.Line (xPos, yPos)-(xPos2, yPos2), lineColor, B
End Sub

Sub printPoolFormNumberCounts(xPos As Integer, yPos As Integer, headerWidth As Integer)
  Dim xPos2 As Integer, yPos2 As Integer
  Dim i As Integer
  Dim savYpos As Integer
  Dim txt As String
  Dim pnts As Integer
  Dim sqlstr As String
  Dim pointsLinePos As Integer
  txt = "Aantallen"
  printSectionHeader xPos, yPos, headerWidth, txt
  
  sqlstr = "SELECT pointTypeDescription as descr, pointPointsAward as pnt, pointPointsMargin as mrg from tblPoolPoints p "
  sqlstr = sqlstr & "INNER JOIN tblPointTypes t on p.pointTypeId = t.pointTypeId "
  sqlstr = sqlstr & " WHERE p.poolId = " & thisPool
  sqlstr = sqlstr & " AND t.pointTypeCategory = 6 "
  sqlstr = sqlstr & " ORDER BY t.pointTypeListOrder"
  
  Set rs = New ADODB.Recordset
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  printObj.FillStyle = vbFSTransparent
  savYpos = printObj.CurrentY
  Do While Not rs.EOF
    yPos = printObj.CurrentY
    fontSizing 12
    printObj.CurrentX = xPos + 50
    printObj.CurrentY = yPos + 50
    txt = Mid(rs!descr, 8) & "("
    If rs!mrg > 0 Then txt = txt & "±" & rs!mrg & ", "
    txt = txt & rs!pnt & "p):"
    If printObj.TextWidth(txt) > pointsLinePos Then pointsLinePos = printObj.TextWidth(txt)
    printObj.Print txt
    fontSizing 16
    yPos2 = yPos + printObj.TextHeight("AA")
    xPos2 = xPos + headerWidth
    printObj.Line (xPos, yPos)-(xPos2, yPos2), lineColor, B
    printObj.CurrentY = yPos2
    rs.MoveNext
  Loop
  pointsLinePos = pointsLinePos + xPos + 150
  'points line
  xPos = pointsLinePos
  printObj.Line (xPos, savYpos)-(xPos, yPos2)
  rs.Close
  Set rs = Nothing
End Sub

Sub printPoolFormFinaleBlock(whichFinal As Integer, xPos As Integer, yPos As Integer)
'5 = 8th; 2 = 4th; 3 = half; 4 = final
Dim finaleType As String
Dim pntTeamName As Integer
Dim pntTeamPlace As Integer
Dim pnts As Integer
Dim txt As String
Dim sqlstr As String
Dim xPos1 As Integer, yPos1 As Integer
Dim xPos2 As Integer, yPos2 As Integer
Dim matchCount As Integer
Dim columnWidth As Integer
Dim columnNr As Integer
Dim shiftHor As Integer, shiftVert As Integer
Dim startYpos As Integer
Dim lineHeight As Integer  'to save vertical position in case of an empty block
Dim headerWidth As Integer
Dim widthAdjust As Integer
  widthAdjust = 0
  yPos1 = yPos
  xPos1 = xPos
  
  Select Case whichFinal
  Case 5
    finaleType = "Achtste Finale"
  Case 2
    finaleType = "Kwart Finale"
  Case 3
    finaleType = "Halve Finale"
  Case 4
    finaleType = "Finale"
    widthAdjust = 40
  Case 7
    finaleType = "Kleine finale"
    widthAdjust = 30
  End Select
  pnts = getPointsFor(LCase(finaleType) & " team", cn)
  pntTeamName = CStr(pnts)
  pnts = getPointsFor(LCase(finaleType) & " positie", cn)
  pntTeamPlace = CStr(pnts)
  If pntTeamName + pntTeamPlace = 0 Then
    txt = ""
  Else
    txt = " ("
    
    If pntTeamName > 0 Then
      txt = txt & pntTeamName
      If whichFinal = 5 Or whichFinal = 2 Then txt = txt & " pnt per genoemd team"
    End If
    If pntTeamPlace > 0 Then
      If pntTeamName > 0 Then txt = txt & " / "
      txt = txt & pntTeamPlace
      If whichFinal = 5 Or whichFinal = 2 Then txt = txt & " pnt voor team op de juiste plaats"
    End If
    If whichFinal = 3 Or whichFinal = 4 Then txt = txt & "pnt"
    txt = txt & ")"
  End If
  ' get the matches
  Set rs = New ADODB.Recordset
  sqlstr = "Select matchnumber, matchteamA, matchteamB from tblTournamentSchedule "
  sqlstr = sqlstr & "where tournamentid = " & thisTournament
  sqlstr = sqlstr & " AND matchType = " & whichFinal
  sqlstr = sqlstr & " order by matchnumber"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If rs.EOF Then
    'no matches for this catgegory print empty block
      finaleType = ""
      headerWidth = printObj.ScaleWidth / 4 - 50
  Else
    If rs.RecordCount > 2 Then
      headerWidth = printObj.ScaleWidth
    Else
      headerWidth = printObj.ScaleWidth * rs.RecordCount / 4 - 20
    End If
    'make it a little bit ssmaller when for the 3rd place match
    If whichFinal = 7 Then
      headerWidth = printObj.ScaleWidth / 4 - 50
      widthAdjust = 20
    End If
  End If
  printSectionHeader xPos1, yPos1, headerWidth, finaleType, txt
  
  fontSizing 14
  startYpos = printObj.CurrentY 'remember postition for the square around the section
  xPos1 = xPos
  yPos1 = startYpos + 50
  xPos2 = xPos1
  yPos2 = yPos1
  lineHeight = printObj.TextHeight("99")
  With rs
    'print square around finals
    matchCount = rs.RecordCount
    columnWidth = printObj.ScaleWidth / 4 - 20
    columnNr = 0
    Do While Not .EOF
      fontSizing 8
      printObj.CurrentX = xPos1 + 50
      'shift down to center in block
      shiftVert = printObj.TextHeight("99") / 2
      printObj.CurrentY = yPos1 + shiftVert * 2
      printObj.Print Format(!matchNumber, "0"); ":"
      fontSizing 12
      'move starting point for square
      shiftHor = printObj.TextWidth("00:")
      printObj.CurrentX = xPos1 + shiftHor
      printObj.CurrentY = yPos1
      'print teamcode
      'fontSizing 12
      printObj.Print !matchteamA; ":";
      fontSizing 14
      shiftVert = printObj.TextHeight("99")
      'draw square
      printObj.FillStyle = vbFSTransparent
      printObj.Line (xPos1 + shiftHor - 20, yPos1)-(xPos1 + columnWidth - 40 - widthAdjust, yPos1 + shiftVert), lineColor, B
      yPos1 = printObj.CurrentY
      printObj.CurrentX = xPos1 + shiftHor
      fontSizing 12
      'print teamcode
      printObj.Print !matchteamB; ":";
      'draw square
      printObj.Line (xPos1 + shiftHor - 20, yPos1)-(xPos1 + columnWidth - 40 - widthAdjust, yPos1 + shiftVert), lineColor, B
      'get next match
      .MoveNext
      'shift x to next column
      columnNr = columnNr + 1
      xPos1 = columnWidth * columnNr
      If columnNr > 3 Then 'move back to left margin
      columnNr = 0
        xPos1 = 0
        yPos2 = printObj.CurrentY + 100
      End If
      yPos1 = yPos2 'set vertical position back
    Loop
  End With
  'draw square around section
  
  yPos2 = printObj.CurrentY
  
  If matchCount > 2 Then 'full width
    xPos1 = 0
    xPos2 = printObj.ScaleWidth - 20
  Else
    If matchCount = 2 Then  'half finals
      xPos1 = 0
      xPos2 = printObj.ScaleWidth / 2 - 20
    Else
      If whichFinal = 7 Then
        xPos1 = printObj.ScaleWidth / 2 + 30
        xPos2 = printObj.ScaleWidth / 2 + columnWidth
        'If rs.RecordCount = 0 Then yPos2 = endYpos
      End If
      If whichFinal = 4 Then
        xPos1 = printObj.ScaleWidth - columnWidth + 20
        xPos2 = printObj.ScaleWidth - 10
      End If
    End If
    
  End If
  If rs.RecordCount = 0 Then
    'need special yPos2
    yPos2 = yPos1 + 2 * lineHeight
    printObj.FillStyle = vbUpwardDiagonal 'vbDiagonalCross '
  Else
    printObj.FillStyle = vbFSTransparent
  End If
  printObj.Line (xPos1, startYpos)-(xPos2, yPos2 + 30), lineColor, B
  
  rs.Close
  Set rs = Nothing
End Sub

Sub printPoolFormFinalSection(xPos As Integer, yPos As Integer)
'onderdeel van formulieren
Dim sqlstr As String
Dim txt As String
Dim pntTeamName As Integer
Dim pntTeamPlace As Integer
Dim thirdPlace As Boolean
  
  'check if there are 8th finals
  If getTournamentInfo("tournamentgroupcount", cn) >= 6 And getTournamentInfo("tournamentteamcount", cn) >= 24 Then
    printPoolFormFinaleBlock 5, xPos, yPos
  End If
  fontSizing 4
  printObj.Print
  
  '1/4 finals
  xPos = printObj.CurrentX
  yPos = printObj.CurrentY
  printPoolFormFinaleBlock 2, xPos, yPos
  fontSizing 4
  printObj.Print
  
  '1/2 finals
  xPos = printObj.CurrentX
  yPos = printObj.CurrentY
  printPoolFormFinaleBlock 3, xPos, yPos
  
  '3/4th place
  xPos = printObj.ScaleWidth / 2 + 30
  printPoolFormFinaleBlock 7, xPos, yPos
  
  'final
  xPos = printObj.ScaleWidth / 4 * 3 + 30
  printPoolFormFinaleBlock 4, xPos, yPos
  
End Sub

Sub printPoolFormGroupBlock(nr As Integer)
Dim sqlstr As String
Dim xLinePos As Integer
Dim yLinePos As Integer
Dim xPos As Integer
Dim txt As String
Dim squarePos(1, 1)
Dim grp As String * 1
Dim iGrp As Integer

Set rs = New ADODB.Recordset
sqlstr = "Select groupLetter, groupPlace, teamName from (tblGroupLayout l"
sqlstr = sqlstr & " INNER JOIN tblTournamentTeamCodes c ON (l.teamId = c.teamId) "
sqlstr = sqlstr & " AND (l.tournamentID = c.tournamentId))"
sqlstr = sqlstr & " INNER JOIN tblTeamNames n on n.teamNameId = c.teamid "
sqlstr = sqlstr & " WHERE l.groupLetter = '" & Chr(64 + nr) & "'"
sqlstr = sqlstr & " AND left(c.teamCode,1) = '" & Chr(64 + nr) & "'"
sqlstr = sqlstr & " AND l.tournamentID = " & thisTournament
sqlstr = sqlstr & " ORDER BY groupletter, groupPlace"
rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly

yLinePos = printObj.CurrentY
iGrp = getTournamentInfo("tournamentGroupCount", cn)

fontSizing 10

xLinePos = (printObj.ScaleWidth / iGrp) * (nr - 1)
xPos = xLinePos + 50
Do While Not rs.EOF
    squarePos(0, 0) = Int(xPos + printObj.ScaleWidth / iGrp - printObj.TextHeight("W") - printObj.TextWidth("W"))
    squarePos(0, 1) = printObj.CurrentY
    squarePos(1, 0) = Int(squarePos(0, 0) + printObj.TextHeight("W"))
    squarePos(1, 1) = Int(squarePos(0, 1) + printObj.TextHeight("W"))

    txt = rs!teamName

    Do While xPos + printObj.TextWidth(txt) > squarePos(0, 0)
        txt = Left(txt, Len(txt) - 1)
    Loop
    printObj.CurrentX = xPos
    printObj.Print txt;
    printObj.FillStyle = vbFSTransparent
    printObj.FillColor = vbWhite
    printObj.DrawWidth = 1

    printObj.Line (squarePos(0, 0), squarePos(0, 1))-(squarePos(1, 0), squarePos(1, 1)), lineColor, B
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
'printObj.CurrentY = yLinePos
End Sub
'

Sub printPoolFormMatchesInstructions()
'print the explanation text for the 2nd poolform page
  Dim pnt As Integer
  fontSizing 12
  subHeader "Uitleg"
  printObj.FontBold = False
  printObj.Print "Vul voor ";
  printObj.FontBold = True
  printObj.Print "alle";
  printObj.FontBold = False
  printObj.Print " wedstrijden je voorspelling in voor rust- en eindstand en toto.";
'  printObj.FontBold = True
  printObj.Print " Ook waar de teams nog niet bekend zijn."
  printObj.FontBold = False
  printObj.Print "(Ook al heb je een ander team op die plaats dan kan je uitslag nog steeds goed zijn)"
  printObj.Print "De uitslag hoeft onderling niet te kloppen. ";
  printObj.Print "Je krijgt punten voor elk vak dat achteraf juist blijkt te zijn ingevuld."
  printObj.Print "Bij 'toto' vul je een 1 in voor winst linker team, een 2 voor winst rechter team en een 3 voor een gelijkspel"
  printObj.FontBold = True
  centerText "Alle uitslagen, ook de toto, gelden na 90 minuten voetbal!"
  printObj.FontBold = False
  printObj.Print
  fontSizing 9
  centerText "(plus de eventuele blessuretijd)"
  printObj.Print
  fontSizing 11
  subHeader "Punten"
  printObj.Print "Ruststand goed: ";
  printObj.FontBold = True
  pnt = getPointsFor(LCase("ruststand goed"), cn)
  printObj.Print pnt; "pnt";
  printObj.FontBold = False
  printObj.Print ", Eindstand goed: ";
  printObj.FontBold = True
  pnt = getPointsFor(LCase("eindstand goed"), cn)
  printObj.Print pnt; "pnt";
  printObj.FontBold = False
  printObj.Print ", Toto goed: ";
  printObj.FontBold = True
  pnt = getPointsFor(LCase("toto goed"), cn)
  printObj.Print pnt; "pnt";
  printObj.FontBold = False
  pnt = getPointsFor(LCase("doelpunten op een dag"), cn)
  printObj.Print ".";
  If pnt > 0 Then
      printObj.FontBold = True
      printObj.Print " BONUS";
      printObj.FontBold = False
      printObj.Print ": totaal doelpunten op één dag goed: ";
      printObj.FontBold = True
      printObj.Print pnt; " pnt"
      printObj.FontBold = False
  End If
  printObj.Print
  

End Sub

Sub printPoolFormMatches()
''wedstrijden op het poolformulier
Dim i As Integer, J As Integer
Dim X As Integer, x2 As Integer, y As Integer
Dim currentColumn As Integer
Dim columnWidth As Integer
Dim columnPos(6) As Integer
Dim columnNames() As String
Dim savYpos As Integer
Dim savLineYpos(1) As Integer  '0= top of the table, 1= bottom of table
Dim yPos As Integer
Dim sqlstr As String
Dim matchDescription As String
Dim savDate As Date  'to skip trhe same date on the form
'print instructions
  printPoolFormMatchesInstructions

'get the match records
  sqlstr = "SELECT matchNumber, matchDate, matchTime, tc1.teamCode as tcA, tn1.teamName as tnA, tc2.teamCode as tcB, tn2.teamName as tnB"
  sqlstr = sqlstr & " FROM (((tblTournamentSchedule ts "
  sqlstr = sqlstr & " LEFT JOIN tblTournamentTeamCodes AS tc1 ON (ts.matchTeamA = tc1.teamCode) AND (ts.tournamentID = tc1.tournamentID)) "
  sqlstr = sqlstr & " LEFT JOIN tblTeamNames AS tn1 ON tc1.teamID = tn1.teamNameID) "
  sqlstr = sqlstr & " LEFT JOIN tblTournamentTeamCodes AS tc2 ON (ts.matchTeamB = tc2.teamCode) AND (ts.tournamentID = tc2.tournamentID)) "
  sqlstr = sqlstr & " LEFT JOIN tblTeamNames AS tn2 ON tc2.teamID = tn2.teamNameID"
  sqlstr = sqlstr & " Where ts.TournamentId = " & thisTournament
  sqlstr = sqlstr & " ORDER BY ts.matchOrder;"
  Set rs = New ADODB.Recordset
  
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
'set column positions
  columnWidth = printObj.ScaleWidth / 2 - 50
  columnPos(0) = 30 'datum
  fontSizing 10
  columnPos(1) = columnPos(0) + printObj.TextWidth("MA 30-11") + 20 ' tijd
  columnPos(2) = columnPos(1) + printObj.TextWidth("00:000") + 10  'nr
  columnPos(3) = columnPos(2) + printObj.TextWidth("123") 'wedstrijd
  columnPos(4) = columnPos(3) + printObj.TextWidth("Zwitserland - N-Macedonie")  'rust
  columnPos(5) = columnPos(4) + (columnWidth - 50 - columnPos(4)) / 2.5 'eind
  columnPos(6) = columnPos(5) + (columnWidth - 50 - columnPos(4)) / 2.5 'toto
  
  columnNames = Split("Datum; tijd; nr; Wedstrijd; rust; eind;toto", ";")
  
  printObj.ForeColor = vbBlack
  printObj.FillStyle = vbFSTransparent
  printObj.FontBold = False
  printObj.FontName = textFont
  fontSizing 10
    
  savYpos = printObj.CurrentY 'save vertical position fot 2nd column
  savLineYpos(0) = savYpos
'    vertLineYPos = printObj.CurrentY
  printObj.CurrentY = savYpos
  For J = 0 To 1
    For i = 0 To 6
      printObj.CurrentX = columnPos(i) + J * (columnWidth + 50)
      If i <> 6 Then
        printObj.Print " ";
      Else
        printObj.CurrentX = columnPos(i) + J * (columnWidth)
      End If
      printObj.Print columnNames(i);
    Next
    
  Next
  printObj.Print
  savYpos = printObj.CurrentY
  printObj.FillStyle = vbFSTransparent
  printObj.Line (0, savLineYpos(0) - 20)-(columnWidth - 50, savYpos), lineColor, B
  printObj.Line (columnWidth + 50, savLineYpos(0) - 20)-(columnWidth * 2 - 50, savYpos), lineColor, B
  yPos = printObj.CurrentY
  printObj.FontName = smallFont
  With rs
    Do While Not .EOF
      printObj.CurrentY = printObj.CurrentY + 50
      'Datum
      printObj.CurrentX = columnPos(0) + currentColumn * (columnWidth + 50)
      If savDate <> !matchDate Or printObj.CurrentY = savYpos Then
          printObj.Print Format(!matchDate, "ddd d-M"); " ";
          savDate = !matchDate
      End If
      'center the time
      printObj.CurrentX = columnPos(1) + currentColumn * (columnWidth + 50) + (columnPos(2) - columnPos(1) - printObj.TextWidth(Format(!matchtime, "HH:NN"))) / 2
      printObj.Print Format(!matchtime, "HH:NN");
      printObj.CurrentX = columnPos(2) + currentColumn * (columnWidth + 50) + (columnPos(3) - columnPos(2) - printObj.TextWidth(Format(!matchNumber, "0"))) / 2
      printObj.Print Format(!matchNumber, "0");
      printObj.CurrentX = columnPos(3) + currentColumn * (columnWidth + 50) + 50
      yPos = printObj.CurrentY
      If nz(!tna, "") > "" Then
        fontSizing 8
        matchDescription = !tca & ":"
        If !tna <> !tca Then
          matchDescription = matchDescription & Left(!tna, 12)
        Else
          matchDescription = matchDescription & "?"
        End If
        matchDescription = matchDescription & " - " & !tcB & ":"
        If !tnb <> !tcB Then
          matchDescription = matchDescription & Left(!tnb, 12)
        Else
          matchDescription = matchDescription & "?"
        End If
        printObj.CurrentY = yPos + 30
        matchDescription = fitText(columnPos(4) - columnPos(3) - 30, matchDescription) 'fit text width
      Else
        fontSizing 10
        matchDescription = !tca & " - " & !tcB
      End If
      printObj.Print matchDescription;
      printObj.CurrentY = yPos
      fontSizing 10
    ' matchscore blocks
      For i = 1 To 2
        X = columnPos(3 + i) + currentColumn * (columnWidth)
        x2 = X + columnPos(6) - columnPos(5) - 50
        printObj.Line (X, yPos - 20)-(x2, printObj.CurrentY + printObj.TextHeight("Z") + 30), lineColor, B
        printObj.CurrentY = yPos
        printObj.CurrentX = X + (x2 - X - printObj.TextWidth("-")) / 2
        printObj.Print "-";
      Next
      'toto block
      X = columnPos(6) + currentColumn * (columnWidth)
      x2 = X + (columnPos(6) - columnPos(5)) / 2 - 50
      printObj.Line (X, yPos - 20)-(x2, printObj.CurrentY + printObj.TextHeight("Z") + 30), lineColor, B
      printObj.CurrentY = yPos
      'finish up
      fontSizing 13
      printObj.Print
      fontSizing 10
      yPos = printObj.CurrentY
      printObj.Line (currentColumn * (columnWidth + 50), yPos)-((currentColumn + 1) * (columnWidth - 50), yPos), lineColor
      rs.MoveNext
      If (.AbsolutePosition - 1) = Int(.RecordCount / 2 + 0.5) Then
        currentColumn = 1
        savLineYpos(1) = printObj.CurrentY  'end point vertical lines
        printObj.CurrentY = savYpos
      End If
    Loop
  End With
  'vertical lines
  For J = 0 To 1
    X = J * (columnWidth + 50)
    For i = 1 To 3
      printObj.Line (columnPos(i) + X, savLineYpos(0))-(columnPos(i) + X, savLineYpos(1)), lineColor
    Next
  Next
  printObj.Line (0, savLineYpos(0))-(columnWidth - 50, savLineYpos(1)), lineColor, B
  printObj.Line (columnWidth + 50, savLineYpos(0))-(columnWidth * 2 - 50, savLineYpos(1)), lineColor, B
  
End Sub

Sub printPoolFormDeliverDate()
Dim yPos As Integer
Dim yPos2 As Integer
' PRINT FFooter
  fontSizing 16
  yPos = printObj.ScaleHeight - printObj.TextHeight("INLEVEREN") - printObj.ScaleTop
  yPos2 = printObj.ScaleHeight - printObj.ScaleTop
  printObj.CurrentY = yPos
  printObj.FillColor = headBGcolor
  printObj.FillStyle = vbFSSolid
  printObj.Line (0, yPos - 20)-(printObj.ScaleWidth, printObj.ScaleHeight - 10), headBGcolor, B
  printObj.CurrentY = yPos
  fontSizing 16
  printObj.FontBold = True
  printObj.ForeColor = headTxtColor
  iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
  centerText "UITERLIJK INLEVEREN OP " & UCase(Format(getPoolInfo("poolFormsTill", cn), "dddd d mmmm yyyy"))


End Sub


'
'Private Sub PrijsAfdr(wat As String, eind As Boolean)
'Dim aant As Integer
'Dim i As Integer
'End Sub
'
Private Sub centerText(txt As String)
    printObj.CurrentX = (printObj.ScaleWidth - printObj.TextWidth(Trim(txt))) \ 2
    printObj.Print txt;
End Sub
'
Function sqlDeelnems(poule As Long) As String

'Dim sqlstr As String
'    sqlstr = "Select * from pooldeelnems"
'    sqlstr = sqlstr & " WHERE PoolID = " & poule
'    sqlstr = sqlstr & " ORDER BY bijnaam "
'    sqlDeelnems = sqlstr
End Function
'
Private Sub printFavourites()

Dim grpCount As Integer
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim cnt As Integer
Dim savX As Integer
Dim savy As Integer
Dim xPos As Integer
Dim col(6) As Integer
Dim pageCol As Integer
Dim yStart As Integer
Dim maxrows As Integer
Dim savYpos As Integer
Dim poolFormCount As Integer
Dim sqlstr As String
Dim headerWidth As Integer
Set rs = New ADODB.Recordset
Dim nwPag As Boolean
Dim finYpos As Integer
Dim favYpos As Integer
Dim favXpos As Integer
  poolFormCount = getpoolFormCount(cn)
  headerWidth = printObj.ScaleWidth
  
  heading1 = "Favorieten voor het toernooi"
  pageCol = printObj.ScaleWidth / 2
  InitPage True, False, , , , True
  'printSectionHeader xPos, yStart, headerWidth, "Groepstanden"
  reportTitle "Groepstanden"
  yStart = printObj.CurrentY
  fontSizing 11
  'groepen
  col(0) = 0
  col(1) = printObj.TextWidth("GROEP A: ") 'teamname
  col(2) = col(1) + printObj.TextWidth("Group A: ZWITESRLAND W ") '1e pl
  col(3) = col(2) + printObj.TextWidth("1e PlXX") '2e pl
  col(4) = col(3) + printObj.TextWidth("1e PlXX") '3e pl
  col(5) = col(4) + printObj.TextWidth("1e PlXX") '4e pl
  
  'printObj.Print
  grpCount = getTournamentInfo("tournamentgroupcount", cn)
  savy = printObj.CurrentY
  For J = 0 To 1
    printObj.CurrentY = savy
    For i = 1 To 4
      printObj.CurrentX = (pageCol * J) + col(i + 1) - printObj.TextWidth(Format(i, "0") & "e pl")
      printObj.Print Format(i, "0"); "e pl";
    Next
  Next
  printObj.Print
  savy = printObj.CurrentY
  xPos = 0
  J = 0
  For i = 1 To grpCount
      If i = grpCount / 2 + 1 Then
          J = J + 1
          printObj.CurrentY = savy
      End If
      sqlstr = "Select * from tblGroupLayout where tournamentID = " & thisTournament
      sqlstr = sqlstr & " AND groupletter = '" & Chr(i + 64) & "'"
      rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
      rs.MoveFirst
      printObj.CurrentX = J * pageCol
      printObj.Print "Groep " & rs!groupletter; ": ";
      Do While Not rs.EOF
          printObj.CurrentX = col(1) + (pageCol * J)
          printObj.Print getTeamInfo(rs!teamID, "teamname", cn); " ";
          For K = 1 To 4
            cnt = getGrpPredictionCount(K, rs!teamID, cn)
            printObj.CurrentX = (pageCol * J) + col(K + 1) - printObj.TextWidth(Format(cnt / poolFormCount, "0.0%"))
            printObj.Print Format(cnt / poolFormCount, "0.0%");
            'If j < 4 Then printObj.Print ", ";
          Next
          printObj.Print
          rs.MoveNext
      Loop
      rs.Close
  Next
  savy = printObj.CurrentY
  On Error Resume Next
  printObj.Line (0, yStart)-(printObj.ScaleWidth - 50, savy), lineColor, B
  On Error GoTo 0
  maxFavYpos = savy
  'achtste finales
  i = getPointsFor("achtste finale team", cn) + getPointsFor("achtste finale positie", cn)
  If i > 0 Then
      printFavFinals 5, 4, "Achtste finales"
      savy = printObj.CurrentY
  End If
  printObj.CurrentY = savy
  'kwart finales
  i = getPointsFor("kwart finale team", cn) + getPointsFor("kwart finale positie", cn)
  If i > 0 Then
      printFavFinals 2, 4, "Kwart finales"
      savy = printObj.CurrentY
  End If
  printObj.CurrentY = savy
  'halve finales
  i = getPointsFor("halve finale team", cn) + getPointsFor("halve finale positie", cn)
  If i > 0 Then
      printFavFinals 3, 4, "Halve finales"
      savy = printObj.CurrentY
      maxFavYpos = savy
  End If
  printObj.CurrentY = savy
  'kleine finale
  i = getPointsFor("kleine finale team", cn) + getPointsFor("kleine finale positie", cn)
  If i > 0 Then
      savYpos = printObj.CurrentY
      printFavFinals 7, 4, "Kleine finale"
      'savy = maxFavYpos
      savX = 3
  Else
      savYpos = printObj.CurrentY
      savX = 1
  End If
'
  'finale
  i = getPointsFor("finale team", cn) + getPointsFor("finale positie", cn)
  If i > 0 Then
      printObj.CurrentY = savYpos
      printFavFinals 4, 4, "Finale", savy, savX
      If savX = 3 Then
          savX = 1
          savy = printObj.CurrentY
      Else
          savy = savYpos
          savX = 3
      End If
      savy = printObj.CurrentY
      maxFavYpos = savy
  End If
  printObj.CurrentY = savy
  printFavTournamentResult savy, savX
  
  printObj.CurrentY = savy
  printFavTopScorers
  
  'printObj.Print
  
  printPoolPrizes
  
  Set rs = Nothing
End Sub

Sub printPoolPrizes(Optional subHdr As Boolean)
Dim txtStr As String
Dim days As Integer
Dim dayMoney As Double
Dim poolCost As Double
Dim moneyLeft As Double
Dim poolFormCount As Integer
Dim yPos As Integer
Dim cols(2) As Integer
Dim prizes(4) As Double
Dim i As Integer
Dim indent As Integer
  If printObj.CurrentY > footerPos - printObj.TextHeight("HALLO") * 6 Then
    'nieuwe pagina graag
    addNewPage False, False, 240
  End If
  If subHdr Then
    subHeading "Prijzen aan het eind van het toernooi", True
  Else
    reportTitle "Prijzen aan het eind van het toernooi", False, False, printObj.CurrentY, 0
  End If
  fontSizing 10
  indent = printObj.TextWidth("1234567890")
  cols(0) = printObj.ScaleWidth / 2 - indent
  cols(1) = printObj.ScaleWidth / 2 + indent
  cols(2) = cols(1) + 2.5 * indent
  yPos = printObj.CurrentY
  days = getTournamentDayCount(cn)
  dayMoney = getPoolInfo("prizehighdayscore", cn)
  dayMoney = dayMoney + getPoolInfo("prizehighdayPosition", cn)
  dayMoney = dayMoney + getPoolInfo("prizeLowdayPosition", cn)
  poolCost = getPoolInfo("poolcost", cn)
  poolFormCount = getpoolFormCount(cn)
  
  prizes(0) = getPoolInfo("prizeLowFinalPosition", cn)
  For i = 1 To 4
    prizes(i) = getPoolInfo("prizePercentage" & i, cn)
  Next
  moneyLeft = poolCost * poolFormCount - (dayMoney * days) - prizes(0)
  
  printObj.CurrentX = indent
  printObj.Print "Inleg: (" & poolFormCount & " x " & Format(poolCost, "currency") & " = )";
  printObj.CurrentX = cols(0) - printObj.TextWidth(Format(poolCost * poolFormCount, "currency"))
  printObj.Print Format(poolCost * poolFormCount, "currency")
  printObj.CurrentX = indent
  printObj.Print "Dagprijzen: (" & days & " x " & Format(dayMoney, "currency") & " = )";
  printObj.CurrentX = cols(0) - printObj.TextWidth(Format(dayMoney * days, "currency"))
  printObj.Print Format(dayMoney * days, "currency")
  printObj.CurrentX = indent
  printObj.Print "Rode lantaarn: ";
  printObj.CurrentX = cols(0) - printObj.TextWidth(Format(prizes(0), "currency"))
  printObj.Print Format(getPrizeMoney(0, cn), "currency")
  printObj.CurrentX = indent
  printObj.Print "Rest: ";
  printObj.CurrentX = cols(0) - printObj.TextWidth(Format(moneyLeft, "currency"))
  printObj.Print Format(moneyLeft, "currency")
  printObj.CurrentY = yPos
  For i = 1 To 4
    If prizes(i) > 0 Then
      printObj.CurrentX = cols(1)
      printObj.Print i & "e prijs (" & Format((prizes(i) / 100), "0%") & "):";
      printObj.CurrentX = cols(2) - printObj.TextWidth(Format(moneyLeft * (prizes(i) / 100), "€ 0.00"))
      printObj.Print Format(moneyLeft * (prizes(i) / 100), "€ 0.00")
    End If
  Next
  On Error Resume Next
  printObj.Line (0, yPos)-(printObj.ScaleWidth - 50, printObj.CurrentY), 1, B
  On Error GoTo 0
End Sub

Private Sub printDailyResults(toMatch As Integer)
Dim i As Integer
Dim J  As Integer
Dim savdat As Date
Dim goals As Integer
Dim sqlstr As String
Dim headerWidth As Integer
Dim rsLcl As ADODB.Recordset
Dim cols() As Integer
Dim colWidth As Integer
Dim colCount As Integer
colCount = 6
Dim matchNr As Integer
matchNr = getMatchNumber(toMatch, cn)
ReDim cols(colCount)

  cols(0) = printObj.TextWidth("EINDSTANDGOED")
  colWidth = (printObj.ScaleWidth - cols(0)) / colCount
  For i = 1 To colCount
    cols(i) = cols(i - 1) + colWidth
  Next
  Set rsLcl = New ADODB.Recordset

  heading1 = "Wie had wat goed?"
  InitPage True, False, , , , True
  If IsLastMatchOfDay(toMatch, cn) Then
    savdat = getMatchInfo(toMatch, "matchdate", cn)
    sqlstr = "Select matchOrder from tblTournamentSchedule WHERE tournamentID = " & thisTournament
    sqlstr = sqlstr & " AND cdbl(matchdate) = " & CDbl(savdat)
    sqlstr = sqlstr & " ORDER by matchOrder"
    rsLcl.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    Do While Not rsLcl.EOF
      printPoolFormResults rsLcl!matchOrder, True
      'printObj.Print
      rsLcl.MoveNext
    Loop
    rsLcl.Close
    printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
    printObj.CurrentX = 0
    If printObj.CurrentY > footerPos - 4 * printObj.TextHeight("HALLO") Then
      heading1 = ""
      addNewPage False, False
    End If
    subHeading "Aantal doelpunten van de dag goed:" ', False, False, printObj.CurrentY, 0
    goals = getGoalsPerDay(savdat, cn)
    
    sqlstr = "Select * from tblCompetitorPools where poolID = " & thisPool
    rsLcl.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    i = 0
    J = 0
    Do While Not rsLcl.EOF
      If getPredictionGoalsPerDay(rsLcl!competitorPoolID, savdat, cn) = goals Then
        printObj.CurrentX = cols(i)
        i = i + 1
        printObj.Print rsLcl!nickName;
        J = J + 1
        If i > colCount Then
          i = 0
          printObj.Print
        End If
      End If
      rsLcl.MoveNext
    Loop
    If J = 0 Then
        printObj.CurrentX = cols(0)
        printObj.Print "NIEMAND!!!"
    End If
    rsLcl.Close
    printObj.Print
    printObj.Line (cols(0), printObj.CurrentY + 50)-(printObj.ScaleWidth - 50, printObj.CurrentY + 50)
    printObj.CurrentX = 0
  Else
    printPoolFormResults toMatch, False
  End If
  Set rsLcl = Nothing
  printObj.Print
End Sub

Sub printPoolFormResults(matchOrder As Integer, fullDay As Boolean)
  Dim matchDescr As String
  Dim sqlstr As String
  Set rs = New ADODB.Recordset
  Dim rust As String
  Dim xStartPos As Integer
  Dim lastYpos As Integer
  Dim eind As String
  Dim toto As Integer
  Dim savYpos As Integer
  Dim formatStr As String
  Dim cols(6) As Integer
  Dim matchNr As Integer
  matchNr = getMatchNumber(matchOrder, cn)
  rust = getMatchResultPartStr(matchOrder, 0, cn)
  eind = getMatchResultPartStr(matchOrder, 1, cn)
  toto = getMatchresult(matchOrder, 7, cn)
  cols(0) = printObj.TextWidth("EINDSTANDGOED")
  formatStr = "\(0\);0;\ "

  
  matchDescr = getMatchDescription(matchOrder, cn, True, True, False)
  'reportTitle "Wedstrijd: " & matchdescr & ":  " & getMatchresultStr(matchnr, True, cn), , , , 1
  On Error Resume Next
  printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
  On Error GoTo 0
  printObj.CurrentX = 0
  subHeading matchOrder & "e wedstrijd: " & matchDescr & ":  " & getMatchResultPartStr(matchOrder, 2, cn), True
  
  savYpos = printObj.CurrentY
  printObj.Line (cols(0), printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
  printObj.CurrentX = 0
  '''RUST
  sqlstr = "Select  cp.nickName, hta & '-' & htb as rust from tblCompetitorPools cp"
  sqlstr = sqlstr & " INNER JOIN tblPrediction_Matchresults r ON cp.competitorPoolID = r.competitorpoolID"
  sqlstr = sqlstr & " WHERE cp.poolID =" & thisPool
  sqlstr = sqlstr & " AND r.matchOrder = " & matchOrder
  sqlstr = sqlstr & " AND hta & '-' & htb = '" & rust & "'"
  sqlstr = sqlstr & " ORDER BY cp.nickname"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  printObj.Print "Ruststand "; Format(rs.RecordCount, formatStr); ":";
 
  printPoolFormResultsBlock rs
  
  rs.Close
  
  printObj.Line (cols(0), printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
  printObj.CurrentX = 0
  
  '''EIND
  sqlstr = "Select  cp.nickName, fta & '-' & ftb as eind from tblCompetitorPools cp"
  sqlstr = sqlstr & " INNER JOIN tblPrediction_Matchresults r ON cp.competitorPoolID = r.competitorpoolID"
  sqlstr = sqlstr & " WHERE cp.poolID =" & thisPool
  sqlstr = sqlstr & " AND r.matchOrder = " & matchOrder
  sqlstr = sqlstr & " AND fta & '-' & ftb = '" & eind & "'"
  sqlstr = sqlstr & " ORDER BY cp.nickname"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  printObj.Print "Eindstand "; Format(rs.RecordCount, formatStr); ":";
'  printObj.Print "Eindstand ("; rs.RecordCount; "):";
  
  printPoolFormResultsBlock rs
  
  rs.Close
  
'  printObj.CurrentY = lastYpos + 50
  printObj.Line (cols(0), printObj.CurrentY)-(printObj.ScaleWidth, printObj.CurrentY)
  
  printObj.CurrentX = 0
  '''TOTO
  savYpos = printObj.CurrentY
  sqlstr = "Select  cp.nickName,  tt from tblCompetitorPools cp"
  sqlstr = sqlstr & " INNER JOIN tblPrediction_Matchresults r ON cp.competitorPoolID = r.competitorpoolID"
  sqlstr = sqlstr & " WHERE cp.poolID =" & thisPool
  sqlstr = sqlstr & " AND r.matchOrder = " & matchOrder
  sqlstr = sqlstr & " AND tt = " & toto
  sqlstr = sqlstr & " ORDER BY cp.nickname"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  printObj.Print "Toto "; Format(rs.RecordCount, formatStr); ":";
  printPoolFormResultsBlock rs
  
  rs.Close
  Set rs = Nothing
  
  printObj.Print
  
End Sub

Sub printPoolFormResultsBlock(rs As ADODB.Recordset)
  Dim cols() As Integer
  Dim colWidth As Integer
  Dim colCount As Integer
  Dim savYpos As Integer
  Dim i As Integer
  colCount = 6
  ReDim cols(colCount)
  cols(0) = printObj.TextWidth("EINDSTANDGOED")
  colWidth = (printObj.ScaleWidth - cols(0)) / colCount
  For i = 1 To colCount
    cols(i) = cols(i - 1) + colWidth
  Next
  i = 0
  'printObj.CurrentY = savYpos
  Do While Not rs.EOF
    If printObj.CurrentY > footerPos - 400 Then
      heading1 = ""
      addNewPage False, False
    End If
    printObj.CurrentX = cols(i)
    printObj.Print rs!nickName;
    rs.MoveNext
    i = i + 1
    If i >= colCount And Not rs.EOF Then
      printObj.Print
      savYpos = printObj.CurrentY
      i = 0
    End If
  Loop
  If rs.RecordCount = 0 Then
    printObj.CurrentX = cols(i)
    printObj.Print "NIEMAND!!"
  Else
    printObj.Print
  End If
End Sub

Sub printFavTopScorers()
Dim aant As Integer
Dim cols() As Integer
Dim sqlstr As String
Dim savy As Integer
Dim favYpos As Integer
Dim savFntgr As Integer
Dim i As Integer
Dim J As Integer
Dim colWidth As Integer
Dim colAant As Integer
colAant = 4
colWidth = printObj.ScaleWidth / colAant
ReDim cols(colAant + 1)
For i = 0 To colAant
    cols(i) = i * colWidth + 50
Next
Set rs = New ADODB.Recordset
cols(5) = printObj.ScaleWidth - 10
sqlstr = "SELECT p.nickName, Count(c.competitorpoolID) AS aantal"
sqlstr = sqlstr & " FROM tblPredictionTopscorers c LEFT JOIN tblPeople p ON c.topscorerplayerID = p.peopleID"
sqlstr = sqlstr & " WHERE c.competitorpoolid In (select competitorpoolid from tblCompetitorPools where poolid= " & thisPool
sqlstr = sqlstr & " ) GROUP BY p.nickname, c.topscorerplayerID"
sqlstr = sqlstr & " ORDER BY Count(c.competitorpoolID) DESC, p.nickname "
rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    rs.MoveLast
End If
aant = rs.RecordCount
i = 1
J = 0
'printObj.Print
favYpos = printObj.CurrentY
printObj.CurrentX = 0
If favYpos > footerPos - Int(aant / 5) * printObj.TextHeight("tekst") - 120 Then
  heading1 = "Topscorers"
  InitPage False, False, 0, False
  favYpos = printObj.CurrentY
Else
  printObj.CurrentY = favYpos
  reportTitle "Topscorers", False, False, favYpos, 0
End If
fontSizing 10
savy = printObj.CurrentY
rs.MoveFirst

Do While Not rs.EOF
    printObj.CurrentX = cols(i - 1)
    If nz(rs!nickName, "") > "" Then
        printObj.Print rs!nickName;
    Else
        printObj.Print "Niet ingevuld";
    End If
    printObj.CurrentX = cols(i) - 500 - printObj.TextWidth(rs!aantal)
    printObj.Print rs!aantal
    J = J + 1
    rs.MoveNext
    If printObj.CurrentY > favYpos Then
        favYpos = printObj.CurrentY
    End If
    If J > aant / colAant Then
        i = i + 1
        J = 0
        printObj.CurrentY = savy
    End If
Loop

rs.Close
Set rs = Nothing

On Error Resume Next
printObj.Line (cols(0) - 50, savy)-(cols(5) - 50, favYpos), lineColor, B
On Error GoTo 0

End Sub
'

Sub printTournamentFav(Plaats As String, col As Integer, rs As ADODB.Recordset, veld As String)
Dim sqlstr As String
Dim yPos As Integer
Dim fntGr As Integer
  yPos = printObj.CurrentY
  fntGr = printObj.Font.Size
  If rs.RecordCount > 0 Then
    rs.MoveFirst
    printObj.FontBold = True
    printObj.CurrentX = col
    printObj.Print Plaats
    printObj.FontBold = False
    Do While Not rs.EOF
      printObj.CurrentX = col + 50
      If nz(rs(veld), 0) = 0 Then
          printObj.Print "Niet ingevuld";
      Else
          printObj.Print getTeamInfo(rs(veld), "teamName", cn); 'getTeamInfo(rs(veld));
      End If
      printObj.CurrentX = col + printObj.TextWidth("123456789012345") - printObj.TextWidth(rs!aantal)
      printObj.Print rs!aantal;
      fontSizing fntGr - 2
      printObj.CurrentY = printObj.CurrentY + 30
      printObj.Print "(" & Format(rs!aantal / getpoolFormCount(cn), "0.0%") & ")"
      printObj.CurrentY = printObj.CurrentY - 30
      fontSizing fntGr
      rs.MoveNext
    Loop
  End If
End Sub

Sub printFavTournamentResult(savy As Integer, savX2 As Integer)
Dim sqlstr As String
Dim rs() As ADODB.Recordset
Dim maxaant As Integer
Dim favXpos As Integer
Dim savX As Integer
Dim aantpos As Integer
Dim startY As Integer
Dim maxFavYpos As Integer
Dim i As Integer
Dim podiumPlaces As Integer
Dim savFntgr As Integer
Dim aantFav As Integer
Dim cols() As Integer
Dim colWidth As Integer
colWidth = printObj.ScaleWidth / 4
'lets start at a new page
  If getTournamentInfo("tournamentThirdPlace", cn) Then
    podiumPlaces = 4
  Else
    podiumPlaces = 2
  End If
  ReDim rs(podiumPlaces) As ADODB.Recordset
  ReDim cols(podiumPlaces + 1)
  For i = 1 To podiumPlaces + 1
    cols(i - 1) = Int((printObj.ScaleWidth / 4) * (i - 1))
    If podiumPlaces = 2 And savX2 = 3 Then
      cols(i - 1) = cols(i) + (printObj.ScaleWidth / 2)
    End If
  Next
  If podiumPlaces = 4 Then savX2 = 0
  cols(podiumPlaces + 1) = printObj.ScaleWidth - 20

  startY = savy
  For i = 1 To podiumPlaces
    sqlstr = "SELECT predictionTeam" & i & ", Count(competitorPoolID) AS aantal"
    sqlstr = sqlstr & " From tblCompetitorPools "
    sqlstr = sqlstr & " WHERE poolid = " & thisPool
    sqlstr = sqlstr & " GROUP BY predictionTeam" & i
    sqlstr = sqlstr & " ORDER BY Count(competitorPoolID) desc"
    Set rs(i) = New ADODB.Recordset
    rs(i).Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs(i).EOF Then
      rs(i).MoveLast
    End If
    favXpos = cols(savX2)
    If maxaant < rs(i).RecordCount Then
      maxaant = rs(i).RecordCount
    End If
  Next
  savFntgr = printObj.FontSize
  'printObj.FontSize = savFntGr
  maxFavYpos = maxaant * printObj.TextHeight("Q") + savy
  printObj.FontSize = savFntgr
  maxFavYpos = maxFavYpos + printObj.TextHeight("Q") + 50
  If maxFavYpos > footerPos - 465 Then
    heading1 = "Favorieten einduitslag"
    addNewPage True, False, podiumPlaces
    'maxFavYpos = printObj.CurrentY
    savy = printObj.CurrentY
    startY = savy
    savFntgr = printObj.FontSize - 2
    printObj.FontSize = savFntgr
    maxFavYpos = maxaant * printObj.TextHeight("Q") + savy
    printObj.FontSize = savFntgr
    maxFavYpos = maxFavYpos + printObj.TextHeight("Q") + 50
  Else
    If savX2 = 3 Then
      reportTitle "Favorieten einduitslag", False, False, savy, savX2 + 1
    Else
      reportTitle "Favorieten einduitslag", False, False, savy, savX2 ' 0 centreert tussenkop
    End If
    fontSizing 10
    savy = printObj.CurrentY
    startY = savy
    savFntgr = printObj.FontSize
    printObj.FontSize = savFntgr - 2
    maxFavYpos = maxaant * printObj.TextHeight("Q") + savy
    printObj.FontSize = savFntgr
    maxFavYpos = maxFavYpos + printObj.TextHeight("Q") + 50
  End If
  For i = 1 To podiumPlaces
    If getPointsFor(Format(i, "0") & "e plaats", cn) Then
      printObj.CurrentY = savy
      printTournamentFav Format(i, "0") & "e plaats", cols(i - 1) + 50, rs(i), "predictionTeam" & i
      On Error Resume Next
      printObj.Line (cols(i - 1), startY)-(cols(i) - 50, maxFavYpos), lineColor, B
      On Error GoTo 0
    End If
    rs(i).Close
    Set rs(i) = Nothing
  Next
  favXpos = 0
  savy = printObj.CurrentY
End Sub

Sub printFavFinals(wedtype As Integer, cols As Integer, koptxt As String, Optional bewaarYpos As Integer, Optional posX As Integer)
Dim sqlstr As String
Dim savX As Integer
Dim savy As Integer
Dim aantpos As Integer
Dim startY As Integer
Dim col() As Integer
Dim i As Integer
Dim J As Integer
Dim team As String
Dim fld As field
Dim maxrows As Integer
Dim maxrows1 As Integer
Dim savMaxRows As Integer
Dim savMaxRows1 As Integer
Dim ttlRows As Integer
Dim maxFinpos As Integer
Dim nwPag As Boolean
Dim finYpos As Integer
Dim favYpos As Integer
Dim favXpos As Integer

Set rs = New ADODB.Recordset

ReDim col(cols + 1) As Integer

    For i = 1 To cols
        col(i) = (i - 1) * printObj.ScaleWidth / cols
    Next
    col(cols + 1) = printObj.ScaleWidth
    savy = printObj.CurrentY
    sqlstr = "Select * from tblTournamentSchedule where tournamentID = " & thisTournament
    sqlstr = sqlstr & " and matchType = " & wedtype
    sqlstr = sqlstr & " ORDER BY matchNumber"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    startY = savy
    'startY = 945

    If rs.RecordCount > 0 Then
        savMaxRows = 0
        maxrows = 0
        rs.MoveFirst
        'bepaal aantal regels dat nodig is eerste team
        Do While Not rs.EOF
            savMaxRows = maxrows + getFavRowCount(rs!matchNumber, "teamNameA", cn)
            If maxrows < savMaxRows Then
                maxrows = savMaxRows
            End If
            rs.MoveNext
        Loop
        rs.MoveFirst
        'bepaal aantal regels dat nodig is voor tweede team
        Do While Not rs.EOF
            savMaxRows = maxrows + getFavRowCount(rs!matchNumber, "teamNameB", cn)
            If maxrows1 < savMaxRows1 Then
                maxrows1 = savMaxRows1
            End If
            rs.MoveNext
        Loop
        ttlRows = maxrows
        If maxrows1 > ttlRows Then ttlRows = maxrows1
        rs.MoveFirst
        If startY + ttlRows * TextHeight("Q") > footerPos - 465 Then  'And wedtype <> 4)
            heading1 = koptxt
            If wedtype = 7 Then
                addNewPage False, False, 2
                maxFavYpos = printObj.CurrentY
                savy = maxFavYpos
                startY = 480
                nwPag = True
            Else
                addNewPage False, False, 20
                maxFavYpos = printObj.CurrentY
                savy = maxFavYpos
                startY = printObj.CurrentY
                nwPag = False
            End If
        Else
            If wedtype = 7 Then
                finYpos = printObj.CurrentY
                reportTitle koptxt, False, False, maxFavYpos, 2
            ElseIf wedtype = 4 Then
                printObj.CurrentY = savy
                If getPointsFor("kleine finale team", cn) + getPointsFor("kleine finale positie", cn) > 0 Then
                    If nwPag Then
                        reportTitle koptxt, False, False, 480, 4
                    Else
                        reportTitle koptxt, False, False, savy, 4
                    End If
                Else
                    reportTitle koptxt, False, False, bewaarYpos, 2
                End If
            Else
                reportTitle koptxt, False, False, maxFavYpos
            End If
            savy = printObj.CurrentY
            startY = savy
        End If

        i = 1
        If wedtype = 4 Then
            i = posX
        End If
        Do While Not rs.EOF
            If i <= cols Then
                printObj.CurrentY = savy
            End If
            fav_finalTeams "teamNameA", "matchTeamA", rs, col(i) ', wedtype = 5
            If maxFavYpos < printObj.CurrentY Then maxFavYpos = printObj.CurrentY
            i = i + 1
            If i <= cols Then
                printObj.CurrentY = savy
            End If
            fav_finalTeams "teamNameB", "matchTeamB", rs, col(i) ', wedtype = 5
            If maxFavYpos < printObj.CurrentY Then maxFavYpos = printObj.CurrentY
            i = i + 1

            If wedtype = 7 And maxFavYpos < printObj.CurrentY Then
                maxFavYpos = printObj.CurrentY
            ElseIf wedtype = 4 Then
                If printObj.CurrentY > maxFavYpos Then
                    maxFavYpos = printObj.CurrentY
                End If
            End If
            maxFavYpos = maxFavYpos + 50
            If i = 5 Then
                printObj.Line (col(1), startY)-(col(3) - 50, maxFavYpos), lineColor, B
                printObj.Line (col(3), startY)-(col(5) - 50, maxFavYpos), lineColor, B
            End If
            If posX = 1 And i = 3 Then
                printObj.Line (col(1), startY)-(col(3) - 50, maxFavYpos), lineColor, B
            End If

            rs.MoveNext
            If i > cols Then
                i = 1
                printObj.CurrentY = maxFavYpos + 50
                savy = printObj.CurrentY
                maxFavYpos = savy
                startY = maxFavYpos
                favYpos = savy
                favXpos = 0
            End If

        Loop

    End If
    rs.Close
    Set rs = Nothing
End Sub
'
Sub fav_finalTeams(teamField As String, cod As String, rs As ADODB.Recordset, col As Integer, Optional morethan1 As Boolean)
'morethan1 :  alleen teams die vaker dan 1 keer worden genoemd
Dim rs1 As ADODB.Recordset
Dim savX As Integer
Dim savy As Integer
Dim aantpos As Integer
Dim sqlstr As String
Dim fntGr As Integer
Dim aantal As Integer
Dim statCol(3) As Integer
  Set rs1 = New ADODB.Recordset
  aantpos = printObj.TextWidth("NIET INGEVULD  1")
  sqlstr = "Select matchOrder, " & teamField & ", count(matchOrder) as ttl"
  sqlstr = sqlstr & " FROM tblPrediction_Finals"
  sqlstr = sqlstr & " WHERE competitorPoolID IN ("
  sqlstr = sqlstr & " SELECT competitorPoolID from tblCompetitorPools WHERE poolid = " & thisPool & ")"
  sqlstr = sqlstr & " GROUP BY matchOrder, " & teamField
  sqlstr = sqlstr & " HAVING matchOrder = " & rs!matchOrder
  sqlstr = sqlstr & " AND " & teamField & " > 0"
  If morethan1 Then sqlstr = sqlstr & " AND " & "count(matchOrder) > 1"
  sqlstr = sqlstr & " ORDER BY count(matchOrder) DESC"
  rs1.Open sqlstr, cn, adOpenStatic, adLockReadOnly
  fontSizing 10
  statCol(0) = col + 20
  statCol(1) = statCol(0) + printObj.TextWidth("#ABDEF")
  statCol(2) = statCol(1) + printObj.TextWidth("TEAMNAAMLAN")
  statCol(3) = statCol(2) + printObj.TextWidth("99")
  printObj.CurrentX = col
  printObj.Print rs(cod) & ": ";
  savX = col + printObj.TextWidth("#ABDEF")
  fntGr = printObj.Font.Size
  Do While Not rs1.EOF
      printObj.CurrentX = statCol(1)
      If nz(rs1(teamField), "") = "" Then
          printObj.Print "Niet ingevuld";
      Else
          printObj.Print getTeamInfo(rs1(teamField), "teamName", cn);
      End If
      printObj.CurrentX = statCol(2) - printObj.TextWidth(rs1!ttl)
      printObj.Print rs1!ttl;
      fontSizing 8
      printObj.CurrentY = printObj.CurrentY + 30
      printObj.CurrentX = statCol(3)
      printObj.Print "(" & Format(rs1!ttl / getpoolFormCount(cn), "0.0%") & ")"
      fontSizing 10
      printObj.CurrentY = printObj.CurrentY - 30
      If maxFavYpos < printObj.CurrentY Then maxFavYpos = printObj.CurrentY
      rs1.MoveNext
  Loop
  rs1.Close
  If morethan1 Then
    'laatste regel toevoegen voor de teams die maar door 1 deelnemer werden gekozen
    sqlstr = "Select matchOrder, " & teamField & ", count(matchOrder) as ttl"
    sqlstr = sqlstr & " FROM tblPrediction_Finals"
    sqlstr = sqlstr & " WHERE competitorPoolID IN ("
    sqlstr = sqlstr & " SELECT competitorPoolID from tblCompetitorPools WHERE poolid = " & thisPool & ")"
    sqlstr = sqlstr & " GROUP BY matchOrder, " & teamField
    sqlstr = sqlstr & " HAVING matchOrder = " & rs!matchOrder
    sqlstr = sqlstr & " AND " & teamField & " > 0"
    sqlstr = sqlstr & " AND " & "count(matchOrder) = 1"
    sqlstr = sqlstr & " ORDER BY count(matchOrder) DESC"
    rs1.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    printObj.CurrentX = statCol(1)
    If rs1.EOF Then
      aantal = 0
    Else
      aantal = rs1.RecordCount
      If aantal > 1 Then
        printObj.Print "" & aantal & " team(s) bij ";
      Else
        printObj.Print "" & aantal & " team bij ";
      End If
      printObj.CurrentX = statCol(2) - printObj.TextWidth(1)
      printObj.Print 1;
      fontSizing 8
      printObj.CurrentY = printObj.CurrentY + 30
      printObj.CurrentX = statCol(3)
      printObj.Print "(" & Format(1 / getpoolFormCount(cn), "0.0%") & ")"
      fontSizing 10
      printObj.CurrentY = printObj.CurrentY - 30
    End If
    rs1.Close
  End If
  Set rs1 = Nothing
End Sub
'

Private Sub printCompetitorPoolForms()
Dim colCnt As Integer
Dim rsPoolForms As ADODB.Recordset
Dim i As Integer
Dim prntCount As Integer 'number of forms printed
Dim lineXpos As Integer
Dim lineYpos As Integer
Dim newLinePos As Integer
Dim matchYpos As Integer
Dim lineHeight As Integer
Dim headHeight  As Integer
Dim TopMargin As Integer
Dim colNum As Integer
Dim colWidth As Integer
Dim colpos() As Integer
Dim formsCnt As Integer 'how many forms on one sheet
Dim formsOnPage As Integer
Dim formHeight 'height of each poolform on the paper
Dim headerHeight As Integer
Dim sqlstr As String
Dim header As String
Dim maxFormsOnPage As Integer
Dim pgNum As Integer
Dim grpMatchAant As Integer

    grpMatchAant = getMatchCount(1, cn)
    rotater.Angle = 0
    
    Set rsPoolForms = New ADODB.Recordset
    
    colNum = 1
    heading1 = "Deelnemers & Voorspellingen"
    InitPage True, False
    printObj.CurrentY = printObj.CurrentY - 50
    headerHeight = printObj.CurrentY
    TopMargin = printObj.CurrentY
    maxFormsOnPage = 2
    If getMatchCount(1, cn) <= 24 Then
        maxFormsOnPage = 3
    End If
    formHeight = (footerPos - TopMargin) / maxFormsOnPage
    fontSizing 8
    lineHeight = printObj.TextHeight("x") '* printRatio
    fontSizing 11
    headHeight = printObj.TextHeight("x") '* printRatio
    fontSizing 12
    formsOnPage = 0
    'get the forms
    sqlstr = "Select * from tblCompetitorPools where poolID =" & thisPool
    sqlstr = sqlstr & " ORDER BY nickName"
    
    With rsPoolForms
      .Open sqlstr, cn, adOpenKeyset, adLockReadOnly
      If .EOF Then
        MsgBox "Geen deelnemers gevonden", vbExclamation + vbOKOnly, "Deelnemerformulieren"
        Exit Sub
      End If
      Do While Not .EOF
        'start printing participants
        If Me.lstCompetitorPools.Selected(.AbsolutePosition - 1) Or Me.optAll Then
'         showInfo True, "Afdrukken deelnemers", !nickName, "Record " & .AbsolutePosition & "/" & .RecordCount
          If formsOnPage = 0 Then
            printObj.CurrentY = TopMargin
          Else
            printObj.CurrentY = formsOnPage * formHeight + TopMargin
          End If
          lineYpos = printObj.CurrentY
          printObj.CurrentX = 30
          printObj.FontBold = True
          fontSizing 16
          printObj.Print
          matchYpos = printObj.CurrentY
          printObj.Line (0, lineYpos)-(printObj.ScaleWidth - 10, matchYpos), &H127419, BF
          printObj.CurrentY = lineYpos
          printObj.ForeColor = vbWhite
          iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
          printObj.CurrentX = 30
          printObj.Print !nickName
          printObj.FontBold = False
          'printObj.CurrentX = 50
          printObj.ForeColor = 1
          'print the blocks
          printCompetitorMatches !competitorPoolID, newLinePos
          printCompetitorGroupStandings !competitorPoolID
          printCompetitorPoolFinals !competitorPoolID
          printCompetitorPoolBottomBlock !competitorPoolID
          'printCompetitorPoolTopScores !competitorPoolID
          'printCompetitorPoolNumbers !competitorPoolID
          
          prntCount = prntCount + 1
          printObj.Line (0, lineYpos)-(printObj.ScaleWidth - 10, newLinePos), , B
        End If 'deeln selected
        
        .MoveNext
        printObj.CurrentX = 0
        If Not .EOF Then
          If Me.lstCompetitorPools.Selected(.AbsolutePosition - 1) Or Me.optAll = True Then
            If formsOnPage = maxFormsOnPage - 1 Then  'next page
              'printObj.Line (0, Helft + 200)-(printObj.ScaleWidth - 10, endEersteDeelnPos + 50), lineColor, B
              formsOnPage = 0
              newLinePos = 0
              'Exit Do
              If Not .EOF Then addNewPage True ' False, False, , False
            Else
              'endEersteDeelnPos = printObj.CurrentY
              If prntCount > 0 Then formsOnPage = formsOnPage + 1
            End If
            printObj.DrawWidth = 1
          End If
        End If
      Loop
    End With
    rsPoolForms.Close
    Set rsPoolForms = Nothing
'    showInfo False
End Sub

Sub printCompetitorPoolBottomBlock(competitorPoolID As Long)
'print the ranking, topscorer and numbers on the bottom block of this poolCompetitor
  Dim sqlstr As String
  Dim i As Integer
  Dim prntTxt As String
  Dim finalStartNr As Integer
  Dim rsFinals As ADODB.Recordset
  Dim colWidth As Integer
  Dim thisCol As Integer
  Dim blockStartPos As Integer
  Dim headerStartPos As Integer 'to remember vertical block header position
  Dim blockEndPos As Integer 'the vertical position after this block
  Dim blockXpos As Integer 'horizontal position bock textlines
  colWidth = printObj.ScaleWidth / 5
  
  Set rs = New ADODB.Recordset
  
  'ranking
  sqlstr = "Select * from tblCompetitorPools WHERE competitorpoolID = " & competitorPoolID
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  printObj.CurrentY = printObj.CurrentY + 30
  headerStartPos = printObj.CurrentY
  printObj.CurrentX = 60
  subHeading "Eindstand"
  blockStartPos = printObj.CurrentY
  thisCol = 0
  blockXpos = 60  'start horizontal  posiition
  blockEndPos = printObj.CurrentY
  For i = 1 To 4
    If i > 2 And Not getTournamentInfo("tournamentThirdPlace", cn) Then
      Exit For
    End If
    printObj.CurrentX = blockXpos
    printObj.Print Format(i, "0: ") & getTeamInfo(rs.Fields("predictionTeam" & Format(i, "0")), "teamName", cn)
  Next
  rs.Close
  If blockEndPos < printObj.CurrentY Then blockEndPos = printObj.CurrentY
  'topscorer
  printObj.CurrentY = headerStartPos
  thisCol = thisCol + 1
  blockXpos = thisCol * colWidth + 60
  sqlstr = "Select * from tblPredictionTopScorers WHERE competitorPoolID = " & competitorPoolID
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  printObj.CurrentX = blockXpos
  subHeading "Topscorer (goals)"
  blockStartPos = printObj.CurrentY
  Do While Not rs.EOF 'maybe in the future more then 1 topscorer
    printObj.CurrentX = blockXpos
    printObj.Print getPlayerInfo(rs!topScorerPlayerID, "nickName", cn) & Format(rs!topScorergoals, "    (##0)")
    rs.MoveNext
  Loop
  rs.Close
  If blockEndPos < printObj.CurrentY Then blockEndPos = printObj.CurrentY
  If blockEndPos < printObj.CurrentY Then blockEndPos = printObj.CurrentY
  'numbers
  printObj.CurrentY = headerStartPos
  thisCol = thisCol + 1
  blockXpos = thisCol * colWidth + 60
  printObj.CurrentX = blockXpos
  subHeading "Overigen"
  blockStartPos = headerStartPos
  blockXpos = blockXpos + printObj.TextWidth("Overigen: ")
  printObj.CurrentY = headerStartPos
  sqlstr = "Select * from tblPrediction_Numbers WHERE competitorPoolID = " & competitorPoolID
   sqlstr = sqlstr & " ORDER BY predictiontypeID DESC"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  i = Int(rs.RecordCount / 2 + 0.5)
  Do While Not rs.EOF
    If rs.AbsolutePosition = i + 1 Then
      blockXpos = blockXpos + colWidth
      printObj.CurrentY = headerStartPos
    End If
    printObj.CurrentX = blockXpos
    prntTxt = getPredictionInfo(rs!predictiontypeid, "pointTypeDescription", cn)
    If LCase(Left(prntTxt, 6)) = "aantal" Then
      prntTxt = Mid(prntTxt, 7)
    End If
    printObj.Print prntTxt & ": " & Format(rs!predictionNumber, "###0")
    rs.MoveNext
  Loop
  rs.Close
  If blockEndPos < printObj.CurrentY Then blockEndPos = printObj.CurrentY
  printObj.Line (0, headerStartPos)-(printObj.ScaleWidth - 10, blockEndPos), , B
  For i = 1 To 2
    printObj.Line (colWidth * i, headerStartPos)-(colWidth * i, blockEndPos)
  Next
  Set rs = Nothing
End Sub

Sub printCompetitorPoolFinals(competitorPoolID As Long)
  Dim sqlstr As String
  Dim i As Integer
  Dim prntTxt As String
  Dim finalStartNr As Integer
  Dim rsFinals As ADODB.Recordset
  Dim colWidth As Integer
  Dim thisCol As Integer
  Dim blockStartPos As Integer
  Dim headerStartPos As Integer 'to remember vertical block header position
  Dim blockEndPos As Integer 'the vertical position after this block
  Dim blockXpos As Integer 'horizontal position bock textlines
  colWidth = printObj.ScaleWidth / 5
  
  sqlstr = "Select * from tblPrediction_Finals WHERE competitorPoolID = " & competitorPoolID
  Set rsFinals = New ADODB.Recordset
  rsFinals.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If rsFinals.EOF Then
    MsgBox "ERRORRRR", vbOKOnly + vbExclamation, "FOUT"
    Exit Sub
  End If
  rsFinals.MoveFirst
  '8th finals
  printObj.CurrentY = printObj.CurrentY + 30
  headerStartPos = printObj.CurrentY
  printObj.CurrentX = 60
  subHeading "Achtste finales"
  blockStartPos = printObj.CurrentY
  finalStartNr = getFirstFinalMatchNumber(cn)
  thisCol = 0
  blockXpos = 60  'start horizontal  posiition
  For i = finalStartNr To finalStartNr + 7
    rsFinals.Find "matchOrder = " & i
    If rsFinals.EOF Then
      MsgBox "ERRORRRR", vbOKOnly + vbExclamation, "FOUT"
      Exit Sub
    End If
    If i = finalStartNr + 4 Then
      thisCol = thisCol + 1
      printObj.CurrentY = blockStartPos
      blockXpos = thisCol * colWidth + 60
    End If
    printObj.CurrentX = blockXpos
    prntTxt = Format(getMatchNumber(rsFinals!matchOrder, cn), "00:")
    prntTxt = prntTxt & getTeamInfo(rsFinals!teamnameA, "teamName", cn)
    prntTxt = prntTxt & "-" & getTeamInfo(rsFinals!teamnameB, "teamName", cn)
    printObj.Print fitText(colWidth, prntTxt)
    rsFinals.MoveFirst
  Next
  blockEndPos = printObj.CurrentY
  'q-finals
  thisCol = thisCol + 1
  printObj.CurrentY = blockStartPos
  blockXpos = thisCol * colWidth + 60
  printObj.CurrentY = headerStartPos
  printObj.CurrentX = blockXpos
  subHeading "Kwart finales"
  printObj.CurrentY = blockStartPos
  For i = finalStartNr + 8 To finalStartNr + 11
    rsFinals.Find "matchOrder = " & i
    If rsFinals.EOF Then
      MsgBox "ERRORRRR", vbOKOnly + vbExclamation, "FOUT"
      Exit Sub
    End If
    printObj.CurrentX = blockXpos
    prntTxt = Format(getMatchNumber(rsFinals!matchOrder, cn), "00:")
    prntTxt = prntTxt & getTeamInfo(rsFinals!teamnameA, "teamName", cn)
    prntTxt = prntTxt & "-" & getTeamInfo(rsFinals!teamnameB, "teamName", cn)
    printObj.Print fitText(colWidth, prntTxt)
    rsFinals.MoveFirst
  Next
  '½-finals
  thisCol = thisCol + 1
  printObj.CurrentY = blockStartPos
  blockXpos = thisCol * colWidth + 60
  printObj.CurrentY = headerStartPos
  printObj.CurrentX = blockXpos
  subHeading "Halve finales"
  'printObj.CurrentY = blockStartPos + 105
  printObj.Print
  For i = finalStartNr + 12 To finalStartNr + 13
    rsFinals.Find "matchOrder = " & i
    If rsFinals.EOF Then
      MsgBox "ERRORRRR", vbOKOnly + vbExclamation, "FOUT"
      Exit Sub
    End If
    printObj.CurrentX = blockXpos
    prntTxt = Format(getMatchNumber(rsFinals!matchOrder, cn), "00:")
    prntTxt = prntTxt & getTeamInfo(rsFinals!teamnameA, "teamName", cn)
    prntTxt = prntTxt & "-" & getTeamInfo(rsFinals!teamnameB, "teamName", cn)
    printObj.Print fitText(colWidth, prntTxt)
'    printObj.Print
    rsFinals.MoveFirst
  Next
  'final and (if played) third place
  thisCol = thisCol + 1
  printObj.CurrentY = blockStartPos
  blockXpos = thisCol * colWidth + 60
  printObj.CurrentY = headerStartPos
  printObj.CurrentX = blockXpos
  If getTournamentInfo("tournamentThirdPlace", cn) Then
    rsFinals.MoveLast
    rsFinals.MovePrevious
    rsFinals.MoveFirst
    subHeading "3e plaats"
    rsFinals.Find "matchOrder = " & i
    If rsFinals.EOF Then
      MsgBox "ERRORRRR", vbOKOnly + vbExclamation, "FOUT"
      Exit Sub
    End If
    printObj.CurrentX = blockXpos
    prntTxt = Format(getMatchNumber(rsFinals!matchOrder, cn), "00:")
    prntTxt = prntTxt & getTeamInfo(rsFinals!teamnameA, "teamName", cn)
    prntTxt = prntTxt & "-" & getTeamInfo(rsFinals!teamnameB, "teamName", cn)
    printObj.Print fitText(colWidth, prntTxt)
    fontSizing 5
    printObj.Print
    fontSizing 9
    printObj.Line (colWidth * 4, printObj.CurrentY - 40)-(printObj.ScaleWidth, printObj.CurrentY - 40)
  End If
  printObj.CurrentX = blockXpos
  subHeading "Finale"
  rsFinals.MoveLast
  printObj.CurrentX = blockXpos
  prntTxt = Format(getMatchNumber(rsFinals!matchOrder, cn), "00:")
  prntTxt = prntTxt & getTeamInfo(rsFinals!teamnameA, "teamName", cn)
  prntTxt = prntTxt & "-" & getTeamInfo(rsFinals!teamnameB, "teamName", cn)
  printObj.Print fitText(colWidth, prntTxt)
  'print boxes around the sections
  'printObj.Line (0, blockStartPos)-(Printer.ScaleWidth, blockStartPos)
  printObj.Line (0, headerStartPos)-(printObj.ScaleWidth - 10, blockEndPos), , B
  For i = 2 To 4
    printObj.Line (colWidth * i, headerStartPos)-(colWidth * i, blockEndPos)
  Next
  
End Sub

Sub printCompetitorGroupStandings(competitorID As Long)
'groepswedstrijden
Dim sqlstr As String
Dim i As Integer
Dim rsGrpStanding As ADODB.Recordset
Dim colCnt As Integer
Dim colWidth As Integer
Dim thisCol As Integer

Dim blockStartPos As Integer
Dim headerStartPos As Integer 'to remember vertical block header position
Dim blockEndPos As Integer 'the vertical position after this block
  colCnt = getTournamentInfo("tournamentGroupCount", cn)
  colWidth = printObj.ScaleWidth / colCnt
  
  Set rsGrpStanding = New ADODB.Recordset
  sqlstr = "Select groupLetter, "
  For i = 1 To 4
    sqlstr = sqlstr & "predictionGroupPosition" & Format(i, "0") & " as p" & Format(i, "0")
    If i < 4 Then sqlstr = sqlstr & ", "
  Next
  sqlstr = sqlstr & " FROM tblPredictionGroupResults WHERE competitorPoolID = " & competitorID
  printObj.CurrentY = printObj.CurrentY + 30
  'heading
  
  headerStartPos = printObj.CurrentY
  printObj.CurrentX = 60
  subHeading "Groepsstanden"
  printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 10, printObj.CurrentY)
  blockStartPos = printObj.CurrentY
  With rsGrpStanding
    .Open sqlstr, cn, adOpenStatic, adLockReadOnly
    thisCol = 0
    Do While Not .EOF
      printObj.CurrentY = blockStartPos
      printObj.CurrentX = thisCol * colWidth + 60
      printObj.FontUnderline = True
      printObj.ForeColor = &H4000&  'dark green
      printObj.Print "Groep " & !groupletter
      printObj.ForeColor = 1
      printObj.FontUnderline = False
      For i = 1 To 4
        printObj.CurrentX = thisCol * colWidth + 60
        printObj.Print getGroupTeamName(!groupletter, i, cn);
        printObj.CurrentX = colWidth * (thisCol + 1) - printObj.TextWidth("123")
        printObj.Print .Fields(i)
      Next
      blockEndPos = printObj.CurrentY 'bottom position of the bock
      printObj.Line (thisCol * colWidth, blockStartPos)-(thisCol * colWidth, blockEndPos), , B
      thisCol = thisCol + 1
      .MoveNext
    Loop
    .Close
    printObj.Line (0, headerStartPos)-(printObj.ScaleWidth - 10, blockEndPos), , B
  End With
End Sub

Sub printCompetitorMatches(competitorID As Long, newLinePos As Integer)
Dim sqlstr As String
Dim pr As String 'what to print
Dim matchCol As Integer
Dim colpos(6) As Integer
Dim nwColumn As Boolean
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim lineXpos As Integer  'hor position for lines
Dim lineYpos As Integer 'vertical position for lines
Dim pos1 As Integer 'for vertical lines in matches block
Dim pos2 As Integer

Dim grpMatchCnt As Integer 'the number of group matches

Dim savdat As String 'to avoid double dates
Dim rsMatches As ADODB.Recordset
Set rsMatches = New ADODB.Recordset
    
    
    grpMatchCnt = getMatchCount(1, cn)
    
    printObj.FillStyle = vbFSTransparent
    fontSizing 9
    'define column positions for the matches
    colpos(0) = 50                                        'date
    colpos(1) = colpos(0) + printObj.TextWidth("99-99:")  'nr
    colpos(2) = colpos(1) + printObj.TextWidth("991")      'match
    colpos(3) = colpos(2) + printObj.TextWidth("WWW-WWW") 'halftime
    colpos(4) = colpos(3) + printObj.TextWidth("11-11")   'fulltime
    colpos(5) = colpos(4) + printObj.TextWidth("11-11")   'toto
    colpos(6) = colpos(5) + printObj.TextWidth("99")   'end of block

  sqlstr = sqlCompetitorMatches(competitorID)
  rsMatches.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  
  fontSizing 10
  printObj.FontBold = True
  printObj.ForeColor = vbBlue
  printObj.CurrentX = 50
  printObj.Print "Groepswedstrijden";
  printObj.CurrentX = printObj.ScaleWidth * 0.75 + 50
  printObj.Print "Finales";
  printObj.ForeColor = 1
  printObj.FontBold = False
  fontSizing 10
  printObj.Print
  printObj.Line (0, printObj.CurrentY - 10)-(printObj.ScaleWidth - 10, printObj.CurrentY + 10), lineColor, B
  lineYpos = printObj.CurrentY + 10
  printObj.CurrentY = lineYpos
  fontSizing 9
  lineXpos = 0
  With rsMatches
    K = 0
    matchCol = 1
    Do While Not .EOF
      printObj.CurrentX = lineXpos + colpos(1) - printObj.TextWidth(!shortMatchDate & ":") - 20
      If savdat <> !shortMatchDate Or printObj.CurrentY = lineYpos Then
        printObj.Print !shortMatchDate;
        savdat = !shortMatchDate
      End If
      pr = !matchNumber & ":"
      printObj.CurrentX = lineXpos + colpos(2) - printObj.TextWidth(pr) - 20
      printObj.Print pr;
      
      pr = !shortMatchDesc
      If !shortMatchDesc = " - " Then
        pr = !shortCodeDesc
      End If
      printObj.CurrentX = lineXpos + colpos(2) + 60
      printObj.Print pr;
      printObj.CurrentX = lineXpos + colpos(3)
      printObj.Print !r1; "-"; !r2;
      printObj.CurrentX = colpos(4) + lineXpos
      printObj.Print !e1; "-"; !e2;
      printObj.CurrentX = lineXpos + colpos(5)
      printObj.Print !tt;
      printObj.Print
      If newLinePos < printObj.CurrentY Then newLinePos = printObj.CurrentY
      .MoveNext
      If grpMatchCnt < 25 Then
        nwColumn = (.AbsolutePosition - 1) Mod (grpMatchCnt / 3) = 0  '= Int(grpWedsAant / 2) Or .AbsolutePosition = grpWedsAant
      Else
        nwColumn = (.AbsolutePosition - 1) Mod 16 = 0
      End If
      nwColumn = nwColumn Or .AbsolutePosition - 1 = grpMatchCnt
      If nwColumn And K < 3 Then
        printObj.CurrentY = lineYpos
        K = K + 1
        If (.AbsolutePosition - 1) = grpMatchCnt Then K = 3
        lineXpos = (printObj.ScaleWidth / 4) * K
      End If
    Loop
    .Close
  End With
  J = UBound(colpos)
  For i = 1 To 4
    pos1 = printObj.ScaleWidth / 4 * (i - 1) + colpos(0)
    printObj.Line (pos1 - 50, lineYpos)-(pos1 - 50, newLinePos + 20)
    For J = 1 To UBound(colpos)
      If J <> 2 Then
      pos1 = printObj.ScaleWidth / 4 * (i - 1) + colpos(J) - 20
      printObj.Line (pos1, lineYpos)-(pos1, newLinePos + 20), , B
      End If
    Next
    pos1 = printObj.ScaleWidth - 20
    printObj.Line (pos1, lineYpos)-(pos1, newLinePos + 20)
  Next
  printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 10, printObj.CurrentY)

  Set rsMatches = Nothing
End Sub

Private Sub btnPrntAllAfterDay_Click()
Dim i As Integer
Dim curWed As Integer
Dim lastMatch As Integer
Dim savdat As Date
Dim msg As String
Dim printTo As Integer
'stand in toernooi
  If adminLogin Then
    MsgBox "Rapporten worden naar Preview gestuurd, niet naar de printer"
    printTo = 1  'preview
  Else
    printTo = 0  'printer
  End If
 ' Me.Hide
  ' Me.Show
  savdat = getMatchInfo(Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex), "matchdate", cn)
  setCombo Me.cmbMatchesPlayed, getHighestDayMatchNr(savdat, cn)
  curWed = Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
  lastMatch = getLastMatchPlayed(cn) 'Returns the matchOrder number
  'Me.upDnToMatch.value = getHighestDayMatchNr(savdat, cn)

  If curWed > 0 Then
    msg = "Dagstanden, grafiek en voorspellingen afgedrukt"
    showInfo True, "Afdrukken", "Stand van zaken in toernooi", "Wedstrijd: " & curWed
    DoEvents
    optPrintDoc_Click 8  ' stand in het toernooi
    btnPrint_Click printTo
    optPrintDoc_Click 9 'deelnemers resultaat wedstrijden
    showInfo True, "Afdrukken", "Deelnemer resultaten", "Dag: " & Format(getMatchInfo(curWed, "matchdate", cn), "d MMM")
    btnPrint_Click printTo
    'stand op punten
    DoEvents
    optPrintDoc_Click 4
    Me.poolFormOrder(1) = True
    If getpoolFormCount(cn) <= 58 Then
      Me.chkCombi = 1
    Else
      Me.chkCombi = 0
    End If
    showInfo True, "Afdrukken", "Stand in de pool", "Wedstrijd: " & curWed
    btnPrint_Click printTo
    'stand alfabetisch
    'Screen.MousePointer = vbHourglass
    DoEvents
    If getpoolFormCount(cn) > 58 Then
      Me.chkCombi = False
      optPrintDoc_Click 4
      Me.poolFormOrder(0) = True
      showInfo True, "Afdrukken", "Stand alfabetisch", "Wedstrijd: " & curWed
      btnPrint_Click printTo
    End If
    
    'positie per wedstrijd
    optPrintDoc_Click 10
    showInfo True, "Afdrukken", "Posities per wedstrijd", "Wedstrijd: " & curWed
    btnPrint_Click printTo

    'punten per wedstrijd alfabetisch
    optPrintDoc_Click 5
    Me.poolFormOrder(0) = True
    showInfo True, "Afdrukken", "Puntenopbouw", "Wedstrijd: " & getMatchNumber(lastMatch, cn)
    Me.optLandscape = True
    toMatch = lastMatch
    btnPrint_Click printTo
    DoEvents
    'punten opbouw alfabetisch
    optPrintDoc_Click 6
    Me.poolFormOrder(0) = True
    Me.optLandscape = True
    showInfo True, "Afdrukken", "Punten per wedstrijd", "Wedstrijd: " & getMatchNumber(lastMatch, cn)
    toMatch = lastMatch
    btnPrint_Click printTo
    DoEvents
    'grafiek alfabetisch
    optPrintDoc_Click 7
    Me.poolFormOrder(0) = True
    showInfo True, "Afdrukken", "Grafiek", "Wedstrijd: " & curWed
    btnPrint_Click printTo
  ''voorspellingen
  End If
  curWed = getLastMatchPlayed(cn)
  If curWed < getMatchCount(0, cn) Then
    savdat = getMatchInfo(curWed + 1, "matchdate", cn)
    For i = curWed + 1 To getMatchCount(0, cn)
      
      If Format(getMatchInfo(i, "matchdate", cn), "d-m-yyyy") = Format(savdat, "d-m-yyyy") Then
        optPrintDoc_Click 3
        Me.upDnForMatch.value = i
        showInfo True, "Afdrukken", "Voorspelling", "Wedstrijd: " & i
        btnPrint_Click printTo
      End If
    Next
  End If
  showInfo False
  Screen.MousePointer = vbDefault
  MsgBox msg, vbOKOnly + vbInformation, "Afdrukken"
  Me.Show
End Sub
'
Sub EindStandAfdrukken()
Dim i As Integer
Dim curWed As Integer
Dim savdat As Date
Me.chkEindstand = 1
Dim msg As String

msg = "Voor alle " & getpoolFormCount(cn) & " deelnemers de toernooi- en de poolstand afdrukken?"
msg = msg & vbNewLine & "(Klik 'Nee' om één afdruk te maken en later te kopiëren)"
If MsgBox(msg, vbYesNo, "Eindstand") = vbYes Then
    Me.upDnCopies = getpoolFormCount(cn)
End If
Me.upDnToMatch.value = getLastMatchPlayed(cn)
'stand in toernooi
showInfo True, "Afdrukken", "Eindstand toernooi", "Wedstrijd: " & Me.upDnToMatch.value
DoEvents
optPrintDoc_Click 8
Me.chkDblSide.value = 0
btnPrint_Click 0
'stand op punten
DoEvents
optPrintDoc_Click 4
Me.poolFormOrder(1) = True
showInfo True, "Afdrukken", "Stand op punten", "Wedstrijd: " & Me.upDnToMatch.value
Me.chkDblSide.value = 0
btnPrint_Click 0


''punten per wedstrijd alfabetisch
'DoEvents
'optPrintDoc_Click 6
'Me.poolFormOrder(0) = True
'Me.chkDblSide.value = 0
'showInfo True, "Afdrukken", "Punten per wedstrijd", "Wedstrijd: " & Me.vscrlTM.value
'btnPrint_Click 0
''punten opbouw alfabetisch
'DoEvents
'optPrintDoc_Click 8
'Me.poolFormOrder(0) = True
'Me.optLandscape = True
'Me.chkDblSide.value = 0
'showInfo True, "Afdrukken", "Puntenopbouw", "Wedstrijd: " & getLastMatchPlayed(cn)
'btnPrint_Click 0
''grafiek alfabetisch
'DoEvents
'optPrintDoc_Click 5
'Me.poolFormOrder(0) = True
'Me.chkDblSide.value = 0
'showInfo True, "Afdrukken", "Grafiek", "Wedstrijd: " & Me.vscrlTM.value
'btnPrint_Click 0
'
'klaar
showInfo False
Screen.MousePointer = vbDefault
MsgBox "Eindstand afgedrukt", vbOKOnly + vbInformation, "Afdrukken"

End Sub
'
Private Sub btnFinalPlayerPrint_Click()
    EindStandAfdrukken
End Sub
'
Private Sub cmbPrinters_Click()
   Dim prntr As Printer

   For Each prntr In Printers
      If cmbPrinters.List(cmbPrinters.ListIndex) = prntr.DeviceName Then
         Set Printer = prntr
      End If
   Next
End Sub
'
Sub btnPrint_Click(Index As Integer)
Dim i As Integer
Dim reportSelect As Integer  'hich report
Dim prntr As Printer
Dim curMatch As Integer
  'Me.Hide
  'check the printer
  For Each prntr In Printers
    If cmbPrinters.List(cmbPrinters.ListIndex) = prntr.DeviceName Then
      Set Printer = prntr
    End If
  Next
  'set the text rotator object
  Set rotater = New rotator
  
  If Me.optPortrait Then
    Printer.Orientation = vbPRORPortrait
  Else
    Printer.Orientation = vbPRORLandscape
  End If
  'index is either : preview(1)  or printer(0)
  If Index = 0 Then  'send to printer
    
    Set printObj = Printer
    'check duplex mode
    If printObj.Duplex <> 0 Then
      If Me.chkDblSide Then
        On Error Resume Next
        If Printer.Orientation = vbPRORPortrait Then
          Printer.Duplex = 2
        Else
          Printer.Duplex = 3
        End If
      Else
        Printer.Duplex = 1
      End If
      On Error GoTo 0
    End If
    Printer.FontTransparent = True
    If Me.upDnCopies = 0 Then Me.upDnCopies = 1
    Printer.Copies = Me.upDnCopies.value
  Else      'send to printPreview
    'instantiate object to printpreview
    Set printPrev = New frmPrintPreview
    Me.Visible = False
    printPrev.Show
    For i = printPrev.printPages.UBound To 1 Step -1
      Unload printPrev.printPages(i)
    Next
    If printPrev.printPages.UBound = 0 Then
        Set printObj = printPrev.printPages(0) 'we are 'printing' to the first page of the control array
    End If
  End If
  
  Set rotater.Device = printObj ' used to print texts in an angle
  
  'which report are we printing
  For i = 0 To 10
      If Me.optPrintDoc(i).value = True Then
          reportSelect = i
          Exit For
      End If
  Next
  If Index = 0 Then
    write2Log "Afdrukken op printer rapport nr:" & reportSelect
  Else
    write2Log "Afdrukvoorbeeld rapport nr:" & reportSelect
  End If
  DoEvents
  Select Case reportSelect
  Case 0
      printPoolForm
  Case 1
      printCompetitorPoolForms
  Case 2
      'Favorieten
      printFavourites
  Case 3
      'voorspellingen voor wedstrijd
      printMatchPredictions Me.upDnForMatch
  Case 4
      'Stand in pool
      printPoolStandings Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
  Case 5
      'samenvatting stand
      printPoolPoints Me.poolFormOrder(0)
  Case 6
      'punten per wedstrijd
      printPoolPointsPerMatch
  Case 7
      printSkyline
  Case 8
      'toernooi stand
      printTournamentStandings Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
  Case 9
      'Dagresultaat
      printDailyResults Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
  Case 10
      'plaats per dag
      printPlaceAfterDay Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.ListIndex)
  End Select

  'Picture1.Visible = True
  DoEvents
  If Index = 0 Then
    'send it  to the printer
      Printer.EndDoc
  Else
    With printPrev
      'show the first page
      .pageContent.PaintPicture printObj.Image, 0, 0, printObj.width, printObj.Height
    End With
  End If
  'release printObj
  Set printObj = Nothing
  'Me.Show

End Sub


Sub printPlaceAfterDay(afterMatch As Integer)
'print a table with names (vertical) and days (horizontal) to sho posittion after match afterMatch
Dim leftMargin As Integer 'where does the table start (largest competitor name)
Dim colWidth As Integer 'width of the columns per day
Dim tableWidth As Integer
Dim rightMargin As Integer ' leave some room for extra info
Dim row As Integer
Dim col As Integer
Dim dayCount As Integer
Dim sqlstr As String
Dim topYpos As Integer
Dim tableYbottom As Integer
Dim tableYpos As Integer
Dim plts As Integer
Dim plStr As String
Dim nameStr As String
Dim rsDays As ADODB.Recordset
Set rsDays = New ADODB.Recordset

  Set rs = New ADODB.Recordset
'  MsgBox "Sorry deze doet het nog niet", vbOKOnly + vbInformation, "Plaats per dag"
'  Exit Sub
  heading1 = "Positie in de pool per wedstrijd na " & toMatch & "e wedstrijd " & getMatchDescription(toMatch, cn, True)
  If Me.chkEindstand Then heading1 = "Positie in de pool per wedstrijd"
  InitPage True, False
  dayCount = getTournamentDayCount(cn)
  leftMargin = printObj.TextWidth(getLongestNickName(cn))
  rightMargin = leftMargin / 2
  tableWidth = printObj.ScaleWidth - leftMargin - rightMargin
  colWidth = tableWidth / dayCount
  horline 0
  topYpos = printObj.CurrentY
  sqlstr = "Select matchDate, last(matchOrder) as matchOrder "
  sqlstr = sqlstr & " from tblTournamentSchedule"
  sqlstr = sqlstr & " WHERE tournamentid = " & thisTournament
  sqlstr = sqlstr & " GROUP BY matchdate"
  sqlstr = sqlstr & " ORDER BY matchdate"
  rsDays.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  fontSizing 9
  printObj.CurrentX = leftMargin + (colWidth - printObj.TextHeight("WW")) / 2
  Do While Not rsDays.EOF
    fontSizing 10
    rotater.Angle = 90
    printObj.CurrentY = topYpos + printObj.TextWidth("wo 14 dec") + 10
    rotater.PrintText Format(rsDays!matchDate, "ddd d MMM")
    printObj.CurrentX = printObj.CurrentX + colWidth '+ printObj.TextHeight("WW")
    rsDays.MoveNext
  Loop
  rotater.Angle = 0
  printObj.CurrentY = printObj.CurrentY + 50
  tableYpos = printObj.CurrentY
  horline 0
  
  sqlstr = "Select c.nickname, p.matchOrder, p.positionTotal, p.positionDay  "
  sqlstr = sqlstr & " from tblCompetitorPools c INNER JOIN tblCompetitorPoints p "
  sqlstr = sqlstr & " ON c.competitorpoolid = p.competitorpoolid "
  sqlstr = sqlstr & " WHERE p.matchOrder IN ( "
  sqlstr = sqlstr & " SELECT last(t.matchOrder) as matchorder from tblTournamentSchedule t"
  sqlstr = sqlstr & " WHERE t.tournamentid = " & thisTournament
  sqlstr = sqlstr & " GROUP BY t.matchDate"
  sqlstr = sqlstr & " ORDER BY t.matchDate"
  sqlstr = sqlstr & " ) AND c.poolid =  " & thisPool
  sqlstr = sqlstr & " GROUP BY c.nickname, p.matchorder, p.positionTotal, p.positionDay   "
  sqlstr = sqlstr & " ORDER BY c.nickname, p.matchorder "
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  fontSizing 10.5
  nameStr = ""
  Do While Not rs.EOF
    nameStr = rs!nickName
    printObj.CurrentX = 30
    printObj.FontBold = False
    printObj.ForeColor = vbBlack
    printObj.Print rs!nickName;
    For col = 1 To dayCount
        plts = nz(rs!positionTotal, 0)
        If plts > 0 Then
          plStr = Format(plts, 0)
          printObj.ForeColor = vbBlack
          printObj.FontBold = False
  '        printObj.BackColor = vbWhite
          printObj.CurrentX = leftMargin + ((col * colWidth) - (colWidth + printObj.TextWidth(plStr)) / 2)
          If plts = 1 Then
            printObj.ForeColor = vbBlue
  '          printObj.BackColor = vbYellow
            printObj.FontBold = True
          End If
          If plts = getpoolFormCount(cn) Then
            printObj.ForeColor = vbRed
          End If
          If rs!positionDay = 1 Then
            printObj.ForeColor = &H8000&
            printObj.FontBold = True
          End If
          printObj.Print plStr;
          rs.MoveNext
        End If
        'If rs!nickName = "Wolf" Then Stop
        If Not rs.EOF Then
          If nameStr <> rs!nickName Then
            Exit For
          End If
        Else
          Exit Do
        End If
    Next
    
    printObj.Print
    horline 0
    'rs.MoveNext
  Loop
  printObj.Print
  horline 0
  tableYbottom = printObj.CurrentY
  'vertical lines
  printObj.Line (leftMargin, topYpos)-(leftMargin, tableYbottom)
  For col = 1 To dayCount
    printObj.Line (leftMargin + (col * colWidth), topYpos)-(leftMargin + (col * colWidth), tableYbottom)
  Next
End Sub

Sub printTournamentStandings(toMatch As Integer)
'Dim kopje As String
'    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool - Stand van zaken"
'    kopje = Format(GetWedInfo(toMatch, "datum"), "dddd d mmmm") & ": "
'    kopje = kopje & GetWedInfo(toMatch, "naam1") & " vs " & GetWedInfo(toMatch, "naam2")
    heading1 = "Stand van zaken na " & toMatch & "e wedstrijd " & getMatchDescription(toMatch, cn, True)
    If Me.chkEindstand Then heading1 = "Toernooi uitslagen, statistieken en eindstand"
    InitPage True, False
    printMatchResults
    printGroupStanding
    printTnFinals
    printTopScorers

    printStatistics
    printObj.Print
    printPoolPrizes True

End Sub
'
Sub printTopScorers()
Dim sqlstr As String
Dim rs As New ADODB.Recordset
Dim rsED As New ADODB.Recordset 'voor de eigen doelpunten
Dim i As Integer
Dim grps As Integer
Dim colNu As Integer
Dim numpos As Integer
Dim datPos As Integer
Dim wedPos As Integer
Dim uitslPos As Integer
Dim newYpos As Integer
Dim yPos As Integer
Dim edYpos As Integer
Dim aantpos As Integer
Dim col(5) As Integer
    col(0) = 50
    col(1) = (printObj.ScaleWidth - col(0)) / 5
    col(2) = (printObj.ScaleWidth - col(0)) / 5 * 2
    col(3) = (printObj.ScaleWidth - col(0)) / 5 * 3
    col(4) = (printObj.ScaleWidth - col(0)) / 5 * 4
    col(5) = (printObj.ScaleWidth - col(0))
    aantpos = (printObj.ScaleWidth - col(0)) / 4
    'goals & penalties
    sqlstr = "SELECT tblPeople.nickname as nickname, Count(tblPeople.peopleID) AS cnt, tblTeamNames.teamShortName"
    sqlstr = sqlstr & " FROM ((tblMatchEvents INNER JOIN tblPeople ON tblMatchEvents.playerId = tblPeople.peopleID) "
    sqlstr = sqlstr & " INNER JOIN tblTeamPlayers ON (tblMatchEvents.tournamentID = tblTeamPlayers.tournamentID) AND (tblPeople.peopleID = tblTeamPlayers.playerID)) "
    sqlstr = sqlstr & " INNER JOIN tblTeamNames ON tblTeamPlayers.teamID = tblTeamNames.teamNameID"
    sqlstr = sqlstr & " WHERE (((tblMatchEvents.eventId) <= 2) AND ((tblMatchEvents.tournamentId) = " & thisTournament & "))"
    sqlstr = sqlstr & " GROUP BY tblMatchEvents.tournamentId, tblPeople.nickname, tblTeamNames.teamShortName"
    sqlstr = sqlstr & " ORDER BY Count(tblPeople.peopleID) DESC, tblPeople.nickname"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    'own goals
    sqlstr = "SELECT tblPeople.nickname as nickname  , Count(tblPeople.peopleID) AS cnt, tblTeamNames.teamShortName"
    sqlstr = sqlstr & " FROM ((tblMatchEvents INNER JOIN tblPeople ON tblMatchEvents.playerId = tblPeople.peopleID) "
    sqlstr = sqlstr & " INNER JOIN tblTeamPlayers ON (tblMatchEvents.tournamentID = tblTeamPlayers.tournamentID) AND (tblPeople.peopleID = tblTeamPlayers.playerID)) "
    sqlstr = sqlstr & " INNER JOIN tblTeamNames ON tblTeamPlayers.teamID = tblTeamNames.teamNameID"
    sqlstr = sqlstr & " WHERE (((tblMatchEvents.eventId) = 3) AND ((tblMatchEvents.tournamentId) = " & thisTournament & "))"
    sqlstr = sqlstr & " GROUP BY tblMatchEvents.tournamentId, tblPeople.nickname, tblTeamNames.teamShortName"
    sqlstr = sqlstr & " ORDER BY Count(tblPeople.peopleID) DESC, tblPeople.nickname"
    rsED.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If Not rsED.EOF Then
        rsED.MoveLast
        rsED.MoveFirst
    End If
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    If rs.RecordCount > 0 Then
        fontSizing 12
        printObj.ForeColor = vbBlue
        printObj.FontBold = True
        If toMatch <> getMatchCount(0, cn) Then
          printObj.Print "Topscorers tot nu toe: "
        Else
          printObj.Print "Topscorers"
        End If
        yPos = printObj.CurrentY
        printObj.FontBold = False
        printObj.ForeColor = 1
        fontSizing 8
        Do While Not rs.EOF
'            i = i + 1
'            printObj.CurrentX = col(colNu)
'            printObj.Print Left(rs!nickName, 12) & " (" & LCase(rs!teamshortname) & ")";
'            printObj.CurrentX = col(colNu) + aantpos - printObj.TextWidth("1234567890")
'            printObj.Print rs!cnt
'            newYpos = printObj.CurrentY
'
'
'            rs.MoveNext
'            If colNu = 5 Then
'              colNu = 0
'              printObj.Print
'            End If
'
'            If i = Int(rs.RecordCount / 5) And Not rs.EOF Then
'                i = 0
'                colNu = colNu + 1
'                printObj.CurrentY = yPos
'            End If
            printObj.CurrentX = col(colNu)
            printObj.Print Left(rs!nickName, 20) & " (" & LCase(rs!teamshortname) & ")";
            printObj.CurrentX = col(colNu) + aantpos - printObj.TextWidth("1234567890")
            printObj.Print rs!cnt;

            rs.MoveNext
            colNu = colNu + 1
            If colNu = 5 Then
              colNu = 0
              printObj.Print
            End If
'            If i = Int((rs.RecordCount) / 5) + 1 And Not rs.EOF Then
'                i = 0
'                colNu = colNu + 1
'                newYpos = printObj.CurrentY
'                printObj.CurrentY = yPos
'            End If

        Loop
        printObj.Print
        newYpos = printObj.CurrentY

        If rsED.RecordCount > 0 Then
            printObj.CurrentY = newYpos
            printObj.ForeColor = vbBlue
            printObj.FontBold = True
            i = 0
            colNu = 0
            printObj.CurrentX = col(colNu)
            printObj.Print "Eigen doelpunten:"
            edYpos = printObj.CurrentY
            printObj.FontBold = False
            printObj.ForeColor = 1
            Do While Not rsED.EOF
                i = i + 1
                printObj.CurrentX = col(colNu)
                printObj.Print rsED!nickName & " (" & LCase(rsED!teamshortname) & ")";
                printObj.CurrentX = col(colNu) + aantpos - printObj.TextWidth("1234567890")
                printObj.Print rsED!cnt;


                rsED.MoveNext
                colNu = colNu + 1
                If colNu = 5 Then
                  colNu = 0
                  printObj.Print
                End If
'                If i = Int(rsED.RecordCount / 5) + 1 And Not rsED.EOF Then
'                    i = 0
'                    colNu = colNu + 1
'                    newYpos = printObj.CurrentY
'                    printObj.CurrentY = edYpos
'                End If
            Loop
            rsED.Close
            printObj.Print
            newYpos = printObj.CurrentY
        End If
        rs.Close
        On Error Resume Next
        printObj.Line (0, yPos)-(printObj.ScaleWidth - 50, newYpos), lineColor, B
        On Error GoTo 0
        printObj.CurrentY = newYpos
        printObj.Print
    End If
End Sub
'
Sub printStatistics()
Dim yPos As Integer
Dim prStr As String
Dim col(6) As Integer
Dim pntFormat As String
Dim dp As Integer
Dim gelijk As Integer
Dim gele As Integer
Dim rode As Integer
Dim pens As Integer
Dim eigdp As Integer
Dim wd As Integer
wd = toMatch 'getMatchNumber(toMatch, cn)
  dp = getEventCount(wd, doelp, cn) + getEventCount(wd, penalty, cn) + getEventCount(wd, eigdoelp, cn)
  eigdp = getEventCount(wd, eigdoelp, cn)
  pens = getEventCount(wd, penalty, cn) + getEventCount(wd, penaltyMis, cn)
  gele = getEventCount(wd, geel, cn)
  rode = getEventCount(wd, rood, cn)
  gelijk = getDrawCount(wd, cn)

    pntFormat = "0"
    col(0) = 50
    col(1) = (printObj.ScaleWidth - col(0)) / 6
    col(2) = (printObj.ScaleWidth - col(0)) / 6 * 2
    col(3) = (printObj.ScaleWidth - col(0)) / 6 * 3
    col(4) = (printObj.ScaleWidth - col(0)) / 6 * 4
    col(5) = (printObj.ScaleWidth - col(0)) / 6 * 5
    col(6) = printObj.ScaleWidth - 50
    fontSizing 12
    printObj.ForeColor = vbBlue
    printObj.FontBold = True
    printObj.Print "Statistieken"
    yPos = printObj.CurrentY
    printObj.FontBold = False
    printObj.ForeColor = 1
    fontSizing 10
    printObj.CurrentX = col(0)
    prStr = "Doelpunten: " & Format(dp, pntFormat)
    printObj.Print prStr;
    printObj.CurrentX = col(1)
    prStr = "Penalties: " & Format(pens, pntFormat)
    printObj.Print prStr;
    printObj.CurrentX = col(2)
    prStr = "Gele kaarten: " & Format(gele, pntFormat)
    printObj.Print prStr;
    printObj.CurrentX = col(3)
    prStr = "Rode kaarten: " & Format(rode, pntFormat)
    printObj.Print prStr;
    printObj.CurrentX = col(4)
    prStr = "Gelijke spelen: " & Format(gelijk, pntFormat)
    printObj.Print prStr;
    printObj.CurrentX = col(5)
    prStr = "Eigen doelpunten: " & Format(eigdp, pntFormat)
    printObj.Print prStr
    printObj.ForeColor = vbBlue
    printObj.FontItalic = True
    centerText getpoolFormCount(cn) & " deelnemers aan de pool"
    printObj.Print
    printObj.FontItalic = False
    printObj.ForeColor = 1
    On Error Resume Next
    printObj.Line (0, yPos)-(col(6), printObj.CurrentY), lineColor, B
    On Error GoTo 0
End Sub
'
Sub printTnFinals()
Dim sqlstr As String
Dim savdat As Date
Dim rs As New ADODB.Recordset
Dim rsUitsl As New ADODB.Recordset
Dim i As Integer
Dim grps As Integer
Dim col(5) As Integer
Dim currentCol As Integer
Dim matchCol(5) As Integer
Dim teams() As String
Dim winners() As Long
Dim txtStr As String
Dim newYpos As Integer
Dim yPos As Integer
Dim topYpos As Integer
Dim mType As Integer
Dim uitsl As String
Dim colNr As Integer
Dim grpAant As Integer
  grpAant = getTournamentInfo("tournamentGroupCount", cn)
  col(0) = 20
  col(1) = printObj.ScaleWidth / 3 + col(0)
  col(2) = printObj.ScaleWidth / 3 * 2 + col(0)
  col(3) = printObj.ScaleWidth
  col(4) = printObj.ScaleWidth / 6 + col(0)
  col(5) = printObj.ScaleWidth / 2 + col(0)
  sqlstr = "Select * from tblTournamentSchedule "
  sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchtype <> 1"
  sqlstr = sqlstr & " order by matchOrder"
  rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
  printObj.FontBold = True
  fontSizing 12
  printObj.ForeColor = vbBlue
  printObj.Print "Finales"
  topYpos = printObj.CurrentY
  colNr = 0
  printObj.CurrentX = col(colNr)
  fontSizing 10
  If grpAant > 4 Then
    printObj.Print "Achtste finales";
    colNr = colNr + 1
    printObj.CurrentX = col(colNr)
  End If
  printObj.Print "Kwart finales";
  colNr = colNr + 1
  printObj.CurrentX = col(colNr)
  printObj.Print "Halve finales";
  If colNr < 2 Then
    colNr = colNr + 1
    printObj.CurrentX = col(colNr)
    printObj.Print "Finale";
  End If
  yPos = printObj.CurrentY
  printObj.FontBold = False
  printObj.ForeColor = 1
  fontSizing 8
  matchCol(0) = printObj.TextWidth("00")
  matchCol(1) = matchCol(0) + printObj.TextWidth("0")
  matchCol(2) = matchCol(1) + printObj.TextWidth("wo 29 juni ")
  matchCol(3) = matchCol(2) + printObj.TextWidth(" 20u:")
  matchCol(4) = matchCol(3) + printObj.TextWidth("MWWl")
  matchCol(5) = col(1) - printObj.TextWidth("0-0(0-0)nl:0-0,zwi wnsXX")
  printObj.Print
  yPos = printObj.CurrentY
  Do While Not rs.EOF

    mType = rs!matchType
    teams = getScheduleTeamNames(rs!matchOrder, cn, True)
    Select Case mType
    Case 5 'achtste
      If grpAant > 4 Then
          currentCol = 0
      End If
    Case 2 'KwartFinale
      If grpAant > 4 Then
          currentCol = 1
      Else
          currentCol = 0
      End If
    Case 4 'Finale
      currentCol = 2
      If grpAant <= 4 Then
          printObj.CurrentY = yPos
      End If
    Case Else
      If grpAant > 4 Then
          currentCol = 2
      Else
          currentCol = 1
      End If
    End Select
    printObj.CurrentX = col(currentCol) + matchCol(0) - printObj.TextWidth(Format(rs!matchNumber, "0"))
    printObj.Print Format(rs!matchNumber, "0");
    printObj.CurrentX = col(currentCol) + matchCol(1)
    If savdat <> rs!matchDate Then
      printObj.Print Format(rs!matchDate, "ddd d mmm");
      savdat = rs!matchDate
    End If
    printObj.CurrentX = col(currentCol) + matchCol(2)
    printObj.Print Format(rs!matchtime, "HH\u"); " ";
    printObj.CurrentX = col(currentCol) + matchCol(3)
    printObj.Print teams(0);
    'printObj.CurrentX = col(currentCol) + matchCol(4)
    printObj.Print " - "; teams(1);
    printObj.CurrentX = col(currentCol) + matchCol(5)
    If matchPlayed(rs!matchOrder, cn) Then
      printObj.Print getMatchResultPartStr(rs!matchOrder, 2, cn)
      If rs!matchOrder = getMatchCount(0, cn) Then
        printObj.CurrentY = printObj.CurrentY + 50
        printObj.Line (col(currentCol), printObj.CurrentY)-(printObj.ScaleWidth, printObj.CurrentY)
        printObj.CurrentY = printObj.CurrentY + 50
        winners = getWinners(False, cn)
        printObj.CurrentX = col(currentCol) ' + matchCol(5) - printObj.TextWidth("Kampioen " & txtStr)
        fontSizing 12
        printObj.ForeColor = vbBlue
        printObj.FontBold = True
        printObj.Print "Eindstand";
        printObj.ForeColor = 1
        printObj.FontBold = False
        txtStr = getTeamInfo(CInt(winners(1)), "TeamName", cn)
        printObj.CurrentX = col(currentCol) + printObj.TextWidth("Eindstand: ")
        printObj.Print "1: ";
        printObj.FontBold = True
        printObj.Print txtStr
        printObj.FontBold = False
        printObj.CurrentX = col(currentCol) + printObj.TextWidth("Eindstand: ")
        txtStr = getTeamInfo(CInt(winners(2)), "TeamName", cn)
        
        printObj.Print "2: " & txtStr;
        
        'printObj.Print " kampioen"

      End If
    Else
      printObj.Print
    End If
    rs.MoveNext
    If Not rs.EOF Then
      If rs!matchType <> mType Then
        If newYpos < printObj.CurrentY Then
          newYpos = printObj.CurrentY
        End If
        If rs!matchType <> 7 And rs!matchType <> 4 Then '7 = kleine/ 4 = grote finale
          printObj.CurrentY = yPos
        Else
          printObj.FontBold = True
          fontSizing 12
          printObj.ForeColor = vbBlue
          printObj.CurrentX = col(2)
          If rs!matchType = 7 Then
            printObj.Print "Derde plaats"
          ElseIf grpAant > 4 Then
            printObj.CurrentX = col(2)
            printObj.Print "Finale"
          End If
          printObj.FontBold = False
          printObj.ForeColor = 1
          fontSizing 8
        End If
      End If
    End If

  Loop
  On Error Resume Next
  printObj.Line (col(0) - 20, topYpos)-(col(1) - 50, newYpos), lineColor, B
  printObj.Line (col(1) - 20, topYpos)-(col(2) - 50, newYpos), lineColor, B
  printObj.Line (col(2) - 20, topYpos)-(col(3) - 50, newYpos), lineColor, B
  On Error GoTo 0
  printObj.CurrentY = newYpos
  printObj.Print
End Sub
'
'
Sub printGroupStanding()
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim i As Integer
Dim grps As Integer
Dim grpCnt As Integer
Dim col() As Integer
Dim currentCol As Integer
Dim grpCol(7) As Integer
Dim pos As Integer 'de positie van het team in de groep
Dim teamName As String
Dim yPos As Integer
Dim savYpos As Integer
Set rs = New ADODB.Recordset
  grpCnt = getTournamentInfo("tournamentGroupCount", cn) / 2
  ReDim col(grpCnt)
  If UBound(col) = 3 Then
    col(0) = 250
  Else
    col(0) = 10
  End If
  For i = 1 To UBound(col)
    col(i) = col(i - 1) + printObj.ScaleWidth / grpCnt
  Next
    
    printObj.FontBold = True
    fontSizing 12
    printObj.ForeColor = vbBlue
    printObj.Print "Groepstanden"
    savYpos = printObj.CurrentY
    yPos = printObj.CurrentY + 20
    printObj.FontBold = False
    printObj.ForeColor = 1
    fontSizing 8
    grpCol(0) = 40
    grpCol(1) = grpCol(0) + printObj.TextWidth("1234567890123")
    grpCol(2) = grpCol(1) + printObj.TextWidth("000")
    grpCol(3) = grpCol(2) + printObj.TextWidth("000")
    grpCol(4) = grpCol(3) + printObj.TextWidth("000")
    grpCol(5) = grpCol(4) + printObj.TextWidth("000")
    grpCol(6) = grpCol(5) + printObj.TextWidth("000")
    grpCol(7) = grpCol(6) + printObj.TextWidth("000")


    grps = getTournamentInfo("tournamentGroupCount", cn)
    currentCol = 0
    For i = 1 To grps
        printObj.CurrentY = yPos
        sqlstr = "Select * from tblGroupLayout"
        sqlstr = sqlstr & " Where tournamentID  = " & thisTournament
        sqlstr = sqlstr & " AND groupLetter = '" & Chr(i + 64) & "'"
        sqlstr = sqlstr & " order by teamPoints DESC, mPl, teamPosition, groupPlace"
        rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
        printObj.CurrentX = col(currentCol) + grpCol(0)
        printObj.Print "groep " & Chr(i + 64);
        printObj.CurrentX = col(currentCol) + grpCol(1)
        printObj.Print "sp";
        printObj.CurrentX = col(currentCol) + grpCol(2)
        printObj.Print "W";
        printObj.CurrentX = col(currentCol) + grpCol(3)
        printObj.Print "V";
        printObj.CurrentX = col(currentCol) + grpCol(4)
        printObj.Print "G";
        printObj.CurrentX = col(currentCol) + grpCol(5)
        printObj.Print "P";
        printObj.CurrentX = col(currentCol) + grpCol(6)
        printObj.Print "v-t"
        Do While Not rs.EOF
            pos = pos + 1
            teamName = getTeamInfo(rs!teamID, "teamName", cn)
            printObj.CurrentX = col(currentCol) + grpCol(0)
            If rs!teamposition <> 0 Then
                printObj.Print Format(rs!teamposition, "0"); ". "; teamName;
            Else
                printObj.Print Format(pos, "0"); ". "; teamName;
            End If
            printObj.CurrentX = col(currentCol) + grpCol(1)
            printObj.Print Format(rs!mPl, "0");
            printObj.CurrentX = col(currentCol) + grpCol(2)
            printObj.Print Format(rs!mWon, "0");
            printObj.CurrentX = col(currentCol) + grpCol(3)
            printObj.Print Format(rs!mLost, "0");
            printObj.CurrentX = col(currentCol) + grpCol(4)
            printObj.Print Format(rs!mDraw, "0");
            printObj.CurrentX = col(currentCol) + grpCol(5)
            printObj.Print Format(rs!teamPoints, "0");
            printObj.CurrentX = col(currentCol) + grpCol(6)
            printObj.Print Format(rs!mScored, "0"); "-"; Format(rs!mAgainst, "0")
            rs.MoveNext
        Loop
        printObj.Line (col(currentCol), yPos)-(col(currentCol) + grpCol(7) + 50, printObj.CurrentY), lineColor, B
        currentCol = currentCol + 1
        If currentCol >= 3 Then
          If currentCol = grpCnt Then
            currentCol = 0
            yPos = printObj.CurrentY + 50
          End If
        End If
        pos = 0
        rs.Close
    Next
    printObj.Line (0, savYpos - 10)-(printObj.ScaleWidth - 10, printObj.CurrentY + 30), lineColor, B
    printObj.Print
End Sub
'
Sub printMatchResults()
Dim sqlstr As String
Dim rs As New ADODB.Recordset
Dim rsUitsl As New ADODB.Recordset
Dim i As Integer
Dim grps As Integer
Dim col(3) As Integer
Dim currentCol As Integer
Dim matchCol(4)
Dim newYpos As Integer
Dim teamNames() As String
Dim yPos As Integer
Dim savdat As Date

Dim thisMatch As Integer
    col(0) = 0
    col(1) = printObj.ScaleWidth / 3
    col(2) = printObj.ScaleWidth / 3 * 2
    col(3) = printObj.ScaleWidth
    sqlstr = "Select * from tblTournamentSchedule "
    sqlstr = sqlstr & " WHERE tournamentID= " & thisTournament
    sqlstr = sqlstr & " AND matchType = 1"
    sqlstr = sqlstr & " order by matchOrder"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    printObj.FontBold = True
    fontSizing 12
    printObj.ForeColor = vbBlue
    printObj.Print "Groepswedstrijden"
    yPos = printObj.CurrentY
    printObj.FontBold = False
    fontSizing 8
    printObj.ForeColor = 1
    'position of match columns
    matchCol(0) = printObj.TextWidth("0")
    matchCol(1) = matchCol(0) + printObj.TextWidth("wo 29 ")
    matchCol(2) = matchCol(1) + printObj.TextWidth(" 20u : ")
    matchCol(3) = matchCol(2) + printObj.TextWidth("(22) ")
    matchCol(4) = col(1) - printObj.TextWidth("0-0 (0-0)")
    Do While Not rs.EOF
        thisMatch = rs!matchNumber
        teamNames = getScheduleTeamNames(rs!matchOrder, cn)
        i = i + 1
        printObj.CurrentX = col(currentCol) + matchCol(0)
        If savdat <> rs!matchDate Then
          printObj.Print Format(rs!matchDate, "ddd d ");
          savdat = rs!matchDate
        End If
        printObj.CurrentX = col(currentCol) + matchCol(1)
        printObj.Print " "; Format(rs!matchtime, "HH\u"); " : ";
        'printObj.CurrentX = col(currentCol) + matchCol(3) - printObj.TextWidth(Format(thisMatch, "\(0\)\ "))
        printObj.CurrentX = col(currentCol) + matchCol(2)
        printObj.Print teamNames(0) & " - " & teamNames(1) & "  ";
        printObj.Print Format(rs!matchNumber, "\(0\)\ ");
        printObj.CurrentX = col(currentCol) + matchCol(4)
        If matchPlayed(rs!matchOrder, cn) Then
            printObj.Print getMatchresultStr(rs!matchOrder, True, cn)
        Else
            printObj.Print
        End If
        rs.MoveNext
        If i = rs.RecordCount / 3 Then
            If newYpos < printObj.CurrentY Then
                newYpos = printObj.CurrentY
            End If
            i = 0
            printObj.CurrentY = yPos
            currentCol = currentCol + 1
        End If
    Loop
    printObj.Line (10, yPos)-(printObj.ScaleWidth - 20, newYpos), lineColor, B
    printObj.Line (col(1), yPos)-(col(1), newYpos)
    printObj.Line (col(2), yPos)-(col(2), newYpos)
    printObj.Print
End Sub
'
Private Sub addNewPage(poolName As Boolean, Optional headerBgFill As Boolean, Optional headerPos As Integer, Optional noFooter As Boolean, Optional extraLarge As Boolean, Optional noSubHead As Boolean)
  If TypeOf printObj Is Printer Then
    Printer.NewPage
  Else
    Load printPrev.printPages(printPrev.printPages.UBound + 1)
    printPrev.printPages(printPrev.printPages.UBound).Visible = True
    printPrev.printPages(printPrev.printPages.UBound).AutoRedraw = True
    Set printObj = printPrev.printPages(printPrev.printPages.UBound)
    printPrev.btnNext.Enabled = printPrev.printPages.UBound > 0
  End If
  InitPage poolName, headerBgFill, headerPos, noFooter, , extraLarge, noSubHead
End Sub
'
Private Sub fontSizing(fntSize As Integer)
' !!!! Font.Size for object and FontSize for printer !!!
    Printer.FontSize = fntSize
    With printObj.Font
        .Size = Printer.FontSize '* printRatio
    End With
End Sub

Sub initializeVars()

  currentMatch = getLastMatchPlayed(cn) 'Returns ORDER number
  toMatch = currentMatch
  With Printer 'use printer to be able to get the values
    .FontName = textFont
    .FontUnderline = 0
    .FontSize = 18
    lineHeight18 = .TextHeight("Dummy")
    .FontSize = 10
    lineHeight10 = .TextHeight("Dummy")
    .FontSize = 8
    lineHeight8 = .TextHeight("Dummy")
    .FontSize = 12
    lineHeight12 = .TextHeight("Dummy")
  End With
  MakeColors

End Sub

Private Sub Form_Load()
  Set cn = New ADODB.Connection
  With cn
    .ConnectionString = lclConn
    .CursorLocation = adUseClient
    .Open
  End With

  initializeVars
  updateForm
  
  centerForm Me
  UnifyForm Me

End Sub

Sub fillMatchCombo()
Dim sqlstr As String
Dim matchStr As String
  Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblTournamentSchedule "
  sqlstr = sqlstr & " WHERE tournamentID = " & thisTournament
  sqlstr = sqlstr & " AND matchPlayed = -1"
  sqlstr = sqlstr & " ORDER BY matchOrder"
  rs.Open sqlstr, cn, adOpenForwardOnly, adLockReadOnly
  With rs
    Do While Not .EOF
      matchStr = getMatchDescrCombo(rs!matchOrder, cn)
      Me.cmbMatchesPlayed.AddItem matchStr
      Me.cmbMatchesPlayed.ItemData(Me.cmbMatchesPlayed.NewIndex) = rs!matchOrder
      .MoveNext
    Loop
    .Close
  End With
  Set rs = Nothing
End Sub

Sub updateForm()
Dim i As Integer
Dim prntr As Printer
Dim matchDate As Date
Dim lastMatch As Integer 'last match ORDER number
'set some controls on the right place
  Me.picCompetitorList.Top = 90
  Me.picCompetitorList.Left = 3090
  Me.picPrnterSettings.Left = 3090
  Me.picPrnterSettings.Top = 2760

  cmbPrinters.Clear
  'Load the combo with all available printers
  For Each prntr In Printers
    cmbPrinters.AddItem prntr.DeviceName
    If Printer.DeviceName = prntr.DeviceName Then 'Current default
        cmbPrinters.Text = prntr.DeviceName
    End If
  Next
  lastMatch = getLastMatchPlayed(cn)
  'fill match combo
  fillMatchCombo
  
  'button to print everything for the next day
  If lastMatch >= 1 Then
      matchDate = getMatchInfo(lastMatch, "matchdate", cn)
      Me.btnPrntAllAfterDay.Enabled = getAllMatchesPlayedOnDay(matchDate, cn)
      'button to print the results for each participant at end of tournament
      Me.btnFinalPlayerPrint.Enabled = lastMatch = getMatchCount(0, cn)
      Me.btnStickers.Enabled = Me.btnFinalPlayerPrint.Enabled
      'option to print everything at the end of the tournament
      Me.chkEindstand.Enabled = Me.btnFinalPlayerPrint.Enabled
    
      
      Me.txtToMatch.Enabled = True
  End If
  ' Me.chkDblSide.Enabled = printersettings
  Me.upDnToMatch.Max = getCount("Select tournamentID from tblTournamentSchedule where tournamentID = " & thisTournament, cn)
  Me.upDnToMatch = lastMatch
  setCombo Me.cmbMatchesPlayed, lastMatch
  Me.upDnForMatch.Max = Me.upDnToMatch.Max + 1
  Me.upDnForMatch = lastMatch + 1
  Me.upDnCopies = 1
  Me.optPortrait = True
  Me.optPrintDoc(2).Enabled = getCount("Select competitorPoolID from tblCompetitorPools where poolID = " & thisPool, cn) > 0
  Me.optPrintDoc(2).Enabled = Me.optPrintDoc(7).Enabled
  Me.optPrintDoc(3).Enabled = Me.optPrintDoc(7).Enabled
  Me.optPrintDoc(7).Enabled = currentMatch > 0
  Me.optPrintDoc(4).Enabled = currentMatch > 0
  'Me.optPrintDoc(5).Visible = False
  'Me.optPrintDoc(6).Visible = False
  Me.optPrintDoc(5).Enabled = currentMatch > 0
  Me.optPrintDoc(6).Enabled = currentMatch > 0
  Me.optPrintDoc(8).Enabled = currentMatch > 0
  Me.optPrintDoc(9).Enabled = currentMatch > 0
  Me.optPrintDoc(0).value = True
  optPrintDoc_Click 0
  Screen.MousePointer = Default
  ' Me.chkDblSide.Visible = true
  Me.width = 6730
  Me.Height = 6000

End Sub

Function RandomColor() As Long
    RandomColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
End Function


Private Sub printSkyline()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sqlstr As String
Dim pnt As Integer
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim l As Integer
Dim xPos As Integer
Dim yPos As Integer
Dim yBot As Integer
Dim tmpX As Integer
Dim tmpY As Integer
Dim oldYpos As Integer
Dim bottom As Integer
Dim highNow As Integer
Dim longName As Integer
Dim matchCount As Integer
Dim poolCount As Integer
Dim poolOnPage As Integer
Dim pageCount As Integer
Dim graphWidth As Integer
Dim pntCount As Integer
Dim maxHeight As Integer
Dim scorePos As Integer
Dim maximal As Integer
Dim scaleFact As Double
Dim factor As Integer
Dim curPag As Integer
Dim poolsPage1 As Integer
Dim poolsPage1Pos As Integer
Dim matchDesc As String
Dim matchOrder As Integer
Dim matchNr As Integer
Dim txtStr As String
matchOrder = toMatch
matchNr = getMatchNumber(matchOrder, cn)
'MakeColors
'
txtStr = "Grafiek t/m " & matchOrder & "e wedstrijd" ' & ", Uitslag: " & getMatchResultPartStr(matchNr, 2, cn)

heading1 = txtStr
If Me.chkEindstand <> 0 Then
    heading1 = "Grafiek Eindstand"
End If

InitPage False, False, 0, True
Set rs = New ADODB.Recordset
fontSizing 8
xPos = 0 'printObj.CurrentX + printObj.TextWidth("0") + printObj.ScaleLeft
yPos = printObj.CurrentY
sqlstr = "Select competitorPoolID, nickName from tblCompetitorPools"
sqlstr = sqlstr & " WHERE poolid =  " & thisPool
sqlstr = sqlstr & " Order BY nickName"
rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
  rs.MoveLast
  rs.MoveFirst
  footerPos = printObj.ScaleHeight - 240
  longName = printObj.TextWidth(Left(getLongestNickName(cn), 15))
  longName = longName + printObj.TextWidth("00(99)")
  bottom = footerPos - longName + 60
  yBot = footerPos - TextHeight("999")
  poolCount = rs.RecordCount
 ' Me.optPortrait = True
  If Me.optLandscape Then 'landscape
    poolOnPage = 40  'maxHeight on landscape
  Else
    poolOnPage = 26 'maxHeight on portrait
  End If
  pageCount = 1
  'set margin, top, wdth and other values
  If poolCount > poolOnPage Then 'we need an extra page
    pageCount = (poolCount / (poolOnPage + 3) + 0.5)
  End If
  poolOnPage = Int(poolCount / pageCount + 0.5)  'so we have more room, wider graphs
  matchCount = getMatchCount(0, cn)
  highNow = getHighesPts(Me.upDnToMatch, cn) 'what is the maxHeight height of a graph, set a factor accordingly
  If highNow > 250 Then
    factor = 50
  ElseIf highNow > 150 Then
    factor = 25
  ElseIf highNow > 100 Then
    factor = 10
  Else
    factor = 5
  End If
  Do While pntCount <= highNow / factor
    pntCount = pntCount + factor
  Loop
  maxHeight = Int(highNow / pntCount + 1) * pntCount
  pntCount = maxHeight / factor
  scorePos = Int((bottom - yPos) / pntCount)
  'legenda
  printObj.FillStyle = vbSolid
  oldYpos = bottom
  fontSizing 6
  poolsPage1Pos = printObj.TextWidth("99: XXXX-XXXX") + 20
  printObj.ForeColor = vbBlack
  For i = 0 To matchOrder - 1
      printObj.FillColor = thisColor(i)
      printObj.Line (xPos, oldYpos)-(xPos + poolsPage1Pos, oldYpos - printObj.TextHeight("W")), lineColor, B
      printObj.CurrentX = xPos + 40
      SetForeCol thisColor(i)
      matchNr = getMatchNumber(i + 1, cn)
      printObj.Print getMatchDescription(matchNr, cn, False, False, True) 'getWedTeams(i + 1)
      oldYpos = oldYpos - printObj.TextHeight("W")
      printObj.ForeColor = vbBlack
  Next
  fontSizing 8
  poolsPage1Pos = poolsPage1Pos + printObj.TextWidth("000")
  printObj.Line (xPos + poolsPage1Pos + 40, yPos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, yPos)
  printObj.Line -(printObj.ScaleWidth + 2 * printObj.ScaleLeft, bottom)
  printObj.Line -(xPos + poolsPage1Pos + 40, bottom)
  printObj.Line -(xPos + poolsPage1Pos + 40, yPos)
  For i = 0 To pntCount
      yPos = bottom - i * scorePos
      fontSizing 8
      printObj.Line (xPos + poolsPage1Pos + 40, yPos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, yPos)
      printObj.CurrentX = xPos + poolsPage1Pos + 40 - TextWidth(CStr(i * maxHeight / pntCount)) - 40
      printObj.CurrentY = yPos - TextHeight("999") / 2
      printObj.Print i * maxHeight / pntCount
  Next
  maximal = (i - 1) * pntCount
  scaleFact = (bottom - yPos) / maxHeight
  'fontSizing 4
  printObj.FontBold = False
  rs.MoveFirst
  'thisColor(0) = 15
  curPag = 1
  graphWidth = Int((printObj.ScaleWidth - (2 * printObj.ScaleLeft) - xPos - poolsPage1Pos) / poolOnPage)
  xPos = poolsPage1Pos + 40
  i = 0 'horizontale positie eerste deelnemer
  poolsPage1 = poolOnPage - i
  Do While Not rs.EOF
      i = i + 1
      oldYpos = bottom
  '    If curPag > 1 Then poolsPage1 = poolOnPage
      For J = 0 To matchOrder - 1
          printObj.FillColor = thisColor(J)
          
         ' pnt = Int(getDeelnPnt(GetWedNum(j + 1), rs!deelnemID, 1) * (scaleFact) + 0.5)
          pnt = Int(getPoolFormPoints(rs!competitorPoolID, J + 1, 41, cn) * (scaleFact) + 0.5)
          printObj.Line (xPos + 10 + graphWidth * (i - 1), oldYpos)-(xPos + graphWidth * (i - 1) + graphWidth - 10, oldYpos - pnt), lineColor, B
  
          oldYpos = oldYpos - pnt
      Next
      fontSizing 8
      printObj.CurrentX = xPos + graphWidth * (i - 1) + (graphWidth - printObj.TextWidth(Format(pnt, "999"))) / 2
      printObj.CurrentY = oldYpos - printObj.TextHeight(Format(pnt, "##"))
  
      printObj.Print getPoolFormPoints(rs!competitorPoolID, matchOrder, 43, cn)
      printObj.CurrentX = xPos + graphWidth * (i - 1) + (graphWidth - TextWidth("W")) / 2
      tmpX = printObj.CurrentX
  
      printObj.CurrentY = bottom + printObj.TextWidth(Trim(rs!nickName) & " ")
      tmpY = printObj.CurrentY
      printObj.FontBold = False
      fontSizing 10
      Set rotater.Device = printObj
      printObj.CurrentY = bottom + 50
      printObj.CurrentX = xPos + graphWidth * (i - 1) + (graphWidth + printObj.TextWidth("W")) / 2
      rotater.Angle = 270
      rotater.PrintText Left(rs!nickName, 15) & " (" & getPoolFormPoints(rs!competitorPoolID, matchOrder, 46, cn) & ")"
      rs.MoveNext
      printObj.DrawWidth = 1
      If i = poolOnPage And Not rs.EOF Then
          addNewPage False, False, 0, True
          curPag = curPag + 1
          printObj.Line (xPos, yPos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, yPos)
          printObj.Line -(printObj.ScaleWidth + 2 * printObj.ScaleLeft, bottom)
          printObj.Line -(xPos, bottom)
          printObj.Line -(xPos, yPos)
  
          For i = 0 To pntCount
              yPos = bottom - i * scorePos
              fontSizing 8
              printObj.Line (xPos, yPos)-(printObj.ScaleWidth + 2 * printObj.ScaleLeft, yPos)
              printObj.CurrentX = xPos - TextWidth(CStr(i * maxHeight / pntCount)) - 10
              printObj.CurrentY = yPos - TextHeight("99") / 2
              printObj.Print i * maxHeight / pntCount
          Next
          i = 0
          printObj.FontBold = False
          printObj.FillStyle = vbSolid
      End If
  Loop
End If
rs.Close
Set rs = Nothing
End Sub
'
Private Sub InitPage(poolName As Boolean, Optional bgFill As Boolean, Optional headerPos As Integer, Optional noFooter As Boolean, Optional pagNr As Boolean, Optional extraLarge As Boolean, Optional noSubHead As Boolean)
    
    'start with the pageFooter
    If Not noFooter Or (noFooter And Me.chkNwePagKop) Then
      pageFooter  'if this is the only page or not the first page do footer
      footerPos = printObj.CurrentY
    Else
      footerPos = printObj.ScaleHeight - 240
    End If
    'print the page header
    If poolName Then
      pageHeader
    Else
      printObj.CurrentY = 0
    End If
    
    ' print the first heading
    If Not noSubHead Then
      reportTitle heading1, pagNr, bgFill, headerPos, , extraLarge
    End If

End Sub
'

Private Sub btnClose_Click()
On Error Resume Next
  
  On Error GoTo 0
  Printer.KillDoc
  Unload Me
  
End Sub

Private Sub pageHeader()
Dim headerText As String ' top of the page
Dim lineWidth As Integer
Dim fnt As String
Dim yPos As Integer ' vertical position
  headerText = getOrganisation(cn) & " - " & getPoolInfo("poolName", cn)
  With printObj
    .ForeColor = headBGcolor
    fnt = .FontName  'remember old font
    .FontName = headingFont
    lineWidth = .DrawWidth  'remember drawwidth
    .DrawWidth = 2
    printObj.Line (0, 0)-(.ScaleWidth, 0), headBGcolor
    fontSizing 2
    printObj.Print 'small linebreak
    fontSizing 16
    printObj.FontBold = True
      centerText headerText
    printObj.FontBold = False
    printObj.Print
    yPos = .CurrentY
    printObj.Line (0, yPos)-(.ScaleWidth, yPos), headBGcolor
    fontSizing 1 'small newline
    printObj.Print
    headerHeight = .CurrentY 'store the headerHeight
    .DrawWidth = lineWidth 'reset DraWidht
    .ForeColor = vbBlack
    .FontName = fnt 'reset font
  End With
End Sub

Private Sub reportTitle(txt As String, Optional pagNr As Boolean, Optional bgFill As Boolean, Optional yPos As Integer, Optional hAlign As Integer, Optional extraLarge As Boolean)

Dim savYpos As String

    fontSizing 16
    If extraLarge Then fontSizing 20
    printObj.FillColor = headBGcolor
    If bgFill Then
        printObj.FillStyle = vbFSSolid
        printObj.ForeColor = vbWhite
        
        printObj.Line (0, headerHeight)-(printObj.ScaleWidth - 20, headerHeight + printObj.TextHeight("W")), lineColor, B
    Else
        printObj.ForeColor = RGB(0, 128, 0)
        printObj.FillStyle = vbFSTransparent
    End If
    printObj.FontItalic = True
    printObj.FontBold = True
    printObj.CurrentY = headerHeight
    If yPos > 0 Then printObj.CurrentY = yPos

    iBKMode = SetBkMode(printObj.hdc, TRANSPARENT)
    Select Case hAlign  'set horizontal position
    Case 0  'center
        centerText txt
    Case 1  'left align
        printObj.CurrentX = 0
    Case 2   ' ¼ page
        printObj.CurrentX = Int(printObj.ScaleWidth / 4) - printObj.TextWidth(txt) / 2
    Case 3    '½ page
        printObj.CurrentX = Int(printObj.ScaleWidth / 2) - printObj.TextWidth(txt) / 2
    Case 4    '¾ page
        printObj.CurrentX = Int(printObj.ScaleWidth / 4) * 3 - printObj.TextWidth(txt) / 2
    End Select
    If hAlign <> 0 Then printObj.Print txt; 'if hAlign is 0 then txt is allready printed ;-)
    'If hAlign = 0 Then printObj.Print 'add new line if text is centered
    
    savYpos = printObj.CurrentY 'set vertical position
    '
    printObj.CurrentY = printObj.CurrentY + lineHeight18 - lineHeight10  'set the vertical position
    fontSizing 9
    printObj.CurrentX = printObj.ScaleWidth - printObj.TextWidth("blad 12")  'position of the page nr
    If pagNr Then
      If TypeOf printObj Is Printer Then
          If printObj.Page > 1 Then
              printObj.Print "blad "; printObj.Page;
          End If
      Else
          If printObj.Index > 0 Then
              printObj.Print "blad "; printObj.Index + 1;
          End If
      End If
    End If
    fontSizing 12
    printObj.Print
    headerHeight = printObj.CurrentY  'set the height of the header plus reportTitle (if printed)
    'reset default settings
    printObj.FillStyle = vbFSTransparent
    printObj.ForeColor = vbBlack
    printObj.FontItalic = False
    printObj.FontBold = False
End Sub

Function getAant(deeln As Long, vanwat As String)
''haal het aantal scores op van 'vanwat' bij deeln
Dim rs As New ADODB.Recordset
Dim sqlstr As String
  sqlstr = "SELECT * from tblCompetitorPoints"
  sqlstr = sqlstr & " Where competitorPoolID = " & deeln
  sqlstr = sqlstr & " AND " & vanwat & " > 0"
  sqlstr = sqlstr & " AND matchOrder <= " & Me.upDnToMatch
  
  With rs
    .Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If Not .EOF Then
        .MoveLast
    End If
    getAant = .RecordCount
    .Close
  End With
  Set rs = Nothing
End Function
'
Function GetPntDeelnem(deeln As Long, vanwat As String)
'Dim rsdeelnScore As New ADODB.Recordset
'Dim pnt As Integer
'Dim sqlstr As String
'    sqlstr = "SELECT * from deelnempnt"
'    sqlstr = sqlstr & " Where deelnID =" & deeln
'    sqlstr = sqlstr & " AND " & vanwat & " > 0"
'    sqlstr = sqlstr & " order by wednum"
'    rsdeelnScore.Open sqlstr, cn, adOpenStatic, adLockReadOnly
'    If rsdeelnScore.RecordCount > 0 Then
'        rsdeelnScore.MoveLast
'        GetPntDeelnem = rsdeelnScore(vanwat)
'        If UCase(Left(vanwat, 7)) = UCase("pntfin4") Then
'            rsdeelnScore.MoveFirst
'            pnt = 0
'            Do While Not rsdeelnScore.EOF
'                pnt = pnt + rsdeelnScore(vanwat)
'                rsdeelnScore.MoveNext
'            Loop
'            GetPntDeelnem = pnt
'        ElseIf UCase(Left(vanwat, 7)) = UCase("pntfin2") Then
'            rsdeelnScore.MoveFirst
'            pnt = 0
'            Do While Not rsdeelnScore.EOF
'                pnt = pnt + rsdeelnScore(vanwat)
'                rsdeelnScore.MoveNext
'            Loop
'            GetPntDeelnem = pnt
'        End If
'    Else
'        GetPntDeelnem = 0
'    End If
End Function

Function getPointsForThis(id As Long)
'helper function to include connection to database
  getPointsForThis = getPointsForID(id, cn)
End Function
'

'Sub poolPointsReportHeader(leftMargin As Integer)
'
'    printObj.CurrentX = leftMargin
'    printObj.Print "Naam";
'    printObj.CurrentX = printObj.TextWidth("123456789012345")
'    ReDim Preserve pntPos(1)
'    pntPos(0) = 0
'    pntPos(1) = printObj.CurrentX - colwidth
'    printObj.Print
'    top2Ypos = printObj.CurrentY
'    printObj.CurrentX = pntPos(1) + colwidth
'    fontSizing 8
'    'we print the second line first to be able to calculate the positions
'    'uitslagen
'    printObj.Print "rust";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "eind";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "toto"; '("; Format(getPointsForThis(3), pntFormat); "p)";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    'doelpunten van de dag
'    printObj.Print "dlp"; '("; Format(getPointsForThis(28), pntFormat); "p)";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "total";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    'groepstanden
'    grpStndBegin = UBound(pntPos)  'the item in the position array where the groups stands startwith
'    For i = 1 To grpCount
'        printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'        printObj.Print " "; Chr(i + 64);
'        ReDim Preserve pntPos(UBound(pntPos) + 1)
'        pntPos(UBound(pntPos)) = printObj.CurrentX
'    Next
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "tot";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    'finales
'    If grpCount > 4 Then
'      '8e finales
'        fin8Begin = UBound(pntPos)
'
'        For i = 1 To grpCount
'            printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'            printObj.Print " "; Chr(i + 64);
'            ReDim Preserve pntPos(UBound(pntPos) + 1)
'            pntPos(UBound(pntPos)) = printObj.CurrentX
'        Next
'        printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'        printObj.Print "tot";
'        ReDim Preserve pntPos(UBound(pntPos) + 1)
'        pntPos(UBound(pntPos)) = printObj.CurrentX
'    End If
'    '4e finales
'    fin4Begin = UBound(pntPos)
'    For i = 1 To 4
'        printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'        printObj.Print " "; Format(i, "0"); " ";
'        ReDim Preserve pntPos(UBound(pntPos) + 1)
'        pntPos(UBound(pntPos)) = printObj.CurrentX
'    Next
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "tot";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    '½ finales
'    fin2Begin = UBound(pntPos)
'    For i = 1 To 2
'        printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'        printObj.Print " "; Format(i, "0"); " ";
'        ReDim Preserve pntPos(UBound(pntPos) + 1)
'        pntPos(UBound(pntPos)) = printObj.CurrentX
'    Next
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "tot";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    'finales (klein & groot)
'    finBegin = UBound(pntPos)
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    If thirdPlace Then
'        printObj.Print " kl "; '("; Format(getPointsForThis(30), pntFormat);
''        If getPointsForThis(31) > 0 Then
''            printObj.Print "/"; Format(getPointsForThis(31), pntFormat);
''        End If
''        printObj.Print ")";
'        ReDim Preserve pntPos(UBound(pntPos) + 1)
'        pntPos(UBound(pntPos)) = printObj.CurrentX
'        printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'        printObj.Print " gr "; '("; Format(getPointsForThis(11), pntFormat);
''        If getPointsForThis(12) > 0 Then
''            printObj.Print "/"; Format(getPointsForThis(12), pntFormat);
''        End If
''        printObj.Print ")";
'        ReDim Preserve pntPos(UBound(pntPos) + 1)
'        pntPos(UBound(pntPos)) = printObj.CurrentX
'        printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    Else
'        printObj.Print "("; Format(getPointsForThis(11), pntFormat);
'        If getPointsForThis(12) > 0 Then
'            printObj.Print "/"; Format(getPointsForThis(12), pntFormat);
'        End If
'        printObj.Print ")";
'        ReDim Preserve pntPos(UBound(pntPos) + 1)
'        pntPos(UBound(pntPos)) = printObj.CurrentX
'        printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    End If
'    EindstBegin = UBound(pntPos)
'    ' Format(getPointsForThis(15), pntFormat); "/"; Format(getPointsForThis(14), pntFormat); "/"; Format(getPointsForThis(13), pntFormat); "/"; Format(getPointsForThis(29), pntFormat); ")";
'    printObj.Print " 1 ";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print " 2 ";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    If thirdPlace Then
'        printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'        printObj.Print " 3 ";
'        ReDim Preserve pntPos(UBound(pntPos) + 1)
'        pntPos(UBound(pntPos)) = printObj.CurrentX
'        printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'        printObj.Print " 4 ";
'        ReDim Preserve pntPos(UBound(pntPos) + 1)
'        pntPos(UBound(pntPos)) = printObj.CurrentX
'    End If
'    AantBegin = UBound(pntPos)
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "dp";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "gs";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print " gl";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "rd";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "pn";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "ed";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    TopScBegin = UBound(pntPos)
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth ' + printObj.TextWidth("sc")
'    printObj.Print " ts";
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    TTLBegin = UBound(pntPos)
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth + printObj.TextWidth("123")
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    PosBegin = UBound(pntPos)
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth + printObj.TextWidth("123")
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.CurrentX
'    GeldBegin = UBound(pntPos)
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print "";
'    'laatste kolom
'    ReDim Preserve pntPos(UBound(pntPos) + 1)
'    pntPos(UBound(pntPos)) = printObj.ScaleWidth - 50
''Now we print the first line
'    printObj.CurrentY = topYpos
'    fontSizing 10
'    printObj.CurrentX = (pntPos(1) + pntPos(grpStndBegin) + colwidth - printObj.TextWidth("Wedstrijdpunten")) / 2
'    printObj.Print "Wedstrijdpunten";
'    If grpCount > 4 Then
'        printObj.CurrentX = (pntPos(grpStndBegin) + pntPos(fin8Begin) + colwidth - printObj.TextWidth("Groepstand (" & Format(getPointsForThis(8), pntFormat) & "p)")) / 2
'    Else
'        printObj.CurrentX = (pntPos(grpStndBegin) + pntPos(fin4Begin) + colwidth - printObj.TextWidth("Groepstand (" & Format(getPointsForThis(8), pntFormat) & "p)")) / 2
'    End If
'    printObj.Print "Groepstand"; ' (" & Format(getPointsForThis(8), pntFormat) & "p)";
'    If grpCount > 4 Then
'        printObj.CurrentX = (pntPos(fin8Begin) + pntPos(fin4Begin) + colwidth - printObj.TextWidth("8e Finalisten (" & Format(getPointsForThis(6), pntFormat) & "/" & Format(getPointsForThis(7), pntFormat) & "p)")) / 2
'        printObj.Print "8e Finalisten"; ' (" & Format(getPointsForThis(4), pntFormat);
''        If getPointsForThis(5) > 0 Then
''            printObj.Print "/" & Format(getPointsForThis(5), pntFormat);
''        End If
''        printObj.Print "p)";
'    End If
'    printObj.CurrentX = (pntPos(fin4Begin) + pntPos(fin2Begin) + colwidth - printObj.TextWidth("4e fin.(" & Format(getPointsForThis(6), pntFormat) & "/" & Format(getPointsForThis(7), pntFormat) & "p)")) / 2
'    printObj.Print "¼ finalisten"; '(" & Format(getPointsForThis(6), pntFormat);
''    If getPointsForThis(7) > 0 Then
''        printObj.Print "/" & Format(getPointsForThis(7), pntFormat);
''    End If
''    printObj.Print "p)";
'    printObj.CurrentX = (pntPos(fin2Begin) + pntPos(finBegin) + colwidth - printObj.TextWidth("½finale")) / 2 '(" & Format(getPointsForThis(9), pntFormat) & "/" & Format(getPointsForThis(10), pntFormat) & "p)")) / 2
'    printObj.Print "½ finale"; ' (" & Format(getPointsForThis(9), pntFormat);
''    If getPointsForThis(10) > 0 Then
''        printObj.Print "/" & Format(getPointsForThis(10), pntFormat);
''    End If
''    printObj.Print "p)";
'    printObj.CurrentX = (pntPos(finBegin) + pntPos(EindstBegin) + colwidth - printObj.TextWidth("Finale")) / 2
'    printObj.Print "Finale";
'    printObj.CurrentX = (pntPos(EindstBegin) + pntPos(AantBegin) + colwidth - printObj.TextWidth("Eindstand")) / 2
'    printObj.Print "Eindstand";
'    printObj.CurrentX = (pntPos(AantBegin) + pntPos(TopScBegin) + colwidth - printObj.TextWidth("Statistiek")) / 2
'    printObj.Print "Statistiek";
'    printObj.CurrentX = pntPos(TopScBegin) + colwidth
'    printObj.Print "";
'    printObj.CurrentX = (pntPos(TTLBegin) + pntPos(PosBegin) + colwidth - printObj.TextWidth("Ttl")) / 2
'    printObj.Print "Ttl";
'    printObj.CurrentX = (pntPos(PosBegin) + pntPos(GeldBegin) + colwidth - printObj.TextWidth("Pos")) / 2
'    printObj.Print "Pos";
'    printObj.CurrentX = (pntPos(GeldBegin) + pntPos(GeldBegin + 1) + colwidth - printObj.TextWidth("Geld")) / 2
'    printObj.Print "Geld";
'    fontSizing 8
'    printObj.CurrentY = top2Ypos
'    printObj.CurrentX = pntPos(UBound(pntPos)) + colwidth
'    printObj.Print
'    printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
'
'End Sub

Sub printPoolPoints(alfabet As Boolean)
Dim rs As New ADODB.Recordset
'Dim rsScore As New adodb.Recordset
Dim sqlstr As String
'Dim bedr As Currency
'Dim geldold As Currency
'Dim savy As Integer
Dim leftMargin As Integer
Dim pntPos() As Integer
Dim pnt As Integer
Dim aant As Integer
Dim prntPts As Integer
Dim geld As Double
Dim geldttl As Double
Dim txtStr As String
Dim prStr As String
Dim topYpos As Integer
Dim top2Ypos As Integer
Dim botY As Integer
Dim lastDeelnPos As Integer
Dim maxY As Integer
Dim grp As String
Dim i As Integer
Dim J As Integer
Dim ipos As Integer
Dim has8eFin As Boolean
Dim thirdPlace As Boolean
Dim grpCount As Integer
Dim matchNr As Integer
Dim prTtl As Boolean
Dim colWidth As Integer
Dim grpStndBegin As Integer '6
Dim fin8Begin As Integer    '15
Dim fin4Begin As Integer    '24
Dim fin2Begin As Integer    '29
Dim finBegin As Integer     '32
'
Dim EindstBegin As Integer  '34
Dim AantBegin As Integer    '38
Dim TopScBegin As Integer   '43
Dim TTLBegin As Integer     '44
Dim PosBegin As Integer     '45
Dim GeldBegin As Integer    '46
Dim pntFormat As String
Dim tmp As String
Dim yposnu As Integer
Dim final8Pts(8) As Integer

'
    grpCount = getTournamentInfo("tournamentgroupcount", cn)
    Select Case grpCount
    Case 4
      colWidth = 250
    Case 6
      colWidth = 200
    Case Else
        colWidth = 140
    End Select
    has8eFin = grpCount > 4
    thirdPlace = getTournamentInfo("tournamentThirdPlace", cn)
    If getLastMatchPlayed(cn) = getMatchCount(0, cn) Then
        pntFormat = "0"
    Else
        pntFormat = "0;;\ ;-"
    End If

    leftMargin = 0
    fontSizing 10
    printObj.Print
    matchNr = getMatchNumber(toMatch, cn)
    If Me.chkEindstand = False Then
        If alfabet Then
            'txtStr = "Puntenopbouw t/m wedstrijd " & matchNr
            txtStr = "Puntenopbouw t/m " & toMatch & "e wedstrijd: " & getMatchDescription(toMatch, cn) & ", Uitslag: " & getMatchResultPartStr(toMatch, 2, cn)

        Else
            txtStr = "Puntenopbouw t/m " & toMatch & "e wedstrijd (hoog-laag)"
        End If
    Else
        If alfabet Then
            txtStr = "Eindstand (alfabetisch)"
        Else
            txtStr = "Eindstand (op score)"
        End If
    End If
    
    heading1 = txtStr
    InitPage True, False, , , , True
    
    printObj.FontItalic = False
    printObj.FontBold = False
    topYpos = printObj.CurrentY
    On Error Resume Next
    printObj.Line (0, topYpos)-(printObj.ScaleWidth - 50, topYpos)
    On Error GoTo 0
    printObj.CurrentX = leftMargin
    'we need the past position in the pool (ther could be more then one
    sqlstr = getPoolFormSql(False, toMatch)
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveLast
        lastDeelnPos = rs!positionTotal
    End If
    rs.Close
    'now reset the recordset
    sqlstr = getPoolFormSql(alfabet, toMatch)
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly

    If rs.RecordCount = 0 Then
        printObj.Print "Geen deelnemers gevonden"
        Exit Sub
    End If
    fontSizing 10
'    poolPointsReportHeader leftMargin
    printObj.CurrentX = leftMargin
    printObj.Print "Naam (pos)";
    printObj.CurrentX = printObj.TextWidth("123456789012345")
    ReDim Preserve pntPos(1)
    pntPos(0) = 0
    pntPos(1) = printObj.CurrentX - colWidth
    printObj.Print
    top2Ypos = printObj.CurrentY
    printObj.CurrentX = pntPos(1) + colWidth
    fontSizing 8
    'we print the second line first to be able to calculate the positions
    'uitslagen
    printObj.Print "rust";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "eind";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "toto"; '("; Format(getPointsForThis(3), pntFormat); "p)";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    'doelpunten van de dag
    printObj.Print "dlp"; '("; Format(getPointsForThis(28), pntFormat); "p)";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "total";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    'groepstanden
    grpStndBegin = UBound(pntPos)  'the item in the position array where the groups stands startwith
    For i = 1 To grpCount
        printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
        printObj.Print " "; Chr(i + 64);
        ReDim Preserve pntPos(UBound(pntPos) + 1)
        pntPos(UBound(pntPos)) = printObj.CurrentX
    Next
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "tot";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    'finales
    If grpCount > 4 Then
      '8e finales
        fin8Begin = UBound(pntPos)

        For i = 1 To grpCount
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print " "; Chr(i + 64);
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
        Next
        printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
        printObj.Print "tot";
        ReDim Preserve pntPos(UBound(pntPos) + 1)
        pntPos(UBound(pntPos)) = printObj.CurrentX
    End If
    '4e finales
    fin4Begin = UBound(pntPos)
    For i = 1 To 4
        printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
        printObj.Print " "; Format(i, "0"); " ";
        ReDim Preserve pntPos(UBound(pntPos) + 1)
        pntPos(UBound(pntPos)) = printObj.CurrentX
    Next
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "tot";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    '½ finales
    fin2Begin = UBound(pntPos)
    For i = 1 To 2
        printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
        printObj.Print " "; Format(i, "0"); " ";
        ReDim Preserve pntPos(UBound(pntPos) + 1)
        pntPos(UBound(pntPos)) = printObj.CurrentX
    Next
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "tot";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    'finales (klein & groot)
    finBegin = UBound(pntPos)
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    If thirdPlace Then
        printObj.Print " kl "; '("; Format(getPointsForThis(30), pntFormat);
'        If getPointsForThis(31) > 0 Then
'            printObj.Print "/"; Format(getPointsForThis(31), pntFormat);
'        End If
'        printObj.Print ")";
        ReDim Preserve pntPos(UBound(pntPos) + 1)
        pntPos(UBound(pntPos)) = printObj.CurrentX
        printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
        printObj.Print " gr "; '("; Format(getPointsForThis(11), pntFormat);
'        If getPointsForThis(12) > 0 Then
'            printObj.Print "/"; Format(getPointsForThis(12), pntFormat);
'        End If
'        printObj.Print ")";
        ReDim Preserve pntPos(UBound(pntPos) + 1)
        pntPos(UBound(pntPos)) = printObj.CurrentX
        printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    Else
        printObj.Print "("; Format(getPointsForThis(11), pntFormat);
        If getPointsForThis(12) > 0 Then
            printObj.Print "/"; Format(getPointsForThis(12), pntFormat);
        End If
        printObj.Print ")";
        ReDim Preserve pntPos(UBound(pntPos) + 1)
        pntPos(UBound(pntPos)) = printObj.CurrentX
        printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    End If
    EindstBegin = UBound(pntPos)
    ' Format(getPointsForThis(15), pntFormat); "/"; Format(getPointsForThis(14), pntFormat); "/"; Format(getPointsForThis(13), pntFormat); "/"; Format(getPointsForThis(29), pntFormat); ")";
    printObj.Print " 1 ";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print " 2 ";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    If thirdPlace Then
        printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
        printObj.Print " 3 ";
        ReDim Preserve pntPos(UBound(pntPos) + 1)
        pntPos(UBound(pntPos)) = printObj.CurrentX
        printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
        printObj.Print " 4 ";
        ReDim Preserve pntPos(UBound(pntPos) + 1)
        pntPos(UBound(pntPos)) = printObj.CurrentX
    End If
    AantBegin = UBound(pntPos)
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "dp";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "gs";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print " gl";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "rd";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "pn";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "ed";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    TopScBegin = UBound(pntPos)
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth ' + printObj.TextWidth("sc")
    printObj.Print " ts";
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    TTLBegin = UBound(pntPos)
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth + printObj.TextWidth("123")
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    PosBegin = UBound(pntPos)
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth + printObj.TextWidth("123")
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.CurrentX
    GeldBegin = UBound(pntPos)
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print "";
    'laatste kolom
    ReDim Preserve pntPos(UBound(pntPos) + 1)
    pntPos(UBound(pntPos)) = printObj.ScaleWidth - 50
'Now we print the first line
    printObj.CurrentY = topYpos
    fontSizing 10
    printObj.CurrentX = (pntPos(1) + pntPos(grpStndBegin) + colWidth - printObj.TextWidth("Wedstrijdpunten")) / 2
    printObj.Print "Wedstrijdpunten";
    If grpCount > 4 Then
        printObj.CurrentX = (pntPos(grpStndBegin) + pntPos(fin8Begin) + colWidth - printObj.TextWidth("Groepstand (" & Format(getPointsForThis(8), pntFormat) & "p)")) / 2
    Else
        printObj.CurrentX = (pntPos(grpStndBegin) + pntPos(fin4Begin) + colWidth - printObj.TextWidth("Groepstand (" & Format(getPointsForThis(8), pntFormat) & "p)")) / 2
    End If
    printObj.Print "Groepstand"; ' (" & Format(getPointsForThis(8), pntFormat) & "p)";
    If grpCount > 4 Then
        printObj.CurrentX = (pntPos(fin8Begin) + pntPos(fin4Begin) + colWidth - printObj.TextWidth("8e Finalisten (" & Format(getPointsForThis(6), pntFormat) & "/" & Format(getPointsForThis(7), pntFormat) & "p)")) / 2
        printObj.Print "8e Finalisten"; ' (" & Format(getPointsForThis(4), pntFormat);
'        If getPointsForThis(5) > 0 Then
'            printObj.Print "/" & Format(getPointsForThis(5), pntFormat);
'        End If
'        printObj.Print "p)";
    End If
    printObj.CurrentX = (pntPos(fin4Begin) + pntPos(fin2Begin) + colWidth - printObj.TextWidth("4e fin.(" & Format(getPointsForThis(6), pntFormat) & "/" & Format(getPointsForThis(7), pntFormat) & "p)")) / 2
    printObj.Print "¼ finalisten"; '(" & Format(getPointsForThis(6), pntFormat);
'    If getPointsForThis(7) > 0 Then
'        printObj.Print "/" & Format(getPointsForThis(7), pntFormat);
'    End If
'    printObj.Print "p)";
    printObj.CurrentX = (pntPos(fin2Begin) + pntPos(finBegin) + colWidth - printObj.TextWidth("½finale")) / 2 '(" & Format(getPointsForThis(9), pntFormat) & "/" & Format(getPointsForThis(10), pntFormat) & "p)")) / 2
    printObj.Print "½ finale"; ' (" & Format(getPointsForThis(9), pntFormat);
'    If getPointsForThis(10) > 0 Then
'        printObj.Print "/" & Format(getPointsForThis(10), pntFormat);
'    End If
'    printObj.Print "p)";
    printObj.CurrentX = (pntPos(finBegin) + pntPos(EindstBegin) + colWidth - printObj.TextWidth("Finale")) / 2
    printObj.Print "Finale";
    printObj.CurrentX = (pntPos(EindstBegin) + pntPos(AantBegin) + colWidth - printObj.TextWidth("Eindstand")) / 2
    printObj.Print "Eindstand";
    printObj.CurrentX = (pntPos(AantBegin) + pntPos(TopScBegin) + colWidth - printObj.TextWidth("Statistiek")) / 2
    printObj.Print "Statistiek";
    printObj.CurrentX = pntPos(TopScBegin) + colWidth
    printObj.Print "";
    printObj.CurrentX = (pntPos(TTLBegin) + pntPos(PosBegin) + colWidth - printObj.TextWidth("Ttl")) / 2
    printObj.Print "Ttl";
    printObj.CurrentX = (pntPos(PosBegin) + pntPos(GeldBegin) + colWidth - printObj.TextWidth("Pos")) / 2
    printObj.Print "Pos";
    printObj.CurrentX = (pntPos(GeldBegin) + pntPos(GeldBegin + 1) + colWidth - printObj.TextWidth("Geld")) / 2
    printObj.Print "Geld";
    fontSizing 8
    printObj.CurrentY = top2Ypos
    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
    printObj.Print
    printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)

'''''''''''' START DE TABEL
    With rs
      Do While Not .EOF
        printObj.CurrentX = leftMargin
        If !positionTotal = 1 Then
          printObj.ForeColor = vbBlue
          printObj.FontBold = True
        End If
        If !positionTotal = lastDeelnPos Then
          printObj.ForeColor = vbRed
        End If
        printObj.Print !nickName; " (" & !positionTotal & ")";
        printObj.ForeColor = 1
        printObj.FontBold = False
        pnt = printAant(!competitorPoolID, pntPos(2), "ptsHt")
        pnt = pnt + printAant(!competitorPoolID, pntPos(3), "ptsFt")
        pnt = pnt + printAant(!competitorPoolID, pntPos(4), "ptsToto")
        pnt = pnt + printAant(!competitorPoolID, pntPos(5), "ptsDayGoals")
        printObj.CurrentX = pntPos(6) - printObj.TextWidth(Format(pnt, pntFormat))
        printObj.FontBold = True
        printObj.Print Format(pnt, pntFormat);
        printObj.FontBold = False
        pnt = 0
        prntPts = 0
        'pointsgrpA-H
        For i = 1 To grpCount
          If grpPlayedAll(Chr(i + 64), cn) Then
              pntFormat = "0"
          Else
              pntFormat = "0;;\ ;-"
          End If
          prntPts = getPoolFormPoints(!competitorPoolID, Me.upDnToMatch, 7 + i, cn, Chr(i + 64))
          pnt = pnt + prntPts
          printObj.CurrentX = (pntPos(i + 5) + pntPos(i + 6) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
          printObj.Print Format(prntPts, pntFormat);
        Next
        If grpCount > 4 Then
          printObj.CurrentX = pntPos(fin8Begin) - printObj.TextWidth(Format(pnt, pntFormat))
        Else
          printObj.CurrentX = pntPos(fin4Begin) - printObj.TextWidth(Format(pnt, pntFormat))
        End If
        printObj.FontBold = True
        printObj.Print Format(pnt, pntFormat);
        printObj.FontBold = False
        pnt = 0
        prntPts = 0
        If grpCount > 4 Then
        'pointsTeamFinals8 A-H / 8e finales
          For i = 1 To grpCount
            If grpPlayedAll(Chr(i + 64), cn) Then
              pntFormat = "0"
              prntPts = getFin8Points(!competitorPoolID, Me.upDnToMatch, Chr(i + 64), cn)
            Else
              prntPts = 0
              pntFormat = "0;;\ ;-"
            End If
            pnt = pnt + prntPts
            printObj.CurrentX = (pntPos(fin8Begin - 1 + i) + pntPos(i + fin8Begin) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
            printObj.Print Format(prntPts, pntFormat);
          Next
          printObj.CurrentX = pntPos(fin4Begin) - printObj.TextWidth(Format(pnt, pntFormat))
          printObj.FontBold = True
          If grpPlayedAll("A", cn) Then
            pntFormat = "0"
          Else
            pntFormat = "0;;\ ;-"
          End If
          printObj.Print Format(pnt, pntFormat);
          printObj.FontBold = False
          pnt = 0
          prntPts = 0
        Else
          For i = 1 To grpCount  'kwartfinales /pointsTeamsFinals4 A-H als er GEEN 8e finale is
            If grpPlayedAll(Chr(i + 64), cn) Then
              pntFormat = "0"
              prntPts = getPoolFormPoints(!competitorID, toMatch, 4, cn, Chr(i + 64))
            Else
              prntPts = 0
              pntFormat = "0;;\ ;-"
            End If
            pnt = pnt + prntPts
            printObj.CurrentX = (pntPos(fin4Begin - 1 + i) + pntPos(i + fin4Begin) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
            printObj.Print Format(prntPts, pntFormat);
          Next
          printObj.FontBold = True
          If grpPlayedAll("A", cn) Then
            pntFormat = "0"
          Else
            pntFormat = "0;;\ ;-"
          End If
          printObj.CurrentX = pntPos(fin2Begin) - printObj.TextWidth(Format(pnt, pntFormat))
          printObj.Print Format(pnt, pntFormat);
          printObj.FontBold = False
          pnt = 0
          prntPts = 0
        End If

        If grpCount > 4 Then
          'If !competitorPoolID = 7 Then Stop
            For i = 1 To 4
              prntPts = getPoolFormPoints(!competitorPoolID, toMatch, 25 + i, cn, grp)
              'If prntPts > 0 Then Stop
              printObj.CurrentX = (pntPos(i + fin4Begin - 1) + pntPos(i + fin4Begin) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
              pnt = pnt + prntPts
              If toMatch >= getFirstFinalMatchNumber(cn) Then
                pntFormat = "0"
              Else
                pntFormat = "0;;\ ;-"
              End If

              printObj.Print Format(prntPts, pntFormat);
            Next
            prTtl = True
          If prTtl > 0 Then pntFormat = "0"
          printObj.CurrentX = pntPos(fin2Begin) - printObj.TextWidth(Format(pnt, pntFormat))
          printObj.FontBold = True
          printObj.Print Format(pnt, pntFormat);
          printObj.FontBold = False
        End If
        pnt = 0
        prntPts = 0
        'Semi finals
        For i = 1 To 2
          If getFinalmatchOrder(2, True, cn) - 1 <= getLastMatchPlayed(cn) Then
              pntFormat = "0"
          Else
              pntFormat = "0;;\ ;-"
          End If
          'If !competitorPoolId = 183 Then Stop
          prntPts = getPoolFormPoints(!competitorPoolID, toMatch, 30 + i, cn, Chr(i + 64))
          pnt = pnt + prntPts
          printObj.CurrentX = (pntPos(i + fin2Begin - 1) + pntPos(i + fin2Begin) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
          printObj.Print Format(prntPts, pntFormat);
        Next
        printObj.CurrentX = pntPos(finBegin) - printObj.TextWidth(Format(pnt, pntFormat))
        printObj.FontBold = True
        printObj.Print Format(pnt, pntFormat);
        printObj.FontBold = False
        If getFinalmatchOrder(3, True, cn) <= getLastMatchPlayed(cn) Then
          pntFormat = "0"
        Else
          pntFormat = "0;;\ ;-"
        End If
        If thirdPlace Then
          prntPts = getPoolFormPoints(!competitorPoolID, toMatch, 34, cn)
          printObj.CurrentX = pntPos(32) + (pntPos(33) - pntPos(32) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
          printObj.Print Format(prntPts, pntFormat);
        End If
'        If !competitorPoolID = 46 Then Stop
       prntPts = getPoolFormPoints(!competitorPoolID, toMatch, 37, cn)
       printObj.CurrentX = (pntPos(finBegin) + pntPos(EindstBegin) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
       printObj.FontBold = True
       printObj.Print Format(prntPts, pntFormat);
       printObj.FontBold = False

       pntFormat = "0;;\ ;-"
       If getLastMatchPlayed(cn) = getMatchCount(0, cn) Then
        pntFormat = "0"  'eindstand
         For i = 1 To 2
             prntPts = getPoolFormEndPoints(!competitorPoolID, i, cn)
             printObj.CurrentX = (pntPos(finBegin + i) + pntPos(EindstBegin + i) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
             printObj.Print Format(prntPts, pntFormat);
         Next
         pntFormat = "0;;\ ;-"
         If getLastMatchPlayed(cn) >= getMatchCount(0, cn) - 1 Then pntFormat = "0"
         If getTournamentInfo("tournamentThirdplace", cn) Then
          For i = 3 To 4
              prntPts = getPoolFormEndPoints(!competitorPoolID, i, cn)
              printObj.CurrentX = (pntPos(EindstBegin - 1 + i) + pntPos(EindstBegin + i) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
              printObj.Print Format(prntPts, pntFormat);
          Next
        End If
      End If
       pntFormat = "0;;\ ;-"
       If getLastMatchPlayed(cn) = getMatchCount(0, cn) Then
        'statistics
           pntFormat = "0"
           For i = dpAant To pensAant
             prntPts = getStatsPointsFor(i, !competitorPoolID, cn)
             printObj.CurrentX = (pntPos(AantBegin + i - dpAant) + pntPos(AantBegin + i - dpAant + 1) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
             printObj.Print Format(prntPts, pntFormat);
           Next
         'topscorer
          'pointsTopscorers
          prntPts = getPoolFormPoints(!competitorPoolID, toMatch, 39, cn)
          printObj.CurrentX = (pntPos(TopScBegin) + pntPos(TTLBegin) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
          printObj.Print prntPts;
       End If
'       'plaats en geld
       pntFormat = "0"
       If !positionTotal = 1 Then
         printObj.ForeColor = vbBlue
         printObj.FontBold = True
       End If
       If !positionTotal = lastDeelnPos Then
         printObj.ForeColor = vbRed
       End If
       'If !competitorPoolId = 125 Then Stop
       prntPts = getPoolFormPoints(!competitorPoolID, toMatch, 43, cn)
       printObj.CurrentX = (pntPos(TTLBegin) + pntPos(PosBegin) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
       printObj.Print Format(prntPts, pntFormat);
       prntPts = getPoolFormPoints(!competitorPoolID, toMatch, 46, cn)
       printObj.CurrentX = (pntPos(PosBegin) + pntPos(GeldBegin) + colWidth - printObj.TextWidth(Format(prntPts, pntFormat))) / 2
       printObj.Print Format(prntPts, pntFormat);
       printObj.ForeColor = 1
       printObj.FontBold = False
      ' If !competitorPoolID = 15 Then Stop
       If IsLastMatchOfDay(toMatch, cn) Then
         If toMatch <> getMatchCount(0, cn) Then
          geld = getTotalMoney(!competitorPoolID, getMatchPrevDay(toMatch + 1, cn)) 'lastPos money total'getPoolFormPoints(!competitorPoolID, tomatch, 50, cn)
        Else
          geld = getTotalMoney(!competitorPoolID, toMatch) 'lastPos money total'getPoolFormPoints(!competitorPoolID, tomatch, 50, cn)
        End If
       Else
         geld = getTotalMoney(!competitorPoolID, getMatchPrevDay(toMatch, cn)) 'lastPos money total'getPoolFormPoints(!competitorPoolID, tomatch, 50, cn)
       End If
       printObj.CurrentX = pntPos(GeldBegin + 1) - colWidth - printObj.TextWidth(Format(geld, "currency"))
       printObj.Print Format(geld, "currency");
       printObj.Print
       printObj.ForeColor = 1
       printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
       prntPts = 0

      .MoveNext
        
      'start a new page if necessary   BAD PROGRAMMING, just copied from beginning of the routine
        If printObj.CurrentY >= footerPos - 360 Then 'onderkant pagina
          If Not rs.EOF Then
            botY = printObj.CurrentY
            printObj.Line (pntPos(1) + 75, topYpos)-(pntPos(1) + 75, top2Ypos)
            printObj.Line (pntPos(grpStndBegin) + 75, topYpos)-(pntPos(6) + 75, top2Ypos)
            If grpCount > 4 Then
                printObj.Line (pntPos(fin8Begin) + 75, topYpos)-(pntPos(fin8Begin) + 75, top2Ypos)
            End If
            printObj.Line (pntPos(fin4Begin) + 75, topYpos)-(pntPos(fin4Begin) + 75, top2Ypos)
            printObj.Line (pntPos(fin2Begin) + 75, topYpos)-(pntPos(fin2Begin) + 75, top2Ypos)
            printObj.Line (pntPos(finBegin) + 75, topYpos)-(pntPos(finBegin) + 75, top2Ypos)
            printObj.Line (pntPos(EindstBegin) + 75, topYpos)-(pntPos(EindstBegin) + 75, top2Ypos)
            printObj.Line (pntPos(AantBegin) + 75, topYpos)-(pntPos(AantBegin) + 75, top2Ypos)
            printObj.Line (pntPos(TopScBegin) + 75, topYpos)-(pntPos(TopScBegin) + 75, top2Ypos)
            printObj.Line (pntPos(TTLBegin) + 75, topYpos)-(pntPos(TTLBegin) + 75, top2Ypos)
            printObj.Line (pntPos(PosBegin) + 75, topYpos)-(pntPos(PosBegin) + 75, top2Ypos)
            printObj.Line (pntPos(GeldBegin) + 75, topYpos)-(pntPos(GeldBegin) + 75, top2Ypos)
            For i = 1 To UBound(pntPos) - 1
                printObj.Line (pntPos(i) + 75, top2Ypos)-(pntPos(i) + 75, botY)
            Next
            printObj.Line (printObj.ScaleWidth - 50, topYpos)-(printObj.ScaleWidth - 50, botY)
            heading1 = heading1 & " (vervolg)"
            addNewPage False, False, 1, , False
            printObj.Line (0, topYpos)-(printObj.ScaleWidth - 50, topYpos)
            printObj.CurrentX = leftMargin
            printObj.CurrentY = topYpos
            fontSizing 10
            printObj.Print "Naam";
            printObj.CurrentX = printObj.TextWidth("123456789012345")
            ReDim Preserve pntPos(1)
            pntPos(0) = 0
            pntPos(1) = printObj.CurrentX - colWidth
            printObj.Print
            top2Ypos = printObj.CurrentY
            printObj.CurrentX = pntPos(1) + colWidth
            fontSizing 8
            'we print the second line first to be able to calculate the positions
            'uitslagen
            printObj.Print "rust";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "eind";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "toto"; '("; Format(getPointsForThis(3), pntFormat); "p)";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            'doelpunten van de dag
            printObj.Print "dlp"; '("; Format(getPointsForThis(28), pntFormat); "p)";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "total";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            'groepstanden
            grpStndBegin = UBound(pntPos)  'the item in the position array where the groups stands startwith
            For i = 1 To grpCount
                printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
                printObj.Print " "; Chr(i + 64);
                ReDim Preserve pntPos(UBound(pntPos) + 1)
                pntPos(UBound(pntPos)) = printObj.CurrentX
            Next
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "tot";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            'finales
            If grpCount > 4 Then
              '8e finales
                fin8Begin = UBound(pntPos)
        
                For i = 1 To grpCount
                    printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
                    printObj.Print " "; Chr(i + 64);
                    ReDim Preserve pntPos(UBound(pntPos) + 1)
                    pntPos(UBound(pntPos)) = printObj.CurrentX
                Next
                printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
                printObj.Print "tot";
                ReDim Preserve pntPos(UBound(pntPos) + 1)
                pntPos(UBound(pntPos)) = printObj.CurrentX
            End If
            '4e finales
            fin4Begin = UBound(pntPos)
            For i = 1 To 4
                printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
                printObj.Print " "; Format(i, "0"); " ";
                ReDim Preserve pntPos(UBound(pntPos) + 1)
                pntPos(UBound(pntPos)) = printObj.CurrentX
            Next
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "tot";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            '½ finales
            fin2Begin = UBound(pntPos)
            For i = 1 To 2
                printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
                printObj.Print " "; Format(i, "0"); " ";
                ReDim Preserve pntPos(UBound(pntPos) + 1)
                pntPos(UBound(pntPos)) = printObj.CurrentX
            Next
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "tot";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            'finales (klein & groot)
            finBegin = UBound(pntPos)
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            If thirdPlace Then
                printObj.Print " kl "; '("; Format(getPointsForThis(30), pntFormat);
        '        If getPointsForThis(31) > 0 Then
        '            printObj.Print "/"; Format(getPointsForThis(31), pntFormat);
        '        End If
        '        printObj.Print ")";
                ReDim Preserve pntPos(UBound(pntPos) + 1)
                pntPos(UBound(pntPos)) = printObj.CurrentX
                printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
                printObj.Print " gr "; '("; Format(getPointsForThis(11), pntFormat);
        '        If getPointsForThis(12) > 0 Then
        '            printObj.Print "/"; Format(getPointsForThis(12), pntFormat);
        '        End If
        '        printObj.Print ")";
                ReDim Preserve pntPos(UBound(pntPos) + 1)
                pntPos(UBound(pntPos)) = printObj.CurrentX
                printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            Else
                printObj.Print "("; Format(getPointsForThis(11), pntFormat);
                If getPointsForThis(12) > 0 Then
                    printObj.Print "/"; Format(getPointsForThis(12), pntFormat);
                End If
                printObj.Print ")";
                ReDim Preserve pntPos(UBound(pntPos) + 1)
                pntPos(UBound(pntPos)) = printObj.CurrentX
                printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            End If
            EindstBegin = UBound(pntPos)
            ' Format(getPointsForThis(15), pntFormat); "/"; Format(getPointsForThis(14), pntFormat); "/"; Format(getPointsForThis(13), pntFormat); "/"; Format(getPointsForThis(29), pntFormat); ")";
            printObj.Print " 1 ";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print " 2 ";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            If thirdPlace Then
                printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
                printObj.Print " 3 ";
                ReDim Preserve pntPos(UBound(pntPos) + 1)
                pntPos(UBound(pntPos)) = printObj.CurrentX
                printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
                printObj.Print " 4 ";
                ReDim Preserve pntPos(UBound(pntPos) + 1)
                pntPos(UBound(pntPos)) = printObj.CurrentX
            End If
            AantBegin = UBound(pntPos)
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "dp";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "gs";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print " gl";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "rd";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "pn";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "ed";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            TopScBegin = UBound(pntPos)
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth ' + printObj.TextWidth("sc")
            printObj.Print " ts";
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            TTLBegin = UBound(pntPos)
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth + printObj.TextWidth("123")
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            PosBegin = UBound(pntPos)
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth + printObj.TextWidth("123")
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.CurrentX
            GeldBegin = UBound(pntPos)
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print "";
            'laatste kolom
            ReDim Preserve pntPos(UBound(pntPos) + 1)
            pntPos(UBound(pntPos)) = printObj.ScaleWidth - 50
        'Now we print the first line
            printObj.CurrentY = topYpos
            fontSizing 10
            printObj.CurrentX = (pntPos(1) + pntPos(grpStndBegin) + colWidth - printObj.TextWidth("Wedstrijdpunten")) / 2
            printObj.Print "Wedstrijdpunten";
            If grpCount > 4 Then
                printObj.CurrentX = (pntPos(grpStndBegin) + pntPos(fin8Begin) + colWidth - printObj.TextWidth("Groepstand (" & Format(getPointsForThis(8), pntFormat) & "p)")) / 2
            Else
                printObj.CurrentX = (pntPos(grpStndBegin) + pntPos(fin4Begin) + colWidth - printObj.TextWidth("Groepstand (" & Format(getPointsForThis(8), pntFormat) & "p)")) / 2
            End If
            printObj.Print "Groepstand"; ' (" & Format(getPointsForThis(8), pntFormat) & "p)";
            If grpCount > 4 Then
                printObj.CurrentX = (pntPos(fin8Begin) + pntPos(fin4Begin) + colWidth - printObj.TextWidth("8e Finalisten (" & Format(getPointsForThis(6), pntFormat) & "/" & Format(getPointsForThis(7), pntFormat) & "p)")) / 2
                printObj.Print "8e Finalisten"; ' (" & Format(getPointsForThis(4), pntFormat);
        '        If getPointsForThis(5) > 0 Then
        '            printObj.Print "/" & Format(getPointsForThis(5), pntFormat);
        '        End If
        '        printObj.Print "p)";
            End If
            printObj.CurrentX = (pntPos(fin4Begin) + pntPos(fin2Begin) + colWidth - printObj.TextWidth("4e fin.(" & Format(getPointsForThis(6), pntFormat) & "/" & Format(getPointsForThis(7), pntFormat) & "p)")) / 2
            printObj.Print "¼ finalisten"; '(" & Format(getPointsForThis(6), pntFormat);
        '    If getPointsForThis(7) > 0 Then
        '        printObj.Print "/" & Format(getPointsForThis(7), pntFormat);
        '    End If
        '    printObj.Print "p)";
            printObj.CurrentX = (pntPos(fin2Begin) + pntPos(finBegin) + colWidth - printObj.TextWidth("½finale")) / 2 '(" & Format(getPointsForThis(9), pntFormat) & "/" & Format(getPointsForThis(10), pntFormat) & "p)")) / 2
            printObj.Print "½ finale"; ' (" & Format(getPointsForThis(9), pntFormat);
        '    If getPointsForThis(10) > 0 Then
        '        printObj.Print "/" & Format(getPointsForThis(10), pntFormat);
        '    End If
        '    printObj.Print "p)";
            printObj.CurrentX = (pntPos(finBegin) + pntPos(EindstBegin) + colWidth - printObj.TextWidth("Finale")) / 2
            printObj.Print "Finale";
            printObj.CurrentX = (pntPos(EindstBegin) + pntPos(AantBegin) + colWidth - printObj.TextWidth("Eindstand")) / 2
            printObj.Print "Eindstand";
            printObj.CurrentX = (pntPos(AantBegin) + pntPos(TopScBegin) + colWidth - printObj.TextWidth("Statistiek")) / 2
            printObj.Print "Statistiek";
            printObj.CurrentX = pntPos(TopScBegin) + colWidth
            printObj.Print "";
            printObj.CurrentX = (pntPos(TTLBegin) + pntPos(PosBegin) + colWidth - printObj.TextWidth("Ttl")) / 2
            printObj.Print "Ttl";
            printObj.CurrentX = (pntPos(PosBegin) + pntPos(GeldBegin) + colWidth - printObj.TextWidth("Pos")) / 2
            printObj.Print "Pos";
            printObj.CurrentX = (pntPos(GeldBegin) + pntPos(GeldBegin + 1) + colWidth - printObj.TextWidth("Geld")) / 2
            printObj.Print "Geld";
            fontSizing 8
            printObj.CurrentY = top2Ypos
            printObj.CurrentX = pntPos(UBound(pntPos)) + colWidth
            printObj.Print
            printObj.Line (0, printObj.CurrentY)-(printObj.ScaleWidth - 50, printObj.CurrentY)
          End If
        End If
      Loop
    End With
    'finish of with the vertical lines etc
    botY = printObj.CurrentY
    printObj.Line (pntPos(1) + 75, topYpos)-(pntPos(1) + 75, top2Ypos)
    printObj.Line (pntPos(grpStndBegin) + 75, topYpos)-(pntPos(6) + 75, top2Ypos)
    If grpCount > 4 Then
        printObj.Line (pntPos(fin8Begin) + 75, topYpos)-(pntPos(fin8Begin) + 75, top2Ypos)
    End If
    printObj.Line (pntPos(fin4Begin) + 75, topYpos)-(pntPos(fin4Begin) + 75, top2Ypos)
    printObj.Line (pntPos(fin2Begin) + 75, topYpos)-(pntPos(fin2Begin) + 75, top2Ypos)
    printObj.Line (pntPos(finBegin) + 75, topYpos)-(pntPos(finBegin) + 75, top2Ypos)
    printObj.Line (pntPos(EindstBegin) + 75, topYpos)-(pntPos(EindstBegin) + 75, top2Ypos)
    printObj.Line (pntPos(AantBegin) + 75, topYpos)-(pntPos(AantBegin) + 75, top2Ypos)
    printObj.Line (pntPos(TopScBegin) + 75, topYpos)-(pntPos(TopScBegin) + 75, top2Ypos)
    printObj.Line (pntPos(TTLBegin) + 75, topYpos)-(pntPos(TTLBegin) + 75, top2Ypos)
    printObj.Line (pntPos(PosBegin) + 75, topYpos)-(pntPos(PosBegin) + 75, top2Ypos)
    printObj.Line (pntPos(GeldBegin) + 75, topYpos)-(pntPos(GeldBegin) + 75, top2Ypos)
    For i = 1 To UBound(pntPos) - 1
        printObj.Line (pntPos(i) + 75, top2Ypos)-(pntPos(i) + 75, botY)
    Next
    printObj.Line (printObj.ScaleWidth - 50, topYpos)-(printObj.ScaleWidth - 50, botY)
End Sub

Function printAant(poolFormID As Long, pos As Integer, fld As String)
Dim aant As Integer
Dim pnt As Long
    Select Case fld
    Case "ptsHt"
    pnt = getPointsForID(1, cn)
    Case "ptsFt"
    pnt = getPointsForID(2, cn)
    Case "ptsToto"
    pnt = getPointsForID(3, cn)
    Case "ptsDayGoals"
    pnt = getPointsForID(28, cn)
    End Select
    If LCase(Left(fld, 9)) = "pointsgrp" Then
        pnt = getPointsForID(8, cn)
    End If

    aant = getAant(poolFormID, fld)
'    printObj.CurrentX = pos - printObj.TextWidth("(" & Format(aant, "0") & "x) " & Format(aant * pnt, "0"))
    printObj.CurrentX = pos - printObj.TextWidth(Format(aant * pnt, "0"))
    printObj.FontItalic = True
'    printObj.Print "(" & Format(aant, "0"); "x) ";
    printObj.FontItalic = False
    printObj.Print Format(aant * pnt, "0");
    printAant = aant * pnt
End Function
'
Sub printPoolStandings(matchOrder As Integer)
Dim lastPos As Integer
Dim colpos(6)
Dim alfabet As Boolean
Dim poolColWidth As Integer
alfabet = True

If Me.poolFormOrder(1) Then alfabet = False

  Set rs = New ADODB.Recordset
  'leftMargin = printObj.CurrentX
  poolColWidth = (printObj.ScaleWidth + 2 * printObj.ScaleLeft) \ 2
  fontSizing 10
  colpos(0) = printObj.TextWidth("999.")
  colpos(1) = colpos(0) + poolColWidth / 4 - 200
  colpos(2) = colpos(1) + poolColWidth / 10
  colpos(3) = colpos(2) + poolColWidth / 10
  colpos(4) = colpos(3) + poolColWidth / 6 + 200
  colpos(5) = colpos(4) + poolColWidth / 6 - 100
  colpos(6) = colpos(5) + poolColWidth / 6 - 100

  If alfabet Then
      colpos(0) = Me.CurrentX + 40
  End If

'  fontSizing 16
'  printObj.FontBold = True
  
  InitPage True, False, 0, 0, 0, 0, True
  printPoolFormStandingLegenda
' bepaal de laatste plaats
  rs.Open getPoolFormSql(False, matchOrder), cn, adOpenStatic, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveLast
    lastPos = nz(rs!pointsGrandTotal, 0)
  Else
    Exit Sub
  End If
  rs.Close
  Set rs = Nothing
 
  printObj.CurrentX = 0
  
  'druk de tabel af
  poolstandingsTable alfabet, colpos(), printObj.CurrentY, matchOrder, lastPos
  
  If Me.chkCombi Then 'tabel op andere volgorder op hetzelfde velletje
    printObj.Print
    alfabet = Not alfabet
    poolstandingsTable alfabet, colpos(), printObj.CurrentY, matchOrder, lastPos
  End If
End Sub

Sub poolstandingsTable(alfabet As Boolean, colpos() As Variant, poolTopPos As Integer, matchOrder As Integer, lastPos As Integer)
  Dim col As Integer
  Dim lastTtl As Integer
  Dim i As Integer
  Dim poolColWidth As Integer
  Dim prStr As String
  Dim pts As Integer
  Dim moneyPrev As Currency
  Dim moneySum As Currency
  Dim leftMargin As Integer
  Dim yLinePos As Integer
  Dim curYpos As Integer
  Dim savy As Integer
  Dim tmp As String
  Dim txtStr As String
  Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If alfabet Then
      txtStr = "Stand alfabetisch na " & matchOrder & "e wedstrijd: " & getMatchDescription(matchOrder, cn, True, True, False, True) '& ": " & getMatchResultPartStr(matchOrder, 2, cn)
      colpos(0) = printObj.CurrentX + 40
    Else
      txtStr = "Stand op punten na " & matchOrder & "e wedstrijd " & getMatchDescription(matchOrder, cn) & ", uitslag: " & getMatchresultStr(matchOrder, True, cn)
      colpos(0) = printObj.TextWidth("999.")
    End If
    If Me.chkEindstand Then
      If alfabet Then
          txtStr = "Eindstand alfabetisch"
      Else
          txtStr = "Eindstand"
      End If
    End If
    'subHeading txtStr, True
    reportTitle txtStr, False, False, printObj.CurrentY, 0, False
    poolTopPos = printObj.CurrentY
    fontSizing 10
    
     rs.Open getPoolFormSql(alfabet, matchOrder), cn, adOpenStatic, adLockReadOnly

    leftMargin = printObj.CurrentX
    poolColWidth = (printObj.ScaleWidth + 2 * printObj.ScaleLeft) \ 2
    
    For col = 0 To 1
      If Not alfabet Then
        printObj.CurrentX = col * poolColWidth
        'printObj.Print "pos";
      End If
      printObj.CurrentX = colpos(0) + col * poolColWidth
      printObj.Print "Naam";
      If alfabet Then printObj.Print " (pl)";
      printObj.CurrentX = colpos(1) + col * poolColWidth
      printObj.Print "had  +";
      printObj.CurrentX = colpos(2) + col * poolColWidth
      printObj.Print "erbij =";
      printObj.CurrentX = colpos(3) + col * poolColWidth + printObj.TextWidth("999") - printObj.TextWidth("nu")
      printObj.Print "nu";
      printObj.CurrentX = colpos(4) - printObj.TextWidth("Geld") + col * poolColWidth
      printObj.Print "Geld";
      printObj.CurrentX = colpos(5) - printObj.TextWidth("erbij") + col * poolColWidth
      printObj.Print "erbij";
      printObj.CurrentX = colpos(6) - printObj.TextWidth("totaal") + col * poolColWidth
      printObj.Print "totaal";
    Next
    printObj.CurrentY = printObj.CurrentY + 50
    yLinePos = printObj.CurrentY + TextHeight("test")
    On Error Resume Next
    printObj.Line (leftMargin, yLinePos)-(printObj.ScaleWidth, yLinePos)
    On Error GoTo 0
    printObj.CurrentY = printObj.CurrentY + 50
    poolTopPos = printObj.CurrentY
    printObj.CurrentX = 0
    With rs
      If .RecordCount > 0 Then
        .MoveFirst
        lastTtl = 0
        col = 0
        i = 0
        Do While Not .EOF
        
          i = i + 1
          'If !competitorPoolID = 23 Then Stop
          If i = Int(.RecordCount / 2 + 0.5) + 1 Then
            col = poolColWidth
            printObj.CurrentY = poolTopPos
          End If
          printObj.CurrentX = printObj.CurrentX + colpos(0) - printObj.TextWidth(!positionTotal) - printObj.TextWidth("..") + col
          If Not alfabet Then
            If lastTtl <> !pointsGrandTotal Then printObj.Print !positionTotal & ".";
          End If
          printObj.FontBold = !positionTotal = 1
          printObj.FontItalic = nz(!pointsGrandTotal, 0) = lastPos
          prStr = Left(!nickName, 12)
          If alfabet Then
            prStr = prStr & " (" & !positionTotal & ")"
          End If
          If !pointsGrandTotal = lastPos Then
            printObj.ForeColor = vbRed
          ElseIf nz(!positionTotal, 0) = 1 Then
            printObj.ForeColor = vbBlue
          ElseIf nz(!positionDay, 0) = 1 Then
            printObj.ForeColor = &H8000&
          Else
            printObj.ForeColor = 0
          End If
          printObj.CurrentX = colpos(0) + col
          printObj.FontUnderline = nz(!positionDay, 0) = 1
  
          printObj.Print prStr;
          printObj.FontBold = False
          printObj.FontItalic = False
          printObj.ForeColor = 0
          printObj.FontUnderline = False
          If matchOrder > 1 Then
              pts = getTotalPointsPrevDay(!poolFormID, matchOrder, cn) 'lastPos total
              moneyPrev = getTotalMoney(!poolFormID, getMatchPrevDay(matchOrder, cn)) 'lastPos money total
          Else
              pts = 0
              moneyPrev = 0
          End If
  
          printObj.CurrentX = colpos(1) + col + printObj.TextWidth("999") - printObj.TextWidth(CStr(pts))
          printObj.Print Format(pts, "##0");
          printObj.FontBold = False
          pts = getTotalDayPoints(!competitorPoolID, matchOrder, getMatchInfo(matchOrder, "matchDate", cn), cn)
          printObj.CurrentX = colpos(2) + col + printObj.TextWidth("999") - printObj.TextWidth(CStr(pts))
          printObj.FontUnderline = nz(!positionDay, 0) = 1
          If !positionDay = 1 Then
            printObj.ForeColor = &H8000&
          Else
            printObj.ForeColor = 0
          End If
          printObj.Print Format(pts, "##0");
          printObj.ForeColor = 0
          printObj.FontUnderline = False
          printObj.FontBold = !positionTotal = 1
          printObj.FontItalic = nz(!pointsGrandTotal, 0) = lastPos
          pts = nz(!pointsGrandTotal, 0)
          If !pointsGrandTotal = lastPos Then
              printObj.ForeColor = vbRed
          ElseIf !positionTotal = 1 Then
              printObj.ForeColor = vbBlue
          Else
              printObj.ForeColor = 0
          End If
          printObj.CurrentX = colpos(3) + col + printObj.TextWidth("999") - printObj.TextWidth(CStr(pts))
          If !pointsGrandTotal = lastPos Then
              printObj.ForeColor = &H80&
          ElseIf !positionTotal = 1 Then
              printObj.ForeColor = &HC00000
          Else
              printObj.ForeColor = 0
          End If
          printObj.Print Format(!pointsGrandTotal, "##0");
          printObj.ForeColor = 0
          printObj.FontBold = False
          printObj.FontItalic = False
          tmp = Format(moneyPrev, "€ ##0.00")
          printObj.CurrentX = colpos(4) - printObj.TextWidth(tmp) + col
          printObj.Print tmp;   '= geld
          tmp = Format(!moneyDaytotal, "€ ##0.00")
          printObj.CurrentX = colpos(5) - printObj.TextWidth(tmp) + col
          printObj.Print tmp;
          moneySum = 0
          tmp = Format(moneyPrev + !moneyDaytotal, "€ ##0.00")
          printObj.CurrentX = colpos(6) - printObj.TextWidth(tmp) + col
          printObj.Print tmp;   '= geld
          printObj.Print
          lastTtl = nz(!pointsGrandTotal, 0)
          rs.MoveNext
        Loop
        curYpos = printObj.CurrentY
        printObj.Line (poolColWidth, yLinePos)-(poolColWidth, curYpos)
        printObj.Line (colpos(4) - printObj.TextWidth("Geld") - 400, yLinePos)-(colpos(4) - printObj.TextWidth("Geld") - 400, curYpos)
        printObj.Line (colpos(4) - printObj.TextWidth("Geld") - 400 + poolColWidth, yLinePos)-(colpos(4) - printObj.TextWidth("Geld") - 400 + poolColWidth, curYpos)
        printObj.Line (leftMargin, curYpos)-(printObj.ScaleWidth + printObj.ScaleLeft * 2, curYpos)
      End If
      .Close
    End With
End Sub

Sub printPoolFormStandingLegenda()
    printObj.FontItalic = False
    printObj.FontBold = False
    fontSizing 10
    printObj.CurrentX = (printObj.ScaleWidth - printObj.TextWidth("onderstreept=daghoogste, vet=bovenaan, cursief=onderaan")) / 2
    printObj.Print "(";
    printObj.FontUnderline = True
    printObj.ForeColor = &H8000&
    printObj.Print "onderstreept";
    printObj.FontUnderline = False
    printObj.ForeColor = 0
    printObj.Print "= daghoogste, ";
    printObj.ForeColor = vbBlue
    printObj.FontBold = True
    printObj.Print "vet";
    printObj.FontBold = False
    printObj.ForeColor = 0
    printObj.Print "= bovenaan, ";
    printObj.FontItalic = True
    printObj.ForeColor = vbRed
    printObj.Print "cursief";
    printObj.FontItalic = False
    printObj.ForeColor = 0
    printObj.Print "= onderaan)"
End Sub

'
Function getPoolFormSql(alfabet As Boolean, matchOrdrNr As Integer) As String
Dim sqlstr As String
    sqlstr = "SELECT c.competitorpoolid as poolFormID, c.nickName as nickname, p.*"
    sqlstr = sqlstr & " FROM tblCompetitorPools c INNER JOIN tblCompetitorPoints p "
    sqlstr = sqlstr & " ON p.competitorPoolID = c.competitorPoolID"
    sqlstr = sqlstr & " WHERE p.matchOrder = " & matchOrdrNr
    sqlstr = sqlstr & " AND c.poolID = " & thisPool
    If alfabet Then
        sqlstr = sqlstr & " ORDER BY nickname"
    Else
        sqlstr = sqlstr & " ORDER BY p.pointsGrandTotal DESC, c.nickname ASC"
    End If
    getPoolFormSql = sqlstr

End Function
'
Private Sub lstCompetitorPools_Click()
    Me.optSelection.value = True
End Sub


Private Sub pageFooter()
Dim printWidth
Dim i As Double
Dim fontnaam As String
Dim yPos As Integer
Dim fontHeight As Integer
    printObj.ForeColor = headBGcolor
    printWidth = printObj.DrawWidth
    printObj.DrawWidth = 2
    printObj.FontItalic = True
    printObj.FontBold = False
    fontnaam = printObj.FontName
    printObj.FontName = "Garamond"
    fontSizing 10
    yPos = printObj.ScaleHeight - printObj.TextHeight("w")
    printObj.Line (0, yPos - 15)-(printObj.ScaleWidth, yPos - 15)
    printObj.CurrentY = yPos + 12
    fontSizing 8
    centerText "vbPool 2.0 © 2004-" & Year(Now) & " jota services"
    printObj.FontName = fontnaam
    printObj.FontBold = False
    printObj.FontItalic = False
    yPos = printObj.CurrentY + 50
    On Error Resume Next
    'printObj.Line (0, yPos)-(printObj.ScaleWidth, yPos)
    On Error GoTo 0
    printObj.ForeColor = vbBlack
    printObj.DrawWidth = printWidth
End Sub

'
'
Sub printPointInfo(inclpnt As Boolean)
Dim infostr As String
Dim pntToto As Integer
Dim pntRust As Integer
Dim pntEind As Integer
Dim pntDp As Integer
pntToto = getPointsForID(3, cn)
pntRust = getPointsForID(1, cn)
pntEind = getPointsForID(2, cn)
pntDp = getPointsForID(28, cn)
'get the length of the info str to center the text
infostr = "Samenstelling punten: rust goed "
If inclpnt Then infostr = infostr & pntRust & ":pnt"
infostr = infostr & ", eindstand goed "
If inclpnt Then infostr = infostr & pntEind & ":pnt"
infostr = infostr & ", toto goed "
If inclpnt Then infostr = infostr & pntToto & ":pnt"
infostr = infostr & ", aantal doelpunten van de dag goed "
If inclpnt Then infostr = infostr & pntDp & ":pnt"

fontSizing 10
printObj.CurrentX = (printObj.ScaleWidth - printObj.TextWidth(infostr)) / 2

printObj.Print "Samenstelling punten: ";
printObj.FontUnderline = True
printObj.Print "rust goed";
printObj.FontUnderline = False
If inclpnt Then printObj.Print ": " & pntRust; "pnt";
printObj.Print ", ";
printObj.FontBold = True
printObj.Print "eindstand goed";
printObj.FontBold = False
If inclpnt Then printObj.Print ": " & pntEind; "pnt";
printObj.Print ", ";
printObj.FontItalic = True
printObj.Print "toto goed";
printObj.FontItalic = False
If inclpnt Then printObj.Print ": " & pntToto; "pnt";
printObj.Print ", ";
printObj.ForeColor = vbBlue
printObj.Print "aantal doelpunten van de dag goed";
printObj.ForeColor = 1
If inclpnt Then printObj.Print ": " & pntDp; "pnt"
printObj.CurrentY = printObj.CurrentY + 50
'
'
End Sub
'

Sub printPoolPointsPerMatch()
''print de deelnemers en hun punten per matchStr
Dim rs As New ADODB.Recordset
Dim rsPnt As New ADODB.Recordset
Dim rsWeds As New ADODB.Recordset
Dim matchResult As String
Dim sqlstr As String
Dim xPos As Integer
Dim posX() As Integer
Dim X As Integer
Dim i As Integer
Dim colWidthFactor As Double
Dim topY As Integer
Dim botY As Integer
Dim topYpos As Integer
Dim colWidth As Integer
Dim ttlColWidth As Integer
Dim matchStr As String
Dim pntFormat As String
Dim moneyStr As String
Dim lastPosition As Integer
Dim naam As String
Dim headTxt() As String
Dim headCnt As Integer
Dim grpTtl As Integer
Dim fin8ttl As Integer
Dim fin4ttl As Integer
Dim fin2ttl As Integer
Dim fin34ttl As Integer
Dim matchNr As Integer
Dim txtStr As String
Dim verttxtHeight 'de hoogte van de verticale text bovenin
Dim infostr As String
  matchNr = getMatchNumber(toMatch, cn)
  txtStr = "Punten per wedstrijd t/m " & getMatchDescription(toMatch, cn, True, True) '& ", Uitslag: " & getMatchResultPartStr(matchNr, 2, cn)

  'headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool"
  lastPosition = getLastPoolFormPosition(toMatch, cn)
  heading1 = txtStr
  InitPage True, False, , True
  printObj.CurrentY = printObj.CurrentY - 50
  topYpos = printObj.CurrentY
  printPointInfo True 'druk de inforegel over de punten toekenning af
  topY = printObj.CurrentY
  printObj.Line (0, topY)-(printObj.ScaleWidth - 50, topY)
  fontSizing 8
  'sqlstr = "SELECT pooldeelnems.deelnemID, pooldeelnems.bijnaam, deelnempnt.grandTotaal"
  sqlstr = "Select c.competitorPoolID, c.nickname, p.pointsGrandTotal"
  sqlstr = sqlstr & " FROM (tblCompetitorPools c INNER JOIN tblCompetitorPoints p ON c.competitorPoolID = p.competitorPoolID) "
  sqlstr = sqlstr & " INNER JOIN tblTournamentSchedule s ON p.matchOrder= s.matchOrder"
  sqlstr = sqlstr & " Where c.poolid= " & thisPool
  sqlstr = sqlstr & " And s.matchOrder = " & toMatch
  sqlstr = sqlstr & " And s.tournamentID = " & thisTournament
  If Me.poolFormOrder(1) = True Then
      sqlstr = sqlstr & " order by pointsGrandTotal DESC"
  Else
      sqlstr = sqlstr & " order by nickname"
  End If
  '
  rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
  sqlstr = "Select * from tblTournamentSchedule where tournamentID=" & thisTournament
  sqlstr = sqlstr & " order by matchOrder"
  rsWeds.Open sqlstr, cn, adOpenStatic, adLockReadOnly
  verttxtHeight = printObj.TextWidth("12345678901234567890123456789")
  printObj.CurrentY = verttxtHeight
  fontSizing 12
  printObj.CurrentX = printObj.TextWidth(Left(getLongestNickName(cn), 10))
  fontSizing 10
  ReDim posX(1)
  posX(1) = printObj.CurrentX
  colWidthFactor = 0.92 * 64 / getMatchCount(0, cn)
  With rsWeds
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        rotater.Angle = 90
        printObj.CurrentX = posX(UBound(posX)) + 2
        matchStr = getMatchDescription(!matchOrder, cn, False, False, True)
        matchResult = getMatchresultStr(!matchOrder, True, cn)
        fontSizing 9
        If matchResult > "" Then
          rotater.PrintText matchStr & ": " & matchResult
        Else
          rotater.PrintText matchStr
        End If
        rotater.Angle = 0
        fontSizing 10
        xPos = printObj.CurrentX + printObj.TextWidth("99") * colWidthFactor
        ReDim Preserve posX(UBound(posX) + 1)
        posX(UBound(posX)) = xPos
        rsWeds.MoveNext
        'Debug.Print UBound(posX), posX(UBound(posX))
      Loop
    End If
  End With
  rotater.Angle = 90
  printObj.CurrentX = posX(UBound(posX))
  headCnt = headCnt + 1
  ReDim Preserve headTxt(headCnt)
  headTxt(headCnt) = " groepstanden"
  rotater.PrintText headTxt(headCnt)
  
  If getTournamentInfo("tournamentgroupCount", cn) > 4 Then
    xPos = printObj.CurrentX + printObj.TextWidth("199") * colWidthFactor
    ReDim Preserve posX(UBound(posX) + 1)
    posX(UBound(posX)) = xPos
    rotater.Angle = 90
    printObj.CurrentX = posX(UBound(posX))
    headCnt = headCnt + 1
    ReDim Preserve headTxt(headCnt)
    headTxt(headCnt) = " 8e Finales"
    rotater.PrintText headTxt(headCnt)
  End If
  xPos = printObj.CurrentX + printObj.TextWidth("199") * colWidthFactor
  ReDim Preserve posX(UBound(posX) + 1)
  posX(UBound(posX)) = xPos
  rotater.Angle = 90
  printObj.CurrentX = posX(UBound(posX))
  
  headCnt = headCnt + 1
  ReDim Preserve headTxt(headCnt)
  headTxt(headCnt) = " Kw Finales"
  rotater.PrintText headTxt(headCnt)
  
  xPos = printObj.CurrentX + printObj.TextWidth("99") * colWidthFactor
  ReDim Preserve posX(UBound(posX) + 1)
  posX(UBound(posX)) = xPos
  rotater.Angle = 90
  printObj.CurrentX = posX(UBound(posX))
  headCnt = headCnt + 1
  ReDim Preserve headTxt(headCnt)
  headTxt(headCnt) = " Hv Finales"
  rotater.PrintText headTxt(headCnt)
  
  xPos = printObj.CurrentX + printObj.TextWidth("99") * colWidthFactor
  ReDim Preserve posX(UBound(posX) + 1)
  posX(UBound(posX)) = xPos
  rotater.Angle = 90
  printObj.CurrentX = posX(UBound(posX))
  headCnt = headCnt + 1
  ReDim Preserve headTxt(headCnt)
  headTxt(headCnt) = " Finales"
  rotater.PrintText headTxt(headCnt)
  
  xPos = printObj.CurrentX + printObj.TextWidth("99") * colWidthFactor
  ReDim Preserve posX(UBound(posX) + 1)
  posX(UBound(posX)) = xPos
  rotater.Angle = 90
  printObj.CurrentX = posX(UBound(posX))
  headCnt = headCnt + 1
  ReDim Preserve headTxt(headCnt)
  headTxt(headCnt) = " Eindstand"
  rotater.PrintText headTxt(headCnt)
  
  xPos = printObj.CurrentX + printObj.TextWidth("99") * colWidthFactor
  ReDim Preserve posX(UBound(posX) + 1)
  posX(UBound(posX)) = xPos
  rotater.Angle = 90
  printObj.CurrentX = posX(UBound(posX))
  headCnt = headCnt + 1
  ReDim Preserve headTxt(headCnt)
  headTxt(headCnt) = " Topscorers"
  rotater.PrintText headTxt(headCnt)
    
  xPos = printObj.CurrentX + printObj.TextWidth("99") * colWidthFactor
  ReDim Preserve posX(UBound(posX) + 1)
  posX(UBound(posX)) = xPos
  rotater.Angle = 90
  printObj.CurrentX = posX(UBound(posX))
  headCnt = headCnt + 1
  ReDim Preserve headTxt(headCnt)
  headTxt(headCnt) = " Stats"
  rotater.PrintText headTxt(headCnt)
  
  xPos = printObj.CurrentX + printObj.TextWidth("99") * colWidthFactor
  ReDim Preserve posX(UBound(posX) + 1)
  posX(UBound(posX)) = xPos
  rotater.Angle = 90
  printObj.CurrentX = posX(UBound(posX))
  headCnt = headCnt + 1
  ReDim Preserve headTxt(headCnt)
  headTxt(headCnt) = " Totaal"
  rotater.PrintText headTxt(headCnt)
  
  xPos = printObj.CurrentX + printObj.TextWidth("999") * colWidthFactor
  ReDim Preserve posX(UBound(posX) + 1)
  posX(UBound(posX)) = xPos
  rotater.Angle = 90
  printObj.CurrentX = posX(UBound(posX))
  headCnt = headCnt + 1
  ReDim Preserve headTxt(headCnt)
  headTxt(headCnt) = " Positie"
  rotater.PrintText headTxt(headCnt)
  
  xPos = printObj.CurrentX + printObj.TextWidth("99") * colWidthFactor
  ReDim Preserve posX(UBound(posX) + 1)
  posX(UBound(posX)) = xPos
  rotater.Angle = 90
  headCnt = headCnt + 1
  ReDim Preserve headTxt(headCnt)
  headTxt(headCnt) = "geld"
  printObj.CurrentX = printObj.ScaleWidth - 100 - printObj.TextWidth("geld")
  printObj.CurrentY = verttxtHeight - printObj.TextHeight("Geld")
  printObj.Print headTxt(headCnt);
    
  xPos = printObj.CurrentX + printObj.TextWidth("geld") * colWidthFactor
  printObj.Print
  topYpos = printObj.CurrentY + 50
  ReDim Preserve posX(UBound(posX) + 1)
  posX(UBound(posX)) = xPos
  printObj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
  printObj.CurrentY = topYpos
  printObj.CurrentX = 0
  colWidth = posX(2) - posX(1)
  botY = printObj.CurrentY
  
  Do While Not rs.EOF
    pntFormat = "0;;\ ;-"
    naam = rs!nickName
    fontSizing 10
    Do While printObj.TextWidth(naam) > printObj.TextWidth("1234567890")
        naam = Left(naam, Len(naam) - 1)
    Loop
    fontSizing 10
    
    sqlstr = "SELECT s.matchtime, p.*, s.matchplayed"
    sqlstr = sqlstr & " FROM tblCOmpetitorPoints p INNER JOIN tblTournamentSchedule s ON p.matchOrder= s.matchOrder"
    sqlstr = sqlstr & " Where s.matchOrder <=" & toMatch
    sqlstr = sqlstr & " AND s.tournamentid = " & thisTournament
    sqlstr = sqlstr & " AND p.competitorpoolID = " & rs!competitorPoolID
    sqlstr = sqlstr & " ORDER BY s.matchOrder"
    rsPnt.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    i = 0
    With rsPnt
      rotater.Angle = 90
      printObj.ForeColor = vbBlack
      fontSizing 10
      printObj.Print naam;
      printObj.ForeColor = vbBlack
      fontSizing 8
      Do While Not .EOF
        i = i + 1
        printObj.CurrentX = posX(i) + (colWidth - printObj.TextWidth(Format(nz(!pointsDay, 0), pntFormat))) / 2
        printObj.FontItalic = nz(!ptsToto, 0) <> 0
        printObj.FontBold = nz(!ptsFt, 0) <> 0
        printObj.FontUnderline = nz(!ptsHt, 0) > 0
        If nz(!ptsDayGoals, 0) > 0 Then
            printObj.ForeColor = vbBlue
        End If
        printObj.Print Format(nz(!pointsDay, 0), pntFormat);
        printObj.FontBold = False
        printObj.FontItalic = False
        printObj.FontUnderline = False
        printObj.ForeColor = 1

        .MoveNext
        rotater.Angle = 90
      Loop
      If Not .RecordCount = 0 Then
        .MoveLast
        If !positionTotal = 1 Then
            printObj.ForeColor = &HC00000
            printObj.FontBold = True
        Else
            printObj.ForeColor = vbBlack
            printObj.FontBold = False
        End If
        ttlColWidth = posX(UBound(posX) - 10) - posX(UBound(posX) - 11)
        If toMatch = getMatchCount(0, cn) Then pntFormat = "0"
        grpTtl = getTotalPointsForFieldUntilMatch(rs!competitorPoolID, toMatch, "pointsGroupStanding", cn)
        fin8ttl = getTotalPointsForFieldUntilMatch(rs!competitorPoolID, toMatch, "pointsFinals_8", cn)
        fin4ttl = getTotalPointsForFieldUntilMatch(rs!competitorPoolID, toMatch, "pointsFinals_4", cn)
        fin2ttl = getTotalPointsForFieldUntilMatch(rs!competitorPoolID, toMatch, "pointsFinals_2", cn)
        printObj.CurrentX = posX(UBound(posX) - 11) + (ttlColWidth - printObj.TextWidth(Format(grpTtl, pntFormat))) / 2
        printObj.Print Format(grpTtl, pntFormat);
        ttlColWidth = posX(UBound(posX) - 9) - posX(UBound(posX) - 10)
        If getTournamentInfo("tournamentGroupCount", cn) > 4 Then
            printObj.CurrentX = posX(UBound(posX) - 10) + (ttlColWidth - printObj.TextWidth(Format(fin8ttl, pntFormat))) / 2
            printObj.Print Format(fin8ttl, pntFormat);
            ttlColWidth = posX(UBound(posX) - 8) - posX(UBound(posX) - 9)
        End If
        printObj.CurrentX = posX(UBound(posX) - 9) + (ttlColWidth - printObj.TextWidth(Format(fin4ttl, pntFormat))) / 2
        printObj.Print Format(fin4ttl, pntFormat);
        ttlColWidth = posX(UBound(posX) - 7) - posX(UBound(posX) - 8)
        printObj.CurrentX = posX(UBound(posX) - 8) + (ttlColWidth - printObj.TextWidth(Format(fin2ttl, pntFormat))) / 2
        printObj.Print Format(fin2ttl, pntFormat);
        ttlColWidth = posX(UBound(posX) - 6) - posX(UBound(posX) - 7)
        printObj.CurrentX = posX(UBound(posX) - 7) + (ttlColWidth - printObj.TextWidth(Format(nz(!pointsFinal, 0) + nz(!pointsFinals_34, 0), pntFormat))) / 2
        printObj.Print Format(nz(!pointsFinal, 0) + nz(!pointsFinals_34, 0), pntFormat);
        ttlColWidth = posX(UBound(posX) - 5) - posX(UBound(posX) - 6)
        printObj.CurrentX = posX(UBound(posX) - 6) + (ttlColWidth - printObj.TextWidth(Format(!pointsFinalStanding, pntFormat))) / 2
        printObj.Print Format(!pointsFinalStanding, pntFormat);
        ttlColWidth = posX(UBound(posX) - 4) - posX(UBound(posX) - 5)
        printObj.CurrentX = posX(UBound(posX) - 5) + (ttlColWidth - printObj.TextWidth(Format(nz(!pointsTopscorers, 0) + nz(!pointsOther, 0), pntFormat))) / 2
        printObj.Print Format(nz(!pointsTopscorers, 0), pntFormat);
        ttlColWidth = posX(UBound(posX) - 3) - posX(UBound(posX) - 4)
        printObj.CurrentX = posX(UBound(posX) - 4) + (ttlColWidth - printObj.TextWidth(Format(nz(!pointsOther, 0) + nz(!pointsTopscorers, 0), pntFormat))) / 2
        printObj.Print Format(nz(!pointsOther, 0), pntFormat);
        ttlColWidth = posX(UBound(posX) - 2) - posX(UBound(posX) - 3)
        printObj.CurrentX = posX(UBound(posX) - 3) + (ttlColWidth - printObj.TextWidth(Format(nz(!pointsGrandTotal, 0), pntFormat))) / 2
        printObj.Print Format(nz(!pointsGrandTotal, 0), pntFormat);
        ttlColWidth = posX(UBound(posX) - 1) - posX(UBound(posX) - 2)
        printObj.CurrentX = posX(UBound(posX) - 2) + (ttlColWidth - printObj.TextWidth(Format(nz(!positionTotal, 0), pntFormat))) / 2
        printObj.Print Format(nz(!positionTotal, 0), pntFormat);
        printObj.CurrentX = printObj.ScaleWidth - 50 - printObj.TextWidth(Format(nz(!moneyTotal, 0), "currency"))
        printObj.ForeColor = vbBlack
        printObj.FontItalic = False
        printObj.FontBold = False
        If IsLastMatchOfDay(toMatch, cn) Then
          If toMatch < getMatchCount(0, cn) Then
            moneyStr = Format(getTotalMoney(!competitorPoolID, getMatchPrevDay(toMatch + 1, cn)), "currency") 'lastPos money total'getPoolFormPoints(!competitorPoolID, matchNr, 50, cn)
          Else
            moneyStr = Format(getTotalMoney(!competitorPoolID, toMatch), "currency")  'lastPos money total'getPoolFormPoints(!competitorPoolID, matchNr, 50, cn)
          End If
        Else
          moneyStr = Format(getTotalMoney(!competitorPoolID, getMatchPrevDay(toMatch, cn)), "currency") 'lastPos money total'getPoolFormPoints(!competitorPoolID, matchNr, 50, cn)
        End If
        
        printObj.Print moneyStr;
      End If
      printObj.Print
    End With
    printObj.Line (0, printObj.CurrentY + 10)-(posX(UBound(posX)), printObj.CurrentY + 10)
    printObj.CurrentY = printObj.CurrentY + 10
    printObj.CurrentX = 0
    botY = printObj.CurrentY
    If botY >= footerPos And rs.AbsolutePosition < rs.RecordCount Then
      'nieuwe pagina
      'eerste de lijntjes
      For i = 1 To UBound(posX)
          printObj.Line (posX(i), topY)-(posX(i), botY)
      Next
      i = 0
      addNewPage True, False
      printObj.CurrentY = printObj.CurrentY - 50
      topYpos = printObj.CurrentY
      printPointInfo True 'druk de inforegel over de punten toekenning af
      topY = printObj.CurrentY
      printObj.Line (0, topY)-(printObj.ScaleWidth - 50, topY)
      fontSizing 12
      printObj.CurrentY = verttxtHeight
      printObj.CurrentX = printObj.TextWidth("123456789012345")
      fontSizing 10
      With rsWeds
        If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
            Set rotater.Device = printObj
            i = i + 1
            rotater.Angle = 90
            printObj.CurrentX = posX(i) + 2
            matchStr = getMatchDescription(!matchOrder, cn, False, False, True)
            matchResult = getMatchresultStr(!matchOrder, True, cn)
            fontSizing 9
            If matchResult > "" Then
              rotater.PrintText matchStr & ": " & matchResult
            Else
              rotater.PrintText matchStr
            End If
            rotater.Angle = 0
            fontSizing 10
            .MoveNext
          Loop
        End If
      End With
      rotater.Angle = 90
      If getTournamentInfo("tournamentgroupcount", cn) > 4 Then
          X = 11
      Else
          X = 10
      End If
      For i = 0 To UBound(headTxt) - 2
        printObj.CurrentX = posX(UBound(posX) - (X - i))
        rotater.PrintText headTxt(i + 1)
      Next
      printObj.CurrentX = printObj.ScaleWidth - 100 - printObj.TextWidth("geld")
      printObj.CurrentY = verttxtHeight - printObj.TextHeight("Geld")
      printObj.Print "geld"

'
'      x = x - 1
'      rotater.PrintText headTxt(x - 10)
'      If getTournamentInfo("tournamentGroupCOunt", cn) > 4 Then
'          printObj.CurrentX = posX(UBound(posX) - x)
'          x = x - 1
'          rotater.PrintText " 8e Finalisten"
'      End If
'      printObj.CurrentX = posX(UBound(posX) - x)
'      x = x - 1
'      rotater.PrintText " Kw Finalisten"
'      printObj.CurrentX = posX(UBound(posX) - x)
'      x = x - 1
'      rotater.PrintText " Hv Finalisten"
'      printObj.CurrentX = posX(UBound(posX) - x)
'      x = x - 1
'      rotater.PrintText " Finalisten"
'      printObj.CurrentX = posX(UBound(posX) - x)
'      x = x - 1
'      rotater.PrintText " Eindstand"
'      printObj.CurrentX = posX(UBound(posX) - x)
'      x = x - 1
'      rotater.PrintText " Topscorers"
'      printObj.CurrentX = posX(UBound(posX) - x)
'      x = x - 1
'      rotater.PrintText " Overigen"
'      printObj.CurrentX = posX(UBound(posX) - x)
'      x = x - 1
'      rotater.PrintText " Totaal"
'      printObj.CurrentX = posX(UBound(posX) - x)
'      x = x - 1
'      rotater.PrintText " positie"
'      printObj.CurrentX = posX(UBound(posX) - x)
'      printObj.CurrentY = verttxtHeight - printObj.TextHeight("Geld")
'      printObj.Print " geld"
      topYpos = printObj.CurrentY + 50
      printObj.Line (0, topYpos)-(posX(UBound(posX)), topYpos)
      printObj.CurrentY = topYpos
      printObj.CurrentX = 0
      i = i + 1
    End If
    rs.MoveNext
    rsPnt.Close
  Loop
  rs.Close
  For i = 1 To UBound(posX)
      printObj.Line (posX(i), topY)-(posX(i), botY)
  Next
  i = 0
  Set rs = Nothing
  Set rsPnt = Nothing
End Sub
'
Sub printMatchPredictions(matchOrder As Integer)
Dim sqlstr As String
Dim rs As New ADODB.Recordset
Dim rsDeeln As New ADODB.Recordset
Dim cloneRS As ADODB.Recordset
Dim findStr As String
'Dim xpos As Integer
Dim cols(4) As Integer
'Dim naampos
Dim rowCnt As Integer
Dim thisRow As Integer
Dim yStart As Integer
Dim lineXstart As Integer
Dim lineYstart As Integer
Dim lineXend As Integer
Dim lineYend As Integer
Dim colpos(3) As Integer
Dim col As Integer
Dim topColPos As Integer
Dim i As Integer
Dim matchNr As Integer
  matchNr = getMatchNumber(matchOrder, cn)
'    headerText = GetOrgNaam(thisPool) & " " & getTournamentInfo("toernooi") & " voetbalpool" & " - Voorspelling"
    cols(0) = 0
    cols(1) = printObj.ScaleWidth / 3
    cols(2) = printObj.ScaleWidth / 3 * 2
    cols(3) = printObj.ScaleWidth
    cols(4) = printObj.ScaleWidth
    col = 3
    
    heading1 = "Voorspellingen " & matchOrder & "e wedstrijd (" & getMatchDescription(matchOrder, cn, True, True, False) & ")"
    InitPage True, False
'
    'horline 0
    topColPos = printObj.CurrentY
    printObj.Print
    colpos(0) = 50
    colpos(1) = printObj.TextWidth("0-000")
    colpos(2) = colpos(1) + printObj.TextWidth("0-000")
    colpos(3) = colpos(2) + printObj.TextWidth("0-000")
    printObj.ForeColor = RGB(0, 51, 0)
    For i = 0 To col - 1
        printObj.CurrentX = cols(i) + colpos(0)
        printObj.Print "Rust";
        printObj.CurrentX = cols(i) + colpos(1)
        printObj.Print "Eind";
        printObj.CurrentX = cols(i) + colpos(2)
        printObj.Print "Toto";
        printObj.CurrentX = cols(i) + colpos(3)
        printObj.Print "Pool";
    Next
    printObj.ForeColor = 0
    printObj.Print
    'horline 0
    yStart = printObj.CurrentY
    sqlstr = "SELECT ftA, ftB, htA,htB,tt, matchorder "
    sqlstr = sqlstr & " FROM tblPrediction_Matchresults p INNER JOIN "
    sqlstr = sqlstr & " tblCompetitorPools c ON p.competitorpoolid = c.competitorpoolid"
    sqlstr = sqlstr & " WHERE matchorder =" & matchOrder
    sqlstr = sqlstr & " AND c.poolid= " & thisPool
    sqlstr = sqlstr & " GROUP BY p.ftA, p.ftB, p.htA, p.htB, p.tt, p.matchorder"
    sqlstr = sqlstr & " ORDER BY htA, htB, ftA, ftB, tt"
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    sqlstr = "SELECT ftA, ftB, htA,htB,tt, matchorder, nickname "
    sqlstr = sqlstr & " FROM tblPrediction_Matchresults p INNER JOIN "
    sqlstr = sqlstr & " tblCompetitorPools c ON p.competitorpoolid = c.competitorpoolid"
    sqlstr = sqlstr & " WHERE matchorder = " & matchOrder
    sqlstr = sqlstr & " AND poolid = " & thisPool
    sqlstr = sqlstr & " ORDER BY nickName"
    rsDeeln.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    rsDeeln.MoveLast
    rowCnt = Int(rsDeeln.RecordCount / col + 0.5) + 1
    rsDeeln.MoveFirst
    rs.MoveFirst
    i = 0
    Do While Not rs.EOF
        Set cloneRS = rsDeeln.Clone
        findStr = "ftA = " & rs!ftA
        findStr = findStr & " and ftB = " & rs!ftB
        findStr = findStr & " and htA = " & rs!htA
        findStr = findStr & " and htB = " & rs!htB
        findStr = findStr & " and tt = " & rs!tt
        cloneRS.Filter = findStr
        If cloneRS.EOF Or cloneRS.BOF Then
            rsDeeln.MoveLast
            rsDeeln.MoveNext
        End If
        printObj.CurrentX = cols(i)
        lineXstart = printObj.CurrentX
        lineYstart = printObj.CurrentY
        printObj.CurrentX = cols(i) + colpos(0)
        printObj.Print rs!htA & "-" & rs!htB;
        printObj.CurrentX = cols(i) + colpos(1)
        printObj.FontBold = True
        printObj.Print rs!ftA & "-" & rs!ftB;
        printObj.FontBold = False
        printObj.CurrentX = cols(i) + colpos(2)
        printObj.Print rs!tt;
        cloneRS.MoveFirst
        Do While Not cloneRS.EOF
            printObj.CurrentX = cols(i) + colpos(3)
            printObj.Print cloneRS!nickName
            cloneRS.MoveNext
            thisRow = thisRow + 1
        Loop
        lineXend = cols(i + 1) - 100
        lineYend = printObj.CurrentY
        printObj.Line (lineXstart, lineYstart)-(lineXend, lineYend), lineColor, B
        rs.MoveNext
        If thisRow >= rowCnt Then
            i = i + 1
            printObj.CurrentY = yStart
            thisRow = 0
        End If
        cloneRS.Close
        Set cloneRS = Nothing
    Loop
    rs.Close
    rsDeeln.Close
    Set rs = Nothing
    Set rsDeeln = Nothing
End Sub

Sub SetForeCol(kl As Long)
Dim r As Integer
Dim g As Integer
Dim b As Integer
    r = &HFF& And kl
    g = (&HFF00& And kl) \ 256
    b = (&HFF0000 And kl) \ 65536
    If r * 0.3 + g * 0.59 + b * 0.11 < 128 Then
        printObj.ForeColor = vbWhite
    Else
        printObj.ForeColor = vbBlack
    End If

End Sub

Sub MakeColors()
Dim i As Integer
Dim A As Integer
Dim C As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim klCol As Collection
Dim forecol As Integer
    Set klCol = New Collection
    For r = 0 To 255 Step 63
       For g = 0 To 255 Step 63
         For b = 0 To 255 Step 63
            klCol.Add RGB(r, g, b)
         Next
       Next
    Next
For A = 0 To 64
    i = Int(Rnd() * klCol.Count) + 1
    thisColor(A) = klCol(i)
    klCol.Remove i
Next
End Sub

Function fitText(toWidth As Integer, txt As String)
'cut of text if to wide
Dim tmpTxt As String
  tmpTxt = txt
  Do While printObj.TextWidth(tmpTxt) > toWidth
    tmpTxt = Left(tmpTxt, Len(tmpTxt) - 1)
  Loop
  fitText = tmpTxt
End Function

Function getTotalPts(poolFormID As Long, matchNr As Integer)
'haal de tussenstand op voor deze deelnemer na deze wedstrijd
Dim rs As New ADODB.Recordset
Dim sqlstr As String
  sqlstr = "Select pointsGrandTotal from tblCompetitorPoints"
  sqlstr = sqlstr & " Where competitorPoolID = " & poolFormID
  sqlstr = sqlstr & " AND matchOrder = " & matchNr
  rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
  If Not rs.EOF Then
    getTotalPts = rs!pointsGrandTotal
  Else
    getTotalPts = 0
  End If
  rs.Close
  Set rs = Nothing

End Function

Function getTotalMoney(poolFormID As Long, matchNr As Integer)
'haal de tussenstand op voor deze deelnemer na deze wedstrijd
Dim rs As New ADODB.Recordset
Dim sqlstr As String
    sqlstr = "Select * from tblCompetitorPoints"
    sqlstr = sqlstr & " Where competitorPoolId = " & poolFormID
    sqlstr = sqlstr & " AND matchOrder = " & matchNr
    rs.Open sqlstr, cn, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
      getTotalMoney = rs!moneyTotal
    Else
      getTotalMoney = 0
    End If
    rs.Close
    Set rs = Nothing
End Function

