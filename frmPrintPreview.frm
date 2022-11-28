VERSION 5.00
Begin VB.Form frmPrintPreview 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Afdruk voorbeeld"
   ClientHeight    =   10215
   ClientLeft      =   10710
   ClientTop       =   1635
   ClientWidth     =   9225
   FillColor       =   &H000000FF&
   HelpContextID   =   460
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   9225
   Begin VB.PictureBox vscrlHolder 
      Align           =   4  'Align Right
      Height          =   9405
      Left            =   8940
      Negotiate       =   -1  'True
      ScaleHeight     =   9345
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   450
      Width           =   285
      Begin VB.VScrollBar VScroll1 
         Height          =   10005
         LargeChange     =   5000
         Left            =   0
         SmallChange     =   1000
         TabIndex        =   6
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.PictureBox hscrlHolder 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   300
      ScaleWidth      =   9165
      TabIndex        =   3
      Top             =   9855
      Width           =   9225
      Begin VB.HScrollBar HScroll1 
         Height          =   285
         LargeChange     =   5000
         Left            =   0
         SmallChange     =   1000
         TabIndex        =   4
         Top             =   0
         Width           =   7200
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   390
      ScaleWidth      =   9165
      TabIndex        =   0
      Top             =   0
      Width           =   9225
      Begin VB.ComboBox cmbZoom 
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Text            =   "100%"
         Top             =   45
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton btnPrint 
         Caption         =   "Afdrukken"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7275
         TabIndex        =   7
         Top             =   30
         Width           =   1080
      End
      Begin VB.CommandButton btnNext 
         Caption         =   "Volgende>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6150
         TabIndex        =   2
         Top             =   30
         Width           =   1080
      End
      Begin VB.CommandButton btnPrev 
         Caption         =   "< Vorige"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5025
         TabIndex        =   1
         Top             =   30
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom"
         Height          =   285
         Left            =   225
         TabIndex        =   11
         Top             =   60
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.PictureBox pageHolder 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   11595
      Left            =   120
      Negotiate       =   -1  'True
      ScaleHeight     =   11535
      ScaleWidth      =   8085
      TabIndex        =   8
      Top             =   480
      Width           =   8145
      Begin VB.PictureBox pageContent 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F9F9F9&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   9735
         Left            =   840
         ScaleHeight     =   9735
         ScaleMode       =   0  'User
         ScaleWidth      =   5851.5
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   6345
      End
      Begin VB.PictureBox printPages 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
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
         Height          =   10695
         Index           =   0
         Left            =   120
         ScaleHeight     =   10106.83
         ScaleLeft       =   245
         ScaleMode       =   0  'User
         ScaleTop        =   245
         ScaleWidth      =   7500
         TabIndex        =   9
         Top             =   120
         Width           =   7500
      End
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim currentPage As Integer
Dim printRatio As Double
Dim zoomFactor As Double 'zoom factor

Private Sub btnNext_Click()
    zoomFactor = val(Me.cmbZoom) / 100
    If currentPage < Me.printPages.UBound Then
        currentPage = currentPage + 1
        Me.pageContent.Cls
        With Me.printPages(currentPage)
            .Visible = True
            Me.pageContent.Move .Left * zoomFactor, .Top * zoomFactor, .width * zoomFactor, .Height * zoomFactor
            Me.pageContent.PaintPicture .Image, 0, 0, .width * zoomFactor, .Height * zoomFactor
            Me.pageContent.Refresh
        End With
        Me.printPages(currentPage).ZOrder
    End If
    Me.btnPrev.Enabled = currentPage > 0
    Me.btnNext.Enabled = currentPage < Me.printPages.UBound
End Sub

Private Sub btnPrint_Click()
    With frmPrintDialog
        .btnPrint_Click 0
    End With
End Sub

Private Sub btnPrev_Click()
    zoomFactor = val(Me.cmbZoom) / 100
    If currentPage > 0 Then
        currentPage = currentPage - 1
        Me.pageContent.Cls
        With Me.printPages(currentPage)
            .Visible = True
            Me.pageContent.Move .Left * zoomFactor, .Top * zoomFactor, .width * zoomFactor, .Height * zoomFactor
            Me.pageContent.PaintPicture .Image, 0, 0, .width * zoomFactor, .Height * zoomFactor
            Me.pageContent.Refresh
        End With
        Me.printPages(currentPage).ZOrder
    End If
    Me.btnPrev.Enabled = currentPage > 0
    Me.btnNext.Enabled = currentPage < Me.printPages.UBound
End Sub

Private Function ScalePicPreviewToPrinter(picPreview As PictureBox) As Double
    
    Dim Ratio As Double ' Ratio between Printer and Picture
    Dim LRGap As Double, TBGap As Double
    Dim HeightRatio As Double, WidthRatio As Double
    Dim PgWidth As Double, PgHeight As Double
    Dim smtemp As Long
    Dim factor As Double
    smtemp = Printer.ScaleMode
    ' Get the physical page size in twips ('Inches):
    PgWidth = Printer.width '/ 1440
    PgHeight = Printer.Height ' / 1440
    'A4
    factor = PgHeight / Printer.Height
    ' Find the size of the non-printable area on the printer to
    ' use to offset coordinates. These formulas assume the
    ' printable area is centered on the page:
    
    LRGap = (Printer.width - Printer.ScaleWidth) / 2 * factor
    TBGap = (Printer.Height - Printer.ScaleHeight) / 2 * factor
    'Me.printPages(0).Container.BackColor = vbBlue
    Printer.ScaleMode = smtemp
    
    
    ' Compare the height and with ratios to determine the
    ' Ratio to use and how to size the picture box:
    HeightRatio = picPreview.Container.ScaleHeight / PgHeight
    WidthRatio = picPreview.Container.ScaleWidth / PgWidth
    
    If HeightRatio < WidthRatio Then
        Ratio = HeightRatio
    Else
        Ratio = WidthRatio
    End If
    Ratio = 1
    'Ratio = picPreview.FontSize / 8
    picPreview.Container.Height = PgHeight * Ratio
    picPreview.Container.width = PgWidth * Ratio
    Me.printPages(0).Top = TBGap * Ratio
    Me.printPages(0).Left = LRGap * Ratio
    Me.printPages(0).Height = Me.printPages(0).Container.Height - 2 * TBGap * Ratio
    Me.printPages(0).width = Me.printPages(0).Container.width - 2 * LRGap * Ratio
    ' Set default properties of picture box to match printer
    ' There are many that you could add here:
    picPreview.Container.Scale (0, 0)-(PgWidth, PgHeight)
    picPreview.Font.Name = Printer.Font.Name
    picPreview.FontSize = Printer.FontSize * Ratio
    picPreview.ForeColor = Printer.ForeColor
    picPreview.FillStyle = vbTransparent
    picPreview.Cls
    
    ScalePicPreviewToPrinter = Int(Ratio * 100) / 100
'    picPreview.ScaleMode = 1
End Function


Private Sub cmbZoom_Click()
  zoomFactor = val(Me.cmbZoom) / 100 '* 100
  Me.pageHolder.AutoRedraw = True
  Me.pageHolder.Move Me.pageHolder.Left, Me.pageHolder.Top, Printer.width * zoomFactor, Printer.Height * zoomFactor
  Me.pageContent.Cls
  DoEvents
  With Me.printPages(currentPage)
      Me.pageContent.Move .Left * zoomFactor, .Top * zoomFactor, Printer.ScaleWidth * zoomFactor, Printer.ScaleHeight * zoomFactor
      Me.pageContent.PaintPicture .Image, 0, 0, .width * zoomFactor, .Height * zoomFactor
      Me.pageContent.Refresh
  End With
  setScrollBars
End Sub

Private Sub Form_Load()
    Dim prtWidth As Integer
    Dim prtHeight As Integer
    Dim scm As Double
    Dim i As Integer
    
    Me.Font.Size = Printer.FontSize
    Me.Font.Name = Printer.Font.Name
    
    printRatio = ScalePicPreviewToPrinter(Me.printPages(0))
    
    Me.btnPrev.Enabled = False
    currentPage = 0
    For i = 25 To 200 Step 25
        Me.cmbZoom.AddItem i & "%"
    Next
    Me.printPages(0).ScaleMode = vbTwips
    With Me.printPages(0)
        Me.pageContent.Move .Left, .Top, .width, .Height
        Me.pageContent.PaintPicture .Image, 0, 0, .width, .Height
    End With
    Me.width = Me.pageHolder.width + 1000
    Me.Height = Screen.Height - 1500
    cmbZoom_Click
    centerForm Me
    Me.Visible = True
    
    
End Sub

Sub setScrollBars()
    Me.VScroll1.Height = Me.vscrlHolder.Height
    Me.HScroll1.width = Me.hscrlHolder.width - Me.vscrlHolder.width
    If Me.pageHolder.Height > Me.ScaleHeight Then
      Me.VScroll1.Max = Me.pageHolder.Height - Me.ScaleHeight + 1500
    Else
      Me.VScroll1.Max = 0
    End If
    If Me.ScaleWidth > Me.pageHolder.width Then
      Me.HScroll1.Max = 0
    Else
      Me.HScroll1.Max = Me.pageHolder.width - Me.ScaleWidth
    End If
End Sub

Private Sub Form_Resize()
Dim i As Integer
  Me.pageHolder.Left = 100
  Me.pageHolder.Top = 100 + Me.picButtons.Height
  For i = 0 To Me.printPages.UBound - 1
  Me.printPages(i).Left = Me.HScroll1 * -1 - Me.printPages(0).ScaleLeft
  Me.printPages(i).Top = Me.VScroll1 * -1 + 450
  Me.printPages(i).Top = 240
  Me.printPages(i).Left = 240
  Me.printPages(i).ScaleTop = Printer.ScaleTop - 240 '(Printer.Height - Printer.ScaleHeight) / -2
  Me.printPages(i).ScaleLeft = Printer.ScaleLeft - 240 '(Printer.Width - Printer.ScaleWidth) / -2
  Next
  Me.btnNext.Enabled = Me.printPages.UBound > 0
  setScrollBars
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
  For i = Me.printPages.Count - 1 To 1 Step -1
    Unload Me.printPages(i)
    'Set Me.printPages(i) = Nothing
  Next
  frmPrintDialog.Visible = True
End Sub

Private Sub HScroll1_Change()
    Me.pageHolder.Left = Me.HScroll1 * -1 + 450
End Sub

Private Sub VScroll1_Change()
    Me.pageHolder.Top = Me.VScroll1 * -1 + 450
End Sub


