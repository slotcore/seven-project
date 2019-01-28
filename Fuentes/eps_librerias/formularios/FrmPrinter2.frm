VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrinter2 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VSPrinter7LibCtl.VSPrinter VSPrinter1 
      Height          =   7125
      Left            =   375
      TabIndex        =   0
      Top             =   210
      Width           =   9570
      _cx             =   16880
      _cy             =   12568
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   37.4888691006233
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
End
Attribute VB_Name = "FrmPrinter2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim xNumPagina As Integer
Dim AnchoLinea As Integer
Dim AltoHoja As Integer

Private xNum As Long

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        Laod xVS
        Cargar
    End If
End Sub

Private Sub Form_Load()
    ''od xVS
    VSPrinter1 = xVS1
    xNum = 1
    Set xNum = Controls.Add("VSPrinter7LibCtl.VSPrinter", "xVS")
    
    VSPrinter7LibCtl.VSPrinter

'    Vp.PaperSize = pprA4
'    If Prin_OrientacionHoja = 2 Then
'        '0 = horinzontal
'        Vp.Orientation = orLandscape
'        AnchoLinea = 16000
'        AltoHoja = 11000
'    Else
'        '1 = Vertical
'        Vp.Orientation = orPortrait
'        AnchoLinea = 11000
'        AltoHoja = 16000
'    End If
'    SeEjecuto = False
'    xNumPagina = 0
End Sub


Private Sub Form_Resize()
    Vp.Height = Me.Height - 200
    Vp.Width = Me.Width - 200
End Sub

Private Sub VSPrinter1_Click()

End Sub
