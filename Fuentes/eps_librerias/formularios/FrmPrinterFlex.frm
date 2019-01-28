VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrinterFlex 
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VSPrinter7LibCtl.VSPrinter VS 
      Height          =   6480
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   9795
      _cx             =   17277
      _cy             =   11430
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
      MarginRight     =   1080
      MarginBottom    =   1080
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
      Zoom            =   33.659839715049
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
Attribute VB_Name = "FrmPrinterFlex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NomEmp As String
Public NumRUC As String
Public Titulo1 As String
Public Titulo2 As String
Dim xFila As Integer
Dim SeEjecuto As Boolean

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Dim xFila As Integer
        Dim xFilaInicial As Integer
                
        With FrmPrinterFlex.VS
            .MarginTop = 1000
            .MarginRight = 900
            .BrushColor = &H80000005
            .StartDoc
            CrearCabeceraVS
            
            
            xFilaInicial = 1600
            xFila = xFilaInicial
            .FontSize = 10
            .TextAlign = taCenterMiddle
            .TextBox Titulo1, 1000, xFila, 10000, 250, True, False, False
            
            .FontSize = 8
            xFila = xFila + 300
            .TextAlign = taLeftMiddle
            .TextBox Titulo2, 1000, xFila, 10000, 200, True, False, False
            
            xFila = xFila + 250
            ImprimirCabeceraRepo xFila
            
            xFila = xFila + 300
            
            ImprimirRepo xFila
            .EndDoc
        End With
    End If
End Sub

Sub ImprimirRepo(xFila As Integer)
    Dim A, B As Integer
    Dim xCol As Integer
    xCol = 1000
        
    FrmPrinterFlex.VS.FontSize = 7
        
    For B = 1 To xFg.Rows - 1
    
        For A = 1 To xFg.Cols - 1
        
        If xFg.ColAlignment(A) = flexAlignLeftCenter Then FrmPrinterFlex.VS.TextAlign = taLeftMiddle
        If xFg.ColAlignment(A) = flexAlignCenterCenter Then FrmPrinterFlex.VS.TextAlign = taCenterMiddle
        If xFg.ColAlignment(A) = flexAlignRightCenter Then FrmPrinterFlex.VS.TextAlign = taRightMiddle
        
        FrmPrinterFlex.VS.TextBox xFg.TextMatrix(B, A), xCol, xFila, xFg.ColWidth(A), 300, True, False, False
        xCol = xCol + xFg.ColWidth(A)
        
        Next A
        xCol = 1000
        xFila = xFila + 200
        xFila = LeerFila(xFila, FrmPrinterFlex.VS.FontSize)
    Next B
End Sub

Function LeerFila(xFila As Integer, xFontSize As Variant) As Integer
    If xFila >= 16000 Then
        FrmPrinterFlex.VS.NewPage
        CrearCabeceraVS
        
        xFila = 1550
        ImprimirCabeceraRepo xFila
        LeerFila = 1850
        
        FrmPrinterFlex.VS.FontSize = xFontSize
    Else
        LeerFila = xFila
    End If
End Function


Sub ImprimirCabeceraRepo(xFila As Integer)
    Dim A As Integer
    Dim xCol As Integer
    xCol = 1000
    
    FrmPrinterFlex.VS.FontSize = 7
    FrmPrinterFlex.VS.TextAlign = taCenterMiddle
    
    For A = 1 To xFg.Cols - 1
        FrmPrinterFlex.VS.TextBox xFg.TextMatrix(0, A), xCol, xFila, xFg.ColWidth(A), 300, True, False, True
        xCol = xCol + xFg.ColWidth(A)
    Next A
    FrmPrinterFlex.VS.TextAlign = taLeftMiddle
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    VS.Left = 0
    VS.Top = 0
    Me.Caption = "Almacen - " & Titulo1
End Sub

Private Sub Form_Resize()
    VS.Height = Me.Height - 380
    VS.Width = Me.Width - 120
End Sub

Sub CrearCabeceraVS()
    Dim xCad As String
    
    FrmPrinterFlex.VS.TextAlign = taLeftTop
    FrmPrinterFlex.VS.FontName = "Courier New"
    FrmPrinterFlex.VS.FontBold = True
    FrmPrinterFlex.VS.FontSize = 9
    
    FrmPrinterFlex.VS.CurrentX = 1000:      FrmPrinterFlex.VS.CurrentY = 1000
    FrmPrinterFlex.VS.Paragraph = "EMPRESA   : " & NomEmp
    
    FrmPrinterFlex.VS.CurrentX = 8800:      FrmPrinterFlex.VS.CurrentY = 1000
    FrmPrinterFlex.VS.Paragraph = "FECHA     : " & Format(date, "dd/mm/yy")
    
    FrmPrinterFlex.VS.CurrentX = 1000:      FrmPrinterFlex.VS.CurrentY = 1200
    FrmPrinterFlex.VS.Paragraph = "Nº R.U.C. :  " & NumRUC
    
    FrmPrinterFlex.VS.CurrentX = 8800:      FrmPrinterFlex.VS.CurrentY = 1200
    FrmPrinterFlex.VS.Paragraph = "Nº Pagina : " & "0001"
    
    FrmPrinterFlex.VS.DrawLine 1000, 1450, 11000, 1450
End Sub


