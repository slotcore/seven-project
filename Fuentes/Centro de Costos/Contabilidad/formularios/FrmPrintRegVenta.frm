VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrintRegVenta 
   Caption         =   "Contabilidad - Registro de Ventas"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VSPrinter7LibCtl.VSPrinter Vp 
      Height          =   7515
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _cx             =   20955
      _cy             =   13256
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
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      MarginLeft      =   720
      MarginTop       =   1440
      MarginRight     =   720
      MarginBottom    =   720
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
      Zoom            =   75
      ZoomMode        =   0
      ZoomMax         =   400
      ZoomMin         =   25
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   255
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
      Navigation      =   1
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
End
Attribute VB_Name = "FrmPrintRegVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SeEjecuto As Boolean

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Cargar
    End If
End Sub

Private Sub Form_Load()
    Vp.Orientation = orLandscape
    Vp.PaperSize = pprA4
    SeEjecuto = False
End Sub

Sub Cargar()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.preuni " _
        & " From alm_inventario ORDER BY alm_inventario.descripcion", xCon
    
    With Vp
        ' set up
        .FontName = "Courier New"
        .FontSize = 10
        '.ColorMode = cmColor
        .TextColor = &H80000008 'RGB(200, 200, 200)
        '.PenColor = RGB(190, 190, 190)
        
        'MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone
        '.PageBorder = pbTopBottom
        
        .StartDoc
            Cabecera
            
            Dim A, xFila As Integer
            xFila = 2300
            Rst.MoveFirst
            .FontSize = 6
            For A = 1 To FrmRegVentasOtros.Fg1.Rows - 1
                DoEvents
                .TextAlign = taLeftTop
                .CurrentX = 800:  .CurrentY = xFila: .Paragraph = FrmRegVentasOtros.Fg1.TextMatrix(A, 1) 'Nº Reg
                .CurrentX = 1400: .CurrentY = xFila: .Paragraph = FrmRegVentasOtros.Fg1.TextMatrix(A, 2) 'Fecha Documento
                .CurrentX = 2200: .CurrentY = xFila: .Paragraph = FrmRegVentasOtros.Fg1.TextMatrix(A, 3) 'Tipo de Documento
                .CurrentX = 2700: .CurrentY = xFila: .Paragraph = FrmRegVentasOtros.Fg1.TextMatrix(A, 4) 'Nº Documentos
                .CurrentX = 4000: .CurrentY = xFila: .Paragraph = FrmRegVentasOtros.Fg1.TextMatrix(A, 5) 'Nº Ruc
                .CurrentX = 5100: .CurrentY = xFila: .Paragraph = FrmRegVentasOtros.Fg1.TextMatrix(A, 6) 'Cliente
                
'                .CurrentX = 2200: .CurrentY = xFila: .Paragraph = FrmRegComVen.Fg1.TextMatrix(A, 3) 'Tipo de Documento
'                .CurrentX = 3000: .CurrentY = xFila: .Paragraph = FrmRegComVen.Fg1.TextMatrix(A, 4) '
'                .CurrentX = 3500: .CurrentY = xFila: .Paragraph = FrmRegComVen.Fg1.TextMatrix(A, 5)
'                .CurrentX = 4800: .CurrentY = xFila: .Paragraph = FrmRegComVen.Fg1.TextMatrix(A, 6)
'                .CurrentX = 5800: .CurrentY = xFila: .Paragraph = Mid(FrmRegComVen.Fg1.TextMatrix(A, 7), 1, 33)
                
                .TextAlign = taRightTop
                .TextBox FrmRegVentasOtros.Fg1.TextMatrix(A, 7), 8200, xFila, 600, "250"
                
                .TextAlign = taRightTop
                .TextBox FrmRegVentasOtros.Fg1.TextMatrix(A, 8), 9000 - 400, xFila, 1000, "250" 'ina 1
                
                .TextAlign = taRightTop
                .TextBox FrmRegVentasOtros.Fg1.TextMatrix(A, 9), 9900 - 500, xFila, 1000, "250" 'ina 2
                
                .TextAlign = taRightTop
                .TextBox FrmRegVentasOtros.Fg1.TextMatrix(A, 10), 10800 - 600, xFila, 1000, "250" 'ina 3
                
                .TextAlign = taRightTop
                .TextBox FrmRegVentasOtros.Fg1.TextMatrix(A, 11), 11700 - 700, xFila, 1000, "250"
                
                .TextAlign = taRightTop
                .TextBox FrmRegVentasOtros.Fg1.TextMatrix(A, 12), 12600 - 800, xFila, 1000, "250"
                
                .TextAlign = taRightTop
                .TextBox FrmRegVentasOtros.Fg1.TextMatrix(A, 13), 13500 - 900, xFila, 1000, "250"
                
                .TextAlign = taRightTop
                .TextBox FrmRegVentasOtros.Fg1.TextMatrix(A, 14), 14400 - 1000, xFila, 1000, "250"
                
'                .TextAlign = taRightTop
'                .TextBox FrmRegComVen.Fg1.TextMatrix(A, 16), 15300 - 1100, xFila, 1000, "250"
'
'                .TextAlign = taRightTop
'                .TextBox FrmRegComVen.Fg1.TextMatrix(A, 17), 16200 - 1200, xFila, 1000, "250"
                
                If xFila >= 10900 Then
                    Vp.DrawLine 800, 11090, 16100, 11090
                    .NewPage
                    .TextAlign = taLeftTop
                    Cabecera
                    .FontSize = 6
                    xFila = 2300
                Else
                    xFila = xFila + 200
                End If
            Next A
            Vp.DrawLine 800, 11090, 16100, 11090
        .EndDoc
        .ScrollIntoView 0, 0
    End With
End Sub

Sub Cabecera()
    Dim xMes, xMoneda As String
    xMes = Format(FrmRegVentasOtros.TxtFchIni.Valor, "mmmm")
    xMoneda = "Nuevos Soles"
    Vp.FontSize = 10
    Vp.CurrentX = 800: Vp.CurrentY = 700: Vp.Paragraph = NomEmp
    Vp.CurrentX = 13900: Vp.CurrentY = 700: Vp.Paragraph = "FECHA : " + Format(Date, "dd/mm/yyyy")
    
    Vp.CurrentX = 800: Vp.CurrentY = 950: Vp.Paragraph = "R.U.C. Nº : " + NumRUC

    Vp.CurrentX = 7000: Vp.CurrentY = 1100:  Vp.Paragraph = "REGISTRO DE VENTAS MES DE " & UCase(Trim(xMes)) & " " & AnoTra
    Vp.CurrentX = 7300: Vp.CurrentY = 1350:  Vp.Paragraph = "(Expresado en " + xMoneda + ")":
    Vp.FontSize = 6
    Vp.DrawLine 800, 1800, 16100, 1800
    Vp.CurrentX = 800:    Vp.CurrentY = 1900:  Vp.Paragraph = "Nº Reg."
    Vp.CurrentX = 1400:   Vp.CurrentY = 1900:  Vp.Paragraph = "Fch. Doc."
    Vp.CurrentX = 2200:   Vp.CurrentY = 1900:  Vp.Paragraph = "T.D."
    Vp.CurrentX = 2700:   Vp.CurrentY = 1900:  Vp.Paragraph = "Nº Documento"
    Vp.CurrentX = 4000:   Vp.CurrentY = 1900:  Vp.Paragraph = "Nº R.U.C."
    Vp.CurrentX = 5100:   Vp.CurrentY = 1900:  Vp.Paragraph = "Proveedor"
    Vp.CurrentX = 8500:   Vp.CurrentY = 1900:  Vp.Paragraph = "T.C."
    Vp.CurrentX = 9200 - 300: Vp.CurrentY = 1900:  Vp.Paragraph = "Valor Exp."
    Vp.CurrentX = 10100 - 400: Vp.CurrentY = 1900:  Vp.Paragraph = "Oper Exo."
    Vp.CurrentX = 11000 - 500: Vp.CurrentY = 1900:  Vp.Paragraph = "Base Imp."
    Vp.CurrentX = 11900 - 600: Vp.CurrentY = 1900:  Vp.Paragraph = "Imp.I.G.V."
    Vp.CurrentX = 12800 - 700: Vp.CurrentY = 1900:  Vp.Paragraph = "Imp.I.S.C."
    Vp.CurrentX = 13700 - 800: Vp.CurrentY = 1900:  Vp.Paragraph = "Otro Trib."
    Vp.CurrentX = 14600 - 900: Vp.CurrentY = 1900:  Vp.Paragraph = "Imp. Total"
'    Vp.CurrentX = 15500 - 1000: Vp.CurrentY = 1900:  Vp.Paragraph = "Imp.I.S.C."
'    Vp.CurrentX = 16400 - 1100: Vp.CurrentY = 1900:  Vp.Paragraph = "Imp. TOTAL"
    
    Vp.DrawLine 800, 2200, 16100, 2200
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Height > 500 Then Vp.Height = Me.Height - 500
    Vp.Top = 1
'    vp.Left = 10
    Vp.Width = Me.Width - 200
    Err.Clear
End Sub
