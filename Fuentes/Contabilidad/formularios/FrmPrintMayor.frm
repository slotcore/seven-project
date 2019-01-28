VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrintMayor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Reporte Libro Mayor"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "FrmPrintMayor.frx":0000
   KeyPreview      =   -1  'True
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
      Zoom            =   39.8040961709706
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
Attribute VB_Name = "FrmPrintMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim ARR_ACUMULA(2) As Double    '--ARRAY QUE ACUMULARA EL DEBE Y HABER POR PAGINA
                                '--ARR_ACUMULA(0) ::DEBE
                                '--ARR_ACUMULA(1) ::HABER
                                '--ARR_ACUMULA(2) ::SALDO

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Cargar
    End If
End Sub

Private Sub Form_Load()
    Vp.PaperSize = pprA4
    SeEjecuto = False
End Sub

Sub Cargar()
 
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
'''            Rst.MoveFirst
            .FontSize = 6
            For A = 1 To FrmConsultaMayor3.Fg1.Rows - 1

                .TextAlign = taLeftTop
                .CurrentX = 900:  .CurrentY = xFila: .Paragraph = FrmConsultaMayor3.Fg1.TextMatrix(A, 2) '--NUM REG
                If InStr(FrmConsultaMayor3.Fg1.TextMatrix(A, 2), "Cta Nº  :") <> 0 Then
                    .CurrentX = 900:  .CurrentY = xFila: .Paragraph = FrmConsultaMayor3.Fg1.TextMatrix(A, 2)
                    .TextAlign = taRightTop
                    .TextBox FrmConsultaMayor3.Fg1.TextMatrix(A, 10), 9000, xFila, 1000, "250" '--SALDO ANTERIOR
                    GoTo SIG_FIL
                End If
                .TextAlign = taLeftTop
                .CurrentX = 1600: .CurrentY = xFila: .Paragraph = FrmConsultaMayor3.Fg1.TextMatrix(A, 3) '--LIBRO
                .CurrentX = 3300: .CurrentY = xFila: .Paragraph = FrmConsultaMayor3.Fg1.TextMatrix(A, 4) '--TIPO DOC
                .CurrentX = 3600: .CurrentY = xFila: .Paragraph = FrmConsultaMayor3.Fg1.TextMatrix(A, 5) '--FECHA DOC
                .CurrentX = 4400: .CurrentY = xFila: .Paragraph = FrmConsultaMayor3.Fg1.TextMatrix(A, 6) '--NUM DOC
                .CurrentX = 5600: .CurrentY = xFila: .Paragraph = FrmConsultaMayor3.Fg1.TextMatrix(A, 7) '--TIPO CAMBIO
                
                .TextAlign = taRightTop
                .TextBox FrmConsultaMayor3.Fg1.TextMatrix(A, 8), 6600, xFila, 1000, "250" '--DEBE
                .TextBox FrmConsultaMayor3.Fg1.TextMatrix(A, 9), 7800, xFila, 1000, "250" '--HABER
                .TextBox FrmConsultaMayor3.Fg1.TextMatrix(A, 10), 9000, xFila, 1000, "250" '--SALDO
                
                If Trim(FrmConsultaMayor3.Fg1.TextMatrix(A, 2)) <> "" Then
                    ARR_ACUMULA(0) = ARR_ACUMULA(0) + NulosN(FrmConsultaMayor3.Fg1.TextMatrix(A, 8)) '--DEBE
                    ARR_ACUMULA(1) = ARR_ACUMULA(1) + NulosN(FrmConsultaMayor3.Fg1.TextMatrix(A, 9)) '--HABER
                    ARR_ACUMULA(2) = ARR_ACUMULA(2) + NulosN(FrmConsultaMayor3.Fg1.TextMatrix(A, 10)) '--SALDO
                End If
                
SIG_FIL:
                If xFila >= 15500 Then
                    .DrawLine 900, 15800, 11000, 15800
                    .TextAlign = taLeftTop
                    '--PONER DATOS AL FINAL DE HOJA
                    .TextAlign = taLeftTop
                    .TextBox "VAN ==>", 5600, 15850, 800, "250"
                    .TextAlign = taRightTop
                    .TextBox CStr(Format(ARR_ACUMULA(0), FORMAT_MONTO)), 6600, 15850, 1000, "250" '--DEBE
                    .TextBox CStr(Format(ARR_ACUMULA(1), FORMAT_MONTO)), 7800, 15850, 1000, "250" '--HABER
                    .NewPage
                    Cabecera
                    '--PONER DATOS AL INICIO DE HOJA
                    .TextAlign = taLeftTop
                    .TextBox "VIENEN ==>", 5600, 2200, 800, "250"
                    .TextAlign = taRightTop
                    .TextBox CStr(Format(ARR_ACUMULA(0), FORMAT_MONTO)), 6600, 2200, 1000, "250"
                    .TextBox CStr(Format(ARR_ACUMULA(1), FORMAT_MONTO)), 7800, 2200, 1000, "250"
                    
                    .FontSize = 6
                    xFila = 2400
                Else
                    xFila = xFila + 200
                End If
            Next A
            Vp.DrawLine 900, 15700, 11000, 15700
        .EndDoc
        .ScrollIntoView 0, 0
    End With
End Sub

Sub Cabecera()
    Dim xMoneda As String
    Dim nPeriodo As String

'    If FrmConsultaMayor3.OptSoles.Value = True Then
'        xMoneda = "Nuevos Soles"
'    Else
'        xMoneda = "Dolares Americanos"
'    End If
    
    If FrmConsultaMayor3.opt_fecha(0).Value = True Then
        If CDate(FrmConsultaMayor3.TxtFchIni.Valor) < CDate(FrmConsultaMayor3.TxtFchFin.Valor) Then
            nPeriodo = " Del: " + FrmConsultaMayor3.TxtFchIni.Valor + " Al: " + FrmConsultaMayor3.TxtFchFin.Valor
        Else
            nPeriodo = "Al: " + FrmConsultaMayor3.TxtFchIni.Valor
        End If
    Else
        If FrmConsultaMayor3.lbl_periodo(0).Caption = FrmConsultaMayor3.lbl_periodo(1).Caption Then
            nPeriodo = "Periodo : " + FrmConsultaMayor3.lbl_periodo(0).Caption
        Else
            nPeriodo = "Periodo : De " + FrmConsultaMayor3.lbl_periodo(0).Caption & " A " & FrmConsultaMayor3.lbl_periodo(1).Caption
        End If
        
    End If

    Vp.TextAlign = taLeftTop
    Vp.FontSize = 6
    Vp.CurrentX = 900: Vp.CurrentY = 600: Vp.Paragraph = NomEmp
    Vp.CurrentX = 9400: Vp.CurrentY = 600: Vp.Paragraph = "Fecha: " + Format(Date, "dd/mm/yy")
    Vp.CurrentX = 9400: Vp.CurrentY = 750: Vp.Paragraph = "Hora: " + Format(Now(), "hh:mm:ss AM/PM")
    Vp.CurrentX = 9400: Vp.CurrentY = 900: Vp.Paragraph = "Pág.: " + CStr(Vp.PageCount)
    
    Vp.CurrentX = 900: Vp.CurrentY = 750: Vp.Paragraph = "R.U.C. Nº : " + NumRUC
    
    Vp.TextAlign = taCenterMiddle
    Vp.FontSize = 10
    Vp.TextBox "LIBRO MAYOR", 900, 900, 10000, "250"
    Vp.FontSize = 7
    Vp.TextBox nPeriodo, 900, 1100, 10000, "250"
    Vp.TextBox "(Expresado en " + xMoneda + ")", 900, 1250, 10000, "250"

    Vp.FontSize = 6
    Vp.DrawLine 900, 1800, 11000, 1800
    Vp.TextAlign = taLeftTop
    Vp.CurrentX = 900:   Vp.CurrentY = 1900:  Vp.Paragraph = "Num.Reg."
    Vp.CurrentX = 1600:   Vp.CurrentY = 1900:  Vp.Paragraph = "Libro"
    Vp.CurrentX = 3300:   Vp.CurrentY = 1900:  Vp.Paragraph = "T.D."
    Vp.CurrentX = 3600:   Vp.CurrentY = 1900:  Vp.Paragraph = "Fch.Doc."
    Vp.CurrentX = 4400:   Vp.CurrentY = 1900:  Vp.Paragraph = "Nº Documento"
    Vp.CurrentX = 5600:   Vp.CurrentY = 1900:  Vp.Paragraph = "T.C."
    Vp.CurrentX = 6600:   Vp.CurrentY = 1900:  Vp.Paragraph = "  - D E B E -"
    Vp.CurrentX = 7800:   Vp.CurrentY = 1900:  Vp.Paragraph = "  -H A B E R-"
    Vp.CurrentX = 9000:   Vp.CurrentY = 1900:  Vp.Paragraph = "  -S A L D O-"
    
    Vp.DrawLine 900, 2150, 11000, 2150
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Height > 500 Then Vp.Height = Me.Height - 500
    Vp.Top = 1
    Vp.Left = 10
    Vp.Width = Me.Width - 200
    Err.Clear
End Sub
