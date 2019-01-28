VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrintEquiFicha 
   Caption         =   "Contabilidad - Reporte Registro de Compras"
   ClientHeight    =   7515
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "FrmPrintEquiFicha.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7515
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VSPrinter7LibCtl.VSPrinter Vp 
      Height          =   7515
      Left            =   0
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
Attribute VB_Name = "FrmPrintEquiFicha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SeEjecuto As Boolean, vStrCons As String, vNomEmp As String, vNumRuc As String
Dim RsPrintEquipo As New ADODB.Recordset
Dim m_Form As String, m_id As Long
Property Let propId(ByVal pData As Long)
    m_id = pData
End Property
Property Let propFormulario(ByVal pData As String)
    m_Form = pData
End Property
Private Sub CabeceraPrintFicEquipo()
    With Vp
        .StartDoc
        .CurrentX = 1000: .CurrentY = 500: .TextAlign = taCenterMiddle
        .TextColor = vbBlack: .FontBold = True: .FontSize = 14
        Vp = "FICHA DE EQUIPO"
        'CLASE DEL EQUIPO
        .TextAlign = taLeftMiddle
        .FontBold = True
        .FontSize = 12
        .TextBox "CLASE DE EQUIPO: ", 1500, 1500, 3000, 400
        
        .FontBold = False
        .FontSize = 12
        .TextBox Trim(RsPrintEquipo("DESCEQCLASE")), 4100, 1500, 3000, 400
        'DESCRIPCION
        .FontBold = True
        .FontSize = 12
        .TextBox "DESCRIPCION: ", 1500, 2000, 3000, 400

        .FontBold = False
        .FontSize = 12
        .TextBox Trim(RsPrintEquipo("descripcion")), 4100, 2000, 3000, 400
        'AREA
        .FontBold = True
        .FontSize = 12
        .TextBox "AREA: ", 1500, 2500, 3000, 400

        .FontBold = False
        .FontSize = 12
        .TextBox Trim(RsPrintEquipo("AREA")), 4100, 2500, 3000, 400
        'EMPLEADO
        .FontBold = True
        .FontSize = 12
        .TextBox "EMPLEADO: ", 1500, 3000, 3000, 400

        .FontBold = False
        .FontSize = 12
        .TextBox NulosC(RsPrintEquipo("NOMEMP")), 4100, 3000, 4000, 400
        'UNIDAD DE MEDIDA
        .FontBold = True
        .FontSize = 12
        .TextBox "UNI. DE MEDIDA: ", 1500, 3500, 3000, 400

        .FontBold = False
        .FontSize = 12
        .TextBox NulosC(RsPrintEquipo("UMEDIDA")), 4100, 3500, 3000, 400
        'CAPACIDAD
        .FontBold = True
        .FontSize = 12
        .TextBox "CAPACIDAD: ", 1500, 4000, 3000, 400

        .FontBold = False
        .FontSize = 12
        .TextBox Format(NulosN(RsPrintEquipo("cap")), "0.00"), 4100, 4000, 3000, 400
        'FOTO
        .DrawLine 8120, 1450, 11430, 1450 'LINEA HORIZONTAL SUPERIOR
        .DrawLine 8120, 1450, 8120, 4750 'LINEA VERTICAL IZQ
        .DrawLine 11430, 1450, 11430, 4750 'LINEA VERTICAL DERECHA
        .DrawLine 8120, 4750, 11430, 4750 'LINEA HORIZONTAL INFERIOR
        .CalcPicture = FrmManEquipos.ImgFoto.Picture
        .X1 = 8180: .X2 = 11400
        .Y1 = 1500: .Y2 = 4700
        .Picture = FrmManEquipos!ImgFoto.Picture
        'CARACTERISTICAS
        .SpaceAfter = 130
        .FontBold = True
        .FontSize = 12
        .TextBox "CARACTERISTICAS: ", 1500, 4500, 3000, 400
               
        .FontBold = False
        .FontSize = 12
        .TextAlign = taJustMiddle
        .CurrentX = 1500: .CurrentY = .CurrentY
        .MarginLeft = 1500
        .Paragraph = NulosC(RsPrintEquipo("caracteristica"))
        'Vp.BrushColor = RGB(255, 255, 255)
        '.TextBox Trim(RsPrintEquipo("carecteristica")), 4100, 10000, 6000, 0, False, False, True
        .EndDoc
    End With
End Sub

Sub ImprimirFichaEquipo()
    Vp.Top = 0: Vp.Left = 0
    RsPrintEquipo.CursorLocation = adUseClient
    
    RST_Busq RsPrintEquipo, "SELECT man_equipo.id, man_equipoclase.descripcion AS desceqclase, man_equipo.descripcion, pla_area.descripcion AS AREA, " _
        & " pla_empleados.nom + ' ' + pla_empleados.ape AS NOMEMP, mae_unidades.descripcion AS UMEDIDA, man_equipo.cap, man_equipoclase.id AS IDQUCLASE, " _
        & " pla_area.id AS IDAREA, pla_empleados.id as IDEMP, mae_unidades.id AS IDUM, man_equipo.caracteristica, man_equipo.codigo " _
        & " FROM mae_unidades RIGHT JOIN (pla_empleados RIGHT JOIN (pla_area RIGHT JOIN (man_equipoclase RIGHT JOIN man_equipo " _
        & " ON man_equipoclase.id = man_equipo.idclaequ) ON pla_area.id = man_equipo.idarea) ON pla_empleados.id = man_equipo.idemp) " _
        & " ON mae_unidades.id = man_equipo.idunimed WHERE man_equipo.id = " & m_id & "", xCon
    
    CabeceraPrintFicEquipo
    Set RsPrintEquipo = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Select Case m_Form
            Case "FormRptCompVta"
                vNomEmp = "RINCON"
                vNumRuc = "00000000000"
                Cargar
            Case "FormFicEqu"
                ImprimirFichaEquipo
        End Select
    End If
End Sub
Private Sub Form_Load()
    Select Case m_Form
        Case "FormRptCompVta"
            Vp.Orientation = orLandscape
        Case "FormFicEqu"
            Vp.Orientation = orPortrait
    End Select
    Vp.PaperSize = pprA4
    SeEjecuto = False
End Sub
Sub Cargar()
    With Vp
        Vp.Top = 0: Vp.Left = 0
        .FontName = "Courier New"
        .FontSize = 10
        .TextColor = &H80000008 'RGB(200, 200, 200)
        'MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone
        .StartDoc
            Cabecera
            Dim A As Long, xFila As Long
            xFila = 2300
'            .FontSize = 11
            For A = 1 To FrmPrintRptCompras.Dg1.Rows - 1
                .TextAlign = taLeftTop
'                .CurrentX = 800:  .CurrentY = xFila: .Paragraph = FrmRegComVen.Fg1.TextMatrix(A, 1), '.CurrentX = 1400: .CurrentY = xFila: .Paragraph = FrmRegComVen.Fg1.TextMatrix(A, 2), '.CurrentX = 2200: .CurrentY = xFila: .Paragraph = FrmRegComVen.Fg1.TextMatrix(A, 3), '.CurrentX = 3000: .CurrentY = xFila: .Paragraph = FrmRegComVen.Fg1.TextMatrix(A, 4)
'                .CurrentX = 3500: .CurrentY = xFila: .Paragraph = FrmRegComVen.Fg1.TextMatrix(A, 5), '.CurrentX = 4800: .CurrentY = xFila: .Paragraph = FrmRegComVen.Fg1.TextMatrix(A, 6), '.CurrentX = 5800: .CurrentY = xFila: .Paragraph = Mid(FrmRegComVen.Fg1.TextMatrix(A, 7), 1, 33)
    
                .TextAlign = taLeftTop
                .FontSize = 11
                .TextBox FrmPrintRptCompras.Dg1.TextMatrix(A, 1), 800, xFila, 1000, 300     'TIPO DOC
                .TextBox FrmPrintRptCompras.Dg1.TextMatrix(A, 2), 2000, xFila, 1500, 300     'NRO DOC
                .TextBox FrmPrintRptCompras.Dg1.TextMatrix(A, 3), 4000, xFila, 3000, 300     'PROVEEDOR / CLIENTE
                .TextBox FrmPrintRptCompras.Dg1.TextMatrix(A, 4), 7500, xFila, 1000, 300    'FEC EMIS
                .TextBox FrmPrintRptCompras.Dg1.TextMatrix(A, 5), 9000, xFila, 1000, 300    'MONEDA
                .TextAlign = taRightTop
                .TextBox FrmPrintRptCompras.Dg1.TextMatrix(A, 6), 10000, xFila, 1000, 300    'TIP CAMBIO
                .TextAlign = taRightTop
                .TextBox Format(FrmPrintRptCompras.Dg1.TextMatrix(A, 7), "######0.000"), 11400, xFila, 1500, 300    'IMPORTE $
                .TextAlign = taRightTop
                .TextBox Format(FrmPrintRptCompras.Dg1.TextMatrix(A, 8), "######0.00"), 13200, xFila, 1500, 300    'IMPORTE S/.
                .TextAlign = taRightTop
                .TextBox Format(FrmPrintRptCompras.Dg1.TextMatrix(A, 9), "######0.00"), 14800, xFila, 1500, 300    'SALDO
                If xFila >= 10900 Then
                    Vp.DrawLine 800, 11090, 17000, 11090
                    .NewPage
                    .TextAlign = taLeftTop
                    Cabecera
                    .FontSize = 6
                    xFila = 2300
                Else
                    xFila = xFila + 200
                End If
            Next
            Vp.DrawLine 800, 11090, 16500, 11090
            
            Vp.FontSize = 12
            Vp.TextBox "TOTAL", 10000, 11100, 1500, 300
            Vp.TextBox FrmPrintRptCompras.TxtTotalDol.Text, 11400, 11100, 1500, 300
            Vp.TextBox FrmPrintRptCompras.TxtTotalSol.Text, 13200, 11100, 1500, 300
            Vp.TextBox FrmPrintRptCompras.TxtTotSaldo.Text, 14800, 11100, 1500, 300
            .EndDoc
            .ScrollIntoView 0, 0, 0, 0
    End With
End Sub
Sub Cabecera()
    Dim vFecha1 As String, vFecha2 As String, vProveedor As String
    vProveedor = Trim(FrmPrintRptCompras.TxtProv.Text)
    
    Vp.FontSize = 13
    Vp.TextBox vNomEmp, 800, 200, 5000, 400 'NOMBRE EMP
    
    Vp.TextBox vNumRuc, 800, 400, 3000, 400  'NUM RUC
    
    Vp.TextAlign = taCenterTop
    If UCase(Trim(FrmPrintRptCompras.CboTipoRPT.Text)) = "COMPRAS" Then
        Vp.TextBox "REPORTE DE COMPRAS", 7000, 800, 5000, 400    ''Vp.CurrentX = 7000: Vp.CurrentY = 800:  Vp.Paragraph = "REGISTRO DE COMPRAS"
        If Trim(FrmPrintRptCompras.TxtProv.Text) <> "" Then
            Vp.TextBox "PROVEEDOR : " & vProveedor, 800, 1500, 5000, 400  'Vp.CurrentX = 6600: Vp.CurrentY = 1500: Vp.Paragraph = "PROVEEDOR : " & FrmPrintRptCompras.TxtProv.Text
        End If
    ElseIf UCase(Trim(FrmPrintRptCompras.CboTipoRPT.Text)) = "VENTAS" Then
        Vp.TextBox "REPORTE DE VENTAS", 7000, 800, 5000, 400    ''Vp.CurrentX = 7000: Vp.CurrentY = 800:  Vp.Paragraph = "REGISTRO DE COMPRAS"
        If Trim(FrmPrintRptCompras.TxtProv.Text) <> "" Then
            Vp.TextBox "CLIENTE : " & vProveedor, 800, 1500, 5000, 400  'Vp.CurrentX = 6600: Vp.CurrentY = 1500: Vp.Paragraph = "PROVEEDOR : " & FrmPrintRptCompras.TxtProv.Text
        End If
    End If
'
    vFecha1 = Trim(FrmPrintRptCompras.TextBoxFecha1.Valor): vFecha2 = Trim(FrmPrintRptCompras.TextBoxFecha2.Valor)
    Vp.TextBox "DESDE: " & vFecha1 & " HASTA: " & vFecha2 & "", 7000, 1100, 5000, 400
    'Vp.CurrentX = 6500: Vp.CurrentY = 1100: Vp.Paragraph = "DESDE: " & vFecha1 & " HASTA: " & vFecha2

    Vp.FontSize = 10
    Vp.DrawLine 800, 1800, 16500, 1800
    'Vp.CurrentX = 800:    Vp.CurrentY = 1900:  Vp.Paragraph = "Nº de Docum."
    Vp.TextAlign = taJustMiddle
    Vp.CurrentX = 800:  Vp.CurrentY = 1900: Vp.Paragraph = "Tip. Doc."
    Vp.CurrentX = 2000:    Vp.CurrentY = 1900:  Vp.Paragraph = "Nº de Docum."
    If UCase(Trim(FrmPrintRptCompras.CboTipoRPT.Text)) = "COMPRAS" Then
        Vp.CurrentX = 4000:   Vp.CurrentY = 1900:  Vp.Paragraph = "Proveedor"
    ElseIf UCase(Trim(FrmPrintRptCompras.CboTipoRPT.Text)) = "VENTAS" Then
        Vp.CurrentX = 4000:   Vp.CurrentY = 1900:  Vp.Paragraph = "Cliente"
    End If
    Vp.CurrentX = 7500:   Vp.CurrentY = 1900:  Vp.Paragraph = "Fec. Emisión"
    Vp.CurrentX = 9000:   Vp.CurrentY = 1900:  Vp.Paragraph = "Moneda"
    
    Vp.TextAlign = taRightTop
    Vp.TextBox "Tip. Cambio", 10000, 1900, 1000, 250
    'Vp.CurrentX = 9000:   Vp.CurrentY = 1900:  Vp.Paragraph = "Tip. Cambio"
    
    Vp.TextAlign = taRightTop
    Vp.TextBox "Importe $", 11400, 1900, 1500, 250
    'Vp.CurrentX = 10500:   Vp.CurrentY = 1900:  Vp.Paragraph = "Importe $"
    
    Vp.TextAlign = taRightTop
    Vp.TextBox "Importe S/.", 13200, 1900, 1500, 250
    'Vp.CurrentX = 12500:   Vp.CurrentY = 1900:  Vp.Paragraph = "Importe S/."
    
    Vp.TextAlign = taRightTop
    Vp.TextBox "Saldo", 14800, 1900, 1500, 250
    'Vp.CurrentX = 14500:   Vp.CurrentY = 1900:  Vp.Paragraph = "Saldo"
    Vp.DrawLine 800, 2200, 16500, 2200
End Sub

Private Sub Form_Resize()
    Vp.Width = Me.Width
    On Error Resume Next
    Vp.Height = Me.Height - 500
End Sub
