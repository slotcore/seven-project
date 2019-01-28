VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VSPrinter7LibCtl.VSPrinter VP 
      Height          =   7080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10770
      _cx             =   18997
      _cy             =   12488
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
      Zoom            =   37.2549019607843
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstEmpleados As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim xControlForm As VSFlexGrid

Sub CargarEmpleados()
    RST_Busq RstEmpleados, "SELECT pla_empleados.*, UCase(pla_empleados!apepat)+' '+UCase(pla_empleados!apemat)+', '+pla_empleados!nom AS apenom, " _
        & " pla_categoria1.cuspp, mae_ocupacion.descripcion AS descocu FROM mae_ocupacion RIGHT JOIN (pla_empleados LEFT JOIN pla_categoria1 ON " _
        & " pla_empleados.id = pla_categoria1.idemp) ON mae_ocupacion.id = pla_categoria1.idocu", xCon

    
    
    'SELECT pla_empleados.*, UCase(pla_empleados!apepat)+' '+UCase(pla_empleados!apemat)+', '+pla_empleados!nom AS apenom, pla_categoria1.cuspp " _
        & " FROM pla_empleados LEFT JOIN pla_categoria1 ON pla_empleados.id = pla_categoria1.idemp", xCon
    'SELECT pla_empleados.*, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS apenom " _
        & " FROM pla_empleados ", xCon
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        
        CargarEmpleados
        Set xControlForm = FrmEmisionPlanilla.fg1
        ImprimirPlanilla
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
End Sub

Sub ImprimirPlanilla()
    Dim A As Integer
    With VP
        .StartDoc
        For A = 2 To xControlForm.Rows - 1
            CrearBoleta VP, 1500, A
            .NewPage
        Next A
        
        .EndDoc
    End With
End Sub

'Sub aaa()
'    With VP
'        .StartDoc
'
'            .ZoomMode = zmPageWidth
'            .FontSize = 8
'             VP = ""
'            .StartTable
'
'                '-----------------------------------------------------------
'                ' creamos la tabla 13 columnas x 30 filas
'                '-----------------------------------------------------------
'                .TableBorder = tbNone
'
'                .TableCell(tcCols) = 13
'                .TableCell(tcRows) = 32
'
'                .TableCell(tcFontName) = "Times New Roman"
'                .TableCell(tcAlign, 2, 10, 4, 10) = taCenterMiddle
'                .TableCell(tcText, 2, 10) = "BOLETA DE PAGO"
'                .TableCell(tcText, 3, 10) = "MENSUAL"
'                .TableCell(tcText, 4, 10) = "Nº 0000000000"
'
'                .BrushStyle = bsTransparent
'
'                .TableCell(tcAlign, 2, 2) = taLeftMiddle
'                .TableCell(tcText, 2, 2) = "Nº R.U.C."
'                .TableCell(tcText, 3, 2) = "Razon Social"
'                .TableCell(tcText, 4, 2) = "Direccion"
'
'                .TableCell(tcText, 6, 2) = "Periodo"
'                .TableCell(tcText, 7, 2) = "Apellidos Y Nombres"
'                .TableCell(tcText, 8, 2) = "Nº CUSSP / O.N.P."
'                .TableCell(tcText, 9, 2) = "Fch. Ingreso"
'                .TableCell(tcText, 10, 2) = "Cargo"
'
'                .TableCell(tcText, 12, 2) = "Nº Dias Trabajados"
'                .TableCell(tcText, 13, 2) = "Nº  Horas Extras 60%"
'
'
'                .TableCell(tcText, 7, 8) = "Nº D.N.I. "
'                .TableCell(tcText, 8, 8) = "Nº ESSALUD"
'                .TableCell(tcText, 9, 8) = "Fch. Cese"
'
'                .TableCell(tcText, 12, 8) = "Nº Horas Ordinarias"
'                .TableCell(tcText, 13, 8) = "Nº  Horas Extras 100%"
'
'                .TableCell(tcAlign, 15, 2, 27, 12) = taCenterMiddle
'                .TableCell(tcText, 15, 2) = "Remuneraciones"
'                .TableCell(tcText, 15, 6) = "Descuentos"
'                .TableCell(tcText, 15, 10) = "Aportaciones del Empleador"
'
'                .TableCell(tcText, 16, 2) = "Descripcion"
'                .TableCell(tcText, 16, 6) = "Descripcion"
'                .TableCell(tcText, 16, 10) = "Descripcion"
'
'                .TableCell(tcText, 16, 4) = "Importe"
'                .TableCell(tcText, 16, 8) = "Importe"
'                .TableCell(tcText, 16, 12) = "Importe"
'
'                .TableCell(tcText, 29, 2) = "Total Remuneracion"
'                .TableCell(tcText, 29, 6) = "Total Descuentos"
'                .TableCell(tcText, 21, 10) = "Total Aportaciones"
'                .TableCell(tcText, 29, 10) = "Importe Neto a Pagar"
'
'
'                .TableCell(tcAlign, 32, 2, 32, 12) = taCenterMiddle
'                .TableCell(tcText, 32, 3) = "Empleado"
'                .TableCell(tcText, 32, 11) = "Empleador"
'                '-----------------------------------------------------------
'                ' set some column widths (default width is 0.5in)
'                '-----------------------------------------------------------
'                .TableCell(tcColWidth, , 1) = "0.05in"
'                .TableCell(tcColWidth, , 2) = "0.7in"
'                .TableCell(tcColWidth, , 3) = "0.7in"
'                .TableCell(tcColWidth, , 4) = "0.7in"
'                .TableCell(tcColWidth, , 5) = "0.02in"
'                .TableCell(tcColWidth, , 6) = "0.7in"
'                .TableCell(tcColWidth, , 7) = "0.7in"
'                .TableCell(tcColWidth, , 8) = "0.7in"
'                .TableCell(tcColWidth, , 9) = "0.02in"
'                .TableCell(tcColWidth, , 10) = "0.7in"
'                .TableCell(tcColWidth, , 11) = "0.7in"
'                .TableCell(tcColWidth, , 12) = "0.7in"
'                .TableCell(tcColWidth, , 13) = "0.05in"
'
'                .TableCell(tcRowHeight, 1, 1) = "0.05in"
'                .TableCell(tcRowHeight, 2, 2) = "0.15in"
'                .TableCell(tcRowHeight, 3, 3) = "0.15in"
'                .TableCell(tcRowHeight, 4, 4) = "0.15in"
'                .TableCell(tcRowHeight, 5, 5) = "0.05in"
'                .TableCell(tcRowHeight, 6, 6) = "0.15in"
'                .TableCell(tcRowHeight, 7, 7) = "0.15in"
'                .TableCell(tcRowHeight, 8, 8) = "0.15in"
'                .TableCell(tcRowHeight, 9, 9) = "0.15in"
'                .TableCell(tcRowHeight, 10, 10) = "0.15in"
'                .TableCell(tcRowHeight, 11, 11) = "0.05in"
'                .TableCell(tcRowHeight, 12, 12) = "0.15in"
'                .TableCell(tcRowHeight, 13, 13) = "0.15in"
'                .TableCell(tcRowHeight, 14, 14) = "0.05in"
'                .TableCell(tcRowHeight, 15, 15) = "0.15in"
'                .TableCell(tcRowHeight, 16, 16) = "0.15in"
'                .TableCell(tcRowHeight, 17, 17) = "0.15in"
'                .TableCell(tcRowHeight, 18, 18) = "0.15in"
'                .TableCell(tcRowHeight, 19, 19) = "0.15in"
'                .TableCell(tcRowHeight, 20, 20) = "0.15in"
'                .TableCell(tcRowHeight, 21, 21) = "0.15in"
'                .TableCell(tcRowHeight, 22, 22) = "0.15in"
'                .TableCell(tcRowHeight, 23, 23) = "0.15in"
'                .TableCell(tcRowHeight, 24, 24) = "0.15in"
'                .TableCell(tcRowHeight, 25, 25) = "0.15in"
'                .TableCell(tcRowHeight, 26, 26) = "0.15in"
'                .TableCell(tcRowHeight, 27, 27) = "0.15in"
'                .TableCell(tcRowHeight, 28, 28) = "0.15in"
'                .TableCell(tcRowHeight, 29, 29) = "0.15in"
'                .TableCell(tcRowHeight, 30, 30) = "0.15in"
'
'                .TableCell(tcColSpan, 2, 10, 2, 12) = 3
'                .TableCell(tcColSpan, 3, 10, 3, 12) = 3
'                .TableCell(tcColSpan, 4, 10, 4, 12) = 3
'
'                .TableCell(tcColSpan, 2, 2, 2, 3) = 2
'                .TableCell(tcColSpan, 3, 2, 3, 3) = 2
'                .TableCell(tcColSpan, 4, 2, 4, 3) = 2
'
'                .TableCell(tcColSpan, 6, 2, 6, 3) = 2
'                .TableCell(tcColSpan, 7, 2, 7, 3) = 2
'                .TableCell(tcColSpan, 8, 2, 8, 3) = 2
'                .TableCell(tcColSpan, 9, 2, 9, 3) = 2
'                .TableCell(tcColSpan, 10, 2, 10, 3) = 2
'
'                .TableCell(tcColSpan, 12, 2, 12, 3) = 2
'                .TableCell(tcColSpan, 13, 2, 13, 3) = 2
'
'                .TableCell(tcColSpan, 7, 8, 7, 10) = 3
'                .TableCell(tcColSpan, 8, 8, 8, 10) = 3
'                .TableCell(tcColSpan, 9, 8, 9, 10) = 3
'                .TableCell(tcColSpan, 10, 8, 10, 10) = 3
'
'                .TableCell(tcColSpan, 12, 8, 12, 10) = 3
'                .TableCell(tcColSpan, 13, 8, 13, 10) = 3
'
'                'unimos las cabeceras de remuneraciones, descuentos, aportaciones
'                .TableCell(tcColSpan, 15, 2, 15, 4) = 3
'                .TableCell(tcColSpan, 15, 6, 15, 8) = 3
'                .TableCell(tcColSpan, 15, 10, 15, 12) = 3
'
'                .TableCell(tcColSpan, 16, 2, 16, 3) = 2
'                .TableCell(tcColSpan, 16, 6, 16, 7) = 2
'                .TableCell(tcColSpan, 16, 10, 16, 11) = 2
'
'                'unimos el pie de las cabeceras
'                .TableCell(tcColSpan, 29, 2, 29, 3) = 2
'                .TableCell(tcColSpan, 29, 6, 29, 7) = 2
'                .TableCell(tcColSpan, 29, 10, 29, 11) = 2
'                .TableCell(tcColSpan, 21, 10, 21, 11) = 2
'
'
'                .TableCell(tcBackColor, 2, 2, 4, 3) = &H8000000F
'                .TableCell(tcBackColor, 6, 2, 10, 3) = &H8000000F
'                .TableCell(tcBackColor, 12, 2, 13, 3) = &H8000000F
'
'                .TableCell(tcBackColor, 7, 8, 10, 10) = &H8000000F
'                .TableCell(tcBackColor, 12, 8, 13, 10) = &H8000000F
'
'                .TableCell(tcBackColor, 15, 2, 16, 4) = &H8000000F
'                .TableCell(tcBackColor, 15, 6, 16, 8) = &H8000000F
'                .TableCell(tcBackColor, 15, 10, 16, 12) = &H8000000F
'
'                'imprimimos lineas y formas
'                .BrushColor = &H80&
'                .DrawRectangle 7600, 1700, 10640, 2350
'
'                .DrawLine 2000, 7740, 4000, 7740
'                .DrawLine 8100, 7740, 10100, 7740
'
'            .EndTable
'
'            .CurrentY = 9000
'
'            .StartTable
'
'                '-----------------------------------------------------------
'                ' creamos la tabla 13 columnas x 30 filas
'                '-----------------------------------------------------------
'                .TableBorder = tbNone
'
'                .TableCell(tcCols) = 13
'                .TableCell(tcRows) = 32
'
'                .TableCell(tcFontName) = "Times New Roman"
'                .TableCell(tcAlign, 2, 10, 4, 10) = taCenterMiddle
'                .TableCell(tcText, 2, 10) = "BOLETA DE PAGO"
'                .TableCell(tcText, 3, 10) = "MENSUAL"
'                .TableCell(tcText, 4, 10) = "Nº 0000000000"
'
'                .BrushStyle = bsTransparent
'
'                .TableCell(tcAlign, 2, 2) = taLeftMiddle
'                .TableCell(tcText, 2, 2) = "Nº R.U.C."
'                .TableCell(tcText, 3, 2) = "Razon Social"
'                .TableCell(tcText, 4, 2) = "Direccion"
'
'                .TableCell(tcText, 6, 2) = "Periodo"
'                .TableCell(tcText, 7, 2) = "Apellidos Y Nombres"
'                .TableCell(tcText, 8, 2) = "Nº CUSSP / O.N.P."
'                .TableCell(tcText, 9, 2) = "Fch. Ingreso"
'                .TableCell(tcText, 10, 2) = "Cargo"
'
'                .TableCell(tcText, 12, 2) = "Nº Dias Trabajados"
'                .TableCell(tcText, 13, 2) = "Nº  Horas Extras 60%"
'
'
'                .TableCell(tcText, 7, 8) = "Nº D.N.I. "
'                .TableCell(tcText, 8, 8) = "Nº ESSALUD"
'                .TableCell(tcText, 9, 8) = "Fch. Cese"
'
'                .TableCell(tcText, 12, 8) = "Nº Horas Ordinarias"
'                .TableCell(tcText, 13, 8) = "Nº  Horas Extras 100%"
'
'                .TableCell(tcAlign, 15, 2, 27, 12) = taCenterMiddle
'                .TableCell(tcText, 15, 2) = "Remuneraciones"
'                .TableCell(tcText, 15, 6) = "Descuentos"
'                .TableCell(tcText, 15, 10) = "Aportaciones del Empleador"
'
'                .TableCell(tcText, 16, 2) = "Descripcion"
'                .TableCell(tcText, 16, 6) = "Descripcion"
'                .TableCell(tcText, 16, 10) = "Descripcion"
'
'                .TableCell(tcText, 16, 4) = "Importe"
'                .TableCell(tcText, 16, 8) = "Importe"
'                .TableCell(tcText, 16, 12) = "Importe"
'
'                .TableCell(tcText, 29, 2) = "Total Remuneracion"
'                .TableCell(tcText, 29, 6) = "Total Descuentos"
'                .TableCell(tcText, 21, 10) = "Total Aportaciones"
'                .TableCell(tcText, 29, 10) = "Importe Neto a Pagar"
'
'
'                .TableCell(tcAlign, 32, 2, 32, 12) = taCenterMiddle
'                .TableCell(tcText, 32, 3) = "Empleado"
'                .TableCell(tcText, 32, 11) = "Empleador"
'                '-----------------------------------------------------------
'                ' set some column widths (default width is 0.5in)
'                '-----------------------------------------------------------
'                .TableCell(tcColWidth, , 1) = "0.05in"
'                .TableCell(tcColWidth, , 2) = "0.7in"
'                .TableCell(tcColWidth, , 3) = "0.7in"
'                .TableCell(tcColWidth, , 4) = "0.7in"
'                .TableCell(tcColWidth, , 5) = "0.02in"
'                .TableCell(tcColWidth, , 6) = "0.7in"
'                .TableCell(tcColWidth, , 7) = "0.7in"
'                .TableCell(tcColWidth, , 8) = "0.7in"
'                .TableCell(tcColWidth, , 9) = "0.02in"
'                .TableCell(tcColWidth, , 10) = "0.7in"
'                .TableCell(tcColWidth, , 11) = "0.7in"
'                .TableCell(tcColWidth, , 12) = "0.7in"
'                .TableCell(tcColWidth, , 13) = "0.05in"
'
'                .TableCell(tcRowHeight, 1, 1) = "0.05in"
'                .TableCell(tcRowHeight, 2, 2) = "0.15in"
'                .TableCell(tcRowHeight, 3, 3) = "0.15in"
'                .TableCell(tcRowHeight, 4, 4) = "0.15in"
'                .TableCell(tcRowHeight, 5, 5) = "0.05in"
'                .TableCell(tcRowHeight, 6, 6) = "0.15in"
'                .TableCell(tcRowHeight, 7, 7) = "0.15in"
'                .TableCell(tcRowHeight, 8, 8) = "0.15in"
'                .TableCell(tcRowHeight, 9, 9) = "0.15in"
'                .TableCell(tcRowHeight, 10, 10) = "0.15in"
'                .TableCell(tcRowHeight, 11, 11) = "0.05in"
'                .TableCell(tcRowHeight, 12, 12) = "0.15in"
'                .TableCell(tcRowHeight, 13, 13) = "0.15in"
'                .TableCell(tcRowHeight, 14, 14) = "0.05in"
'                .TableCell(tcRowHeight, 15, 15) = "0.15in"
'                .TableCell(tcRowHeight, 16, 16) = "0.15in"
'                .TableCell(tcRowHeight, 17, 17) = "0.15in"
'                .TableCell(tcRowHeight, 18, 18) = "0.15in"
'                .TableCell(tcRowHeight, 19, 19) = "0.15in"
'                .TableCell(tcRowHeight, 20, 20) = "0.15in"
'                .TableCell(tcRowHeight, 21, 21) = "0.15in"
'                .TableCell(tcRowHeight, 22, 22) = "0.15in"
'                .TableCell(tcRowHeight, 23, 23) = "0.15in"
'                .TableCell(tcRowHeight, 24, 24) = "0.15in"
'                .TableCell(tcRowHeight, 25, 25) = "0.15in"
'                .TableCell(tcRowHeight, 26, 26) = "0.15in"
'                .TableCell(tcRowHeight, 27, 27) = "0.15in"
'                .TableCell(tcRowHeight, 28, 28) = "0.15in"
'                .TableCell(tcRowHeight, 29, 29) = "0.15in"
'                .TableCell(tcRowHeight, 30, 30) = "0.15in"
'
'                .TableCell(tcColSpan, 2, 10, 2, 12) = 3
'                .TableCell(tcColSpan, 3, 10, 3, 12) = 3
'                .TableCell(tcColSpan, 4, 10, 4, 12) = 3
'
'                .TableCell(tcColSpan, 2, 2, 2, 3) = 2
'                .TableCell(tcColSpan, 3, 2, 3, 3) = 2
'                .TableCell(tcColSpan, 4, 2, 4, 3) = 2
'
'                .TableCell(tcColSpan, 6, 2, 6, 3) = 2
'                .TableCell(tcColSpan, 7, 2, 7, 3) = 2
'                .TableCell(tcColSpan, 8, 2, 8, 3) = 2
'                .TableCell(tcColSpan, 9, 2, 9, 3) = 2
'                .TableCell(tcColSpan, 10, 2, 10, 3) = 2
'
'                .TableCell(tcColSpan, 12, 2, 12, 3) = 2
'                .TableCell(tcColSpan, 13, 2, 13, 3) = 2
'
'                .TableCell(tcColSpan, 7, 8, 7, 10) = 3
'                .TableCell(tcColSpan, 8, 8, 8, 10) = 3
'                .TableCell(tcColSpan, 9, 8, 9, 10) = 3
'                .TableCell(tcColSpan, 10, 8, 10, 10) = 3
'
'                .TableCell(tcColSpan, 12, 8, 12, 10) = 3
'                .TableCell(tcColSpan, 13, 8, 13, 10) = 3
'
'                'unimos las cabeceras de remuneraciones, descuentos, aportaciones
'                .TableCell(tcColSpan, 15, 2, 15, 4) = 3
'                .TableCell(tcColSpan, 15, 6, 15, 8) = 3
'                .TableCell(tcColSpan, 15, 10, 15, 12) = 3
'
'                .TableCell(tcColSpan, 16, 2, 16, 3) = 2
'                .TableCell(tcColSpan, 16, 6, 16, 7) = 2
'                .TableCell(tcColSpan, 16, 10, 16, 11) = 2
'
'                'unimos el pie de las cabeceras
'                .TableCell(tcColSpan, 29, 2, 29, 3) = 2
'                .TableCell(tcColSpan, 29, 6, 29, 7) = 2
'                .TableCell(tcColSpan, 29, 10, 29, 11) = 2
'                .TableCell(tcColSpan, 21, 10, 21, 11) = 2
'
'
'                .TableCell(tcBackColor, 2, 2, 4, 3) = &H8000000F
'                .TableCell(tcBackColor, 6, 2, 10, 3) = &H8000000F
'                .TableCell(tcBackColor, 12, 2, 13, 3) = &H8000000F
'
'                .TableCell(tcBackColor, 7, 8, 10, 10) = &H8000000F
'                .TableCell(tcBackColor, 12, 8, 13, 10) = &H8000000F
'
'                .TableCell(tcBackColor, 15, 2, 16, 4) = &H8000000F
'                .TableCell(tcBackColor, 15, 6, 16, 8) = &H8000000F
'                .TableCell(tcBackColor, 15, 10, 16, 12) = &H8000000F
'
'                'imprimimos lineas y formas
'                .BrushColor = &H80&
'                .DrawRectangle 7600, 1700, 10640, 2350
'
'                .DrawLine 2000, 7740, 4000, 7740
'                .DrawLine 8100, 7740, 10100, 7740
'
'            .EndTable
'
'        .EndDoc
'    End With
'End Sub

Sub CrearBoleta(xControl As VSPrinter, Posicion As Integer, PosEnElFlex As Integer)
    xControl.CurrentY = Posicion
    xControl.FontSize = 8
     VP = ""
    xControl.StartTable
    
        '-----------------------------------------------------------
        ' creamos la tabla 13 columnas x 32 filas
        '-----------------------------------------------------------
        xControl.TableBorder = tbNone
        
        xControl.TableCell(tcCols) = 13
        xControl.TableCell(tcRows) = 32
        
        xControl.TableCell(tcFontName) = "Times New Roman"
        xControl.TableCell(tcAlign, 2, 10, 4, 10) = taCenterMiddle
        xControl.TableCell(tcText, 2, 10) = "BOLETA DE PAGO"
        xControl.TableCell(tcText, 3, 10) = "MENSUAL"
        xControl.TableCell(tcText, 4, 10) = "Nº 0000000000"
        
        xControl.BrushStyle = bsTransparent
        
        xControl.TableCell(tcAlign, 2, 2) = taLeftMiddle
        xControl.TableCell(tcText, 2, 2) = "Nº R.U.C."
        xControl.TableCell(tcText, 3, 2) = "Razon Social"
        xControl.TableCell(tcText, 4, 2) = "Direccion"

        xControl.TableCell(tcText, 6, 2) = "Periodo"
        xControl.TableCell(tcText, 7, 2) = "Apellidos Y Nombres"
        xControl.TableCell(tcText, 8, 2) = "Nº CUSSP / O.N.P."
        xControl.TableCell(tcText, 9, 2) = "Fch. Ingreso"
        xControl.TableCell(tcText, 10, 2) = "Cargo"

        xControl.TableCell(tcText, 12, 2) = "Nº Dias Trabajados"
        xControl.TableCell(tcText, 13, 2) = "Nº  Horas Extras 60%"
        
        xControl.TableCell(tcText, 7, 8) = "Nº D.N.I. "
        xControl.TableCell(tcText, 8, 8) = "Nº ESSALUD"
        xControl.TableCell(tcText, 9, 8) = "Fch. Cese"

        xControl.TableCell(tcText, 12, 8) = "Nº Horas Ordinarias"
        xControl.TableCell(tcText, 13, 8) = "Nº  Horas Extras 100%"
        
        xControl.TableCell(tcAlign, 15, 2, 27, 12) = taCenterMiddle
        xControl.TableCell(tcText, 15, 2) = "Remuneraciones"
        xControl.TableCell(tcText, 15, 6) = "Descuentos"
        xControl.TableCell(tcText, 15, 10) = "Aportaciones del Empleador"
        
        xControl.TableCell(tcText, 16, 2) = "Descripcion"
        xControl.TableCell(tcText, 16, 6) = "Descripcion"
        xControl.TableCell(tcText, 16, 10) = "Descripcion"
        
        xControl.TableCell(tcText, 16, 4) = "Importe"
        xControl.TableCell(tcText, 16, 8) = "Importe"
        xControl.TableCell(tcText, 16, 12) = "Importe"
        
        xControl.TableCell(tcText, 29, 2) = "Total Remuneracion"
        xControl.TableCell(tcText, 29, 6) = "Total Descuentos"
        xControl.TableCell(tcText, 21, 10) = "Total Aportaciones"
        xControl.TableCell(tcText, 29, 10) = "Importe Neto a Pagar"
        
        
        xControl.TableCell(tcAlign, 32, 2, 32, 12) = taCenterMiddle
        xControl.TableCell(tcText, 32, 3) = "Empleado"
        xControl.TableCell(tcText, 32, 11) = "Empleador"
        '-----------------------------------------------------------
        ' set some column widths (default width is 0.5in)
        '-----------------------------------------------------------
        xControl.TableCell(tcColWidth, , 1) = "0.05in"
        xControl.TableCell(tcColWidth, , 2) = "0.7in"
        xControl.TableCell(tcColWidth, , 3) = "0.7in"
        xControl.TableCell(tcColWidth, , 4) = "0.7in"
        xControl.TableCell(tcColWidth, , 5) = "0.02in"
        xControl.TableCell(tcColWidth, , 6) = "1in"
        xControl.TableCell(tcColWidth, , 7) = "0.7in"
        xControl.TableCell(tcColWidth, , 8) = "0.7in"
        xControl.TableCell(tcColWidth, , 9) = "0.02in"
        xControl.TableCell(tcColWidth, , 10) = "0.7in"
        xControl.TableCell(tcColWidth, , 11) = "0.7in"
        xControl.TableCell(tcColWidth, , 12) = "0.7in"
        xControl.TableCell(tcColWidth, , 13) = "0.05in"

        xControl.TableCell(tcRowHeight, 1, 1) = "0.05in"
        xControl.TableCell(tcRowHeight, 2, 2) = "0.15in"
        xControl.TableCell(tcRowHeight, 3, 3) = "0.15in"
        xControl.TableCell(tcRowHeight, 4, 4) = "0.15in"
        xControl.TableCell(tcRowHeight, 5, 5) = "0.05in"
        xControl.TableCell(tcRowHeight, 6, 6) = "0.15in"
        xControl.TableCell(tcRowHeight, 7, 7) = "0.15in"
        xControl.TableCell(tcRowHeight, 8, 8) = "0.15in"
        xControl.TableCell(tcRowHeight, 9, 9) = "0.15in"
        xControl.TableCell(tcRowHeight, 10, 10) = "0.15in"
        xControl.TableCell(tcRowHeight, 11, 11) = "0.05in"
        xControl.TableCell(tcRowHeight, 12, 12) = "0.15in"
        xControl.TableCell(tcRowHeight, 13, 13) = "0.15in"
        xControl.TableCell(tcRowHeight, 14, 14) = "0.05in"
        xControl.TableCell(tcRowHeight, 15, 15) = "0.15in"
        xControl.TableCell(tcRowHeight, 16, 16) = "0.15in"
        xControl.TableCell(tcRowHeight, 17, 17) = "0.15in"
        xControl.TableCell(tcRowHeight, 18, 18) = "0.15in"
        xControl.TableCell(tcRowHeight, 19, 19) = "0.15in"
        xControl.TableCell(tcRowHeight, 20, 20) = "0.15in"
        xControl.TableCell(tcRowHeight, 21, 21) = "0.15in"
        xControl.TableCell(tcRowHeight, 22, 22) = "0.15in"
        xControl.TableCell(tcRowHeight, 23, 23) = "0.15in"
        xControl.TableCell(tcRowHeight, 24, 24) = "0.15in"
        xControl.TableCell(tcRowHeight, 25, 25) = "0.15in"
        xControl.TableCell(tcRowHeight, 26, 26) = "0.15in"
        xControl.TableCell(tcRowHeight, 27, 27) = "0.15in"
        xControl.TableCell(tcRowHeight, 28, 28) = "0.15in"
        xControl.TableCell(tcRowHeight, 29, 29) = "0.15in"
        xControl.TableCell(tcRowHeight, 30, 30) = "0.15in"

        xControl.TableCell(tcColSpan, 2, 10, 2, 12) = 3
        xControl.TableCell(tcColSpan, 3, 10, 3, 12) = 3
        xControl.TableCell(tcColSpan, 4, 10, 4, 12) = 3
        
        xControl.TableCell(tcColSpan, 2, 2, 2, 3) = 2
        xControl.TableCell(tcColSpan, 3, 2, 3, 3) = 2
        xControl.TableCell(tcColSpan, 4, 2, 4, 3) = 2

        xControl.TableCell(tcColSpan, 6, 2, 6, 3) = 2
        xControl.TableCell(tcColSpan, 7, 2, 7, 3) = 2
        xControl.TableCell(tcColSpan, 8, 2, 8, 3) = 2
        xControl.TableCell(tcColSpan, 9, 2, 9, 3) = 2
        xControl.TableCell(tcColSpan, 10, 2, 10, 3) = 2

        xControl.TableCell(tcColSpan, 12, 2, 12, 3) = 2
        xControl.TableCell(tcColSpan, 13, 2, 13, 3) = 2

        xControl.TableCell(tcColSpan, 7, 8, 7, 10) = 3
        xControl.TableCell(tcColSpan, 8, 8, 8, 10) = 3
        xControl.TableCell(tcColSpan, 9, 8, 9, 10) = 3
        xControl.TableCell(tcColSpan, 10, 8, 10, 10) = 3

        xControl.TableCell(tcColSpan, 12, 8, 12, 10) = 3
        xControl.TableCell(tcColSpan, 13, 8, 13, 10) = 3
        
        'unimos las cabeceras de remuneraciones, descuentos, aportaciones
        xControl.TableCell(tcColSpan, 15, 2, 15, 4) = 3
        xControl.TableCell(tcColSpan, 15, 6, 15, 8) = 3
        xControl.TableCell(tcColSpan, 15, 10, 15, 12) = 3
        
        xControl.TableCell(tcColSpan, 16, 2, 16, 3) = 2
        xControl.TableCell(tcColSpan, 16, 6, 16, 7) = 2
        xControl.TableCell(tcColSpan, 16, 10, 16, 11) = 2
        
        'unimos el pie de las cabeceras
        xControl.TableCell(tcColSpan, 29, 2, 29, 3) = 2
        xControl.TableCell(tcColSpan, 29, 6, 29, 7) = 2
        xControl.TableCell(tcColSpan, 29, 10, 29, 11) = 2
        xControl.TableCell(tcColSpan, 21, 10, 21, 11) = 2
        
        'unimos los campos que se van a mostrar
        xControl.TableCell(tcColSpan, 2, 4, 2, 8) = 5 'Campo Nº R.U.C.
        xControl.TableCell(tcText, 2, 4) = NulosC(NumRuc)
        xControl.TableCell(tcColSpan, 3, 4, 3, 8) = 5 ' Campo Razon socia
        xControl.TableCell(tcText, 3, 4) = NulosC(NomEmp)
        xControl.TableCell(tcColSpan, 4, 4, 4, 8) = 5 ' Campo Direccion
        xControl.TableCell(tcText, 4, 4) = NulosC(DirEmp)
        
        RstEmpleados.MoveFirst
        RstEmpleados.Find "id = " & NulosN(xControlForm.TextMatrix(PosEnElFlex, 2)) & ""
   If RstEmpleados.EOF = False And RstEmpleados.BOF = False Then
        xControl.TableCell(tcColSpan, 6, 4, 6, 8) = 5   'campo para poner el periodo de trabajo
        xControl.TableCell(tcColSpan, 7, 4, 7, 7) = 4   'campo para el nombre del trabajador
        xControl.TableCell(tcText, 7, 4) = NulosC(NulosC(RstEmpleados("apenom")))
        xControl.TableCell(tcText, 7, 11) = NulosC(RstEmpleados("numdoc"))
        
        xControl.TableCell(tcColSpan, 8, 4, 8, 6) = 3  'campo para poner el Nº de CUSSP o ONP
        'xControl.TableCell(tcColSpan, 8, 11, 8, 12) = 1  'campo para poner el Nº de ESSALUD
        xControl.TableCell(tcText, 8, 4) = NulosC(RstEmpleados("cuspp"))
        xControl.TableCell(tcText, 8, 11) = NulosC(RstEmpleados("numessalud"))
        
        xControl.TableCell(tcColSpan, 9, 4, 9, 6) = 3  'campo para poner la fecha de ingreso del trabajador
        xControl.TableCell(tcText, 9, 4) = NulosC(RstEmpleados("fching"))
        xControl.TableCell(tcText, 10, 11) = "99/99/99"
        
        
        
        
        xControl.TableCell(tcColSpan, 10, 4, 10, 7) = 4 'campo para poner la fecha de cese del trabajador
        xControl.TableCell(tcText, 10, 4) = NulosC(RstEmpleados("descocu"))
    End If


'            xControl.TableCell(tcColSpan, 12, 4, 12, 4) = 1 'campo dias trabajados
'            xControl.TableCell(tcColSpan, 13, 4, 13, 4) = 1 'campo horas extras al 60%
'
'            xControl.TableCell(tcColSpan, 7, 11, 7, 12) = 3  'campo para el dni
'            xControl.TableCell(tcColSpan, 8, 11, 8, 12) = 3  'campo para el numero de esalud
'            xControl.TableCell(tcColSpan, 9, 11, 9, 12) = 3  'campo para la fecha de cese
'            xControl.TableCell(tcColSpan, 10, 11, 10, 12) = 3 'campo horas extras al 100%
        Dim A, B, xFila, xIdConcep As Integer
        Dim Descripcion As String
        Dim Total As Double
        
        'Extraemos las aportaciones del trabajador
        
        xFila = 17
        For A = 1 To 12
            
            xControl.TableCell(tcColSpan, xFila, 2, xFila, 3) = 2
            xFila = xFila + 1
        Next A
        
        xFila = 17
        For B = 3 To xControlForm.Cols - 1
            If Mid(xControlForm.TextMatrix(1, B), 1, 1) = "I" Then
                xIdConcep = Val(Mid(xControlForm.TextMatrix(1, B), 3, 4))
                Descripcion = Busca_Codigo(xIdConcep, "id", "descripcion", "mae_concepingresosdet", "N", xCon)
                xControl.TableCell(tcText, xFila, 2) = Descripcion
                xControl.TableCell(tcText, xFila, 4) = xControlForm.TextMatrix(PosEnElFlex, B)
                Total = Total + NulosN(xControlForm.TextMatrix(PosEnElFlex, B))
                xFila = xFila + 1
            End If
        Next B
        
        xControl.TableCell(tcText, 29, 4) = Format(Total, "0.00")
        
                
        xControl.TableCell(tcBackColor, 2, 2, 4, 3) = &H8000000F
        xControl.TableCell(tcBackColor, 6, 2, 10, 3) = &H8000000F
        xControl.TableCell(tcBackColor, 12, 2, 13, 3) = &H8000000F
        
        xControl.TableCell(tcBackColor, 7, 8, 10, 10) = &H8000000F
        xControl.TableCell(tcBackColor, 12, 8, 13, 10) = &H8000000F
        
        xControl.TableCell(tcBackColor, 15, 2, 16, 4) = &H8000000F
        xControl.TableCell(tcBackColor, 15, 6, 16, 8) = &H8000000F
        xControl.TableCell(tcBackColor, 15, 10, 16, 12) = &H8000000F
        
        'imprimimos lineas y formas
        xControl.BrushColor = &H80&
        xControl.DrawRectangle 7600, 1700, 10640, 2350
        
        xControl.DrawLine 2000, 7740, 4000, 7740
        xControl.DrawLine 8100, 7740, 10100, 7740
        
    xControl.EndTable
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Height > 500 Then VP.Height = Me.Height - 500
    VP.Top = 1
    VP.Left = 10
    VP.Width = Me.Width - 200
    Err.Clear
End Sub
