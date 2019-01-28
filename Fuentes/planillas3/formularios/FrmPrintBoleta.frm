VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrintBoleta 
   Caption         =   "Planillas - Impresión de Boletas"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Zoom            =   37.2217275155833
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
Attribute VB_Name = "FrmPrintBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstEmpleados As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim xControlForm As VSFlexGrid

'--------
Dim RstIng As New ADODB.Recordset '--ingreso
Dim RstDesc As New ADODB.Recordset '--descuento
Dim RstApo As New ADODB.Recordset '--aporte

Public Sub pRecibeRsts(RstIngreso As ADODB.Recordset, _
                        RstDescuento As ADODB.Recordset, _
                        RstAporte As ADODB.Recordset, _
                        RstEmp As ADODB.Recordset)

    Dim mIdEmp As Variant
    
    If RstEmp.State = 0 Then Exit Sub
    If RstIngreso.State = 0 Or RstDescuento.State = 0 Or RstAporte.State = 0 Then Exit Sub
    If RstEmp.RecordCount = 0 Then Exit Sub
    Set RstIng = RstIngreso
    Set RstDesc = RstDescuento
    Set RstApo = RstAporte
    Set RstEmpleados = RstEmp
    
    With VP
        .StartDoc
        RstEmpleados.MoveFirst
        Do While Not RstEmpleados.EOF
            mIdEmp = RstEmpleados("idemp")
            RstIng.Filter = "mIdEmp=" & mIdEmp
            RstDesc.Filter = "mIdEmp=" & mIdEmp
            RstApo.Filter = "mIdEmp=" & mIdEmp
            '---------
            If RstIng.RecordCount <> 0 Then RstIng.MoveFirst
            If RstDesc.RecordCount <> 0 Then RstDesc.MoveFirst
            If RstApo.RecordCount <> 0 Then RstApo.MoveFirst
            
            If RstIng.RecordCount <> 0 Or RstDesc.RecordCount <> 0 Or RstApo.RecordCount <> 0 Then

                If RstEmpleados.Bookmark <> 1 Then .NewPage
                
                CrearBoleta VP, 1400
                
                '---segunda copia
                If RstIng.RecordCount <> 0 Then RstIng.MoveFirst
                If RstDesc.RecordCount <> 0 Then RstDesc.MoveFirst
                If RstApo.RecordCount <> 0 Then RstApo.MoveFirst
                CrearBoleta VP, 9000
                '''''''''''''''''''''''''''
            End If
            RstEmpleados.MoveNext
        Loop
        .EndDoc
    End With
  

End Sub

Private Sub CargarEmpleados(mIdEmp As Variant)
    Set RstEmpleados = Nothing
    Dim nSQL As String
    nSQL = "SELECT pla_empleados.*, UCase(pla_empleados!apepat)+' '+UCase(pla_empleados!apemat)+', '+pla_empleados!nom AS apenom, " _
        & " pla_categoria1.cuspp, mae_ocupacion.descripcion AS descocu FROM mae_ocupacion RIGHT JOIN (pla_empleados LEFT JOIN pla_categoria1 ON " _
        & " pla_empleados.id = pla_categoria1.idemp) ON mae_ocupacion.id = pla_categoria1.idocu WHERE pla_empleados.id= " & mIdEmp & "; "
        
    RST_Busq RstEmpleados, nSQL, xCon

    
    
    'SELECT pla_empleados.*, UCase(pla_empleados!apepat)+' '+UCase(pla_empleados!apemat)+', '+pla_empleados!nom AS apenom, pla_categoria1.cuspp " _
        & " FROM pla_empleados LEFT JOIN pla_categoria1 ON pla_empleados.id = pla_categoria1.idemp", xCon
    'SELECT pla_empleados.*, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS apenom " _
        & " FROM pla_empleados ", xCon
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        
'        CargarEmpleados
'        Set xControlForm = FrmEmisionPlanilla.Fg1
'        ImprimirPlanilla
    End If
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
End Sub

Private Sub CrearBoleta(xControl As VSPrinter, Posicion As Integer)
    '===================================================================================================
    'Creado : //09 Por: Enrique Pollongo
    'Propósito: Mostrar el reporte de la boleta en pantalla
    '
    'Entradas:  xControl=Componente del reporte
    '           Posiocion=Indica si es horizontal o vertical
    '
    'Resultados: Reporte de la boleta de pago
    '
    'Modificado: 05/01/11 Por: Johan Castro
    '           Mostrar el titulo del reporte segun el documento ingresado
    '           Si no hay horas de trabajo asignadas al personal, estas se mostraran vacias en el reporte
    '===================================================================================================


    Dim A&
    xControl.CurrentY = Posicion
    xControl.FontSize = 8
     VP = ""
    xControl.StartTable
    xControl.Top = 0

    '-----------------------------------------------------------
    ' creamos la tabla 13 columnas x 32 filas
    '-----------------------------------------------------------
    xControl.TableBorder = tbNone
    
    xControl.TableCell(tcCols) = 13
    xControl.TableCell(tcRows) = 32
    
    xControl.TableCell(tcFontName) = "Times New Roman"
    xControl.TableCell(tcAlign, 2, 10, 4, 10) = taCenterMiddle
    
    
    'xControl.TableCell(tcText, 2, 10) = "BOLETA DE PAGO"
    xControl.TableCell(tcText, 2, 10) = NulosC(RstEmpleados("tipodoc"))
    
    
    xControl.TableCell(tcText, 3, 10) = "MENSUAL"
    
    xControl.BrushStyle = bsTransparent
    
    xControl.TableCell(tcAlign, 2, 2) = taLeftMiddle
    xControl.TableCell(tcText, 2, 2) = "Nº R.U.C."
    xControl.TableCell(tcText, 3, 2) = "Razón Social"
    xControl.TableCell(tcText, 4, 2) = "Dirección"

    xControl.TableCell(tcText, 6, 2) = "Periodo"
    xControl.TableCell(tcText, 7, 2) = "Apellidos Y Nombres"
    xControl.TableCell(tcText, 8, 2) = "Nº CUSSP"
    xControl.TableCell(tcText, 9, 2) = "Fch.Ingreso"
    xControl.TableCell(tcText, 10, 2) = "Cargo"

    xControl.TableCell(tcText, 12, 2) = "Nº Dias Trabajados"
    xControl.TableCell(tcText, 13, 2) = "Nº Horas Extras 25%"
    
    xControl.TableCell(tcText, 7, 8) = "Nº D.N.I."
    xControl.TableCell(tcText, 8, 8) = "Nº Autogenerado"
    xControl.TableCell(tcText, 9, 8) = "Fch. Cese"
'''''    xControl.TableCell(tcText, 10, 8) = "Fondo Pensiones"

    xControl.TableCell(tcText, 12, 8) = "Nº Horas Ordinarias"
    xControl.TableCell(tcText, 13, 8) = "Nº Horas Extras 35%"
    
    xControl.TableCell(tcAlign, 15, 2, 27, 12) = taCenterMiddle
    xControl.TableCell(tcText, 15, 2) = "Remuneraciones"
    xControl.TableCell(tcText, 15, 6) = "Descuentos y Aportes del Trabajador"
    xControl.TableCell(tcText, 15, 10) = "Aportaciones del Empleador"
    
    xControl.TableCell(tcAlign, 16, 2, 16, 2) = taLeftMiddle:      xControl.TableCell(tcText, 16, 2) = "Descripción"
    xControl.TableCell(tcAlign, 16, 6, 16, 6) = taLeftMiddle:      xControl.TableCell(tcText, 16, 6) = "Descripción"
    xControl.TableCell(tcAlign, 16, 10, 16, 10) = taLeftMiddle:    xControl.TableCell(tcText, 16, 10) = "Descripción"
    
    xControl.TableCell(tcAlign, 16, 4, 16, 4) = taRightMiddle:     xControl.TableCell(tcText, 16, 4) = "Importe"
    xControl.TableCell(tcAlign, 16, 8, 16, 8) = taRightMiddle:     xControl.TableCell(tcText, 16, 8) = "Importe"
    xControl.TableCell(tcAlign, 16, 12, 16, 12) = taRightMiddle:   xControl.TableCell(tcText, 16, 12) = "Importe"
    
    xControl.TableCell(tcAlign, 26, 2, 26, 12) = taLeftMiddle
    xControl.TableCell(tcText, 26, 2) = "Total Remuneración"
    xControl.TableCell(tcText, 26, 6) = "Total Descuentos"
    xControl.TableCell(tcText, 26, 10) = "Total Aportaciones"
           
    xControl.TableCell(tcColSpan, 32, 3, 32, 4) = 2
    xControl.TableCell(tcText, 32, 3) = "Empleado"
    xControl.TableCell(tcAlign, 32, 3, 32, 3) = taCenterMiddle
    
    xControl.TableCell(tcColSpan, 32, 10, 32, 11) = 2
    xControl.TableCell(tcText, 32, 10) = "Empleador"
    xControl.TableCell(tcAlign, 32, 10, 32, 10) = taCenterMiddle
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
    xControl.TableCell(tcRowHeight, 27, 27) = "0.04in"
    xControl.TableCell(tcRowHeight, 28, 28) = "0.15in"
    xControl.TableCell(tcRowHeight, 29, 29) = "0.15in"
    xControl.TableCell(tcRowHeight, 30, 30) = "0.15in"
    
    '--uniendo celdas
    
    
    '--del cuadro de la boleta
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
    
    
    'unimos los campos que se van a mostrar
    xControl.TableCell(tcColSpan, 2, 4, 2, 8) = 5 'Campo Nº R.U.C.
    xControl.TableCell(tcText, 2, 4) = NulosC(NumRuc)
    xControl.TableCell(tcColSpan, 3, 4, 3, 8) = 5 ' Campo Razon socia
    xControl.TableCell(tcText, 3, 4) = NulosC(NomEmp)
    xControl.TableCell(tcColSpan, 4, 4, 4, 8) = 5 ' Campo Direccion
    xControl.TableCell(tcText, 4, 4) = NulosC(DirEmp)
    
   If RstEmpleados.EOF = False And RstEmpleados.BOF = False Then
        
        xControl.TableCell(tcColSpan, 6, 4, 6, 8) = 5   'campo para poner el periodo de trabajo
        xControl.TableCell(tcText, 6, 4) = NulosC(RstEmpleados("periodo"))
        
        xControl.TableCell(tcColSpan, 7, 4, 7, 7) = 4   'campo para el nombre del trabajador
        xControl.TableCell(tcText, 7, 4) = NulosC(RstEmpleados("apenom"))

        xControl.TableCell(tcColSpan, 7, 11, 7, 11) = 2
        xControl.TableCell(tcText, 7, 11) = NulosC(RstEmpleados("numdoc"))

        xControl.TableCell(tcColSpan, 8, 4, 8, 7) = 4  'campo para poner el Nº de CUSSP o ONP
        xControl.TableCell(tcText, 8, 4) = NulosC(RstEmpleados("cuspp"))

        xControl.TableCell(tcColSpan, 8, 11, 8, 12) = 2  'campo para poner el Nº de ESSALUD
        xControl.TableCell(tcText, 8, 11) = NulosC(RstEmpleados("numessalud"))

        xControl.TableCell(tcColSpan, 9, 4, 9, 6) = 3   'campo para poner la fecha de ingreso del trabajador
        xControl.TableCell(tcText, 9, 4) = NulosC(RstEmpleados("fchingreso"))

        xControl.TableCell(tcColSpan, 9, 11, 9, 11) = 2 'campo para poner la fecha de cese del trabajador
        xControl.TableCell(tcText, 9, 11) = NulosC(RstEmpleados("fchcese"))

        xControl.TableCell(tcColSpan, 10, 4, 10, 7) = 4  'campo para poner el cargo
        xControl.TableCell(tcText, 10, 4) = NulosC(RstEmpleados("cargo"))

        xControl.TableCell(tcText, 4, 10) = NulosC(RstEmpleados("numboleta"))
        
        
        xControl.TableCell(tcColSpan, 12, 4, 12, 7) = 4
        xControl.TableCell(tcText, 12, 4) = NulosC(RstEmpleados("diatrabajo"))  'campo dias trabajados
        
        If NulosN(RstEmpleados("diatrabajo")) = 0 Then
            xControl.TableCell(tcText, 12, 4) = ""
                    
        Else
            xControl.TableCell(tcColSpan, 13, 4, 13, 7) = 4
            xControl.TableCell(tcText, 13, 4) = NulosC(RstEmpleados("totalhe1"))  'campo horas extras al 60%
    
            xControl.TableCell(tcColSpan, 12, 11, 12, 12) = 2
            xControl.TableCell(tcText, 12, 11) = NulosC(RstEmpleados("totalhn"))   'campo para el numero de esalud
            
            xControl.TableCell(tcColSpan, 13, 11, 13, 12) = 2
            xControl.TableCell(tcText, 13, 11) = NulosC(RstEmpleados("totalhe2")) 'campo horas extras al 100%
        End If
    End If

       Dim xFila&
       Dim sTotalIng As Double
       Dim sTotalDesc As Double
       Dim sTotalApo As Double
       
       'Extraemos las aportaciones del trabajador
       '--uniendo celdas del detalle
       xFila = 17
       For A = 1 To 12
           
           xControl.TableCell(tcColSpan, xFila, 2, xFila, 3) = 2
           xControl.TableCell(tcColSpan, xFila, 6, xFila, 7) = 2
           xControl.TableCell(tcColSpan, xFila, 10, xFila, 11) = 2
           
           xFila = xFila + 1
       Next A
       
       '--remuneraciones
       xFila = 17
       Do While Not RstIng.EOF
           If NulosN(RstIng("imptot")) <> 0 Then
               xControl.TableCell(tcAlign, xFila, 2, xFila, 2) = taLeftMiddle
               xControl.TableCell(tcText, xFila, 2) = NulosC(RstIng("nomcorto"))
               xControl.TableCell(tcAlign, xFila, 4, xFila, 4) = taRightTop
               xControl.TableCell(tcText, xFila, 4) = Format(NulosN(RstIng("imptot")), FORMAT_MONTO)
               sTotalIng = sTotalIng + NulosN(RstIng("imptot"))
               xFila = xFila + 1
           End If
           RstIng.MoveNext
       Loop
       xControl.TableCell(tcAlign, 26, 4, 26, 4) = taRightTop
       xControl.TableCell(tcText, 26, 4) = Format(sTotalIng, FORMAT_MONTO)
       '--descuentos
       xFila = 17
       Do While Not RstDesc.EOF
           If NulosN(RstDesc("imptot")) <> 0 Then
               xControl.TableCell(tcAlign, xFila, 6, xFila, 6) = taLeftMiddle
               xControl.TableCell(tcText, xFila, 6) = NulosC(RstDesc("nomcorto"))
               xControl.TableCell(tcAlign, xFila, 8, xFila, 8) = taRightTop
               xControl.TableCell(tcText, xFila, 8) = Format(NulosN(RstDesc("imptot")), FORMAT_MONTO)
               sTotalDesc = sTotalDesc + NulosN(RstDesc("imptot"))
               xFila = xFila + 1
           End If
           RstDesc.MoveNext
       Loop
       xControl.TableCell(tcAlign, 26, 8, 26, 8) = taRightTop
       xControl.TableCell(tcText, 26, 8) = Format(sTotalDesc, FORMAT_MONTO)
       
       '--aportes
       xFila = 17
       Do While Not RstApo.EOF
           If NulosN(RstApo("imptot")) <> 0 Then
               xControl.TableCell(tcAlign, xFila, 10, xFila, 10) = taLeftMiddle
               xControl.TableCell(tcText, xFila, 10) = NulosC(RstApo("nomcorto"))
               xControl.TableCell(tcAlign, xFila, 12, xFila, 12) = taRightTop
               xControl.TableCell(tcText, xFila, 12) = Format(NulosN(RstApo("imptot")), FORMAT_MONTO)
               sTotalApo = sTotalApo + NulosN(RstApo("imptot"))
               xFila = xFila + 1
           End If
           RstApo.MoveNext
       Loop
       xControl.TableCell(tcAlign, 26, 12, 26, 12) = taRightTop
       xControl.TableCell(tcText, 26, 12) = Format(sTotalApo, FORMAT_MONTO)
       xControl.TableCell(tcColSpan, 28, 2) = 12
       xControl.TableCell(tcText, 28, 2) = "Neto a Pagar: " & Format(Format((sTotalIng - sTotalDesc), FORMAT_MONTO), FORMAT_MONTO) & "      SON: " & NumeroLetra((sTotalIng - sTotalDesc), NulosN(RstEmpleados("idmon")))
       
       '---------------------

       xControl.TableCell(tcBackColor, 2, 2, 4, 3) = &H8000000F
       xControl.TableCell(tcBackColor, 6, 2, 10, 3) = &H8000000F
       xControl.TableCell(tcBackColor, 12, 2, 13, 3) = &H8000000F

       xControl.TableCell(tcBackColor, 7, 8, 10, 10) = &H8000000F
       xControl.TableCell(tcBackColor, 12, 8, 13, 10) = &H8000000F

       xControl.TableCell(tcBackColor, 15, 2, 16, 4) = &H8000000F
       xControl.TableCell(tcBackColor, 15, 6, 16, 8) = &H8000000F
       xControl.TableCell(tcBackColor, 15, 10, 16, 12) = &H8000000F
       
       'imprimimos lineas y formas
       '-- primera copia
       xControl.BrushColor = &H80&
       xControl.DrawRectangle 8100, 1600, 11080, 2350
       
       xControl.DrawLine 2500, 7540, 4500, 7540
       xControl.DrawLine 8100, 7540, 10100, 7540
       
       '--segunda copia
       Dim mAlto&
       mAlto = 7600
       xControl.BrushColor = &H80&
       xControl.DrawRectangle 8100, mAlto + 1600, 11080, mAlto + 2350
       
       xControl.DrawLine 2500, mAlto + 7540, 4500, mAlto + 7540
       xControl.DrawLine 8100, mAlto + 7540, 10100, mAlto + 7540
               
               
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
