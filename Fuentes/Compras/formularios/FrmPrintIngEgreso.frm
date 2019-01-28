VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrintIngEgreso 
   Caption         =   "Reporte de ingresos y Egresos de Almacén"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VSPrinter7LibCtl.VSPrinter Vp 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   11055
      _cx             =   19500
      _cy             =   11245
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
      Zoom            =   33.0365093499555
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
Attribute VB_Name = "FrmPrintIngEgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPRINTINGEGRESO.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA IMPRIMIR INFORMACION DEL FORMULARIO FrmConsIngAlmacen, ESTE
'*                    FOMULARIO SE INVOCA DESDE FrmConsIngAlmacen
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 18/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit
Dim SeEjecuto As Boolean           ' VARIABLE PARA CONTROLAR QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ

'*****************************************************************************************************
'* Nombre           : CabeceraIngresos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GENERA LA CABECERA PARA CUANDO SEA UN INGRESO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub CabeceraIngresos()
    Dim vTitulo2 As String
    Dim vFecha1 As String, vFecha2 As String, vProveedor As String
    Vp.FontSize = 8
    Vp.TextAlign = taLeftMiddle
    Vp.FontBold = True
    Vp.TextBox NomEmp, 1000, 600, 5000, 400
    Vp.TextBox NumRUC, 1000, 800, 2000, 400
    Vp.TextAlign = taRightMiddle
    Vp.TextBox "Fecha: " & Date & "", 1000, 800, 10200, 400
    Vp.TextAlign = taLeftMiddle
    
    Vp.TextAlign = taCenterMiddle
    Vp.FontSize = 10
    Vp.TextBox "REPORTE DE INGRESOS", 1000, 1000, 10200, 400
    
    vFecha1 = Trim(FrmConsIngAlmacen.TxtFec1.Valor): vFecha2 = Trim(FrmConsIngAlmacen.TxtFec2.Valor)
    Vp.TextBox "DESDE: " & vFecha1 & " HASTA: " & vFecha2 & "", 1000, 1200, 10200, 400
    
    Vp.TextAlign = taLeftMiddle
    
    If FrmConsIngAlmacen.LblProveedor.Caption <> "" Then
        vTitulo2 = "Proveedor: " & FrmConsIngAlmacen.LblProveedor.Caption
    End If
    
    If FrmConsIngAlmacen.lblProducto.Caption <> "" Then
        If vTitulo2 <> "" Then
            vTitulo2 = vTitulo2 & "  "
        End If
        vTitulo2 = vTitulo2 & "Producto: " & FrmConsIngAlmacen.lblProducto.Caption
    End If
    
    Vp.TextAlign = taLeftMiddle
    Vp.FontSize = 8
    Vp.TextBox vTitulo2, 1000, 1500, 10200, 400
    
    FormatoNormal
    Vp.SpaceAfter = 130
    
    Vp.TextBox "Fec. Operac.", 1000, 1900, 900, 300 '
    Vp.TextBox "Tipo Doc.", 2000, 1900, 700, 300 '
    Vp.TextBox "Nro. de Doc.", 2800, 1900, 1100, 300 '
    Vp.TextBox "Fec. Emis.", 4400, 1900, 700, 300 '
    Vp.TextBox "Proveedor", 5200, 1900, 2400, 300 '
    Vp.TextBox "Producto", 7700, 1900, 2200, 300 '
    Vp.TextAlign = taLeftMiddle
    Vp.TextBox "Uni. Med.", 10000, 1900, 600, 300 '
    Vp.TextAlign = taCenterTop
    Vp.TextBox "Cant.", 10700, 1900, 600, 300 '
    
    Vp.DrawLine 800, 1850, 11300, 1850   'PARA LINEA DE ENCABEZ PAGINA HORIZONTAL
    Vp.DrawLine 800, 2200, 11300, 2200
End Sub

'*****************************************************************************************************
'* Nombre           : PrintIngresos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME EL DETALLE DE LOS INGRESOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub PrintIngresos()
    With Vp
        Vp.Top = 0: Vp.Left = 0
        Vp.Orientation = orPortrait
        FormatoNormal
        .TextColor = &H80000008
        ' MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone
        .StartDoc
            CabeceraIngresos
            Dim A As Long, xFila As Long
            xFila = 2300
            For A = 1 To FrmConsIngAlmacen.Fg1.Rows - 1
                FormatoNormal
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 1), 1000, xFila, 900, 300     ' FECHA OPERAC
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 2), 2000, xFila, 700, 300     ' TIPO DOC
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 3), 2800, xFila, 1500, 300    ' NRO DE DOC
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 4), 4400, xFila, 700, 300     ' FECHA EMISION
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 5), 5200, xFila, 2400, 300    ' PROVEEDOR / CLIENTE
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 7), 7700, xFila, 2200, 300    ' PRODUCTO
                .TextAlign = taLeftMiddle
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 8), 10000, xFila, 600, 300    ' UNI MED
                .TextAlign = taRightMiddle
                .TextBox Format(FrmConsIngAlmacen.Fg1.TextMatrix(A, 9), "####0.0000"), 10700, xFila, 600, 300  '  'CANTIDAD
                If xFila >= 14500 Then
                    .NewPage
                    CabeceraIngresos
                    xFila = 2300
                Else
                    xFila = xFila + 200
                End If
            Next
        .EndDoc
        .ScrollIntoView 0, 0, 0, 0
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : cabeceraSalida
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME LA CABECERA DE LAS SALIDAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub cabeceraSalida()
    Dim vFecha1 As String, vFecha2 As String, vProveedor As String
    Dim vTitulo3 As String
    Vp.FontSize = 8
    Vp.TextAlign = taLeftMiddle
    Vp.FontBold = True
    Vp.TextBox "Empresa: Rincon", 1000, 600, 5000, 400
    Vp.TextBox "R.U.C.: 0000000000", 1000, 800, 2000, 400
    Vp.TextAlign = taRightBottom
    Vp.TextBox "Fecha: " & Date & "", 1000, 800, 13000, 400
    
    Vp.TextAlign = taCenterMiddle
    Vp.FontSize = 10
    Vp.TextBox "REPORTE DE SALIDAS", 1000, 1000, 13000, 400
        
    vFecha1 = Trim(FrmConsIngAlmacen.TxtFec1.Valor): vFecha2 = Trim(FrmConsIngAlmacen.TxtFec2.Valor)
    Vp.TextBox "DESDE: " & vFecha1 & " HASTA: " & vFecha2 & "", 1000, 1200, 13000, 400
    
    Vp.TextAlign = taLeftMiddle
    If Trim(FrmConsIngAlmacen.LblProveedor.Caption) <> "" Then
        vTitulo3 = "Cliente: " & FrmConsIngAlmacen.LblProveedor.Caption
    End If
    If vTitulo3 <> "" Then
        vTitulo3 = vTitulo3 & "   "
    End If
    If FrmConsIngAlmacen.lblProducto.Caption <> "" Then
        vTitulo3 = vTitulo3 & "Producto: " & FrmConsIngAlmacen.lblProducto.Caption
    End If
    
    Vp.FontSize = 8
    Vp.TextAlign = taLeftMiddle
    Vp.TextBox vTitulo3, 1000, 1500, 13000, 300

    FormatoNormal
    Vp.TextBox "Fec. Operac.", 1000, 1900, 900, 250 '
    Vp.TextBox "Tipo Doc.", 2000, 1900, 700, 250 '
    Vp.TextBox "Nro. de Doc.", 2800, 1900, 1400, 250 '
    Vp.TextBox "Fec. Emis.", 4300, 1900, 700, 250 '
    Vp.TextBox "Cliente", 5100, 1900, 2500, 250 '
    Vp.TextBox "Producto", 7700, 1900, 2800, 250
    Vp.TextAlign = taLeftMiddle
    Vp.TextBox "Uni. Med.", 10600, 1900, 600, 250
    
    Vp.TextAlign = taCenterTop
    Vp.TextBox "Cant.", 11300, 1900, 600, 250
    
    Vp.TextBox "Area", 12000, 1900, 2000, 250

    Vp.DrawLine 1000, 1850, 14000, 1850 'PARA LINEA DE ENCABEZ PAGINA HORIZONTAL
    Vp.DrawLine 1000, 2200, 14000, 2200
End Sub

'*****************************************************************************************************
'* Nombre           : PrintSalida
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME EL DETALLE DE LA SALIDA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub PrintSalida()
    With Vp
        Vp.Top = 0: Vp.Left = 0
        FormatoNormal
        .TextColor = &H80000008
        ' MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone
        .StartDoc
            cabeceraSalida
            Dim A As Long, xFila As Long
            xFila = 2300
            For A = 1 To FrmConsIngAlmacen.Fg1.Rows - 1
                FormatoNormal
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 1), 1000, xFila, 900, 300   '  'FECHA OPERAC
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 2), 2000, xFila, 700, 300  '   'TIPO DOC
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 3), 2800, xFila, 1400, 300 '    'NRO DE DOC
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 4), 4300, xFila, 700, 300  '   'FECHA EMISION
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 5), 5100, xFila, 2500, 300 '    'PROVEEDOR / CLIENTE
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 7), 7700, xFila, 2800, 300  '  'PRODUCTO
                .TextAlign = taLeftMiddle
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 8), 10600, xFila, 600, 300 '   'UNI MED
                .TextAlign = taRightMiddle
                .TextBox Format(FrmConsIngAlmacen.Fg1.TextMatrix(A, 9), "#####0.0000"), 11300, xFila, 600, 300    'CANTIDAD
                
                Vp.TextAlign = taLeftMiddle
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 11), 12000, xFila, 2000, 300 'AREA
                If xFila >= 10900 Then
                    .NewPage
                    Cabecera
                    xFila = 2300
                Else
                    xFila = xFila + 200
                End If
            Next
        .EndDoc
        .ScrollIntoView 0, 0, 0, 0
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : FormatoNormal
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DA FORMATO AL TEXTO DEL REPORTE
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub FormatoNormal()
    Vp.TextAlign = taLeftMiddle
    Vp.FontSize = 8
    Vp.FontBold = False
End Sub

'*****************************************************************************************************
'* Nombre           : Cabecera
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME EL ENCABEZADO DE LA PAGINA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cabecera()
    Dim vFecha1 As String, vFecha2 As String, vProveedor As String
    
    Vp.FontSize = 13
    Vp.TextAlign = taCenterTop
    Vp.FontBold = True
    Vp.CurrentY = 600
    If FrmConsIngAlmacen.OptIng.Value = True Then
        Vp = "REPORTE DE INGRESOS"
    ElseIf FrmConsIngAlmacen.OptSal.Value = True Then
        Vp = "REPORTE DE SALIDAS"
        Vp.TextBox "Solicitante: " & Trim(FrmConsIngAlmacen.TxtIdSolicitante.Text) & " - " & Trim(FrmConsIngAlmacen.LblSolicitante.Caption) & "", 800, 500, 3000, 400
    End If
    FormatoNormal
    Vp.SpaceAfter = 130
    
    Vp.TextAlign = taCenterMiddle
    Vp.FontSize = 13
    Vp.CurrentY = 1000
    vFecha1 = Trim(FrmConsIngAlmacen.TxtFec1.Valor): vFecha2 = Trim(FrmConsIngAlmacen.TxtFec2.Valor)
    Vp = "DESDE: " & vFecha1 & " HASTA: " & vFecha2 & ""
    
    FormatoNormal
    Vp.TextBox "Fec. Operac.", 300, 1900, 900, 300
    Vp.TextBox "Tipo Doc.", 1300, 1900, 700, 300
    Vp.TextBox "Nro. de Doc.", 2050, 1900, 1100, 300
    Vp.TextBox "Fec. Emis.", 3200, 1900, 700, 300
    If FrmConsIngAlmacen.OptIng.Value = True Then
        Vp.TextBox "Proveedor", 4000, 1900, 2000, 300
    Else
        Vp.TextBox "Cliente", 4000, 1900, 2000, 300
    End If
    Vp.TextBox "Responsable", 6050, 1900, 2000, 300
    Vp.TextBox "Producto", 8150, 1900, 2000, 300
    Vp.TextBox "Cant.", 10200, 1900, 600, 300
    Vp.TextBox "Uni. Med.", 11000, 1900, 600, 300
    If FrmConsIngAlmacen.OptSal.Value = True Then
        Vp.TextBox "Solicitante", 12000, 1900, 2000, 300
        Vp.TextBox "Area", 14300, 1900, 2000, 300
    End If
    
    Vp.DrawLine 800, 1800, 16000, 1800 'PARA LINEA DE ENCABEZ PAGINA HORIZONTAL
    Vp.DrawLine 800, 2200, 16000, 2200
End Sub

'*****************************************************************************************************
'* Nombre           : Cargar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIMIMOS EL DETALLE DEL REPORTE
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cargar()
    With Vp
        Vp.Top = 0: Vp.Left = 0
        FormatoNormal
        .TextColor = &H80000008
        ' MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone
        .StartDoc
            Cabecera
            Dim A As Long, xFila As Long
            xFila = 2300
            For A = 1 To FrmConsIngAlmacen.Fg1.Rows - 1
                FormatoNormal
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 1), 300, xFila, 900, 300      ' FECHA OPERAC
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 2), 1300, xFila, 700, 300     ' TIPO DOC
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 3), 2050, xFila, 1100, 300    ' NRO DE DOC
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 4), 3200, xFila, 700, 300     ' FECHA EMISION
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 5), 4000, xFila, 2000, 300    ' PROVEEDOR / CLIENTE
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 6), 6050, xFila, 2000, 300    ' RESPONSABLE
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 7), 8150, xFila, 2000, 300    ' PRODUCTO
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 8), 10250, xFila, 600, 300    ' CANTIDAD
                .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 9), 11000, xFila, 600, 300    ' UNI MED
                If FrmConsIngAlmacen.OptSal.Value = True Then
                    .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 10), 12000, xFila, 2100, 300 ' SOLICITANTE
                    .TextBox FrmConsIngAlmacen.Fg1.TextMatrix(A, 11), 14300, xFila, 2000, 300 ' AREA
                End If
                
                If xFila >= 10900 Then
                    Vp.DrawLine 800, 11090, 16000, 11090
                    .NewPage
                    Cabecera
                    xFila = 2300
                Else
                    xFila = xFila + 200
                End If
            Next
            
            Vp.DrawLine 800, 11090, 16000, 11090
        .EndDoc
        .ScrollIntoView 0, 0, 0, 0
    End With
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJCUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        If FrmConsIngAlmacen.OptIng.Value = True Then
            PrintIngresos
        Else
            PrintSalida
        End If
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    Vp.PaperSize = pprA4
    Vp.Orientation = orLandscape  'HORIZONTAL
    SeEjecuto = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Vp.Width = Me.Width
    Vp.Height = Me.Height - 500
End Sub
