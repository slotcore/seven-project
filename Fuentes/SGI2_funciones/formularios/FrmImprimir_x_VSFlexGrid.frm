VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmImprimir_x_VSFlexGrid 
   Caption         =   "Impresión de Consulta"
   ClientHeight    =   7305
   ClientLeft      =   3720
   ClientTop       =   2655
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   9615
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Align           =   2  'Align Bottom
      Height          =   7290
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   9615
      _cx             =   16960
      _cy             =   12859
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
      Zoom            =   37.5
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
Attribute VB_Name = "FrmImprimir_x_VSFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----IMPRESION DE DATOS POR VSFlexGrid
'----POR: JOHAN CASTRO
'----06/10/07
'--------------

Option Explicit
Dim M_FILA As Long                  '--INDICA LA DISTANCIA PARA EMPEZAR A IMPRIMIR EN EL REPORTE
                                    '--PUEDE SER EL TITULO, ENCABEZADO O DETALLE
                   
Dim M_POS_INCIAL As Long            '--INDICA LA POSICION EN EL REPORTE PARA IMPRIMIR LA INFORMACION
Dim Q_COLS, Q_ROW, Q_COL, Q_COL_1 As Integer
Dim M_SIZE_RPT As Long              '--INDICA EL ANCHO DEL REPORTE

Dim M_SALTO_HOJA As Long            '--INDICA EL SALDO DE PAGINA,
                                    '--DEPENDE DEL TIPO DE LA ORIENTACION DE LAHOJA
Dim F_ES_HORIZONTAL As Boolean      '--INDICA LA ORIENTACION DE LA HOJA

Dim M_LEFT As Long                  '--DISTANCIA INICIAL PARA EMPEZAR A IMPRIMIR
                                    '--ESTA EN FUNCION DE LA ORIENTACION DE LA HOJA Y EL ANCHO DE LA GRILLA
                                    
Const M_SEPARACION  As Long = 65    '--SEPARACION DE COLUMNAS
Const M_HEIGHT As Long = 170        '--ALTO DE CELDA
Const M_SIZE_DETALLE As Integer = 6.5 '--INDICA EL TAMAÑO DE LA LETRA DEL DETALLE DEL REPORTE
                                    '--SI SE CAMBIA A UN TAMAÑO SUPERIOR, AUMENTAR M_HEIGHT,
                                    '--PUES POSIBLEMENTE NO SE VEA LA INFORMACION EN EL REPORTE
                                    
Const M_SIZE_ENCABEZADO As Integer = 7 '--INDICA EL TAMAÑO DE LA LETRA DEL ENCABEZADO
Const M_HEIGHT_ENCABEZADO_TMP As Integer = 225 '--ALTO DE CELDA
Dim M_HEIGHT_ENCABEZADO As Integer
Dim Q_COL_INICIAL As Integer        '-- INDICA LA POSICION INICIAL DE GRILLA A IMPRIMIR

Const M_SIZE_EMPRESA As Integer = 8     '--TAMAÑO DE TEXTO DE ENCABEZADO EMPRESA
Const M_HEIGHT_EMPRESA As Integer = 300 '--ALTO DE CELDA

Const M_SIZE_RUC As Integer = 7         '--TAMAÑO DE TEXTO DE ENCABEZADO RUC DE EMPRESA
Const M_HEIGHT_RUC As Integer = 300     '--ALTO DE CELDA

Const M_SIZE_SISTEMA = 7                '--TAMAÑO DE TEXTO DE NOMBRE SISTEMA
Const M_HEIGHT_SISTEMA As Integer = 180 '--ALTO DE CELDA

Dim SGI_JC As New SGI2_funciones.JC_Varios

 
'

''Public Property Let proptitulo1(pData As String)
''    m_titulo1 = pData
''End Property
''Public Property Let proptitulo2(pData As String)
''    m_titulo2 = pData
''End Property



'------------------------------------------------
Sub PONER_TITULO_FRM(N_CAPTION As String)
    If N_CAPTION = "" Then
        Me.Caption = "Impresión de Consulta"
    Else
        Me.Caption = N_CAPTION
    End If
End Sub

Private Sub PONER_NOMBRE_SISTEMA()
    vp.TextAlign = taRightBottom
    vp.FontSize = M_SIZE_SISTEMA
    vp.FontBold = False
    vp.TextBox Nomsis, M_LEFT, 500, M_SIZE_RPT, M_HEIGHT_SISTEMA
   
End Sub

Private Sub PONER_TITULO(T_TITULO As String, _
                        T_TITULO_1 As String, _
                        T_PERIODO As String, _
                        Optional F_TITULO_EN_HOJAS As Boolean = True)
                
    '--ESTA FUNCION RECIBIRA COMO PARAMETRO EL TITULO DEL REPORTE
    '--TAMBIEN EL PERIODO DE LA CONSULTA
    M_FILA = 500
    If F_TITULO_EN_HOJAS = False Then Exit Sub
    '--
    PONER_EMPRESA
    PONER_NOMBRE_SISTEMA
    PONER_FECHA
    '--
    M_FILA = 650
    If F_ES_HORIZONTAL = False Then M_FILA = 750
    vp.TextAlign = taCenterMiddle
    If T_TITULO <> "" Then '--DEL TITULO
        vp.FontSize = 10
        vp.FontBold = True
        vp.TextBox T_TITULO, M_LEFT, M_FILA, M_SIZE_RPT, 600
    End If
    If T_PERIODO <> "" Then '--DEL PERIODO
        M_FILA = M_FILA + 350
        vp.FontSize = 7
        vp.TextBox T_PERIODO, M_LEFT, M_FILA, M_SIZE_RPT, 400
        vp.FontBold = False
    End If
    If T_TITULO_1 <> "" Then '--DEL TITULO_1
        M_FILA = M_FILA + 250
        vp.FontSize = 8
        vp.TextBox T_TITULO_1, M_LEFT, M_FILA, M_SIZE_RPT, 400
        vp.FontBold = False
        M_FILA = M_FILA + 100
    End If
    vp.TextAlign = taLeftMiddle
'--
End Sub

Private Sub PONER_EMPRESA()
    vp.TextAlign = taLeftMiddle
    vp.FontSize = M_SIZE_EMPRESA
    vp.FontBold = False
    vp.TextBox "Empresa: " + NomEmp, M_LEFT, M_FILA, M_SIZE_RPT, M_HEIGHT_EMPRESA
    M_FILA = M_FILA + 150
    vp.FontSize = M_SIZE_RUC
    vp.TextBox "R.U.C.:  " + NumRUC, M_LEFT, M_FILA, M_SIZE_RPT, M_HEIGHT_RUC
End Sub

Private Sub PONER_FECHA()
    '--DE LA FECHA
    vp.TextAlign = taRightMiddle
    vp.FontSize = 7
    vp.FontBold = False
    vp.TextBox Format(Date, "dd/mm/yy"), M_LEFT, 650, M_SIZE_RPT, 200
End Sub

Private Sub PONER_ENCABEZADO(GRID As Object, _
                            Optional F_TITULO_EN_HOJAS As Boolean = True, _
                            Optional F_ENCABEZADO_EN_HOJAS As Boolean = True, _
                            Optional F_SALTO_HOJA As Boolean = False)
  
    Dim GRUPO_TEXTO As String
    Dim GRUPO_CUENTA As Integer
    Dim M_CELDA_ZISE As Integer '--ES EL TAMAÑO HORIZONTAL DE LA CELDA
    Dim Q_ROW1 As Long
    Dim Q_COL_ANTERIOR As Integer '--ES LA POSICION ANTERIOR A LA COLUMNA ACTUAL
    
    If F_ENCABEZADO_EN_HOJAS = False Then Exit Sub
    If F_TITULO_EN_HOJAS = True Then
        M_FILA = M_FILA + 400
    End If
    
    vp.FontSize = M_SIZE_ENCABEZADO
    vp.FontBold = True
    '***************************************************************************
    '--LINEA SUPERIOR
    vp.DrawLine M_LEFT, M_FILA - 40, M_SIZE_RPT + M_LEFT, M_FILA - 40
    '***************************************************************************
    
    For Q_ROW1 = 0 To GRID.FixedRows - 1
        M_POS_INCIAL = M_LEFT
        GRUPO_CUENTA = 0
        '--SABER EL ENCABEZADO
        M_HEIGHT_ENCABEZADO = GRID.RowHeight(Q_ROW1)
        If M_HEIGHT_ENCABEZADO < M_HEIGHT_ENCABEZADO_TMP Then
            M_HEIGHT_ENCABEZADO = M_HEIGHT_ENCABEZADO_TMP
        End If
        '----
        For Q_COL = 1 To GRID.Cols - 1
            If GRID.ColWidth(Q_COL) <> 0 Then
                GRID.Row = Q_ROW1
                GRID.Col = Q_COL

                If Q_COL_INICIAL = Q_COL Then
                    GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW1, Q_COL))
                    Q_COL_ANTERIOR = Q_COL
                Else
                    Q_COL_ANTERIOR = OBTENER_COL_ANTERIOR(GRID, CInt(Q_COL))
                End If
                '-----
                vp.TextAlign = COL_ALINEACION(GRID, CInt(Q_COL))
                '--COLOR AL TEXTO
                vp.TextColor = GRID.CellForeColor
                '--VER GRUPOS
                If GRID.MergeCells = flexMergeFree And GRID.MergeRow(Q_ROW1) = True And CStr(GRID.TextMatrix(Q_ROW1, Q_COL_ANTERIOR)) = CStr(GRID.TextMatrix(Q_ROW1, Q_COL)) Then
                    If GRUPO_CUENTA = 0 Then
                        GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW1, Q_COL))
                        M_CELDA_ZISE = TAMANO_CELDA_GRUPO(GRID, CInt(Q_ROW1), CInt(Q_COL))
                    End If
                Else
                    GRUPO_TEXTO = "xxxxxxxxx"
                    GRUPO_CUENTA = 0
                    M_CELDA_ZISE = GRID.ColWidth(Q_COL)
                End If
                '--------
                If (GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW1, Q_COL)) Or GRUPO_TEXTO = "xxxxxxxxx") And GRUPO_CUENTA = 0 Then
                    vp.TextBox CStr(GRID.TextMatrix(Q_ROW1, Q_COL)), M_POS_INCIAL, M_FILA, M_CELDA_ZISE, M_HEIGHT_ENCABEZADO
                End If
                
                If GRID.MergeCells = flexMergeFree Then GRUPO_CUENTA = 1
                       
                M_POS_INCIAL = M_POS_INCIAL + M_SEPARACION + GRID.ColWidth(Q_COL)

                
                vp.TextColor = vbBlack
                '--
            End If
        Next Q_COL
        
         M_FILA = M_FILA + 200
       
        '---
        If M_HEIGHT_ENCABEZADO_TMP < M_HEIGHT_ENCABEZADO Then
            M_FILA = M_FILA + (M_HEIGHT_ENCABEZADO - M_HEIGHT_ENCABEZADO_TMP)
        End If
        '---
    Next Q_ROW1
    '***************************************************************************
    '--LINEA INFERIOR
    vp.DrawLine M_LEFT, M_FILA + 40, M_SIZE_RPT + M_LEFT, M_FILA + 40
    '***************************************************************************
    If F_SALTO_HOJA = True Then M_FILA = M_FILA + 200
    vp.FontBold = False
    'vp.DrawLine 800, 2200, 16500, 2200
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    M_LEFT = 300
    Nomsis = LeerLineaINI(Trim(App.PATH) + "\sgi.ini", "NOMBRE", "SOFTWARE")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vp.Height = Me.Height - 500
    Err.Clear
End Sub

Sub PONER_DATOS(GRID As Object, _
                T_TITULO As String, _
                Optional T_TITULO_1 As String = "", _
                Optional T_PERIODO As String = "", _
                Optional F_TITULO_EN_HOJAS As Boolean = True, _
                Optional F_ENCABEZADO_EN_HOJAS As Boolean = True)
                
    On Error GoTo ERROR
    Dim GRUPO_TEXTO As String
    Dim GRUPO_CUENTA As Integer
    Dim M_CELDA_ZISE As Integer '--ES EL TAMAÑO HORIZONTAL DE LA CELDA
    Dim Q_COL_ANTERIOR As Integer '--ES LA POSICION ANTERIOR A LA COLUMNA ACTUAL
    With vp
        .Top = 0 ': vp.Left = 0
        .PaperSize = pprA4
        .FontName = "Arial" '"Courier New"
        .Zoom = 75
        .FontSize = 10
        .TextColor = &H80000008 'RGB(200, 200, 200)
        'MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone
        If ES_HORIZONTAL(GRID) = True Then
            .Orientation = orLandscape
        Else
            .Orientation = orPortrait
        End If
       
        .StartDoc
        
        PONER_TITULO T_TITULO, T_TITULO_1, T_PERIODO

        '------
        PONER_ENCABEZADO GRID, True
        M_FILA = M_FILA + 200
        vp.FontSize = M_SIZE_DETALLE
        For Q_ROW = GRID.FixedRows To GRID.Rows - 1
            DoEvents
            M_POS_INCIAL = M_LEFT
            GRUPO_CUENTA = 0
            For Q_COL = 1 To GRID.Cols - 1
                
                If GRID.ColWidth(Q_COL) <> 0 Then
                    '--ALINEACION
                    GRID.Row = Q_ROW
                    GRID.Col = Q_COL
                    '-----
                    If Q_COL_INICIAL = Q_COL Then
                        GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW, Q_COL))
                        Q_COL_ANTERIOR = Q_COL
                    Else
                        Q_COL_ANTERIOR = OBTENER_COL_ANTERIOR(GRID, CInt(Q_COL))
                    End If
                    '-----
                    .TextAlign = COL_ALINEACION(GRID, CInt(Q_COL))
                    '--COLOR AL TEXTO
                    .TextColor = GRID.CellForeColor
                    '--VER GRUPOS
                    If GRID.MergeCells = flexMergeFree And GRID.MergeRow(Q_ROW) = True And CStr(GRID.TextMatrix(Q_ROW, Q_COL_ANTERIOR)) = CStr(GRID.TextMatrix(Q_ROW, Q_COL)) Then
                        If GRUPO_CUENTA = 0 Then
                            GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW, Q_COL))
                            M_CELDA_ZISE = TAMANO_CELDA_GRUPO(GRID, CInt(Q_ROW), CInt(Q_COL))
                        End If
                    Else
                        GRUPO_TEXTO = "xxxxxxxxx"
                        GRUPO_CUENTA = 0
                        M_CELDA_ZISE = GRID.ColWidth(Q_COL)
                    End If
                    '--------
                    If (GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW, Q_COL)) Or GRUPO_TEXTO = "xxxxxxxxx") And GRUPO_CUENTA = 0 Then
                        .TextBox CStr(GRID.TextMatrix(Q_ROW, Q_COL)), M_POS_INCIAL, M_FILA, M_CELDA_ZISE, M_HEIGHT
                    End If
                    
                    If GRID.MergeCells = flexMergeFree Then GRUPO_CUENTA = 1
                   
                    M_POS_INCIAL = M_POS_INCIAL + M_SEPARACION + GRID.ColWidth(Q_COL)
                    '--
                    vp.TextColor = vbBlack
                    '--
                End If
            Next Q_COL
            If M_FILA >= M_SALTO_HOJA Then
                .NewPage
                PONER_TITULO T_TITULO, T_TITULO_1, T_PERIODO, F_TITULO_EN_HOJAS
                PONER_ENCABEZADO GRID, F_TITULO_EN_HOJAS, F_ENCABEZADO_EN_HOJAS, True
                .FontSize = M_SIZE_DETALLE
            Else
                M_FILA = M_FILA + 200
            End If
                        
        Next Q_ROW
        '------
        .EndDoc
        .ScrollIntoView 0, 0, 0, 0

    End With
    Me.MousePointer = vbDefault
    Exit Sub
ERROR:
    Me.MousePointer = vbDefault
    SGI_JC.SHOW_ERROR
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SGI_JC = Nothing
    vp.Clear
End Sub

Private Function ES_HORIZONTAL(GRID As Object) As Boolean
    
    M_SIZE_RPT = 0
    For Q_COL = 1 To GRID.Cols - 1
        If GRID.ColWidth(Q_COL) <> 0 Then
            If M_SIZE_RPT = 0 Then Q_COL_INICIAL = Q_COL
            M_SIZE_RPT = M_SIZE_RPT + GRID.ColWidth(Q_COL) + M_SEPARACION
        End If
    Next Q_COL
    If M_SIZE_RPT > 10900 Then
        ES_HORIZONTAL = True
        M_SALTO_HOJA = 11000 '--SE CAMBIA HOJA
        F_ES_HORIZONTAL = True
    Else
        M_SALTO_HOJA = 14500
    End If
    
    '--OBTENER EL INICIO DE IMPRESION DE DATOS
    'M_LEFT = 300
    If ES_HORIZONTAL = True Then
        M_LEFT = (15800 - M_SIZE_RPT) / 2
    Else
        M_LEFT = (12000 - M_SIZE_RPT) / 2
    End If
    If M_LEFT < 0 Then M_LEFT = 300
    If M_SIZE_RPT < 5000 Then
        M_LEFT = 300
        M_SIZE_RPT = 8000
    End If
End Function


Private Function TAMANO_CELDA_GRUPO(GRID As Object, K_ROW As Integer, Q_COL_INI As Integer) As Integer
    '--ESTA FUNCION CALCULARA EL TAMAÑO HORIZONTAL DEL GRUPO
    Dim X_POS  As Integer
    Dim N_VALOR As String
    Dim M_ZISE_GRUPO As Integer
    
    N_VALOR = CStr(GRID.TextMatrix(K_ROW, Q_COL_INI))
    For X_POS = Q_COL_INI + 1 To GRID.Cols - 1
        If GRID.ColWidth(X_POS) <> 0 Then
            If GRID.MergeCells = flexMergeFree And GRID.MergeRow(K_ROW) = True And N_VALOR = CStr(GRID.TextMatrix(K_ROW, X_POS)) Then
                M_ZISE_GRUPO = M_ZISE_GRUPO + GRID.ColWidth(X_POS) + M_SEPARACION
            Else
                Exit For
            End If
        End If
    Next
    M_ZISE_GRUPO = M_ZISE_GRUPO + GRID.ColWidth(Q_COL_INI) + M_SEPARACION
    TAMANO_CELDA_GRUPO = M_ZISE_GRUPO
End Function


Private Function OBTENER_COL_ANTERIOR(GRID As Object, Q_COL_INI As Integer) As Integer

    Dim X_POS  As Integer
    Dim N_VALOR As String
    Dim M_ZISE_GRUPO As Integer
    If X_POS = 1 Then
        OBTENER_COL_ANTERIOR = 1
        Exit Function
    End If
    For X_POS = Q_COL_INI - 1 To 1 Step -1
        If GRID.ColWidth(X_POS) <> 0 Then
            OBTENER_COL_ANTERIOR = X_POS
            Exit Function
        End If
    Next
    
End Function


Private Function COL_ALINEACION(GRID As Object, Col As Integer) As Variant
    '--ESTA FUNCION DEVOLVERA LA CONSTANTE DE ALINEACION QUE SOPORTA EL REPORTE EN FUNCION A LA ALINEACION DEL GRID
    '--------------------------------------------
    '--------------------------------------------
    '--                     reporte  VSFlexGrid
    '--Center   Bottom       4       5
    '--Center   Middle       7       - (*)
    '--Center   Top          1       3
    '--Center   Center       -       4
    '--Just     Bottom       10      -
    '--Just     Middle       11      -
    '--Just     Top          9       -
    '--Just     Center       -       -
    '--Left     Bottom       3       2
    '--Left     Middle       6       - (*)
    '--Left     Top          0       0
    '--Left     Center       -       1
    '--Rigth    Bottom       5       8
    '--Rigth    Middle       8       - (*)
    '--Rigth    Top          2       6
    '--Rigth    Center       -       7
    '--------------------------------------------
    '--------------------------------------------
    Dim ALINEACION As Integer
    Select Case GRID.ColAlignment(Col)
        Case 3, 4, 5:   ALINEACION = 7 '--CENTRADO
        Case 2, 0, 1:   ALINEACION = 6 '--DERECHO
        Case 6, 7, 8:   ALINEACION = 8 '--IZQUIERDO
        Case Else
            ALINEACION = 7
    End Select
    COL_ALINEACION = ALINEACION
End Function


Public Sub PONER_GRID_EN_RPT(GRID As Object)
 With vp
        .Top = 0 ': vp.Left = 0
        .PaperSize = pprA4
        .FontName = "Arial" '"Courier New"
        .Zoom = 75
        .FontSize = 10
        .TextColor = &H80000008 'RGB(200, 200, 200)
        'MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone
        If ES_HORIZONTAL(GRID) = True Then
            .Orientation = orLandscape
        Else
            .Orientation = orPortrait
        End If
    .StartDoc
    
    .RenderControl = GRID.hWnd
    .EndDoc
End With
End Sub


Private Sub S(GRID As Object, Q_ROW2 As Integer)
Dim GRID1 As VSFlexGrid
GRID1.M_HEIGHT_ENCABEZADO
End Sub
