VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form FrmPrintKardex 
   Caption         =   "Contabilidad - Reporte Kardex"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "FrmPrintKardex.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   11910
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
      Zoom            =   42.3295454545455
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
Attribute VB_Name = "FrmPrintKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean

Dim ARR_ACUMULA(1, 3) As Double   '--array que acumulara totales por pagina
                                '--ARR_ACUMULA(?,0) ::unidades entradas
                                '--ARR_ACUMULA(?,1) ::unidades salida
                                '--ARR_ACUMULA(?,2) ::importes entradas
                                '--ARR_ACUMULA(?,3) ::importes salida
                                
                                '--ARR_ACUMULA(0,?) ::totales por producto
                                '--ARR_ACUMULA(1,?)::totales por hoja

Dim UnRegistro  As Boolean '--indica el si ya se mostro el primer registro para que al momento de imprimir el segundo
                            '--imprima filas en blanco para indentificar saltos de grupo
                            
Dim mSeparacion& '--indica la separacion que tendra cuando se  desee imprimir por modulo de almacen

Dim xFila&

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        If MostrarValorizado = True Then
            Me.Caption = "Contabilidad - Consulta de Kardex"
            mSeparacion = 0
        Else
            Me.Caption = "Almacén - Consulta de Kardex"
            mSeparacion = 150
        End If
        Me.WindowState = 2
    End If
End Sub

Private Sub Form_Load()
    Vp.PaperSize = pprA4
    SeEjecuto = False
End Sub

Sub Cargar2()
    If MostrarValorizado = True Then
        Me.Caption = "Contabilidad - Consulta de Kardex"
        mSeparacion = 0
    Else
        Me.Caption = "Almacén - Consulta de Kardex"
        mSeparacion = 150
    End If
    
    With Vp
        .FontName = "Courier New"
        .FontSize = 10
        .TextColor = &H80000008
        .Orientation = orPortrait
        
        'MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone
        '.PageBorder = pbTopBottom
        Erase ARR_ACUMULA()
        .StartDoc
        
        Cabecera
        
        Dim A&, mRowResumen&
        xFila = 2000
        UnRegistro = False
        .FontSize = 7
        '-----
        .TextAlign = taLeftTop
        '--codigo
        .TextBox "Código", 700, xFila, 950, "200"
        .TextBox ": " & FrmVerKardex.txtCodItem.Text, 1800, xFila, 1500, "200"
        '--unidad
        .TextBox "Unidad:", 3700, xFila, 950, "200"
        .TextBox FrmVerKardex.TxtUnidad.Text, 4300, xFila, 950, "200"
        '--descripcion
        pCompararUltimaFila
        .TextBox "Descripción", 700, xFila, 950, "200"
        .TextBox ": " & FrmVerKardex.TxtDesc.Text, 1800, xFila, 6500, "200"
        
        xFila = xFila + 250
        
        For A = FrmVerKardex.Fg1.FixedRows To FrmVerKardex.Fg1.Rows - 1
            DoEvents
            .TextAlign = taLeftTop
            '--fecha
            .CurrentX = 700:  .CurrentY = xFila: .Paragraph = FrmVerKardex.Fg1.TextMatrix(A, 1)
            '--T.D.
            .CurrentX = 1400 + mSeparacion * 1: .CurrentY = xFila: .Paragraph = FrmVerKardex.Fg1.TextMatrix(A, 2)
            '--numero doc
            .CurrentX = 1650 + mSeparacion * 2: .CurrentY = xFila: .Paragraph = FrmVerKardex.Fg1.TextMatrix(A, 3)
            
            .TextAlign = taRightTop
            '--ingreso entrada
            .TextBox FrmVerKardex.Fg1.TextMatrix(A, 4), 2550 + mSeparacion * 3, xFila, 900, "200"
            '--ingreso salida
            .TextBox FrmVerKardex.Fg1.TextMatrix(A, 5), 3400 + mSeparacion * 4, xFila, 900, "200"
            '--ingreso saldo
            .TextBox FrmVerKardex.Fg1.TextMatrix(A, 6), 4300 + mSeparacion * 5, xFila, 900, "200"
            
            '********************************************
            If MostrarValorizado = True Then
                '** contable
                '--ingreso precio unitario
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 7), 5300, xFila, 800, "200"
                '--importes entrada
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 8), 6100, xFila, 800, "200"
                '--importes salida
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 9), 6700, xFila, 800, "200"
                '--importes saldo
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 10), 7500, xFila, 800, "200"
                '--precio promedio
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 11), 8500, xFila, 800, "200"
            End If
             '********************************************
            
            .TextAlign = taLeftTop
            '--cliente/proveedor - origen/destino
            If MostrarValorizado = True Then
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 12), 9600, xFila, 3000, "200"
            Else
                '--cliente/proveedor - origen/destino
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 12), 5400 + mSeparacion * 6, xFila, 1700, "200"
                '--Nº Documento
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 13), 7200 + mSeparacion * 6, xFila, 1800, "200"
                '--Modulo
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 14), 9250 + mSeparacion * 6, xFila, 1800, "200"
            End If
            '--acumulando los totales
            
            ARR_ACUMULA(0, 0) = ARR_ACUMULA(0, 0) + NulosN(FrmVerKardex.Fg1.TextMatrix(A, 4)) '--unidades entrada
            ARR_ACUMULA(0, 1) = ARR_ACUMULA(0, 1) + NulosN(FrmVerKardex.Fg1.TextMatrix(A, 5)) '--unidades salida
            '--
            ARR_ACUMULA(0, 2) = ARR_ACUMULA(0, 2) + NulosN(FrmVerKardex.Fg1.TextMatrix(A, 8)) '--importes entrada
            ARR_ACUMULA(0, 3) = ARR_ACUMULA(0, 3) + NulosN(FrmVerKardex.Fg1.TextMatrix(A, 9)) '--importes salida
            
            pCompararUltimaFila
            
        Next A
        Vp.DrawLine 700, 15100, 11000, 15100
        .EndDoc
        .ScrollIntoView 0, 0
    End With
    Erase ARR_ACUMULA()
End Sub

Sub Cargar()
    If MostrarValorizado = True Then
        Me.Caption = "Contabilidad - Consulta de Kardex"
        mSeparacion = 0
    Else
        Me.Caption = "Almacén - Consulta de Kardex"
        mSeparacion = 150
    End If
    
    With Vp
        .FontName = "Courier New"
        .FontSize = 10
        .TextColor = &H80000008 'RGB(200, 200, 200)
        '.PenColor = RGB(190, 190, 190)
        .Orientation = orPortrait
        
        'MUESTRA BORDES DEL REPORTE
        .PageBorder = pbNone
        '.PageBorder = pbTopBottom
        Erase ARR_ACUMULA()
        .StartDoc
        
        Me.MousePointer = vbHourglass
        
        Cabecera
        
        Dim A&, mRowResumen&
        xFila = 2000
        UnRegistro = False
        .FontSize = 7
        For mRowResumen = FrmResuMov.Fg1.FixedRows To FrmResuMov.Fg1.Rows - 1
            DoEvents
            If NulosN(FrmResuMov.Fg1.TextMatrix(mRowResumen, 5)) = 0 And NulosN(FrmResuMov.Fg1.TextMatrix(mRowResumen, 6)) = 0 Then GoTo SiguienteReg:
            'If UnRegistro = True Then xFila = xFila + 200
            '--descargar el form detalle
            
            
            Unload FrmVerKardex
            FrmVerKardex.Visible = False
            FrmVerKardex.txtCodItem.Text = FrmResuMov.Fg1.TextMatrix(mRowResumen, 1)
            FrmVerKardex.LblIdProducto.Caption = FrmResuMov.Fg1.TextMatrix(mRowResumen, 10)
            FrmVerKardex.TxtFchIni.Valor = FrmResuMov.TxtFchIni.Valor
            FrmVerKardex.TxtFchFin.Valor = FrmResuMov.TxtFchFin.Valor

            FrmVerKardex.Visible = False
            FrmVerKardex.pCargarRpt
            '*******************************************************************************
            '--colocando los encabezados
            
            pCompararUltimaFila
            
            xFila = xFila - 160
            
            .TextAlign = taLeftTop
            '--codigo
            .TextBox "Código", 700, xFila, 950, "200"
            .TextBox ": " & FrmResuMov.Fg1.TextMatrix(mRowResumen, 1), 1800, xFila, 1500, "200"
            '--unidad
            .TextBox "Unidad:", 3700, xFila, 950, "200"
            .TextBox FrmResuMov.Fg1.TextMatrix(mRowResumen, 3), 4300, xFila, 950, "200"
            '--descripcion
            pCompararUltimaFila
            .TextBox "Descripción", 700, xFila, 950, "200"
            .TextBox ": " & FrmResuMov.Fg1.TextMatrix(mRowResumen, 2), 1800, xFila, 6500, "200"
            
            pCompararUltimaFila
            
            UnRegistro = True

            
            '-----
            For A = FrmVerKardex.Fg1.FixedRows To FrmVerKardex.Fg1.Rows - 1
                DoEvents
                .TextAlign = taLeftTop
                '--fecha
                .CurrentX = 700:  .CurrentY = xFila: .Paragraph = FrmVerKardex.Fg1.TextMatrix(A, 1)
                '--T.D.
                .CurrentX = 1400 + mSeparacion * 1: .CurrentY = xFila: .Paragraph = FrmVerKardex.Fg1.TextMatrix(A, 2)
                '--numero doc
                .CurrentX = 1650 + mSeparacion * 2: .CurrentY = xFila: .Paragraph = FrmVerKardex.Fg1.TextMatrix(A, 3)
                
                .TextAlign = taRightTop
                '--ingreso entrada
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 4), 2550 + mSeparacion * 3, xFila, 900, "200"
                '--ingreso salida
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 5), 3400 + mSeparacion * 4, xFila, 900, "200"
                '--ingreso saldo
                .TextBox FrmVerKardex.Fg1.TextMatrix(A, 6), 4300 + mSeparacion * 5, xFila, 900, "200"
                
                '********************************************
                If MostrarValorizado = True Then
                    '** contable
                    '--ingreso precio unitario
                    .TextBox FrmVerKardex.Fg1.TextMatrix(A, 7), 5300, xFila, 800, "200"
                    '--importes entrada
                    .TextBox FrmVerKardex.Fg1.TextMatrix(A, 8), 6100, xFila, 800, "200"
                    '--importes salida
                    .TextBox FrmVerKardex.Fg1.TextMatrix(A, 9), 6700, xFila, 800, "200"
                    '--importes saldo
                    .TextBox FrmVerKardex.Fg1.TextMatrix(A, 10), 7500, xFila, 800, "200"
                    '--precio promedio
                    .TextBox FrmVerKardex.Fg1.TextMatrix(A, 11), 8500, xFila, 800, "200"
                End If
                 '********************************************
                
                .TextAlign = taLeftTop
                '--cliente/proveedor - origen/destino
                If MostrarValorizado = True Then
                    .TextBox FrmVerKardex.Fg1.TextMatrix(A, 12), 9600, xFila, 3000, "200"
                Else
                    '--cliente/proveedor - origen/destino
                    .TextBox FrmVerKardex.Fg1.TextMatrix(A, 12), 5400 + mSeparacion * 6, xFila, 1700, "200"
                    '--Nº Documento
                    .TextBox FrmVerKardex.Fg1.TextMatrix(A, 13), 7200 + mSeparacion * 6, xFila, 1800, "200"
                    '--Modulo
                    .TextBox FrmVerKardex.Fg1.TextMatrix(A, 14), 9250 + mSeparacion * 6, xFila, 1800, "200"
                End If
                '--acumulando los totales
                
                ARR_ACUMULA(0, 0) = ARR_ACUMULA(0, 0) + NulosN(FrmVerKardex.Fg1.TextMatrix(A, 4)) '--unidades entrada
                ARR_ACUMULA(0, 1) = ARR_ACUMULA(0, 1) + NulosN(FrmVerKardex.Fg1.TextMatrix(A, 5)) '--unidades salida
                '--
                ARR_ACUMULA(0, 2) = ARR_ACUMULA(0, 2) + NulosN(FrmVerKardex.Fg1.TextMatrix(A, 8)) '--importes entrada
                ARR_ACUMULA(0, 3) = ARR_ACUMULA(0, 3) + NulosN(FrmVerKardex.Fg1.TextMatrix(A, 9)) '--importes salida
                
                pCompararUltimaFila
            Next A
            
            Unload FrmVerKardex
SiguienteReg:
        Next mRowResumen
        Vp.DrawLine 700, 15100, 11000, 15100
        .EndDoc
        .ScrollIntoView 0, 0
    End With
    Erase ARR_ACUMULA()
    Me.MousePointer = vbDefault
End Sub

Sub Cabecera()
    Dim nPeriodo As String
    '--del periodo
    If CDate(FrmResuMov.TxtFchIni.Valor) < CDate(FrmResuMov.TxtFchFin.Valor) Then
        nPeriodo = " Del: " + FrmResuMov.TxtFchIni.Valor + " Al: " + FrmResuMov.TxtFchFin.Valor
    Else
        nPeriodo = "Al: " + FrmResuMov.TxtFchIni.Valor
    End If

    Vp.TextAlign = taLeftTop
    Vp.FontSize = 6
    Vp.CurrentX = 700: Vp.CurrentY = 600: Vp.Paragraph = NomEmp
    Vp.CurrentX = 9400: Vp.CurrentY = 600: Vp.Paragraph = "Fecha: " + Format(Date, "dd/mm/yy")
    Vp.CurrentX = 9400: Vp.CurrentY = 750: Vp.Paragraph = "Hora: " + Format(Now(), "hh:mm:ss AM/PM")
    Vp.CurrentX = 9400: Vp.CurrentY = 900: Vp.Paragraph = "Pág.: " + CStr(Vp.PageCount)
    
    Vp.CurrentX = 700: Vp.CurrentY = 750: Vp.Paragraph = "R.U.C. Nº : " + NumRUC
    
    Vp.TextAlign = taCenterMiddle
    Vp.FontSize = 10
    Vp.TextBox "CONSULTA DE KARDEX", 700, 900, 10000, "300"
    Vp.FontSize = 8
    Vp.TextBox nPeriodo, 700, 1100, 10000, "200"

    Vp.FontSize = 6
    
    Vp.DrawLine 700, 1500, 11400, 1500
    
    Vp.TextAlign = taLeftTop
    Vp.CurrentX = 700:    Vp.CurrentY = 1550:  Vp.Paragraph = "Fch.Doc"
    Vp.CurrentX = 1350 + mSeparacion * 1: Vp.CurrentY = 1550: Vp.Paragraph = "T.D."
    Vp.CurrentX = 1700 + mSeparacion * 2: Vp.CurrentY = 1550: Vp.Paragraph = "NºDocumento"
    
    Vp.CurrentX = 3900 + mSeparacion * 3: Vp.CurrentY = 1550: Vp.Paragraph = "Unidades"
    Vp.DrawLine 3135 + mSeparacion * 3, 1700, 5215 + mSeparacion * 4, 1700
    Vp.CurrentX = 3100 + mSeparacion * 3: Vp.CurrentY = 1750: Vp.Paragraph = "Entradas"
    Vp.CurrentX = 3900 + mSeparacion * 4: Vp.CurrentY = 1750: Vp.Paragraph = "Salidas"
    Vp.CurrentX = 4700 + mSeparacion * 5: Vp.CurrentY = 1750: Vp.Paragraph = "Saldo"
    
    If MostrarValorizado = False Then
        Vp.TextBox "Cliente/Prov.- Origen/Dest.", 5200 + mSeparacion * 6, 1650, 2250, "200"
        Vp.TextBox "Nº Documento", 7600 + mSeparacion * 6, 1650, 2000, "200"
        Vp.TextBox "Modulo", 9600 + mSeparacion * 6, 1650, 1500, "200"
    Else
        Vp.CurrentX = 5500:   Vp.CurrentY = 1650:  Vp.Paragraph = "P.Unit."

        Vp.CurrentX = 7000:   Vp.CurrentY = 1550:  Vp.Paragraph = "Importes"
        Vp.DrawLine 6400, 1700, 8300, 1700
        Vp.CurrentX = 6400:   Vp.CurrentY = 1750:  Vp.Paragraph = "Entradas"
        Vp.CurrentX = 7100:   Vp.CurrentY = 1750:  Vp.Paragraph = "Salidas"
        Vp.CurrentX = 7900:   Vp.CurrentY = 1750:  Vp.Paragraph = "Saldo"

        Vp.CurrentX = 8700:   Vp.CurrentY = 1650:  Vp.Paragraph = "P.Prom."
        
        Vp.TextAlign = taCenterTop
        Vp.TextBox "Cliente/Proveedor", 9500, 1550, 1500, "200"
        Vp.TextBox "- Origen/Destino", 9600, 1700, 1500, "200"
        
    End If
    
    Vp.DrawLine 700, 1900, 11400, 1900
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Height > 500 Then Vp.Height = Me.Height - 500
    Vp.Top = 1
    Vp.Left = 1
    Vp.Width = Me.Width - 200
    Err.Clear
End Sub

Private Sub pCompararUltimaFila()
    With Vp
        'If xFila >= 15500 Then'Hoja a4
        If xFila >= 14700 Then
            '.DrawLine 700, 15900, 11000, 15900
            .DrawLine 700, 15100, 11000, 15100
            .TextAlign = taLeftTop
            '--PONER DATOS AL FINAL DE HOJA
            .TextAlign = taLeftTop
    '''                    .TextBox "VAN ==>", 8000, 15850, 800, "200"
    '''                    .TextAlign = taRightTop
    '''                    .TextBox CStr(Format(ARR_ACUMULA(0,0), FORMAT_MONTO)), 8800, 15850, 950, "200" '--unidades entrada
    '''                    .TextBox CStr(Format(ARR_ACUMULA(0,1), FORMAT_MONTO)), 10000, 15850, 950, "200" '--unidades salida
    
    '''                    .TextBox CStr(Format(ARR_ACUMULA(0,2), FORMAT_MONTO)), 10000, 15850, 950, "200" '--importes salida
    '''                    .TextBox CStr(Format(ARR_ACUMULA(0,3), FORMAT_MONTO)), 10000, 15850, 950, "200" '--importes salida
            
            
            .NewPage
            Cabecera
            '--PONER DATOS AL INICIO DE HOJA
            .TextAlign = taLeftTop
    '''                    .TextBox "VIENEN ==>", 8000, 2200, 800, "200"
    '''                    .TextAlign = taRightTop
    '''                    .TextBox CStr(Format(ARR_ACUMULA(0,0), FORMAT_MONTO)), 8800, 15850, 950, "200" '--unidades entrada
    '''                    .TextBox CStr(Format(ARR_ACUMULA(0,1), FORMAT_MONTO)), 10000, 15850, 950, "200" '--unidades salida
    
    '''                    .TextBox CStr(Format(ARR_ACUMULA(0,2), FORMAT_MONTO)), 10000, 15850, 950, "200" '--importes salida
    '''                    .TextBox CStr(Format(ARR_ACUMULA(0,3), FORMAT_MONTO)), 10000, 15850, 950, "200" '--importes salida
            
            .FontSize = 7
            xFila = 2150
        Else
            xFila = xFila + 190
        End If
    End With
End Sub
