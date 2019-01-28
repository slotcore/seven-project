VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JC_VSFlexGrid"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--GRID = VSFlexGrid

Public Enum e_ESTADO_ROW_GRID
    '--SE USA ESTOS VALORES CON LA FINALIDAD DE SER UNICOS
    Fila_Ninguno = -999
    Fila_grupo = -991
    Fila_en_Blanco = -992
    Fila_Total = -993
    Fila_Total_grl = -994
End Enum

Public Enum e_COLOR
    e_NEGRO = vbBlack
    e_AZULMARINO = &H800000
    e_ROJO = vbRed
    e_VERDE = &HC000&
    e_CELESTE = 16755972
End Enum

Public Enum e_RowColVSFlexGrid
    e_Fila = 1
    e_Columna = 2
End Enum

Public Enum e_TipoFrmFecha
    e_Escribir = 1
    e_Seleccion = 2
End Enum

Public FORMAT_MONTO As String
Public FORMAT_DATE As String
Public FORMAT_IMPUESTO As String
Public FORMAT_CANTIDAD As String
Public FORMAT_MEDIA As String
Public FORMAT_PU As String
Public FORMAT_HORA_SIN_SEGUNDO As String
Public FORMAT_HORA_AL_SEGUNDO As String
Public FORMAT_HORA_LARGO As String
Public FORMAT_PORCENTAJE As String
Public INICIO_GRUPO As String

Private Sub Class_Initialize()
    
    FORMAT_MONTO = "###,###,##0.00"
    FORMAT_DATE = "dd/mm/yy"
    FORMAT_IMPUESTO = "#####0.000"
    FORMAT_CANTIDAD = "###,###,##0.00"
    FORMAT_MEDIA = "###,##0.0000"
    FORMAT_PU = "###,##0.000000"
    FORMAT_HORA_SIN_SEGUNDO = "hh:mm AM/PM"
    FORMAT_HORA_AL_SEGUNDO = "hh:mm:ss AM/PM"
    FORMAT_HORA_LARGO = "HH:MM:SS"
    INICIO_GRUPO = "> "
    FORMAT_PORCENTAJE = "##0.00"
    '--------------------------------------------------------------------------
    Nomsis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")

    
End Sub


Function UNIR_CELDAS(GRID As Object, _
                    X_ROW1 As Long, X_COL1 As Integer, _
                    X_ROW2 As Long, X_COL2 As Integer, _
                    x_CAPTION As String, _
                    Optional X_ALINEACION As AlignmentSettings = flexAlignCenterCenter, _
                    Optional F_MERGE_ROW As Boolean = True, _
                    Optional X_MERGECELL As MergeSettings = flexMergeFree)
    
   On Error Resume Next
    With GRID
        '0
        .MergeCells = X_MERGECELL
        .Row = X_ROW2
        If F_MERGE_ROW = True Then
            .MergeRow(X_ROW2) = True
        Else
            .MergeCol(X_COL2) = True
        End If
        .Select X_ROW1, X_COL1, X_ROW2, X_COL2
        .CellAlignment = X_ALINEACION
        .Cell(flexcpText, X_ROW1, X_COL1, X_ROW2, X_COL2) = x_CAPTION
       
    End With
    Err.Clear
End Function


Sub OCULTAR_COL(GRID As Object, COL_INI As Integer, COL_FIN As Integer)
    '--OCULTA LAS COLUMNAS DE ACUERDO AL INCIO Y FIN
    For k = COL_INI To COL_FIN
        GRID.ColWidth(k) = 0
    Next
End Sub

Sub FORMATO_CELDA(GRID As Object, X_ROW1 As Long, X_COL1 As Integer, _
                   Optional x_ForeColor As Long = &H800000, _
                   Optional ES_NEGRITA As Boolean = False, _
                   Optional x_BackColor As Long = &HFFFFFF, _
                   Optional x_CAPTION As String = "")
                   '--::x_ForeColor=negro
                   '--::x_BackColor=blanco
    '--DAR LA FUENTE A LA CELDA
    GRID.Row = X_ROW1: GRID.Col = X_COL1
    If x_CAPTION <> "" Then GRID.TextMatrix(X_ROW1, X_COL1) = x_CAPTION
    GRID.CellFontBold = ES_NEGRITA
    GRID.CellForeColor = x_ForeColor
    GRID.CellBackColor = x_BackColor
End Sub


Sub LimpiarGrid(GRID As Object, Optional CON_FORMATSTRING As Boolean = False, Optional mRowS As Integer = 2)
    GRID.Clear
    GRID.Rows = mRowS
    If CON_FORMATSTRING = True Then
        '--EL FORMATO SE GUARDARA AL MOMENTO DE CARGAR EL FORM.
        GRID.FormatString = GRID.Tag
    End If

End Sub

Function VERIFICAR_LISTA(GRID As Object, M_COL As Integer, N_VALOR As String, Optional SHOW_MSG As Boolean = True) As Boolean
    '--ESTA FUNCION VALIDA QUE SOLO SE AGREGUE UN REGISTRO A LA LISTA
    Dim k As Long
    With GRID
        For k = 0 To .Rows - 1
            If k + 1 = .Rows Then Exit For
            If CStr(.TextMatrix(k + 1, M_COL)) = N_VALOR Then
                If SHOW_MSG = True Then MsgBox "Ya Existe un registro con el mismo Valor", vbInformation, "Mensaje"
                Exit Function
            End If
        Next k
    End With
    VERIFICAR_LISTA = True
End Function

Sub ADD_REG(GRID As Object, Optional M_ESTADO_FILA As e_ESTADO_ROW_GRID = Fila_Ninguno)
    '--AGREGAR UNA FILA A LA GRILLA
    'M_ESTADO_FILA
    GRID.AddItem ""
    If M_ESTADO_FILA <> Fila_Ninguno Then GRID.TextMatrix(GRID.Rows - 1, 1) = M_ESTADO_FILA
End Sub

Sub GRID_COMBOLIST(GRID As Object, _
                    Optional Col As Integer = 2, _
                    Optional J_ColFormat As String = "|...", _
                    Optional J_EditableSettings As EditableSettings = flexEDKbdMouse, _
                    Optional J_SelModeSettings As SelModeSettings = flexSelectionFree)
                    
    GRID.ColComboList(Col) = J_ColFormat
    GRID.Editable = J_EditableSettings
    GRID.SelectionMode = J_SelModeSettings
End Sub

Function GENERAR_SQL_ID(GRID As Object, mCol As Integer, nCampo As String, _
            Optional nTipoIn As String = "IN", Optional fEsNumero As Boolean = True) As String
            
    Dim k&
    Dim nSQL As String
    Dim Apostrofe As String
    If fEsNumero = False Then Apostrofe = "'"
    If fEsNumero = True Then Apostrofe = ""
    nSQL = ""
    
    With GRID
        For k = .FixedRows To .Rows - 1
            If CStr(.TextMatrix(k, mCol)) <> "" Then
                nSQL = nSQL + Apostrofe + CStr(.TextMatrix(k, mCol)) + Apostrofe + ","
            End If
        Next k
    End With
    If nSQL <> "" Then nSQL = " " + nCampo + " " + nTipoIn + " (" + Left(nSQL, Len(nSQL) - 1) + ") "
    GENERAR_SQL_ID = nSQL
End Function

'Public Function GRID_SUMAR_COL(GRID As Object, mCol As Integer, Optional RowInicial As Long = 0, Optional RowFinal As Long = 0) As Double
'    '===================================================================================================
'    'Creado : 10/08/08 Por: Johan Castro
'    'Propósito: Sumar una columna de un grid
'    '
'    'Entradas:  GRID = Control VSFlexGrid
'    '           mCol= numero de columna que sera para sumar
'    '           RowInicial = Fila Inicial; Indicara el inicio del fila para sumar (Opcional)
'    '           RowFinal = Fila Final; Indicara el fin de la fila para sumar(Opcional)
'    '
'    'Resultados:Valor que indica la suma segun parametros ingresados
'
'    'Modificado : 05/09/09 Por: Johan Castro
'    '           Considera en suma rango de filas RowInicial,RowFinal
'    '
'    '===================================================================================================
'    '--
'
'    '--ACUMULA DE ACUERDO A UNA COLUMNA
'    Dim mRow&
'    Dim sTotal As Double
'    Dim sRowInicio As Long
'    Dim sRowFinal As Long
'
'    sTotal = 0#
'    If RowInicial = 0 Or RowInicial <= GRID.FixedRows Then
'        sRowInicio = GRID.FixedRows
'    Else
'        sRowInicio = RowInicial
'    End If
'
'    If sRowFinal = 0 Or sRowFinal >= GRID.Rows - 1 Then
'        sRowFinal = GRID.Rows - 1
'    Else
'        sRowFinal = RowFinal
'    End If
'
'    For mRow = sRowInicio To sRowFinal
'        If IsNumeric(GRID.TextMatrix(mRow, mCol)) = True And GRID.RowHidden(mRow) = False Then
'            sTotal = sTotal + CDbl(GRID.TextMatrix(mRow, mCol))
'        End If
'    Next mRow
'
'    GRID_SUMAR_COL = sTotal
'
'End Function

Public Function GRID_SUMAR_COL(GRID As Object, mCol As Integer, Optional RowInicial As Long = 0, Optional RowFinal As Long = 0) As Double
    '===================================================================================================
    'Creado : 10/08/08 Por: Johan Castro
    'Propósito: Sumar una columna de un grid
    '
    'Entradas:  GRID = Control VSFlexGrid
    '           mCol= numero de columna que sera para sumar
    '           RowInicial = Fila Inicial; Indicara el inicio del fila para sumar (Opcional)
    '           RowFinal = Fila Final; Indicara el fin de la fila para sumar(Opcional)
    '
    'Resultados:Valor que indica la suma segun parametros ingresados
    
    'Modificado : 05/09/09 Por: Johan Castro
    '           Considera en suma rango de filas RowInicial,RowFinal
    'Modificado : 26/05/10 Por: Johan Castro
    '           Correccion de sumatoria utilizando parametro RowInicial
    '===================================================================================================
    '--
    
    '--ACUMULA DE ACUERDO A UNA COLUMNA
    Dim mRow&
    Dim sTotal As Double
    Dim sRowInicio As Long
    Dim sRowFinal As Long
    
    sTotal = 0#
    If RowInicial = 0 Or RowInicial <= GRID.FixedRows Then
        sRowInicio = GRID.FixedRows
    Else
        sRowInicio = RowInicial
    End If
    
    If RowFinal = 0 Or RowFinal >= GRID.Rows - 1 Then
        sRowFinal = GRID.Rows - 1
    Else
        sRowFinal = RowFinal
    End If
    
    For mRow = sRowInicio To sRowFinal
        If IsNumeric(GRID.TextMatrix(mRow, mCol)) = True And GRID.RowHidden(mRow) = False Then
            sTotal = sTotal + NulosN(GRID.TextMatrix(mRow, mCol))
        End If
    Next mRow
    
    GRID_SUMAR_COL = sTotal
        
End Function

Public Sub GRID_COLOR_FONDO(GRID As Object, _
                    X_ROW1 As Long, X_COL1 As Integer, _
                    X_ROW2 As Long, X_COL2 As Integer, _
                    Optional x_BackColor As Long = &H8000000E, Optional x_FillStyle As FillStyleSettings = flexFillRepeat)
                    
        With GRID
            If .FixedRows = .Rows Then Exit Sub
            .Select X_ROW1, X_COL1, X_ROW2, X_COL2
            .FillStyle = x_FillStyle
            .CellBackColor = x_BackColor
            .Select 1, 1, 1, 1
        End With

End Sub


Function GRID_BUSCAR_VALOR(GRID As Object, _
                M_COL As Integer, _
                N_VALOR As String, _
                Optional SHOW_MSG As Boolean = True, _
                Optional COL_DEVOLVER_VALOR As Long = -1, _
                Optional ROW_NO_CONISDERA As Long = -1) As String
    '--ESTA FUNCION BUSCA UN VALOR EN LA GRILLA SEGUN COLUMNA,
    '--COMPARA EL VALOR EN TODA LA COLUMNA, VALIDA QUE COL SEA DIFERENTE A ROW_NO_CONISDERA
    '--DEVUELVE EL VALOR SEGUN COLUMNA=COL_DEVOLVER_VALOR
    Dim k&
    With GRID
        For k = .FixedRows To .Rows - 1
            If CStr(.TextMatrix(k, M_COL)) = N_VALOR And (k <> ROW_NO_CONISDERA) Then
                If SHOW_MSG = True Then MsgBox "Ya Existe un registro con el mismo Valor", vbInformation, "Mensaje"
                If COL_DEVOLVER_VALOR = -1 Then GRID_BUSCAR_VALOR = CStr(.TextMatrix(k, M_COL))
                If COL_DEVOLVER_VALOR <> -1 Then GRID_BUSCAR_VALOR = CStr(.TextMatrix(k, COL_DEVOLVER_VALOR))
                Exit Function
            End If
        Next k
    End With
    GRID_BUSCAR_VALOR = "-1"
End Function


Function GRID_GENERAR_SQL_ID(GRID As Object, _
                            mCol As Integer, _
                            nCampo As String, _
                            Optional nTipoIn As String = "IN", _
                            Optional fEsNumero As Boolean = True, _
                            Optional mColConsidera As Integer = -1, _
                            Optional nValorConsidera As String = "", _
                            Optional nCondicional As String = "") As String
    '===================================================================================================
    'Creado : 15/09/08 Por: Johan Castro
    'Propósito: Generar filtro de varios registros en sentencia SQL
    '
    'Entradas:  GRID = Control VSFlexGrid
    '           mCol= Número de columna que sera para sumar
    '           nCampo = Nombre del campo a filtrar(puede ir acompañado de la tabla.campo)
    '           nTipoIn = Indica si va incluir o excluir los datos de la consulta; Puede ser "IN" o "NOT IN"
    '           fEsNumero = Indica si el dato es numerico ; True=Numerico, False=No es numero
    '           mColConsidera = Indica un filtro adicional del objeto; -1= No se considera este filtro
    '           nValorConsidera = Indica el valor que se debe filtrar en la columna = mColConsidera
    '           nCondicional= Indica si se agrega al final de la sentencia SQL el condicional AND u OR
    '
    'Resultados:Parte de la sentencia SQL que se agregara en WHERE o HAVING
    
    'Modificado : 21/09/10 Por: Johan Castro
    '           Agregar parametro nCondicional
    
    '===================================================================================================
    '--
                            
    Dim nSQL As String
    Dim k&
    nSQL = ""
    
    
    With GRID
        If .Cols - 1 < mCol Then Exit Function
        For k = .FixedRows To .Rows - 1
            If Trim(CStr(.TextMatrix(k, mCol))) <> "" Then
                If mColConsidera <> -1 Then
                    If fEsNumero = True Then
                        If NulosN(.TextMatrix(k, mColConsidera)) = NulosN(Trim(nValorConsidera)) Then
                            nSQL = nSQL & NulosN(.TextMatrix(k, mCol)) & ","
                        End If
                    Else
                        If NulosC(.TextMatrix(k, mColConsidera)) = NulosC(Trim(nValorConsidera)) Then
                            nSQL = nSQL & "'" & NulosC(.TextMatrix(k, mCol)) & "',"
                        End If
                    End If
                Else
                    If fEsNumero = True Then
                        nSQL = nSQL & NulosN(.TextMatrix(k, mCol)) & ","
                    Else
                        nSQL = nSQL & "'" & NulosC(.TextMatrix(k, mCol)) & "',"
                    End If
                End If
            End If
        Next k
    End With
    
    If nSQL <> "" Then nSQL = " " & nCampo & " " & nTipoIn & " (" + Left(nSQL, Len(nSQL) - 1) & ") " & nCondicional & " "
        
    GRID_GENERAR_SQL_ID = nSQL
    
End Function

Public Sub GRID_INSERT(GRID As Object, mPosOrigen As Long, Optional eTipo As e_RowColVSFlexGrid = e_Fila)
    If eTipo = e_Columna Then
        If mPosOrigen > GRID.Cols Then Exit Sub
        GRID.Cols = GRID.Cols + 1
        GRID.ColPosition(GRID.Cols - 1) = mPosOrigen
    Else
        If mPosOrigen > GRID.Rows Then Exit Sub
        GRID.Rows = GRID.Rows + 1
        GRID.RowPosition(GRID.Rows - 1) = mPosOrigen
    End If
End Sub

Public Sub GRID_DELETE(GRID As Object, mIni As Long, mFin As Long, _
                Optional eTipo As e_RowColVSFlexGrid = e_Fila)
                
    On Error Resume Next
    Dim mRow&
    If eTipo = e_Fila Then
        If GRID.Rows - GRID.FixedRows = 0 Then Exit Sub
        
        For mRow = mIni To mFin
            GRID.RemoveItem GRID.Rows - 1
        Next mRow
    Else
        If GRID.Cols = 1 Then Exit Sub
        
        For mRow = mIni To mFin
            '--coloco la columna en la ultima posicion
            GRID.ColPosition(mIni) = GRID.Cols - 1
            '--elimino la ultima fila
            GRID.Cols = GRID.Cols - 1
        Next mRow
        
    End If
    Err.Clear
End Sub

Public Sub GRID_AGRUPAR(GRID As Object, mColCompara As Long, _
                Optional cColor1 As Long = &HFDFFFF, _
                Optional cColor2 As Long = &HE0FEFE)
                
    '--esta procedimiento pinta las filas agrupandolos segun el valor de la una columna
    Dim mRow&
    Dim nCompara As String
    Dim cColor As Long
    Dim fCambia As Boolean
    '--validar que la columna a comparar sea menor o igual al total de columnas
    If GRID.Cols - mColCompara <= 0 Then Exit Sub
    '--restableciendo como fondo blanco en el grid
    If GRID.Rows - GRID.FixedRows = 0 Then Exit Sub
    GRID_COLOR_FONDO GRID, GRID.FixedRows, 1, GRID.Rows - 1, GRID.Cols - 1, vbWhite
    '--
    If GRID.Rows <> 0 Then
        nCompara = NulosC(GRID.TextMatrix(1, mColCompara))
        cColor = cColor2
        fCambia = False
    End If
    For mRow = GRID.FixedRows To GRID.Rows - 1
        DoEvents
        '--agrupando por colores
        If nCompara <> NulosC(GRID.TextMatrix(mRow, mColCompara)) Then
            nCompara = NulosC(GRID.TextMatrix(mRow, mColCompara))
            If fCambia = False Then
                cColor = cColor1
            Else
                cColor = cColor2
            End If
            fCambia = Not fCambia
        End If
        GRID_COLOR_FONDO GRID, mRow, 1, mRow, GRID.Cols - 1, cColor
        '--
    Next mRow
End Sub


Public Sub GRID_ORDENAR(GRID As Object, _
                    FILA_INICIO As Long, COL_ORDEN As Long, _
                    Optional FILA_INICIO1 As Long, Optional COL_ORDEN1 As Long, _
                    Optional TIPO_ORDEN As SortSettings = flexSortGenericAscending)
    If GRID.FixedRows >= GRID.Rows Then Exit Sub
    If FILA_INICIO1 = 0 Then
        GRID.Select FILA_INICIO, COL_ORDEN
    Else
        GRID.Select FILA_INICIO, COL_ORDEN, FILA_INICIO1, COL_ORDEN1
    End If
    GRID.Sort = TIPO_ORDEN

End Sub

Function GRID_COMBINAR(GRID As Object, _
                    X_ROW1 As Long, X_COL1 As Integer, _
                    X_ROW2 As Long, X_COL2 As Integer, _
                    x_CAPTION As String, _
                    Optional X_ALINEACION As AlignmentSettings = flexAlignCenterCenter, _
                    Optional F_MERGE_ROW As Boolean = True, _
                    Optional X_MERGECELL As MergeSettings = flexMergeFree, _
                    Optional x_ForeColor As Long = &H800000, _
                    Optional x_BackColor As Long = &HFFFFFF, _
                    Optional ES_NEGRITA As Boolean = False)

   On Error Resume Next
    With GRID
        '0
        .MergeCells = X_MERGECELL
        .Row = X_ROW2
        If F_MERGE_ROW = True Then
            .MergeRow(X_ROW2) = True
        Else
            .MergeCol(X_COL2) = True
        End If
        .Select X_ROW1, X_COL1, X_ROW2, X_COL2
        .CellAlignment = X_ALINEACION
        .Cell(flexcpText, X_ROW1, X_COL1, X_ROW2, X_COL2) = x_CAPTION
       
        .Row = X_ROW1: GRID.Col = X_COL1
        .CellFontBold = ES_NEGRITA
        .CellForeColor = x_ForeColor
        .CellBackColor = x_BackColor
       
    End With
    Err.Clear
End Function


Public Sub GRID_EXPORTAR_MSEXCELTMP(GRID As Object, Con As ADODB.Connection, _
                    Optional TipoCarga As SaveLoadSettings = flexFileCustomText, _
                    Optional AbrirArchivo As Boolean = False, _
                    Optional nTitulo As String, _
                    Optional nPeriodo As String = "", _
                    Optional nTitulo1 As String = "")
                    
                    
    '--01/08/08
    '--310309
   
    If GRID.Rows - GRID.FixedRows = 0 Then
        MsgBox "No hay Registros Para Exportar", vbInformation, xTitulo
        Exit Sub
    End If
    
    '******************************************************************
    
    'Dim objCommDlg As Object
'    Dim fs As Object
    
    'Set fs = CreateObject("Scripting.FileSystemObject")
    Dim fs As New Scripting.FileSystemObject
    'Dim objCommDlg  As New CommonDialog
    'Dim objCommDlg As CommonDialog
    'Set objCommDlg = CreateObject("MSComDlg.CommonDialog")
    
    'objCommDlg.Filter = "Libro de Microsoft Excel (*.xls|*.xls)"
    'objCommDlg.DialogTitle = "Guardar Como"
    'objCommDlg.ShowSave
    'If Trim(objCommDlg.FileName) = "" Then
        
        'Set objCommDlg = Nothing
        'Exit Sub

    'Else
    Dim xFile As String
    
    xFile = InputBox("Ingrese el nombre del archivo", Nomsis)
    If xFile = "" Then Exit Sub
    
    xFile = App.Path & "\" & xFile & ".xls"
    
    
    
        'If fs.FileExists(objCommDlg.FileName) Then
        If fs.FileExists(Trim(xFile)) Then
            If MsgBox("El archivo ya existe, ¿Desea Sobreescribirlo?", vbYesNo + vbInformation, "Fichero existente") = vbNo Then
                Set fs = Nothing
                'Set objCommDlg = Nothing
                Exit Sub
            End If
        End If
    'End If
    'RutaArchivo = objCommDlg.FileName
    RutaArchivo = xFile
    Set fs = Nothing
    'Set objCommDlg = Nothing
       
    GRID.SaveGrid RutaArchivo, TipoCarga
    
    If AbrirArchivo = True Then '--abrir el archivo
        Dim xFunc As New SGI2_funciones.JC_Varios
        Dim mCol&, mColInsert&, mRow&
        '--abriendo el archivo
        Dim objExcel As Object
        Set objExcel = CreateObject("Excel.Application")
        objExcel.Workbooks.Open RutaArchivo
        objExcel.Visible = False
        objExcel.WindowState = 1
        '----
        Set xCon = Con
        CargaDatosEmpresa
        '----
        mColInsert = 9 + GRID.FixedRows
        With objExcel.ActiveSheet
            .Rows("1:" & mColInsert).Select
            objExcel.Selection.Insert Shift:=xlDown
            '--colocando el ancho de las columnas
            For mCol = 1 To GRID.Cols - 1
                .Columns(mCol).ColumnWidth = GRID.ColWidth(mCol) / 100
            Next
            
            '--colocando el encabezado
            For mRow = 0 To GRID.FixedRows
                For mCol = 1 To GRID.Cols - 1
                    .Cells(mRow + 10, mCol) = GRID.TextMatrix(mRow, mCol)
                Next
            Next mRow
            Dim mColXls&
            mColXls = 1
            For mCol = 1 To GRID.Cols - 1
                If GRID.ColWidth(mCol) < 70 Then
                    .Columns(mColXls).Select
                    objExcel.Selection.Delete Shift:=xlToLeft
                    mColXls = mColXls - 1
                End If
                mColXls = mColXls + 1
            Next mCol
            
            '-----------------------------------------------------------------
            '--datos de la empresa
            .Cells(1, 1) = NomEmp
            'UNIR_CELDA objExcel, .Cells(1, 1).Address & ":" & .Cells(1, 4).Address:
            xFunc.xls_UnirCelda objExcel, .Cells(1, 1).Address & ":" & .Cells(1, 4).Address:
            .Cells(2, 1) = "R.U.C. : " + NumRUC
            xFunc.xls_UnirCelda objExcel, .Cells(2, 1).Address & ":" & .Cells(2, 4).Address:
            
            .Cells(3, 1) = Date
            xFunc.xls_UnirCelda objExcel, .Cells(3, 1).Address & ":" & .Cells(3, 4).Address:
            
            .Cells(1, 1).HorizontalAlignment = -4131
            .Cells(2, 1).HorizontalAlignment = -4131
            .Cells(3, 1).HorizontalAlignment = -4131
            '-----------------------------------------------------------------
            '--DEL TITULO
            .Cells(5, 1) = nTitulo:            .Cells(5, 1).Font.Bold = True
            .Cells(6, 1) = nPeriodo:           .Cells(6, 1).Font.Bold = True
            .Cells(7, 1) = nTitulo1:           .Cells(7, 1).Font.Bold = True
            '-----------------------------------------------------------------
        End With
        '-----
        Set xFunc = Nothing
''''        MsgBox "El archivo se exportó satisfactoriamente", vbInformation, xTitulo
        objExcel.Visible = True
        Set objExcel = Nothing
        Exit Sub
    End If
    
    MsgBox "El archivo se exportó satisfactoriamente", vbInformation, xTitulo
    objExcel.Visible = True
    
End Sub



Public Sub SetearCuadricula(GRID As Object, idFormato As Integer, xCon As ADODB.Connection, TipoFormato As Integer, Optional NumeroFormato As Integer, Optional OcultarCol As Boolean = True)
    'tipo de formato especifica los siguiente
    '1 = con_formatostipodet
    '2 = con_analisistipodet
    
    'NivelFormato = especifica el id de la tabla con_analisistipo
    
    'OcultarCol = especifica si las columnas se van a ocultar columnas
    
    Dim Rst As New ADODB.Recordset
    Dim A As Integer

    GRID.Rows = 2
    GRID.Cols = 1
    GRID.FixedRows = 2
    GRID.WordWrap = True
    GRID.RowHeight(0) = 400
    GRID.RowHeight(1) = 500

    Dim xPos1 As Integer
    Dim xNumCol As Integer
    Dim Cambio As Boolean
    Dim xCol As Integer
    
    Dim xFun As New SGI2_funciones.JC_VSFlexGrid
    
    If TipoFormato = 1 Then
        If NumeroFormato = 0 Then
            RST_Busq Rst, "SELECT con_formatostipodet.*, con_formatostipodet.orden FROM con_formatostipo LEFT JOIN con_formatostipodet " _
                & " ON (con_formatostipo.id = con_formatostipodet.idformatotipo) AND (con_formatostipo.idformato = con_formatostipodet.idformato) " _
                & " Where (((con_formatostipo.idformato) = " & idFormato & ") And ((con_formatostipo.defecto) = -1)) ORDER BY con_formatostipodet.orden", xCon
        Else
            RST_Busq Rst, "SELECT con_formatostipodet.*, con_formatostipodet.orden, con_formatostipo.id FROM con_formatostipo LEFT JOIN con_formatostipodet " _
                & " ON (con_formatostipo.id = con_formatostipodet.idformatotipo) AND (con_formatostipo.idformato = con_formatostipodet.idformato) " _
                & " Where (((con_formatostipo.id) = " & NumeroFormato & ") And ((con_formatostipo.idFormato) = " & idFormato & ")) " _
                & " ORDER BY con_formatostipodet.orden", xCon
        End If
    Else
        If NumeroFormato = 0 Then
            RST_Busq Rst, "SELECT con_analisistipodet.*, con_analisistipodet.orden FROM con_analisistipo LEFT JOIN con_analisistipodet " _
                & " ON (con_analisistipo.id = con_analisistipodet.idformatotipo) AND (con_analisistipo.idformato = con_analisistipodet.idformato) " _
                & " Where (((con_analisistipo.idformato) = " & idFormato & ") And ((con_analisistipo.defecto) = -1)) ORDER BY con_analisistipodet.orden", xCon
        Else
            RST_Busq Rst, "SELECT con_analisistipodet.*, con_analisistipodet.orden FROM con_analisistipo LEFT JOIN con_analisistipodet " _
                & " ON (con_analisistipo.idformato = con_analisistipodet.idformato) AND (con_analisistipo.id = con_analisistipodet.idformatotipo) " _
                & " Where (((con_analisistipo.idFormato) = " & idFormato & ") And ((con_analisistipo.id) = " & NumeroFormato & ")) ORDER BY con_analisistipodet.orden", xCon

        End If
    End If
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xCol = 1
        Do While Not Rst.EOF
            GRID.Cols = GRID.Cols + 1

            GRID.TextMatrix(0, GRID.Cols - 1) = NulosC(Rst("titulo"))
            GRID.TextMatrix(1, GRID.Cols - 1) = NulosC(Rst("abrev"))
            If Rst("mostrar") = -1 Then
                GRID.ColWidth(GRID.Cols - 1) = NulosN(Rst("ancho"))
            Else
                GRID.ColWidth(GRID.Cols - 1) = 0
            End If
        
            If NulosC(Rst("titulo")) = "" Then
                xFun.GRID_COMBINAR GRID, 0, GRID.Cols - 1, 1, GRID.Cols - 1, NulosC(Rst("abrev")), flexAlignCenterCenter, False, , &H80000008, &H8000000F
            End If

            GRID.FixedAlignment(GRID.Cols - 1) = flexAlignCenterCenter 'NulosN(Rst("alineacion"))
            GRID.ColAlignment(GRID.Cols - 1) = NulosN(Rst("alineacion"))
                                       
            '**********************************************************************************
            '--eliminar la columna que no se va usar
            If OcultarCol = False And NulosN(Rst("mostrar")) = 0 Then
                GRID.Cols = GRID.Cols - 1
            End If
            '**********************************************************************************
                        
            Rst.MoveNext
            xCol = xCol + 1
        Loop
    End If

    xFun.GRID_COMBINAR GRID, 0, xPos1, 0, xNumCol, GRID.TextMatrix(0, xPos1), flexAlignCenterCenter, , , &H80000008, &H8000000F
    
    GRID.MergeCells = flexMergeFixedOnly
    
    Set xFun = Nothing
End Sub





