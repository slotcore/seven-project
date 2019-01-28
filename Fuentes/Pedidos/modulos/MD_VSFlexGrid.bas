Attribute VB_Name = "MD_VSFlexGrid"
Public Const FORMAT_MONTO As String = "#,###,##0.00"
Public Const FORMAT_DATE As String = "dd/mm/yy"
Public Const FORMAT_IMPUESTO As String = "#####0.000"
Public Const FORMAT_CANTIDAD As String = "#####0.00"

Public Enum e_ESTADO_ROW_GRID
'--SE USA ESTOS VALORES CON LA FINALIDAD DE SER UNICOS
    Fila_Ninguno = -999
    Fila_grupo = -991
    Fila_en_Blanco = -992
    Fila_Total = -993
    Fila_Total_grl = -994
End Enum


Sub UNIR_CELDAS(GRID As VSFlexGrid, X_ROW1 As Integer, X_COL1 As Integer, X_ROW2 As Integer, X_COL2 As Integer, X_CAPTION As String, Optional X_ALINEACION As AlignmentSettings = flexAlignCenterCenter)
    'On Error Resume Next
    With GRID
        '0
        .MergeCells = flexMergeRestrictAll
        .Row = X_ROW2
        .MergeRow(X_ROW2) = True
        .Select X_ROW1, X_COL1, X_ROW2, X_COL2
        .CellAlignment = X_ALINEACION
        .Cell(flexcpText, X_ROW1, X_COL1, X_ROW2, X_COL2) = X_CAPTION
        
        
    End With
End Sub


Sub OCULTAR_COL(GRID As VSFlexGrid, COL_INI As Integer, COL_FIN As Integer)
    '--OCULTA LAS COLUMNAS DE ACUERDO AL INCIO Y FIN
    For k = COL_INI To COL_FIN
        GRID.ColWidth(k) = 0
    Next
End Sub

Sub FORMATO_CELDA(GRID As VSFlexGrid, X_ROW1 As Integer, X_COL1 As Integer)
    '--DAR LA FUENTE A LA CELDA
    GRID.Row = X_ROW1: GRID.Col = X_COL1
    GRID.CellFontBold = False
    GRID.CellForeColor = &H800000
End Sub


Sub LimpiarGrid(GRID As VSFlexGrid, Optional CON_SELECCION_REGISTROS As Boolean = False)
    GRID.Clear
    GRID.Rows = 2
    '    grid.FormatString = vFormatString
    If CON_SELECCION_REGISTROS = True Then
        '--EL FORMATO SE GUARDARA AL MOMENTO DE CARGAR EL FORM.
        GRID.FormatString = GRID.Tag
    End If

End Sub

Function VERIFICAR_LISTA(GRID As VSFlexGrid, M_COL As Integer, N_VALOR As String) As Boolean
    '--ESTA FUNCION VALIDA QUE SOLO SE AGREGUE UN REGISTRO A LA LISTA
    With GRID
        For k = 0 To .Rows - 1
            If k + 1 = .Rows Then Exit For
            If CStr(.TextMatrix(k + 1, M_COL)) = N_VALOR Then
                MsgBox "Ya Existe in registro con el mismo nombre", vbInformation, "Mensaje"
                Exit Function
            End If
        Next k
    End With
    VERIFICAR_LISTA = True
End Function

Sub ADD_REG(GRID As VSFlexGrid, Optional M_ESTADO_FILA As e_ESTADO_ROW_GRID = Fila_Ninguno)
    '--AGREGAR UNA FILA A LA GRILLA
    'M_ESTADO_FILA
    GRID.AddItem ""
    If M_ESTADO_FILA <> Fila_Ninguno Then GRID.TextMatrix(GRID.Rows - 1, 1) = M_ESTADO_FILA
End Sub

