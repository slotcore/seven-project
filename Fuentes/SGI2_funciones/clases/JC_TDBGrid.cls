VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JC_TDBGrid"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Enum OrdenTDBGRid
     AscendenteGrid = 0
     DescendenteGrid
End Enum

Public Sub TDB_SelDesActCheck(TDGRID As Object, Rst As ADODB.Recordset, Campo As String, valor As String)
    Dim selrow As TrueOleDBGrid70.SelBookmarks
    Set selrow = TDGRID.SelBookmarks
    TDGRID.Visible = False
    For i = 0 To selrow.Count - 1
        Rst.Bookmark = selrow.Item(i)
        Rst.Fields(Campo) = valor
    
    Next
    TDGRID.Visible = True
End Sub

Public Sub TDB_TodosDesActCheck(TDGRID As Object, RstX As ADODB.Recordset, Campo As String, valor As Integer)
    If RstX.State = 0 Then Exit Sub
On Error GoTo ERROR
    TDGRID.Visible = False
    
    If RstX.RecordCount > 0 Then
       Bookmarkx = RstX.Bookmark
       RstX.MoveFirst
       While Not RstX.EOF
             RstX.Fields(Campo) = valor
             RstX.MoveNext
       Wend
       If RstX.RecordCount > 0 Then RstX.Bookmark = Bookmarkx
    End If
    TDGRID.Visible = True
    Exit Sub
ERROR:
    TDGRID.Visible = True
    MsgBox Err.Description & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear
End Sub

'
Public Function TDB_Ordenar(TDGRID As Object, ColIndex, Rst As ADODB.Recordset, Orden As OrdenTDBGRid, Optional Intercambio As Boolean = False, Optional OrdenDefecto As String = "Asc", Optional QuedaCol As Boolean = True) As String
    Dim NomCampo As String
    Dim TituloColumna As String
    'Dim OrdenX As String
    On Error Resume Next
    
    If IsNumeric(ColIndex) = True Then
        If ColIndex > TDGRID.Columns.Count - 1 Then
            TDB_Ordenar = ""
            Exit Function
        End If
    End If
    
    NomCampo = TDGRID.Columns(ColIndex).DataField
    TituloColumna = TDGRID.Columns(ColIndex).Caption
    
    
    If Intercambio = True Then
       If Rst.Sort = "" Then
          Rst.Sort = "[" & NomCampo & "] " & OrdenDefecto
       Else
          If Rst.Sort = "[" & NomCampo & "]" & " asc" Then
             Rst.Sort = "[" & NomCampo & "]" & " desc"
          Else
             Rst.Sort = "[" & NomCampo & "]" & " asc"
          End If
       End If
    Else
       If Orden = AscendenteGrid Then
          Rst.Sort = "[" & NomCampo & "] asc"
       ElseIf Orden = DescendenteGrid Then
          Rst.Sort = "[" & NomCampo & "] desc"
       Else
          TDB_Ordenar = ""
          Exit Function
       End If
    End If
    
    If QuedaCol = True Then
       If IsNumeric(ColIndex) Then
          TDGRID.Col = ColIndex
       Else
          TDGRID.Col = TDGRID.Columns(ColIndex).ColIndex
       End If
    End If
    
    TDB_Ordenar = TituloColumna
    Err.Clear
End Function
'
Public Function TDB_ClickEnCabeza(TDGRID As Object, ColIndex, X As Single, Y As Single) As Boolean
    If X >= TDGRID.Columns(ColIndex).Left And TDGRID.Columns(ColIndex).Left <= X Then
       If Y >= 0 And Y <= 200 Then
            ClickEnCabezaTDBGrid = True
            Exit Function
        End If
    End If
    TDB_ClickEnCabeza = False
End Function

''--TDGRID =  TDBGrid
'
Public Sub TDB_FiltroLimpiar(TDGRID As Object)
    Dim k As Integer
    For k = 0 To TDGRID.Columns.Count - 1
        TDGRID.Columns(k).FilterText = ""
    Next
End Sub

Public Sub TDB_IMPRIMIR(TDGRID As Object, _
                        N_TITULO_FRM As String, _
                        N_TITULO As String, _
                        Optional CON_NUMERO_PAG As Boolean = True)

    On Error GoTo ERROR
    
    Dim x_HeadForeColor As Variant
    Dim x_BackColor As Variant

    x_HeadForeColor = TDGRID.HeadForeColor
    x_BackColor = TDGRID.Style.BackColor
    '---------------
    TDGRID.HeadForeColor = vbBlack
    TDGRID.Style.BackColor = vbWhite
    TDGRID.HeadLines = 1.5 '--ALTO DE ENCABEZADO
    With TDGRID.PrintInfo
        '.Default = False
        .PreviewCaption = N_TITULO_FRM '--TITULO DE FRM
        .PreviewInitZoom = 100 '--ZOOM

        .PageHeaderFont.Italic = True
        .PageHeaderHeight = 500
        .PageHeader = N_TITULO
        
        ' Column headers will be on every page
        .RepeatColumnHeaders = True
        '--DIMENSIONES DE FRM
        .PreviewInitScreenFill = 50
        .PreviewInitHeight = 3500
        .PreviewInitWidth = 5000
        
        '--NUMERO DE PAGINAS
        If CON_NUMERO_PAG = True Then .PageFooter = "Pág: \p de \P "
        
        .PrintPreview

    End With
    '--VOLVIENDO AL ESTADO NATURAL
    TDGRID.HeadForeColor = x_HeadForeColor
    TDGRID.Style.BackColor = x_BackColor
    
    Exit Sub
ERROR:
    'SHOW_ERROR
End Sub
'------
'
Private Sub Class_Initialize()

End Sub



Public Function TDB_FiltroGenerar(TDGRID As Object, Rst As ADODB.Recordset)
    Dim tmp As String
    Dim n As Integer
    Dim k As Integer
    Dim C As Integer
 
    If Rst.State = 0 Then Exit Function
    Err.Clear
    On Error GoTo errHandler
    C = TDGRID.Col
    TDGRID.HoldFields
    tmp = ""
    For k = 0 To TDGRID.Columns.Count - 1
        If Trim(TDGRID.Columns(k).FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            tmp = tmp & TDGRID.Columns(k).DataField & " LIKE '%" & TDGRID.Columns(k).FilterText & "%'"
        End If
    Next

    Rst.Filter = tmp
    TDGRID.Col = C
    TDGRID.EditActive = True

    Err.Clear

    Exit Function

errHandler:
    Rst.Filter = ""
    TDB_FiltroLimpiar TDGRID
    If Err.Number = -2147217825 Then
        MsgBox "La columna no se puede filtrar", vbExclamation
    Else
        MsgBox Err.Source & ":" & vbCrLf & Err.Description
    End If

    Err.Clear

End Function

Public Sub TDB_Actualizar(xfrm As Object, xTabOne As Object, TDGRID As Object, xRst As ADODB.Recordset)
    '===================================================================================================
    'Creado : 06/11/11 Por: Johan Castro
    'Propósito: Limpiar el filtro que tenga el Objeto TDGRID asimismo refrescar el recordset para mostrar
    '           nuevos registros si los hubiera
    '
    'Entradas:  xFrm = Formulario de trabajo
    '           xTabOne = Control TabOne consulta y detalle, en este caso interesa pestaña consulta
    '           TDGRID = Objeto TrueDBGrid
    '           xRst = Recordset principal
    '
    'Resultados: Limpiar filtro y consula actualizada.
    '
    '===================================================================================================

    xfrm.MousePointer = vbHourglass
    
    '--Poniendo enfoque en pestaña consulta
    xTabOne.CurrTab = 0
    
    '--limpiar los filtros
    xRst.Filter = ""
    
    TDB_FiltroLimpiar TDGRID
    
    '--Refrescando el listado
    If xRst.State = 1 Then xRst.Requery
    TDGRID.Refresh
    xfrm.MousePointer = vbDefault

End Sub



