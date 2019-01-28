Attribute VB_Name = "Declaraciones"
'*****************************************************************************************************
'* Nombre Archivo   : DECLARACIONES.BAS
'* Tipo             : MODULO
'* Descripcion      : MODULO DONDE SE DECLARAN LAS VARIABLES PUBLICAS QUE SE UTILIZARAN EN LA CLASE
'*                    ASI COMO LA DEFINICION DE ALGUNAS FUNCIONES PROPIAS DE LA CLASE
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 28/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public xCon As New ADODB.Connection
Public xMes As Integer
Public xTitulo As String

Public NomEmp As String
Public NumRUC As String
Public AnoTra As String
Public Nomsis As String
Public xIdUsuario As Integer

Public CONTABILIZAR As Boolean

Global AP_RUTASY As String
Global AP_RUTABD As String
Global AP_RUTABM As String

Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)

Public Const FORMAT_CANTIDADDECIMAL = "0.0000"

Public Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFacename As String * 33
End Type

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const FW_BOLD = 700    '''''
Public Const FW_NORMAL = 400  '''''
Public Function existeNumeroDoc(TABLA_ As String, CONDICION1 As String, CAMPO1 As String, _
                                            Optional CONDICION2 As String = "", _
                                            Optional CAMPO2 As String = "", _
                                            Optional CONDICION3 As String = "", _
                                            Optional CAMPO3 As String = "") As Boolean
    Dim xRs As New ADODB.Recordset
    Dim xNum As Double
    Dim cSQL As String
    Dim nSQL As String
    
    If CONDICION2 <> "" And CAMPO2 <> "" Then
        nSQL = " AND ((" & CAMPO2 & ") = " & CONDICION2 & ")"
    End If
    
    If nSQL <> "" And CONDICION3 <> "" And CAMPO3 <> "" Then
        nSQL = nSQL & " AND ((" & CAMPO3 & ") = " & CONDICION3 & ")"
    End If
    
    cSQL = "SELECT TOP 1 * " _
        + vbCr + "FROM " & TABLA_ & " " _
        + vbCr + "WHERE ((" & CAMPO1 & ") = " & CONDICION1 & ")" & nSQL
    
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then Exit Function
        
    If xRs.RecordCount > 0 Then
        existeNumeroDoc = True
    Else
        existeNumeroDoc = False
    End If
    Set xRs = Nothing
End Function

Public Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    With Rst
        If .State <> 0 Then
            If TODO Then .Filter = adFilterNone
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
            End If
        End If
    End With
End Sub

Public Function crearModificarLote(MODO_ As Integer, TIPO_ As Integer, _
                            IDITEM_ As Integer, CANTIDAD_ As Double, _
                            IDALM_ As Integer, FCHING_ As String, _
                            IDLOTE_ As Integer, IDLOTEDET_ As Integer, _
                            Optional CANANT_ As Double = 0, Optional IDLOTEANT_ As Integer = 0, _
                            Optional IDLOTEDETANT_ As Integer = 0) As Integer
                            
    ' MODO_=0:Nuevo MODO_=1:Modificar
    ' TIPO_=0:Ingreso TIPO_=1:Salida
    Dim RstLote As New ADODB.Recordset
    Dim RstLoteDet As New ADODB.Recordset
    Dim xIdLoteAnt As Double
    Dim xIdLoteDetAnt As Double
    Dim xIdLote As Double
    Dim xIdLoteDet As Double
    Dim LOTE_ As String
    Dim MAXBD_ As Integer
    Dim NUMERO_ As Integer
    Dim xRs As New ADODB.Recordset
    Dim cSQL As String

On Error GoTo LaCague

    RST_Busq RstLote, "SELECT * FROM alm_inventariolote", xCon
    RST_Busq RstLoteDet, "SELECT * FROM alm_inventariolotedet", xCon
    xIdLote = HallaCodigoTabla("alm_inventariolote", xCon, "id")
    xIdLoteDet = HallaCodigoTabla("alm_inventariolotedet", xCon, "id")
        
    xCon.BeginTrans
    If (MODO_ = 0) Then ' --------------------------------------------------------Nuevo registro
        Select Case TIPO_
            Case 0 ' Agregar
AGREGARNUEVO_:
                LOTE_ = Format(IDITEM_, "0000") & Format(CDate(FCHING_), "yy") & Format(Month(CDate(FCHING_)), "00") & Format(Day(CDate(FCHING_)), "00")
                    
                ' Se verifica el mayor en la base de datos
                cSQL = "SELECT Max(Mid([alm_inventariolote].[descripcion],11,2)) AS orden, alm_inventariolote.iditem, alm_inventariolote.fching " _
                    + vbCr + "FROM alm_inventariolote " _
                    + vbCr + "GROUP BY alm_inventariolote.iditem, alm_inventariolote.fching " _
                    + vbCr + "HAVING (((alm_inventariolote.iditem)=" & IDITEM_ & ") AND ((alm_inventariolote.fching)=CDate('" & FCHING_ & "')))"
                
                Set xRs = Nothing
                RST_Busq xRs, cSQL, xCon
                
                MAXBD_ = 0
                If xRs.State = 0 Then GoTo SALIR_
                If xRs.RecordCount = 0 Then GoTo SALIR_
                
                MAXBD_ = NulosN(xRs("orden"))
SALIR_:
                NUMERO_ = MAXBD_ + 1
                LOTE_ = LOTE_ & Format(NUMERO_, "00")
            
                ' Cabecera
                RstLote.AddNew
                RstLote("id") = xIdLote
                RstLote("iditem") = IDITEM_
                RstLote("fching") = FCHING_
                RstLote("descripcion") = LOTE_
                RstLote("cantidad") = CANTIDAD_
                RstLote.Update
                ' Detalle
                RstLoteDet.AddNew
                RstLoteDet("id") = xIdLoteDet
                RstLoteDet("idlote") = xIdLote
                RstLoteDet("idalm") = IDALM_
                RstLoteDet("cantidad") = CANTIDAD_
                RstLoteDet.Update
                
            Case 1 ' Quitar
                ' ---------CASO ERROR
                If IDLOTE_ = 0 And IDLOTEDET_ = 0 Then GoTo SALIRFUNCION_
                
                xIdLote = IDLOTE_
                xIdLoteDet = IDLOTEDET_
                ' Se actualiza detalle
                RstLoteDet.Filter = "id = " & xIdLoteDet
                If RstLoteDet.RecordCount = 0 Then GoTo LaCague
                RstLoteDet("cantidad") = RstLoteDet("cantidad") - CANTIDAD_
                RstLoteDet.Update
                ' Se actualiza cabecera
                RstLote.Filter = "id = " & xIdLote
                If RstLote.RecordCount = 0 Then GoTo LaCague
                RstLote("cantidad") = RstLote("cantidad") - CANTIDAD_
                RstLote.Update
        End Select
            
    ElseIf (MODO_ = 1) Then ' --------------------------------------------------------Modificar Registro
        ' Lotes Actuales
        xIdLote = IDLOTE_
        xIdLoteDet = IDLOTEDET_
        ' Lotes Anteriores
        xIdLoteAnt = IDLOTEANT_
        xIdLoteDetAnt = IDLOTEDETANT_
        
        Select Case TIPO_
            Case 0 ' Agregar
                ' ---------CASO ERROR
                If IDLOTE_ = 0 And IDLOTEDET_ = 0 Then GoTo SALIRFUNCION_
                
                ' ---------DETALLE
                ' Se disminuye lote anterior
                RstLoteDet.Filter = "id = " & xIdLoteDetAnt
                If RstLoteDet.RecordCount > 0 Then
                    RstLoteDet("cantidad") = RstLoteDet("cantidad") - CANANT_
                    RstLoteDet.Update
                End If
                ' Se actualiza lote actual
                RstLoteDet.Filter = adFilterNone
                RstLoteDet.Filter = "id = " & xIdLoteDet
                If RstLoteDet.RecordCount > 0 Then
                    RstLoteDet("cantidad") = RstLoteDet("cantidad") + CANTIDAD_
                    RstLoteDet.Update
                End If
                
                '----------CABECERA
                ' Se disminuye lote anterior
                RstLote.Filter = "id = " & xIdLoteAnt
                If RstLote.RecordCount > 0 Then
                    RstLote("cantidad") = NulosN(RstLote("cantidad")) - CANANT_
                    RstLote.Update
                End If
                ' Se actualiza lote actual
                RstLote.Filter = adFilterNone
                RstLote.Filter = "id = " & xIdLote
                If RstLote.RecordCount > 0 Then
                    RstLote("cantidad") = NulosN(RstLote("cantidad")) + CANTIDAD_
                    RstLote.Update
                End If
            
            Case 1 ' Quitar
                ' ---------CASO ERROR
                If IDLOTE_ = 0 And IDLOTEDET_ = 0 Then GoTo SALIRFUNCION_
                
                ' ---------DETALLE
                ' Se disminuye lote anterior
                RstLoteDet.Filter = "id = " & xIdLoteDetAnt
                If RstLoteDet.RecordCount > 0 Then
                    RstLoteDet("cantidad") = RstLoteDet("cantidad") + CANANT_
                    RstLoteDet.Update
                End If
                ' Se actualiza lote actual
                RstLoteDet.Filter = adFilterNone
                RstLoteDet.Filter = "id = " & xIdLoteDet
                If RstLoteDet.RecordCount > 0 Then
                    RstLoteDet("cantidad") = RstLoteDet("cantidad") - CANTIDAD_
                    RstLoteDet.Update
                End If
                
                '---------CABECERA
                ' Se disminuye lote anterior
                RstLote.Filter = "id = " & xIdLote
                If RstLote.RecordCount > 0 Then
                    RstLote("cantidad") = NulosN(RstLote("cantidad")) + CANANT_
                    RstLote.Update
                End If
                ' Se actualiza lote actual
                RstLoteDet.Filter = adFilterNone
                RstLoteDet.Filter = "id = " & xIdLoteDet
                If RstLoteDet.RecordCount > 0 Then
                    RstLote("cantidad") = NulosN(RstLote("cantidad")) - CANTIDAD_
                    RstLote.Update
                End If
            End Select
    End If
SALIRFUNCION_:
    xCon.CommitTrans
    crearModificarLote = xIdLoteDet
    Exit Function
LaCague:
    xCon.RollbackTrans
    MsgBox "Ocurrió un error al tratar de generar el lote para el producto"
End Function

Function GrabarMovimiento(FCHMOV_ As String, TIPDOC_ As Integer, NUMSER_ As String, _
                        IDRESP_ As Integer, IDPROV_ As Integer, DESPROV_ As String, _
                        IDESTADO_ As Integer, IDTIPMOV_ As Integer, IDTIPDOCREF_ As Integer, _
                        IDDOCREF_ As Integer, IDALM_ As Integer, RSTDET_ As ADODB.Recordset, _
                        Optional ByRef IDING_ As Integer = 0, Optional NUMDOC_ As String = "", _
                        Optional QUEHACE_ As Integer = 1, Optional MES_ As Integer = 0, _
                        Optional ANIO_ As Integer = 0) As Boolean
    
    ' IDTIPMOV_:
    ' -1: Ingreso, 0: Salida
    
    ' RSTDET:
    '___________________________________________________________________________________________________
    ' iditem|Integer, cantidad|Double, idalm|Integer, idtipo|Integer, idlote|Integer, idlotedet|Integer,
    ' canant|Integer, idloteant|Integer, idlotedetant|Integer
    '___________________________________________________________________________________________________
    
    Dim xId As Double
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim MODO_ As Integer
    Dim TIPO_ As Integer
    Dim IDITEM_ As Integer
    Dim IDLOTE_ As Integer
    Dim IDLOTEDET_ As Integer
    Dim IDLOTEANT_ As Integer
    Dim IDLOTEDETANT_ As Integer
    Dim IDTIPO_ As Integer
    Dim CANTIDAD_ As Double
    Dim CANTANT_ As Double
    
On Error GoTo LaCague

    xCon.BeginTrans
    If IDING_ = 0 Then
        xId = HallaCodigoTabla("alm_ingreso", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM alm_ingreso", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM alm_ingresodet", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = IDING_
        RST_Busq RstCab, "SELECT * FROM alm_ingreso WHERE id = " & xId & "", xCon
        xCon.Execute "DELETE * FROM alm_ingresodet WHERE id = " & xId & " "
        RST_Busq RstDet, "SELECT * FROM alm_ingresodet", xCon
    End If
    ' --------------------------------CABECERA DEL MOVIMIENTO
    RstCab("tipdoc") = TIPDOC_
    RstCab("fching") = FCHMOV_
    RstCab("fchdoc") = FCHMOV_
    RstCab("numser") = NUMSER_
    If NUMDOC_ = "" Then
        RstCab("numdoc") = hallarNumDoc("alm_ingreso", "'" & NUMSER_ & "'", "numser")
    Else
        RstCab("numdoc") = NUMDOC_
    End If
    RstCab("idres") = IDRESP_
    RstCab("idpro") = IDPROV_
    RstCab("nombre") = DESPROV_
    RstCab("estado") = IDESTADO_
    RstCab("tipmov") = IDTIPMOV_
    RstCab("idtipdocref") = IDTIPDOCREF_
    RstCab("iddocref") = IDDOCREF_
    RstCab("idalm") = IDALM_
    If ANIO_ = 0 Then RstCab("ano") = AnoTra Else RstCab("ano") = ANIO_
    If MES_ = 0 Then RstCab("idmes") = xMes Else RstCab("idmes") = MES_
    RstCab.Update
    
    RSTDET_.MoveFirst
    While Not RSTDET_.EOF
        ' --------------CRITERIOS PARA CREAR LOTE
        IDITEM_ = NulosN(RSTDET_("iditem"))
        CANTIDAD_ = NulosN(RSTDET_("cantidad"))
        IDLOTE_ = NulosN(RSTDET_("idlote"))
        IDLOTEDET_ = NulosN(RSTDET_("idlotedet"))
        IDLOTEANT_ = NulosN(RSTDET_("idloteant"))
        IDLOTEDETANT_ = NulosN(RSTDET_("idlotedetant"))
        CANTANT_ = NulosN(RSTDET_("canant"))
        IDTIPO_ = NulosN(RSTDET_("idtipo"))
        If IDLOTE_ = 0 Then MODO_ = 0 Else MODO_ = 1
        If IDTIPMOV_ = -1 Then TIPO_ = 0 Else TIPO_ = 1
        ' --------------GRABAMOS EL DETALLE DEL MOVIMIENTO
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("iditem") = NulosN(RSTDET_("iditem"))
        RstDet("cantidad") = NulosN(RSTDET_("cantidad"))
        RstDet("cantteo") = NulosN(RSTDET_("cantteo"))
        RstDet("idlotedet") = crearModificarLote(MODO_, TIPO_, IDITEM_, CANTIDAD_, IDALM_, FCHMOV_, IDLOTE_, IDLOTEDET_, CANTANT_, IDLOTEANT_, IDLOTEDETANT_)
        RstDet.Update
        
        RSTDET_.MoveNext
    Wend
    
    IDING_ = xId
    ' grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, 8, QUEHACE_, Time, Time, Date, xCon, xId
    
    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    MsgBox "El movimiento se registró con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    GrabarMovimiento = True
    Exit Function

LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo registrar el movimiento por el siguiente motivo :" + Trim(Err.Description)
    GrabarMovimiento = False
    Exit Function
End Function

Public Function calcularProdAnterior(IDREC_ As Integer) As Variant
    Dim xRs As New ADODB.Recordset
    Dim cSQL As String
    
    cSQL = "SELECT pro_recetains.iditem, alm_inventario.descripcion " _
        + vbCr + "FROM (pro_lineadet LEFT JOIN pro_recetains ON pro_lineadet.idrec = pro_recetains.idrec) LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id " _
        + vbCr + "WHERE (((pro_recetains.idrec) = " & IDREC_ & ") And ((alm_inventario.tippro) <= 3)) " _
        + vbCr + "GROUP BY pro_recetains.iditem, alm_inventario.descripcion;"
    
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then calcularProdAnterior = "": Exit Function
    If xRs.RecordCount = 0 Then calcularProdAnterior = "": Exit Function
    
    calcularProdAnterior = NulosC(xRs("descripcion"))
End Function

Public Sub llenarEstado(TIPO_ As Integer, ORIGEN_ As Integer, _
                                Optional ByRef FGGRID As VSFlexGrid, _
                                Optional ByRef COMBO_ As ComboBox, _
                                Optional ESTADOAESTABLECER_ As Double, _
                                Optional COLUMNAGRID_ As Integer, _
                                Optional LLENARTODOSESTADOS_ As Boolean = True, _
                                Optional ESTADOSATOMARENCUENTA_ As String)
    '****************************************************
    ' TIPO=0: ESTABLECER ESTADO; TIPO=1: LLENAR ESTADOS
    ' ORIGEN=0: FLEXGRID; ORIGEN=1: COMBOBOX
    '****************************************************
    
    Dim CAMPOS As String
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim cSQL As String
    
    Select Case TIPO_
        Case 0 ' ----------------------------------ESTABLECER ESTADO
            Select Case ORIGEN_
                Case 0 ' ------------------------------FLEXGRID
                Case 1 ' ------------------------------COMBOBOX
                    For A = 0 To COMBO_.ListCount - 1
                        If COMBO_.ItemData(A) = ESTADOAESTABLECER_ Then
                            COMBO_.ListIndex = A
                            Exit For
                        End If
                    Next A
            End Select
            
        Case 1 ' ----------------------------------LLENAR ESTADOS
            If LLENARTODOSESTADOS_ Then
                cSQL = "SELECT * FROM mae_estados ORDER BY id"
            Else
                cSQL = "SELECT * " _
                    + vbCr + "FROM mae_estados " _
                    + vbCr + "WHERE (((mae_estados.id) In (" & ESTADOSATOMARENCUENTA_ & "))) " _
                    + vbCr + "ORDER BY mae_estados.id;"
            End If
    
            RST_Busq xRs, cSQL, xCon
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then
                MsgBox "No se ha encontrado estados, Ingrese estados", vbInformation, xTitulo
                Exit Sub
            End If
                
            Select Case ORIGEN_
                Case 0 ' --------------------------FLEXGRID
                    xRs.MoveFirst
                    CAMPOS = "#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
                    xRs.MoveNext
                    While Not xRs.EOF
                        CAMPOS = CAMPOS & "|#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
                        xRs.MoveNext
                    Wend
                    FGGRID.ColComboList(COLUMNAGRID_) = CAMPOS
                
                Case 1 ' --------------------------COMBOBOX
                    COMBO_.Clear
                    xRs.MoveFirst
                    While Not xRs.EOF
                        COMBO_.AddItem UCase(NulosC(xRs("descripcion")))
                        COMBO_.ItemData(COMBO_.NewIndex) = NulosN(xRs("id"))
                        
                        xRs.MoveNext
                    Wend
                    COMBO_.ListIndex = 0
            End Select
    End Select
End Sub

Public Sub ImprimirLinea(NUMORDPORD_ As String, ITEM_ As String, IDRECETA_ As Integer, RECETA_ As String, _
                                        FCHPROD_ As String, CANTIDAD_ As Double, RESPONSABLE_ As String, _
                                        UNIMED_ As String, NUMEROTOTPER_ As Integer, _
                                        RSTTAREAS_ As ADODB.Recordset, RSTPERSONAS_ As ADODB.Recordset)
    Dim NUMEROPAGINA_ As Integer
    Dim RSTTEMP_ As New ADODB.Recordset
    Dim RSTTAREASAUX_ As New ADODB.Recordset
    Dim IDAREA_ As Double
    Dim IDRESPONSABLE_ As Double
    Dim CAMBIO_ As Boolean
    Dim NUMEROTOTPERAUX_ As Integer
    
    With FrmVsPrinter.Vs
        .BrushColor = &H80000005
        .FontSize = 11
        .TextAlign = taCenterMiddle
        
        If RSTTAREAS_.State = 0 Then Exit Sub
        If RSTTAREAS_.RecordCount = 0 Then Exit Sub
          
        DEFINIR_RST_TMP RSTTAREASAUX_, RSTTAREAS_
        CARGAR_RST_TMP RSTTAREASAUX_, RSTTAREAS_
        RSTTAREAS_.MoveFirst
        IDAREA_ = NulosN(RSTTAREAS_("idarea"))
        IDRESPONSABLE_ = NulosN(RSTTAREAS_("idsubresp"))
        NUMEROPAGINA_ = 0
        CAMBIO_ = True
        While Not RSTTAREAS_.EOF
            If Not CAMBIO_ Then GoTo SIGUIENTE_
            DEFINIR_RST_TMP RSTTEMP_, RSTTAREASAUX_
            CARGAR_RST_TMP RSTTEMP_, RSTTAREASAUX_
            NUMEROTOTPERAUX_ = NUMEROTOTPER_
            imprimirDetallado NUMORDPORD_, ITEM_, IDRECETA_, RECETA_, FCHPROD_, CANTIDAD_, _
                                        UCase(Busca_Codigo(NulosN(IDRESPONSABLE_), "id", "nombre", "pla_empleados", "N", xCon)), _
                                        UNIMED_, NUMEROPAGINA_, NUMEROTOTPERAUX_, RSTTEMP_, RSTPERSONAS_
            
SIGUIENTE_:
            RSTTAREAS_.MoveNext
            If Not RSTTAREAS_.EOF Then
                If IDAREA_ <> NulosN(RSTTAREAS_("idarea")) Or IDRESPONSABLE_ <> NulosN(RSTTAREAS_("idsubresp")) Then
                    CAMBIO_ = True
                    NUMEROPAGINA_ = NUMEROPAGINA_ + 1
                    IDAREA_ = NulosN(RSTTAREAS_("idarea"))
                    IDRESPONSABLE_ = NulosN(RSTTAREAS_("idsubresp"))
                    .NewPage
                    CrearCabeceraVS NUMEROPAGINA_
                Else
                    CAMBIO_ = False
                End If
            End If
        Wend
        
SIGUIENTE:
    End With
End Sub
Public Function ImprimirOP(RSTCAB_ As ADODB.Recordset) As Boolean
    Dim A As Integer
    Dim numPag As Integer
    Dim Rst As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim B As Integer
    Dim FILA_ As Integer
    Dim COLUMNA_ As Integer
    Dim numper As Double
    Dim xFila As Integer
    Dim Nombre As String
    Dim RSTDET_ As New ADODB.Recordset
    Dim cSQL As String
    
    
    If RSTCAB_.State = 0 Then ImprimirOP = False: Exit Function
    RSTCAB_.Filter = adFilterNone
    If RSTCAB_.RecordCount = 0 Then ImprimirOP = False: Exit Function
        
    With FrmVsPrinter.Vs
    
'        .ExportFormat = vpxDHTML
'        .ExportFile = "c:\report2.htm"
    
        numPag = 0
        .BrushColor = &H80000005
        .FontSize = 11
        .TextAlign = taCenterMiddle
        .StartDoc
            
        '*************************************************************
        ' ------------------------------------------CABECERA DE PAGINA
        '*************************************************************
        FILA_ = 800
        COLUMNA_ = 1000
        numPag = numPag + 1
        CrearCabeceraVS numPag
        
        RSTCAB_.MoveFirst
        While Not RSTCAB_.EOF
            '******************************************************
            ' -----------------------------------------------TITULO
            '******************************************************
            If FILA_ >= 13000 Then
                .NewPage
                FILA_ = 800
                numPag = numPag + 1
                CrearCabeceraVS numPag
            End If
            .FontSize = 12
            .FontBold = True
            .TextAlign = taCenterMiddle
            .TextBox "ORDEN DE PRODUCCION", COLUMNA_, FILA_, 8000, 500, True, False, True
            .FontSize = 10
            .TextAlign = taCenterTop
            .TextBox "Nº", COLUMNA_ + 8100, FILA_, 1900, 250, True, False, True
            FILA_ = FILA_ + 240
            .TextBox NulosC(RSTCAB_("numser")) & "-" & NulosC(RSTCAB_("numdoc")), COLUMNA_ + 8100, FILA_, 1900, 250, True, False, True

            '********************************************************
            ' -----------------------------------------------CABECERA
            '********************************************************
            .TextAlign = taLeftMiddle
            .FontSize = 9
            FILA_ = FILA_ + 300
            .FontBold = True
            .TextBox "Fecha       :", COLUMNA_, FILA_, 1500, 250, True, False, False
            .FontBold = False
            .TextBox Format(NulosC(RSTCAB_("fchpro")), FORMAT_DATE), COLUMNA_ + 1500, FILA_, 7000, 250, True, False, False
            .FontBold = True
            .TextBox "Nº Doc. Ref.:", COLUMNA_ + 6000, FILA_, 1500, 250, True, False, False
            .FontBold = False
            .TextBox NulosC(RSTCAB_("numdocref")), COLUMNA_ + 7500, FILA_, 6000, 250, True, False, False

            FILA_ = FILA_ + 250
            .FontBold = True
            .TextBox "Responsable :", COLUMNA_, FILA_, 1500, 250, True, False, False
            .FontBold = False
            .TextBox NulosC(RSTCAB_("resp")), COLUMNA_ + 1500, FILA_, 6000, 250, True, False, False
            .FontBold = True
            .TextBox "N° Lote     :", COLUMNA_ + 6000, FILA_, 1500, 250, True, False, False
            .FontBold = False
            .TextBox NulosC(RSTCAB_("lote")), COLUMNA_ + 7500, FILA_, 6000, 250, True, False, False

            '*******************
            ' FILA PRODUCTO
            '*******************
            FILA_ = FILA_ + 350
            .TextAlign = taCenterMiddle
            .FontBold = True
            .TextBox "PRODUCTO", COLUMNA_, FILA_, 1750, 500, True, False, True
            .FontBold = False
            .TextBox NulosC(RSTCAB_("codrec")), COLUMNA_ + 1750, FILA_, 1550, 500, True, False, True
            .TextBox NulosC(RSTCAB_("desitem")), COLUMNA_ + 3300, FILA_, 6700, 500, True, False, True

            '*******************
            ' FILA DATOS DE PRODUCCION
            '*******************
            FILA_ = FILA_ + 500
            .TextAlign = taCenterMiddle
            .FontBold = True
            .TextBox "DATOS DE PRODUCCION", COLUMNA_, FILA_, 5050, 500, True, False, True
            .FontBold = False
            .TextAlign = taLeftMiddle
            .TextBox "", COLUMNA_ + 5050, FILA_, 4950, 1250, True, False, True

            FILA_ = FILA_ + 150
            .FontBold = True
            .TextBox " Observaciones:", COLUMNA_ + 5050, FILA_, 2500, 250, True, False, False
            FILA_ = FILA_ + 250
            .FontBold = False
            .TextBox " " & NulosC(RSTCAB_("glosa")), COLUMNA_ + 5050, FILA_, 6000, 250, True, False, False

            FILA_ = FILA_ + 100
            .TextAlign = taCenterMiddle
            .TextBox "Código", COLUMNA_, FILA_, 1750, 500, True, False, True
            .TextBox "U.M.", COLUMNA_ + 1750, FILA_, 1550, 500, True, False, True
            .TextBox "Cantidad", COLUMNA_ + 3300, FILA_, 1750, 500, True, False, True

            FILA_ = FILA_ + 500
            .TextAlign = taLeftMiddle
            .TextBox " " & NulosC(RSTCAB_("coditem")), COLUMNA_, FILA_, 1750, 250, True, False, True
            .TextBox " " & NulosC(RSTCAB_("desunimed")), COLUMNA_ + 1750, FILA_, 1550, 250, True, False, True
            .TextBox " " & Format(NulosN(RSTCAB_("cantidad")), FORMAT_CANTIDADDECIMAL), COLUMNA_ + 3300, FILA_, 1750, 250, True, False, True

            FILA_ = FILA_ + 250

            '*******************************************************
            '------------------------------------------------DETALLE
            '*******************************************************
            cSQL = "SELECT alm_inventario.codpro AS coditem, alm_inventario.descripcion AS desitem, mae_unidades.abrev AS desunimed, pro_recetains.iditem, [pro_recetains]![canpro]*" & NulosN(RSTCAB_("cantidad")) & " AS cantidad, pro_recetains.idunimed " _
            + vbCr + "FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((pro_recetains.idrec)=" & NulosN(RSTCAB_("idrec")) & "));"

            Set RSTDET_ = Nothing
            RST_Busq RSTDET_, cSQL, xCon

            If RSTDET_.State = 0 Then ImprimirOP = False: Exit Function
            If RSTDET_.RecordCount = 0 Then ImprimirOP = False: Exit Function

            FILA_ = FILA_ + 350
            .TextAlign = taCenterMiddle
            .TextBox "Código", COLUMNA_, FILA_, 1750, 500, True, False, True
            .TextBox "Ítem", COLUMNA_ + 1750, FILA_, 5900, 500, True, False, True
            .TextBox "U.M.", COLUMNA_ + 7650, FILA_, 800, 500, True, False, True
            .TextBox "Cantidad", COLUMNA_ + 8450, FILA_, 1550, 500, True, False, True

            FILA_ = FILA_ + 250

            xFila = FILA_
            While Not RSTDET_.EOF
                FILA_ = FILA_ + 250
                If FILA_ >= 16200 Then
                    FILA_ = 800
                    numPag = numPag + 1
                    .NewPage
                    CrearCabeceraVS numPag
                End If

                .FontSize = 8
                .FontBold = False
                .TextAlign = taLeftMiddle
                .TextBox " " & RSTDET_("coditem"), COLUMNA_, FILA_, 1750, 250, True, False, True
                .FontSize = 7
                .TextBox " " & RSTDET_("desitem"), COLUMNA_ + 1750, FILA_, 5900, 250, True, False, True
                .FontSize = 8
                .TextAlign = taCenterMiddle
                .TextBox NulosC(RSTDET_("desunimed")), COLUMNA_ + 7650, FILA_, 800, 250, True, False, True
                .TextAlign = taRightMiddle
                .TextBox Format(RSTDET_("cantidad"), FORMAT_CANTIDADDECIMAL), COLUMNA_ + 8450, FILA_, 1550, 250, True, False, True

                RSTDET_.MoveNext
            Wend

            FILA_ = FILA_ + 1200
            If FILA_ >= 16000 Then
                FILA_ = 2000
                .NewPage
            End If

            .TextBox "_______________________________", COLUMNA_, FILA_, 3500, 250, True, False, False
            .TextBox "_______________________________", COLUMNA_ + 4500, FILA_, 3500, 250, True, False, False

            FILA_ = FILA_ + 200

            .TextAlign = taCenterMiddle
            .TextBox "RESPONSABLE", COLUMNA_, FILA_, 3500, 250, True, False, False
            .TextBox "SUPERVISOR", COLUMNA_ + 4500, FILA_, 3500, 250, True, False, False
            .TextAlign = taRightMiddle
            .TextBox "FO-PRO-001", COLUMNA_ + 8700, FILA_, 1200, 250, True, False, False

            FILA_ = FILA_ + 800

            RSTCAB_.MoveNext
        Wend
        .EndDoc
    End With
    'Muestra la preimagen de la impresion
    FrmVsPrinter.WindowState = 2
    FrmVsPrinter.Show
End Function

Private Sub imprimirDetallado(NUMORDPORD_ As String, ITEM_ As String, IDRECETA_ As Integer, RECETA_ As String, _
                                        FCHPROD_ As String, CANTIDAD_ As Double, RESPONSABLE_ As String, _
                                        UNIMED_ As String, ByRef numPag As Integer, NUMEROTOTPER_ As Integer, _
                                        RSTTAREAS_ As ADODB.Recordset, RSTPERSONAS_ As ADODB.Recordset)
    Dim HORFIN_ As String
    Dim CAMBIO_ As Boolean
    Dim A As Integer
    Dim xRsTarAuxAux As New ADODB.Recordset
    Dim xLinea As Double
    Dim B As Integer
    Dim xColumna As Integer         ' Columna de impresion
    Dim ID_LINEA As Double
    Dim xFila As Integer
                
    With FrmVsPrinter.Vs
        xLinea = 700
        xColumna = 900
        numPag = numPag + 1
        CrearCabeceraVS numPag
        
        ' ------------------------------------------------------CABECERA
        .FontSize = 12
        .FontBold = True
        .TextAlign = taCenterMiddle
        
        .TextBox "ORDEN DE PRODUCCION", xColumna, xLinea, 8000, 500, True, False, True
        .FontSize = 10
        .TextAlign = taCenterTop
        .TextBox "Nº", xColumna + 8100, xLinea, 1900, 250, True, False, True
        xLinea = xLinea + 240
        .FontSize = 9
        .TextBox NUMORDPORD_, xColumna + 8100, xLinea, 1900, 250, True, False, True
        
        .TextAlign = taLeftMiddle
        .FontSize = 9
        .FontBold = False
        xLinea = xLinea + 250
        .TextBox "Producto", xColumna, xLinea, 1500, 250, True, False, False
        .TextBox ITEM_, xColumna + 1500, xLinea, 7000, 250, True, False, False
        .TextBox "Receta", xColumna + 7500, xLinea, 1000, 250, True, False, False
        .TextBox RECETA_, xColumna + 8550, xLinea, 6000, 250, True, False, False
        xLinea = xLinea + 250
        .TextBox "Fecha Prod.", xColumna, xLinea, 1500, 250, True, False, False
        .TextBox FCHPROD_, xColumna + 1500, xLinea, 6000, 250, True, False, False
        .TextBox "Cantidad", xColumna + 7500, xLinea, 1000, 250, True, False, False
        .TextBox Format(CANTIDAD_, "0.00") & " " & UNIMED_, xColumna + 8550, xLinea, 6000, 250, True, False, False
        xLinea = xLinea + 250
        .TextBox "Responsable ", xColumna, xLinea, 1500, 250, True, False, False
        .TextBox RESPONSABLE_, xColumna + 1500, xLinea, 6000, 250, True, False, False
            
        ' --------------------------------------------DETALLE DE TAREAS
        xLinea = xLinea + 300
        .TextAlign = taLeftMiddle
        .FontBold = True
        .TextBox "Detalles de la Linea", xColumna, xLinea, 2500, 250, True, False, False
        
        .FontBold = False
        xLinea = xLinea + 350
        .TextAlign = taCenterMiddle
        .TextBox "Ord.", xColumna, xLinea, 500, 500, True, False, True
        .TextBox "Tarea", xColumna + 500, xLinea, 3500, 500, True, False, True
        .TextBox "Durac.", xColumna + 4000, xLinea, 800, 500, True, False, True
        .TextBox "Hor.Ini", xColumna + 4800, xLinea, 800, 500, True, False, True
        .TextBox "Hor.Fin", xColumna + 5600, xLinea, 800, 500, True, False, True
        .TextBox "Num. Pers.", xColumna + 6400, xLinea, 800, 500, True, False, True
        .TextBox "Unid.x Hora", xColumna + 7200, xLinea, 1000, 500, True, False, True
        .TextBox "%Rdto", xColumna + 8200, xLinea, 800, 500, True, False, True
        .TextBox "Cant. Proc.", xColumna + 9000, xLinea, 1000, 500, True, False, True
        
        xLinea = xLinea + 250
        xFila = xLinea
        
        RSTTAREAS_.MoveFirst
        While Not RSTTAREAS_.EOF
            xLinea = xLinea + 250
            .FontSize = 8
            .FontBold = False
            .TextAlign = taLeftMiddle
            .TextBox " " & NulosN(RSTTAREAS_("idtar")), xColumna, xLinea, 500, 250, True, False, True
            .TextBox " " & UCase(Busca_Codigo(NulosN(RSTTAREAS_("idtar")), "id", "descripcion", "pro_tareas", "N", xCon)), xColumna + 500, xLinea, 3500, 250, True, False, True
            .TextAlign = taCenterMiddle
            .TextBox Format(RSTTAREAS_("durtar"), "HH:mm"), xColumna + 4000, xLinea, 800, 250, True, False, True
            .TextBox Format(RSTTAREAS_("horini"), "HH:mm"), xColumna + 4800, xLinea, 800, 250, True, False, True
            .TextBox Format(RSTTAREAS_("horfin"), "HH:mm"), xColumna + 5600, xLinea, 800, 250, True, False, True
            .TextBox Format(RSTTAREAS_("numop"), "00"), xColumna + 6400, xLinea, 800, 250, True, False, True
            .TextAlign = taRightMiddle
            .TextBox "", xColumna + 7200, xLinea, 1000, 250, True, False, True
            .TextBox "", xColumna + 8200, xLinea, 800, 250, True, False, True
            .TextBox Format(RSTTAREAS_("cantproc"), FORMAT_CANTIDAD) & " ", xColumna + 9000, xLinea, 1000, 250, True, False, True
                        
            RSTTAREAS_.MoveNext
            
            If xLinea >= 16200 Then
                xLinea = 1300
                numPag = numPag + 1
                .NewPage
                CrearCabeceraVS numPag
            End If
        Wend
            
        xLinea = xLinea + 250
        .TextAlign = taRightMiddle
        .TextBox "TOTAL", xColumna, xLinea, 4000, 250, True, False, True
        .TextAlign = taCenterMiddle
        .TextBox Format(NUMEROTOTPER_, "00"), xColumna + 6400, xLinea, 800, 250, True, False, True
        .FontBold = False
        xLinea = xLinea + 400
        .TextBox "RECETA", xColumna + 4750, xLinea, 1500, 250, True, False, True
        .TextBox "CANTIDAD", xColumna + 6250, xLinea, 1000, 250, True, False, True
        .TextAlign = taCenterMiddle
        xLinea = xLinea + 250
        .FontSize = 7
        .TextBox calcularProdAnterior(IDRECETA_), xColumna + 500, xLinea, 4250, 250, True, False, True
        .FontSize = 8
        .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
        .TextAlign = taLeftMiddle
        .TextBox " Hora Ini.", xColumna + 7500, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 8500, xLinea, 1500, 250, True, False, True
        .TextAlign = taCenterMiddle
        xLinea = xLinea + 250
        .TextBox "P1", xColumna, xLinea, 500, 250, True, False, True
        .FontSize = 7
        .TextBox ITEM_, xColumna + 500, xLinea, 4250, 250, True, False, True
        .FontSize = 8
        .TextBox RECETA_, xColumna + 4750, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
        .TextAlign = taLeftMiddle
        .TextBox " Hora Fin", xColumna + 7500, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 8500, xLinea, 1500, 250, True, False, True
        xLinea = xLinea + 250
        .TextAlign = taCenterMiddle
        .TextBox "P2", xColumna, xLinea, 500, 250, True, False, True
        .TextBox "", xColumna + 500, xLinea, 4250, 250, True, False, True
        .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
        xLinea = xLinea + 250
        .TextBox "P3", xColumna, xLinea, 500, 250, True, False, True
        .TextBox "", xColumna + 500, xLinea, 4250, 250, True, False, True
        .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
        
        ' ------------------------------------------------------DETALLE DE PERSONAL
        xLinea = xLinea + 300
        .TextAlign = taLeftMiddle
        .FontBold = True
        .TextBox "Detalles del Personal", xColumna, xLinea, 2500, 250, True, False, False
        xLinea = xLinea + 350
        .FontBold = False
        .TextAlign = taCenterMiddle
        .TextBox "Item", xColumna, xLinea, 500, 500, True, False, True
        .TextBox "PERSONAL", xColumna + 500, xLinea, 3500, 500, True, False, True
        .TextBox "Tarea", xColumna + 4000, xLinea, 800, 500, True, False, True
        .TextBox "Hr.Ini.", xColumna + 4800, xLinea, 1000, 500, True, False, True
        .TextBox "Hr.Ter.", xColumna + 5800, xLinea, 1000, 500, True, False, True
        .TextBox "M.P.", xColumna + 6800, xLinea, 800, 500, True, False, True
        .TextBox "Prod1", xColumna + 7600, xLinea, 600, 500, True, False, True
        .TextBox "Prod2", xColumna + 8200, xLinea, 600, 500, True, False, True
        .TextBox "Prod3", xColumna + 8800, xLinea, 600, 500, True, False, True
        .TextBox "Efic.", xColumna + 9400, xLinea, 600, 500, True, False, True
            
        If RSTPERSONAS_.RecordCount > 0 Then RSTPERSONAS_.MoveFirst
        xLinea = xLinea + 500
        xFila = xLinea
        
        ' Se agrega 5 campos mas para ingresar personal
        NUMEROTOTPER_ = NUMEROTOTPER_ + 5
        For B = 1 To NUMEROTOTPER_
            .FontSize = 10
            .FontBold = False
            .TextAlign = taLeftMiddle
            
            .TextBox " " & Format(B, "00"), xColumna, xLinea, 500, 300, True, False, True
            If Not RSTPERSONAS_.EOF Then
                ' UCase(Busca_Codigo(NulosN(xRs("idper")), "id", "nombre", "pla_empleados", "N", xCon))
                '.TextBox " " & NulosC(RSTPERSONAS_("nombre")), xColumna + 500, xLinea, 3500, 300, True, False, True
                .TextBox " " & UCase(Busca_Codigo(NulosN(RSTPERSONAS_("idper")), "id", "nombre", "pla_empleados", "N", xCon)), xColumna + 500, xLinea, 3500, 300, True, False, True
                .TextBox "", xColumna + 4000, xLinea, 800, 300, True, False, True
                .TextBox "", xColumna + 4800, xLinea, 1000, 300, True, False, True
                .TextBox "", xColumna + 5800, xLinea, 1000, 300, True, False, True
                .TextBox "", xColumna + 6800, xLinea, 800, 300, True, False, True
                .TextBox "", xColumna + 7600, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 8200, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 8800, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 9400, xLinea, 600, 300, True, False, True
                RSTPERSONAS_.MoveNext
            Else
                .TextBox "", xColumna + 500, xLinea, 3500, 300, True, False, True
                .TextBox "", xColumna + 4000, xLinea, 800, 300, True, False, True
                .TextBox "", xColumna + 4800, xLinea, 1000, 300, True, False, True
                .TextBox "", xColumna + 5800, xLinea, 1000, 300, True, False, True
                .TextBox "", xColumna + 6800, xLinea, 800, 300, True, False, True
                .TextBox "", xColumna + 7600, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 8200, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 8800, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 9400, xLinea, 600, 300, True, False, True
            End If
            
            xLinea = xLinea + 300
            If xLinea >= 14750 Then
                xLinea = 1300
                numPag = numPag + 1
                .NewPage
                CrearCabeceraVS numPag
            End If
        Next B
        
        ' -------------------------------------------------OBSERVACIONES
        xLinea = xLinea + 100
        
        If xLinea >= 15500 Then
            xLinea = 1300
            numPag = numPag + 1
            .NewPage
            CrearCabeceraVS numPag
        End If
        
        .TextAlign = taLeftMiddle
        .FontBold = True
        .TextBox "Observaciones", xColumna, xLinea, 2500, 250, True, False, False
        xLinea = xLinea + 450
        .DrawLine xColumna + 500, xLinea, 10000, xLinea
        xLinea = xLinea + 250
        .DrawLine xColumna + 500, xLinea, 10000, xLinea
        xLinea = xLinea + 250
        .DrawLine xColumna + 500, xLinea, 10000, xLinea
        xLinea = xLinea + 250
        .DrawLine xColumna + 500, xLinea, 10000, xLinea
        
    End With
End Sub

Sub CrearCabeceraVS(numPag As Integer)
    Dim xCad As String

    FrmVsPrinter.Vs.TextAlign = taLeftTop
    FrmVsPrinter.Vs.FontName = "Courier New"
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.FontSize = 9

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 200
    FrmVsPrinter.Vs.Paragraph = "EMPRESA   : " & NomEmp

'    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 200
'    FrmVsPrinter.Vs.Paragraph = "FECHA        : " & Format(Date, "dd/mm/yy")

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 400
    FrmVsPrinter.Vs.Paragraph = "Nº R.U.C. : " & NumRUC

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 400
    FrmVsPrinter.Vs.Paragraph = "Nº Pagina    : " & Format(numPag, "0000")

    FrmVsPrinter.Vs.DrawLine 1000, 650, 11000, 650
    
'    With FrmVsPrinter.Vs
'        .Header = "QuickPrinter||Page %d||Prueba||Test"
'        .Footer = Format(Now, "dd-mmm-yy") & "||Prueba"
'        .MarginLeft = "0.5in"
'        .MarginRight = "0.5in"
'        .PageBorder = pbAll
'
'        .StartTable
'        ' create table with four rows and four columns
'        .TableCell(tcCols) = 5
'        .TableCell(tcRows) = 2
'        ' set some column widths (default width is 0.5in)
'        .TableCell(tcColWidth, , 1) = "1in"
'        .TableCell(tcColWidth, , 2) = "1.5in"
'        .TableCell(tcColWidth, , 3) = "2in"
'        .TableCell(tcColWidth, , 4) = "1in"
'        .TableCell(tcColWidth, , 5) = "1.5in"
'        ' assign text to each cell
'
'        .TableCell(tcText, 1, 1) = "EMPRESA"
'        .TableCell(tcText, 1, 2) = NomEmp
'        .TableCell(tcText, 1, 4) = "FECHA"
'        .TableCell(tcText, 1, 5) = Format(Date, "dd/mm/yy")
'        .TableCell(tcText, 2, 1) = "N° R.U.C"
'        .TableCell(tcText, 2, 2) = NumRUC
'        .TableCell(tcText, 2, 4) = "N° PAGINA"
'        .TableCell(tcText, 2, 5) = Format(numPag, "0000")
'
'        .TableBorder = tbAll
''        .TableCell(tcColBorder, 1, 1, 2, 5) = tbNone
'
'        ' format cell (1,1): make it span two columns, with a
'        ' blue background, center alignment, and bold
''        .TableCell(tcColSpan, 1, 1) = 2
''        .TableCell(tcBackColor, 1, 1) = vbBlue
''        .TableCell(tcAlign, 1, 1) = taCenterMiddle
''        .TableCell(tcFontBold, 1, 1) = True
'        ' set row height for row 1
'        ' (default height is calculated to fit the contents)
''        .TableCell(tcRowHeight, 1) = "0.2in"
'        ' format cell (3,2): make is span two columns, with a
'        ' yellow background, center alignment, and bold
''        .TableCell(tcColSpan, 3, 2) = 2
''        .TableCell(tcBackColor, 3, 2) = vbYellow
''        .TableCell(tcAlign, 3, 2) = taCenterMiddle
''        .TableCell(tcFontBold, 3, 2) = True
''        ' set row height for row 3
''        .TableCell(tcRowHeight, 3) = "0.2in"
''        ' set row borders all around
''        .TableBorder = tbAll
''        ' finish table definition
'        .EndTable
'
'
'
'        .StartTable
'        ' create table with four rows and four columns
'        .TableCell(tcCols) = 5
'        .TableCell(tcRows) = 2
'        ' set some column widths (default width is 0.5in)
'        .TableCell(tcColWidth, , 1) = "1in"
'        .TableCell(tcColWidth, , 2) = "1.5in"
'        .TableCell(tcColWidth, , 3) = "2in"
'        .TableCell(tcColWidth, , 4) = "1in"
'        .TableCell(tcColWidth, , 5) = "1.5in"
'        ' assign text to each cell
'
'        .TableCell(tcText, 1, 1) = "EMPRESA"
'        .TableCell(tcText, 1, 2) = NomEmp
'        .TableCell(tcText, 1, 4) = "FECHA"
'        .TableCell(tcText, 1, 5) = Format(Date, "dd/mm/yy")
'        .TableCell(tcText, 2, 1) = "N° R.U.C"
'        .TableCell(tcText, 2, 2) = NumRUC
'        .TableCell(tcText, 2, 4) = "N° PAGINA"
'        .TableCell(tcText, 2, 5) = Format(numPag, "0000")
'
'        .TableBorder = tbAll
'        .EndTable
'    End With
End Sub

Private Sub BuildCalendar(TheDay As Date, _
 ByRef TheCal(), _
 ByRef TheDayRow, ByRef TheDayCol)
 Dim dt As Date
 ' clear array
 ReDim TheCal(7, 1)
 ' initialize date to the first of the month
 dt = TheDay
 While Day(dt) > 1
 dt = dt - 1
 Wend
 ' fill array with dates for current month
 Dim r%, C%
 r = 0
 C = Weekday(dt) - 1
 While Month(dt) = Month(TheDay)
 ' add row if we have to
 If C >= 7 Then
 C = 0
 r = r + 1
 ReDim Preserve TheCal(7, r)
 End If
 ' save day value in the calendar
 TheCal(C, r) = Day(dt)
 ' return TheDate's row and column
 If dt = TheDay Then
 TheDayRow = r
 TheDayCol = C
 End If
 ' increment day
 dt = dt + 1
 C = C + 1
 Wend
End Sub

Function grabarOrdProd(FCHDOC_ As String, IDTIPDOCREF_ As Integer, _
                                    IDDOCREF_ As Integer, IDRESP_ As Integer, _
                                    IDREC_ As Integer, IDUNIMED_ As Integer, _
                                    CANTIDAD_ As Double, IDLINEA_ As Integer, _
                                    EFIC_ As Integer, HORINI_ As String, _
                                    HORFIN_ As String, FCHFIN_ As String, _
                                    NUMOP_ As Integer, REPROC_ As Boolean, _
                                    NUMDOC_ As String, LOTE_ As String, GLOSA_ As String, _
                                    RSTTAR_ As ADODB.Recordset, _
                                    RSTPER_ As ADODB.Recordset, RSTREP_ As ADODB.Recordset, _
                                    Optional NUMSER_ As String = "0001", Optional ByRef IDORD_ As Integer, _
                                    Optional ESTADO_ As Integer = 1, Optional ANIO_ As Integer, _
                                    Optional MES_ As Integer, Optional QUEHACE_ As Integer) As Boolean
    
    Dim RstCab As New ADODB.Recordset
    Dim RstPer As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim RstRep As New ADODB.Recordset
    Dim xId As Double
    
'On Error GoTo LaCague
    xCon.BeginTrans

    If IDORD_ = 0 Then
        ' Obetenemos el Id del registro
        xId = HallaCodigoTabla("pro_ordenprod", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pro_ordenprod", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = IDORD_
        RST_Busq RstCab, "SELECT * FROM pro_ordenprod WHERE id=" & xId, xCon
        ' ELIMINAMOS DETALLES
        xCon.Execute "DELETE * FROM pro_ordenprodtar WHERE idord=" & xId
        xCon.Execute "DELETE * FROM pro_ordenprodpers WHERE idord=" & xId
        xCon.Execute "DELETE * FROM pro_ordenprodreproc WHERE idord=" & xId
    End If
    
    RST_Busq RstTar, "SELECT TOP 1 * FROM pro_ordenprodtar", xCon
    RST_Busq RstPer, "SELECT TOP 1 * FROM pro_ordenprodpers", xCon
    RST_Busq RstRep, "SELECT TOP 1 * FROM pro_ordenprodreproc", xCon
        
    ' ---------------------------------------CABECERA
    RstCab("numser") = NUMSER_
    If NUMDOC_ = "" Then
        RstCab("numdoc") = hallarNumDoc("pro_ordenprod", NUMSER_, "numser")
    Else
        RstCab("numdoc") = NUMDOC_
    End If
    
    RstCab("lote") = LOTE_
    RstCab("idtipdocref") = IDTIPDOCREF_
    RstCab("iddocref") = IDDOCREF_
    RstCab("fchpro") = FCHDOC_
    RstCab("idresp") = IDRESP_
    RstCab("idrec") = IDREC_
    RstCab("idunimed") = IDUNIMED_
    RstCab("cantidad") = CANTIDAD_
    RstCab("idlinea") = IDLINEA_
    RstCab("efic") = EFIC_
    If HORINI_ <> "" Then RstCab("horini") = HORINI_
    If HORFIN_ <> "" Then RstCab("horfin") = HORFIN_
    If FCHFIN_ <> "" Then RstCab("fchfin") = FCHFIN_
    RstCab("numop") = NUMOP_
    RstCab("reproc") = REPROC_
    RstCab("estado") = ESTADO_
    RstCab("glosa") = GLOSA_
    If ANIO_ = 0 Then RstCab("ano") = AnoTra Else RstCab("ano") = ANIO_
    If MES_ = 0 Then RstCab("idmes") = xMes Else RstCab("idmes") = MES_
    RstCab.Update
    ' --------------------------------------TAREAS
    If RSTTAR_.State = 0 Then grabarOrdProd = False: Exit Function
    If RSTTAR_.RecordCount = 0 Then GoTo AGREGARPERSONAL_
    RSTTAR_.MoveFirst
    While Not RSTTAR_.EOF
        RstTar.AddNew
        RstTar("idord") = xId
        RstTar("idtar") = NulosN(RSTTAR_("idtar"))
        RstTar("orden") = NulosN(RSTTAR_("orden"))
        RstTar("cantsum") = NulosN(RSTTAR_("cantsum"))
        RstTar("cantproc") = NulosN(RSTTAR_("cantproc"))
        RstTar("numop") = NulosN(RSTTAR_("numop"))
        RstTar("fchini") = RSTTAR_("fchini")
        RstTar("fchfin") = RSTTAR_("fchfin")
        RstTar("horini") = RSTTAR_("horini")
        RstTar("horfin") = RSTTAR_("horfin")
        RstTar("durtar") = NulosC(RSTTAR_("durtar"))
        RstTar("idsubresp") = NulosN(RSTTAR_("idsubresp"))
        RstTar("idarea") = NulosN(RSTTAR_("idarea"))
        RstTar("activo") = NulosN(RSTTAR_("activo"))
        RstTar.Update
        RSTTAR_.MoveNext
    Wend
AGREGARPERSONAL_:
    ' -------------------------------------PERSONAL
    If RSTPER_.State = 0 Then grabarOrdProd = False: Exit Function
    If RSTPER_.RecordCount = 0 Then GoTo AGREGARREPROCESO_
    RSTPER_.MoveFirst
    While Not RSTPER_.EOF
        RstPer.AddNew
        RstPer("idord") = xId
        RstPer("idper") = NulosN(RSTPER_("idper"))
        RstPer.Update
        RSTPER_.MoveNext
    Wend
AGREGARREPROCESO_:
    ' -------------------------------------REPROCESO
    If RSTREP_.State = 0 Then grabarOrdProd = False: Exit Function
    If RSTREP_.RecordCount = 0 Then GoTo TERMINAR_
    RSTREP_.MoveFirst
    While Not RSTREP_.EOF
        RstRep.AddNew
        RstRep("idord") = xId
        RstRep("idlotedet") = NulosN(RSTREP_("idlotedet"))
        RstRep("cantidad") = NulosN(RSTREP_("cantidad"))
        RstRep.Update
        RSTREP_.MoveNext
    Wend
TERMINAR_:

    IDORD_ = xId
    ' Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, 52, QUEHACE_, Time, Time, Date, xCon, xId
   
    xCon.CommitTrans
    MsgBox "La operación se registró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstTar = Nothing
    Set RstPer = Nothing
    Set RstRep = Nothing
    grabarOrdProd = True
    Exit Function
LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstTar = Nothing
    Set RstPer = Nothing
    Set RstRep = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    grabarOrdProd = False
End Function

Public Function caracteristicaLinea(TIPO_ As Integer, IDLINEA_ As Integer, _
                                            Optional IDTAR_ As Integer, _
                                            Optional RECORDSET_ As ADODB.Recordset) As Double
    '******************************************************************
    'TIPO:0=RENDIMIENTO DE LINEA, TIPO:1=CANTIDAD PROCESADA EN UNA HORA
    'TIPO:2=RENDIMIENTO TOTAL DE LA LINEA
    '******************************************************************
    Dim xRs As New ADODB.Recordset
    Dim RENDIMIENTO_ As Double
    Dim LIMITARLINEA_ As Boolean
    Dim cSQL As String
    
    Select Case TIPO_
        Case 0
            cSQL = "SELECT pro_lineadet.rdmto As valor" _
                + vbCr + "FROM pro_lineadet " _
                + vbCr + "WHERE (((pro_lineadet.idlineadet)=" & IDLINEA_ & ") AND ((pro_lineadet.idtar)=" & IDTAR_ & "));"
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            If xRs.State = 0 Then caracteristicaLinea = 0: Exit Function
            If xRs.RecordCount = 0 Then caracteristicaLinea = 0: Exit Function
            
            xRs.MoveFirst
            caracteristicaLinea = NulosN(xRs("valor"))
            Exit Function
                    
        Case 1
            cSQL = "SELECT pro_lineadet.kghora, pro_lineadet.numop" _
                + vbCr + "FROM pro_lineadet " _
                + vbCr + "WHERE (((pro_lineadet.idlineadet)=" & IDLINEA_ & ") AND ((pro_lineadet.idtar)=" & IDTAR_ & "));"
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            If xRs.State = 0 Then caracteristicaLinea = 0: Exit Function
            If xRs.RecordCount = 0 Then caracteristicaLinea = 0: Exit Function
            
            xRs.MoveFirst
            caracteristicaLinea = NulosN(xRs("kghora")) * NulosN(xRs("numop"))
            Exit Function
        
        Case 2
            LIMITARLINEA_ = True
            If RECORDSET_.State = 0 Then caracteristicaLinea = 1: Exit Function
            If RECORDSET_.RecordCount = 0 Then LIMITARLINEA_ = False
            
            cSQL = "SELECT pro_lineadet.idtar, pro_lineadet.rdmto " _
                + vbCr + "FROM pro_lineadet " _
                + vbCr + "WHERE (((pro_lineadet.idlineadet) = " & IDLINEA_ & "));"
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            If xRs.State = 0 Then caracteristicaLinea = 1: Exit Function
            If xRs.RecordCount = 0 Then caracteristicaLinea = 1: Exit Function
            
            RENDIMIENTO_ = 1
            While Not xRs.EOF
                If LIMITARLINEA_ Then
                    RECORDSET_.Filter = "idtar = " & NulosN(xRs("idtar")) & " AND activo=-1"
                    If RECORDSET_.RecordCount = 0 Then GoTo SIGUIENTE_
                End If
                RENDIMIENTO_ = RENDIMIENTO_ * (NulosN(xRs("rdmto")) / 100)
SIGUIENTE_:
                xRs.MoveNext
            Wend
            
            caracteristicaLinea = RENDIMIENTO_
            Exit Function
    End Select
    
    caracteristicaLinea = 0
End Function

Public Function procesarCronograma(IDORD_ As Integer, IDLINEADET_ As Integer, _
                        RECORDSET_ As ADODB.Recordset, CANTIDADAPROC_ As Double, _
                        HORINITAR_ As String, FCHINITAR_ As Date, _
                        Optional PORCENTAJEAPLICADO_ As Double = 100, _
                        Optional TIPOPROCESAMIENTO_ As Integer = 3, _
                        Optional CONSIDERARREFRIGERIO_ As Boolean = True, _
                        Optional HORINIREFRIGERIO_ As String = "13:00", _
                        Optional HORFINREFRIGERIO_ As String = "14:00", _
                        Optional MINUTOSENTRETAREAS_ As String = "00:10", _
                        Optional PORCENTAJEENTRETAREAS_ As Double = 10) As ADODB.Recordset
    
    '******************************************************************************************************
    ' TIPOPROCESAMIENTO_:0=UNA TAREA DESPUES DE OTRA. TIPOPROCESAMIENTO_:1=UNA TAREA AL PORCENTAJE DE OTRA
    ' TIPOPROCESAMIENTO_:2=UNA TAREA AL MINUTO DE LA OTRA, TIPOPROCESAMIENTO_:3=TAREA SEGUN LINEA DE PRODUCCION
    '******************************************************************************************************
    Dim DURTARNUMERIC_ As Double
    Dim DURTARCADENA_ As String
    Dim FCHINITARANT_ As Date
    Dim FCHFINTARANT_ As Date
    Dim CANTIDADSUMANT_ As Double
    Dim CANTIDADPROCANT_ As Double
    Dim HORINITARANT_ As String
    Dim HORFINTARANT_ As String
    Dim DURTARANT_ As String
    Dim valor As Variant
    Dim RECORDSETTEMP_ As New ADODB.Recordset
    Dim DURACREFRIGERIO_ As String
    Dim RENDIMIENTOTAR_ As Double
    Dim CANTPROCXHORA As Double
    Dim xRs As New ADODB.Recordset
    Dim cSQL As String
    
    Dim h() As String
    Dim tiempo As Double
    Dim intervalo As String
        
    ' -------------------------VALORES DE PROCESAMIENTO
    If TIPOPROCESAMIENTO_ = 2 Then valor = NulosC(MINUTOSENTRETAREAS_) Else valor = NulosN(PORCENTAJEENTRETAREAS_)
    ' -------------------------VALORES INICIALES
    CANTIDADPROCANT_ = CANTIDADAPROC_
    HORINITARANT_ = HORINITAR_
    HORFINTARANT_ = HORINITAR_
    DURTARANT_ = Format(CDate(HORFINTARANT_) - CDate(HORINITARANT_), "HH:mm")
    FCHINITARANT_ = FCHINITAR_
    FCHFINTARANT_ = FCHINITAR_
    DURACREFRIGERIO_ = Format(CDate(HORFINREFRIGERIO_) - CDate(HORINIREFRIGERIO_), "HH:mm")
    
    If RECORDSET_.State = 0 Then Set procesarCronograma = Nothing: Exit Function
    RECORDSET_.Filter = adFilterNone
    If RECORDSET_.RecordCount = 0 Then
        cSQL = "SELECT " & IDORD_ & " AS idord, pro_lineadet.idtar, pro_lineadet.orden, 0 AS cantsum, 0 AS cantproc, pro_lineadet.numop, '" & FCHINITAR_ & "' AS fchini, '" & FCHINITAR_ & "' AS fchfin, '" & HORINITAR_ & "' AS horini, '" & HORINITAR_ & "' AS horfin, '00:00' AS durtar, pro_emp.idemp AS idsubresp, pro_recetatar.idarea, -1 AS activo " _
            + vbCr + "FROM (((pro_lineadet LEFT JOIN pro_recetatar ON (pro_lineadet.idtar = pro_recetatar.idtar) AND (pro_lineadet.idrec = pro_recetatar.idrec)) LEFT JOIN mae_area ON pro_recetatar.idarea = mae_area.id) LEFT JOIN pro_area ON mae_area.id = pro_area.idarea) LEFT JOIN pro_emp ON pro_area.idper = pro_emp.id " _
            + vbCr + "WHERE (((pro_lineadet.idlineadet) = " & IDLINEADET_ & ")) " _
            + vbCr + "ORDER BY pro_lineadet.orden;"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        CARGAR_RST_TMP RECORDSET_, xRs

    End If
    
    DEFINIR_RST_TMP RECORDSETTEMP_, RECORDSET_
    RECORDSET_.MoveFirst
    While Not RECORDSET_.EOF
        RECORDSETTEMP_.AddNew
        If RECORDSET_("activo") = 0 Then GoTo SIGUIENTE

        ' ---------------------------------CANTIDAD PROCESADA
        RENDIMIENTOTAR_ = caracteristicaLinea(0, IDLINEADET_, NulosN(RECORDSET_("idtar")))
        If RENDIMIENTOTAR_ = 0 Then RENDIMIENTOTAR_ = 100
        CANTIDADSUMANT_ = CANTIDADPROCANT_
        RECORDSETTEMP_("cantsum") = NulosN(CANTIDADSUMANT_)
        RECORDSETTEMP_("cantproc") = (NulosN(CANTIDADSUMANT_) * ((RENDIMIENTOTAR_ / 100)))
        CANTIDADPROCANT_ = RECORDSETTEMP_("cantproc")
        
        ' ---------------------------------DURACION DE LA TAREA
        DURTARNUMERIC_ = 0
        CANTPROCXHORA = caracteristicaLinea(1, IDLINEADET_, NulosN(RECORDSET_("idtar")))
        If PORCENTAJEAPLICADO_ = 0 Then PORCENTAJEAPLICADO_ = 100
        CANTPROCXHORA = (CANTPROCXHORA * 100) / PORCENTAJEAPLICADO_
        DURTARNUMERIC_ = CANTIDADPROCANT_ / CANTPROCXHORA
        If DURTARNUMERIC_ > 24 Then DURTARNUMERIC_ = 23.9
        DURTARCADENA_ = ""
        DURTARCADENA_ = Format(Int(DURTARNUMERIC_), "00")
        DURTARCADENA_ = DURTARCADENA_ & ":" & Format(((DURTARNUMERIC_ * 60) Mod 60), "00")
        RECORDSETTEMP_("durtar") = DURTARCADENA_

        ' ---------------------------------HORA DE INICIO DE LA TAREA
        Select Case TIPOPROCESAMIENTO_
            Case 0
                RECORDSETTEMP_("horini") = HORINITARANT_
            Case 1
                If HORINITARANT_ = HORFINTARANT_ Then
                    RECORDSETTEMP_("horini") = HORINITARANT_
                Else
                    h = Split(DURTARANT_, ":")
                    tiempo = (60 * Val(h(0))) + Val(h(1))
                    tiempo = ((valor * tiempo) / 100)
                    tiempo = tiempo / 60
                    intervalo = Format(Int(tiempo), "00")
                    intervalo = intervalo & ":" & Format(((tiempo * 60) Mod 60), "00")
                    RECORDSETTEMP_("horini") = CDate(HORINITARANT_) + CDate(intervalo)
                End If
            Case 2
                If HORINITARANT_ = HORFINTARANT_ Then
                    RECORDSETTEMP_("horini") = HORINITARANT_
                Else
                    RECORDSETTEMP_("horini") = CDate(HORINITARANT_) + CDate(valor)
                End If
            Case 3
                RECORDSETTEMP_("horini") = HORINITARANT_
        End Select
        If CONSIDERARREFRIGERIO_ Then
            If (RECORDSETTEMP_("horini") > CDate(HORINIREFRIGERIO_)) And (RECORDSETTEMP_("horini") < CDate(HORFINREFRIGERIO_)) Then
                RECORDSETTEMP_("horini") = CDate(HORFINREFRIGERIO_)
            End If
        End If
        DURTARANT_ = Format(RECORDSETTEMP_("durtar"), "HH:mm")
        HORINITARANT_ = Format(RECORDSETTEMP_("horini"), "HH:mm")

        ' --------------------------------FECHA DE INICIO DE LA TAREA
        FCHINITARANT_ = CDate(Format(FCHINITARANT_, "dd/mm/yy") & " " & Format(RECORDSETTEMP_("horini"), "HH:mm"))
        RECORDSETTEMP_("fchini") = Format(FCHINITARANT_, "dd/mm/yy")

        ' --------------------------------FECHA DE FIN DE LA TAREA
        FCHFINTARANT_ = FCHINITARANT_ + CDate(RECORDSETTEMP_("durtar"))
        RECORDSETTEMP_("fchfin") = Format(FCHFINTARANT_, "dd/mm/yy")
        
        ' --------------------------------HORA DE FIN DE LA TAREA
        RECORDSETTEMP_("horfin") = Format(FCHFINTARANT_, "HH:mm")
        If CONSIDERARREFRIGERIO_ Then
            If (RECORDSETTEMP_("horfin") > CDate(HORINIREFRIGERIO_)) And (RECORDSETTEMP_("horfin") <= CDate(HORFINREFRIGERIO_)) Then
                RECORDSETTEMP_("horfin") = RECORDSETTEMP_("horfin") + CDate(DURACREFRIGERIO_)
            Else
                If (RECORDSETTEMP_("horini") <= CDate(HORINIREFRIGERIO_)) And (RECORDSETTEMP_("horfin") >= CDate(HORFINREFRIGERIO_)) Then
                    RECORDSETTEMP_("horfin") = RECORDSETTEMP_("horfin") + CDate(DURACREFRIGERIO_)
                End If
            End If
        End If
      
        RECORDSETTEMP_("activo") = NulosN(RECORDSET_("activo"))
        RECORDSETTEMP_("numop") = Format(NulosN(RECORDSET_("numop")), "00")
        RECORDSETTEMP_("idarea") = NulosN(RECORDSET_("idarea"))
        RECORDSETTEMP_("idsubresp") = NulosN(RECORDSET_("idsubresp"))
SIGUIENTE:
        RECORDSETTEMP_("idord") = NulosN(IDORD_)
        RECORDSETTEMP_("idtar") = NulosN(RECORDSET_("idtar"))
        RECORDSETTEMP_("orden") = NulosN(RECORDSET_("orden"))
        RECORDSETTEMP_.Update
        
        RECORDSET_.MoveNext
    Wend
        
    Set procesarCronograma = RECORDSETTEMP_
End Function

Public Sub seleccionarIndiceCombo(VALOR_ As Double, ByRef COMBO_ As ComboBox)
    Dim A As Integer
    
    ' Se llena el estado
    For A = 0 To COMBO_.ListCount - 1
        If COMBO_.ItemData(A) = VALOR_ Then
            COMBO_.ListIndex = A
            Exit For
        End If
    Next A
End Sub

Public Function GENERAR_SQL_ID_RST(Rst As ADODB.Recordset, nDesc As String, _
                            nCampo As String, Optional nTipoIn As String = "IN", _
                            Optional fEsNumero As Boolean = True) As String
    Dim nSQL As String
    Dim k&
    nSQL = ""
    
    If Rst.RecordCount = 0 Then Exit Function Else Rst.MoveFirst
    While Not Rst.EOF
        If Trim(CStr(Rst("" & nDesc & ""))) <> "" Then
            If fEsNumero = True Then
                nSQL = nSQL & NulosN(Rst("" & nDesc & "")) & ","
            Else
                nSQL = nSQL & "'" & NulosC(Rst("" & nDesc & "")) & "',"
            End If
        End If
        Rst.MoveNext
    Wend
    
    If nSQL <> "" Then nSQL = " " & nCampo & " " & nTipoIn & " (" + Left(nSQL, Len(nSQL) - 1) & ") "
        
    GENERAR_SQL_ID_RST = nSQL
End Function

Public Sub preparaRST(ByRef RST_ As ADODB.Recordset, TIPO_ As Integer)
    ' TIPO: 0=REGISTRO DE PLANEAMIENTO, 1=ORDEN DE PRODUCCION, 2=SOLICITUD DE MATERIALES
    ' N: Numerico
    ' D: Double
    ' F: Fecha
    ' C: Caracter
    ' L: Logico
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos() As String
    
    Select Case TIPO_
        Case 0
        Case 1
            ReDim xCampos(14, 3) As String
            
            xCampos(0, 0) = "idord":        xCampos(0, 1) = "N":      xCampos(0, 2) = ""
            xCampos(1, 0) = "idtar":        xCampos(1, 1) = "N":      xCampos(1, 2) = ""
            xCampos(2, 0) = "orden":        xCampos(2, 1) = "D":      xCampos(2, 2) = ""
            xCampos(3, 0) = "cantsum":      xCampos(3, 1) = "N":      xCampos(3, 2) = ""
            xCampos(4, 0) = "cantproc":     xCampos(4, 1) = "N":      xCampos(4, 2) = ""
            xCampos(5, 0) = "numop":        xCampos(5, 1) = "N":      xCampos(5, 2) = ""
            xCampos(6, 0) = "fchini":       xCampos(6, 1) = "F":      xCampos(6, 2) = ""
            xCampos(7, 0) = "fchfin":       xCampos(7, 1) = "F":      xCampos(7, 2) = ""
            xCampos(8, 0) = "horini":       xCampos(8, 1) = "C":      xCampos(8, 2) = "10"
            xCampos(9, 0) = "horfin":       xCampos(9, 1) = "C":      xCampos(9, 2) = "10"
            xCampos(10, 0) = "durtar":       xCampos(10, 1) = "C":      xCampos(10, 2) = "10"
            xCampos(11, 0) = "idsubresp":    xCampos(11, 1) = "N":      xCampos(11, 2) = ""
            xCampos(12, 0) = "idarea":       xCampos(12, 1) = "N":      xCampos(12, 2) = ""
            xCampos(13, 0) = "activo":       xCampos(13, 1) = "N":      xCampos(13, 2) = ""
        Case 2
            ReDim xCampos(5, 3) As String
            
            xCampos(0, 0) = "iditem":           xCampos(0, 1) = "N":      xCampos(0, 2) = ""
            xCampos(1, 0) = "idunimed":         xCampos(1, 1) = "N":      xCampos(1, 2) = ""
            xCampos(2, 0) = "cantidad":         xCampos(2, 1) = "D":      xCampos(2, 2) = ""
            xCampos(3, 0) = "idlote":           xCampos(3, 1) = "N":      xCampos(3, 2) = ""
            xCampos(4, 0) = "idlotedet":        xCampos(4, 1) = "N":      xCampos(4, 2) = ""
            
    End Select
            
    Set RST_ = xFun.CrearRstTMP(xCampos)
    RST_.Open
End Sub

Public Function grabarSolicitud(FCHDOC_ As String, IDTIPDOCREF_ As Integer, _
                                    IDDOCREF_ As Integer, IDRESP_ As Integer, _
                                    NUMDOC_ As String, IDALM_ As Integer, RSTDET_ As ADODB.Recordset, _
                                    Optional NUMSER_ As String = "0001", Optional ByRef IDSOL_ As Integer, _
                                    Optional ESTADO_ As Integer = 1, Optional ANIO_ As Integer, _
                                    Optional MES_ As Integer, Optional QUEHACE_ As Integer) As Boolean
    Dim xId As Double
    Dim xIdDet As Double
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    
On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If IDSOL_ = 0 Then
        ' Obetenemos el Id del registro
        xId = HallaCodigoTabla("pro_solicitudmat", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pro_solicitudmat", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = IDSOL_
        RST_Busq RstCab, "SELECT * FROM pro_solicitudmat WHERE id=" & xId, xCon
        xCon.Execute "DELETE * FROM pro_solicitudmatdet WHERE idsol=" & xId
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM pro_solicitudmatdet", xCon
    
    RstCab("fchdoc") = FCHDOC_
    RstCab("idtipdocref") = IDTIPDOCREF_
    RstCab("iddocref") = IDDOCREF_
    RstCab("numser") = NUMSER_
    If NUMDOC_ = "" Then
        RstCab("numdoc") = Format(hallarNumDoc("pro_solicitudmat", NUMSER_, "numser"), "0000000000")
    Else
        RstCab("numdoc") = NUMDOC_
    End If
    RstCab("idresp") = IDRESP_
    RstCab("idalm") = IDALM_
    RstCab("estado") = ESTADO_
    If ANIO_ = 0 Then RstCab("ano") = AnoTra Else RstCab("ano") = ANIO_
    If MES_ = 0 Then RstCab("idmes") = xMes Else RstCab("idmes") = MES_
    RstCab.Update
    
    xIdDet = HallaCodigoTabla("pro_solicitudmatdet", xCon, "id")
    RSTDET_.MoveFirst
    While Not RSTDET_.EOF
        RstDet.AddNew
        RstDet("id") = xIdDet
        RstDet("idsol") = xId
        RstDet("iditem") = NulosN(RSTDET_("iditem"))
        RstDet("idunimed") = NulosN(RSTDET_("idunimed"))
        RstDet("cantidad") = NulosN(RSTDET_("cantidad"))
        RstDet("idlote") = NulosN(RSTDET_("idlote"))
        RstDet("idlotedet") = NulosN(RSTDET_("idlotedet"))
        RstDet.Update
        
        xIdDet = xIdDet + 1
        RSTDET_.MoveNext
    Wend
            
    'Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, 54, QUEHACE_, Time, Time, Date, xCon, xId
   
    xCon.CommitTrans
    MsgBox "La solicitud de materiales se registró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
    grabarSolicitud = True
    Exit Function
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    grabarSolicitud = False
End Function

'*****************************************************************************************************
'* Nombre         : CargaDatosEmpresa()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : CARGA LOS DATOS DE LA EMPRESA ACTUAL Y LOS ALAMCENA EN LAS  VARIABLES PUBLICAS  YA
'*                  DEFINIDAS
'* Paranetros     :
'* Retorna        :
'*****************************************************************************************************
Sub CargaDatosEmpresa(xCon As ADODB.Connection)
    On Error Resume Next
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    CONTABILIZAR = Rst("procon")
    AnoTra = Rst("anotra")
    
    Set Rst = Nothing
    Err.Clear
End Sub

'*****************************************************************************************************
'* Nombre           : BuscarFrm
'* Tipo             : FUNCION
'* Descripcion      : Buscar si un formulario esta activo, DEVUELVE VERDADERO SI TIENE EXITO
'* Paranetros       : NOMBRE           |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    nTexto           |  String     |   puede ser el nombre(default) o el titulo del
'*                                                       frm
'*                    fBuscarNombreFrm |  Boolean    |   Indica Como va ser la busqueda True::Buscar
'*                                                       por nombre, False::Buscar por titulo
'*                    fCerrarFrm       |  Boolean    |   cerrar la ventana; false::No cerrar la
'*                                                       ventana(default)
'* Devuelve         : Boolean
'*****************************************************************************************************
Public Function BuscarFrm(nTexto As String, Optional fBuscarNombreFrm As Boolean = True, Optional fCerrarFrm As Boolean = False) As Boolean
    Dim frm As Form
    Dim fEncontrado As Boolean
    
    For Each frm In Forms
        fEncontrado = False
        If fBuscarNombreFrm = True Then
            If UCase(frm.Name) = UCase(nTexto) Then fEncontrado = True
        Else
            If UCase(frm.Caption) = UCase(nTexto) Then fEncontrado = True
        End If
        
        If fEncontrado = True Then
           If fCerrarFrm = True Then
              Unload frm
              BuscarFrm = True
              Exit Function
           Else
              On Error Resume Next
              frm.SetFocus
              BuscarFrm = True
              Err.Clear
              Exit Function
           End If
         End If
    Next
    
    BuscarFrm = False
    Err.Clear
End Function

Public Sub pActualizarCantHoras()
    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    Dim HoraFraccion As Double
    Dim Difhora As String

    nSQL = "SELECT pro_controltardet.idctr, pro_controltardet.corr, pro_controltardet.horini, pro_controltardet.horfin, pro_controltardet.cant " _
        + vbCr + " From pro_controltardet " _
        + vbCr + " WHERE (((pro_controltardet.tipo)=1)  AND ((pro_controltardet.horini) Is Not Null) AND ((pro_controltardet.horfin) Is Not Null) ); "

    RST_Busq Rst, nSQL, xCon

    If Rst.RecordCount <> 0 Then Rst.MoveFirst
    Do While Not Rst.EOF
        DoEvents
        HoraFraccion = Convert1HoraFaccion(DiferenciaHoras(Rst("horini"), Rst("horfin"), True))
        ' HoraFraccion
        xCon.Execute "update pro_controltardet set tothor = " & HoraFraccion & ", difhor = " & IIf(Difhora = "", Null, "'" & CDate(Difhora) & "'") & " where idctr = " & Rst("idctr") & " and corr = " & Rst("corr") & ";"
        Rst.MoveNext
    Loop
    Set Rst = Nothing
End Sub

Public Function hallarNumDoc(TABLA_ As String, CONDICION1 As String, CAMPO1 As String, _
                                            Optional CONDICION2 As String = "", _
                                            Optional CAMPO2 As String = "", _
                                            Optional CONDICION3 As String = "", _
                                            Optional CAMPO3 As String = "", _
                                            Optional CAMPOORDEN As String = "numdoc", _
                                            Optional FORMATO_ As String = "0000000000") As String
    Dim xRs As New ADODB.Recordset
    Dim xNum As Double
    Dim cSQL As String
    Dim nSQL As String
    
    If CONDICION2 <> "" And CAMPO2 <> "" Then
        nSQL = " AND ((" & CAMPO2 & ") = " & CONDICION2 & ")"
    End If
    
    If nSQL <> "" And CONDICION3 <> "" And CAMPO3 <> "" Then
        nSQL = nSQL & " AND ((" & CAMPO3 & ") = " & CONDICION3 & ")"
    End If
    
    cSQL = "SELECT TOP 1 * " _
        + vbCr + "FROM " & TABLA_ & " " _
        + vbCr + "WHERE ((" & CAMPO1 & ") = " & CONDICION1 & ")" _
        + vbCr + "ORDER BY " & CAMPOORDEN & " DESC"
    
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Function
    If xRs.RecordCount = 0 Then xNum = 1 Else xNum = NulosN(xRs("numdoc")) + 1
    
    hallarNumDoc = Format(xNum, FORMATO_)
    Set xRs = Nothing
End Function

