Attribute VB_Name = "Declaraciones"
'*****************************************************************************************************
'* Nombre Archivo   : DECLARACIONES.BAS
'* Tipo             : MODULO
'* Descripcion      : MODULO DONDE SE DECLARAN LAS VARIABLES PUBLICAS QUE SE UTILIZARAN EN LA CLASE
'*                    ASI COMO LA DEFINICION DE ALGUNAS FUNCIONES PROPIAS DE LA CLASE
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public xDataSource As String             ' Almacena el Ruta de la Conexion actual
Public xCon As New ADODB.Connection      ' VARIABLE QUE ALMACENARA LA CONECCION QUE SE LE PASE A LA CLASE
Public xTitulo As String                 ' ALMACENA EL TITULO PARA LOS MENSAJES QUE EMITA LA CLASE
Public xNomEmp, xNumRuc As String        ' ALMACENA EL NOMBRE Y EL RUC DE LA EMPRESA, ESTAS VARIABLES SON USADAS EN LOS REPORTES
Public CaracteresNumericos As String     ' ALMACENA CARACTERES NUMERICOS PARA VALIDAR EL EVENTO KEYPRES DE LOS CUADROS DE TEXTO
Public NomEmp As String                  ' ALMACENA EL NOMBRE DE LA EMPRESA
Public NumRUC As String                  ' ALMACENA EL RUC DE LA EMPRESA
Public CONTABILIZAR As Boolean           ' ESPECIFICA SI EL MODULO APLICARA EL MODO CONTABLE
Public xMes As Integer                   ' ALAMCENA EL MES DE TRABAJO ACTUAL
Public AnoTra As String                  ' ESPECIFICA EL AÑO DE TRABAJO ACTUAL
Public xIdEmpresa As Integer             ' ESPECIFICA EL ID DE LA EMPRESA
Public xIdUsuario As Integer             ' almacena el id del usuario
Public xDeDonde As Integer               ' almacenara el un valor para saber de donde se esta invocando a la librerias si del menu de compras o del menu de opciones
                                         ' 1 = menu compras
                                         ' 2 = menu de opciones
                                         
Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)

Public Const FORMAT_CANTIDADDECIMAL = "0.0000"
                                         

'*****************************************************************************************************
'* Nombre         : CargaDatosEmpresa()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : CARGA LOS DATOS DE LA EMPRESA ACTUAL Y LOS ALAMCENA EN LAS  VARIABLES PUBLICAS  YA
'*                  DEFINIDAS
'* Paranetros     :
'* Retorna        :
'*****************************************************************************************************
Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset

    CaracteresNumericos = "0123456789." & Chr(8)

    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    CONTABILIZAR = Rst("procon")
    AnoTra = Rst("anotra")
    Set Rst = Nothing
End Sub

Sub CargaDatosEmpresa_()
    Dim Rst As New ADODB.Recordset
    Dim F As New SistemaLogica.Funciones

    CaracteresNumericos = "0123456789." & Chr(8)

    Set Rst = F.GeneraRstSQL_("SELECT * FROM mae_empresa", xDataSource)
    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    CONTABILIZAR = Rst("procon")
    AnoTra = Rst("anotra")
    Set Rst = Nothing
End Sub

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

Function grabarMovimiento(FCHMOV_ As String, TIPDOC_ As Integer, NUMSER_ As String, GLOSA_ As String, _
                        IDRESP_ As Integer, IDPROV_ As Integer, DESPROV_ As String, _
                        IDESTADO_ As Integer, IDTIPMOV_ As Integer, IDTIPDOCREF_ As Integer, _
                        IDDOCREF_ As Integer, IDALM_ As Integer, RSTDET_ As ADODB.Recordset, _
                        Optional ByRef IDING_ As Integer = 0, Optional NUMDOC_ As String = "", _
                        Optional QUEHACE_ As Integer = 1, Optional MES_ As Integer = 0, _
                        Optional ANIO_ As Integer = 0, Optional NUMDOCREF_ As String = "") As Boolean
    
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
    Dim iDITEM_ As Integer
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
    RstCab("glosa") = GLOSA_
    RstCab("desdocref") = NUMDOCREF_
    If ANIO_ = 0 Then RstCab("ano") = AnoTra Else RstCab("ano") = ANIO_
    If MES_ = 0 Then RstCab("idmes") = xMes Else RstCab("idmes") = MES_
    RstCab.Update
    
    RSTDET_.MoveFirst
    While Not RSTDET_.EOF
        ' --------------CRITERIOS PARA CREAR LOTE
        iDITEM_ = NulosN(RSTDET_("iditem"))
        CANTIDAD_ = NulosN(RSTDET_("cantidad"))
        IDLOTE_ = NulosN(RSTDET_("idlote"))
        IDLOTEANT_ = NulosN(RSTDET_("idloteant"))
        IDLOTEDET_ = NulosN(RSTDET_("idlotedet"))
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
        RstDet("idlotedet") = crearModificarLote(MODO_, TIPO_, iDITEM_, CANTIDAD_, IDALM_, FCHMOV_, IDLOTE_, IDLOTEDET_, CANTANT_, IDLOTEANT_, IDLOTEDETANT_)
        RstDet("hora") = RSTDET_("hora")
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
    grabarMovimiento = True
    Exit Function

LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo registrar el movimiento por el siguiente motivo :" + Trim(Err.Description)
    grabarMovimiento = False
    Exit Function
End Function

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
        + vbCr + "WHERE ((" & CAMPO1 & ") = " & CONDICION1 & ")" & nSQL _
        + vbCr + "ORDER BY " & CAMPOORDEN & " DESC"
    
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Function
    If xRs.RecordCount = 0 Then xNum = 1 Else xNum = NulosN(xRs("numdoc")) + 1
    
    hallarNumDoc = Format(xNum, FORMATO_)
    Set xRs = Nothing
End Function

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

Public Function crearModificarLote(MODO_ As Integer, TIPO_ As Integer, _
                            iDITEM_ As Integer, CANTIDAD_ As Double, _
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
                LOTE_ = Format(iDITEM_, "0000") & Format(CDate(FCHING_), "yy") & Format(Month(CDate(FCHING_)), "00") & Format(Day(CDate(FCHING_)), "00")
                    
                ' Se verifica el mayor en la base de datos
                cSQL = "SELECT Max(Mid([alm_inventariolote].[descripcion],11,2)) AS orden, alm_inventariolote.iditem, alm_inventariolote.fching " _
                    + vbCr + "FROM alm_inventariolote " _
                    + vbCr + "GROUP BY alm_inventariolote.iditem, alm_inventariolote.fching " _
                    + vbCr + "HAVING (((alm_inventariolote.iditem)=" & iDITEM_ & ") AND ((alm_inventariolote.fching)=CDate('" & FCHING_ & "')))"
                
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
                RstLote("iditem") = iDITEM_
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

Public Function PrecioUni(idDocumento, IdItem As Double, DondeBuscar As String) As Double
    '===================================================================================================
    'Creado:     01/07/11 Johan Castro
    'Propósito:  Obtener el Precio unitario del registro de compras vinculado con documentos (de ingreso de almacen, Guia Remision)
    '
    'Entradas:   IdDocumento = Código de Libro
    '            IdItem = Código del Item (Producto, Materia prima, Insumo, etc)
    '            DondeBuscar = Indica el origen del registro
    '
    'Resultados: Precio unitario del item segun el documento ingresado
    '===================================================================================================
    
    Dim xRst As New ADODB.Recordset
    Dim nSQL As String
    
    If DondeBuscar = "AI" Then
        nSQL = "SELECT Sum(IIf([com_compras].[idmon]=2,[com_comprasdet].[imptot]*[con_tc].[impcom],[com_comprasdet].[imptot])) AS importetot, Sum(com_comprasdet.canpro) AS cantidadtot " _
            + vbCr + "FROM (com_compras INNER JOIN (com_comprasdet INNER JOIN alm_ingresodoc ON com_comprasdet.idcom = alm_ingresodoc.iddoc) ON com_compras.id = com_comprasdet.idcom) INNER JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
            + vbCr + "WHERE (((alm_ingresodoc.id)=" & idDocumento & ") AND ((com_comprasdet.iditem)=" & IdItem & "));"
    
'        nSQL = "SELECT Avg(com_comprasdet.preuni) AS preuniprom " _
'            + vbCr + " FROM com_comprasdet INNER JOIN alm_ingresodoc ON com_comprasdet.idcom = alm_ingresodoc.iddoc " _
'            + vbCr + " GROUP BY alm_ingresodoc.id, com_comprasdet.iditem " _
'            + vbCr + " HAVING (((alm_ingresodoc.id)=" & IdDocumento & ") AND ((com_comprasdet.iditem)=" & IdItem & "))"

        Set xRst = Nothing
        RST_Busq xRst, nSQL, xCon
        If xRst.State = 0 Then PrecioUni = 0: Exit Function
        If xRst.RecordCount = 0 Then PrecioUni = 0: Exit Function
        
        If NulosN(xRst("cantidadtot")) = 0 Then
            PrecioUni = 0
        Else
            PrecioUni = NulosN(xRst("importetot")) / NulosN(xRst("cantidadtot"))
        End If
        Exit Function
        
    ElseIf DondeBuscar = "GR" Then
        nSQL = "SELECT vta_guia.id, vta_ventasdet.iditem, Avg(vta_ventasdet.preuni) AS preuniprom " _
            + vbCr + " FROM vta_guia INNER JOIN vta_ventasdet ON vta_guia.iddocven = vta_ventasdet.idvta " _
            + vbCr + " GROUP BY vta_guia.id, vta_ventasdet.iditem " _
            + vbCr + " HAVING (((vta_guia.id)=" & idDocumento & ") AND ((vta_ventasdet.iditem)=" & IdItem & ")); "
       
        Set xRst = Nothing
        RST_Busq xRst, nSQL, xCon
        If xRst.State = 0 Then PrecioUni = 0: Exit Function
        If xRst.RecordCount = 0 Then PrecioUni = 0: Exit Function
        
        PrecioUni = NulosN(xRst("preuniprom"))
        Exit Function
    Else
        PrecioUni = 0
        Exit Function
    End If
End Function

Public Function pHallarPrecioInicial(IDTEM_ As Integer, FECHA_ As String, ANIOTRABAJO_ As Integer) As Double
    Dim xRs As New ADODB.Recordset
    Dim cSQL As String
    Dim SALDOINICIAL_ As Double
    Dim INGRESOCANTIDAD_ As Double
    Dim INGRESOIMPORTE_ As Double
    Dim SALIDACANTIDAD_ As Double
    Dim SALIDAIMPORTE_ As Double
    Dim PRECIOPROMEDIO_ As Double
    
    ' SI ES EL PRIMER DIA DEL AÑO
    If CDate(FECHA_) = CDate("01/01/" & ANIOTRABAJO_) Then
        pHallarPrecioInicial = NulosN(Busca_Codigo("id", NulosC(IDTEM_), "preini", "alm_inventario", "N", xCon))
    Else
        pHallarDatosMovimientos IDTEM_, "01/01/" & ANIOTRABAJO_, CDate(FECHA_), SALDOINICIAL_, INGRESOCANTIDAD_, INGRESOIMPORTE_, SALIDACANTIDAD_, SALIDAIMPORTE_, PRECIOPROMEDIO_
        If INGRESOCANTIDAD_ = SALIDACANTIDAD_ And SALDOINICIAL_ = 0 Then
            pHallarPrecioInicial = 0: Exit Function
        Else
            pHallarPrecioInicial = PRECIOPROMEDIO_: Exit Function
        End If
    End If
End Function

Function KardexMovimientoSQL(xIdItem As Double, xFchIni As Date, xFchFin As Date) As String
     '===================================================================================================
    'Creado : 19/01/12 Por Johan Castro
    'Propósito: Generar Sentencia SQL para mostrar los movimientos al detalle de un item en funcion al periodo indicado
    '
    'Entradas:  xIdItem= Código del item(ver tabla mae_inventario.id)
    '           xFchIni= Fecha Inicial
    '           xFchFin= Fecha Final
    '
    'Resultados: Sentencia SQL listo para ejecutar
    '
    'Nota:  Descripción de los estados (ver tabla mae_estados)
    '--1=Pendiente, 2=Procesado, 3=Aprobado,4=Rechazado,5=Anulado
    '
    'Modificado: 21/01/12 Johan Castro; Se extrae la Sentencia SQL de evento FrmVerKardex.MuestraKardexProm y FrmResuMov.HallarPrecioPromedio que son lo mismo y se centraliza
    '           No mostrar registros que esten con el siguiente estado(1=Pendiente, 4=Rechazado y 5=Anulado)
    '           Agregar filtro en Almacen Ingreso - Salida "AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 "
    '           Agregar filtro en Produccion Insumos, "AND pro_producciondet.estado Not In (1,4,5) AND pro_producciondetins.canutil<>0 AND alm_inventario.tippro=3"
    '           Agregar filtro en Produccion Producto Terminado, "AND pro_producciondet.estado Not In (1,4,5)"
    '           Agregar Variable xSQLFiltroPS para mostrar solo materia prima a partir del 2012 en adelante en Produccion Insumos

    '===================================================================================================

    ' AI = Almacen Ingreso
    ' AS = Almacen Salida
    ' C =  Compras
    ' SM = SOLICUTID DE MATERIALES
    ' PP = PARTE DE PRODUCCION
    'GR = GUIAS DE REMISION
    'PS =
    
    '--ALMACEN INGRESO
    '--ALMACEN SALIDA
    '--COMPRAS
    '--GUIA REMISION => SALIDA
    '--PRODUCCION DETALLE INSUMOS,PRODUCTOS INTERMEDIOS, MATERIA PRIMA => SALIDA
    '--PRODUCCION PRODUCTOS TERMINADOS => SALIDA
    '--VENTAS => SALIDA
    '--VENTAS => DEVOLUCIONES INGRESO

    Dim xCadSQL As String
    Dim xSQLFiltroPS As String '--Util para aplicar un filtro adicional que mostrará solo materia prima en sentencia de "produccion insumos salida"

    If NulosN(AnoTra) >= 2012 Then
        '--Aplicar filtro en produccion de salida para mostrar solo materia prima del 2012 en adelante
        xSQLFiltroPS = " AND alm_inventario.tippro=3  "
    End If

    ' PREPARAMOS LA SELECT PARA ARMAR EL KARDEX
    xCadSQL = "SELECT alm_ingreso.id, alm_ingresodet.iditem, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, alm_ingreso.numser, alm_ingreso.numdoc, " _
                & " alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AI' AS tipo, alm_ingreso.nombre AS entidad, 0 AS aa, " _
                & " (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos ,'Almacén' & iif(cstr(numdocumentos) <>'0', ' - Compras','') as modulo, '' AS registro, '' AS ctanum, '' AS ctanom, IIf([alm_ingreso].[idope]=1,'RECEPCION',IIf([alm_ingreso].[idope]=2,'DESPACHO',IIf([alm_ingreso].[idope]=3,'ENTRADA PRODUCCION',IIf([alm_ingreso].[idope]=4,'SALIDA PRODUCCION','')))) AS desope, alm_ingresodet.hora AS horini, alm_ingresodet.hora AS horfin " _
        + vbCr + " FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id " _
        + vbCr + " WHERE (((alm_ingresodet.iditem)=" & xIdItem & ") AND ((alm_ingreso.fching)>=CDate('" & xFchIni & "') " _
                & " And (alm_ingreso.fching)<=CDate('" & xFchFin & "')) AND ((alm_ingreso.tipmov)=-1)) " _
                & " AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " _
        + vbCr + " UNION ALL " _
        + vbCr + " SELECT alm_ingreso.id, alm_ingresodet.iditem, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, alm_ingreso.numser, alm_ingreso.numdoc, " _
                & " alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AS' AS tipo, alm_ingreso.nombre AS entidad, 0 AS aa, " _
                & " (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos ,'Almacén' & iif(cstr(numdocumentos) <>'0', ' - Compras','') as modulo, '' AS registro,'' AS ctanum, '' AS ctanom, IIf([alm_ingreso].[idope]=1,'RECEPCION',IIf([alm_ingreso].[idope]=2,'DESPACHO',IIf([alm_ingreso].[idope]=3,'ENTRADA PRODUCCION',IIf([alm_ingreso].[idope]=4,'SALIDA PRODUCCION','')))) AS desope, alm_ingresodet.hora AS horini, alm_ingresodet.hora AS horfin  " _
        + vbCr + " FROM (alm_ingreso LEFT JOIN (alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id " _
        + vbCr + " WHERE (((alm_ingresodet.iditem)=" & xIdItem & ") AND ((alm_ingreso.fching)>=CDate('" & xFchIni & "') " _
                & " And (alm_ingreso.fching)<=CDate('" & xFchFin & "')) AND ((alm_ingreso.tipmov)=0)) " _
                & " AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " _
        + vbCr + " UNION ALL " _
        + vbCr + " SELECT com_compras.id, com_comprasdet.iditem, alm_inventario.descripcion, com_compras.fchdoc, com_compras.numser, com_compras.numdoc, " _
                & " com_comprasdet.canpro, IIf([com_compras]![idmon]=2,[com_comprasdet]![preuni]*[con_tc]![impcom],[com_comprasdet]![preuni]) AS preuni, mae_documento.abrev AS descdoc, " _
                & " 'C' AS Tipo, mae_prov.nombre AS entidad, 0 AS aa, 0 AS numdocumentos,'Compras' as modulo,com_compras.numreg AS registro,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS desope, '' AS horini, '' AS horfin " _
        + vbCr + " FROM (alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc)  " _
                & " LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
        + vbCr + " WHERE (((com_comprasdet.iditem)=" & xIdItem & ") AND " _
                & " ((com_compras.fchdoc)>=CDate('" & xFchIni & "') And (com_compras.fchdoc)<=CDate('" & xFchFin & "')) AND ((com_compras.tipcom)=1))"

    xCadSQL = xCadSQL _
        + vbCr + "  Union All" _
        + vbCr + " SELECT vta_guia.id, vta_guiadet.iditem, alm_inventario.descripcion, vta_guia.fecgiro, vta_guia.numser, vta_guia.numdoc, vta_guiadet.canpro, " _
                & " 0 AS preuni, mae_documento.abrev AS desdoc, 'GR' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, IIf([vta_guia]![iddocven]<>0,1,0) AS numdocumentos,'Guia de Remisión' as modulo, '' AS registro,'' AS ctanum, '' AS ctanom, '' AS desope, '' AS horini, '' AS horfin  " _
        + vbCr + " FROM ((mae_cliente RIGHT JOIN vta_guia ON mae_cliente.id = vta_guia.idcli) LEFT JOIN mae_documento ON vta_guia.tipdoc = mae_documento.id) LEFT JOIN (vta_guiadet " _
                & " LEFT JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id) ON vta_guia.id = vta_guiadet.idgui " _
        + vbCr + " WHERE (((vta_guiadet.iditem)=" & xIdItem & ") " _
                & " AND ((vta_guia.fecgiro)>=CDate('" & xFchIni & "') And (vta_guia.fecgiro)<=CDate('" & xFchFin & "'))) " _
        + vbCr + " Union All " _
        + vbCr + " SELECT pro_produccion.id, pro_producciondetins.iditem, alm_inventario.descripcion, pro_produccion.dia, '' As numser, pro_producciondetins.numparte As numdoc, pro_producciondetins.canutil, " _
                & " 0 AS preuni, 'SM' AS desdoc, 'PS' AS tipo, [alm_inventario_1].[descripcion] AS entidad, pro_producciondet.iditem AS aa, 0 AS numdocumentos,'Producción' as modulo, '' AS registro,'' AS ctanum, '' AS ctanom, '' AS desope, pro_producciondet.horini, pro_producciondet.horfin  " _
        + vbCr + " FROM (((pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN (pro_producciondetins LEFT JOIN alm_inventario ON pro_producciondetins.iditem = alm_inventario.id) ON (pro_producciondet.idrec = pro_producciondetins.idrec) " _
                & " AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro)) LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_receta.iditem = alm_inventario_1.id " _
        + vbCr + " WHERE (((pro_producciondetins.iditem)=" & xIdItem & ") AND ((pro_produccion.dia)>=CDate('" & xFchIni & "') And (pro_produccion.dia)<=CDate('" & xFchFin & "')))" _
                & " AND pro_producciondet.estado Not In (1,4,5) AND pro_producciondetins.canutil<>0 " & xSQLFiltroPS _
        + vbCr + " Union All " _
        + vbCr & " SELECT pro_produccion.id, pro_producciondet.iditem, alm_inventario.descripcion, pro_produccion.dia, '' As numser, pro_producciondet.numparte As numdoc, pro_producciondet.cantidad, " _
                & " 0 AS preuni, 'PP' AS desdoc, 'P' AS tipo, 'Producción' AS entidad, pro_producciondet.iditem AS aa, 0 AS numdocumentos ,'Producción' as modulo, '' as registro,'' AS ctanum, '' AS ctanom, '' AS desope, pro_producciondet.horini, pro_producciondet.horfin  " _
        + vbCr & " FROM pro_produccion INNER JOIN (pro_producciondet LEFT JOIN alm_inventario ON pro_producciondet.iditem = alm_inventario.id) ON pro_produccion.id = pro_producciondet.idpro " _
        + vbCr & " WHERE (((pro_producciondet.iditem)=" & xIdItem & ") AND ((pro_produccion.dia)>=CDate('" & xFchIni & "') And (pro_produccion.dia)<=CDate('" & xFchFin & "'))) " _
               & " AND pro_producciondet.estado Not In (1,4,5) AND pro_producciondet.cantidad<>0 "

    xCadSQL = xCadSQL + " UNION All " _
        + vbCr + " SELECT vta_ventas.id, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.fchdoc, vta_ventas.numser, vta_ventas.numdoc, " _
                    & " vta_ventasdet.canpro, vta_ventasdet.preuni, mae_documento.abrev AS descdoc, 'V' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, " _
                    & " 'Ventas' as modulo, vta_ventas.numreg AS registro,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS desope, '' AS horini, '' AS horfin " _
        + vbCr + " FROM ((mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN (vta_ventasdet  " _
                    & " LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
        + vbCr + " WHERE (((vta_ventasdet.iditem)=" & xIdItem & ") " _
                    & " AND ((vta_ventas.fchdoc)>=CDate('" & xFchIni & "') And (vta_ventas.fchdoc)<=CDate('" & xFchFin & "')) AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) " _
                    & " AND ((vta_ventas.iddocref) Is Null Or (vta_ventas.iddocref)=0) )" _
        + vbCr + " UNION All " _
        + vbCr + " SELECT vta_ventas.id, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.fchdoc, vta_ventas.numser, vta_ventas.numdoc, " _
                    & " vta_ventasdet.canpro, vta_ventasdet.preuni, mae_documento.abrev AS descdoc, 'V' AS tipo, mae_cliente.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, " _
                    & " 'Ventas NC' AS modulo, vta_ventas.numreg AS registro,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS desope, '' AS horini, '' AS horfin " _
        + vbCr + " FROM ((mae_cliente RIGHT JOIN (vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_cliente.id = vta_ventas.idcli) RIGHT JOIN (vta_ventasdet " _
                    & " LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id " _
        + vbCr + " WHERE (((vta_ventasdet.iditem)=" & xIdItem & ") AND ((vta_ventas.fchdoc)>=CDate('" & xFchIni & "') And (vta_ventas.fchdoc)<=CDate('" & xFchFin & "')) " _
                    & " AND ((vta_ventas.oriitem)=1 Or (vta_ventas.oriitem)=3) AND ((vta_ventas.iddocref)<>0) AND ((vta_ventas.idmotnotcre)=4))"

    KardexMovimientoSQL = xCadSQL

End Function

Public Sub pHallarDatosMovimientos(iDITEM_ As Integer, FCHINI_ As Date, FCHFIN_ As Date, ByRef SALDOINICIAL_ As Double, ByRef INGRESOCANTIDAD_ As Double, ByRef INGRESOIMPORTE_ As Double, _
                            ByRef SALIDACANTIDAD_ As Double, ByRef SALIDAIMPORTE_ As Double, ByRef PRECIOPROMEDIO_ As Double)
    Dim xCadSQL As String
    Dim UltPreCosto As Double
    Dim mInicioGrupo As Long '--indica la fila inicial de un grupo, cambia cuando cambia de item
    Dim xPrecioUni As Double '--Indica el precio unitario de cada registro
    Dim xPrecioIni As Double
    Dim StockIni As Double
    Dim xRs As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    Dim TIPOPRODUCTO_ As Integer
    Dim xSaldo As Double
    Dim xSaldoImp As Double
    Dim A&
    Dim xFila As Integer
    Dim xTotSal, xTotEnt As Double
    Dim xImpSal, xImpEnt As Double
    Dim cSQL As String
    
    ' AI = Almacen Ingreso
    ' AS = Almacen Salida
    ' C =  Compras
    ' SM = SOLICUTID DE MATERIALES
    ' PP = PARTE DE PRODUCCION
    ' GR = GUIAS DE REMISION
    ' PS =
    
    xTotSal = 0
    xImpSal = 0
    xTotEnt = 0
    xImpEnt = 0
    xPrecioUni = 0
    
    '--Generar la consulta SQL para obtener el detalle de movimientos del kardex
    xCadSQL = KardexMovimientoSQL(CDbl(iDITEM_), FCHINI_, FCHFIN_)
            
    RST_Busq xRs, xCadSQL, xCon
    xRs.Sort = "fchdoc, Tipo, numdoc"
                
    '--obtener el saldo inicial
    If CDate(FCHINI_) <> CDate("01/01/" & AnoTra) Then
        StockIni = SaldoActual(CDbl(iDITEM_), NulosC("01/01/" & AnoTra), NulosC(CDate(FCHINI_) - 1), xCon)
        If StockIni = 0 Then
            xPrecioIni = 0
        Else
            xPrecioIni = pHallarPrecioInicial(iDITEM_, NulosC(FCHINI_), CInt(AnoTra))
        End If
    Else
        StockIni = NulosN(Busca_Codigo("id", NulosC(iDITEM_), "stckini", "alm_inventario", "N", xCon))
        xPrecioIni = NulosN(Busca_Codigo("id", NulosC(iDITEM_), "preini", "alm_inventario", "N", xCon))
    End If
                
    xSaldo = StockIni
    xSaldoImp = xSaldo * xPrecioIni
    xPrecioUni = xPrecioIni
    'xTotEnt = xTotEnt + StockIni
    
    If xRs.RecordCount = 0 Then GoTo SALIR_
    xRs.MoveFirst

    While Not xRs.EOF
        ' ----------------------------------------------INGRESOS
        If xRs("tipo") = "C" Or xRs("tipo") = "AI" Or xRs("tipo") = "P" Then
            If xRs("descdoc") = "NC" Then
                xSaldo = xSaldo - NulosN(xRs("canpro"))
                xTotSal = xTotSal + NulosN(xRs("canpro"))
            Else
                xSaldo = xSaldo + NulosN(xRs("canpro"))
                xTotEnt = xTotEnt + NulosN(xRs("canpro"))
            End If
            
            '--obtener el precio
            If xRs("tipo") = "AI" And xRs("numdocumentos") <> 0 Then
                xPrecioUni = PrecioUni(xRs("id"), CDbl(iDITEM_), NulosC(xRs("tipo")))
            Else
                xPrecioUni = NulosN(xRs("preuni"))
            End If
            
            If xRs("tipo") = "P" Then
                xPrecioUni = 0
                
                cSQL = "SELECT con_librocostodet.impmprima, con_librocostodet.impmanobr, con_librocostodet.impgasfab, con_librocostodet.cantidad " _
                    + vbCr + "FROM con_librocostodet " _
                    + vbCr + "WHERE (((con_librocostodet.idprod)=" & NulosN(xRs("id")) & ") AND ((con_librocostodet.iditem)=" & iDITEM_ & "));"
                
                Set xRsAux = Nothing
                RST_Busq xRsAux, cSQL, xCon
                
                If xRsAux.State = 0 Then Exit Sub
                If xRsAux.RecordCount > 0 Then
                    xPrecioUni = (NulosN(xRsAux("impmprima")) + NulosN(xRsAux("impmanobr")) + NulosN(xRsAux("impgasfab"))) / NulosN(xRs("canpro"))
'                        xPrecioUni = (NulosN(xRs("impmprima")) + NulosN(xRs("impmanobr"))) / NulosN(xRs("cantidad"))
                End If
            End If
                             
            If xRs("descdoc") = "NC" Then
                xSaldoImp = xSaldoImp - (NulosN(xRs("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(xRs("canpro")) * xPrecioUni)
            Else
                xSaldoImp = xSaldoImp + (NulosN(xRs("canpro")) * xPrecioUni)
                xImpEnt = xImpEnt + (NulosN(xRs("canpro")) * xPrecioUni)
            End If
                                
            UltPreCosto = xPrecioUni
        ' ----------------------------------------------------------SALIDAS
        Else
            If xSaldo = 0 Then
                xPrecioUni = 0
            Else
                xPrecioUni = xSaldoImp / xSaldo
            End If
            UltPreCosto = xPrecioUni
            
            If xRs("descdoc") = "NC" Then
                xSaldo = xSaldo + NulosN(xRs("canpro"))
                xTotEnt = xTotEnt + NulosN(xRs("canpro"))
            Else
                xSaldo = xSaldo - NulosN(xRs("canpro"))
                xTotSal = xTotSal + NulosN(xRs("canpro"))
            End If
                                        
            If xRs("descdoc") = "NC" Then
                xSaldoImp = xSaldoImp + (NulosN(xRs("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(xRs("canpro")) * xPrecioUni)
            Else
                xSaldoImp = xSaldoImp - (NulosN(xRs("canpro")) * xPrecioUni)
                xImpSal = xImpSal + (NulosN(xRs("canpro")) * xPrecioUni)
            End If
        End If
        
        xRs.MoveNext
    Wend

SALIR_:
    SALDOINICIAL_ = StockIni
    SALIDACANTIDAD_ = xTotSal
    SALIDAIMPORTE_ = xImpSal
    INGRESOCANTIDAD_ = xTotEnt
    INGRESOIMPORTE_ = xImpEnt
    PRECIOPROMEDIO_ = xPrecioUni
End Sub
