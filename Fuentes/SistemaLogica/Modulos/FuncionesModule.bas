Attribute VB_Name = "FuncionesModule"
Option Explicit

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function mGetMachineName() As String
    Dim dwLen As Long
    Dim strString As String
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    'Show the computer name
    mGetMachineName = strString
End Function

Public Function mGeneraRstSQL(CadSQL As String, _
                                Optional Cnn As ADODB.Connection) As ADODB.Recordset
    Dim mDataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
    If Not Cnn Is Nothing Then Set mDataBase.Connection = Cnn
    mDataBase.ClearParameter
    mDataBase.CommandText = CadSQL
    Set mRecord = mDataBase.GetRecordset
    Set mDataBase = Nothing
    
    Set mGeneraRstSQL = mRecord
End Function

Public Function mGenerarSQLInGrid(GRID As Object, _
                                mCol As Integer, _
                                nCampo As String, _
                                Optional ArmaSQL As Boolean = True, _
                                Optional nTipoIn As String = "IN", _
                                Optional fEsNumero As Boolean = True) As String
            
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
    
    If nSQL <> "" Then
        If ArmaSQL Then
            nSQL = " " & nCampo & " " & nTipoIn & " (" + Left(nSQL, Len(nSQL) - 1) & ") "
        Else
            nSQL = Left(nSQL, Len(nSQL) - 1)
        End If
    End If
    mGenerarSQLInGrid = nSQL
End Function

Public Function mGenerarSQLInRst(Rst As ADODB.Recordset, _
                                nDesc As String, _
                                nCampo As String, _
                                Optional ArmaSQL As Boolean = True, _
                                Optional nTipoIn As String = "IN", _
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
    
    If nSQL <> "" Then
        If ArmaSQL Then
            nSQL = " " & nCampo & " " & nTipoIn & " (" + Left(nSQL, Len(nSQL) - 1) & ") "
        Else
            nSQL = Left(nSQL, Len(nSQL) - 1)
        End If
    End If
        
    mGenerarSQLInRst = nSQL
End Function

Function mCentrarFrame(mFrame As Frame, mForm As Form) As Boolean
    On Error GoTo ERROR
    If mForm.WindowState <> 2 Then
        mFrame.Top = (mForm.ScaleHeight / 2) - (mFrame.Height / 2)
        mFrame.Left = (mForm.ScaleWidth / 2) - (mFrame.Width / 2)
    End If
    mCentrarFrame = True
    Exit Function
ERROR:
    mCentrarFrame = False
    Exit Function
End Function

Function mNuloNumeric(value As Variant) As Double
    If Trim(value) = "" Or IsNull(value) Then
        mNuloNumeric = 0
    Else
        'Si el valor no es nulo retorna el valor original
        If IsNumeric(value) Then
            mNuloNumeric = mConvertDBL(value)
        Else
            mNuloNumeric = 0
        End If
    End If
End Function

Function mConvertDBL(texto)
    Dim retval, SepDecimal, SepMiles
    If texto = "" Then texto = "0"
    retval = Replace(Trim(texto), " ", "")
    If CDbl("3,24") = 324 Then
        SepDecimal = "."
        SepMiles = ","
    Else
        SepDecimal = ","
        SepMiles = "."
    End If
    If InStr(retval, SepDecimal) > 0 Then
        If InStr(retval, SepMiles) > 0 Then
            If InStr(retval, SepDecimal) > InStr(retval, SepMiles) Then
                retval = Replace(retval, SepMiles, "")
            Else
                retval = Replace(retval, SepDecimal, "")
                retval = Replace(retval, SepMiles, SepDecimal)
            End If
        End If
    Else
        retval = Replace(retval, SepMiles, SepDecimal)
    End If
    mConvertDBL = CDbl(retval)
End Function

Function mNuloString(value As Variant) As String
    If Trim(value) = "" Or IsNull(value) Then
        mNuloString = ""
    Else
        'Si el valor no es nulo retorna el valor original
        If IsNumeric(value) Then
            mNuloString = Trim(value)
        Else
            mNuloString = Trim(value)
        End If
    End If
End Function

Function mStringFormat(value As Variant, formato As String) As String
    mStringFormat = Format(mNuloString(value), formato)
End Function

Function mConvertirFechaANumero(value As Date) As Double
    mConvertirFechaANumero = mNuloNumeric(mStringFormat(Year(value), "0000") & mStringFormat(Month(value), "00") & mStringFormat(Day(value), "00"))
End Function

Function mConvertirHoraANumero(value As Date) As Double
    mConvertirHoraANumero = mNuloNumeric(mStringFormat(Hour(value), "00") & mStringFormat(Minute(value), "00") & mStringFormat(Second(value), "00"))
End Function

Function mConvertirNumeroAFecha(value As Double) As Date
    Dim Anho As Long
    Dim Mes As Long
    Dim Dia As Long
    
    Anho = mNuloNumeric(Mid(mNuloString(value), 1, 4))
    Mes = mNuloNumeric(Mid(mNuloString(value), 5, 2))
    Dia = mNuloNumeric(Mid(mNuloString(value), 7, 2))
    
    mConvertirNumeroAFecha = DateSerial(Anho, Mes, Dia)
End Function

Function mHallaInsumosParteProduccion(IdParteProduccion As Long, Optional mConexion As ADODB.Connection) As ADODB.Recordset
    Dim cSQL As String
    Dim cSQLInterna As String
    Dim mIdTipoOrdProduccion As Long
    Dim database As New SistemaData.EDataBase
    Dim xRs As New ADODB.Recordset
        
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    mIdTipoOrdProduccion = mNuloNumeric(mKeyValue("OrdenProduccion", mConexion))
    
    cSQLInterna = "SELECT pro_ordenprod.id As idord, pro_solicitudmat.id AS idsol, alm_ingresodet.id As idmov, alm_ingresodet.iditem, Sum(alm_ingresodet.cantteo) AS canteoing, 0 AS canteosal, Sum(alm_ingresodet.cantidad) AS canrealing, 0 AS canrealsal " _
        + vbCr + "FROM ((pro_ordenprod INNER JOIN pro_solicitudmat ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_ingreso ON pro_solicitudmat.id = alm_ingreso.iddocref) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id " _
        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((alm_ingreso.tipmov) = True)) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat.id, alm_ingresodet.id, alm_ingresodet.iditem, 0; " _
        + vbCr + "UNION " _
        + vbCr + "SELECT pro_ordenprod.id As idord, pro_solicitudmat_1.id AS idsol, alm_ingresodet.id AS idmov, alm_ingresodet.iditem, Sum(alm_ingresodet.cantteo) AS canteoing, 0 AS canteosal, Sum(alm_ingresodet.cantidad) AS canrealing, 0 AS canrealsal " _
        + vbCr + "FROM pro_ordenprod INNER JOIN (pro_solicitudmat INNER JOIN (pro_solicitudmat AS pro_solicitudmat_1 INNER JOIN (alm_ingreso INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) ON pro_solicitudmat_1.id = alm_ingreso.iddocref) ON pro_solicitudmat.id = pro_solicitudmat_1.iddocref) ON pro_ordenprod.id = pro_solicitudmat.iddocref " _
        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((alm_ingreso.tipmov) = True)) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat_1.id, alm_ingresodet.id, alm_ingresodet.iditem, 0; " _
        + vbCr + "UNION " _
        + vbCr + "SELECT pro_ordenprod.id As idord, pro_solicitudmat.id AS idsol, alm_ingresodet.id As idmov, alm_ingresodet.iditem, 0 AS canteoing, Sum(alm_ingresodet.cantteo) AS canteosal, 0 AS canrealing, Sum(alm_ingresodet.cantidad) AS canrealsal " _
        + vbCr + "FROM ((pro_ordenprod INNER JOIN pro_solicitudmat ON pro_ordenprod.id = pro_solicitudmat.iddocref) INNER JOIN alm_ingreso ON pro_solicitudmat.id = alm_ingreso.iddocref) INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id " _
        + vbCr + "Where (((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((alm_ingreso.tipmov) = False)) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat.id, alm_ingresodet.id, alm_ingresodet.iditem, 0, 0; " _
        + vbCr + "UNION " _
        + vbCr + "SELECT pro_ordenprod.id As idord, pro_solicitudmat_1.id AS idsol, alm_ingresodet.id As idmov, alm_ingresodet.iditem, 0 AS canteoing, Sum(alm_ingresodet.cantteo) AS canteosal, 0 AS canrealing, Sum(alm_ingresodet.cantidad) AS canrealsal " _
        + vbCr + "FROM pro_ordenprod INNER JOIN (pro_solicitudmat INNER JOIN (pro_solicitudmat AS pro_solicitudmat_1 INNER JOIN (alm_ingreso INNER JOIN alm_ingresodet ON alm_ingreso.id = alm_ingresodet.id) ON pro_solicitudmat_1.id = alm_ingreso.iddocref) ON pro_solicitudmat.id = pro_solicitudmat_1.iddocref) ON pro_ordenprod.id = pro_solicitudmat.iddocref " _
        + vbCr + "WHERE (((pro_solicitudmat.idtipdocref) = " & mIdTipoOrdProduccion & ") And ((alm_ingreso.tipmov) = False)) " _
        + vbCr + "GROUP BY pro_ordenprod.id, pro_solicitudmat_1.id, alm_ingresodet.id, alm_ingresodet.iditem, 0, 0, 0;"
    
    cSQL = "SELECT pro_produccion.id AS idprod, pro_producciondet.idproddet, pro_produccion.fchdoc AS fecha, pro_producciondet.idrec, pro_producciondet.iditem AS iditemprod, pro_producciondet.idunimed, pro_producciondet.canprog AS canteoprod, pro_producciondet.cantidad AS canrealprod, conmov.iditem, Sum(conmov.canteoing) AS canteoing, Sum(conmov.canteosal) AS canteosal, Sum(conmov.canrealing) AS canrealing, Sum(conmov.canrealsal) AS canrealsal, Sum([conmov].[canteosal]-[conmov].[canteoing]) AS canteomov, Sum([conmov].[canrealsal]-[conmov].[canrealing]) AS canrealmov " _
        + vbCr + "FROM (pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN " _
        + vbCr + "( " _
        + vbCr + cSQLInterna _
        + vbCr + ") " _
        + vbCr + "AS conmov ON pro_producciondet.idord = conmov.idord " _
        + vbCr + "WHERE (((pro_produccion.id)=" & IdParteProduccion & ")) " _
        + vbCr + "GROUP BY pro_produccion.id, pro_producciondet.idproddet, pro_produccion.fchdoc, pro_producciondet.idrec, pro_producciondet.iditem, pro_producciondet.idunimed, pro_producciondet.canprog, pro_producciondet.cantidad, conmov.iditem"
        
    database.CommandText = cSQL
    Set xRs = database.GetRecordset
    If xRs.State = 0 Then Err.Raise &HFFFFFF01, , "[HallaInsumosParteProduccion] El recordset se encuentra en un estado incoherente" + Trim(Err.Description)
    
    Set mHallaInsumosParteProduccion = xRs
End Function

Public Function mHallaNumeroDocumento(nombreTabla As String, condicion1 As String, _
                                            campo1 As String, _
                                            Optional mConexion As ADODB.Connection, _
                                            Optional condicion2 As String = "", _
                                            Optional campo2 As String = "", _
                                            Optional condicion3 As String = "", _
                                            Optional campo3 As String = "", _
                                            Optional campoOrden As String = "numdoc", _
                                            Optional formato As String = "0000000000") As String
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim database As New SistemaData.EDataBase
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    
    If condicion2 <> "" And campo2 <> "" Then
        nSQL = " AND ((" & campo2 & ") = " & condicion2 & ")"
    End If
    
    If nSQL <> "" And condicion3 <> "" And campo3 <> "" Then
        nSQL = nSQL & " AND ((" & campo3 & ") = " & condicion3 & ")"
    End If
    
    database.CommandText = "SELECT TOP 1 * " _
        + vbCr + "FROM " & nombreTabla & " " _
        + vbCr + "WHERE ((" & campo1 & ") = " & condicion1 & ")" & nSQL _
        + vbCr + "ORDER BY " & campoOrden & " DESC"
    
    Set xRs = database.GetRecordset()
    
    If xRs.State = 0 Then Exit Function
    If xRs.RecordCount = 0 Then
        mHallaNumeroDocumento = 1
    Else
        mHallaNumeroDocumento = NulosN(xRs("numdoc")) + 1
    End If
    
    mHallaNumeroDocumento = Format(mHallaNumeroDocumento, formato)
    Set xRs = Nothing
    Set database = Nothing
End Function

Public Function mKeyValue(key As String, Optional mConexion As ADODB.Connection) As String
    Dim xRs As New ADODB.Recordset
    Dim xNum As Double
    Dim cSQL As String
    Dim nSQL As String
    Dim database As New SistemaData.EDataBase
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
        
    database.CommandText = "SELECT TOP 1 * " _
                    + vbCr + "FROM mae_configuracion " _
                    + vbCr + "WHERE ((llave) = '" & key & "')"
    
    Set xRs = database.GetRecordset()
    
    If xRs.State = 0 Then Exit Function
    If xRs.RecordCount = 0 Then
        MsgBox "No se encontro la llave especificada: " & key, vbInformation + vbOKOnly + vbDefaultButton1, "Funciones - KeyValue"
        mKeyValue = ""
    Else
        mKeyValue = NulosC(xRs("valor"))
    End If
    
    Set xRs = Nothing
    Set database = Nothing
End Function

Public Sub mMostrarMensajeError(Mensaje As String, Titulo As String, Optional Pila As String, Optional mErrorNumber)
    If IsNumeric(mErrorNumber) Then
        If mErrorNumber = vbObjectError + 1 Then Exit Sub
    End If
    MsgBox Mensaje, vbInformation + vbOKOnly + vbDefaultButton1, Titulo
End Sub

Public Function mBuscaCodigoTabla(valor_busca As Variant, campo_codigo As String, Campo_descripcion As String, Tabla As String, Tipo As String, Optional mConexion As ADODB.Connection)
    Dim RstBusca As New ADODB.Recordset
    Dim database As New SistemaData.EDataBase
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    
    If Tipo = "C" Then
        database.CommandText = "SELECT " & campo_codigo & ", " & Campo_descripcion & "  FROM  " & Tabla & " WHERE " & campo_codigo & " = '" & valor_busca & "'"
    End If
    If Tipo = "N" Then
        database.CommandText = "SELECT " & campo_codigo & ", " & Campo_descripcion & "  FROM  " & Tabla & " WHERE " & campo_codigo & " = " & valor_busca & ""
    End If
    Set RstBusca = database.GetRecordset
    If RstBusca.RecordCount = 0 Then
        mBuscaCodigoTabla = ""
    Else
        mBuscaCodigoTabla = mNuloString(RstBusca(Campo_descripcion))
    End If
    Exit Function
    
BloqueError:
    mBuscaCodigoTabla = ""
    mMostrarMensajeError Err.Description, "[BuscaCodigoTabla] " & Err.Source
End Function

Public Function mFechaInicioMovimientos(IdAlmacen As Long, Optional mConexion As ADODB.Connection) As Date
    Dim xRs As New ADODB.Recordset
    Dim database As New SistemaData.EDataBase
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
        
    database.CommandText = "SELECT alm_inventarioinicial.fechavigencia " _
        + vbCr + "FROM alm_inventarioinicial " _
        + vbCr + "WHERE (((alm_inventarioinicial.idalm)=" & IdAlmacen & ") AND ((alm_inventarioinicial.idestado)=" & mNuloNumeric(mKeyValue("EstadoAprobadoInventarioInicial", mConexion)) & "))"
    
    Set xRs = database.GetRecordset
    If xRs.State = 0 Then Err.Raise &HFFFFFF01, , "Error en el estado del Recordset"
    If xRs.RecordCount = 0 Then
        mFechaInicioMovimientos = Empty
        Set database = Nothing
        Exit Function
    End If
    
    mFechaInicioMovimientos = mConvertirNumeroAFecha(xRs("fechavigencia"))
    Set database = Nothing
    Exit Function
    
BloqueError:
    mFechaInicioMovimientos = Empty
    Set database = Nothing
    mMostrarMensajeError Err.Description, "[FechaInicioMovimientos] " & Err.Source
End Function

Function mSaldoActual(IdItem As Long, IdAlmacen As Long, _
                                FchInicio As Date, _
                                FchFinal As Date, _
                                Optional mConexion As ADODB.Connection, _
                                Optional xTipo As Long = 0) As Double
    Dim database As New SistemaData.EDataBase
    Dim Rst As New ADODB.Recordset
    Dim mSQL As String
    Dim xTotEnt As Double
    Dim xTotSal As Double
    Dim A As Long
    
    ' Se abre la conexion a la base de datos
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    '--CArgar solo cuando xTipo sea todos o Entradas
    If xTipo = 0 Or xTipo = 1 Then
        database.CommandText = "SELECT SUM(M.canpro) AS total " _
            + vbCr + "FROM (" _
            + vbCr + mKardexMovimientoSQL(IdItem, IdAlmacen, FchInicio, FchFinal, mConexion, 1) _
            + vbCr + ") AS M " _
            + vbCr + "WHERE (((M.tipo)<>'II'))"
        '--Cargar rst
        Set Rst = database.GetRecordset
        If Rst.RecordCount <> 0 Then
            xTotEnt = 0
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                xTotEnt = xTotEnt + NulosN(Rst("total"))
                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next A
        End If
    
    End If
    Set Rst = Nothing
    '--CArgar solo cuando xTipo sea todos o Entradas
    If xTipo = 0 Or xTipo = 2 Then
        'CARGAMOS TODAS LAS SALIDAS
        database.ClearParameter
        database.CommandText = "SELECT SUM(M.canpro) AS total " _
            + vbCr + "FROM (" _
            + vbCr + mKardexMovimientoSQL(IdItem, IdAlmacen, FchInicio, FchFinal, mConexion, 2) _
            + vbCr + ") AS M " _
            + vbCr + "WHERE (((M.tipo)<>'II'))"
            
        Set Rst = database.GetRecordset
        If Rst.RecordCount <> 0 Then
            xTotSal = 0
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                xTotSal = xTotSal + NulosN(Rst("total"))
                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next A
        End If
    End If
    Set Rst = Nothing
    
    If xTipo = 0 Then
        mSaldoActual = NulosN(xTotEnt) - NulosN(xTotSal) + mSaldoInicial(IdItem, IdAlmacen, mConexion)
        
    ElseIf xTipo = 1 Then '--Solo entradas
        mSaldoActual = NulosN(xTotEnt)
        
    ElseIf xTipo = 2 Then '--Solo Salidas
        mSaldoActual = NulosN(xTotSal)
    End If
End Function

Public Function mCostoInicial(IdItem As Long, IdAlmacen As Long, Optional mConexion As ADODB.Connection) As Double
    Dim database As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    database.CommandText = "SELECT alm_inventarioinicialdet.costo " _
            + vbCr + "FROM alm_inventarioinicial INNER JOIN alm_inventarioinicialdet ON alm_inventarioinicial.idinventarioinicial = alm_inventarioinicialdet.idinventarioinicial " _
            + vbCr + "WHERE (((alm_inventarioinicialdet.iditem)=" & IdItem & ") AND ((alm_inventarioinicial.idalm)=" & IdAlmacen & ") AND ((alm_inventarioinicial.idestado)=" & F.NuloNumeric(F.KeyValue("EstadoAprobadoInventarioInicial", mConexion)) & "))"
    
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount = 0 Then mCostoInicial = 0: Exit Function
    mRecord.MoveFirst
    mCostoInicial = F.NuloNumeric(mRecord("costo"))
    Exit Function
    
BloqueError:
    mCostoInicial = 0
    mMostrarMensajeError Err.Description, "[CostoInicial] " & Err.Source
End Function

Public Function mCostoActual(IdItem As Long, IdAlmacen As Long, _
                            FchInicio As Date, FchFinal As Date, _
                            Optional mConexion As ADODB.Connection, _
                            Optional xTipo As Long = 0) As Double
    Dim database As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones
    Dim mRecord As New ADODB.Recordset
    Dim mCostoEntradas As Double
    Dim mCostoSalidas As Double
    Dim A As Long
    
On Error GoTo BloqueError
'    If Not mConexion Is Nothing Then Set database.Connection = mConexion
'    database.CommandText = "SELECT alm_kardexdet.costounitariopromedio AS costo " _
'            + vbCr + "FROM alm_kardex INNER JOIN alm_kardexdet ON alm_kardex.idkardex = alm_kardexdet.idkardex " _
'            + vbCr + "WHERE (((alm_kardex.iditem)=" & IdItem & ") AND ((alm_kardexdet.idalm)=" & IdAlmacen & "))"
'
'    Set mRecord = database.GetRecordset
'    If mRecord.RecordCount = 0 Then mCostoActual = 0: Exit Function
'    mRecord.MoveFirst
'    mCostoActual = F.NuloNumeric(mRecord("costo"))
'    Exit Function
    
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    '--Cargar solo cuando xTipo sea todos o Entradas
    If xTipo = 0 Or xTipo = 1 Then
        database.CommandText = "SELECT SUM(M.costo) AS total " _
            + vbCr + "FROM (" _
            + vbCr + mKardexMovimientoSQL(IdItem, IdAlmacen, FchInicio, FchFinal, mConexion, 1) _
            + vbCr + ") AS M " _
            + vbCr + "WHERE (((M.tipo)<>'II'))"
        '--Cargar mRecord
        Set mRecord = database.GetRecordset
        If mRecord.RecordCount <> 0 Then
            mCostoEntradas = 0
            mRecord.MoveFirst
            For A = 1 To mRecord.RecordCount
                mCostoEntradas = mCostoEntradas + F.NuloNumeric(mRecord("total"))
                mRecord.MoveNext
                If mRecord.EOF = True Then
                    Exit For
                End If
            Next A
        End If
    
    End If
    Set mRecord = Nothing
    '--CArgar solo cuando xTipo sea todos o Entradas
    If xTipo = 0 Or xTipo = 2 Then
        'CARGAMOS TODAS LAS SALIDAS
        database.ClearParameter
        database.CommandText = "SELECT SUM(M.costo) AS total " _
            + vbCr + "FROM (" _
            + vbCr + mKardexMovimientoSQL(IdItem, IdAlmacen, FchInicio, FchFinal, mConexion, 2) _
            + vbCr + ") AS M " _
            + vbCr + "WHERE (((M.tipo)<>'II'))"
            
        Set mRecord = database.GetRecordset
        If mRecord.RecordCount <> 0 Then
            mCostoSalidas = 0
            mRecord.MoveFirst
            For A = 1 To mRecord.RecordCount
                mCostoSalidas = mCostoSalidas + F.NuloNumeric(mRecord("total"))
                mRecord.MoveNext
                If mRecord.EOF = True Then
                    Exit For
                End If
            Next A
        End If
    End If
    Set mRecord = Nothing
    
    If xTipo = 0 Then
        mCostoActual = F.NuloNumeric(mCostoEntradas) - F.NuloNumeric(mCostoSalidas) + mCostoInicial(IdItem, IdAlmacen, mConexion)
        
    ElseIf xTipo = 1 Then '--Solo entradas
        mCostoActual = F.NuloNumeric(mCostoEntradas)
        
    ElseIf xTipo = 2 Then '--Solo Salidas
        mCostoActual = F.NuloNumeric(mCostoSalidas)
    End If
    Exit Function
    
BloqueError:
    mCostoActual = 0
    mMostrarMensajeError Err.Description, "[CostoActual] " & Err.Source
End Function

Public Function mSQL_MovTotalizado(Cad_IdItem_In As String, _
                                    IdAlmacen As Long, _
                                    FechaInicio As Date, _
                                    FechaFin As Date, _
                                    Optional mConexion As ADODB.Connection, _
                                    Optional All_Items As Boolean = False, _
                                    Optional FiltraAlmacen As Boolean = False) As String
    
    Dim mSQL As String
    Dim QUERYA As String
    Dim QUERYB As String
    Dim QUERYC As String
    Dim QUERYD As String
    Dim FechaInicioCab As Date
    Dim FechaFinCab As Date
    
    If All_Items Then
        FechaInicioCab = mFechaInicioMovimientos(IdAlmacen, mConexion)
        FechaFinCab = Date
    Else
        FechaInicioCab = FechaInicio
        FechaFinCab = FechaFin
    End If
    If FechaInicio < mFechaInicioMovimientos(IdAlmacen, mConexion) Then
        FechaInicio = mFechaInicioMovimientos(IdAlmacen, mConexion)
    End If
    
    '***********
    ' CABECERA
    '***********
    QUERYA = "SELECT M.iditem, M.idtippro, M.tippro, M.coditem, M.item, M.unimed, LAST(M.costounitariopromedio) As costouniprom " _
            + vbCr + "FROM ( " _
            + vbCr + mSQL_MovDetallado(Cad_IdItem_In, IdAlmacen, FechaInicioCab, FechaFinCab, mConexion, , , , , FiltraAlmacen) _
            + vbCr + ") AS M " _
            + vbCr + "GROUP BY M.iditem, M.idtippro, M.tippro, M.coditem, M.item, M.unimed"

    '**************
    ' ENTRADAS
    '**************
    QUERYB = "SELECT M.iditem, SUM(M.costo) As costotot, SUM(M.cantidad) As cantot " _
            + vbCr + "FROM ( " _
            + vbCr + mSQL_MovDetallado(Cad_IdItem_In, IdAlmacen, FechaInicio, FechaFin, mConexion, 1, False, , , FiltraAlmacen) _
            + vbCr + ") AS M " _
            + vbCr + "GROUP BY M.iditem "

    '**************
    ' SALIDAS
    '**************
    QUERYC = "SELECT M.iditem, SUM(M.costo) As costotot, SUM(M.cantidad) As cantot " _
            + vbCr + "FROM ( " _
            + vbCr + mSQL_MovDetallado(Cad_IdItem_In, IdAlmacen, FechaInicio, FechaFin, mConexion, 2, False, , , FiltraAlmacen) _
            + vbCr + ") AS M " _
            + vbCr + "GROUP BY M.iditem "
            
    ' ***********************
    ' INVENTARIO INICIAL
    ' ***********************
    QUERYD = mSQL_MovHistoricoTotalizado(IdAlmacen, FechaInicio - 1, Cad_IdItem_In, mConexion, FiltraAlmacen)
    
    mSQL = "SELECT A.iditem, A.idtippro, A.tippro, A.coditem, A.item, A.unimed, B.cantot As canent, C.cantot As cansal, ((IIf(D.canini IS NULL, 0, D.canini) + IIF(D.canent IS NULL, 0, D.canent)) - IIF(D.cansal IS NULL, 0, D.cansal)) As canini, B.costotot As costoent, C.costotot As costosal, ((IIf(D.costoini IS NULL, 0, D.costoini) + IIF(D.costoent IS NULL, 0, D.costoent)) - IIF(D.costosal IS NULL, 0, D.costosal)) As costoini, A.costouniprom, D.costouniprom As costoiniuniprom " _
            + vbCr + "FROM (( " _
            + vbCr + "( " _
            + vbCr + QUERYA _
            + vbCr + ") AS A LEFT JOIN " _
            + vbCr + "( " _
            + vbCr + QUERYB _
            + vbCr + ") AS B ON A.iditem = B.iditem) LEFT JOIN " _
            + vbCr + "( " _
            + vbCr + QUERYC _
            + vbCr + ") AS C ON A.iditem = C.iditem) LEFT JOIN " _
            + vbCr + "( " _
            + vbCr + QUERYD _
            + vbCr + ") AS D ON A.iditem = D.iditem " _
            + vbCr + "GROUP BY A.iditem, A.idtippro, A.tippro, A.coditem, A.item, A.unimed, B.cantot, C.cantot, ((IIf(D.canini IS NULL, 0, D.canini) + IIF(D.canent IS NULL, 0, D.canent)) - IIF(D.cansal IS NULL, 0, D.cansal)), B.costotot, C.costotot, ((IIf(D.costoini IS NULL, 0, D.costoini) + IIF(D.costoent IS NULL, 0, D.costoent)) - IIF(D.costosal IS NULL, 0, D.costosal)), A.costouniprom, D.costouniprom " _
            + vbCr + "ORDER BY A.coditem "

    mSQL_MovTotalizado = mSQL
End Function

Public Function mSQL_MovHistoricoTotalizado(IdAlmacen As Long, _
                                    FechaConsulta As Date, _
                                    Optional Cad_IdItem_In As String = "", _
                                    Optional mConexion As ADODB.Connection, _
                                    Optional FiltraAlmacen As Boolean = False) As String
    
    Dim mSQL As String
    Dim QUERYA As String
    Dim QUERYB As String
    Dim QUERYC As String
    Dim QUERYD As String
    Dim FechaInicioMovimientos As Date
    Dim mFiltroFecha As String
    
    FechaInicioMovimientos = mFechaInicioMovimientos(IdAlmacen, mConexion)
    
    ' Inventario Inicial
    If FechaConsulta < FechaInicioMovimientos Then
        mFiltroFecha = "WHERE (((con_librocostotemp.fecha)=CDate('" & FechaInicioMovimientos & "')) AND ((alm_ingreso.idtipdocref)=" & mNuloNumeric(mKeyValue("IdDocumentoInventarioInicial", mConexion)) & ")) "
    Else
        mFiltroFecha = "WHERE (((con_librocostotemp.fecha)>=CDate('" & FechaInicioMovimientos & "') And (con_librocostotemp.fecha)<=CDate('" & FechaConsulta & "'))) "
    End If
    
    QUERYA = "SELECT M.iditem, M.idtippro, M.tippro, M.coditem, M.item, M.unimed, CULTCOST.costounitariopromedio As costouniprom " _
            + vbCr + "FROM ( " _
            + vbCr + "SELECT CM.iditem, CM.idtippro, CM.tippro, CM.coditem, CM.item, CM.unimed " _
            + vbCr + "FROM ( " _
            + vbCr + mSQL_MovDetallado(Cad_IdItem_In, IdAlmacen, FechaInicioMovimientos, FechaConsulta, mConexion, , , , , FiltraAlmacen) _
            + vbCr + ") AS CM " _
            + vbCr + "GROUP BY CM.iditem, CM.idtippro, CM.tippro, CM.coditem, CM.item, CM.unimed " _
            + vbCr + ") AS M LEFT JOIN ( " _
            + vbCr + "SELECT CULTMOV.iditem, con_librocostotemp.idmovdet, con_librocostotemp.costounitariopromedio " _
            + vbCr + "FROM con_librocostotemp " _
            + vbCr + "INNER JOIN ( " _
            + vbCr + "SELECT alm_ingresodet.iditem, Max(con_librocostotemp.idlibrocostotemp) AS idlibrocostotemp " _
            + vbCr + "FROM alm_ingreso INNER JOIN (alm_ingresodet INNER JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id " _
            + vbCr + mFiltroFecha _
            + vbCr + "GROUP BY alm_ingresodet.iditem " _
            + vbCr + ") AS CULTMOV ON con_librocostotemp.idlibrocostotemp = CULTMOV.idlibrocostotemp " _
            + vbCr + ") AS CULTCOST ON M.iditem = CULTCOST.iditem"
            
    ' Entradas
    QUERYB = "SELECT M.iditem, SUM(M.costo) As costotot, SUM(M.cantidad) As cantot " _
            + vbCr + "FROM ( " _
            + vbCr + mSQL_MovDetallado(Cad_IdItem_In, IdAlmacen, FechaInicioMovimientos, FechaConsulta, mConexion, 1, False, , , FiltraAlmacen) _
            + vbCr + ") AS M " _
            + vbCr + "GROUP BY M.iditem "

    ' Salidas
    QUERYC = "SELECT M.iditem, SUM(M.costo) As costotot, SUM(M.cantidad) As cantot " _
            + vbCr + "FROM ( " _
            + vbCr + mSQL_MovDetallado(Cad_IdItem_In, IdAlmacen, FechaInicioMovimientos, FechaConsulta, mConexion, 2, False, , , FiltraAlmacen) _
            + vbCr + ") AS M " _
            + vbCr + "GROUP BY M.iditem "
    
    ' Inventario Inicial
    QUERYD = "SELECT alm_inventarioinicialdet.iditem, (alm_inventarioinicialdet.costo*alm_inventarioinicialdet.cantidad) AS costotot, alm_inventarioinicialdet.cantidad AS cantot " _
            + vbCr + "FROM alm_inventarioinicial INNER JOIN alm_inventarioinicialdet ON alm_inventarioinicial.idinventarioinicial = alm_inventarioinicialdet.idinventarioinicial " _
            + vbCr + "WHERE (((alm_inventarioinicial.idalm)=" & IdAlmacen & ") AND ((alm_inventarioinicial.idestado)=" & mNuloNumeric(mKeyValue("EstadoAprobadoInventarioInicial", mConexion)) & "))"

    If Cad_IdItem_In <> "" Then
        QUERYD = QUERYD & " AND ((alm_inventarioinicialdet.iditem) IN (" & Cad_IdItem_In & ")) "
    End If
    
    QUERYD = QUERYD + vbCr + "GROUP BY alm_inventarioinicialdet.iditem, (alm_inventarioinicialdet.costo*alm_inventarioinicialdet.cantidad), alm_inventarioinicialdet.cantidad"
    
    mSQL = "SELECT ATOT.iditem, ATOT.idtippro, ATOT.tippro, ATOT.coditem, ATOT.item, ATOT.unimed, ATOT.costouniprom, BTOT.cantot As canent, CTOT.cantot As cansal, DTOT.cantot As canini, BTOT.costotot As costoent, CTOT.costotot As costosal, DTOT.costotot As costoini " _
            + vbCr + "FROM (( " _
            + vbCr + "( " _
            + vbCr + QUERYA _
            + vbCr + ") AS ATOT LEFT JOIN " _
            + vbCr + "( " _
            + vbCr + QUERYB _
            + vbCr + ") AS BTOT ON ATOT.iditem = BTOT.iditem) LEFT JOIN " _
            + vbCr + "( " _
            + vbCr + QUERYC _
            + vbCr + ") AS CTOT ON ATOT.iditem = CTOT.iditem) LEFT JOIN " _
            + vbCr + "( " _
            + vbCr + QUERYD _
            + vbCr + ") AS DTOT ON ATOT.iditem = DTOT.iditem "
    
    mSQL_MovHistoricoTotalizado = mSQL
End Function

Public Function mSaldoInicial(IdItem As Long, IdAlmacen As Long, _
                                Optional mConexion As ADODB.Connection) As Double
    Dim database As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones
    Dim mRecord As New ADODB.Recordset
    Dim QUERYA As String
    Dim QUERYB As String
    Dim QUERYC As String
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    database.CommandText = "SELECT alm_inventarioinicialdet.cantidad " _
            + vbCr + "FROM alm_inventarioinicial INNER JOIN alm_inventarioinicialdet ON alm_inventarioinicial.idinventarioinicial = alm_inventarioinicialdet.idinventarioinicial " _
            + vbCr + "WHERE (((alm_inventarioinicialdet.iditem)=" & IdItem & ") AND ((alm_inventarioinicial.idalm)=" & IdAlmacen & ") AND ((alm_inventarioinicial.idestado)=" & F.NuloNumeric(F.KeyValue("EstadoAprobadoInventarioInicial", mConexion)) & "))"
        
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount = 0 Then mSaldoInicial = 0: Exit Function
    mRecord.MoveFirst
    mSaldoInicial = F.NuloNumeric(mRecord("cantidad"))
    Exit Function
    
BloqueError:
    mSaldoInicial = 0
    mMostrarMensajeError Err.Description, "[SaldoInicial] " & Err.Source
End Function

Public Function mErrorDescriptionDLL(Optional ByVal lLastDLLError As Long) As String
    Dim sBuff As String * 256
    Dim lCount As Long
    Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100, FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
    Const FORMAT_MESSAGE_FROM_HMODULE = &H800, FORMAT_MESSAGE_FROM_STRING = &H400
    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000, FORMAT_MESSAGE_IGNORE_INSERTS = &H200
    Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

    If lLastDLLError = 0 Then
        'Use Err object to get dll error number
        lLastDLLError = Err.LastDllError
    End If

    lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
        mErrorDescriptionDLL = Left$(sBuff, lCount - 2)    'Remove line feeds
    End If
End Function

Public Function mMachineName() As String
    mMachineName = ""
End Function

Public Function mExisteDocumento(Tabla As String, Valor As String, _
                                        Optional Conexion As ADODB.Connection, _
                                        Optional Campo As String = "numdoc", _
                                        Optional Valor2 As String, _
                                        Optional campo2 As String = "numser", _
                                        Optional Valor3 As String, _
                                        Optional campo3 As String, _
                                        Optional IdDocumento As Long, _
                                        Optional CampoIdDocumento As String) As Boolean
    Dim nSQL As String
    Dim database As New SistemaData.EDataBase
    Dim xRs As New ADODB.Recordset

On Error GoTo BloqueError
    If Valor2 <> "" And campo2 <> "" Then
        nSQL = " AND ((" & campo2 & ") = " & Valor2 & ")"
    End If
    
    If nSQL <> "" And Valor3 <> "" And campo3 <> "" Then
        nSQL = nSQL & " AND ((" & campo3 & ") = " & Valor3 & ")"
    End If
    
    If nSQL <> "" And CampoIdDocumento <> "" And IdDocumento <> 0 Then
        nSQL = nSQL & " AND ((" & CampoIdDocumento & ") <> " & IdDocumento & ")"
    End If
    
    If Not Conexion Is Nothing Then Set database.Connection = Conexion
    
    database.CommandText = "SELECT " & Tabla & "." & Campo & " " _
        + vbCr + "FROM " & Tabla & " " _
        + vbCr + "WHERE ((" & Campo & " = " + Valor + "))" & nSQL
    
    Set xRs = database.GetRecordset
    If xRs.RecordCount = 0 Then
        mExisteDocumento = False
    Else
        mExisteDocumento = True
    End If
    Exit Function
    
BloqueError:
    mExisteDocumento = False
    mMostrarMensajeError Err.Description, "[FechaInicioMovimientos] " & Err.Source
End Function

Public Function mAlmacenaEn(IdItem As Long, Optional Conexion As ADODB.Connection) As Long
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BloqueError
    If Not Conexion Is Nothing Then Set database.Connection = Conexion
    database.CommandText = "SELECT alm_almacenajeauto.idalm " _
                            + vbCr + "FROM alm_almacenajeauto " _
                            + vbCr + "WHERE (((alm_almacenajeauto.iditem)=" & IdItem & "))"
                            
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount = 0 Then mAlmacenaEn = 0: Exit Function
    mAlmacenaEn = mNuloNumeric(mRecord("idalm"))
    Exit Function
    
BloqueError:
    mAlmacenaEn = 0
    mMostrarMensajeError Err.Description, "[AlmacenaEn] " & Err.Source
End Function

Public Function mDespachaEn(IdItem As Long, Optional Conexion As ADODB.Connection) As Long
    Dim database As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BloqueError
    If Not Conexion Is Nothing Then Set database.Connection = Conexion
    database.CommandText = "SELECT alm_despachoauto.idalm " _
                            + vbCr + "FROM alm_despachoauto " _
                            + vbCr + "WHERE (((alm_despachoauto.iditem)=" & IdItem & "))"
                            
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount = 0 Then mDespachaEn = 0: Exit Function
    mDespachaEn = mNuloNumeric(mRecord("idalm"))
    Exit Function
    
BloqueError:
    mDespachaEn = 0
    mMostrarMensajeError Err.Description, "[DespachaEn] " & Err.Source
End Function

Function mKardexMovimientoSQL(IdItem As Long, _
                                IdAlmacen As Long, _
                                xFchIni As Date, _
                                xFchFin As Date, _
                                Optional mConexion As ADODB.Connection, _
                                Optional xTipo As Long = 0, _
                                Optional MuestraInvIni As Boolean = True) As String
    ' AI = Almacen Ingreso
    ' AS = Almacen Salida
    ' C =  Compras
    ' SM = SOLICUTID DE MATERIALES
    ' PP = PARTE DE PRODUCCION
    'GR = GUIAS DE REMISION
    'PS =

    Dim xCadSQL As String
    Dim xSQLFiltroPS As String '--Util para aplicar un filtro adicional que mostrará solo materia prima en sentencia de "produccion insumos salida"
    Dim F As New SistemaLogica.Funciones
    Dim mFechaInicioMovimientos As Date
    Dim mSQLFechaMov As String
    Dim mSQLIdItem As String
    Dim mSQLIdAlmacen As String
    Dim mSQLInvInicial As String
    Dim mSQLExclude As String
    Dim IdTipoInvInicial As Long

    mFechaInicioMovimientos = F.FechaInicioMovimientos(IdAlmacen, mConexion)
    IdTipoInvInicial = F.NuloNumeric(F.KeyValue("IdDocumentoInventarioInicial", mConexion))

    ' PREPARAMOS LA SELECT PARA ARMAR EL KARDEX
    If xTipo = 0 Or xTipo = 1 Then
        ' Ingresos de Almacen
        If IdItem = 0 Then
            mSQLIdItem = ""
        Else
            mSQLIdItem = "AND ((alm_ingresodet.iditem)=" & IdItem & ") "
        End If
        If IdAlmacen = 0 Then
            mSQLIdAlmacen = ""
        Else
            mSQLIdAlmacen = "AND ((alm_ingreso.idalm)=" & IdAlmacen & ") "
        End If
        If Not IsDate(mFechaInicioMovimientos) Or mFechaInicioMovimientos = Empty Then
            mSQLFechaMov = ""
        Else
            mSQLFechaMov = "AND (alm_ingreso.fching)>=CDate('" & mFechaInicioMovimientos & "') "
        End If
        If MuestraInvIni Then
            mSQLInvInicial = ""
        Else
            mSQLInvInicial = "AND ((alm_ingreso.idtipdocref)<>" & F.NuloNumeric(F.KeyValue("IdDocumentoInventarioInicial", mConexion)) & ") "
        End If
        
        xCadSQL = "SELECT alm_ingreso.id, alm_ingresodet.idmovdet, alm_ingresodet.iditem, alm_inventario.codpro, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, alm_ingreso.numser, alm_ingreso.numdoc, alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, IIf(alm_ingreso.idtipdocref=121,'II','AI') AS tipo, alm_ingreso.tipmov, alm_ingreso.nombre AS entidad, 0 AS aa, (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos, 'Almacén' & IIf(CStr(numdocumentos)<>'0',' - Compras','') AS modulo, '' AS registro, '' AS ctanum, '' AS ctanom, IIf([alm_ingreso].[idope]=1,'RECEPCION',IIf([alm_ingreso].[idope]=2,'DESPACHO',IIf([alm_ingreso].[idope]=3,'ENTRADA PRODUCCION',IIf([alm_ingreso].[idope]=4,'SALIDA PRODUCCION','')))) AS desope, " _
                    & "alm_ingresodet.hora AS horini, alm_ingresodet.hora AS horfin, alm_inventario.tippro AS idtippro, mae_tipoproducto.descripcion AS tippro, alm_inventario.idunimed, mae_unidades.abrev AS unimed, (con_librocostotemp.costoprimo + IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS costo, con_librocostotemp.costounitariopromedio  " _
            + vbCr + "FROM (alm_ingreso LEFT JOIN ((((alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id " _
            + vbCr + "WHERE (((alm_ingreso.fching)>=CDate('" & xFchIni & "') And (alm_ingreso.fching)<=CDate('" & xFchFin & "')) AND ((alm_ingreso.tipmov)=-1)) AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad <> 0 " & mSQLIdItem & mSQLIdAlmacen & mSQLFechaMov & mSQLInvInicial
        
        ' Ingreso por Compras
        If IdItem = 0 Then
            mSQLIdItem = ""
        Else
            mSQLIdItem = "AND ((com_comprasdet.iditem)=" & IdItem & ") "
        End If
        If IdAlmacen = 0 Then
            mSQLIdAlmacen = ""
        Else
            mSQLIdAlmacen = "AND ((com_compras.idalm)=" & IdAlmacen & ") "
        End If
        If Not IsDate(mFechaInicioMovimientos) Or mFechaInicioMovimientos = Empty Then
            mSQLFechaMov = ""
        Else
            mSQLFechaMov = " AND (com_compras.fchdoc)>=CDate('" & mFechaInicioMovimientos & "')"
        End If
        '**********************************************
        ' Exlusione de compras
        mSQLExclude = mSQLExclude & " AND ((alm_inventario.tippro)<> 5) "
        '**********************************************
        xCadSQL = xCadSQL _
            + vbCr + "UNION ALL " _
            + vbCr + "SELECT com_compras.id, 0 As idmovdet, com_comprasdet.iditem, alm_inventario.codpro, alm_inventario.descripcion, com_compras.fchdoc, com_compras.numser, com_compras.numdoc, com_comprasdet.canpro, IIf([com_compras]![idmon]=2,[com_comprasdet]![preuni]*[con_tc]![impcom],[com_comprasdet]![preuni]) AS preuni, mae_documento.abrev AS descdoc, 'C' AS Tipo, -1 AS tipmov, mae_prov.nombre AS entidad, 0 AS aa, 0 AS numdocumentos, 'Compras' AS modulo, com_compras.numreg AS registro, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctanom, '' AS desope, '' AS horini, '' AS horfin, alm_inventario.tippro AS idtippro, mae_tipoproducto.descripcion AS tippro, alm_inventario.idunimed, mae_unidades.abrev AS unimed, com_comprasdet.imptot AS costo, com_comprasdet.preuni AS costounitariopromedio " _
            + vbCr + "FROM (((alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((com_compras.fchdoc)>=CDate('" & xFchIni & "') And (com_compras.fchdoc)<=CDate('" & xFchFin & "')) AND ((com_compras.tipcom)=1)) " & mSQLIdItem & mSQLIdAlmacen & mSQLFechaMov & mSQLExclude
    
    End If
    
    If xTipo = 0 Or xTipo = 2 Then
        If xTipo = 0 Then xCadSQL = xCadSQL + vbCr + "UNION ALL "
        ' Salidas de Almacen
        If IdItem = 0 Then
            mSQLIdItem = ""
        Else
            mSQLIdItem = "AND ((alm_ingresodet.iditem)=" & IdItem & ") "
        End If
        If IdAlmacen = 0 Then
            mSQLIdAlmacen = ""
        Else
            mSQLIdAlmacen = "AND ((alm_ingreso.idalm)=" & IdAlmacen & ") "
        End If
        If Not IsDate(mFechaInicioMovimientos) Or mFechaInicioMovimientos = Empty Then
            mSQLFechaMov = ""
        Else
            mSQLFechaMov = " AND (alm_ingreso.fching)>=CDate('" & mFechaInicioMovimientos & "')"
        End If
        xCadSQL = xCadSQL _
            + vbCr + "SELECT alm_ingreso.id, alm_ingresodet.idmovdet, alm_ingresodet.iditem, alm_inventario.codpro, alm_inventario.descripcion, alm_ingreso.fching AS fchdoc, alm_ingreso.numser, alm_ingreso.numdoc, alm_ingresodet.cantidad AS canpro, alm_ingresodet.preuni, mae_documento.abrev AS descdoc, 'AS' AS tipo, alm_ingreso.tipmov, alm_ingreso.nombre AS entidad, 0 AS aa, (SELECT Count(1) AS numdocumentos FROM alm_ingresodoc WHERE (((alm_ingresodoc.id)=alm_ingreso.id))) AS numdocumentos, 'Almacén' & IIf(CStr(numdocumentos)<>'0',' - Compras','') AS modulo, '' AS registro, '' AS ctanum, '' AS ctanom, IIf([alm_ingreso].[idope]=1,'RECEPCION',IIf([alm_ingreso].[idope]=2,'DESPACHO',IIf([alm_ingreso].[idope]=3,'ENTRADA PRODUCCION',IIf([alm_ingreso].[idope]=4,'SALIDA PRODUCCION','')))) AS desope, " _
                    & "alm_ingresodet.hora AS horini, alm_ingresodet.hora AS horfin, alm_inventario.tippro AS idtippro, mae_tipoproducto.descripcion AS tippro, alm_inventario.idunimed, mae_unidades.abrev AS unimed, (con_librocostotemp.costoprimo + IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS costo, con_librocostotemp.costounitariopromedio  " _
            + vbCr + "FROM (alm_ingreso LEFT JOIN ((((alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id " _
            + vbCr + "WHERE (((alm_ingreso.fching)>=CDate('" & xFchIni & "') And (alm_ingreso.fching)<=CDate('" & xFchFin & "')) AND ((alm_ingreso.tipmov)=0)) AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " & mSQLIdItem & mSQLIdAlmacen & mSQLFechaMov
        
    End If
    
    mKardexMovimientoSQL = xCadSQL
End Function

Function mSQL_MovDetallado(Optional Cad_IdItem_In As String = "", _
                                Optional IdAlmacen As Long = 0, _
                                Optional FechaInicio As Date = Empty, _
                                Optional FechaFin As Date = Empty, _
                                Optional mConexion As ADODB.Connection, _
                                Optional FiltraTipoMovimiento As Integer = 0, _
                                Optional MuestraInvIni As Boolean = True, _
                                Optional MuestraCompras As Boolean = False, _
                                Optional IdMovimientoDetalle As Long = 0, _
                                Optional FiltraAlmacen As Boolean = False) As String

    Dim xCadSQL As String
    Dim mFchIniMov As Date
    Dim mSQLFechaMov As String
    Dim mSQLIdItem As String
    Dim mSQLIdAlmacen As String
    Dim mSQLIdMovimientoDetalle As String
    Dim mSQLInvInicial As String
    Dim mSQLExclude As String
    Dim mSQLDocumento As String
    Dim IdTipoInvInicial As Long

    mFchIniMov = mFechaInicioMovimientos(IdAlmacen, mConexion)
    IdTipoInvInicial = mNuloNumeric(mKeyValue("IdDocumentoInventarioInicial", mConexion))
    
    mSQLDocumento = "SELECT alm_ingreso.id, alm_ingreso.idtipdocref, alm_ingreso.iddocref, IIf([alm_ingreso].[idtipdocref]=110,[pro_solicitudmat].[numser] & '-' & [pro_solicitudmat].[numdoc],IIf([alm_ingreso].[idtipdocref]=71,[alm_recepcion].[numser] & '-' & [alm_recepcion].[numdoc],IIf([alm_ingreso].[idtipdocref]=114,[alm_devolucion].[numser] & '-' & [alm_devolucion].[numdoc],IIf([alm_ingreso].[idtipdocref]=92,[com_ordencompra].[numser] & '-' & [com_ordencompra].[numdoc],IIf([alm_ingreso].[idtipdocref]=119,[alm_transferencia].[numser] & '-' & [alm_transferencia].[numdoc],IIf([alm_ingreso].[idtipdocref]=120,[pro_produccion].[numser] & '-' & [pro_produccion].[numdoc],IIf([alm_ingreso].[idtipdocref]=111,[alm_tomainventario].[numser] & '-' & [alm_tomainventario].[numdoc], " _
            & "IIf([alm_ingreso].[idtipdocref]=9,[vta_guia].[numser] & '-' & [vta_guia].[numdoc],IIf([alm_ingreso].[idtipdocref]=1,[vta_ventas].[numser] & '-' & [vta_ventas].[numdoc],IIf([alm_ingreso].[idtipdocref]=121,[alm_tomainventario_1].[numser] & '-' & [alm_tomainventario_1].[numdoc],'')))))))))) AS numdocrefconcat, IIf([alm_ingreso].[idtipdocref]=110,[pro_solicitudmat].[numser],IIf([alm_ingreso].[idtipdocref]=71,[alm_recepcion].[numser],IIf([alm_ingreso].[idtipdocref]=114,[alm_devolucion].[numser],IIf([alm_ingreso].[idtipdocref]=92,[com_ordencompra].[numser],IIf([alm_ingreso].[idtipdocref]=119,[alm_transferencia].[numser],IIf([alm_ingreso].[idtipdocref]=120,[pro_produccion].[numser],IIf([alm_ingreso].[idtipdocref]=111,[alm_tomainventario].[numser],IIf([alm_ingreso].[idtipdocref]=9,[vta_guia].[numser], " _
            & "IIf([alm_ingreso].[idtipdocref]=1,[vta_ventas].[numser],IIf([alm_ingreso].[idtipdocref]=121,[alm_tomainventario_1].[numser],'')))))))))) AS numserref, IIf([alm_ingreso].[idtipdocref]=110,[pro_solicitudmat].[numdoc],IIf([alm_ingreso].[idtipdocref]=71,[alm_recepcion].[numdoc],IIf([alm_ingreso].[idtipdocref]=114,[alm_devolucion].[numdoc],IIf([alm_ingreso].[idtipdocref]=92,[com_ordencompra].[numdoc],IIf([alm_ingreso].[idtipdocref]=119,[alm_transferencia].[numdoc],IIf([alm_ingreso].[idtipdocref]=120,[pro_produccion].[numdoc],IIf([alm_ingreso].[idtipdocref]=111,[alm_tomainventario].[numdoc],IIf([alm_ingreso].[idtipdocref]=9,[vta_guia].[numdoc],IIf([alm_ingreso].[idtipdocref]=1,[vta_ventas].[numdoc],IIf([alm_ingreso].[idtipdocref]=121,[alm_tomainventario_1].[numdoc],'')))))))))) AS numdocref " _
        + vbCr + "FROM (((((((((((alm_ingreso LEFT JOIN vta_ventas ON alm_ingreso.iddocref = vta_ventas.id) LEFT JOIN com_compras ON alm_ingreso.iddocref = com_compras.id) LEFT JOIN alm_devolucion ON alm_ingreso.iddocref = alm_devolucion.id) LEFT JOIN vta_guia ON alm_ingreso.iddocref = vta_guia.id) LEFT JOIN alm_tomainventario ON alm_ingreso.iddocref = alm_tomainventario.idtomainventario) LEFT JOIN pro_produccion ON alm_ingreso.iddocref = pro_produccion.id) LEFT JOIN alm_recepcion ON alm_ingreso.iddocref = alm_recepcion.id) LEFT JOIN pro_solicitudmat ON alm_ingreso.iddocref = pro_solicitudmat.id) LEFT JOIN com_ordencompra ON alm_ingreso.iddocref = com_ordencompra.id) LEFT JOIN alm_transferencia ON alm_ingreso.iddocref = alm_transferencia.idtransferencia) LEFT JOIN alm_inventarioinicial ON alm_ingreso.iddocref = alm_inventarioinicial.idinventarioinicial) " _
            & "LEFT JOIN alm_tomainventario AS alm_tomainventario_1 ON alm_inventarioinicial.iddocref = alm_tomainventario_1.idtomainventario"

    ' PREPARAMOS LA SELECT PARA ARMAR EL KARDEX
    If FiltraTipoMovimiento = 0 Or FiltraTipoMovimiento = 1 Then
        '**********************
        ' INGRESOS
        '**********************
        If Cad_IdItem_In = "" Then
            mSQLIdItem = ""
        Else
            mSQLIdItem = "AND ((alm_ingresodet.iditem) IN (" & Cad_IdItem_In & ")) "
        End If
        If IdAlmacen = 0 Then
            mSQLIdAlmacen = ""
        Else
            mSQLIdAlmacen = "AND ((alm_ingreso.idalm)=" & IdAlmacen & ") "
        End If
        If IdMovimientoDetalle = 0 Then
            mSQLIdMovimientoDetalle = ""
        Else
            mSQLIdMovimientoDetalle = "AND ((alm_ingresodet.idmovdet)=" & IdMovimientoDetalle & ") "
        End If
        If Not IsDate(mFchIniMov) Or mFchIniMov = Empty Then
            mSQLFechaMov = ""
        Else
            mSQLFechaMov = "AND (alm_ingreso.fching)>=CDate('" & mFchIniMov & "') "
        End If
        If IsDate(FechaInicio) And FechaInicio <> Empty Then
            mSQLFechaMov = mSQLFechaMov & "AND (alm_ingreso.fching)>=CDate('" & FechaInicio & "') "
        End If
        If IsDate(FechaFin) And FechaFin <> Empty Then
            mSQLFechaMov = mSQLFechaMov & "AND (alm_ingreso.fching)<=CDate('" & FechaFin & "') "
        End If
        If FiltraAlmacen Then
            mSQLIdAlmacen = mSQLIdAlmacen & "AND ((alm_ingreso.idalm) IN (SELECT alm_almacenes.id FROM alm_almacenes WHERE alm_almacenes.vismov = -1)) "
        End If
        
        mSQLInvInicial = "AND ((alm_ingreso.idtipdocref)<>" & mNuloNumeric(mKeyValue("IdDocumentoInventarioInicial", mConexion)) & ") "
        
        xCadSQL = "SELECT alm_ingreso.id As iddoc, alm_ingreso.id As idmov, alm_ingresodet.idmovdet, con_librocostotemp.idlibrocostotemp, alm_ingresodet.iditem, alm_inventario.codpro As coditem, alm_inventario.descripcion As item, alm_ingreso.fching AS fchmov, alm_ingreso.numser, alm_ingreso.numdoc, alm_ingreso.numser & '-' & alm_ingreso.numdoc As numdocconcat, alm_ingresodet.cantidad, mae_documento.abrev AS doc, alm_ingreso.idtipdocref, alm_ingreso.tipmov, 'I' As tipmovcad, alm_inventario.tippro AS idtippro, mae_tipoproducto.descripcion AS tippro, alm_inventario.idunimed, mae_unidades.abrev AS unimed, (con_librocostotemp.costoprimo + IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS costo, con_librocostotemp.costounitariopromedio, alm_ingreso.idalm, alm_almacenes.descripcion AS alm  " _
                & ", con_librocostotemp.costounitario, con_librocostotemp.costoprimo, con_librocostotemp.costomod, con_librocostotemp.costocif, mae_documento.descripcion AS tipdocref, alm_ingreso.iddocref, 0 AS idtipdocrefanex, '' AS tipdocrefanex, alm_ingresodet.iddocref As iddocrefanex, '' As docrefanex, CNUMDOCREF.numserref, CNUMDOCREF.numdocref, CNUMDOCREF.numdocrefconcat  " _
            + vbCr + "FROM (((alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id) LEFT JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id) LEFT JOIN ((((alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN ( " _
            + vbCr + mSQLDocumento _
            + vbCr + ") AS CNUMDOCREF ON alm_ingreso.id = CNUMDOCREF.id " _
            + vbCr + "WHERE (((alm_ingreso.tipmov)=-1)) AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad <> 0 " & mSQLIdItem & mSQLIdAlmacen & mSQLFechaMov & mSQLInvInicial & mSQLIdMovimientoDetalle
        
        If MuestraInvIni Then
            '**********************
            ' INVENTARIO INICIAL
            '**********************
            If Cad_IdItem_In = "" Then
                mSQLIdItem = ""
            Else
                mSQLIdItem = "AND ((alm_ingresodet.iditem) IN (" & Cad_IdItem_In & ")) "
            End If
            If IdMovimientoDetalle = 0 Then
                mSQLIdMovimientoDetalle = ""
            Else
                mSQLIdMovimientoDetalle = "AND ((alm_ingresodet.idmovdet)=" & IdMovimientoDetalle & ") "
            End If
            If IdAlmacen = 0 Then
                mSQLIdAlmacen = ""
            Else
                mSQLIdAlmacen = "AND ((alm_ingreso.idalm)=" & IdAlmacen & ") "
            End If
            
            If Not IsDate(mFchIniMov) Or mFchIniMov = Empty Then
                mSQLFechaMov = ""
            Else
                mSQLFechaMov = "AND (alm_ingreso.fching)>=CDate('" & mFchIniMov & "') "
            End If
            If IsDate(FechaInicio) And FechaInicio <> Empty Then
                mSQLFechaMov = mSQLFechaMov & "AND (alm_ingreso.fching)>=CDate('" & FechaInicio & "') "
            End If
            If FiltraAlmacen Then
                mSQLIdAlmacen = mSQLIdAlmacen & "AND ((alm_ingreso.idalm) IN (SELECT alm_almacenes.id FROM alm_almacenes WHERE alm_almacenes.vismov = -1)) "
            End If
                                   
            mSQLInvInicial = "AND ((alm_ingreso.idtipdocref)=" & mNuloNumeric(mKeyValue("IdDocumentoInventarioInicial", mConexion)) & ") "
            
            xCadSQL = xCadSQL _
                + vbCr + "UNION ALL " _
                + vbCr + "SELECT alm_ingreso.id As iddoc, alm_ingreso.id As idmov, alm_ingresodet.idmovdet, con_librocostotemp.idlibrocostotemp, alm_ingresodet.iditem, alm_inventario.codpro As coditem, alm_inventario.descripcion As item, alm_ingreso.fching AS fchmov, alm_ingreso.numser, alm_ingreso.numdoc, alm_ingreso.numser & '-' & alm_ingreso.numdoc As numdocconcat, alm_ingresodet.cantidad, mae_documento.abrev AS doc, alm_ingreso.idtipdocref, alm_ingreso.tipmov, 'I' As tipmovcad, alm_inventario.tippro AS idtippro, mae_tipoproducto.descripcion AS tippro, alm_inventario.idunimed, mae_unidades.abrev AS unimed, (con_librocostotemp.costoprimo + IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS costo, con_librocostotemp.costounitariopromedio, alm_ingreso.idalm, alm_almacenes.descripcion AS alm  " _
                    & ", con_librocostotemp.costounitario, con_librocostotemp.costoprimo, con_librocostotemp.costomod, con_librocostotemp.costocif, mae_documento.descripcion AS tipdocref, alm_ingreso.iddocref, 0 AS idtipdocrefanex, '' AS tipdocrefanex, alm_ingresodet.iddocref As iddocrefanex, '' As docrefanex, CNUMDOCREF.numserref, CNUMDOCREF.numdocref, CNUMDOCREF.numdocrefconcat  " _
                + vbCr + "FROM (((alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id) LEFT JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id) LEFT JOIN ((((alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN ( " _
                + vbCr + mSQLDocumento _
                + vbCr + ") AS CNUMDOCREF ON alm_ingreso.id = CNUMDOCREF.id " _
                + vbCr + "WHERE (((alm_ingreso.tipmov)=-1)) AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad <> 0 " & mSQLIdItem & mSQLIdAlmacen & mSQLFechaMov & mSQLInvInicial & mSQLIdMovimientoDetalle
        
        End If
        
        If MuestraCompras Then
            '**********************
            ' COMPRAS
            '**********************
            If Cad_IdItem_In = "" Then
                mSQLIdItem = ""
            Else
                mSQLIdItem = "AND ((com_comprasdet.iditem) IN (" & Cad_IdItem_In & ")) "
            End If
            If IdMovimientoDetalle = 0 Then
                mSQLIdMovimientoDetalle = ""
            Else
                mSQLIdMovimientoDetalle = "AND ((com_compras.id) Is Null) "
            End If
            If IdAlmacen = 0 Then
                mSQLIdAlmacen = ""
            Else
                mSQLIdAlmacen = "AND ((com_compras.idalm)=" & IdAlmacen & ") "
            End If
            If FiltraAlmacen Then
                mSQLIdAlmacen = mSQLIdAlmacen & "AND ((com_compras.idalm) IN (SELECT alm_almacenes.id FROM alm_almacenes WHERE alm_almacenes.vismov = -1)) "
            End If
            
            If Not IsDate(mFchIniMov) Or mFchIniMov = Empty Then
                mSQLFechaMov = ""
            Else
                mSQLFechaMov = "AND (com_compras.fchdoc)>=CDate('" & mFchIniMov & "') "
            End If
            If IsDate(FechaInicio) And FechaInicio <> Empty Then
                mSQLFechaMov = mSQLFechaMov & "AND (com_compras.fchdoc)>=CDate('" & FechaInicio & "') "
            End If
            If IsDate(FechaFin) And FechaFin <> Empty Then
                mSQLFechaMov = mSQLFechaMov & "AND (com_compras.fchdoc)<=CDate('" & FechaFin & "') "
            End If
            '**********************************************
            ' Exlusione de compras
            mSQLExclude = mSQLExclude & " AND ((alm_inventario.tippro)<> 5) "
            '**********************************************
                
            xCadSQL = xCadSQL _
                + vbCr + "UNION ALL " _
                + vbCr + "SELECT com_compras.id As iddoc, 0 As idmov, 0 As idmovdet, 0 AS idlibrocostotemp, com_comprasdet.iditem, alm_inventario.codpro As coditem, alm_inventario.descripcion As item, com_compras.fchdoc As fchmov, com_compras.numser, com_compras.numdoc, com_compras.numser & '-' & com_compras.numdoc As numdocconcat, com_comprasdet.canpro As cantidad, mae_documento.abrev AS doc, 0 As idtipdocref, -1 AS tipmov, 'I' As tipmovcad, alm_inventario.tippro AS idtippro, mae_tipoproducto.descripcion AS tippro, alm_inventario.idunimed, mae_unidades.abrev AS unimed, com_comprasdet.imptot AS costo, com_comprasdet.preuni AS costounitariopromedio, com_compras.idalm, alm_almacenes.descripcion AS alm, com_comprasdet.preuni As costounitario, com_comprasdet.imptot As costoprimo, 0 As costomod, 0 As costocif, '' AS tipdocref, 0 AS iddocref, 0 AS idtipdocrefanex, '' AS tipdocrefanex, 0 As iddocrefanex, '' As docrefanex, '' As numserref, '' As numdocref, '' As numdocrefconcat " _
                + vbCr + "FROM ((((alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN alm_almacenes ON com_compras.idalm = alm_almacenes.id " _
                + vbCr + "WHERE (((com_compras.tipcom)=1)) " & mSQLIdItem & mSQLIdAlmacen & mSQLFechaMov & mSQLExclude & mSQLIdMovimientoDetalle
        End If
    End If
    
    If FiltraTipoMovimiento = 0 Or FiltraTipoMovimiento = 2 Then
        If FiltraTipoMovimiento = 0 Then xCadSQL = xCadSQL + vbCr + "UNION ALL "
        '**********************
        ' SALIDAS
        '**********************
        If Cad_IdItem_In = "" Then
            mSQLIdItem = ""
        Else
            mSQLIdItem = "AND ((alm_ingresodet.iditem) IN (" & Cad_IdItem_In & ")) "
        End If
        If IdMovimientoDetalle = 0 Then
            mSQLIdMovimientoDetalle = ""
        Else
            mSQLIdMovimientoDetalle = "AND ((alm_ingresodet.idmovdet)=" & IdMovimientoDetalle & ") "
        End If
        If IdAlmacen = 0 Then
            mSQLIdAlmacen = ""
        Else
            mSQLIdAlmacen = "AND ((alm_ingreso.idalm)=" & IdAlmacen & ") "
        End If
        If FiltraAlmacen Then
            mSQLIdAlmacen = mSQLIdAlmacen & "AND ((alm_ingreso.idalm) IN (SELECT alm_almacenes.id FROM alm_almacenes WHERE alm_almacenes.vismov = -1)) "
        End If
        If Not IsDate(mFchIniMov) Or mFchIniMov = Empty Then
            mSQLFechaMov = ""
        Else
            mSQLFechaMov = "AND (alm_ingreso.fching)>=CDate('" & mFchIniMov & "') "
        End If
        If IsDate(FechaInicio) And FechaInicio <> Empty Then
            mSQLFechaMov = mSQLFechaMov & "AND (alm_ingreso.fching)>=CDate('" & FechaInicio & "') "
        End If
        If IsDate(FechaFin) And FechaFin <> Empty Then
            mSQLFechaMov = mSQLFechaMov & "AND (alm_ingreso.fching)<=CDate('" & FechaFin & "') "
        End If
        
        xCadSQL = xCadSQL _
            + vbCr + "SELECT alm_ingreso.id As iddoc, alm_ingreso.id As idmov, alm_ingresodet.idmovdet, con_librocostotemp.idlibrocostotemp, alm_ingresodet.iditem, alm_inventario.codpro As coditem, alm_inventario.descripcion As item, alm_ingreso.fching AS fchmov, alm_ingreso.numser, alm_ingreso.numdoc, alm_ingreso.numser & '-' & alm_ingreso.numdoc As numdocconcat, alm_ingresodet.cantidad, mae_documento.abrev AS doc, alm_ingreso.idtipdocref, alm_ingreso.tipmov, 'S' As tipmovcad, alm_inventario.tippro AS idtippro, mae_tipoproducto.descripcion AS tippro, alm_inventario.idunimed, mae_unidades.abrev AS unimed, (con_librocostotemp.costoprimo + IIf(con_librocostotemp.costomod Is Null Or con_librocostotemp.costomod = 0, 0, con_librocostotemp.costomod) + IIf(con_librocostotemp.costocif Is Null Or con_librocostotemp.costocif = 0, 0, con_librocostotemp.costocif)) AS costo, con_librocostotemp.costounitariopromedio, alm_ingreso.idalm, alm_almacenes.descripcion AS alm  " _
                & ", con_librocostotemp.costounitario, con_librocostotemp.costoprimo, con_librocostotemp.costomod, con_librocostotemp.costocif, mae_documento.descripcion AS tipdocref, alm_ingreso.iddocref, 0 AS idtipdocrefanex, '' AS tipdocrefanex, alm_ingresodet.iddocref As iddocrefanex, '' As docrefanex, CNUMDOCREF.numserref, CNUMDOCREF.numdocref, CNUMDOCREF.numdocrefconcat " _
            + vbCr + "FROM (((alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.idtipdocref = mae_documento.id) LEFT JOIN alm_almacenes ON alm_ingreso.idalm = alm_almacenes.id) LEFT JOIN ((((alm_ingresodet LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) LEFT JOIN con_librocostotemp ON alm_ingresodet.idmovdet = con_librocostotemp.idmovdet) ON alm_ingreso.id = alm_ingresodet.id) LEFT JOIN ( " _
            + vbCr + mSQLDocumento _
            + vbCr + ") AS CNUMDOCREF ON alm_ingreso.id = CNUMDOCREF.id " _
            + vbCr + "WHERE (((alm_ingreso.tipmov)=0)) AND alm_ingreso.estado Not In (1,4,5) AND alm_ingresodet.cantidad<>0 " & mSQLIdItem & mSQLIdAlmacen & mSQLFechaMov & mSQLIdMovimientoDetalle
        
    End If
    
    If xCadSQL <> "" Then
        xCadSQL = "SELECT * FROM ( " & xCadSQL & " ) As MOVDET ORDER BY MOVDET.iditem, MOVDET.fchmov, MOVDET.tipmov, MOVDET.numdocconcat"
    End If
    
    mSQL_MovDetallado = xCadSQL
End Function

Public Function mGetItemCollection(mCollection As Collection, mKey As String) As Object
On Error GoTo ErrHandler
    Set mGetItemCollection = mCollection.Item(mKey)
    Exit Function
ErrHandler:
    Set mGetItemCollection = Nothing
End Function

Public Sub mExportarExcelRecordSet(mRecordSet As ADODB.Recordset)
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object

    
    Dim recArray As Variant
    
    Dim strDB As String
    Dim fldCount As Long
    Dim recCount As Long
    Dim iCol As Long
    Dim iRow As Long
        
    ' Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)
    
    ' Copy field names to the first row of the worksheet
    fldCount = mRecordSet.Fields.Count
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).value = mRecordSet.Fields(iCol - 1).Name
    Next
        
    ' Check version of Excel
    If Val(Mid(xlApp.Version, 1, InStr(1, xlApp.Version, ".") - 1)) > 8 Then
        'EXCEL 2000,2002,2003, or 2007: Use CopyFromRecordset
         
        ' Copy the recordset to the worksheet, starting in cell A2
        xlWs.Cells(2, 1).CopyFromRecordset mRecordSet
        'Note: CopyFromRecordset will fail if the recordset
        'contains an OLE object field or array data such
        'as hierarchical recordsets
        
    Else
        'EXCEL 97 or earlier: Use GetRows then copy array to Excel
    
        ' Copy recordset to an array
        recArray = mRecordSet.GetRows
        'Note: GetRows returns a 0-based array where the first
        'dimension contains fields and the second dimension
        'contains records. We will transpose this array so that
        'the first dimension contains records, allowing the
        'data to appears properly when copied to Excel
        
        ' Determine number of records

        recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array
        

        ' Check the array for contents that are not valid when
        ' copying the array to an Excel worksheet
        For iCol = 0 To fldCount - 1
            For iRow = 0 To recCount - 1
                ' Take care of Date fields
                If IsDate(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                ' Take care of OLE object fields or array fields
                ElseIf IsArray(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = "Array Field"
                End If
            Next iRow 'next record
        Next iCol 'next field
            
        ' Transpose and Copy the array to the worksheet,
        ' starting in cell A2
        xlWs.Cells(2, 1).Resize(recCount, fldCount).value = _
            TransposeDim(recArray)
    End If

    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
  
    ' Display Excel and give user control of Excel's lifetime
    xlApp.Visible = True
    xlApp.UserControl = True
    
    ' Close ADO objects
    mRecordSet.Close
    Set mRecordSet = Nothing
    
    ' Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing

End Sub

Private Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = v(Y, X)
        Next Y
    Next X
    
    TransposeDim = tempArray
End Function

Public Function mMesCerradoOpcion(IdMes As Integer, IdOpcion As Long, Optional mConexion As ADODB.Connection) As Boolean
    Dim database As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones
    Dim mRecord As New ADODB.Recordset
    
On Error GoTo BloqueError
    If Not mConexion Is Nothing Then Set database.Connection = mConexion
    database.CommandText = "SELECT var_cierre.estado " _
            + vbCr + "FROM var_cierre " _
            + vbCr + "WHERE (((var_cierre.idmes)=" & IdMes & " ) AND ((var_cierre.idform)=" & IdOpcion & "))"
    
    Set mRecord = database.GetRecordset
    If mRecord.RecordCount = 0 Then mMesCerradoOpcion = False: Exit Function
    mRecord.MoveFirst
    mMesCerradoOpcion = Not CBool(mRecord("estado"))
    Exit Function
    
BloqueError:
    mMesCerradoOpcion = False
    mMostrarMensajeError Err.Description, "[MesCerradoOpcion] " & Err.Source
End Function

Public Function mRetornarMesFecha(Fecha As Date) As Integer
    mRetornarMesFecha = Month(Fecha)
End Function

Public Function mRetornarAnhoFecha(Fecha As Date) As Integer
    mRetornarAnhoFecha = Year(Fecha)
End Function

Public Function mRetornarPrimerDiaMes(Fecha As Date) As Date
    mRetornarPrimerDiaMes = CDate("01/" & mRetornarMesFecha(Fecha) & "/" & mRetornarAnhoFecha(Fecha) & "")
End Function

Public Function mRetornarUltimoDiaMes(Fecha As Date) As Date
    Dim Mestrabajo As Integer
    Dim AnhoTrabajo As Integer
    
    AnhoTrabajo = mRetornarAnhoFecha(Fecha)
    Mestrabajo = mRetornarMesFecha(Fecha)
    If Mestrabajo = 12 Then
        Mestrabajo = 1
        AnhoTrabajo = AnhoTrabajo + 1
    Else
        Mestrabajo = Mestrabajo + 1
    End If
    mRetornarUltimoDiaMes = mRetornarPrimerDiaMes(CDate("01/" & Mestrabajo & "/" & AnhoTrabajo & "")) - 1
End Function

Public Sub mLlenarCombo(cb, Rst As ADODB.Recordset, ValueColumn As String, BoundColumn As String)
'    cb.Clear
'    Set cb.RowSource = Rst
'    cb.BoundColumn = "categoryid"
'    cb.ListField = "categoryname"
End Sub

Public Function mCompararConCriterio(Numero1 As Double, Numero2 As Double, Optional NumeroDecimales As Integer = 4) As Boolean
    Dim mNumeroComparar As Double
    Dim mNumeroFinal As Double
    
    mNumeroComparar = 1 / (10 ^ (NumeroDecimales - 1))
    mNumeroFinal = Format(Abs(Numero1 - Numero2), "###,###,##0.0000")
    
    If mNumeroFinal > mNumeroComparar Then
        mCompararConCriterio = False
    Else
        mCompararConCriterio = True
    End If
End Function
