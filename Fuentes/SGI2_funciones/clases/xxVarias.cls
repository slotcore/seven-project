VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Varias"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TIPO_CAMBIO
    'VALORES PARA BUSCAR EL TIPO DE CAMBIO
    compra = 1
    Venta = 2
End Enum

Public Enum TIPO_MOVIMIENTO
    'VALORES PARA BUSCAR EL TIPO DE CAMBIO
    origen = 1
    destino = 2
End Enum

Function DetallarModulo(IdOriDes As Integer, TipoMov As TIPO_MOVIMIENTO, Con As ADODB.Connection) As Boolean
    Dim Rst As New ADODB.Recordset
    If TipoMov = origen Then
        RST_Busq Rst, "SELECT * FROM tes_origen WHERE id = " & IdOriDes & "", Con
    Else
        RST_Busq Rst, "SELECT * FROM tes_destino WHERE id = " & IdOriDes & "", Con
    End If
    
    If Rst.RecordCount <> 0 Then
        DetallarModulo = NulosN(Rst("detallar"))
    End If
    Set Rst = Nothing
End Function


Function SeleccionaMes(Coneccion As ADODB.Connection) As Integer
    Set xCon = Coneccion
    FrmSeteaMes.Show vbModal
    SeleccionaMes = xMes
End Function


Function NomMes(nummes As Integer) As String
    'Valor Requerido un dato de tipo integer
    'Devuelve el nombre del Mes
    Select Case nummes
        Case 1
            NomMes = "Enero"
        Case 2
            NomMes = "Febrero"
        Case 3
            NomMes = "Marzo"
        Case 4
            NomMes = "Abril"
        Case 5
            NomMes = "Mayo"
        Case 6
            NomMes = "Junio"
        Case 7
            NomMes = "Julio"
        Case 8
            NomMes = "Agosto"
        Case 9
            NomMes = "Setiembre"
        Case 10
            NomMes = "Octubre"
        Case 11
            NomMes = "Noviembre"
        Case 12
            NomMes = "Diciembre"
    End Select
End Function

Function HallaTipoCambio(Dia As String, Moneda As Integer, Tipo As TIPO_CAMBIO, Coneccion As ADODB.Connection) As Double
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT con_tc.fecha, con_tc.idmon, con_tc.impcom,con_tc.impven From con_tc " _
        & " WHERE (((con_tc.fecha)=CDate('" & Dia & "')) AND ((con_tc.idmon)=" & Moneda & "))", Coneccion

    If Rst.RecordCount <> 0 Then
        If Tipo = compra Then
            HallaTipoCambio = NulosN(Rst("impcom"))
        Else
            HallaTipoCambio = NulosN(Rst("impven"))
        End If
    Else
        HallaTipoCambio = 0
    End If
    Set Rst = Nothing
End Function

Function DevuelveNumAsiento(Libro As Integer, IdMovimiento As Double, Mes As Integer, Coneccion As ADODB.Connection) As String
    Dim RstDia As New ADODB.Recordset
    
    RST_Busq RstDia, "SELECT * FROM con_diario WHERE idmes = " & Mes & " AND " _
        & " idlib = " & Libro & " AND idmov = " & IdMovimiento & "", Coneccion
            
    If RstDia.RecordCount <> 0 Then
        DevuelveNumAsiento = RstDia("numasi") & ""
    Else
        DevuelveNumAsiento = ""
    End If
    Set RstDia = Nothing
End Function

Function NuevoNumAsiento(Libro As Integer, Mes As Integer, Coneccion As ADODB.Connection) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT TOP 1 con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)=" & Libro & ")) ORDER BY numasi DESC ", Coneccion
    
    If Rst.RecordCount = 0 Then
        NuevoNumAsiento = "0001"
    Else
        Rst.MoveLast
        NuevoNumAsiento = Format(NulosN(Rst("numasi")) + 1, "0000")
    End If
    Set Rst = Nothing
    
End Function

Function HallaNumdocVenta(TipoDoc As Integer, NumSerie As String, Coneccion As ADODB.Connection) As String
    Dim Rst As New ADODB.Recordset
    Dim xNum As Double
    RST_Busq Rst, "SELECT top 1 vta_ventas.tipdoc, vta_ventas.numser, vta_ventas.numdoc From vta_ventas " _
        & " WHERE (((vta_ventas.tipdoc)=" & TipoDoc & ") AND ((vta_ventas.numser)='" & NumSerie & "')) " _
        & " ORDER BY vta_ventas.numdoc DESC ", Coneccion

    If Rst.RecordCount <> 0 Then
        Rst.MoveLast
        xNum = NulosN(Rst("numdoc"))
        xNum = xNum + 1
        HallaNumdocVenta = Format(xNum, "0000000000")
    Else
        xNum = 1
        HallaNumdocVenta = Format(xNum, "0000000000")
    End If
    Set Rst = Nothing
End Function

Function HallaNumGuia(NumSerie As String, Coneccion As ADODB.Connection) As String
    Dim Rst As New ADODB.Recordset
    Dim xNum As Double
    RST_Busq Rst, "SELECT top 1 numdoc FROM vta_guia WHERE numser =  '" & NumSerie & "' ORDER BY numdoc DESC", Coneccion

    If Rst.RecordCount <> 0 Then
        Rst.MoveLast
        xNum = Val(Rst("numdoc"))
        xNum = xNum + 1
        HallaNumGuia = Format(xNum, "0000000000")
    Else
        xNum = 1
        HallaNumGuia = Format(xNum, "0000000000")
    End If
    Set Rst = Nothing
End Function

Sub CargaDatosEmpresa(xCon As ADODB.Connection, NomEmp As String, NumRUC As String)
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

Function TotalProducido(IdProducto As Integer, FchInicio As String, FchFinal As String, xCon As ADODB.Connection) As Double
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim xTotal As Double
    Dim xCad As String
    
    'CARGAMOS TODAS LAS ENTRADAS
    xCad = "SELECT Sum(com_comprasdet.canpro) AS total FROM alm_inventario RIGHT JOIN (mae_prov LEFT JOIN ((mae_documento RIGHT JOIN " _
        & " (com_compras LEFT JOIN  con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN com_comprasdet " _
        & " ON com_compras.id = com_comprasdet.idcom)  ON mae_prov.id = com_compras.idpro) ON alm_inventario.id = com_comprasdet.iditem  " _
        & " WHERE (((com_comprasdet.iditem)=" & IdProducto & ") AND ((com_compras.fchdoc)>=CDate('" & FchInicio & "') And (com_compras.fchdoc)<=CDate('" & FchFinal & "')) " _
        & " AND ((con_tc.idmon)=2))" _
        & " Union " _
        & " SELECT alm_ingresodet.cantidad AS total FROM (alm_ingreso LEFT JOIN mae_documento ON alm_ingreso.tipdoc = mae_documento.id) LEFT JOIN " _
        & " (alm_ingresodet  LEFT JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id  " _
        & " WHERE (((alm_ingresodet.iditem)=" & IdProducto & ") AND ((alm_ingreso.fching)>=CDate('" & FchInicio & "') And (alm_ingreso.fching)<=CDate('" & FchFinal & "')) " _
        & " AND ((alm_ingreso.tipmov)=-1)) " _
        & " Union " _
        & " SELECT Sum([cantidad]) AS total FROM pro_produccion LEFT JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        & " WHERE (((pro_producciondet.iditem)=" & IdProducto & ") AND ((pro_produccion.dia)>=CDate('" & FchInicio & "') And (pro_produccion.dia)<=CDate('" & FchFinal & "')))"

    RST_Busq Rst, xCad, xCon
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            xTotal = xTotal + NulosN(Rst("total"))
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        
        TotalProducido = xTotal
    Else
        TotalProducido = 0
    End If
End Function

Function CambioOpcionLiro(idFormato As Integer, Con As ADODB.Connection, TipoFormatos As Integer) As Boolean
    Set xCon = Con
    xIdFormatos = TipoFormatos
    FrmConfFormato.LblIdLibro.Caption = idFormato
    FrmConfFormato.Show vbModal
    If Grabo = True Then
        CambioOpcionLiro = True
    Else
        CambioOpcionLiro = False
    End If
End Function

Function GrabarOperacion(IdUsuario As Integer, IdFormulario As Integer, IdOperacion As Integer, HorIni As Date, HorFin As Date, Fecha As Date, xCon As ADODB.Connection, IdMovimiento As Double)
    Dim xId As Integer
    xId = HallaCodigoTabla("var_edicion", xCon, "id")
    
    xCon.Execute "INSERT INTO var_edicion ( id, idusuario, idform, idoperacion, horini, horfin, fchope, idmov )" _
        & " SELECT " & xId & " AS Expr1, " & IdUsuario & " AS Expr2, " & IdFormulario & " AS Expr3, " & IdOperacion & " AS Expr4, " _
        & " '" & HorIni & "' AS Expr5, '" & HorFin & "' AS Expr6, '" & CDate(Fecha) & "' AS Expr7, " & IdMovimiento & " as Expr8 "
End Function

Sub VerMovimientos(CadenaSQL As String, Con As ADODB.Connection)
    Set xCon = Con
    FrmVerMovimiento.TxtSQL.Text = CadenaSQL
    FrmVerMovimiento.Show vbModal
End Sub


Sub VerMovimientos1(Idform As Double, Idmov As Double, Con As ADODB.Connection)
    '--30/04/09
    
    Set xCon = Con
    Dim CadenaSQL As String
    
    CadenaSQL = "SELECT var_edicion.*, UCase(mae_usuarios!ape) & ', ' & mae_usuarios!nom AS apenom, mae_formularios.descripcion AS formulario, " _
            & " mae_tipooperacion.descripcion AS operacion FROM ((var_edicion LEFT JOIN mae_usuarios ON var_edicion.idusuario = mae_usuarios.id) " _
            & " LEFT JOIN mae_formularios ON var_edicion.idform = mae_formularios.id) LEFT JOIN mae_tipooperacion ON var_edicion.idoperacion = mae_tipooperacion.id " _
            & " WHERE (((var_edicion.idform)=" & Idform & ") AND ((var_edicion.idmov)=" & Idmov & "))"
            
    FrmVerMovimiento.TxtSQL.Text = CadenaSQL
    FrmVerMovimiento.Show vbModal
End Sub





