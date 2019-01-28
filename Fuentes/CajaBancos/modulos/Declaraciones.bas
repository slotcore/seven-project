Attribute VB_Name = "Declaraciones"
'*****************************************************************************************************
'* Nombre Archivo   : DECLARACIONES.BAS
'* Tipo             : MODULO
'* Descripcion      : MODULO DONDE SE DECLARAN LAS VARIABLES PUBLICAS QUE SE UTILIZARAN EN LA CLASE
'*                    ASI COMO LA DEFINICION DE ALGUNAS FUNCIONES PROPIAS DE LA CLASE
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 12/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public xCon As New ADODB.Connection             ' ALMACENA LA CONECCION ACTUAL A LA BASE DE DATOS
Public xTitulo As String                        ' TITULO DE LA CLASE
Public NomEmp, NumRuc As String                 ' ALMACENA EL NUMBRE DE LA EMPRESA Y EL NUMERO DE RUC
Public AnoTra As String                         ' ALAMCENA EL AÑO DE TRABAJO
Public xMes As Integer                          ' ALMACENA EL MES DE TRABJO ACTUAL
Public xIdUsuario As Integer                    ' ALMACENA EL ID DEL USUARIO

Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Declare Sub InitCommonControls Lib "Comctl32" ()

Public T_ToolTipText() As String

Dim oPDF As cPDF
Dim xFilaInicial As Integer
Dim xNumPag As Integer

Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)


'*****************************************************************************************************
'* Nombre Archivo   : CargaDatos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA EMPRESA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub CargaDatos()
    Dim rst As New ADODB.Recordset
    
    RST_Busq rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = rst("nomemp")
    NumRuc = rst("numruc")
    AnoTra = rst("anotra")
    Set rst = Nothing
End Sub

Sub ImprimirOperacion(xTipo As Integer, xIdMov As Double, xFchIni As String, xCon As ADODB.Connection)
'Modificado 26/01/11 Johan Castro
'Agregar paramentro xTipo que indique si es 1 = Ingresos o 2 = Egresos
'Agregar parametro xIdMov para indicar el codigo del movimiento a imprimir.
'Agregar variable nSQLFiltro para filtrar por movimiento.

    
    Dim xForm As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim nSQLFiltro As String '--Almacenara el filtro por movimiento
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(8, 5) As String
    
    xCampos(0, 0) = "Nº Registro":   xCampos(0, 1) = "registro":         xCampos(0, 2) = "950":    xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    xCampos(1, 0) = "Fch. Mov":      xCampos(1, 1) = "fchope":         xCampos(1, 2) = "1000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "Importe":       xCampos(2, 1) = "importe":        xCampos(2, 2) = "1000":    xCampos(2, 3) = "N":    xCampos(2, 4) = "N"
    xCampos(3, 0) = "M":             xCampos(3, 1) = "simbolo":        xCampos(3, 2) = "450":     xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    xCampos(4, 0) = "Origen":        xCampos(4, 1) = "origen":        xCampos(4, 2) = "2500":    xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
    xCampos(5, 0) = "TD":            xCampos(5, 1) = "abrev":          xCampos(5, 2) = "600":     xCampos(5, 3) = "C":    xCampos(5, 4) = "N"
    xCampos(6, 0) = "Nº Documento":  xCampos(6, 1) = "numerodoc":         xCampos(6, 2) = "1200":    xCampos(6, 3) = "C":    xCampos(6, 4) = "N"
    xCampos(7, 0) = "Glosa":         xCampos(7, 1) = "glosa":          xCampos(7, 2) = "2200":    xCampos(7, 3) = "C":    xCampos(7, 4) = "N"
    
    
    If xIdMov <> 0 Then
        nSQLFiltro = " and tes_caja.id=" & xIdMov & " "
    Else
        nSQLFiltro = " and tes_caja.fchreg=CDate('" & xFchIni & "') "
    End If
    
'    xForm.SQLCad = "SELECT 0 as xsel, tes_caja.id, tes_caja.fchreg, tes_caja.tipmov, tes_caja.fchope & '' AS fchope, tes_caja.numreg, tes_caja.glosa, mae_moneda.simbolo, " _
'        & " tes_cajaorigendet.iddoc, tes_documentos.abrev, tes_documentos.descripcion AS descdoc, tes_origen.descripcion AS descori, " _
'        & " IIf(IsNull(tes_cajaorigendet!numser)=-1,tes_cajaorigendet!numdoc,tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc) AS numdoc, iif(tes_caja.tipmov=1,'Ingreso','Egreso') AS tipo, " _
'        & " tes_cajaori.importe & '' AS importe, tes_caja.idmon, tes_documentos.abrev AS desdocabre, IIf([tes_caja].[numreg] Is Null,'',Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS registro " _
'        & " FROM (((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN (tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) " _
'        & " ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) " _
'        & " ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id " _
'        & " WHERE tes_caja.tipmov=" & xTipo & nSQLFiltro & " ORDER BY tes_caja.numreg DESC;"
        
        
    'consulta para obtener listado de
    xForm.SQLCad = "SELECT 0 as xsel,tes_caja.id, IIf(tes_caja.tipmov=1,'Ingreso','Egreso') AS Tipo1, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([tes_caja].[numreg],4) AS registro, " _
        + vbCr + " tes_origen.descripcion AS origen, mae_moneda.simbolo, tes_cajaori.importe, tes_caja.fchope & '' as fchope, tes_documentos.abrev, IIf(IsNull(tes_cajaorigendet!numser)=-1,tes_cajaorigendet!numdoc,tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc) AS numerodoc, tes_caja.glosa " _
        + vbCr + " FROM ((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) INNER JOIN (tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) ON tes_caja.id = tes_cajaori.idtes) " _
        + vbCr + " LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id " _
        + vbCr + " WHERE tes_caja.tipmov=" & xTipo & nSQLFiltro & " ORDER BY tes_caja.numreg DESC;"
    
    
    If xIdMov <> 0 Then
        RST_Busq xRs, xForm.SQLCad, xCon
        '--Utilizar esta utilidad de posicionarse en la ultima fila para mostrar reporte, caso contrario muestra reporte duplicado
        If xRs.RecordCount <> 0 Then xRs.MoveLast
        '--Imprimir registro
        Imprimir xRs, 1
    Else
        xForm.Titulo = "Operaciones a Imprimir"
        Set xForm.Coneccion = xCon
        Set xRs = Nothing
        Set xRs = xForm.Seleccionar(xCampos)
        
        If xRs.State = 1 Then
                        
            If xRs.RecordCount <> 0 Then
                xRs.MoveFirst
                Imprimir xRs, 1
            End If
        End If
    
    End If
     
Set xForm = Nothing
Set xRs = Nothing

End Sub

Function PreparaRSTImp() As ADODB.Recordset
'26/01/11 Johan Castro
'       Agregar campo tipcam

    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "abrev":        xCampos(0, 1) = "C":      xCampos(0, 2) = "10"
    xCampos(1, 0) = "acuenta":      xCampos(1, 1) = "D":      xCampos(1, 2) = "2"
    xCampos(2, 0) = "fchdoc":       xCampos(2, 1) = "C":      xCampos(2, 2) = "10"
    xCampos(3, 0) = "simbolo":      xCampos(3, 1) = "C":      xCampos(3, 2) = "10"
    xCampos(4, 0) = "numdoc":       xCampos(4, 1) = "C":      xCampos(4, 2) = "20"
    xCampos(5, 0) = "numruc":       xCampos(5, 1) = "C":      xCampos(5, 2) = "20"
    xCampos(6, 0) = "nombre":       xCampos(6, 1) = "C":      xCampos(6, 2) = "100"
    xCampos(7, 0) = "imptot":       xCampos(7, 1) = "D":      xCampos(7, 2) = "20"
    xCampos(8, 0) = "idtes":        xCampos(8, 1) = "N":      xCampos(8, 2) = "2"
    xCampos(9, 0) = "tipcam":       xCampos(9, 1) = "D":      xCampos(9, 2) = "10"
    
    
    Set PreparaRSTImp = xFun.CrearRstTMP(xCampos)
    PreparaRSTImp.Open
End Function

Sub CrearCabeceraVS()
    Dim xCad As String
    
    FrmPrinter.VS.TextAlign = taLeftTop
    FrmPrinter.VS.FontName = "Courier New"
    FrmPrinter.VS.FontBold = True
    FrmPrinter.VS.FontSize = 9
    
    FrmPrinter.VS.CurrentX = 1000:      FrmPrinter.VS.CurrentY = 1200
    FrmPrinter.VS.Paragraph = "EMPRESA   : " & NomEmp
    
    FrmPrinter.VS.CurrentX = 8800:      FrmPrinter.VS.CurrentY = 1200
    FrmPrinter.VS.Paragraph = "FECHA     : " & Format(Date, "dd/mm/yy")
    
    FrmPrinter.VS.CurrentX = 1000:      FrmPrinter.VS.CurrentY = 1400
    FrmPrinter.VS.Paragraph = "Nº R.U.C. :" & NumRuc
    
    FrmPrinter.VS.CurrentX = 8800:      FrmPrinter.VS.CurrentY = 1400
    FrmPrinter.VS.Paragraph = "Nº Pagina     : " & "0001"
    
    FrmPrinter.VS.DrawLine 1000, 1650, 11000, 1650
End Sub

Sub Imprimir(RstOpe As ADODB.Recordset, Opcion As Integer)
'Agregar campo xtipo en consulta RstCab

    'RstOpe = que contiene todos los movimientos que se imprimiram
    ' OPCION = 1 SE ABRE EL DOCUMENTO PDF
    ' OPCION = 2 SE ENVIA POR CORREO EL ARCHIVO PDF
    
    Dim Li As Integer
    Dim strSource As String
    Dim xArea, xEmp, xDir, xCuerpo, xCad  As String
    Dim xEmpleado As String
    Dim Pagina As Integer
    Dim Lineas As Integer
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDetDes As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim RstOri As New ADODB.Recordset
    
    Dim A, B, C As Integer
    Dim xFilaAct As Integer
    
    xNumPag = 0

On Error GoTo Cerrado
    
    With FrmPrinter.VS
    
        FrmPrinter.VS.StartDoc
        CrearCabeceraVS
        xFilaInicial = 1700
        
'        RstOpe.MoveFirst
        For A = 1 To RstOpe.RecordCount
            Set RstCab = Nothing
            Set RstDet = Nothing
            Set RstDetDes = Nothing
            Set RstDia = Nothing
            Set RstOri = Nothing
            
            RST_Busq RstCab, "SELECT tes_caja.id, tes_cajaori.idori, tes_origen.descripcion, mae_banconumcta.numcue, mae_bancos.descripcion AS nomban, mae_moneda.simbolo, " _
                & " tes_caja.fchope, con_tc.impven, tes_cajaori.importe, tes_caja.glosa, Mid([numreg],1,2) & [mae_libros]![codsun] & Mid([numreg],3,4) AS xnumreg,iif(tes_caja.tipmov=1,'INGRESO','EGRESO') AS xtipo " _
                & " FROM ((((((tes_caja INNER JOIN tes_cajaori ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN tes_origen ON tes_cajaori.idori = tes_origen.id) " _
                & " LEFT JOIN mae_banconumcta ON tes_cajaori.idbcocta = mae_banconumcta.id) LEFT JOIN mae_bancos ON mae_banconumcta.idban = mae_bancos.id) LEFT JOIN mae_moneda " _
                & " ON tes_caja.idmon = mae_moneda.id) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id " _
                & " WHERE (((tes_caja.id)=" & RstOpe("id") & "))", xCon
        
            RST_Busq RstOri, "SELECT tes_cajaorigendet.idtes, tes_cajaorigendet.numdoc, tes_mediopago.descripcion AS descmedpag, tes_documentos.descripcion AS descdoc" _
                & " FROM (tes_cajaorigendet LEFT JOIN tes_mediopago ON tes_cajaorigendet.idmedpag = tes_mediopago.id) LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id " _
                & " WHERE (tes_cajaorigendet.idtes=" & RstOpe("id") & ")", xCon
        
            
            '**************************************************************************************************
            ' CARGAMOS LOS DESTINOS
'            Dim RstDes As New ADODB.Recordset
            Dim RstTem As New ADODB.Recordset
'            RST_Busq RstDet, "SELECT tes_cajadestino.* From tes_cajadestino WHERE (((tes_cajadestino.idtes)=" & RstOpe("id") & ")) ", xCon

            ' CREAMOS EL RECORDSET PARA ALMACENAR LOS DATOS QUE SE IMPRIMIRAN
            Set RstDetDes = PreparaRSTImp
            
'            Set RstDetDes.ActiveConnection = Nothing
'
'            If RstDet.RecordCount <> 0 Then
'                RstDet.MoveFirst
'                For C = 1 To RstDet.RecordCount
'                    If RstDet("idmod") = 1 Then   ' SI ES UNA OPERACION CON COMPRAS
'                        RST_Busq RstTem, "SELECT 0 as tipcam, mae_documento.abrev, tes_cajadestinodet.acuenta, com_compras.fchdoc, mae_moneda.simbolo, " _
'                            & " [com_compras]![numser] & '-' & [com_compras]![numdoc] AS numdoc, mae_prov.numruc, mae_prov.nombre, com_compras.imptot, " _
'                            & " tes_cajadestinodet.idtes FROM (tes_cajadestinodet LEFT JOIN ((com_compras LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) " _
'                            & " LEFT JOIN mae_moneda ON com_compras.idmon = mae_moneda.id) ON tes_cajadestinodet.iddoc = com_compras.id) LEFT JOIN mae_prov " _
'                            & " ON com_compras.idpro = mae_prov.id WHERE (((tes_cajadestinodet.idtes)=" & RstDet("idtes") & "))", xCon
'                    End If
'                    If NulosN(RstDet("idmod")) = 0 Or NulosN(RstDet("idmod")) = 6 Then
'                        RST_Busq RstTem, "SELECT 0 as tipcam,'' AS abrev, tes_cajadestino.importe AS acuenta, tes_caja.fchope AS fchdoc, mae_moneda.simbolo, '' AS numdoc, " _
'                            & " '' AS numruc, tes_destino.descripcion AS nombre, 0 AS imptot, tes_cajadestino.idtes FROM ((tes_caja RIGHT JOIN tes_cajadestino " _
'                            & "  ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN tes_destino " _
'                            & " ON tes_cajadestino.iddes = tes_destino.id WHERE (((tes_cajadestino.idtes)=" & RstDet("idtes") & ") AND ((tes_cajadestino.iddes)=" & RstDet("iddes") & "))", xCon
'                    End If
'
'                    If RstDet("idmod") = 9 Then   ' SI ES UNA OPERACION CON HONORARIOS
'                        RST_Busq RstTem, "SELECT 0 as tipcam, mae_documento.abrev, tes_cajadestinodet.acuenta, com_honorarios.fchdoc, mae_moneda.simbolo, " _
'                            & " [com_honorarios]![numser] & '-' & [com_honorarios]![numdoc] AS numdoc, mae_prov.numruc, mae_prov.nombre, com_honorarios.imptot, " _
'                            & " tes_cajadestinodet.idtes FROM (tes_cajadestinodet LEFT JOIN ((com_honorarios LEFT JOIN mae_documento ON com_honorarios.tipdoc = mae_documento.id) " _
'                            & " LEFT JOIN mae_moneda ON com_honorarios.idmon = mae_moneda.id) ON tes_cajadestinodet.iddoc = com_honorarios.id) LEFT JOIN mae_prov " _
'                            & " ON com_honorarios.idpro = mae_prov.id WHERE (((tes_cajadestinodet.idtes)=" & RstDet("idtes") & "))", xCon
'                    End If
'                    If RstDet("idmod") = 2 Then   ' SI ES UNA OPERACION CON VENTAS
'                        RST_Busq RstTem, "SELECT 0 as tipcam, mae_documento.abrev, tes_cajadestinodet.acuenta, vta_ventas.fchdoc, mae_moneda.simbolo, " _
'                            & " [vta_ventas]![numser] & '-' & [vta_ventas]![numdoc] AS numdoc, mae_cliente.numruc, mae_cliente.nombre, vta_ventas.imptotdoc as imptot, " _
'                            & " tes_cajadestinodet.idtes FROM (tes_cajadestinodet LEFT JOIN ((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
'                            & " LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) ON tes_cajadestinodet.iddoc = vta_ventas.id) LEFT JOIN mae_cliente " _
'                            & " ON vta_ventas.idcli = mae_cliente.id WHERE (((tes_cajadestinodet.idtes)=" & RstDet("idtes") & "))", xCon
'                    End If
'
'                    If RstTem.State <> 0 Then
'                        If RstTem.RecordCount <> 0 Then
'                            RstTem.MoveFirst
'                            For B = 1 To RstTem.RecordCount
'                                RstDetDes.AddNew
'                                RstDetDes("abrev") = RstTem("abrev")
'                                RstDetDes("acuenta") = RstTem("acuenta")
'                                RstDetDes("fchdoc") = RstTem("fchdoc")
'                                RstDetDes("simbolo") = RstTem("simbolo")
'                                RstDetDes("numdoc") = NulosC(RstTem("numdoc"))
'                                RstDetDes("numruc") = RstTem("numruc")
'                                RstDetDes("nombre") = NulosC(RstTem("nombre"))
'                                RstDetDes("imptot") = RstTem("imptot")
'                                RstDetDes("idtes") = RstTem("idtes")
'                                RstDetDes("tipcam") = RstTem("tipcam")
'
'                                RstTem.MoveNext
'                                If RstTem.EOF = True Then Exit For
'
'                            Next B
'                        End If
'                    End If
'
'                    RstDet.MoveNext
'                    If RstDet.EOF = True Then Exit For
'                Next C
'            End If
            
'            RST_Busq RstDia, "SELECT con_planctas.cuenta, con_planctas.descripcion, con_diario.impdebsol, con_diario.imphabsol, con_diario.impdebdol, " _
'                & " con_diario.imphabdol, con_diario.idmov, con_diario.idlib, mae_documento.abrev, con_diario.rnumerodoc, mae_moneda.simbolo, con_diario.tc " _
'                & " FROM ((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) " _
'                & " LEFT JOIN mae_moneda ON con_diario.ridmon = mae_moneda.id WHERE (((con_diario.idmov)=" & RstOpe("id") & ") AND ((con_diario.idlib)=6))", xCon
            
            '**************************************************************************************************
            Dim nSQL As String
            
            nSQL = "SELECT con_diario.idmov AS idtes, Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null,'',mae_libros.codsun) & con_diario.numasi AS registro, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, con_diario.fchdoc AS fchope, " _
                + vbCr + " IIf(con_diario.aplicatc=-1,con_diario.tc,IIf(con_tc.impven Is Null,0,con_tc.impven)) AS tipcam, mae_moneda.simbolo AS monope, con_diario.rglosaope AS glosaope, con_diario.rregistro AS registroref, con_diario.rnumerodoc AS numdoc, " _
                + vbCr + " IIf(con_diario.ridtipper=1,mae_prov.numruc,IIf(con_diario.ridtipper=2,mae_cliente.numruc,IIf(con_diario.ridtipper=3,pla_empleados.numdoc,IIf(con_diario.ridtipper=5,mae_bancos.numruc,'')))) AS numruc, " _
                + vbCr + " IIf(con_diario.ridtipper=0,tes_documentos.abrev,mae_documento.abrev) AS tdocdesc, con_diario.rfchope AS fchdoc, " _
                + vbCr + " IIf(con_diario.ridtipper=1,mae_prov.nombre,IIf(con_diario.ridtipper=2,mae_cliente.nombre,IIf(con_diario.ridtipper=3,pla_empleados.nombre,IIf(con_diario.ridtipper=5,mae_bancos.descripcion,'')))) AS apenom, " _
                + vbCr + " con_diario.impdebsol, con_diario.imphabdol, con_diario.impdebdol, con_diario.imphabsol, " _
                + vbCr + " con_diario.impdebsol+con_diario.imphabdol+con_diario.impdebdol+con_diario.imphabsol AS impreal, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol1, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol1, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol1, " _
                + vbCr + " IIf(con_diario.idmon = 2, con_diario.imphabdol, IIf(TipCam = 0 Or con_diario.imphabsol = 0, 0, (con_diario.imphabsol / TipCam))) As imphaberdol1 " _
                + vbCr + " FROM ((((pla_empleados RIGHT JOIN (mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) " _
                + vbCr + " LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper) ON pla_empleados.id = con_diario.ridper) LEFT JOIN tes_documentos ON con_diario.rtipdoc = tes_documentos.id) LEFT JOIN mae_bancos ON con_diario.ridper = mae_bancos.id) LEFT JOIN mae_documento AS mae_documento_1 ON con_diario.rtipdoc1 = mae_documento_1.id) LEFT JOIN mae_moneda AS mae_moneda_1 ON con_diario.ridmon = mae_moneda_1.id " _
                + vbCr + " WHERE (((con_diario.idmov)=" & RstOpe("id") & ") AND ((con_diario.idlib)=6)) "

            RST_Busq RstTem, nSQL, xCon
            
            '--filtrando solo el destino
            If RstCab("xtipo") = "INGRESO" Then
                RstTem.Filter = "impdebsol=0 and impdebdol=0"
            Else
                RstTem.Filter = "imphabsol=0 and imphabdol=0"
            End If
            
            If RstTem.RecordCount <> 0 Then
                RstTem.MoveFirst
                Do While Not RstTem.EOF
                    RstDetDes.AddNew
                    RstDetDes("abrev") = NulosC(RstTem("tdocdesc"))
                    RstDetDes("acuenta") = NulosN(RstTem("impreal"))
                    RstDetDes("fchdoc") = NulosC(RstTem("fchdoc"))
                    RstDetDes("simbolo") = NulosC(RstTem("monope"))
                    RstDetDes("numdoc") = NulosC(RstTem("numdoc"))
                    RstDetDes("numruc") = NulosC(RstTem("numruc"))
                    If NulosC(RstTem("apenom")) = "" Then
                        RstDetDes("nombre") = NulosC(RstTem("ctadesc"))
                    Else
                        RstDetDes("nombre") = NulosC(RstTem("apenom"))
                    End If
                    RstDetDes("imptot") = NulosN(RstTem("impreal"))
                    RstDetDes("idtes") = RstTem("idtes")
                    RstDetDes("tipcam") = NulosN(RstTem("tipcam"))
                    
                    RstTem.MoveNext
                    
                Loop
            End If
                
            RstTem.Filter = ""
            RstTem.Sort = "ctanum,apenom,numdoc"
            
            'xFilaAct = CabeceraVoucher(xFilaInicial, RstCab, RstOri, RstDetDes, RstDia)
            
            xFilaAct = CabeceraVoucher(xFilaInicial, RstCab, RstOri, RstDetDes, RstTem)
            
            xFilaInicial = xFilaAct
            RstOpe.MoveNext
            If RstOpe.EOF = True Then Exit For
        Next A
        
        FrmPrinter.VS.EndDoc
    End With
    
    FrmPrinter.Caption = "Imprimiendo Voucher de " & LCase(RstCab("xtipo"))
    
    
    FrmPrinter.Show
    Exit Sub

Cerrado:
    If Err.Number = 1 Then
    End If
End Sub

Function LeerFila(xFila As Integer, xFontSize As Variant) As Integer
    If xFila >= 16000 Then
        CrearCabeceraVS
        LeerFila = 1700
        'xFilaInicial = 1700
        FrmPrinter.VS.NewPage
        CrearCabeceraVS
        FrmPrinter.VS.FontSize = xFontSize
    Else
        LeerFila = xFila
    End If
End Function

Function CabeceraVoucher(xFila As Integer, xRstCab As ADODB.Recordset, xRstDet As ADODB.Recordset, xRstDetDes As ADODB.Recordset, RstDia As ADODB.Recordset) As Integer
    Dim CualquierCosa As String
    
    FrmPrinter.VS.BrushColor = &H80000005
    FrmPrinter.VS.FontSize = 11
    FrmPrinter.VS.TextAlign = taCenterMiddle
    FrmPrinter.VS.TextBox "COMPROBANTE DE CAJA - BANCOS  (" & UCase(xRstCab("xtipo")) & " )", 1000, xFila, 7500, 450, True, False, True
    
    FrmPrinter.VS.FontSize = 8
    FrmPrinter.VS.TextBox "Nº Registro", 8520, xFila, 2440, 200, True, False, True
    FrmPrinter.VS.FontSize = 10
    FrmPrinter.VS.TextBox xRstCab("xnumreg"), 8520, xFila + 200, 2440, 250, True, False, True
    
    FrmPrinter.VS.FontSize = 7
    FrmPrinter.VS.TextAlign = taLeftMiddle
    
    xFila = xFila + 450 + 100
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    FrmPrinter.VS.TextBox "Nombre Banco    :", 1000, xFila, 1700, 200, True, False, False
    FrmPrinter.VS.TextBox NulosC(xRstCab("nomban")), 2600, xFila, 4200, 200, True, False, False
    
    FrmPrinter.VS.TextBox "Fch. Emisión    :", 7000, xFila, 1700, 200, True, False, False
    FrmPrinter.VS.TextBox NulosC(xRstCab("fchope")), 8600, xFila, 2360, 200, True, False, False
    
    xFila = xFila + 200
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    FrmPrinter.VS.TextBox "Nº Cta. Cte.    :", 1000, xFila, 1700, 200, True, False, False
    FrmPrinter.VS.TextBox NulosC(xRstCab("numcue")), 2600, xFila, 4200, 200, True, False, False
    
    FrmPrinter.VS.TextBox "Moneda          :", 7000, xFila, 1700, 200, True, False, False
    FrmPrinter.VS.TextBox NulosC(xRstCab("simbolo")), 8600, xFila, 2360, 200, True, False, False
    
    xFila = xFila + 200
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    FrmPrinter.VS.TextBox "Tip. Documento  :", 1000, xFila, 1700, 200, True, False, False
    If xRstDet.State = 1 Then
        If xRstDet.RecordCount <> 0 Then
            FrmPrinter.VS.TextBox NulosC(xRstDet("descdoc")), 2600, xFila, 4200, 200, True, False, False
        End If
    End If
    
    FrmPrinter.VS.TextBox "T.C.             :", 7000, xFila, 1700, 200, True, False, False
    FrmPrinter.VS.TextBox Format(NulosC(xRstCab("impven")), "0.000"), 8600, xFila, 2360, 200, True, False, False
    
    xFila = xFila + 200
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    FrmPrinter.VS.TextBox "Nº Documento    :", 1000, xFila, 1700, 200, True, False, False
    If xRstDet.State = 1 Then
        If xRstDet.RecordCount <> 0 Then
            FrmPrinter.VS.TextBox NulosC(xRstDet("numdoc")), 2600, xFila, 4200, 200, True, False, False
        End If
    End If
    
    FrmPrinter.VS.TextBox "Importe         :", 7000, xFila, 1700, 200, True, False, False
    FrmPrinter.VS.TextBox Format(NulosN(xRstCab("importe")), "0.00"), 8600, xFila, 2360, 200, True, False, False
    
    xFila = xFila + 200
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    FrmPrinter.VS.TextBox "Medio de Pago   :", 1000, xFila, 1700, 200, True, False, False
    If xRstDet.State = 1 Then
        If xRstDet.RecordCount <> 0 Then
            FrmPrinter.VS.TextBox Mid(NulosC(xRstDet("descmedpag")), 1, 47), 2600, xFila, 4200, 200, True, False, False
        End If
    End If
    
    FrmPrinter.VS.TextBox "Situación       :", 7000, xFila, 1700, 200, True, False, False
    FrmPrinter.VS.TextBox "", 8600, xFila, 2360, 200, True, False, False
    
    xFila = xFila + 200
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    FrmPrinter.VS.TextBox "Glosa           :", 1000, xFila, 1700, 200, True, False, False
    FrmPrinter.VS.TextBox NulosC(xRstCab("glosa")), 2600, xFila, 8360, 200, True, False, False
    
    xFila = xFila + 250
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    FrmPrinter.VS.FontSize = 9
    FrmPrinter.VS.TextAlign = taLeftTop
    FrmPrinter.VS.TextBox "DESTINO DEL " & UCase(NulosC(xRstCab("xtipo"))) & "", 1000, xFila, 7500, 250, True, False, False
    
    xFila = xFila + 270
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    FrmPrinter.VS.FontSize = 7
    FrmPrinter.VS.TextAlign = taCenterMiddle
    FrmPrinter.VS.TextBox "Nº Doc. Iden.", 1000, xFila, 1000, 400, True, False, True
    FrmPrinter.VS.TextBox "Nombre / Razón Social", 2000, xFila, 3560, 400, True, False, True
    FrmPrinter.VS.TextBox "Fch. Doc.", 5560, xFila, 900, 400, True, False, True
    FrmPrinter.VS.TextBox "T.D.", 6460, xFila, 400, 400, True, False, True
    FrmPrinter.VS.TextBox "Nº Documento", 6860, xFila, 1200, 400, True, False, True
    FrmPrinter.VS.TextBox "M.", 8060, xFila, 400, 400, True, False, True
    FrmPrinter.VS.TextBox "T.C.", 8460, xFila, 500, 400, True, False, True
    FrmPrinter.VS.TextBox "Imp. Orig.", 8960, xFila, 1000, 400, True, False, True
    FrmPrinter.VS.TextBox "Imp. Abon.", 9960, xFila, 1000, 400, True, False, True
    
    Dim B As Integer
    Dim xTotal As Double
    
    FrmPrinter.VS.FontSize = 6
    xFila = xFila + 400
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    xTotal = 0
    
    If xRstDetDes.RecordCount <> 0 Then
        xRstDetDes.MoveFirst
        
        For B = 1 To xRstDetDes.RecordCount
            'FrmPrinter.VS.FontSize = 6
            FrmPrinter.VS.TextAlign = taCenterMiddle
            
            FrmPrinter.VS.TextBox NulosC(xRstDetDes("numruc")), 1000, xFila, 1000, 200, True, False, True
            FrmPrinter.VS.TextBox NulosC(Mid(xRstDetDes("nombre"), 1, 31)), 2000, xFila, 3560, 200, True, False, True
            FrmPrinter.VS.TextBox Format(xRstDetDes("fchdoc"), "dd/mm/yy"), 5560, xFila, 900, 200, True, False, True
            FrmPrinter.VS.TextBox NulosC(xRstDetDes("abrev")), 6460, xFila, 400, 200, True, False, True
            FrmPrinter.VS.TextBox xRstDetDes("numdoc"), 6860, xFila, 1200, 200, True, False, True
            FrmPrinter.VS.TextBox NulosC(xRstDetDes("simbolo")), 8060, xFila, 400, 200, True, False, True
            FrmPrinter.VS.TextBox NulosC(xRstDetDes("tipcam")), 8460, xFila, 500, 200, True, False, True
            FrmPrinter.VS.TextBox Format(NulosN(xRstDetDes("imptot")), "0.00"), 8960, xFila, 1000, 200, True, False, True
            FrmPrinter.VS.TextBox Format(NulosN(xRstDetDes("acuenta")), "0.00"), 9960, xFila, 1000, 200, True, False, True
            
            xTotal = xTotal + NulosN(xRstDetDes("acuenta"))
            xFila = xFila + 200
            xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
            xRstDetDes.MoveNext
            If xRstDetDes.EOF = True Then Exit For
        Next B
    End If
    
    FrmPrinter.VS.TextBox "TOTAL ==>", 8960, xFila, 1000, 200, True, False, True
    FrmPrinter.VS.TextBox Format(xTotal, "0.00"), 9960, xFila, 1000, 200, True, False, True
    
    xFila = xFila + 250
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    FrmPrinter.VS.FontSize = 9
    FrmPrinter.VS.TextAlign = taCenterMiddle
    FrmPrinter.VS.TextBox "ASIENTO CONTABLE", 1000, xFila, 9960, 250, True, False, True
    
    xFila = xFila + 250
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    FrmPrinter.VS.FontSize = 7
    
    FrmPrinter.VS.TextBox "Nº Cuenta", 1000, xFila, 1000, 400, True, False, True
    FrmPrinter.VS.TextBox "Nombre de la Cuenta", 2000, xFila, 2860, 400, True, False, True
    FrmPrinter.VS.TextBox "T.D.", 4860, xFila, 400, 400, True, False, True
    FrmPrinter.VS.TextBox "Nº Documento", 5260, xFila, 1200, 400, True, False, True
    FrmPrinter.VS.TextBox "M.", 6460, xFila, 400, 400, True, False, True
    FrmPrinter.VS.TextBox "IMPORTE EN M.E.", 6860, xFila, 1800, 200, True, False, True
    FrmPrinter.VS.TextBox "DEBE", 6860, xFila + 200, 900, 200, True, False, True
    FrmPrinter.VS.TextBox "HABER", 7760, xFila + 200, 900, 200, True, False, True
    FrmPrinter.VS.TextBox "T.C.", 8660, xFila, 500, 400, True, False, True
    FrmPrinter.VS.TextBox "IMPORTE EN M.N.", 9160, xFila, 1800, 200, True, False, True
    FrmPrinter.VS.TextBox "DEBE", 9160, xFila + 200, 900, 200, True, False, True
    FrmPrinter.VS.TextBox "HABER", 10060, xFila + 200, 900, 200, True, False, True
    
    Dim xTotalDebSol, xTotalDebDol, xTotalHabSol, xTotalHabDol As Double
    FrmPrinter.VS.FontSize = 6
    
    If RstDia.RecordCount <> 0 Then
        xFila = xFila + 400
        xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
        RstDia.MoveFirst
        For B = 1 To RstDia.RecordCount
''            FrmPrinter.VS.TextBox RstDia("cuenta"), 1000, xFila, 1000, 200, True, False, True
''            FrmPrinter.VS.TextBox Mid(RstDia("descripcion"), 1, 28), 2000, xFila, 2860, 200, True, False, True
''            FrmPrinter.VS.TextBox NulosC(RstDia("abrev")), 4860, xFila, 400, 200, True, False, True
''            FrmPrinter.VS.TextBox RstDia("rnumerodoc"), 5260, xFila, 1200, 200, True, False, True
''            FrmPrinter.VS.TextBox RstDia("simbolo"), 6460, xFila, 400, 200, True, False, True
''            FrmPrinter.VS.TextBox Format(RstDia("impdebdol"), "0.00"), 6860, xFila, 900, 200, True, False, True
''            FrmPrinter.VS.TextBox Format(RstDia("imphabdol"), "0.00"), 7760, xFila, 900, 200, True, False, True
''            FrmPrinter.VS.TextBox Format(RstDia("tc"), "0.000"), 8660, xFila, 500, 200, True, False, True
''            FrmPrinter.VS.TextBox Format(RstDia("impdebsol"), "0.00"), 9160, xFila, 900, 200, True, False, True
''            FrmPrinter.VS.TextBox Format(RstDia("imphabsol"), "0.00"), 10060, xFila, 900, 200, True, False, True
''
''            xTotalDebSol = xTotalDebSol + NulosN(RstDia("impdebsol"))
''            xTotalHabSol = xTotalHabSol + NulosN(RstDia("imphabsol"))
''
''            xTotalDebDol = xTotalDebDol + NulosN(RstDia("impdebdol"))
''            xTotalHabDol = xTotalHabDol + NulosN(RstDia("imphabdol"))
            
            
            FrmPrinter.VS.TextBox RstDia("ctanum"), 1000, xFila, 1000, 200, True, False, True
            FrmPrinter.VS.TextBox Mid(RstDia("ctadesc"), 1, 28), 2000, xFila, 2860, 200, True, False, True
            FrmPrinter.VS.TextBox NulosC(RstDia("tdocdesc")), 4860, xFila, 400, 200, True, False, True
            FrmPrinter.VS.TextBox RstDia("numdoc"), 5260, xFila, 1200, 200, True, False, True
            FrmPrinter.VS.TextBox RstDia("monope"), 6460, xFila, 400, 200, True, False, True
            FrmPrinter.VS.TextBox Format(RstDia("impdebedol1"), "0.00"), 6860, xFila, 900, 200, True, False, True
            FrmPrinter.VS.TextBox Format(RstDia("imphaberdol1"), "0.00"), 7760, xFila, 900, 200, True, False, True
            FrmPrinter.VS.TextBox Format(RstDia("tipcam"), "0.000"), 8660, xFila, 500, 200, True, False, True
            FrmPrinter.VS.TextBox Format(RstDia("impdebesol1"), "0.00"), 9160, xFila, 900, 200, True, False, True
            FrmPrinter.VS.TextBox Format(RstDia("imphabersol1"), "0.00"), 10060, xFila, 900, 200, True, False, True
        
            xTotalDebSol = xTotalDebSol + NulosN(RstDia("impdebesol1"))
            xTotalHabSol = xTotalHabSol + NulosN(RstDia("imphabersol1"))
            
            xTotalDebDol = xTotalDebDol + NulosN(RstDia("impdebedol1"))
            xTotalHabDol = xTotalHabDol + NulosN(RstDia("imphaberdol1"))
            
            xFila = xFila + 200
            xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
            RstDia.MoveNext
            If RstDia.EOF = True Then Exit For
        Next B
    End If
    
    FrmPrinter.VS.TextBox "TOTAL ==> ", 5260, xFila, 1200, 200, True, False, True
    
    FrmPrinter.VS.TextBox Format(xTotalDebDol, "0.00"), 6860, xFila, 900, 200, True, False, True
    FrmPrinter.VS.TextBox Format(xTotalHabDol, "0.00"), 7760, xFila, 900, 200, True, False, True
    FrmPrinter.VS.TextBox Format(xTotalDebSol, "0.00"), 9160, xFila, 900, 200, True, False, True
    FrmPrinter.VS.TextBox Format(xTotalHabSol, "0.00"), 10060, xFila, 900, 200, True, False, True
    
    xFila = xFila + 200
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    xFila = xFila + 200
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    xFila = xFila + 200
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    xFila = xFila + 200
    xFila = LeerFila(xFila, FrmPrinter.VS.FontSize)
    
    CabeceraVoucher = xFila
End Function


