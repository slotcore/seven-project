Attribute VB_Name = "Declaraciones"
'*****************************************************************************************************
'* Nombre Archivo   : DECLARACIONES.BAS
'* Tipo             : MODULO
'* Descripcion      : MODULO EN EL QUE SE DEFINEN LA PRINCIPALES VARIABLES A UTILIZARCE EN LA CLASE,
'*                    ADEMAS AQUI SE DEFINEN FUNCIONES QUE SERAN USADAS UNICAMENTE EN LA CLASE ACTUAL
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public xCon As New ADODB.Connection   ' CONECCION A LA BASE DE DATOS
Public xTitulo As String              ' TITULO PARA LA CLASE CUANDO SE MUESTRE UN MENSAJE
Public NomEmp As String               ' NOMBRE DE LA EMPRESA
Public NomSIS As String               ' NOMBRE DEL SISTEMA
Public NumRUC As String               ' NUMERO DE RUC DE LA EMPRESA
Public AnoTra As String               ' AÑO DE TRABAJO ACTUAL
Public CONTABILIZAR As Boolean        ' VARIABLE QUE INDICA SI SE HARAN PROCESOS CONTABLES  TRUE = SE HARA PROCESO CONTABLE ; FALSE = NO SE HARA PROCES CONTABLE
Public xMes As Integer                ' INDICA EL MES DE TRABAJO ACTUAL
Public xOrigen As Integer             ' ESPECIFICA DE DONDE ES INVOCADO EL FORMAULARIO ; 0 = MENU PRINCIPAL; 1 = ALGUNA LIBRERIA O FUNCION
Public xIdUsuario As Integer          ' ALAMACENA EL ID DEL USUARIO

Public AP_RUTASY As String            ' RUTA DEL SISTEMA
Public AP_RUTABD As String            ' RUTA DE LA BASE DE DATOS
Public AP_RUTABM As String            ' RUTA DE LOS ARCHIVOS GRAFICOS DEL SISTEMA
Public AP_AÑODAT As String            ' AÑO DE TRABAJO
Public AP_MESTRA As Integer           ' MES DE TRABAJO

Public xDeDonde As Integer            ' ESPECIFICA SI SE SINCRONIZA EL INGRESO DE DATOS CON LAS DEMAS BASES DE DATOS
Public IdCompraReg As Double         ' alamcenara el id de la compra registrada esta funcion es para cuando se halla llamado el formulario desde otro formulario

Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)


'*****************************************************************************************************
'* Nombre           : HallaNumCuenta
'* Tipo             : FUNCION
'* Descripcion      : HALLA LA CUENTA CONTABLE DEL DOCUMENTO ESPECIFICADO, ESTA FUNCION DEVUELDE EL ID
'*                    DE LA CUENTA CONTABLE (CAMPO Id TABLA con_planctas), SI NO ENCUENTRA DEVUELVE 0
'* Paranetros       : NOMBRE         |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    TipoDocumento  |  STRING           |  ESPECIFICA EL ID DEL TIPO DE DOCUMENTO
'*                    TipoMoneda     |  STRING           |  ESPECIFICA EL ID DE LA MONEDA
'* Devuelve         : INTEGER
'* Observaciones    : LOS PARAMETROS DE LA FUNCION DEBIERON DE SER DE TIPO INTEGER, TENER EN CUENTA
'*                    PARA FUTURAS REVISIONES
'*****************************************************************************************************
Function HallaNumCuenta(TipoDocumento As String, TipoMoneda As String) As Integer
    'Funcion que permite hallar el numero de cuenta contable del tipo de documento seleccionado
    If NulosC(TipoDocumento) = "" Or NulosC(TipoMoneda) = "" Then
        Exit Function
    End If
    
    Dim Rst As New ADODB.Recordset
    
    ' BUSCA EL TIPO DE DOCUMENTO Y SU MONEDA PARA HALLAR LA CUENTA CONTABLE DEL DOCUMENTO SEÑALADO
    RST_Busq Rst, "SELECT mae_documentocta.iddoc, mae_documentocta.idmon, mae_documentocta.tipope, mae_documentocta.idcuen, " _
        & " con_planctas.cuenta FROM mae_documentocta LEFT JOIN con_planctas ON mae_documentocta.idcuen = con_planctas.id " _
        & " WHERE (((mae_documentocta.iddoc)=" & NulosN(TipoDocumento) & ") AND ((mae_documentocta.idmon)=" & NulosN(TipoMoneda) & ") " _
        & " AND ((mae_documentocta.tipope)=0))", xCon

    If Rst.RecordCount = 0 Then
        MsgBox "El documento especificado no tiene asignado una cuenta contable en facturas por pagar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        HallaNumCuenta = 0
    Else
        HallaNumCuenta = NulosN(Rst("idcuen"))
    End If
End Function

'*****************************************************************************************************
'* Nombre           : CargaDatosEmpresa
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA EMPRESA: Nombre de la empresa, Numero de Ruc, Año de trabajo
'*                    TAMBIEN CARGA LOS SIGUIENTES DATOS DEL SISTEMA: Nombre del sistema, Ruta de la
'*                    base de datos, Ruta del sistema, Ruta de los archivo de grafico, ADEMAS ESPECIFICA
'*                    SI EL SISTEMA HARA PROCESOS CONTABLES
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    CONTABILIZAR = Rst("procon")
    AnoTra = Rst("anotra")
    Set Rst = Nothing
    NomSIS = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

'*****************************************************************************************************
'* Nombre           : TasaImpuestoDocumento
'* Tipo             : FUNCION
'* Descripcion      : BUSCA LA TASA DE IMPUESTO ASIGNADA AL DOCUMENTO, DEVUELVE UN VALOR ENTERO
'* Paranetros       : NOMBRE          |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    IdTipoDocumento |  INTEGER          |  ESPECIFICA EL ID DEL TIPO DE DOCUMENTO
'* Devuelve         : INTEGER
'* Observaciones    : EL VALOR DEVUELTO POR ESTA FUNCION DEVERIA DE SER UN VALOR DOUBLE, TENER EN
'*                    CUENTA PARA FUTURAS REVISIONES
'*****************************************************************************************************
Function TasaImpuestoDocumento(IdTipoDocumento As Integer) As Integer
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, mae_impuestos.Abrev AS abreimp, " _
        & " mae_impuestos.idcuen AS cuentaimp FROM mae_documento LEFT JOIN mae_impuestos ON " _
        & " mae_documento.idimp = mae_impuestos.id WHERE (((mae_documento.id)=" & NulosN(IdTipoDocumento) & "))", xCon

    If Rst.RecordCount <> 0 Then
        TasaImpuestoDocumento = NulosN(Rst("tasa"))
    Else
        TasaImpuestoDocumento = 0
    End If
    Set Rst = Nothing
End Function

'*****************************************************************************************************
'* Nombre           : HallaDatosImpuestoDocumento
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BUSCA EL VALOR DE UN DETERMINADO CAMPO EN LA TABLA mae_impuestos Y DEVUELVE SU
'*                    VALOR, DEVUELVE UN VALOR ENTERO, 0 SI NO TIENE EXITO EN LA BUSQUEDA
'* Paranetros       : NOMBRE         |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    TipoDocumento  | INTEGER           |  ESPECIFICA EL ID DEL TIPO DE DOCUMENTO
'*                    CampoDevolver  | STRING            |  ESPECIFICA EL VALOR DEL CAMPO QUE DEVOLVERA
'* Devuelve         : INTEGER
'* Observaciones    : ESTA FUNCION VERIA DE DEVOLVER UN VALOR VARIAN, YA QUE PUEDE DEVOLVER CUALQUIER
'*                    CAMPO NO SIENDO ESTE NECESARIAMENTE DEL TIPO ENTERO, TENER EN CUENTA ESTE CAMBIO
'*                    PARA FUTURAS REVISIONES
'*****************************************************************************************************
Function HallaDatosImpuestoDocumento(TipoDocumento As Integer, CampoDevolver As String) As Integer
    'TipoDocumento = tipo de documento del que se determinara el dato
    'CampoDevolver
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, mae_impuestos.Abrev AS abreimp, " _
        & " mae_impuestos.idcuen AS cuentaimp FROM mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id " _
        & " WHERE (((mae_documento.id)=" & TipoDocumento & "))", xCon

    If Rst.RecordCount <> 0 Then
        HallaDatosImpuestoDocumento = Rst(CampoDevolver)
    Else
        HallaDatosImpuestoDocumento = 0
    End If
    Set Rst = Nothing
End Function


Sub EscribirAsiento(xAñoTra As Integer, xMesAct As Integer, xidMon As Integer, IdLib As Integer, xIdMov As Double, _
                    xNumAsiento As String, xTc As Double, xIdCuenta, xFchDoc As String, xFchAsiento As String, _
                    xIdTipDoc As Integer, xImporte As Double, xTipo As Integer, xCon As ADODB.Connection, _
                    Optional xEsDestino As Boolean = False)
                    
    ' xTipo = 1  DEBE
    ' xTipo = 2  HABER
    
    '--20/10/10 Se agrega parametro xEsDestino; Para controlar la recursiva solo cuando es destino
    
    'Dim RstDia As New ADODB.Recordset
    Dim xImpSol, xImpDol As Double
    Dim xImpHabSol, xImpDebSol, xImpDebDol, xImpHabDol As Double
    
    ' DETERMINAMOS LA MONEDA ORIGEN
    If xidMon = 1 Then
        xImpSol = xImporte
        xImpDol = xImporte / xTc
    End If
    
    If xidMon = 2 Then
        xImpSol = xImporte * xTc
        xImpDol = xImporte
    End If
    
    '--verificar si es nota de credito
    
    If xIdTipDoc = 7 Then
        If xTipo = 1 Then
            xTipo = 2
        Else
            xTipo = 1
        End If
    End If
    
    
    If xTipo = 1 Then
        xImpDebSol = xImpSol:         xImpHabSol = 0
        xImpDebDol = xImpDol:         xImpHabDol = 0
    Else
        xImpDebSol = 0:               xImpHabSol = xImpSol
        xImpDebDol = 0:               xImpHabDol = xImpDol
    End If

On Error GoTo LaCague

    xCon.Execute "INSERT INTO con_diario (año, idmes, idlib, idmov, numasi, tc, idcue, fchasi, fchdoc, impdebdol, imphabdol, impdebsol, imphabsol)" _
        & " SELECT " & xAñoTra & " AS Expr1, " & xMesAct & " AS Expr2, " & IdLib & " As Expr3, " & xIdMov & " As Expr4,'" & xNumAsiento & "' AS Exp5," _
        & " " & xTc & " As Expr6, " & xIdCuenta & " As Expr7,cdate('" & xFchAsiento & "') As Expr8,cdate('" & xFchDoc & "') As Expr9, " & xImpDebDol & " As Expr10 ," _
        & " " & xImpHabDol & " AS Expr11," & xImpDebSol & " AS Expr12," & xImpHabSol & " AS Expr13"

    
    '--Verificar si la cuenta tiene destino
    '--xEsDestino=false :: Verificar si la cuenta tiene destinos para escribir los asientos
    '--xEsDestino=true  :: Indica que se esta escribiendo el registro que hace referencia a una cta destino
    If xEsDestino = False Then
        
        ' buscamos si la cuenta tiene destino
        Dim xRs As New ADODB.Recordset
        
        RST_Busq xRs, "SELECT * FROM con_planctas WHERE id = " & xIdCuenta & "", xCon
        If xRs.RecordCount <> 0 Then
            If xRs("ctadesdeb") <> 0 And xRs("ctadeshab") <> 0 Then
                ' GRABAMOS EL DEBE DEL DESTINO
    '            xCon.Execute "INSERT INTO con_diario (año, idmes, idlib, idmov, numasi, tc, idcue, fchasi, fchdoc, impdebdol, imphabdol, impdebsol, imphabsol)" _
                    & " SELECT " & xAñoTra & " AS Expr1, " & xMesAct & " AS Expr2, " & IdLib & " As Expr3, " & xIdMov & " As Expr4,'" & xNumAsiento & "' AS Exp5," _
                    & " " & xTc & " As Expr6, " & xRs("ctadesdeb") & " As Expr7,cdate('" & xFchAsiento & "') As Expr8,cdate('" & xFchDoc & "') As Expr9, " _
                    & " " & xImpDebDol & " As Expr10 , " & xImpHabDol & " AS Expr11," _
                    & " " & xImpDebSol & " AS Expr12," & xImpHabSol & " AS Expr13"
                '--escribir el destino del debe
                EscribirAsiento xAñoTra, xMesAct, xidMon, IdLib, xIdMov, xNumAsiento, xTc, xRs("ctadesdeb"), xFchDoc, xFchAsiento, xIdTipDoc, xImporte, 1, xCon, True
            
                ' GRABAMOS EL HABER DEL DESTINO
                ' AQUI INVERTIMOS LOS IMPORTE PARA LOGRAR EL ASIENTO HABER
    '            xCon.Execute "INSERT INTO con_diario (año, idmes, idlib, idmov, numasi, tc, idcue, fchasi, fchdoc, impdebdol, imphabdol, impdebsol, imphabsol)" _
                    & " SELECT " & xAñoTra & " AS Expr1, " & xMesAct & " AS Expr2, " & IdLib & " As Expr3, " & xIdMov & " As Expr4,'" & xNumAsiento & "' AS Exp5," _
                    & " " & xTc & " As Expr6, " & xRs("ctadeshab") & " As Expr7,cdate('" & xFchAsiento & "') As Expr8,cdate('" & xFchDoc & "') As Expr9, " _
                    & " " & xImpHabDol & " As Expr10, " & xImpDebDol & " AS Expr11," _
                    & " " & xImpHabSol & " AS Expr12, " & xImpDebSol & " AS Expr13"
                '--escribir el destino del haber
                EscribirAsiento xAñoTra, xMesAct, xidMon, IdLib, xIdMov, xNumAsiento, xTc, xRs("ctadeshab"), xFchDoc, xFchAsiento, xIdTipDoc, xImporte, 2, xCon, True
            
            'Else
            '   MsgBox "La cuenta contable no tiene los destino correctamente distribuidos, verifique la informacion en el maestro del plan de cuentas", vbInformation + vbOKOnly + vbDefaultButton1, ""
            End If
        End If
        Set xRs = Nothing
        
    End If

    Exit Sub
    
LaCague:
    Resume
End Sub
