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

Public xAño As Integer                 ' ESPECIFICA EL AÑO DE TRABAJO
Public xCon As New ADODB.Connection    ' ESPECIFICA LA CONECCION ACTIVA A LA BADE DE DATOS
Public xTitulo As String               ' ESPECIFICA EL TITULO PARA LOS CUADRO DE TEXTO MSGBOX
Public xMes As Integer                 ' ESPECIFICA EL MES DE TRABAJO ACTUAL
Public NomEmp As String                ' ESPECIFICA EL NOMBRE DE LA EMPRESA DE TRABAJO ACTUAL
Public NumRUC As String                ' ESPECIFICA EL NUMERO DE RUC DE LA EMPRESA DE TRABAJO ACTUAL
Public AnoTra As String                ' ESPECIFICA EL AÑO DE TRABAJO ACTUAL
Public CONTABILIZAR As Boolean         ' ESPECIFICA SI SE REALIZARAN LOS PROCESOS CONTABLES
Public NomSis As String                ' ESPECIFICA EL NOMBRE DEL SISTEMA
Public xIdUsuario As Integer           ' ESPECIFICA EL ID DEL USUARIO ACTUAL
Public xDeDonde As Integer             ' ESPECIFICA DE DONDE ES INVOCADO EL FORMULARIO
Public AP_RUTASY As String             ' ALMACENA LA RUTA DEL SISTEMA
Public AP_RUTABD As String             ' ALMACENA LA RUTA DE LA BASE DE DATOS
Public AP_RUTABM As String             ' ALMACENA LA RUTA DE LOS ARCHIVOS BMPS
Public AP_AÑODAT As String             ' ALMACENA EL AÑO DE TRABAJO ACTUAL
Public AP_MESTRA As Integer            ' ALMACENA EL MES DE TRABAJO ACTUAL
Public xConTMP As New ADODB.Connection ' ESPECIFICA UNA CONECCION A UNA BASE DE DATOS TEMPORAL

Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)

'*****************************************************************************************************
'* Nombre           : RellenaNumdoc
'* Tipo             : FUNCION
'* Descripcion      : RELLENA CON 0 EL NUMERO DE UN DOCUMENTO, ESTA FUNCION DEVUELVE UNA CADENA
'* Paranetros       : NOMBRE    |  TIPO    |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    xnumser   |  Long    |  ESPECIFICA EL NUMERO DE SERIE DEL DOCUMENTO
'*                    xnumdoc   |  Long    |  ESPECIFICA EL NUMERO DE DOCUMENTO
'* Devuelve         : String
'*****************************************************************************************************
Function RellenaNumdoc(xnumser As Long, xnumdoc As Long) As String
    RellenaNumdoc = Format(xnumser, "0000") & "-" & Format(xnumdoc, "0000000000")
End Function

'*****************************************************************************************************
'* Nombre           : ActualizaNroDocumento
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE   |  TIPO     |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    xnumdoc  |  Long     |  ESPECIFICA EL NUMERO DE DOCUMENTO
'*                    xtipdoc  |  Long     |  ESPECIFICA EL TIPO DE DOCUMENTO
'*                    xSerie   |  Long     |  ESPECIFICA EL NUMERO DE SERIE DEL DOCUMENTO
'* Devuelve         :
'*****************************************************************************************************
Sub ActualizaNroDocumento(xnumdoc As Long, xtipdoc As Long, xSerie As Long)
    Dim Rst As New ADODB.Recordset
    Dim NumDoc As String

    RST_Busq Rst, "SELECT * from mae_Series where iddoc = " & xtipdoc & " and numser = " & xSerie & "" _
        & " ORDER BY numser,numdoc ", xCon

    If Rst.RecordCount = 0 Then
        Rst.AddNew
        Rst("numdoc") = xnumdoc
        Rst.Update
    Else
        Rst("numdoc") = xnumdoc
        Rst.Update
    End If
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CargaDatosEmpresa
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA EMPRESA
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
    xAño = AnoTra
    

    NomSis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_AÑODAT = Rst("anotra")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
    Set Rst = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : AbriConeccion
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ESTABLECE LA CONECCION CON LA BASE DE DATOS calendario.mdb
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub AbriConeccion()
    Dim xFun As New eps_librerias.FuncionesData
    Dim RutaSys As String
    RutaSys = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    
    xFun.F_BASEDATOS = RutaSys & "calendario.mdb"
    xFun.F_GRUPOTRABAJO = RutaSys & "\seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    Set xConTMP = xFun.AbrirConeccion
End Sub

'*****************************************************************************************************
'* Nombre           : LlenarDatos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LLENA DATOS EN LA TABLA event DE LA BASE DE DATOS calendario.mdb
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub LlenarDatos()
    Dim Rst As New ADODB.Recordset
    Dim RstEve As New ADODB.Recordset
    Dim RstPro As New ADODB.Recordset
    Dim A As Integer
    Dim B As Integer
    Dim xBody As String
    Dim xDia As Date
    
    Set xConTMP = Nothing
    AbriConeccion
    
    xConTMP.Execute "DELETE * FROM event"
    RST_Busq RstEve, "SELECT * FROM event", xConTMP
    
    RST_Busq Rst, "SELECT DISTINCT ped_pedido.id, mae_cliente.nombre, mae_cliente.dir, ped_pedido.idtipped, ped_pedidodetent.fchent, " _
        & " ped_pedidodetent.estado FROM mae_cliente RIGHT JOIN (alm_inventario RIGHT JOIN (ped_pedido LEFT JOIN ped_pedidodetent " _
        & " ON ped_pedido.id = ped_pedidodetent.idped) ON alm_inventario.id = ped_pedidodetent.iditem) ON mae_cliente.id = ped_pedido.idcli " _
        & " WHERE (((ped_pedido.idtipped)=2) AND ((ped_pedidodetent.estado)=2))" _
        & " Union " _
        & " SELECT DISTINCT ped_pedido.id, mae_cliente.nombre, mae_cliente.dir, ped_pedido.idtipped, ped_pedidodetent.fchent, ped_pedidodetent.estado " _
        & " FROM (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) LEFT JOIN (alm_inventario RIGHT JOIN ped_pedidodetent " _
        & " ON alm_inventario.id = ped_pedidodetent.iditem) ON ped_pedido.id = ped_pedidodetent.idped WHERE (((ped_pedido.idtipped)=1) " _
        & " AND ((ped_pedidodetent.estado)=2))", xCon
    
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Dim xHorIni, HorFin As String
        
        xDia = Rst("fchent")
        xHorIni = "00:00:00"
            
        For A = 1 To Rst.RecordCount
            Set RstPro = Nothing
            
            RstEve.AddNew
            RstEve("EventID") = A
            
            RstEve("StartDateTime") = Format(Rst("fchent"), "dd/mm/yyyy") & " " & Format(xHorIni, "hh:mm:ss")        ' "8:00:00"
            
            HorFin = ConvertHora(ConvertSeg(Format(xHorIni, "hh:mm:ss")) + ConvertSeg("02:00:00"))
            
            RstEve("EndDateTime") = Format(Rst("fchent"), "dd/mm/yyyy") & " " & Format(HorFin, "hh:mm:ss")
            RstEve("RecurrenceState") = 0
            RstEve("IsAllDayEvent") = 0
            RstEve("Subject") = Trim(Rst("nombre"))
            RstEve("Location") = Trim(Rst("dir"))
            RstEve("RemainderSoundFile") = ""
            RstEve("Created") = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
            RstEve("Modified") = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
            RstEve("BusyStatus") = 2
            RstEve("ImportanceLevel") = 1
            RstEve("LabelID") = 2
            RstEve("RecurrencePatternID") = 0
            RstEve("ScheduleID") = 0
            RstEve("ISRecurrenceExceptionDeleted") = 0
            RstEve("RExceptionStartTimeOrig") = "00:00:00"
            RstEve("RExceptionEndTimeOrig") = "00:00:00"
            RstEve("IsMeeting") = 0
            RstEve("IsPrivate") = 0
            RstEve("IsReminder") = 0
            RstEve("ReminderMinutesBeforeStart") = 15
            RstEve("CustomPropertiesXMLData") = "<Calendar CompactMode='1'/>"
            
            If Rst("idtipped") = 1 Then
                RST_Busq RstPro, "SELECT DISTINCT ped_pedido.id, mae_cliente.nombre, mae_cliente.dir, ped_pedido.idtipped, ped_pedidodetent.fchent, " _
                    & " ped_pedidodetent.estado, alm_inventario.descripcion, ped_pedidodetent.canpro, mae_unidades.abrev " _
                    & " FROM mae_unidades RIGHT JOIN (mae_cliente RIGHT JOIN (alm_inventario RIGHT JOIN (ped_pedido LEFT JOIN ped_pedidodetent " _
                    & " ON ped_pedido.id = ped_pedidodetent.idped) ON alm_inventario.id = ped_pedidodetent.iditem) ON mae_cliente.id = ped_pedido.idcli) " _
                    & " ON mae_unidades.id = ped_pedidodetent.idunimed WHERE (((ped_pedido.id)=" & Rst("id") & ") AND ((ped_pedido.idtipped)=1) " _
                    & " AND ((ped_pedidodetent.fchent)=CDate('" & Rst("fchent") & "')) AND ((ped_pedidodetent.estado)=2))", xCon
            Else
                RST_Busq RstPro, "SELECT DISTINCT ped_pedido.id, mae_cliente.nombre, mae_cliente.dir, ped_pedido.idtipped, ped_pedidodetent.fchent, " _
                    & " ped_pedidodetent.estado, alm_inventario.descripcion, ped_pedidodetent.canpro, mae_unidades.abrev " _
                    & " FROM mae_unidades RIGHT JOIN (mae_cliente RIGHT JOIN (alm_inventario RIGHT JOIN (ped_pedido LEFT JOIN ped_pedidodetent " _
                    & " ON ped_pedido.id = ped_pedidodetent.idped) ON alm_inventario.id = ped_pedidodetent.iditem) ON mae_cliente.id = ped_pedido.idcli) " _
                    & " ON mae_unidades.id = ped_pedidodetent.idunimed WHERE (((ped_pedido.id)=" & Rst("id") & ") AND ((ped_pedido.idtipped)=2) " _
                    & " AND ((ped_pedidodetent.fchent)=CDate('" & Rst("fchent") & "')) AND ((ped_pedidodetent.estado)=2))", xCon
            End If
            
            B = 0
            xBody = ""
            If RstPro.RecordCount <> 0 Then
                RstPro.MoveFirst
                For B = 1 To RstPro.RecordCount
                    xBody = xBody + Trim(RstPro("descripcion")) & " | " & Trim(RstPro("abrev")) & " | " & Format(RstPro("canpro"), "0.00")
                    RstPro.MoveNext
                    
                    If RstPro.EOF = True Then Exit For
                    xBody = xBody & Chr(13)
                Next B
            End If
            ' escribimos el contenido del cuerpo
            RstEve("Body") = Mid(xBody, 1, 250)
            
            RstEve.Update
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
            
            If xDia = Rst("fchent") Then
                xHorIni = HorFin
            Else
                xHorIni = "00:00:00"
            End If
        Next A
    End If
End Sub
