Attribute VB_Name = "declaraciones"
'*****************************************************************************************************
'* Nombre Archivo   : DECLARACIONES.BAS
'* Tipo             : MODULO
'* Descripcion      : MODULO DONDE SE DECLARAN LAS PRICIPALES VARIABLES UTILIZADAS EN LA CLASE ASI
'*                    COMO FUNCIONES QUE SOLO SERAN UTILIZADAS EN ESTA CLASE
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 09/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public xCon As New ADODB.Connection      ' VARIABLE QUE ALAMCENA LA CONECCION PRINCIPAL A LA BASE DE DATOS
Public NomEmp As String                  ' ALMACENA EL NOMBRE DE LA EMPRESA
Public NomSIS As String                  ' ALMACENA EL NOMBRE DEL SISTEMA
Public NumRUC As String                  ' ALMACENA EL NUMERO DE RUC DE LA EMPRESA
Public AP_RUTASY As String               ' ALMACENA LA RUTA DEL SISTEMA
Public AP_RUTABD As String               ' ALMACENA LA RUTA DE LA BASE DE DATOS
Public AP_RUTABM As String               ' ALMACENA LA RUTA DE LOS ARCHIVOS DE GRAFICO
Public AP_AÑODAT As String               ' ALMACENA EL AÑO DE TRABAJO DE LA DATA
Public AP_MESTRA As Integer              ' ALMACENA EL MES DE TRABAJO ACTUAL
Public xTitulo As String                 ' ALMACENA EL TITULO DE LA CLASE
Public AnoTra As String                  ' ALMACENA EL AÑO DE TRABAJO ACTUAL DEL SISTEMA

Public xIdUsuario As Integer             ' Especifica el id del usuario que accede al sistema
Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)

'*****************************************************************************************************
'* Nombre Archivo   : CargaDatosEmpresa
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA EMPRESA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = Rst("nomemp")
    NumRUC = Rst("numruc")
    AnoTra = Rst("anotra")
    
    Set Rst = Nothing
    NomSIS = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : AbrirConecciones
'* Tipo             : FUNCION
'* Descripcion      : ESTABLECE LA CONECCION A UNA BASE DE DATOS ESTA FUNCION DEVUELVE UNA VARIABLE DE
'*                    CONECCION
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Ruta         |  String    |  RUTA DE LA BASE DE DATOS QUE SE DESEA ACCEDER
'* DEVUELVE         : ADODB.Connection
'*****************************************************************************************************
Function AbrirConecciones(Ruta As String) As ADODB.Connection
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCone As ADODB.Connection
    
    xFun.F_BASEDATOS = Ruta
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xCone = xFun.AbrirConeccion
    Set xFun = Nothing
    Set AbrirConecciones = xCone
End Function


Sub MostrarAcumulado(xFgTot As VSFlexGrid, xFgCopy As VSFlexGrid, xTipo As String, xOrigen As Integer, _
                    xDatosAdicionales As Boolean, xFechaIni As Date, xObjBarra As ProgressBar)
    '===================================================================================================
    'Creado: 19/05/11 Por Johan Castro
    'Propósito: Mostrar el resumen de los productos terminados e intermedios en una ventana
    '
    'Entradas:  xFgTot = Objeto Flex Grid del total
    '           xFgCopy = Objeto Flex Grid que se va copiar al total
    '           xTipo = Indica el tipo de informacion
    '                   T = Terminados, PI = Intermedios (Para productos,materia prima, insumos)
    '           xOrigen = Indica si la consulta es lo siguiente:
    '                      1.- Productos(Terminado o intermendio)
    '                      2.- Materia Prima o Insumos
    '           xDatosAdicionales = Indica si se procede a mostrar los datos como StockInicia, Total Producido, etc
    '                               False = No se muestra los datos
    '                               True = Se procede a recorrer todos los registros para mostrar los datos
    '           xFechaIni = Indica la fecha inicial
    '           xObjBarra = Objeto de progressbar para mostrar el incremento en barra
    '
    'Resultados: Ventana con el resumen de los productos, Adicionalmente se muestra 4 columnas adicionales
    '             1.- Stock Ini = Indica el stock incial antes de la fecha de inicio
    '             2.- Producido/Comprado = Indica el total producido o comprado desde la fecha de inicio hasta la fecha actual
    '             3.- Total = Stock Ini + Producido/Comprado
    '             4.- Diferencia = Total - LoProgramado; Si > 0 = Pintar de color Azul(Item con Stock)
    '                                                    Si < 0 = Pintar de color Rojo(Item sin Stock)
    'Modificado :
    '
    '===================================================================================================




    '--declarar variables
    Dim xFilaCopy As Long '--indica la fila de la grilla que se va copiar al total
    Dim xColCopy As Long '--indica la columna de la grilla que se va copiar al total
    
    Dim xFilaTotal As Long '--india la fila de la grilla del total
    Dim xColTotal As Long '--indica la columna de la grilla del total
    
    
    Dim xColIni As Long '--indica la columna de incio de mes
    
    Dim xRegEncontrado As Boolean '--indica si el registro se encuentra en la grilla de los totales
                                  '--False=No se encuentra; se procede a agregar nueva fila con todos los datos
                                  '--True=Encontra en la grilla, solo se procede a acumular los valores
    
    
    '--Definir el encabezado
    If xFgTot.Cols = 6 Then
        '--Definir el tipo de información
        xFgTot.Cols = xFgTot.Cols + 1
        xFgTot.TextMatrix(0, xFgTot.Cols - 1) = "Tipo"
        xFgTot.ColWidth(xFgTot.Cols - 1) = 500
        '--Definir los periodos
        For xColCopy = 6 To xFgCopy.Cols - 1
            xFgTot.Cols = xFgTot.Cols + 1
            xFgTot.TextMatrix(0, xFgTot.Cols - 1) = xFgCopy.TextMatrix(0, xColCopy)
            
        Next xColCopy
        
        
    End If
    
    
    
    '--recorrer todas las filas del objeto a copiar los datos
    xFilaCopy = xFgCopy.FixedRows
    
    For xFilaCopy = xFgCopy.FixedRows To xFgCopy.Rows - 1

        '--reiniciando variable
        xRegEncontrado = False
        
        '--buscar si el registro esta en la grilla del total
        For xFilaTotal = xFgTot.FixedRows To xFgTot.Rows - 1
            '--se procedera a buscar por codigo de producto
            If xFgTot.TextMatrix(xFilaTotal, 1) = xFgCopy.TextMatrix(xFilaCopy, 1) Then
                '--salir si se encuentra, cambiar valor a variable para identificar que se ha encontrado el registro
                xRegEncontrado = True
                Exit For
            End If
        Next xFilaTotal
        
        If xRegEncontrado = True Then
            '--si ya existe el registro agregar en la misma fila los datos
            '--pintar la celda de producto para identificar aquellos productos que estan en terminado e intermedios
            GRID_COLOR_FONDO xFgTot, xFilaTotal, 4, xFilaTotal, 6, vbYellow
            '--cambiar el tipo a Ambos(Terminado e Intermedio)
            xFgTot.TextMatrix(xFilaTotal, 6) = "A"
        Else
            '--si no existe agregar nueva fila en grilla del total
            xFgTot.Rows = xFgTot.Rows + 1
            xFilaTotal = xFgTot.Rows - 1
            
            xFgTot.TextMatrix(xFilaTotal, 1) = xFgCopy.TextMatrix(xFilaCopy, 1)
            xFgTot.TextMatrix(xFilaTotal, 2) = xFgCopy.TextMatrix(xFilaCopy, 2)
            xFgTot.TextMatrix(xFilaTotal, 3) = xFgCopy.TextMatrix(xFilaCopy, 3)
            xFgTot.TextMatrix(xFilaTotal, 4) = xFgCopy.TextMatrix(xFilaCopy, 4)
            xFgTot.TextMatrix(xFilaTotal, 5) = xFgCopy.TextMatrix(xFilaCopy, 5)
            xFgTot.TextMatrix(xFilaTotal, 6) = xTipo
            
        End If
        
        
        '--acumular todos los datos segun los meses
        For xColIni = 6 To xFgCopy.Cols - 1
            '--acumular el valor anterior si hibiera mas el valor actual
            xFgTot.TextMatrix(xFilaTotal, xColIni + 1) = Format(NulosN(xFgTot.TextMatrix(xFilaTotal, xColIni + 1)) + NulosN(xFgCopy.TextMatrix(xFilaCopy, xColIni)), FORMAT_MONTO)
        Next xColIni
        
    Next xFilaCopy
    
    
    
    If xDatosAdicionales = True Then
    '--recorrer todas las filas del total para mostrar los datos adicionales como stock inicial, total producido, etc
        
        Dim xStkIni, xTotPro, xTotal As Double
        Dim AnoTra As Integer
        Dim RstTodProd As New Recordset
        
        '--agregando las ultimas columnas
        xFgTot.Cols = xFgTot.Cols + 4
        xFgTot.TextMatrix(0, xFgTot.Cols - 4) = "Stock Ini"
        
        If xOrigen = 1 Then
            xFgTot.TextMatrix(0, xFgTot.Cols - 3) = "Producido"
        Else
            xFgTot.TextMatrix(0, xFgTot.Cols - 3) = "Comprado"
        End If
        
        xFgTot.TextMatrix(0, xFgTot.Cols - 2) = "Total"
        xFgTot.TextMatrix(0, xFgTot.Cols - 1) = "Diferencia"
        xFgTot.ColWidth(xFgTot.Cols - 1) = 1100
        
        AnoTra = Year(Now)
        
        '----------------------------------------------------
        xObjBarra.Max = xFgTot.Rows - 1
        xObjBarra.Value = 1
        DoEvents
        '----------------------------------------------------

        
        For xFilaTotal = 1 To xFgTot.Rows - 1
            '----------------------------------------------------
            xObjBarra.Value = xFilaTotal
'            DoEvents
            '----------------------------------------------------
            xStkIni = SaldoActual(NulosN(xFgTot.TextMatrix(xFilaTotal, 1)), CDate("01/01/" & Format(AnoTra, "0000")), xFechaIni - 1, xCon)
            
            If xOrigen = 1 Then '--consultar lo producido
                xTotPro = HallarTotalProducido(NulosN(xFgTot.TextMatrix(xFilaTotal, 1)), xFechaIni)
            Else '--consultar lo comprado
                '--xtipo = Entradas
                xTotPro = SaldoActual(NulosN(xFgTot.TextMatrix(xFilaTotal, 1)), CDate(xFechaIni), Date, xCon, 1)
            End If
            
            xFgTot.TextMatrix(xFilaTotal, xFgTot.Cols - 4) = Format(xStkIni, FORMAT_MONTO)
            xFgTot.TextMatrix(xFilaTotal, xFgTot.Cols - 3) = Format(xTotPro, FORMAT_MONTO)
            xFgTot.TextMatrix(xFilaTotal, xFgTot.Cols - 2) = Format(xTotPro + xStkIni, FORMAT_MONTO)
            
            xTotal = ((xTotPro + xStkIni) - NulosN(xFgTot.TextMatrix(xFilaTotal, xFgTot.Cols - 5)))
            
            If xTotal > 0 Then
                FORMATO_CELDA xFgTot, xFilaTotal, xFgTot.Cols - 1, &HFF0000, True, , Format(xTotal, FORMAT_MONTO)
            Else
                FORMATO_CELDA xFgTot, xFilaTotal, xFgTot.Cols - 1, &HC0&, True, , Format(xTotal, FORMAT_MONTO)
            End If
            
        Next xFilaTotal
        
        xFgTot.FrozenCols = 6
    
    End If
    
    '--ordenar por nombre de producto
    GRID_ORDENAR xFgTot, 1, 4, 1, 4, flexSortGenericAscending
    
End Sub


Function HallarTotalProducido(xIdProducto As Long, Desde As Date) As Double
    Dim xRst As New ADODB.Recordset
    Dim xSQL As String
    
    xSQL = "SELECT Sum(pro_producciondet.cantidad) AS SumaDecantidad FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        & " WHERE (((pro_produccion.dia)>=CDate('" & Desde & "'))) GROUP BY pro_producciondet.iditem HAVING (((pro_producciondet.iditem)=" & xIdProducto & "))"

    RST_Busq xRst, xSQL, xCon
    
    If xRst.RecordCount <> 0 Then
        HallarTotalProducido = NulosN(xRst("SumaDecantidad"))
    Else
        HallarTotalProducido = 0
    End If
    Set xRst = Nothing
End Function

