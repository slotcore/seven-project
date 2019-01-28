VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JC_Varios"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum e_ARR_XX
'--SE USA ESTOS VALORES CON LA FINALIDAD DE SER UNICOS
    X_MES = 0
    X_TRIMESTRE = 1
    X_SEMESTRE = 2
End Enum

'----------constante paara funcion ocultar_botonfrm------------------
Const SC_SIZE = &HF000
Const SC_MOVE = &HF010
Const SC_MINIMIZE = &HF020
Const SC_MAXIMIZE = &HF030
Const SC_CLOSE = &HF060
Const SC_RESTORE = &HF120

Const MF_SEPARATOR = &H800
Const MF_BYPOSITION = &H400
Const MF_BYCOMMAND = &H0
'---------------------------------------------------------------------
Public Enum e_bdTipoOperacion
    e_Insertar = 0
    e_Modificar = 1
    e_Eliminar = 2
End Enum

Public Sub Llenar_Anyo(obj As Object, Optional Todos As Boolean = False)
    '--obj:: listbox,combobox
    On Error Resume Next
    With obj
        .Clear
        If Todos = True Then
         .AddItem "Todos los Años"
        End If
        For i = Year(Date) To 2001 Step -1
            .AddItem i
        Next i
        .ListIndex = 0
    End With
End Sub


Public Sub Llenar_Trimestre(obj As Object, Optional Todos As Boolean = False)
    '--obj:: listbox,combobox
    With obj
        .Clear
        If Todos = True Then
         .AddItem "Todos los Trimestres"
        End If
        .AddItem "Enero - Marzo"
        .AddItem "Abril - Junio"
        .AddItem "Julio - Septiembre"
        .AddItem "Octubre - Diciembre"
        .ListIndex = 0
    End With
End Sub

Public Sub Llenar_Semestre(obj As Object, Optional Todos As Boolean = False)
    '--obj:: listbox,combobox
    With obj
        .Clear
        If Todos = True Then
         .AddItem "Todos los Semestres"
        End If
        .AddItem "1er. Semestre"
        .AddItem "2do. Semestre"
        .ListIndex = 0
    End With
End Sub

Public Sub Llenar_Mes(obj As Object, Optional Todos As Boolean = False)
    '--obj:: listbox,combobox
    With obj
        .Clear
        If Todos = True Then
            .AddItem "Todos los Meses"
        End If
        .AddItem "Enero"
        .AddItem "Febrero"
        .AddItem "Marzo"
        .AddItem "Abril"
        .AddItem "Mayo"
        .AddItem "Junio"
        .AddItem "Julio"
        .AddItem "Agosto"
        .AddItem "Septiembre"
        .AddItem "Octubre"
        .AddItem "Noviembre"
        .AddItem "Diciembre"
       ''''''''''''''''
        If Todos = True Then
            .ListIndex = Month(Date)
        Else
            .ListIndex = Month(Date) - 1
        End If


    End With
End Sub

Public Sub ls_activar_chek(ls As Object, Optional N_VALOR As String = "")
    '--obj:: listbox,combobox
    '--SELECCIONARA TODOS LOS ITEM'S DE LA LISTA
    On Error Resume Next
    Dim k As Integer
    For k = ls.ListCount - 1 To 0 Step -1
        ls.ListIndex = k
        If N_VALOR <> "" Then
            If ls.Text = N_VALOR Then
                ls.Selected(k) = True
                Exit For
            End If
        Else
            ls.Selected(k) = True
        End If

    Next
    Err.Clear
End Sub


Public Function ArchivoExiste(Path As String) As Boolean
    On Error GoTo Fallo
    X = GetAttr(Path)
    ArchivoExiste = True
    Exit Function
Fallo:
    ArchivoExiste = False
End Function

Public Function EstaAbiertoPrograma(ClassName As String) As Boolean
    If FindWindow(ClassName, vbNullString) Then
        EstaAbiertoPrograma = True
    End If
End Function


Public Sub abrir_shape(Rst As ADODB.Recordset, sql As String)
    Dim bdshape As New ADODB.Connection
    'Dim Path As String
    Dim dsn As String
    Set bdshape = New ADODB.Connection
    bdshape.Provider = "msdatashape"


    bdshape.ConnectionString = xCon.ConnectionString
    bdshape.Open
    Rst.CursorLocation = adUseClient
    Rst.Open sql, bdshape, adOpenDynamic, adLockBatchOptimistic
    Set bdshape = Nothing
End Sub


Public Sub LimpiaText(txt As Variant, Optional Limpiar_Tag As Boolean = False)
    Err.Clear
    On Error Resume Next
    Dim obj As Variant
    For Each obj In txt
        If TypeName(obj) = "TextBox" Then
            obj.Text = ""
        ElseIf TypeName(obj) = "TextBoxFecha" Then
            obj.valor = ""
        ElseIf TypeName(obj) = "Label" Then
            obj.Caption = ""
        ElseIf TypeName(obj) = "DTPicker" Then
            obj.Value = ""
        ElseIf TypeName(obj) = "ComboBox" Or TypeName(obj) = "ListBox" Then
            obj.ListIndex = -1
        Else
            obj.Text = ""
            obj.ListIndex = -1
            obj.Text = ""
        End If

      If Limpiar_Tag = True Then
        obj.Tag = ""
      End If
    Next
    Err.Clear
End Sub

Public Function validar_numero(ascii As Integer) As Boolean
    Dim car As String
    car = Chr(ascii)
    If (car >= "0" And car <= "9") Or ascii = 8 Or ascii = 13 Or ascii = 46 Then
     validar_numero = True 'devuelve true si es numero
    Else
     validar_numero = band 'devuelve falso si no es numero
    End If
End Function

Public Function validar_letras(ascii As Integer) As Boolean
    Dim car As String
    car = UCase(Chr(ascii))
    If (car >= "A" And car <= "Z") Or ascii = 13 Or ascii = 32 Or ascii = 8 Or ascii = 241 Or ascii = 209 Then
     validar_letras = True
    Else
     validar_letras = band
    End If
End Function

Public Sub habilitar(txt As Variant, band As Boolean)
    Dim obj As Variant
    For Each obj In txt
      obj.Enabled = band
    Next
End Sub

Public Sub habilitar_Locked(txt As Variant, band As Boolean)
    Dim obj As Variant
    For Each obj In txt
      obj.Locked = band
    Next
End Sub

Public Sub Ocultar(txt As Variant, band As Boolean)
    On Error Resume Next
    Dim obj As Variant
    For Each obj In txt
        obj.Visible = band
    Next
    Err.Clear
End Sub

Public Function Validar(txt As Variant) As Integer
    Dim obj As Variant
    For Each obj In txt
        If Trim(obj.Tag) = "" And Trim(obj.Text) = "" Then
            Validar = obj.index
            Exit Function
        End If
    Next
    Validar = -1
End Function

Public Function validar_Combo(txt As Variant) As Integer
    Dim obj As Variant
    For Each obj In txt
        If Trim(obj.Text) = "" Then
            validar_Combo = obj.index
            Exit Function
        End If
    Next
    validar_Combo = -1
End Function

Public Function validar_Fecha(dtpk As Variant) As Integer
    Dim obj As Variant
    For Each obj In dtpk
        If TypeName(obj) = "DTPicker" Then
            If IsDate(obj.Value) = False And obj.Tag = "" Then
                validar_Fecha = obj.index
                Exit Function
            End If
        ElseIf TypeName(obj) = "TextBoxFecha" Then
            If IsDate(obj.valor) = False And obj.Tag = "" Then
                validar_Fecha = obj.index
                Exit Function
            End If
        End If
    Next
    validar_Fecha = -1
End Function

Public Sub LlenarCombo(Rst As ADODB.Recordset, cb, Campo As String, Optional Min As Boolean = False)
    cb.Clear
    If Rst.RecordCount > 0 Then Rst.MoveFirst
    While Not Rst.EOF
     If Min = False Then
        cb.AddItem Rst.Fields(Campo) & ""
     ElseIf Min = True Then
        cb.AddItem StrConv(Rst.Fields(Campo), 3)
     End If
     Rst.MoveNext
    Wend
End Sub


Public Sub CentrarFrm(frm As Object)
    '--frm formulario
    On Error Resume Next
    If frm.WindowState <> 2 Then
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
    ' frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
    End If
    Err.Clear
End Sub

Sub VentanaFlotante(HwndFrm As Long, hWndFrmPadre As Long)
    ret = SetParent(HwndFrm, hWndFrmPadre)
End Sub

Public Sub DEFINIR_RST_TMP(RST_TEMPORAL As ADODB.Recordset, _
                    RST_ORIGEN As ADODB.Recordset, _
                    Optional Val As String = "")

   Dim vCampos As Long

   vCampos = RST_ORIGEN.Fields.Count
   Set RST_TEMPORAL = Nothing
   Set RST_TEMPORAL = New ADODB.Recordset

   For i = 0 To vCampos - 1
       'RST_TEMPORAL.Fields.Append rst_origen.Fields(i).Name, rst_origen.Fields(i).Type, rst_origen.Fields(i).DefinedSize, adFldIsNullable
       RST_TEMPORAL.Fields.Append RST_ORIGEN.Fields(i).Name, RST_ORIGEN.Fields(i).Type, -1, adFldIsNullable
   Next
    '-SI ES RST_TEMPORAL ES PARA SELECIONAR ADICIONAR EL CAMPO VAL,
   If Val <> "" Then RST_TEMPORAL.Fields.Append Val, adVarChar, 2, adFldIsNullable
   RST_TEMPORAL.Open

End Sub



Public Sub CARGAR_RST_TMP(RST_TEMPORAL As ADODB.Recordset, _
                RST_ORIGEN As ADODB.Recordset, _
                Optional Val As String = "", _
                Optional Valor_Val As Integer = 0, _
                Optional F_UNSOLO_REGISTRO As Boolean = False)

    Dim N_CAMPO As String

    If F_UNSOLO_REGISTRO = False And RST_ORIGEN.RecordCount > 0 Then RST_ORIGEN.MoveFirst

    While Not RST_ORIGEN.EOF

        RST_TEMPORAL.AddNew

        For i = 0 To RST_ORIGEN.Fields.Count - 1
            N_CAMPO = RST_TEMPORAL.Fields(i).Name
            If RST_ORIGEN.Fields(N_CAMPO) & "" <> "" Then
               RST_TEMPORAL.Fields(N_CAMPO) = RST_ORIGEN.Fields(N_CAMPO)
            Else

            End If
        Next

        If Val <> "" Then
           RST_TEMPORAL.Fields(Val) = CStr(Valor_Val)
        End If
        If F_UNSOLO_REGISTRO = True Then Exit Sub
        RST_ORIGEN.MoveNext

    Wend

End Sub


Public Sub CARGAR_ARR_XX(ARR_XX() As String, T_TIPO As e_ARR_XX)
    If T_TIPO = X_MES Then
        ReDim ARR_XX(11, 2)
        ARR_XX(0, 0) = "Enero":        ARR_XX(0, 1) = "Ene":          ARR_XX(0, 2) = "1"
        ARR_XX(1, 0) = "Febrero":      ARR_XX(1, 1) = "Feb":          ARR_XX(1, 2) = "2"
        ARR_XX(2, 0) = "Marzo":        ARR_XX(2, 1) = "Mar":          ARR_XX(2, 2) = "3"
        ARR_XX(3, 0) = "Abril":        ARR_XX(3, 1) = "Abr":          ARR_XX(3, 2) = "4"
        ARR_XX(4, 0) = "Mayo":         ARR_XX(4, 1) = "May":          ARR_XX(4, 2) = "5"
        ARR_XX(5, 0) = "Junio":        ARR_XX(5, 1) = "Jun":          ARR_XX(5, 2) = "6"
        ARR_XX(6, 0) = "Julio":        ARR_XX(6, 1) = "Jul":          ARR_XX(6, 2) = "7"
        ARR_XX(7, 0) = "Agosto":       ARR_XX(7, 1) = "Ago":          ARR_XX(7, 2) = "8"
        ARR_XX(8, 0) = "Septiembre":   ARR_XX(8, 1) = "Sep":          ARR_XX(8, 2) = "9"
        ARR_XX(9, 0) = "Octubre":      ARR_XX(9, 1) = "Oct":          ARR_XX(9, 2) = "10"
        ARR_XX(10, 0) = "Noviembre":   ARR_XX(10, 1) = "Nov":         ARR_XX(10, 2) = "11"
        ARR_XX(11, 0) = "Diciembre":   ARR_XX(11, 1) = "Dic":         ARR_XX(11, 2) = "12"
    ElseIf T_TIPO = X_TRIMESTRE Then
        ReDim ARR_XX(3, 2)
        ARR_XX(0, 0) = "Enero - Marzo":         ARR_XX(0, 1) = "Ene-Mar":          ARR_XX(0, 2) = "1"
        ARR_XX(1, 0) = "Abril - Junio":         ARR_XX(1, 1) = "Abr-Jun":          ARR_XX(1, 2) = "2"
        ARR_XX(2, 0) = "Julio - Septiembre":    ARR_XX(2, 1) = "Jul-Sep":          ARR_XX(2, 2) = "3"
        ARR_XX(3, 0) = "Octubre - Diciembre":   ARR_XX(3, 1) = "Oct-Dic":          ARR_XX(3, 2) = "4"
    ElseIf T_TIPO = X_SEMESTRE Then
        ReDim ARR_XX(1, 2)
        ARR_XX(0, 0) = "Enero - Junio":         ARR_XX(0, 1) = "1er. Sem":         ARR_XX(0, 2) = "1"
        ARR_XX(1, 0) = "Julio - Diciembre":     ARR_XX(1, 1) = "2do. Sem":         ARR_XX(1, 2) = "2"
    End If

End Sub


Public Sub SHOW_ERROR(Optional FRM_NOMBRE As String, _
                      Optional EVENTO_NOMBRE As String, _
                      Optional SHOW_MSG As Boolean = True, _
                      Optional MSG_INCIAL As String = "")

    If Err.Number = 0 Then Exit Sub
    Dim N_ORIGEN As String
    N_ORIGEN = FRM_NOMBRE + "." + EVENTO_NOMBRE
    If Trim(N_ORIGEN) <> "" Then N_ORIGEN = vbCr + N_ORIGEN
    If SHOW_MSG = True Then
        MsgBox IIf(MSG_INCIAL = "", "", MSG_INCIAL + vbCr) + Err.Description + vbCr + _
               Err.Source + vbCr + _
               CStr(Err.Number) + _
               N_ORIGEN, vbCritical, "Error"
    End If
    Err.Clear
End Sub


Public Sub OPEN_CONEX_TMP(X_CONEX As ADODB.Connection, X_RUTA_BD As String)

    SeEjecutoEmp = False
    Dim xFun As New eps_librerias.FuncionesData

    xFun.F_BASEDATOS = X_RUTA_BD
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"

    Set X_CONEX = xFun.AbrirConeccion
    Set xFun = Nothing

End Sub

Public Function CARGAR_LISTA_ANYOS_ACTIVOS(ls As Object, _
        cnn As ADODB.Connection) As Boolean

    On Error GoTo ERROR
    Set xCon = cnn
    CargaDatosEmpresa
    '--ESTA FUNCION CARGA LOS AÑOS ACTIVOS DE UNA EMPRESA
    Dim CnnTmp As New ADODB.Connection '--CONEX TEMPORAL
    Dim rst_select As New ADODB.Recordset
    '--SI LA BASE DE BATOS PRINCIPAL EXISTE
    If ArchivoExiste(AP_RUTABD + "data.mdb") = False Then
        MsgBox "No existe la ruta a la Base de Datos Principal", vbCritical, "Mensaje..."
        Exit Function
    End If
    '--ABRIENDO LA CONEXION PARA OBTENER EL LISTADO DE LOS AÑOS
    OPEN_CONEX_TMP CnnTmp, AP_RUTABD + "data.mdb"
    If CnnTmp.State = 0 Then GoTo salir
    '--CARGAR DATOS
    RST_Busq rst_select, "SELECT anotra FROM mae_empresa WHERE numruc = '" + NumRUC + "' ORDER BY anotra DESC ", CnnTmp
    If rst_select.State = 0 Then GoTo salir
    LLENAR_LISTA rst_select, ls, "anotra", True

salir:
    Set rst_select = Nothing
    CnnTmp.Close

    CARGAR_LISTA_ANYOS_ACTIVOS = True

    Exit Function
ERROR:
    Set rst_select = Nothing
    If CnnTmp.State = 1 Then CnnTmp.Close
    SHOW_ERROR
End Function

Public Sub LLENAR_LISTA(Rst As ADODB.Recordset, _
            ob As Object, _
            Campo As String, _
            Optional Min As Boolean = False)

    '--ob:: listbox o combobox
    '--campo::campo a llenar
    '--min::texto en minuscula primer caracter
    ob.Clear
    If Rst.State = 0 Then Exit Sub
    If Rst.RecordCount > 0 Then Rst.MoveFirst
    While Not Rst.EOF
     If Min = False Then
        ob.AddItem Rst.Fields(Campo) & ""
     ElseIf Min = True Then
        ob.AddItem StrConv(Rst.Fields(Campo), 3)
     End If
     Rst.MoveNext
    Wend
End Sub


Public Sub PONER_COLOR_ESTADO(xCONEX As ADODB.Connection, obj As Object, ID_ESTADO As Integer)
    '--ESTA FUNCTION CAMBIA EL COLOR AL OBJETO SEGUN ID_ESTADO (VER TABLA mae_estados)
    Dim RST_TMP As New ADODB.Recordset
    If xCONEX.State = 0 Then Exit Sub
    RST_Busq RST_TMP, "SELECT mae_estados.color, mae_estados.descripcion " _
                      + vbCr + " From mae_estados " _
                      + vbCr + " WHERE (((mae_estados.id)=" + CStr(ID_ESTADO) + "));", xCONEX
    If RST_TMP.State = 0 Then Exit Sub
    obj.Caption = RST_TMP.Fields("descripcion")
    If IsNull(RST_TMP.Fields("color")) = False Then obj.ForeColor = Val(RST_TMP.Fields("color"))
    Set RST_TMP = Nothing

End Sub


Public Sub CARGAR_DLL_EPSBUSCAR( _
                                x_CON As ADODB.Connection, _
                                x_RS As ADODB.Recordset, _
                                x_SQL As String, _
                                xCampos() As String, _
                                x_TITULO As String, _
                                X_ORDEN As String, _
                                x_CAMPO_BUSCA As String, _
                                Optional x_FORMA_BUSCA As FormasBusqueda = Principio, _
                                Optional x_CRITERIO As String = "", _
                                Optional xRstTmp As ADODB.Recordset)

    On Error GoTo ERROR

    Dim xform As New eps_librerias.FormBuscar
    Dim xxRst As New ADODB.Recordset

    Err.Clear
    If NulosC(x_SQL) <> "" Then
        xform.SQLCad = x_SQL
    Else
        Set xform.RstConsulta = xRstTmp
    End If
    xform.TITULO = x_TITULO
    xform.FormaBusca = x_FORMA_BUSCA
    xform.Ordenado = X_ORDEN
    xform.CampoBusca = x_CAMPO_BUSCA
    xform.Criterio = x_CRITERIO
    Set xform.Coneccion = x_CON
    Set xxRst = Nothing
    Set xxRst = xform.BuscarReg(xCampos)
    Erase xCampos
    Set x_RS = xxRst
    Set xform = Nothing
    Exit Sub
ERROR:
'Resume
    Erase xCampos
    Set xform = Nothing
    SHOW_ERROR
End Sub


Public Sub CARGAR_DLL_EPSBUSCAR_SEL( _
                                x_CON As ADODB.Connection, _
                                x_RS As ADODB.Recordset, _
                                x_SQL As String, _
                                xCampos() As String, _
                                x_TITULO As String, _
                                Optional xRstTmp As ADODB.Recordset)
  On Error GoTo ERROR

    Dim xform As New eps_librerias.FormSeleccion
    Err.Clear
    If NulosC(x_SQL) <> "" Then
        xform.SQLCad = x_SQL
    Else
        'Set xform.RstConsulta = xRstTmp
    End If

    xform.TITULO = x_TITULO
    Set xform.Coneccion = x_CON
    Set x_RS = Nothing
    Set x_RS = xform.Seleccionar(xCampos)
    Set xform = Nothing
    Erase xCampos
    Exit Sub

ERROR:
    Erase xCampos
    Set xform = Nothing
    SHOW_ERROR
End Sub


Public Sub CARGAR_DLL_EPSBUSCAR_FILTRO( _
                                x_CON As ADODB.Connection, _
                                x_RS As ADODB.Recordset, _
                                xCampos() As String, _
                                TDB_GRID As Object)
    On Error GoTo ERROR
    '--TDB_GRID ES UN TdbGrid
    Dim xform As New eps_librerias.FormFiltrar
    'CONDICION DE INGRESO PARA XCAMPOS()
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Err.Clear
    Set xform.Coneccion = x_CON
    Set xform.Rst = x_RS        'recorset que llena el grid
    Set x_RS = xform.FiltrarReg(xCampos)
    Set TDB_GRID.DataSource = x_RS
    TDB_GRID.Refresh
    Set xform = Nothing
    Erase xCampos
    Exit Sub
ERROR:
    Erase xCampos
    Set xform = Nothing
    SHOW_ERROR
End Sub


'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
Public Sub FrmOcultarBoton(FrmhWnd As Long, mTipo As Integer)
    '--mTipo
    Dim hwnd&, hMenu&, Success&
    Dim i%

    hwnd = FrmhWnd
    hMenu = GetSystemMenu(hwnd, 0)

    'Quitar todos (va de 0 a 8)
'    If mTipo = -1 Then
'        For i = 8 To 0 Step -1
'             Success = DeleteMenu(hMenu, i, MF_BYPOSITION)
'        Next
'        Exit Sub
'    End If

    'Usa esto para quitar los menús que te interesen:
    Select Case mTipo
        Case 0: Success = DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)
        Case 1: Success = DeleteMenu(hMenu, SC_MOVE, MF_BYCOMMAND)
        Case 2: Success = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
        Case 3: Success = DeleteMenu(hMenu, SC_MINIMIZE, MF_BYCOMMAND)
        Case 4: Success = DeleteMenu(hMenu, SC_MAXIMIZE, MF_BYCOMMAND)
        Case 6: Success = DeleteMenu(hMenu, SC_RESTORE, MF_BYCOMMAND)
    End Select

End Sub


Public Function SincronizarBD(cn As ADODB.Connection, nTabla As String, mIDTabla, QueHace) As Boolean

    '--ArrCampos(?,0)= Nombre del Campo;    ArrCampos(?,1)=Valor;   ArrCampos(?,2)=Tipo de Dato (N=Numero,F=Fecha,C=Texto)
    '--mTipoOperacion::1=Insertar;      2=Modificar;        3=Eliminar
    '-------------
    '--En ArrCampos1(?.2) Contiene la misma estructura que ArrCampos(?,2), solo que en este array se usara cuando
    '--mTipoOperacion = Modificar, Eliminar, los valores se usaran para condicionar la sentencia sql
    '--ejm. UPDATE NomTabla SET {Valores de ArrCampos(?,2)} WHERE {Valores de ArrCampos1(?,2)}

    Err.Clear
    'On Error GoTo ERROR
    Set xCon = cn
    CargaDatosEmpresa
    '----------------------------
    Dim RstRuta As New ADODB.Recordset  '--CARGA RUTAS DE BD'S
    Dim RstTmp As New ADODB.Recordset
    Dim CnnTmp As New ADODB.Connection  '--CONEXIÓN TEMPORAL
    Dim RstTabla As New ADODB.Recordset '--
    RST_Busq RstRuta, "SELECT ruta,anotra,nomemp FROM mae_empresa WHERE sincronizar = -1 ORDER BY anotra ASC ", xCon
    '--CARGAR RST TEMPORAL
    If RstRuta.RecordCount = 0 Then
        MsgBox "No hay Base de Datos A Sincronizar", vbInformation
        Set RstRuta = Nothing
        Exit Function
    End If
    RstRuta.MoveFirst
    '----
    DoEvents
    '-------
    Dim mRow&, mColCampo&
    Dim RstMaestro As New ADODB.Recordset   '--
    Dim mID As Long

    RST_Busq RstMaestro, "SELECT * FROM " + nTabla + " WHERE id = " + CStr(mIDTabla) + " ;", xCon
    If RstMaestro.BOF = True Or RstMaestro.EOF = True Or RstMaestro.RecordCount = 0 Then
        MsgBox "No se encuentra el registro en la Base Maestra", vbExclamation, xTitulo
        Exit Function
    End If

    nTabla = LCase(nTabla)
    Select Case nTabla
        Case "alm_inventario"
            ReDim xCampos(11)
            xCampos(0) = "id"
            xCampos(1) = "idmae"
            xCampos(2) = "codpro"
            xCampos(3) = "descripcion"
            xCampos(4) = "desctec"
            xCampos(5) = "idunimed"
            xCampos(6) = "idmon"
            xCampos(7) = "tippro"
            xCampos(8) = "idclas"
            xCampos(9) = "idsubclas"
            xCampos(10) = "idfam"
            xCampos(11) = "caracteristica"

        Case "mae_tipoproducto"
            ReDim xCampos(4)
            xCampos(0) = "id"
            xCampos(1) = "idmae"
            xCampos(2) = "descripcion"
            xCampos(3) = "prefijo"
            xCampos(4) = "codsun"

            Case "mae_familia"
            ReDim xCampos(4)
            xCampos(0) = "id"
            xCampos(1) = "idmae"
            xCampos(2) = "descripcion"
            xCampos(3) = "idtippro"
            xCampos(4) = "prefijo"

        Case "mae_clase"
            ReDim xCampos(5)
            xCampos(0) = "id"
            xCampos(1) = "idmae"
            xCampos(2) = "idfam"
            xCampos(3) = "descripcion"
            xCampos(4) = "observacion"
            xCampos(5) = "prefijo"

        Case "mae_subclase"
            ReDim xCampos(5)
            xCampos(0) = "id"
            xCampos(1) = "idmae"
            xCampos(2) = "idclas"
            xCampos(3) = "descripcion"
            xCampos(4) = "observacion"
            xCampos(5) = "prefijo"


        Case "mae_cliente", "mae_prov"
            If nTabla = "mae_cliente" Then ReDim xCampos(16)
            If nTabla = "mae_prov" Then ReDim xCampos(21)

            xCampos(0) = "id"
            xCampos(1) = "idmae"
            xCampos(2) = "tipper"
            xCampos(3) = "numruc"
            xCampos(4) = "nombre"
            xCampos(5) = "dir"
            xCampos(6) = "iddis"
            xCampos(7) = "iddep"
            xCampos(8) = "tel"
            xCampos(9) = "fax"
            xCampos(10) = "email"
            xCampos(11) = "pagweb"
            xCampos(12) = "activo"
            If nTabla = "mae_cliente" Then
                xCampos(13) = "nomcli1"
                xCampos(14) = "nomcli2"
                xCampos(15) = "apecli1"
                xCampos(16) = "apecli2"
            ElseIf nTabla = "mae_prov" Then
                xCampos(13) = "nompro1"
                xCampos(14) = "nompro2"
                xCampos(15) = "apepro1"
                xCampos(16) = "apepro2"
                xCampos(17) = "ageret"
                xCampos(18) = "ageper"
                xCampos(19) = "tipdoc"
                xCampos(20) = "email"
                xCampos(21) = "numdoc"
            End If

        Case "mae_moneda"

        Case "mae_unidades"

        Case Else
            MsgBox "No a especificado el nombre de la tabla", vbInformation, xTitulo
            Exit Function
    End Select

    '------------------------------------------------
    '--TRABAJAR ENTRE LAS BASE DE TADOS
    Do While Not RstRuta.EOF
        '--SI EL ARCHIVO EXISTE
        If ArchivoExiste(AP_RUTABD + RstRuta.Fields(0) & "") = False Then
            MsgBox "No existe la ruta a la Base de Datos" + vbCr + _
                            "Empresa: " + RstRuta.Fields("nomemp") & "" + vbCr + _
                            "Año: " + CStr(RstRuta.Fields(1)), vbCritical, xTitulo
            GoTo salir
        End If
        '--ABRIENDO LA CONEXION A LA BASE DE DATOS
        OPEN_CONEX_TMP CnnTmp, AP_RUTABD + RstRuta.Fields(0) & ""
        If CnnTmp.State = 0 Then
            MsgBox "Error en la Conexión a " + AP_RUTABD + RstRuta.Fields(0) & "", vbExclamation, xTitulo
            GoTo salir
        End If
        '--------------------------------------
        '--EJECUTANDO LA SENTENCIA
        CnnTmp.BeginTrans

        If QueHace = 2 Then '--MODIFICAR
            '--COMPROBAR QUE EXISTA EL REGISTRO, SI EXISTE SE CREARA EL REGISTRO DE NUEVO
            RST_Busq RstTabla, "SELECT * FROM " + nTabla + " WHERE idmae = " + CStr(mIDTabla) + "; ", CnnTmp
            If RstTabla.BOF = True Or RstTabla.EOF = True Or RstTabla.RecordCount = 0 Then
                Set RstTabla = Nothing
                GoTo IrInsertar1:
            End If
        End If

        If QueHace = 1 Then '--NUEVO
IrInsertar1:
            RST_Busq RstTabla, "SELECT TOP 1 * FROM " + nTabla + " ;", CnnTmp
            mID = 0
            mID = HallaCodigoTabla(nTabla, CnnTmp, "id")
            RstTabla.AddNew
            RstTabla.Fields("id") = mID
        End If

        For mRow = 0 To UBound(xCampos())
            If LCase(xCampos(mRow)) <> "id" Then
                For mColCampo = 0 To RstMaestro.Fields.Count - 1
                    If LCase(xCampos(mRow)) = LCase(RstMaestro.Fields(mColCampo).Name) Then
                        RstTabla.Fields(xCampos(mRow)) = RstMaestro.Fields(mColCampo)
                        Exit For
                    ElseIf LCase(xCampos(mRow)) = "idmae" Then
                        RstTabla.Fields("idmae") = mIDTabla
                        Exit For
                    End If
                Next mColCampo
            End If
        Next mRow
        RstTabla.Update
        '--------------------------------------
        CnnTmp.CommitTrans
        CnnTmp.Close
        '--------------------------------------
        RstRuta.MoveNext
    Loop
salir:
    Set RstRuta = Nothing
    SincronizarBD = True
    '--LIMPIAR EL ARRAY
    Erase xCampos()
    Exit Function

ERROR:
'    Resume
    Set RstTmp = Nothing
    CnnTmp.RollbackTrans
    CnnTmp.Close
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error - Unificado"
End Function

Public Function UNIFICADO_RST(CnnLocal As ADODB.Connection, nSQL As String) As ADODB.Recordset

    On Error GoTo ERROR
    Dim RstTmp As New ADODB.Recordset
    Set xCon = CnnLocal

    CargaDatosEmpresa
    '--ESTA FUNCION CARGA LOS AÑOS ACTIVOS DE UNA EMPRESA
    Dim CnnTmp As New ADODB.Connection '--CONEX TEMPORAL
    '--SI LA BASE DE BATOS PRINCIPAL EXISTE
    If ArchivoExiste(AP_RUTABD + "data.mdb") = False Then
        MsgBox "No existe la ruta a la Base de Datos Principal", vbCritical, "Mensaje..."
        Exit Function
    End If
    '--ABRIENDO LA CONEXION PARA OBTENER EL LISTADO DE LOS AÑOS
    OPEN_CONEX_TMP CnnTmp, AP_RUTABD + "data.mdb"
    If CnnTmp.State = 0 Then GoTo salir
    '--CARGAR DATOS
    RST_Busq RstTmp, nSQL, CnnTmp
salir:
    Set UNIFICADO_RST = RstTmp
    CnnTmp.Close
    Exit Function
ERROR:
    If CnnTmp.State = 1 Then CnnTmp.Close
    SHOW_ERROR
End Function


Public Sub UNIFICADO_CONEX(CnnLocal As ADODB.Connection, CnnUnif As ADODB.Connection)

    Set xCon = CnnLocal
    CargaDatosEmpresa
    '--SI LA BASE DE BATOS PRINCIPAL EXISTE
    If ArchivoExiste(AP_RUTABD + "data.mdb") = False Then
        MsgBox "No existe la ruta a la Base de Datos Principal", vbCritical, "Mensaje..."
        Exit Sub
    End If
    '--ABRIENDO LA CONEXION PARA OBTENER EL LISTADO DE LOS AÑOS
    OPEN_CONEX_TMP CnnUnif, AP_RUTABD + "data.mdb"

End Sub


Public Function ConvertHora(mTotalSegundos) As String
    '--esta funcion devuelve la hora segun la cantidad de segundos como parametro
    Dim mTmp&
    Dim mHora As Integer
    Dim mMin As Integer
    Dim mSeg As Integer

    mTmp = mTotalSegundos
    mHora = (mTotalSegundos - (mTotalSegundos Mod 3600)) / 3600

    mTmp = mTotalSegundos Mod 3600

    mMin = (mTmp - (mTmp Mod 60)) / 60
    mSeg = mTmp Mod 60

    ConvertHora = Format(mHora, "00") & ":" & Format(mMin, "00") & ":" & Format(mSeg, "00")

End Function

Public Function ConvertSeg(dHora As String) As Long
    '--esta funcion devuelve los segundos segun la hora como parametro
    Dim mHora&
    Dim mMin&
    Dim mSeg&
    Dim nTmp As String

    Dim PosSeparador As Integer '--indica la separacion ":"

    If IsDate(dHora) = False Then
        ConvertSeg = -1
        '--se asume que el valor ingresado es de la sig. forma
        '-- "1008:15:20" hora:minuto:segundo

'        nTmp = Right(dHora, 2)
'        mSeg = NulosN(Right(dHora, 2))
'
'        dHora = mID(dHora, 1, Len(dHora) - 3)
'
'        mMin = NulosN(Right(dHora, 2)) * 60
'        nTmp = Right(dHora, 2)
'        dHora = Replace(dHora, ":" & nTmp, "")
'        mHora = NulosN(dHora) * 3600

        '**********************************
        '--de la hora
        If dHora <> "" Then
            nTmp = dHora
            PosSeparador = InStr(nTmp, ":")

            mHora = NulosN(mID(nTmp, 1, PosSeparador - 1)) * 3600

            '--del minuto
            If nTmp <> "" Then
                nTmp = mID(nTmp, PosSeparador + 1)
                PosSeparador = InStr(nTmp, ":")
                mMin = NulosN(mID(nTmp, 1, PosSeparador - 1)) * 60
            End If
            '--del segundo
            If nTmp <> "" Then
                nTmp = mID(nTmp, PosSeparador + 1)
                mSeg = NulosN(nTmp)
            End If
        End If
        '**********************************

    Else

        mHora = Hour(dHora) * 3600
        mMin = Minute(dHora) * 60
        mSeg = Second(dHora)
    End If

    ConvertSeg = mHora + mMin + mSeg

End Function


Public Sub RstRegistroEliminar(Rst As ADODB.Recordset, nCampo As String, _
                               nValor As String, fEsNumero As Boolean)

    '===================================================================================================
    'Creado : 05/02/08 Johan Castro
    'Propósito: Eliminar registros dependiendo de parametro ingresado
    '
    'Entradas:  Rst = Recorset que sera tratado
    '           nCampo = Campo a filtrar
    '           nValor = Valor del campo para hacer un filtro
    '           fEsNumero = Indica si es numerico el campo
    '                       True = Numerico; False = Cadena
    '
    'Resultados: Registro(s) eliminado(s)
    '
    '===================================================================================================
    Dim Apostrofe As String
    If Rst.State = 0 Then Exit Sub
    Rst.Filter = ""
    If Rst.RecordCount = 0 Then Exit Sub

    If fEsNumero = False Then Apostrofe = "'"
    If fEsNumero = True Then Apostrofe = ""

    Rst.MoveFirst
    '--buscando el registro
    '--con filtro
    Rst.Filter = nCampo & " = " & Apostrofe & nValor & Apostrofe
    If Rst.EOF = False Or Rst.BOF = False Or Rst.RecordCount <> 0 Then Rst.MoveFirst
    Do While Not Rst.EOF
        Rst.Delete
        Rst.Update
        Rst.MoveNext
    Loop
    Rst.Filter = ""
    If Rst.EOF = False Or Rst.BOF = False Or Rst.RecordCount <> 0 Then Rst.MoveFirst

End Sub

Public Function RstRegistroSumar(Rst As ADODB.Recordset, nCampo As String, _
                               Optional nCampoCondicion As String = "", _
                               Optional nValorCondicion As String = "", _
                               Optional nTipoDato As String = "N", _
                               Optional fEsIgual As Boolean = True) As Double
    '===================================================================================================
    'Creado : 08/02/08 Johan Castro
    'Propósito: Sumar los registros de un campo, aplicando un filtro previo(opcional)
    '
    'Entradas:  Rst = Recorset que sera tratado
    '           nCampo = Campo a utilizar
    '           nCampoCondicion = Campo para hacer un filtro, si es vacio se suman todos los registros
    '           nValorCondicion = Indica el valo para hacer el filtro, depende del parametro nCampoCondicion
    '           nTipoDato = Tipo de dato del campo condicion, depende del parametro nCampoCondicion
    '                       N:Numero; C:Caracter; F:fecha
    '           fEsIgual = Indica que se va sumar, depende del parametro nCampoCondicion
    '                      true = sumar todos los iguales al filtro; false = sumar todos los diferente al filtro
    '                      Nota: Si no hay filtro se suman todos los registros
    'Resultados: Numero, resultado de la suma
    '
    '===================================================================================================

    Dim Acumulado As Double
    Dim fSeEncuentraCampo As Boolean
    If Rst.State = 0 Then Exit Function
    If Rst.RecordCount = 0 Then Exit Function
    Rst.MoveFirst
    '--buscando el registro
    If Rst.EOF = True Or Rst.BOF = True Or Rst.RecordCount = 0 Then Rst.MoveFirst
    Do While Not Rst.EOF
        If nCampoCondicion <> "" Then
            Select Case UCase(nTipoDato)
                Case "N"
                    If fEsIgual = True Then
                        If (NulosN(Rst(nCampoCondicion)) = NulosN(nValorCondicion)) Then Acumulado = Acumulado + NulosN(Rst(nCampo))
                    Else
                        If (NulosN(Rst(nCampoCondicion)) <> NulosN(nValorCondicion)) Then Acumulado = Acumulado + NulosN(Rst(nCampo))
                    End If

                Case "C"
                    If fEsIgual = True Then
                        If (NulosC(Rst(nCampoCondicion)) = NulosC(nValorCondicion)) Then Acumulado = Acumulado + NulosN(Rst(nCampo))
                    Else
                        If (NulosC(Rst(nCampoCondicion)) <> NulosC(nValorCondicion)) Then Acumulado = Acumulado + NulosN(Rst(nCampo))
                    End If

                Case "F"
                    If IsDate(Rst(nCampoCondicion)) = True Then
                        If fEsIgual = True Then
                            If (CDate(Rst(nCampoCondicion)) = CDate(nValorCondicion)) Then Acumulado = Acumulado + NulosN(Rst(nCampo))
                        Else
                            If (CDate(Rst(nCampoCondicion)) <> CDate(nValorCondicion)) Then Acumulado = Acumulado + NulosN(Rst(nCampo))
                        End If
                    End If
            End Select
        Else
            Acumulado = Acumulado + NulosN(Rst(nCampo))
        End If
        Rst.MoveNext
    Loop

    RstRegistroSumar = Acumulado

End Function

Public Function RstRegistroGenerarId(Rst As ADODB.Recordset, nCampo As String, nCampoCondicion As String, _
                               Optional nTipoIn As String = "IN", Optional fEsNumero As Boolean = True) As String
    '===================================================================================================
    'Creado : 11/02/08 Johan Castro
    'Propósito: Unir varios registros en una cadena
    '
    'Entradas:  Rst = Recorset que sera tratado
    '           nCampo = Campo a utilizar
    '           nCampoCondicion =
    '           nTipoIn = Indica si va utilizar la expresion IN (xxxx) o NOT IN (xxxxx);
    '                     Si el parametro es vacio solo se muestra cadena con separador y sin comillas
    '           fEsNumero = Indica si el campo a concatenar es numerico o texto
    '
    'Resultados: Cadena segun el campo indicado

    'Modificado: 16/12/10 Johan Castro
    '            Modificar el resultado de la funcion,
    '            cuando nTipoIn es vacio no mostrar expresion [IN][NOT IN]
    '===================================================================================================

    '--pendiente incorpar condicional
    Dim nSQL As String
    nSQL = ""
    If Rst.State = 0 Then Exit Function
    Rst.Filter = ""
    If Rst.RecordCount = 0 Then Exit Function

    Rst.MoveFirst

    Do While Not Rst.EOF
        If fEsNumero = False Then
            If Trim(nTipoIn) <> "" Then
                nSQL = nSQL & "'" & NulosC(Rst.Fields(nCampo)) & "',"
            Else
                nSQL = nSQL & NulosC(Rst.Fields(nCampo)) & ","
            End If
        Else
            nSQL = nSQL & NulosN(Rst.Fields(nCampo)) & ","
        End If
        Rst.MoveNext
    Loop

    If Trim(nTipoIn) <> "" Then
        If nSQL <> "" Then nSQL = " " + nCampo + " " + nTipoIn + " (" + Left(nSQL, Len(nSQL) - 1) + ") "
    Else
        If nSQL <> "" Then nSQL = Left(nSQL, Len(nSQL) - 1)
    End If

    RstRegistroGenerarId = nSQL

End Function


Public Sub RstRegistroReemplazar(Rst As ADODB.Recordset, nCampo As String, _
                               nValor As String, fEsNumero As Boolean, _
                               nCampoPaste As String, nValorPaste As String)

    '===================================================================================================
    'Creado : 15/02/08 Johan Castro
    'Propósito: Reemplazar un dato en un(os) registro(s)
    '
    'Entradas:  Rst = Recorset que sera tratado
    '           nCampo = Campo a aplicar filtro
    '           fEsNumero = Indica si el campo a filtrar es numerico o texto
    '                       False = Texto; True = Numérico
    '           nCampoPaste = Campo cuyo contenido sera reemplado, este campo puede ser igual al parametro nCampo
    '           nValorPaste = Valor para reemplazar.
    'Resultados: Recordset que fueron reemplazados sus registros
    '
    '===================================================================================================


    Dim Apostrofe As String
    If Rst.State = 0 Then Exit Sub
    Rst.Filter = ""
    If Rst.RecordCount = 0 Then Exit Sub

    If fEsNumero = False Then Apostrofe = "'"
    If fEsNumero = True Then Apostrofe = ""

    Rst.MoveFirst
    '--con filtro
    Rst.Filter = nCampo & " = " & Apostrofe & nValor & Apostrofe
    If Rst.EOF = False Or Rst.BOF = False Or Rst.RecordCount <> 0 Then
        Rst.Fields(nCampoPaste) = nValorPaste
        Rst.Update
    End If
    Rst.Filter = ""

End Sub


'Public Sub FrmNavegador(frm As Form, Toolbar As Variant, KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'    If KeyCode = vbKeyN And Shift = 4 And Toolbar.Buttons(1).Enabled = True Then 'nuevo
'      CallByName frm, "accionesToolBar", VbMethod, 1
'    ElseIf KeyCode = vbKeyG And Shift = 4 And Toolbar.Buttons(2).Enabled = True Then 'guardar
'      ' CallByName frm, "accionesToolBar", VbMethod, 2
'    ElseIf KeyCode = vbKeyD And Shift = 4 And Toolbar.Buttons(3).Enabled = True Then 'modificar
'       'CallByName frm, "accionesToolBar", VbMethod, 3
'    ElseIf KeyCode = vbKeyM And Shift = 4 And Toolbar.Buttons(4).Enabled = True Then 'eliminar
'       'CallByName frm, "accionesToolBar", VbMethod, 4
'    ElseIf KeyCode = vbKeyC And Shift = 4 And Toolbar.Buttons(5).Enabled = True Then 'cancelar
'       'CallByName frm, "accionesToolBar", VbMethod, 5
'    ElseIf KeyCode = vbKeyR And Shift = 4 And Toolbar.Buttons(13).Enabled = True Then  'salir
'       'CallByName frm, "accionesToolBar", VbMethod, 13
'    End If
'    Err.Clear
'End Sub

Public Function RstRegistroDatos(Rst As ADODB.Recordset, _
                                  Optional nTipo As Integer = 0, Optional mSeparador As String = "|") As String
    '===================================================================================================
    'Creado : 25/07/08
    'Propósito: Mostrar la informacion que contiene el recorset, tanto en registros como campos
    '
    'Entradas:  RstTmp = recorset que sera tratado
    '           nTipo = Indica el la forma de mostrar los datos (0::todos los registros,1::todos los campos)
    '           mSeparador = Indica el caracter que separa de las demas columnas
    '
    'Resultados: Se presentan dos informes
    '            1::lista de registros
    '            2::lista de campos; Una sola cadena
    '===================================================================================================

    Dim tTexto As Variant
    Dim k&
    tTexto = ""
    If Rst.State = 0 Then Exit Function
    If nTipo = 0 Then
        If Rst.RecordCount <> 0 Then Rst.MoveFirst

        Do While Not Rst.EOF
            For k = 0 To Rst.Fields.Count - 1
                tTexto = tTexto & Rst.Fields(k) & mSeparador
            Next k
            tTexto = tTexto & vbCrLf
            Rst.MoveNext
        Loop

    ElseIf nTipo = 1 Then
        For k = 0 To Rst.Fields.Count - 1
            tTexto = tTexto & Rst.Fields(k).Name & mSeparador
        Next k
    End If
    RstRegistroDatos = tTexto
End Function

Public Function RstRegistroCuenta(Rst As ADODB.Recordset, _
                               Optional nCampoCondicion As String = "", _
                               Optional nValorCondicion As String = "", _
                               Optional nTipoDato As String = "N", _
                               Optional fEsIgual As Boolean = True) As Double

    '===================================================================================================
    'Creado : 25/07/08
    'Propósito: Contar los registros, que puede ser con una condicional
    '
    'Entradas:  rst = recorset que sera tratado
    '           nCampoCondicion  = nombre del campo si desea contar de un campo en especial
    '           nValorCondicion  = valor que depende de nCampoCondicion
    '           nTipoDato = es el tipo de dato de la comparacion( N:Numero,C:Caracter,F:fecha)
    '           fEsIgual = Indica si se aplicara el conteo con el campo de comparacion ; true:: son iguales; false:: es diferente
    '
    'Resultados: numerico que indica la cantidad de registros
    '===================================================================================================

    Dim Acumulado As Double
    Dim fSeEncuentraCampo As Boolean
    If Rst.State = 0 Then Exit Function
    If Rst.RecordCount = 0 Then Exit Function
    Rst.MoveFirst
    '--buscando el registro
    If Rst.EOF = True Or Rst.BOF = True Or Rst.RecordCount = 0 Then Rst.MoveFirst
    Do While Not Rst.EOF
        If nCampoCondicion <> "" Then
            Select Case UCase(nTipoDato)
                Case "N"
                    If fEsIgual = True Then
                        If (NulosN(Rst(nCampoCondicion)) = NulosN(nValorCondicion)) Then Acumulado = Acumulado + 1
                    Else
                        If (NulosN(Rst(nCampoCondicion)) <> NulosN(nValorCondicion)) Then Acumulado = Acumulado + 1
                    End If

                Case "C"
                    If fEsIgual = True Then
                        If (NulosC(Rst(nCampoCondicion)) = NulosC(nValorCondicion)) Then Acumulado = Acumulado + 1
                    Else
                        If (NulosC(Rst(nCampoCondicion)) <> NulosC(nValorCondicion)) Then Acumulado = Acumulado + 1
                    End If

                Case "F"
                    If IsDate(Rst(nCampoCondicion)) = True Then
                        If fEsIgual = True Then
                            If (CDate(Rst(nCampoCondicion)) = CDate(nValorCondicion)) Then Acumulado = Acumulado + 1
                        Else
                            If (CDate(Rst(nCampoCondicion)) <> CDate(nValorCondicion)) Then Acumulado = Acumulado + 1
                        End If
                    End If
            End Select
        Else
            Acumulado = Acumulado + 1
        End If
        Rst.MoveNext
    Loop

    RstRegistroCuenta = Acumulado

End Function

Public Function ConvertHoraFaccion(dHoraIni As String, dHoraFin As String) As Double
    '===================================================================================================
    'Creado : 24/09/08
    'Propósito: convertir una hora en especial a faccion. ejm 03:30 >> 3.5
    '
    'Entradas:  dHoraIni = Hora inicial
    '           dHoraIni = Hora final
    '
    'Resultados: hora convertida a fraccion, 0 = caso contrario sea error o falte alguna hora
    '===================================================================================================

    '--esta funcion devuelve los segundos segun la hora como parametro
    Dim mHora As Double
    Dim mMin As Double
    Dim mSeg As Double
    Dim nTmp As String

    Dim dHora As Date

    If IsDate(dHoraIni) = False Or IsDate(dHoraFin) = False Then
        ConvertHoraFaccion = 0
        Exit Function
    End If

    If CDate(dHoraIni) < CDate(dHoraFin) Then
        dHora = TimeValue(dHoraIni) - TimeValue(dHoraFin)
    Else
        dHora = CDate((TimeValue(dHoraFin) - TimeValue("00:00")) + (TimeValue("23:59") - TimeValue(dHoraIni)) + CDate("00:01"))
    End If

    mHora = Hour(dHora)
    mMin = Minute(dHora)
    mMin = IIf(mMin = 0, 0, mMin / 60)

    ConvertHoraFaccion = mHora + mMin

End Function

Public Function RstRegistroBuscaCampo(Rst As ADODB.Recordset, nCampo As String) As Boolean
    '===================================================================================================
    'Creado : 25/07/08
    'Propósito: Mostrar la informacion que contiene el recorset, tanto en registros como campos
    '
    'Entradas:  RstTmp = recorset que sera tratado
    '           nTipo = Indica el la forma de mostrar los datos (0::todos los registros,1::todos los campos)
    '
    'Resultados: 1::lista de registros; 2::lista de campos
    '===================================================================================================

    Dim tTexto As Variant
    Dim fEncontrado As Boolean
    Dim k&
    tTexto = ""
    If Rst.State = 0 Then Exit Function
    fEncontrado = False
    For k = 0 To Rst.Fields.Count - 1
        If UCase(Rst.Fields(k).Name) = UCase(nCampo) Then
            fEncontrado = True
            Exit For
        End If
    Next k

    RstRegistroBuscaCampo = fEncontrado

End Function


Private Sub Class_Initialize()

    Nomsis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
End Sub

Public Function DiferenciaHoras(dHoraIni As String, dHoraFin As String, Optional fConsideraDescanso As Boolean = False) As String
    '===================================================================================================
    'Creado : 30/10/08
    'Propósito: Restar horas tanto de un dia o en termino de dia e inicio de otro dia
    '
    'Entradas:  dHoraIni = Hora inicial
    '           dHoraIni = Hora final
    '           fConsideraDescanso = false=no se considera descanso en el dia, true=se resta una hora de
    '           descanso
    '
    'Resultados: hora
    '===================================================================================================

    Dim dHora As Date
    On Error GoTo ERROR

    If IsDate(dHoraIni) = False Or IsDate(dHoraFin) = False Then
        DiferenciaHoras = ""
        Exit Function
    End If

    If CDate(dHoraIni) < CDate(dHoraFin) Then
        If fConsideraDescanso = True Then
            If CDate(dHoraIni) < CDate("13:15:00") And CDate(dHoraFin) > CDate("14:00:00") Then

                DiferenciaHoras = CDate(CDate(CDate(dHoraFin) - CDate("01:00:00")) - CDate(dHoraIni))
            Else

                DiferenciaHoras = CDate(CDate(dHoraFin) - CDate(dHoraIni))
            End If
        Else
            DiferenciaHoras = CDate(CDate(dHoraFin) - CDate(dHoraIni))
        End If
    Else
        DiferenciaHoras = CDate((TimeValue(dHoraFin) - TimeValue("00:00")) + (TimeValue("23:59") - TimeValue(dHoraIni)) + CDate("00:01"))
    End If
    Exit Function
ERROR:
    DiferenciaHoras = ""
    Err.Clear
End Function



Public Function Convert1HoraFaccion(dHora As String) As Double
    '===================================================================================================
    'Creado : 31/10/08
    'Propósito: convertir una hora en especial a faccion. ejm 03:30 >> 3.5
    '
    'Entradas:  dHoraIni = Hora inicial
    '           dHoraIni = Hora final
    '
    'Resultados: hora convertida a fraccion, 0 = caso contrario sea error o falte alguna hora
    '===================================================================================================

    '--esta funcion devuelve los segundos segun la hora como parametro
    Dim mHora As Double
    Dim mMin As Double
    Dim mSeg As Double
    Dim nTmp As String

    If IsDate(dHora) = False Then
        Convert1HoraFaccion = 0
        Exit Function
    End If

    mHora = Hour(dHora)
    mMin = Minute(dHora)
    mMin = IIf(mMin = 0, 0, mMin / 60)
    '--pendiente de implementar los segundos

    Convert1HoraFaccion = mHora + mMin


End Function

Public Function GenerarAsiento(xConn As ADODB.Connection, IdLibro As Integer, IdMovimiento As Double, AnnoTrabajo As String, IdMes As Integer, Optional DestinoAuto As Integer = 1, Optional TipoMovimiento As Integer) As String
    '===================================================================================================
    'Creado : 20/10/08 Por: Johan Castro
    'Propósito: Corregir los Asientos de los distintos libros, asi como los destinos de las cuentas
    '
    'Entradas:  xConn=Conexion activa a la base de datos, antes xcon
    '           IdLibro= Codigo de Libro contable
    '           IdMovimiento= Codigo del Registro que va generar su asiento(compras, ventas, etc)
    '           AnoTrabajo=Año de trabajo
    '           IdMes= Mes de trabajo
    '           DestinoAuto=Indica como se va generar los asientos de destino
    '                       0:No genera Destinos asi las cuetas usadas tengan destino en su configuracion,
    '                       1:Detallado;
    '                       2:Resumen:
    '           TipoMovimiento=SOLO BANCOS Indica si es 1=Ingreso; 2=Egreso
    '
    'Resultados:Asiento Contable corregido

    'Modificado 31/07/09 Por: Johan Castro
    '           Escribir las aperturas de libro Venta, LGD/LGC ' FALTA HACER PRUEBAS
    'Modificado 22/10/09 Por: Johan Castro
    '           Modificar el registro de asiento para banco, inicialmente se utilizaba una consulta que relacionaba las tablas de los modulos que
    '           que entran a banco, ahora se hara consultas por modulo para facilitar el mantenimiento del mismo
    'Modificado 15/09/10 Por:Johan Castro
    '           LGD:: Escribir en asiento del detalle el campo ridlib = 999 esto es cuando los documentos sean reembolsables
    '           Inicialmente se escribia con valor = 0
    'Modificado 25/01/11 Johan Castro
    '           Modificar la escritura del asiento en lgd modulo = 41
    '           1.- Cabecera: Antes vta_gastodebito.id as iddoc ahora 0 as iddoc
    '           2.- Detalle: Antes 0 as iddoc| ahora vta_gastodebitodet.iddoc as iddoc
    '           3.- Detalle: Escribir en campo rregistro vacio cuando el registro sea reembolsable
    '           4.- Escribir en destino de egreso en rregistro vacio cuande sera reembolsable
    'Modificado 28/01/11 Johan Castro
    '           Agrega variable xNumRegistro, indica el numero de registro (mes + cod sunat + asiento)
    'Modificado 25/03/11 Johan Castro
    '           Quitar grupo a consulta de detalle de compras y detalle de ventas, su finalidad es mostrar el asiento tal como se detalle en el módulo
    '           Agregar lineas de codigo para hacer el asiento de igv2 e igv3
    'Modificado 30/03/11 Johan Castro
    '           Considerar las cuentas de anticipos para 2011, dado que estos datos son fijos en el codigo
    '           Cambiar nombre de parametro xcon a xconn por que hay incompatibilidad con variable publica de proyecto con el parametro xcon
    '           se procede a renombrar todas las lineas de codigo que usen parametro xcon a xconn
    'Queda pendiente cuando es no domiciliado
    '           02/05/11 Johan Castro
    '           Modificar consulta de "detalle de concepto Cta haber" para considerar la cuenta asignada al regimen de pensiones
    'Modificado 20/10/11 Johan Castro
    '           Libro Letras, agregar linea para escribir asiento de facturacion, se usa tabla let_letradoc
    '           Agregar linea para escribrir asiento de la nota de debito solo cuando no se agregue manualmente y cuando porte+interes+igv <>0
    'Modificado 23/11/11 Johan Castro
    '           Modificar consultas de libro bancos.
    '           Escribir en diario tipmov = tes_caja.tipmov(1=Ingresos, 2=Egresos), tipo(1=Origen, 2=Destino)
    '           Agregar en origen de egreso el detalle de modulos ventas, lgd, letras
    '           Modificar en asiento automatico tipmov, tipo obtener datos del diario, antes se asignaba valor =0
    '           Modificar en Asiento automatico cuando es resumido(DestinoAuto = 2) corregir el group by
    'Modificado 28/11/11 Johan Castro
    '           Definir variables para la referencias de las operaciones de banco(xrtipdoc2, xrfchope2, xrnumerodoc2, xridtipper2)
    '           Mejorar el select de los detalles de origen y destino de bancos, antes tes_caja.id =xxx,
    '           Ahora Origen(tes_cajaorigendet.idtes=xx) Destino(tes_cajadestinodet.idtes) esto permitira reducir el tiempo al crear el asiento
    '           Modificar campo "iddoc" origen(iddoc = tes_cajaorigendet.iddoc) destino(tes_cajadestinodet.iddoc); "correlativo" origen(iddoc = tes_cajaorigendet.corr) destino(tes_cajadestinodet.corr); iddocpro=0
    '           21/02/12 Johan Castro
    '           Agregar campo ajuste= para guardar si aplica ajuste x dif cambio para todas las consultas de todos los libros
    '           24/02/12
    '           Agregar campo ajuste en "destino ingreso" del módulo lgd
    '===================================================================================================

    Dim dll_varios As New Varias
    Dim nSQL As String
    Dim xAsiento As String
    Dim dFchRegistro As Date
    Dim xAsiento1 As String

    Dim xRst As New ADODB.Recordset
    Dim xNumRegistro As String '--Indica le numero de registro (mes + cod sunat + asiento)

    On Error GoTo ERROR

    Set xCon = xConn

    If IdMes <> 0 Then

        '--definiendo la fecha de registro
        If IdMes <> 0 And IdMes <> 13 Then
            dFchRegistro = CDate("01/" & Format(IdMes, "00") & "/" & AnnoTrabajo)
        End If

        '--obtener el nuevo asiento
        xAsiento = dll_varios.DevuelveNumAsiento(IdLibro, IdMovimiento, IdMes, xConn)

        If xAsiento = "" Then
            xAsiento = dll_varios.NuevoNumAsiento(IdLibro, IdMes, xConn)
        End If

        xAsiento = Format(xAsiento, "0000")

        xAsiento1 = Format(IdMes, "00") + xAsiento '--para las tablas principales


        '--eliminamos los asientos del diario
        xConn.Execute "DELETE * FROM con_diario WHERE año= " & AnnoTrabajo & " and idmes = " & IdMes & " AND idlib = " & IdLibro & " AND idmov = " & IdMovimiento & " ; "

        '---------------------------------------------------------------------------
        '--280909 segun carlos vega ejm. encorp 2008
        '--cuando es cierre no se hace automatico
        If IdMes = 13 Then DestinoAuto = 0
        '---------------------------------------------------------------------------

    Else
        '--cuando es apertura el asiento sera de la sgte. forma

        xAsiento = "0001"
        xAsiento1 = "000001"
        If IdLibro <> 3 Then
            '--eliminamos los asientos del diario
            xConn.Execute "DELETE * FROM con_diario WHERE año= " & AnnoTrabajo & " and idmes = " & IdMes & " AND idlib = 3 AND idmov = " & IdMovimiento & " and iddoc = " & IdMovimiento & "; "
        Else
            '--eliminamos los asientos del diario
            xConn.Execute "DELETE * FROM con_diario WHERE año= " & AnnoTrabajo & " and idmes = " & IdMes & " AND idlib = " & IdLibro & " AND idmov = " & IdMovimiento & " ; "
        End If

    End If

    '***********************************************************************************
    '--numero de registro
    xNumRegistro = Format(IdMes, "00") & Format(Busca_Codigo(IdLibro, "id", "codsun", "mae_libros", "N", xConn), "00") & xAsiento

    '--Si el ejercicio esta vacio se procede a cambiar a algun dato.
    If AnoTra = "" Then CargaDatosEmpresa

    nSQL = ""


    '***********************************************************************************
''''    '--actualizamos el numero de registro
''''    If RstOri.State = 1 Then
''''        RstOri("numreg") = xAsiento1
''''        RstOri.Update
''''    End If
''''    '--acualizamos el numero de registro, y fecha de registro
''''    If RstOrsi.State = 0 Then xConn.Execute "UPDATE com_compras SET com_compras.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((com_compras.id)=" & IdMovimiento & "))"
''''
    '***********************************************************************************


    Select Case IdLibro
        Case 1 '--compras
            '***********************************************************************************
            '--acualizamos el numero de registro, y fecha de registro
            xConn.Execute "UPDATE com_compras SET com_compras.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((com_compras.id)=" & IdMovimiento & "))"
            '***********************************************************************************

            '--total de la compra
            nSQL = "SELECT com_compras.idpro, com_compras.tipdoc, com_compras.numser, com_compras.numdoc, com_compras.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " iif(com_compras.fchreg is null,null,Year(com_compras.fchreg)) AS anotra, iif(com_compras.fchreg is null,null,Month(com_compras.fchreg)) AS idmes, 1 AS idlib, com_compras.id AS idmov, mae_documentocta.idcuen AS idcue, iif(com_compras.numreg is null,null,Right(com_compras.numreg,4)) AS numasi, com_compras.id as iddoc, 0 AS correlativo, IIf(com_compras.tc is null or com_compras.tc=0, con_tc.impven,com_compras.tc) AS tc, " _
                + vbCr + " IIf(mae_documentocta.iddoc=7,IIf(com_compras.idmon=1,com_compras.imptot,0),0) AS impdebsol, " _
                + vbCr + " IIf(mae_documentocta.iddoc=7,IIf(com_compras.idmon=2,com_compras.imptot,0),0) AS impdebdol, " _
                + vbCr + " IIf(mae_documentocta.iddoc<>7,IIf(com_compras.idmon=1,com_compras.imptot,0),0) AS imphabsol, " _
                + vbCr + " IIf(mae_documentocta.iddoc<>7,IIf(com_compras.idmon=2,com_compras.imptot,0),0) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, com_compras.fchreg AS fchasi, com_compras.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,com_compras.idmon, " _
                + vbCr + " 1 AS ridlib, 1 AS ridtipper, com_compras.idpro AS ridper, com_compras.tipdoc AS rtipdoc, com_compras.fchdoc AS rfchope, IIf([com_compras].[numser] Is Null Or [com_compras].[numser]='','',[com_compras].[numser] & '-') & [com_compras].[numdoc] AS rnumerodoc, com_compras.glosa AS rglosa, '' AS rglosaope, Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS rregistro,com_compras.idmon as ridmon, " _
                + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, com_compras.numerodocref AS rnumerodoc1, " _
                + vbCr + " 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, NULL AS rfchope2, '' AS rnumerodoc2, " _
                + vbCr + " IIF(com_compras.tc =0 or com_compras.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (((com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (com_compras.tipdoc = mae_documentocta.iddoc) AND (com_compras.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN mae_docreferencia ON com_compras.idtipdocref = mae_docreferencia.id " _
                + vbCr + " WHERE (((mae_documentocta.tipope)=0)) and com_compras.id = " & IdMovimiento

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

            '--si es apertura salir
            If IdMes = 0 Then GoTo Salto_Apertura:

            '===========================================================================================================================

            '--igv 1 (valor del encabezado - no hay calculo)
            nSQL = " SELECT  com_compras.idpro, com_compras.tipdoc, com_compras.numser, com_compras.numdoc, com_compras.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " iif(com_compras.fchreg is null,null,Year(com_compras.fchreg)) AS anotra, iif(com_compras.fchreg is null,null,Month(com_compras.fchreg)) AS idmes, 1 AS idlib, com_compras.id AS idmov, con_planctas.id AS idcue, iif(com_compras.numreg is null,null,Right(com_compras.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, IIf(com_compras.tc is null or com_compras.tc=0, con_tc.impven,com_compras.tc) AS tc, " _
                + vbCr + " IIf(mae_prov.tipper<>3 and com_compras.tipdoc<>7 and com_compras.tipdoc<>2,IIf(com_compras.idmon=1,com_compras.impigv,0),0) AS impdebsol, " _
                + vbCr + " IIf(mae_prov.tipper<>3 and com_compras.tipdoc<>7 and com_compras.tipdoc<>2,IIf(com_compras.idmon=2,com_compras.impigv,0),0) AS impdebdol, " _
                + vbCr + " IIf( com_compras.tipdoc=7 or com_compras.tipdoc=2 ,IIf(com_compras.idmon=1,com_compras.impigv,0),0) AS imphabsol, " _
                + vbCr + " IIf(com_compras.tipdoc=7 or com_compras.tipdoc=2,IIf(com_compras.idmon=2,com_compras.impigv,0),0) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, com_compras.fchreg AS fchasi, com_compras.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,com_compras.idmon, " _
                + vbCr + " 1 AS ridlib, 1 AS ridtipper, com_compras.idpro AS ridper, com_compras.tipdoc AS rtipdoc, com_compras.fchdoc AS rfchope, IIf([com_compras].[numser] Is Null Or [com_compras].[numser]='','',[com_compras].[numser] & '-') & [com_compras].[numdoc] AS rnumerodoc, com_compras.glosa AS rglosa, '' AS rglosaope,Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS rregistro,com_compras.idmon as ridmon, " _
                + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, com_compras.numerodocref AS rnumerodoc1, " _
                + vbCr + " 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, NULL AS rfchope2, '' AS rnumerodoc2, " _
                + vbCr + " IIF(com_compras.tc =0 or com_compras.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (((mae_documento INNER JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) INNER JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) INNER JOIN (mae_prov INNER JOIN ((com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN mae_docreferencia ON com_compras.idtipdocref = mae_docreferencia.id " _
                + vbCr + " WHERE (((com_compras.impigv)<>0)) and com_compras.id = " & IdMovimiento
            '--igv 2 (valor del encabezado - no hay calculo)
            nSQL = nSQL & vbCr & " UNION " _
                + vbCr + " SELECT  com_compras.idpro, com_compras.tipdoc, com_compras.numser, com_compras.numdoc, com_compras.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " iif(com_compras.fchreg is null,null,Year(com_compras.fchreg)) AS anotra, iif(com_compras.fchreg is null,null,Month(com_compras.fchreg)) AS idmes, 1 AS idlib, com_compras.id AS idmov, con_planctas.id AS idcue, iif(com_compras.numreg is null,null,Right(com_compras.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, IIf(com_compras.tc is null or com_compras.tc=0, con_tc.impven,com_compras.tc) AS tc, " _
                + vbCr + " IIf(mae_prov.tipper<>3 and com_compras.tipdoc<>7 and com_compras.tipdoc<>2,IIf(com_compras.idmon=1,com_compras.impigv2,0),0) AS impdebsol, " _
                + vbCr + " IIf(mae_prov.tipper<>3 and com_compras.tipdoc<>7 and com_compras.tipdoc<>2,IIf(com_compras.idmon=2,com_compras.impigv2,0),0) AS impdebdol, " _
                + vbCr + " IIf( com_compras.tipdoc=7 or com_compras.tipdoc=2 ,IIf(com_compras.idmon=1,com_compras.impigv2,0),0) AS imphabsol, " _
                + vbCr + " IIf(com_compras.tipdoc=7 or com_compras.tipdoc=2,IIf(com_compras.idmon=2,com_compras.impigv2,0),0) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, com_compras.fchreg AS fchasi, com_compras.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,com_compras.idmon, " _
                + vbCr + " 1 AS ridlib, 1 AS ridtipper, com_compras.idpro AS ridper, com_compras.tipdoc AS rtipdoc, com_compras.fchdoc AS rfchope, IIf([com_compras].[numser] Is Null Or [com_compras].[numser]='','',[com_compras].[numser] & '-') & [com_compras].[numdoc] AS rnumerodoc, com_compras.glosa AS rglosa, '' AS rglosaope,Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS rregistro,com_compras.idmon as ridmon, " _
                + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, com_compras.numerodocref AS rnumerodoc1, " _
                + vbCr + " 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, NULL AS rfchope2, '' AS rnumerodoc2, " _
                + vbCr + " IIF(com_compras.tc =0 or com_compras.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (((mae_documento INNER JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) INNER JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) INNER JOIN (mae_prov INNER JOIN ((com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN mae_docreferencia ON com_compras.idtipdocref = mae_docreferencia.id " _
                + vbCr + " WHERE (((com_compras.impigv2)<>0)) and com_compras.id = " & IdMovimiento
            '--igv 3 (valor del encabezado - no hay calculo)
            nSQL = nSQL & vbCr & " UNION " _
                + vbCr + " SELECT  com_compras.idpro, com_compras.tipdoc, com_compras.numser, com_compras.numdoc, com_compras.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " iif(com_compras.fchreg is null,null,Year(com_compras.fchreg)) AS anotra, iif(com_compras.fchreg is null,null,Month(com_compras.fchreg)) AS idmes, 1 AS idlib, com_compras.id AS idmov, con_planctas.id AS idcue, iif(com_compras.numreg is null,null,Right(com_compras.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, IIf(com_compras.tc is null or com_compras.tc=0, con_tc.impven,com_compras.tc) AS tc, " _
                + vbCr + " IIf(mae_prov.tipper<>3 and com_compras.tipdoc<>7 and com_compras.tipdoc<>2,IIf(com_compras.idmon=1,com_compras.impigv3,0),0) AS impdebsol, " _
                + vbCr + " IIf(mae_prov.tipper<>3 and com_compras.tipdoc<>7 and com_compras.tipdoc<>2,IIf(com_compras.idmon=2,com_compras.impigv3,0),0) AS impdebdol, " _
                + vbCr + " IIf( com_compras.tipdoc=7 or com_compras.tipdoc=2 ,IIf(com_compras.idmon=1,com_compras.impigv3,0),0) AS imphabsol, " _
                + vbCr + " IIf(com_compras.tipdoc=7 or com_compras.tipdoc=2,IIf(com_compras.idmon=2,com_compras.impigv3,0),0) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, com_compras.fchreg AS fchasi, com_compras.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,com_compras.idmon, " _
                + vbCr + " 1 AS ridlib, 1 AS ridtipper, com_compras.idpro AS ridper, com_compras.tipdoc AS rtipdoc, com_compras.fchdoc AS rfchope, IIf([com_compras].[numser] Is Null Or [com_compras].[numser]='','',[com_compras].[numser] & '-') & [com_compras].[numdoc] AS rnumerodoc, com_compras.glosa AS rglosa, '' AS rglosaope,Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS rregistro,com_compras.idmon as ridmon, " _
                + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, com_compras.numerodocref AS rnumerodoc1, " _
                + vbCr + " 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, NULL AS rfchope2, '' AS rnumerodoc2, " _
                + vbCr + " IIF(com_compras.tc =0 or com_compras.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (((mae_documento INNER JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) INNER JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) INNER JOIN (mae_prov INNER JOIN ((com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) LEFT JOIN mae_docreferencia ON com_compras.idtipdocref = mae_docreferencia.id " _
                + vbCr + " WHERE (((com_compras.impigv3)<>0)) and com_compras.id = " & IdMovimiento

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================

            '--igv(valor del detalle - calculo de nuevo)
            '--si deseamos obtener el igv del detalle hace lo sgte.
            '--sum(com_comprasdet.canpro * com_comprasdet.preunibru)
''            nSQL = nSQL & vbCr & " Union" _
''                + vbCr + " SELECT com_compras.idpro, com_compras.tipdoc, com_compras.numser, com_compras.numdoc, com_compras.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
''                + vbCr + " IIf(com_compras.fchreg Is Null,Null,Year(com_compras.fchreg)) AS anotra, IIf(com_compras.fchreg Is Null,Null,Month(com_compras.fchreg)) AS idmes, 1 AS idlib, com_compras.id AS idmov, con_planctas.id AS idcue, IIf(com_compras.numreg Is Null,Null,Right(com_compras.numreg,4)) AS numasi, com_compras.tipdoc AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
''                + vbCr + " Sum(IIf(mae_prov.tipper<>3 And com_compras.tipdoc<>7 And com_compras.tipdoc<>2,IIf(com_compras.idmon=1,com_comprasdet.canpro * com_comprasdet.preunibru,com_comprasdet.canpro * com_comprasdet.preunibru*con_tc.impven),0)) * (mae_impuestos.tasa /100) AS impdebsol, " _
''                + vbCr + " Sum(IIf(mae_prov.tipper<>3 And com_compras.tipdoc<>7 And com_compras.tipdoc<>2,IIf(com_compras.idmon=2,com_comprasdet.canpro * com_comprasdet.preunibru,0),0)) * (mae_impuestos.tasa /100) AS impdebdol, " _
''                + vbCr + " Sum(IIf(com_compras.tipdoc=7 Or com_compras.tipdoc=2,IIf(com_compras.idmon=1,com_comprasdet.canpro * com_comprasdet.preunibru,com_comprasdet.canpro * com_comprasdet.preunibru * con_tc.impven),0)) *  (mae_impuestos.tasa /100)  AS imphabsol, " _
''                + vbCr + " Sum(IIf(com_compras.tipdoc=7 Or com_compras.tipdoc=2,IIf(com_compras.idmon=2,com_comprasdet.canpro * com_comprasdet.preunibru,0),0)) * (mae_impuestos.tasa /100) AS imphabdol, " _
''                + vbCr + " 0 AS tipmov, 0 AS prodiv, com_compras.fchreg AS fchasi, com_compras.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, com_compras.idmon " _
''                + vbCr + " FROM mae_prov INNER JOIN ((((mae_documento INNER JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) INNER JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) INNER JOIN (com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_documento.id = com_compras.tipdoc) INNER JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_prov.id = com_compras.idpro " _
''                + vbCr + " GROUP BY com_compras.idpro, com_compras.tipdoc, com_compras.numser, com_compras.numdoc, com_compras.numreg, con_planctas.cuenta, con_planctas.descripcion, IIf(com_compras.fchreg Is Null,Null,Year(com_compras.fchreg)), IIf(com_compras.fchreg Is Null,Null,Month(com_compras.fchreg)), com_compras.id, con_planctas.id, IIf(com_compras.numreg Is Null,Null,Right(com_compras.numreg,4)), com_compras.tipdoc,  con_tc.impven, com_compras.fchreg, com_compras.fchdoc,com_compras.idmon, com_comprasdet.preunibru,mae_impuestos.tasa " _
''                + vbCr + " HAVING (((com_compras.id)= " & IdMovimiento & " ) AND ((com_comprasdet.preunibru)<>0) ) "
''
''            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL,DestinoAuto
''            '===========================================================================================================================


            '--detalle de la compra
            nSQL = " SELECT com_compras.idpro, com_compras.tipdoc, com_compras.numser, com_compras.numdoc, com_compras.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " iif(com_compras.fchreg is null,null,Year(com_compras.fchreg)) AS anotra, iif(com_compras.fchreg is null,null,Month(com_compras.fchreg)) AS idmes, 1 AS idlib, com_compras.id AS idmov, con_planctas.id AS idcue, iif(com_compras.numreg is null,null,Right(com_compras.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, IIf(com_compras.tc is null or com_compras.tc=0, con_tc.impven,com_compras.tc) AS tc, " _
                + vbCr + " (IIf(com_compras.tipdoc<>7,IIf([com_compras].[idmon]=1,[com_comprasdet].[imptot],0),0)) AS impdebsol, " _
                + vbCr + " (IIf(com_compras.tipdoc<>7,IIf([com_compras].[idmon]=2,[com_comprasdet].[imptot],0),0)) AS impdebdol, " _
                + vbCr + " (IIf(com_compras.tipdoc=7,IIf([com_compras].[idmon]=1,[com_comprasdet].[imptot],0),0)) AS imphabsol, " _
                + vbCr + " (IIf(com_compras.tipdoc=7,IIf([com_compras].[idmon]=2,[com_comprasdet].[imptot],0),0)) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, com_compras.fchreg AS fchasi, com_compras.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,com_compras.idmon, " _
                + vbCr + " 1 AS ridlib, 1 AS ridtipper, com_compras.idpro AS ridper, com_compras.tipdoc AS rtipdoc, com_compras.fchdoc AS rfchope, IIf([com_compras].[numser] Is Null Or [com_compras].[numser]='','',[com_compras].[numser] & '-') & [com_compras].[numdoc] AS rnumerodoc, com_compras.glosa AS rglosa,'' AS rglosaope, Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS rregistro,com_compras.idmon as ridmon, " _
                + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, com_compras.numerodocref AS rnumerodoc1, " _
                + vbCr + " 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, NULL AS rfchope2, '' AS rnumerodoc2, " _
                + vbCr + " IIF(com_compras.tc =0 or com_compras.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                + vbCr + " FROM ((alm_inventario LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id) INNER JOIN ((((com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN mae_docreferencia ON com_compras.idtipdocref = mae_docreferencia.id) INNER JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN mae_docreferencia AS mae_docreferencia_1 ON com_compras.idtipdocref = mae_docreferencia_1.id " _
                + vbCr + " WHERE com_compras.id = " & IdMovimiento & ""


            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================

            '--isc detalle
            nSQL = " SELECT com_compras.idpro, com_compras.tipdoc, com_compras.numser, com_compras.numdoc, com_compras.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " iif(com_compras.fchreg is null,null,Year(com_compras.fchreg)) AS anotra, iif(com_compras.fchreg is null,null,Month(com_compras.fchreg)) AS idmes, 1 AS idlib, com_compras.id AS idmov, con_planctas.ctadesdeb AS idcue, iif(com_compras.numreg is null,null,Right(com_compras.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, IIf(com_compras.tc is null or com_compras.tc=0, con_tc.impven,com_compras.tc) AS tc, " _
                + vbCr + " (IIf(com_compras.tipdoc<>7,IIf([com_compras].[idmon]=1,[com_comprasdet].[imptot],0),0)) AS impdebsol, " _
                + vbCr + " (IIf(com_compras.tipdoc<>7,IIf([com_compras].[idmon]=2,[com_comprasdet].[imptot],0),0)) AS impdebdol, " _
                + vbCr + " (IIf(com_compras.tipdoc=7,IIf([com_compras].[idmon]=1,[com_comprasdet].[imptot],0),0)) AS imphabsol, " _
                + vbCr + " (IIf(com_compras.tipdoc=7,IIf([com_compras].[idmon]=2,[com_comprasdet].[imptot],0),0)) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, com_compras.fchreg AS fchasi, com_compras.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,com_compras.idmon, " _
                + vbCr + " 1 AS ridlib, 1 AS ridtipper, com_compras.idpro AS ridper, com_compras.tipdoc AS rtipdoc, com_compras.fchdoc AS rfchope, IIf([com_compras].[numser] Is Null Or [com_compras].[numser]='','',[com_compras].[numser] & '-') & [com_compras].[numdoc] AS rnumerodoc, com_compras.glosa AS rglosa,'' AS rglosaope, Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS rregistro,com_compras.idmon as ridmon, " _
                + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, com_compras.numerodocref AS rnumerodoc1, " _
                + vbCr + " 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, NULL AS rfchope2, '' AS rnumerodoc2, " _
                + vbCr + " IIF(com_compras.tc =0 or com_compras.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (((alm_inventario INNER JOIN mae_impuestos ON alm_inventario.idimpsel = mae_impuestos.id) INNER JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) INNER JOIN (((com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) INNER JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON alm_inventario.id = com_comprasdet.iditem) LEFT JOIN mae_docreferencia ON com_compras.idtipdocref = mae_docreferencia.id " _
                + vbCr + " WHERE (((con_planctas.ctadesdeb)<>0) AND ((alm_inventario.idimpsel)<>0)) and com_compras.id = " & IdMovimiento & ""

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================

        Case 2 '--ventas

            '***********************************************************************************
            '--acualizamos el numero de registro, y fecha de registro
            xConn.Execute "UPDATE vta_ventas SET vta_ventas.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((vta_ventas.id)=" & IdMovimiento & "))"
            '***********************************************************************************

            '--total de la venta
            nSQL = "SELECT vta_ventas.idcli, vta_ventas.tipdoc,vta_ventas.numser, vta_ventas.numdoc, vta_ventas.idmon, vta_ventas.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " iif(vta_ventas.fchreg is null,null,Year(vta_ventas.fchreg)) AS anotra, iif(vta_ventas.fchreg is null,null,Month(vta_ventas.fchreg)) AS idmes, 2 AS idlib, vta_ventas.id AS idmov, con_planctas.id AS idcue, " _
                + vbCr + " iif(vta_ventas.numreg is null,null,Right(vta_ventas.numreg,4)) AS numasi, vta_ventas.id as iddoc, 0 AS correlativo, iif(vta_ventas.tc =0,con_tc.impven,vta_ventas.tc) AS tc, " _
                + vbCr + " IIf(mae_documentocta.iddoc<>7,IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc,0),0) AS impdebsol, " _
                + vbCr + " IIf(mae_documentocta.iddoc<>7,IIf(vta_ventas.idmon=2,vta_ventas.imptotdoc,0),0) AS impdebdol, " _
                + vbCr + " IIf(mae_documentocta.iddoc=7,IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc,0),0) AS imphabsol, " _
                + vbCr + " IIf(mae_documentocta.iddoc=7,IIf(vta_ventas.idmon=2,vta_ventas.imptotdoc,0),0) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, vta_ventas.fchreg AS fchasi, vta_ventas.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, " _
                + vbCr + " 2 AS ridlib, 2 AS ridtipper, vta_ventas.idcli AS ridper, vta_ventas.tipdoc AS rtipdoc, vta_ventas.fchdoc AS rfchope, IIf([vta_ventas].[numser] Is Null Or [vta_ventas].[numser]='','',[vta_ventas].[numser] & '-') & [vta_ventas].[numdoc] AS rnumerodoc, vta_ventas.glosa AS rglosa,'' as rglosaope, Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) AS rregistro,vta_ventas.idmon as ridmon, " _
                + vbCr + " vta_ventas.idtipdocref AS rtipdoc1, vta_ventas.numerodocref AS rnumerodoc1, IIf([vta_ventas_1].[id] Is Null,0,2) AS ridtipper2, IIf([vta_ventas_1].[id] Is Null,0,[vta_ventas_1].[idcli]) AS ridper2, IIf([vta_ventas_1].[id] Is Null,0,[vta_ventas_1].[tipdoc]) AS rtipdoc2, IIf([vta_ventas_1].[id] Is Null,0,[vta_ventas_1].[fchdoc]) AS rfchope2, IIf([vta_ventas_1].[id] Is Null,Null,[vta_ventas_1].[numser] & '-' & [vta_ventas_1].[numdoc]) AS rnumerodoc2, " _
                + vbCr + " IIF(vta_ventas.tc =0 or vta_ventas.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (((vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id " _
                + vbCr + " WHERE (((mae_documentocta.tipope) = -1)) and vta_ventas.id = " & IdMovimiento

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

            '--si es apertura salir
            If IdMes = 0 Then GoTo Salto_Apertura:

            '===========================================================================================================================

            '--igv de la venta
            nSQL = " SELECT vta_ventas.idcli, vta_ventas.tipdoc,vta_ventas.numser, vta_ventas.numdoc, vta_ventas.idmon, vta_ventas.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " iif(vta_ventas.fchreg is null,null,Year(vta_ventas.fchreg)) AS anotra, iif(vta_ventas.fchreg is null,null,Month(vta_ventas.fchreg)) AS idmes, 2 AS idlib, vta_ventas.id AS idmov, con_planctas.id AS idcue, " _
                + vbCr + " iif(vta_ventas.numreg is null,null,Right(vta_ventas.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, iif(vta_ventas.tc =0,con_tc.impven,vta_ventas.tc) AS tc, " _
                + vbCr + " IIf(vta_ventas.tipdoc=7,IIf(vta_ventas.idmon=1,vta_ventas.impigv,0),0) AS impdebsol, " _
                + vbCr + " IIf(vta_ventas.tipdoc=7,IIf(vta_ventas.idmon=2,vta_ventas.impigv,0),0) AS impdebdol, " _
                + vbCr + " IIf(vta_ventas.tipdoc<>7,IIf(vta_ventas.idmon=1,vta_ventas.impigv,0),0) AS imphabsol, " _
                + vbCr + " IIf(vta_ventas.tipdoc<>7,IIf(vta_ventas.idmon=2,vta_ventas.impigv,0),0) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, vta_ventas.fchreg AS fchasi, vta_ventas.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, " _
                + vbCr + " 2 AS ridlib, 2 AS ridtipper, vta_ventas.idcli AS ridper, vta_ventas.tipdoc AS rtipdoc, vta_ventas.fchdoc AS rfchope, IIf([vta_ventas].[numser] Is Null Or [vta_ventas].[numser]='','',[vta_ventas].[numser] & '-') & [vta_ventas].[numdoc] AS rnumerodoc, vta_ventas.glosa AS rglosa,'' as rglosaope, Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) AS rregistro,vta_ventas.idmon as ridmon, " _
                + vbCr + " vta_ventas.idtipdocref AS rtipdoc1, vta_ventas.numerodocref AS rnumerodoc1, IIf([vta_ventas_1].[id] Is Null,0,2) AS ridtipper2, IIf([vta_ventas_1].[id] Is Null,0,[vta_ventas_1].[idcli]) AS ridper2, IIf([vta_ventas_1].[id] Is Null,0,[vta_ventas_1].[tipdoc]) AS rtipdoc2, IIf([vta_ventas_1].[id] Is Null,0,[vta_ventas_1].[fchdoc]) AS rfchope2, IIf([vta_ventas_1].[id] Is Null,Null,[vta_ventas_1].[numser] & '-' & [vta_ventas_1].[numdoc]) AS rnumerodoc2, " _
                + vbCr + " IIF(vta_ventas.tc =0 or vta_ventas.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (((((vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) INNER JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) INNER JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) INNER JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id " _
                + vbCr + " WHERE vta_ventas.id = " & IdMovimiento & " "

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================

            '--detalle de la venta
            nSQL = " SELECT vta_ventas.idcli, vta_ventas.tipdoc,vta_ventas.numser, vta_ventas.numdoc, vta_ventas.idmon, vta_ventas.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " iif(vta_ventas.fchreg is null,null,Year(vta_ventas.fchreg)) AS anotra, iif(vta_ventas.fchreg is null,null,Month(vta_ventas.fchreg)) AS idmes, 2 AS idlib, vta_ventas.id AS idmov, con_planctas.id AS idcue, " _
                + vbCr + " iif(vta_ventas.numreg is null,null,Right(vta_ventas.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, iif(vta_ventas.tc =0,con_tc.impven,vta_ventas.tc) AS tc, " _
                + vbCr + " (IIf(vta_ventas.tipdoc=7,IIf(vta_ventas.idmon=1,vta_ventasdet.imptot,0),0)) AS impdebsol, " _
                + vbCr + " (IIf(vta_ventas.tipdoc=7,IIf(vta_ventas.idmon=2,vta_ventasdet.imptot,0),0)) AS impdebdol, " _
                + vbCr + " (IIf(vta_ventas.tipdoc<>7,IIf(vta_ventas.idmon=1,vta_ventasdet.imptot,0),0)) AS imphabsol, " _
                + vbCr + " (IIf(vta_ventas.tipdoc<>7,IIf(vta_ventas.idmon=2,vta_ventasdet.imptot,0),0)) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, vta_ventas.fchreg AS fchasi, vta_ventas.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, " _
                + vbCr + " 2 AS ridlib, 2 AS ridtipper, vta_ventas.idcli AS ridper, vta_ventas.tipdoc AS rtipdoc, vta_ventas.fchdoc AS rfchope, IIf(vta_ventas.numser Is Null Or vta_ventas.numser='','',vta_ventas.numser & '-') & vta_ventas.numdoc AS rnumerodoc, vta_ventas.glosa AS rglosa,'' as rglosaope, Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4) AS rregistro,vta_ventas.idmon as ridmon, " _
                + vbCr + " vta_ventas.idtipdocref AS rtipdoc1, vta_ventas.numerodocref AS rnumerodoc1, IIf(vta_ventas_1.id Is Null,0,2) AS ridtipper2, IIf(vta_ventas_1.id Is Null,0,vta_ventas_1.idcli) AS ridper2, IIf(vta_ventas_1.id Is Null,0,vta_ventas_1.tipdoc) AS rtipdoc2, IIf(vta_ventas_1.id Is Null,0,vta_ventas_1.fchdoc) AS rfchope2, IIf(vta_ventas_1.id Is Null,Null,vta_ventas_1.numser & '-' & vta_ventas_1.numdoc) AS rnumerodoc2, " _
                + vbCr + " IIF(vta_ventas.tc =0 or vta_ventas.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (((vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) INNER JOIN (vta_ventasdet INNER JOIN (alm_inventario LEFT JOIN con_planctas ON alm_inventario.idcuentaven = con_planctas.id) ON vta_ventasdet.iditem = alm_inventario.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id " _
                + vbCr + " WHERE vta_ventasdet.imptot <> 0 and vta_ventas.id = " & IdMovimiento & ""

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================


        Case 3 '--provisiones diversos


                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE con_proviciones SET con_proviciones.numreg = '" & Format(IdMes, "00") & xAsiento & "' WHERE (((con_proviciones.id)=" & IdMovimiento & "))"
                '***********************************************************************************

                nSQL = "SELECT con_proviciones.idcli, con_proviciones.tipdoc, con_proviciones.numser, con_proviciones.numdoc, con_proviciones.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " IIf(con_proviciones.fchreg Is Null,Null,Year(con_proviciones.fchreg)) AS anotra, IIf(con_proviciones.fchreg Is Null,Null,Month(con_proviciones.fchreg)) AS idmes, 2 AS idlib, con_proviciones.id AS idmov, con_planctas.id AS idcue, " _
                + vbCr + " IIf(con_proviciones.numreg Is Null,Null,Right(con_proviciones.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, con_proviciones.tc , " _
                + vbCr + " IIf(con_provicionesdet.tipo=0,IIf(con_proviciones.idmon=1,con_provicionesdet.imp,0),0) AS impdebsol, " _
                + vbCr + " IIf(con_provicionesdet.tipo=0,IIf(con_proviciones.idmon=2,con_provicionesdet.imp,0),0) AS impdebdol, " _
                + vbCr + " IIf(con_provicionesdet.tipo=-1,IIf(con_proviciones.idmon=1,con_provicionesdet.imp,0),0) AS imphabsol, " _
                + vbCr + " IIf(con_provicionesdet.tipo=-1,IIf(con_proviciones.idmon=2,con_provicionesdet.imp,0),0) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, -1 AS prodiv, con_proviciones.fchreg AS fchasi, con_proviciones.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, con_proviciones.idmon, " _
                + vbCr + " 3 AS ridlib, 0 AS ridtipper, 0 AS ridper, con_proviciones.tipdoc AS rtipdoc, con_proviciones.fchdoc AS rfchope, IIf([con_proviciones].[numser] Is Null Or [con_proviciones].[numser]='','',[con_proviciones].[numser] & '-') & [con_proviciones].[numdoc] AS rnumerodoc, '' as rglosa ,con_proviciones.glosa AS rglosaope, Left([con_proviciones].[numreg],2) & [mae_libros].[codsun] & Right([con_proviciones].[numreg],4) AS rregistro,con_proviciones.idmon as ridmon, " _
                + vbCr + " 0 as rtipdoc1, '' as rnumerodoc1, 0 as ridtipper2,0 as ridper2, 0 as rtipdoc2, null as rfchope2, '' as rnumerodoc2 " _
                + vbCr + " ,-1 as aplicatc, iif(con_proviciones.ajuste is null,0,con_proviciones.ajuste) as ajuste " _
                + vbCr + " FROM mae_libros RIGHT JOIN ((con_proviciones LEFT JOIN con_tc ON con_proviciones.fchdoc = con_tc.fecha) INNER JOIN (con_provicionesdet INNER JOIN con_planctas ON con_provicionesdet.idcuen = con_planctas.id) ON con_proviciones.id = con_provicionesdet.id) ON mae_libros.id = con_proviciones.idlib " _
                + vbCr + " WHERE (((con_proviciones.id)=" & IdMovimiento & " )) "

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================

        Case 4 '--Percepciones

                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE con_percepcion SET con_percepcion.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((con_percepcion.id)=" & IdMovimiento & "))"
                '***********************************************************************************
            '--PERCEPCION VENTA
            '--total de la percepcion
            nSQL = "SELECT con_percepcion.idcli as idpro, con_percepcion.tipdoc , con_percepcion.numser, con_percepcion.numdoc, con_percepcion.idmon as idmonori, con_percepcion.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " Year(con_percepcion.fchreg) AS anotra, Month(con_percepcion.fchreg) AS idmes, 4 AS idlib, con_percepcion.id AS idmov, con_planctas.id AS idcue,iif(con_percepcion.numreg is null,'', Right(con_percepcion.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                + vbCr + " 0 AS impdebsol, " _
                + vbCr + " 0 AS impdebdol, " _
                + vbCr + " IIf(con_percepcion.idmon=1,con_percepcion.imptotper,0) AS imphabsol, " _
                + vbCr + " IIf(con_percepcion.idmon=2,con_percepcion.imptotper,0) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, con_percepcion.fchreg AS fchasi, con_percepcion.fchdoc , 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,con_percepcion.idmon, " _
                + vbCr + " 4 AS ridlib, 2 AS ridtipper, con_percepcion.idcli  AS ridper, con_percepcion.tipdoc AS rtipdoc, con_percepcion.fchdoc AS rfchope, IIf([con_percepcion]![numser] Is Null Or [con_percepcion]![numser]='','',[con_percepcion]![numser] & '-') & [con_percepcion]![numdoc] AS rnumerodoc, Left([con_percepcion].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([con_percepcion].[numreg],4) AS rregistro, '' as rglosa,con_percepcion.glosa AS rglosaope,con_percepcion.idmon as ridmon, " _
                + vbCr + " NULL as rtipdoc1, NULL as rnumerodoc1, NULL as ridtipper2,NULL as ridper2, NULL as rtipdoc2, NULL as rfchope2, NULL as rnumerodoc2 " _
                + vbCr + " ,0 as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (mae_percepcion LEFT JOIN con_planctas ON mae_percepcion.idcuenven = con_planctas.id) RIGHT JOIN ((con_percepcion LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id) ON mae_percepcion.id = con_percepcion.idper " _
                + vbCr + " WHERE con_percepcion.tipo=2 and con_percepcion.id = " & IdMovimiento

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

            '--detalle de la percepcion

            nSQL = " SELECT vta_ventas.idcli AS idpro, vta_ventas.tipdoc, vta_ventas.numser, vta_ventas.numdoc, vta_ventas.idmon as idmonori, vta_ventas.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " Year(con_percepcion.fchreg) AS anotra, Month(con_percepcion.fchreg) AS idmes, 4 AS idlib, con_percepcion.id AS idmov, con_planctas.id AS idcue, iif(con_percepcion.numreg is null,'', Right(con_percepcion.numreg,4)) AS numasi, vta_ventas.id AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                + vbCr + " IIf(vta_ventas.tipdoc=7,0,IIf(con_percepcion.idmon=1,con_percepciondet.impper,0)) AS impdebsol, " _
                + vbCr + " IIf(vta_ventas.tipdoc=7,0,IIf(con_percepcion.idmon=2,con_percepciondet.impper,0)) AS impdebdol, " _
                + vbCr + " Abs(IIf(vta_ventas.tipdoc<>7,0,IIf(con_percepcion.idmon=1,con_percepciondet.impper,IIf(con_tc.impven Is Null,0,con_percepciondet.impper*con_tc.impven)))) AS imphabsol, " _
                + vbCr + " Abs(IIf(vta_ventas.tipdoc<>7,0,IIf(con_percepcion.idmon=2,con_percepciondet.impper,0))) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, con_percepcion.fchreg AS fchasi, con_percepcion.fchdoc , 0 AS idorides, 0 AS idmod, con_percepciondet.iddoc AS iddocpro, 0 AS tipo,con_percepcion.idmon, " _
                + vbCr + " 2 AS ridlib, 2 AS ridtipper, vta_ventas.idcli AS ridper, vta_ventas.tipdoc AS rtipdoc, vta_ventas.fchdoc AS rfchope, IIf([vta_ventas]![numser] Is Null Or [vta_ventas]![numser]='','',[vta_ventas]![numser] & '-') & [vta_ventas]![numdoc] AS rnumerodoc, Left([vta_ventas].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([vta_ventas].[numreg],4) AS rregistro, vta_ventas.glosa AS rglosa,con_percepcion.glosa AS rglosaope,vta_ventas.idmon as ridmon, " _
                + vbCr + " null as rtipdoc1, null as rnumerodoc1, 2 as ridtipper2,con_percepcion.idcli  as ridper2, con_percepcion.tipdoc as rtipdoc2, con_percepcion.fchdoc as rfchope2, IIf([con_percepcion]![numser] Is Null Or [con_percepcion]![numser]='','',[con_percepcion]![numser] & '-') & [con_percepcion]![numdoc] as rnumerodoc2 " _
                + vbCr + " ,0 as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (con_percepcion LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha) INNER JOIN (((con_percepciondet INNER JOIN vta_ventas ON con_percepciondet.iddoc = vta_ventas.id) INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) ON con_percepcion.id = con_percepciondet.id " _
                + vbCr + " WHERE con_percepcion.tipo=2 AND mae_documentocta.tipope=-1 and con_percepcion.id = " & IdMovimiento

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

            '===========================================================================================================================


            '===========================================================================================================================

            '--total de la percepcion-compra
            nSQL = " SELECT con_percepcion.idcli as idpro, con_percepcion.tipdoc , con_percepcion.numser, con_percepcion.numdoc, con_percepcion.idmon as idmonori, con_percepcion.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " Year(con_percepcion.fchreg) AS anotra, Month(con_percepcion.fchreg) AS idmes, 4 AS idlib, con_percepcion.id AS idmov, con_planctas.id AS idcue, iif(con_percepcion.numreg is null,'', Right(con_percepcion.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                + vbCr + " IIf(con_percepcion.idmon=1,con_percepcion.imptotper,0) AS impdebsol, " _
                + vbCr + " IIf(con_percepcion.idmon=2,con_percepcion.imptotper,0) AS impdebdol, " _
                + vbCr + " 0 AS imphabsol, " _
                + vbCr + " 0 AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, con_percepcion.fchreg AS fchasi, con_percepcion.fchdoc , 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,con_percepcion.idmon, " _
                + vbCr + " 4 AS ridlib, 1 AS ridtipper, con_percepcion.idcli  AS ridper, con_percepcion.tipdoc AS rtipdoc, con_percepcion.fchdoc AS rfchope, IIf([con_percepcion]![numser] Is Null Or [con_percepcion]![numser]='','',[con_percepcion]![numser] & '-') & [con_percepcion]![numdoc] AS rnumerodoc, Left([con_percepcion].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([con_percepcion].[numreg],4) AS rregistro, '' as rglosa,con_percepcion.glosa AS rglosaope,con_percepcion.idmon as ridmon, " _
                + vbCr + " null as rtipdoc1, null as rnumerodoc1, null as ridtipper2,null as ridper2, null as rtipdoc2, null as rfchope2, null as rnumerodoc2 " _
                + vbCr + " ,0 as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (mae_percepcion LEFT JOIN con_planctas ON mae_percepcion.idcuencom = con_planctas.id) RIGHT JOIN ((con_percepcion LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id) ON mae_percepcion.id = con_percepcion.idper " _
                + vbCr + " WHERE con_percepcion.tipo=1 and con_percepcion.id = " & IdMovimiento

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================


            '--detalle de la percepcion-compra

            nSQL = " SELECT com_compras.idpro AS idpro, com_compras.tipdoc, com_compras.numser, com_compras.numdoc, com_compras.idmon as idmonori, com_compras.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " Year(con_percepcion.fchreg) AS anotra, Month(con_percepcion.fchreg) AS idmes, 4 AS idlib, con_percepcion.id AS idmov, con_planctas.id AS idcue, iif(con_percepcion.numreg is null,'', Right(con_percepcion.numreg,4)) AS numasi, com_compras.id AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                + vbCr + " Abs(IIf(com_compras.tipdoc<>7,0,IIf(con_percepcion.idmon=1,con_percepciondet.impper,0))) AS impdebsol, " _
                + vbCr + " Abs(IIf(com_compras.tipdoc<>7,0,IIf(con_percepcion.idmon=2,con_percepciondet.impper,0))) AS impdebdol, " _
                + vbCr + " IIf(com_compras.tipdoc=7,0,IIf(con_percepcion.idmon=1,con_percepciondet.impper,IIf(con_tc.impven Is Null,0,con_percepciondet.impper*con_tc.impven))) AS imphabsol, " _
                + vbCr + " IIf(com_compras.tipdoc=7,0,IIf(con_percepcion.idmon=2,con_percepciondet.impper,0)) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, con_percepcion.fchreg AS fchasi, con_percepcion.fchdoc , 0 AS idorides, 0 AS idmod, con_percepciondet.iddoc AS iddocpro, 0 AS tipo,con_percepcion.idmon, " _
                + vbCr + " 1 AS ridlib, 1 AS ridtipper, com_compras.idpro AS ridper, com_compras.tipdoc AS rtipdoc, com_compras.fchdoc AS rfchope, IIf([com_compras]![numser] Is Null Or [com_compras]![numser]='','',[com_compras]![numser] & '-') & [com_compras]![numdoc] AS rnumerodoc, Left([com_compras].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([com_compras].[numreg],4) AS rregistro, com_compras.glosa AS rglosa,con_percepcion.glosa AS rglosaope,com_compras.idmon as ridmon, " _
                + vbCr + " null as rtipdoc1, null as rnumerodoc1, 1 as ridtipper2,con_percepcion.idcli  as ridper2, con_percepcion.tipdoc as rtipdoc2, con_percepcion.fchdoc as rfchope2, IIf([con_percepcion]![numser] Is Null Or [con_percepcion]![numser]='','',[con_percepcion]![numser] & '-') & [con_percepcion]![numdoc] as rnumerodoc2 " _
                + vbCr + " ,0 as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (con_percepcion LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha) INNER JOIN (((con_percepciondet INNER JOIN com_compras ON con_percepciondet.iddoc = com_compras.id) INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (com_compras.tipdoc = mae_documentocta.iddoc) AND (com_compras.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON con_percepcion.id = con_percepciondet.id " _
                + vbCr + " WHERE con_percepcion.tipo=1 AND mae_documentocta.tipope=0 and con_percepcion.id = " & IdMovimiento

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================




        Case 5 '--retenciones

                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE con_retencion SET con_retencion.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((con_retencion.id)=" & IdMovimiento & "))"
                '***********************************************************************************

            '--total de la retencion -venta
            nSQL = "SELECT con_retencion.idpro, con_retencion.iddoc AS tipdoc, con_retencion.numser, con_retencion.numdoc, con_retencion.idmon as idmonori, con_retencion.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " Year(con_retencion.fchreg) AS anotra, Month(con_retencion.fchreg) AS idmes, 5 AS idlib, con_retencion.id AS idmov, con_planctas.id AS idcue,iif(con_retencion.numreg is null,'', Right(con_retencion.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                + vbCr + " IIf(con_retencion.idmon=1,con_retencion.imp,IIf(con_tc.impven Is Null,0,con_retencion.imp*con_tc.impven)) AS impdebsol, " _
                + vbCr + " IIf(con_retencion.idmon=2,con_retencion.imp,0) AS impdebdol, " _
                + vbCr + " 0 AS imphabsol, " _
                + vbCr + " 0 AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, con_retencion.fchreg AS fchasi, con_retencion.fchemi AS fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,con_retencion.idmon, " _
                + vbCr + " 5 AS ridlib, 2 AS ridtipper, con_retencion.idpro AS ridper, con_retencion.iddoc AS rtipdoc, con_retencion.fchemi AS rfchope, IIf([con_retencion]![numser] Is Null Or [con_retencion]![numser]='','',[con_retencion]![numser] & '-') & [con_retencion]![numdoc] AS rnumerodoc, Left([con_retencion].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([con_retencion].[numreg],4) AS rregistro, '' as rglosa,con_retencion.glosa AS rglosaope,con_retencion.idmon as ridmon, " _
                + vbCr + " null as rtipdoc1, null as rnumerodoc1, null as ridtipper2,null as ridper2, null as rtipdoc2, null as rfchope2, null as rnumerodoc2 " _
                + vbCr + " ,0 as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (mae_retencion LEFT JOIN con_planctas ON mae_retencion.idcuenven = con_planctas.id) RIGHT JOIN ((con_retencion LEFT JOIN con_tc ON con_retencion.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) ON mae_retencion.id = con_retencion.idret " _
                + vbCr + " WHERE con_retencion.tipo=2 and con_retencion.id = " & IdMovimiento

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================

            '--total de la retencion -compra
            nSQL = " SELECT con_retencion.idpro, con_retencion.iddoc AS tipdoc, con_retencion.numser, con_retencion.numdoc, con_retencion.idmon as idmonori, con_retencion.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " Year(con_retencion.fchreg) AS anotra, Month(con_retencion.fchreg) AS idmes, 5 AS idlib, con_retencion.id AS idmov, con_planctas.id AS idcue, iif(con_retencion.numreg is null,'', Right(con_retencion.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                + vbCr + " 0 AS impdebsol, " _
                + vbCr + " 0 AS impdebdol, " _
                + vbCr + " IIf(con_retencion.idmon=1,con_retencion.imp,IIf(con_tc.impven Is Null,0,con_retencion.imp*con_tc.impven)) AS imphabsol, " _
                + vbCr + " IIf(con_retencion.idmon=2,con_retencion.imp,0) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, con_retencion.fchreg AS fchasi, con_retencion.fchemi AS fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,con_retencion.idmon, " _
                + vbCr + " 5 AS ridlib, 1 AS ridtipper, con_retencion.idpro AS ridper, con_retencion.iddoc AS rtipdoc, con_retencion.fchemi AS rfchope, IIf([con_retencion]![numser] Is Null Or [con_retencion]![numser]='','',[con_retencion]![numser] & '-') & [con_retencion]![numdoc] AS rnumerodoc, Left([con_retencion].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([con_retencion].[numreg],4) AS rregistro, '' as rglosa,con_retencion.glosa AS rglosaope,con_retencion.idmon as ridmon, " _
                + vbCr + " null as rtipdoc1, null as rnumerodoc1, null as ridtipper2,null as ridper2, null as rtipdoc2, null as rfchope2, null as rnumerodoc2 " _
                + vbCr + " ,0 as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (mae_retencion LEFT JOIN con_planctas ON mae_retencion.idcuencom = con_planctas.id) RIGHT JOIN ((con_retencion LEFT JOIN con_tc ON con_retencion.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) ON mae_retencion.id = con_retencion.idret " _
                + vbCr + " WHERE con_retencion.tipo=1 and con_retencion.id = " & IdMovimiento

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================

            '--detalle de la retencion -venta

            nSQL = " SELECT vta_ventas.idcli AS idpro, vta_ventas.tipdoc, vta_ventas.numser, vta_ventas.numdoc, vta_ventas.idmon as idmonori, vta_ventas.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " Year(con_retencion.fchreg) AS anotra, Month(con_retencion.fchreg) AS idmes, 5 AS idlib, con_retencion.id AS idmov, con_planctas.id AS idcue, iif(con_retencion.numreg is null,'', Right(con_retencion.numreg,4)) AS numasi, vta_ventas.id AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                + vbCr + " Abs(IIf(vta_ventas.tipdoc<>7,0,IIf(con_retencion.idmon=1,con_retenciondet.impret,IIf(con_tc.impven Is Null,0,con_retenciondet.impret*con_tc.impven)))) AS impdebsol, " _
                + vbCr + " Abs(IIf(vta_ventas.tipdoc<>7,0,IIf(con_retencion.idmon=2,con_retenciondet.impret,0))) AS impdebdol, " _
                + vbCr + " IIf(vta_ventas.tipdoc=7,0,IIf(con_retencion.idmon=1,con_retenciondet.impret,IIf(con_tc.impven Is Null,0,con_retenciondet.impret*con_tc.impven))) AS imphabsol, " _
                + vbCr + " IIf(vta_ventas.tipdoc=7,0,IIf(con_retencion.idmon=2,con_retenciondet.impret,0)) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, con_retencion.fchreg AS fchasi, con_retencion.fchemi AS fchdoc, 0 AS idorides, 0 AS idmod, con_retenciondet.iddoc AS iddocpro, 0 AS tipo,con_retencion.idmon, " _
                + vbCr + " 2 AS ridlib,2 AS ridtipper, vta_ventas.idcli AS ridper, vta_ventas.tipdoc AS rtipdoc, vta_ventas.fchdoc AS rfchope, IIf([vta_ventas]![numser] Is Null Or [vta_ventas]![numser]='','',[vta_ventas]![numser] & '-') & [vta_ventas]![numdoc] AS rnumerodoc, Left([vta_ventas].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([vta_ventas].[numreg],4) AS rregistro, vta_ventas.glosa AS rglosa,con_retencion.glosa AS rglosaope,vta_ventas.idmon as ridmon, " _
                + vbCr + " null as rtipdoc1, null as rnumerodoc1, 2 as ridtipper2,con_retencion.idpro as ridper2, con_retencion.iddoc as rtipdoc2, con_retencion.fchemi as rfchope2, IIf([con_retencion]![numser] Is Null Or [con_retencion]![numser]='','',[con_retencion]![numser] & '-') & [con_retencion]![numdoc] as rnumerodoc2 " _
                + vbCr + " ,0 as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (con_retencion LEFT JOIN con_tc ON con_retencion.fchemi = con_tc.fecha) INNER JOIN (((con_retenciondet INNER JOIN vta_ventas ON con_retenciondet.iddoc = vta_ventas.id) INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) ON con_retencion.id = con_retenciondet.id " _
                + vbCr + " WHERE con_retencion.tipo=2 AND mae_documentocta.tipope=-1 and con_retencion.id = " & IdMovimiento

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================

            '--detalle de la retencion -compra

            nSQL = " SELECT com_compras.idpro AS idpro, com_compras.tipdoc, com_compras.numser, com_compras.numdoc, com_compras.idmon as idmonori, com_compras.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                + vbCr + " Year(con_retencion.fchreg) AS anotra, Month(con_retencion.fchreg) AS idmes, 5 AS idlib, con_retencion.id AS idmov, con_planctas.id AS idcue, iif(con_retencion.numreg is null,'', Right(con_retencion.numreg,4)) AS numasi, com_compras.id AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                + vbCr + " IIf(com_compras.tipdoc=7,0,IIf(con_retencion.idmon=1,con_retenciondet.impret,IIf(con_tc.impven Is Null,0,con_retenciondet.impret*con_tc.impven))) AS impdebsol, " _
                + vbCr + " IIf(com_compras.tipdoc=7,0,IIf(con_retencion.idmon=2,con_retenciondet.impret,0)) AS impdebdol, " _
                + vbCr + " Abs(IIf(com_compras.tipdoc<>7,0,IIf(con_retencion.idmon=1,con_retenciondet.impret,IIf(con_tc.impven Is Null,0,con_retenciondet.impret*con_tc.impven)))) AS imphabsol, " _
                + vbCr + " Abs(IIf(com_compras.tipdoc<>7,0,IIf(con_retencion.idmon=2,con_retenciondet.impret,0))) AS imphabdol, " _
                + vbCr + " 0 AS tipmov, 0 AS prodiv, con_retencion.fchreg AS fchasi, con_retencion.fchemi AS fchdoc, 0 AS idorides, 0 AS idmod, con_retenciondet.iddoc AS iddocpro, 0 AS tipo,con_retencion.idmon, " _
                + vbCr + " 2 AS ridlib,2 AS ridtipper, com_compras.idpro AS ridper, com_compras.tipdoc AS rtipdoc, com_compras.fchdoc AS rfchope, IIf([com_compras]![numser] Is Null Or [com_compras]![numser]='','',[com_compras]![numser] & '-') & [com_compras]![numdoc] AS rnumerodoc, Left([com_compras].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([com_compras].[numreg],4) AS rregistro, com_compras.glosa AS rglosa,con_retencion.glosa AS rglosaope,com_compras.idmon as ridmon, " _
                + vbCr + " null as rtipdoc1, null as rnumerodoc1, 2 as ridtipper2,con_retencion.idpro as ridper2, con_retencion.iddoc as rtipdoc2, con_retencion.fchemi as rfchope2, IIf([con_retencion]![numser] Is Null Or [con_retencion]![numser]='','',[con_retencion]![numser] & '-') & [con_retencion]![numdoc] as rnumerodoc2 " _
                + vbCr + " ,0 as aplicatc, 0 as ajuste " _
                + vbCr + " FROM (con_retencion LEFT JOIN con_tc ON con_retencion.fchemi = con_tc.fecha) INNER JOIN (((con_retenciondet INNER JOIN com_compras ON con_retenciondet.iddoc = com_compras.id) INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (com_compras.tipdoc = mae_documentocta.iddoc) AND (com_compras.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON con_retencion.id = con_retenciondet.id " _
                + vbCr + " WHERE con_retencion.tipo=1 AND mae_documentocta.tipope=0 and con_retencion.id = " & IdMovimiento

            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
            '===========================================================================================================================

        Case 6 '--caja y bancos
            '--23/10/09::Optimizar codigo para que su mantenimiento sera mas rapido y la escritura del asiento tambien
            '--Se separara por modulo segun su ingreso en el sistema
            '--25/01/11:: Escribir en destino de egreso en rregistro vacio cuande sera reembolsable
                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE tes_caja SET tes_caja.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((tes_caja.id)=" & IdMovimiento & "))"
                '***********************************************************************************
                Dim xRst1 As New ADODB.Recordset
                Dim xrtipdoc2 As Integer
                Dim xrfchope2 As String
                Dim xrnumerodoc2 As String
                Dim xridtipper2 As Integer
                Dim xridper2 As Long
                Dim nSQLRef As String
                nSQL = "SELECT TOP 1 tes_caja.id AS idtes, tes_cajaorigendet.tipdoc AS rtipdoc2, tes_caja.fchope AS rfchope2, tes_cajaorigendet.numdoc AS rnumerodoc2, 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, mae_bancos.numruc, mae_bancos.descripcion  " _
                    + vbCr + " FROM (tes_caja INNER JOIN ((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) ON tes_cajaori.idbcocta = mae_banconumcta.id " _
                    + vbCr + " WHERE mae_banconumcta.id Is Not Null and tes_cajaorigendet.idmod in (5,6,7) and tes_caja.id = " & IdMovimiento
                RST_Busq xRst1, nSQL, xCon
                If xRst1.RecordCount <> 0 Then
                    xrtipdoc2 = NulosN(xRst1("rtipdoc2"))
                    If IsDate(xRst1("rfchope2")) = True Then xrfchope2 = "'" & CDate(xRst1("rfchope2")) & "'"
                    xrnumerodoc2 = "'" & NulosC(xRst1("rnumerodoc2")) & "'"
                    xridtipper2 = NulosN(xRst1("ridtipper2"))
                    xridper2 = NulosN(xRst1("ridper2"))
                Else
                    xrtipdoc2 = 0
                    xrfchope2 = " null "
                    xrnumerodoc2 = "''"
                    xridtipper2 = 0
                    xridper2 = 0
                End If
                Set xRst1 = Nothing
                '--Se agregara para ver los documentos con que tipo de operacion se esta cobrando o pagando
                nSQLRef = xrtipdoc2 & " as rtipdoc2, " & xrfchope2 & " as rfchope2, " & xrnumerodoc2 & " as rnumerodoc2, " & xridtipper2 & " as ridtipper2, " & xridper2 & " as ridper2 "
                '===========================================================================================================================

                '--origen sin detalle
                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " IIF(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajaori.tc, " _
                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=1,tes_cajaori.importe,0) ) AS impdebsol, " _
                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=2,tes_cajaori.importe,0) ) AS impdebdol, " _
                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=1,tes_cajaori.importe,0) ) AS imphabsol, " _
                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=2,tes_cajaori.importe,0) ) AS imphabdol, " _
                    + vbCr + " tes_caja.tipmov , 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaori.idmod AS idmod, 0 AS iddocpro, 1 AS tipo, tes_caja.idmon,tes_cajaori.importe as impreal, " _
                    + vbCr + " 6 AS ridlib, 0 AS ridtipper, 0 AS ridper, 0 AS rtipdoc, tes_caja.fchope AS rfchope, '' AS rnumerodoc, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([tes_caja].[numreg],4) AS rregistro, NULL AS rglosa,tes_caja.glosa AS rglosaope ,tes_caja.idmon as ridmon, " _
                    + vbCr + " null as rtipdoc1, null as rnumerodoc1, null as ridtipper2,null as ridper2, null as rtipdoc2, null as rfchope2, null as rnumerodoc2 " _
                    + vbCr + " ,-1 as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM ((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idtes = tes_cajaorigendet.idtes) AND (tes_cajaori.idori = tes_cajaorigendet.idori)) ON tes_caja.id = tes_cajaori.idtes " _
                    + vbCr + " WHERE tes_cajaorigendet.idtes Is Null and tes_caja.id = " & IdMovimiento

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, True
                '===========================================================================================================================

                '--destino sin detalle
                '--se agrega una consulta adicional cuya finalidad es mostrar la informacion de banco o caja(numdoc, fchope,tipodoc)
                nSQL = " SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " IIF(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajadestino.tc, " _
                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=1,tes_cajadestino.importe,0) ) AS impdebsol, " _
                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=2,tes_cajadestino.importe,0) ) AS impdebdol, " _
                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=1,tes_cajadestino.importe,0) ) AS imphabsol, " _
                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=2,tes_cajadestino.importe,0) ) AS imphabdol, " _
                    + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestino.idmod , 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,tes_cajadestino.importe as impreal, " _
                    + vbCr + " 6 AS ridlib, 0 AS ridtipper, 0 AS ridper, 0 AS rtipdoc, tes_caja.fchope AS rfchope, '' AS rnumerodoc, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([tes_caja].[numreg],4) AS rregistro, NULL AS rglosa,tes_caja.glosa AS rglosaope ,tes_caja.idmon as ridmon, " _
                    + vbCr + " NULL as rtipdoc1, NULL as rnumerodoc1, " & nSQLRef & " ,-1 as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM ((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) LEFT JOIN tes_cajadestinodet ON (tes_cajadestino.idtes = tes_cajadestinodet.idtes) AND (tes_cajadestino.iddes = tes_cajadestinodet.iddes)) ON tes_caja.id = tes_cajadestino.idtes " _
                    + vbCr + " WHERE tes_cajadestinodet.idtes Is Null and tes_caja.id = " & IdMovimiento

'                nSQL = " SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " IIF(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajadestino.tc, " _
                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=1,tes_cajadestino.importe,0) ) AS impdebsol, " _
                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=2,tes_cajadestino.importe,0) ) AS impdebdol, " _
                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=1,tes_cajadestino.importe,0) ) AS imphabsol, " _
                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=2,tes_cajadestino.importe,0) ) AS imphabdol, " _
                    + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestino.idmod , 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,tes_cajadestino.importe as impreal, " _
                    + vbCr + " 6 AS ridlib, 0 AS ridtipper, 0 AS ridper, 0 AS rtipdoc, tes_caja.fchope AS rfchope, '' AS rnumerodoc, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([tes_caja].[numreg],4) AS rregistro, NULL AS rglosa,tes_caja.glosa AS rglosaope ,tes_caja.idmon as ridmon, " _
                    + vbCr + " NULL as rtipdoc1, NULL as rnumerodoc1, " _
                    + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2 " _
                    + vbCr + " ,-1 as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM (((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) INNER JOIN ((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) " _
                    + vbCr + " LEFT JOIN " _
                    + vbCr + " ( SELECT TOP 1 tes_caja.id AS idtes, tes_cajaorigendet.tipdoc AS rtipdoc2, tes_caja.fchope AS rfchope2, tes_cajaorigendet.numdoc AS rnumerodoc2, 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, mae_bancos.numruc, mae_bancos.descripcion " _
                    + vbCr + " FROM (tes_caja INNER JOIN ((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) ON tes_cajaori.idbcocta = mae_banconumcta.id " _
                    + vbCr + " WHERE mae_banconumcta.id Is Not Null and tes_cajaorigendet.idmod in (5,6,7) and tes_caja.id = " & IdMovimiento & "" _
                    + vbCr + " ) as ref2  ON tes_caja.id = ref2.idtes " _
                    + vbCr + " WHERE tes_cajadestinodet.idtes Is Null and tes_caja.id = " & IdMovimiento


                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, True
                '===========================================================================================================================

                '***********************************************************************************************************************************
                '***********************************************************************************************************************************
                '--origen con detalle

                RST_Busq xRst, "SELECT tes_caja.id, tes_cajaorigendet.idmod " _
                        & " FROM tes_caja INNER JOIN tes_cajaorigendet ON tes_caja.id = tes_cajaorigendet.idtes " _
                        & " GROUP BY tes_caja.id, tes_cajaorigendet.idmod " _
                        & " HAVING (((tes_caja.id)=" & IdMovimiento & "));", xConn

                If xRst.RecordCount <> 0 Then
                    xRst.MoveFirst
                    Do While Not xRst.EOF

                        Select Case xRst("idmod")
                            Case 1 '--compras

                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                     + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajaorigendet.iddoc, tes_cajaorigendet.corr AS correlativo, tes_cajaori.tc, " _
                                     + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS impdebsol, " _
                                     + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS impdebdol, " _
                                     + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS imphabsol, " _
                                     + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS imphabdol, " _
                                     + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaorigendet.idmod, 0 AS iddocpro, 1 AS tipo, tes_caja.idmon,(tes_cajaorigendet.importe+tes_cajaorigendet.acuenta) as impreal, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,1, 6) AS ridlib, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,1, tes_cajaorigendet.idtipper) AS ridtipper, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,com_compras.idpro, tes_cajaorigendet.idper) AS ridper, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,com_compras.tipdoc, tes_cajaorigendet.tipdoc) AS rtipdoc, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,com_compras.fchdoc, tes_cajaorigendet.fchdoc) AS rfchope, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,IIf(com_compras.numser Is Null Or com_compras.numser='','',com_compras.numser & '-') & com_compras.numdoc, IIf(tes_cajaorigendet.numser Is Null Or tes_cajaorigendet.numser='','',tes_cajaorigendet.numser & '-') & tes_cajaorigendet.numdoc) AS rnumerodoc, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,Left(com_compras.numreg,2) & mae_libros_1.codsun & Right(com_compras.numreg,4), Left(tes_caja.numreg,2) & mae_libros.codsun & Right(tes_caja.numreg,4)) AS rregistro, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,com_compras.glosa, tes_cajaorigendet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,com_compras.idmon, tes_caja.idmon) AS ridmon, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,com_compras.idtipdocref,tes_cajaorigendet.idtipdocref) AS rtipdoc1, " _
                                     + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,com_compras.numerodocref,tes_cajaorigendet.numerodocref) As rnumerodoc1, " _
                                     + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                     + vbCr + " FROM ((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN (tes_cajaorigendet LEFT JOIN (com_compras LEFT JOIN mae_libros AS mae_libros_1 ON com_compras.idlib = mae_libros_1.id) ON tes_cajaorigendet.iddoc = com_compras.id) ON (tes_cajaori.idtes = tes_cajaorigendet.idtes) AND (tes_cajaori.idori = tes_cajaorigendet.idori)) LEFT JOIN (con_percepcion LEFT JOIN mae_libros AS mae_libros_2 ON con_percepcion.idlib = mae_libros_2.id) ON tes_cajaorigendet.iddoc = con_percepcion.id) ON tes_caja.id = tes_cajaori.idtes  "
                                    ''WHERE   ((tes_cajaorigendet.idmod)=1 );

'                                     + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
'                                     + vbCr + " FROM ( (((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) ON tes_caja.id = tes_cajaori.idtes) INNER JOIN (tes_cajaorigendet LEFT JOIN (com_compras LEFT JOIN mae_libros AS mae_libros_1 ON com_compras.idlib = mae_libros_1.id) ON tes_cajaorigendet.iddoc = com_compras.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN (con_percepcion LEFT JOIN mae_libros AS mae_libros_2 ON con_percepcion.idlib = mae_libros_2.id) ON tes_cajaorigendet.iddoc = con_percepcion.id ) "

                            Case 2 '--ventas
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                    + vbCr + " IIF(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajaorigendet.iddoc, tes_cajaorigendet.corr AS correlativo, tes_cajaori.tc, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=1,tes_cajaorigendet.acuenta,0) ) AS impdebsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=2,tes_cajaorigendet.acuenta,0) ) AS impdebdol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=1,tes_cajaorigendet.acuenta,0) ) AS imphabsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=2,tes_cajaorigendet.acuenta,0) ) AS imphabdol, " _
                                    + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaorigendet.idmod, 0 AS iddocpro, 1 AS tipo, tes_caja.idmon,(tes_cajaorigendet.importe+tes_cajaorigendet.acuenta) as impreal, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,2,6) AS ridlib, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,2,0) AS ridtipper, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_ventas.idcli,tes_cajaorigendet.idper) AS ridper, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_ventas.tipdoc,tes_cajaorigendet.tipdoc) AS rtipdoc, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_ventas.fchdoc,tes_cajaorigendet.fchdoc) AS rfchope, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,IIf([vta_ventas].[numser] Is Null Or [vta_ventas].[numser]='','',[vta_ventas].[numser] & '-') & [vta_ventas].[numdoc],IIf([tes_cajaorigendet].[numser] Is Null Or [tes_cajaorigendet].[numser]='','',[tes_cajaorigendet].[numser] & '-') & [tes_cajaorigendet].[numdoc]) AS rnumerodoc, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,Left([vta_ventas].[numreg],2) & [mae_libros_1].[codsun] & Right([vta_ventas].[numreg],4),Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS rregistro, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_ventas.glosa,tes_cajaorigendet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_ventas.idmon,tes_caja.idmon) AS ridmon, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_ventas.idtipdocref,tes_cajaorigendet.idtipdocref) AS rtipdoc1, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_ventas.numerodocref,tes_cajaorigendet.numerodocref) AS rnumerodoc1, " _
                                    + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                    + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN (tes_cajaorigendet LEFT JOIN (vta_ventas LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas.idlib = mae_libros_1.id) ON tes_cajaorigendet.iddoc = vta_ventas.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes  "
                                    ''WHERE   ((tes_cajaorigendet.idmod)=2 );

''                                    + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                    + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN (tes_cajaorigendet LEFT JOIN (vta_ventas LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas.idlib = mae_libros_1.id) ON tes_cajaorigendet.iddoc = vta_ventas.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes  ) "


                            Case 4 '--Letras
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajaorigendet.iddoc, tes_cajaorigendet.corr AS correlativo, tes_cajaori.tc, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS impdebsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS impdebdol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS imphabsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS imphabdol, " _
                                        + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaorigendet.idmod, 0 AS iddocpro, 1 AS tipo, tes_caja.idmon,(tes_cajaorigendet.importe+tes_cajaorigendet.acuenta) as impreal, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,37, 6) AS ridlib, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,2, tes_cajaorigendet.idtipper) AS ridtipper, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_letra.idclipro, tes_cajaorigendet.idper) AS ridper, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_letra.tipdoc, tes_cajaorigendet.tipdoc) AS rtipdoc, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_letra.fchemi, tes_cajaorigendet.fchdoc) AS rfchope, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,[let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser], [tes_cajaorigendet].[numser] & ' ' & [tes_cajaorigendet].[numdoc]) AS rnumerodoc, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,Left([let_letra].[numreg],2) & [mae_libros_1].[codsun] & Right([let_letra].[numreg],4), Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS rregistro, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_letra.glosa, tes_cajaorigendet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_letra.idmon, tes_caja.idmon) AS ridmon, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_letra.idtipdocref, tes_cajaorigendet.idtipdocref) AS rtipdoc1, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_letra.numerodocref, tes_cajaorigendet.numerodocref) AS rnumerodoc1, " _
                                        + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                        + vbCr + " FROM ((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((let_letra LEFT JOIN mae_libros AS mae_libros_1 ON let_letra.idlib = mae_libros_1.id) RIGHT JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN (tes_cajaorigendet LEFT JOIN let_letradet ON tes_cajaorigendet.iddoc = let_letradet.corr) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON let_letra.id = let_letradet.idlet) ON tes_caja.id = tes_cajaori.idtes  "
                                        ''WHERE   ((tes_cajaorigendet.idmod)=4 );

''                                        + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                        + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((let_letra LEFT JOIN mae_libros AS mae_libros_1 ON let_letra.idlib = mae_libros_1.id) RIGHT JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN (tes_cajaorigendet LEFT JOIN let_letradet ON tes_cajaorigendet.iddoc = let_letradet.corr) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON let_letra.id = let_letradet.idlet) ON tes_caja.id = tes_cajaori.idtes ) "


                            Case 11 '--liquidacion gasto debito
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                    + vbCr + " IIF(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajaorigendet.iddoc, tes_cajaorigendet.corr AS correlativo, tes_cajaori.tc, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=1,tes_cajaorigendet.acuenta,0) ) AS impdebsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=2,tes_cajaorigendet.acuenta,0) ) AS impdebdol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=1,tes_cajaorigendet.acuenta,0) ) AS imphabsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=2,tes_cajaorigendet.acuenta,0) ) AS imphabdol, " _
                                    + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaorigendet.idmod, 0 AS iddocpro, 1 AS tipo, tes_caja.idmon,(tes_cajaorigendet.importe+tes_cajaorigendet.acuenta) as impreal, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,41,6) AS ridlib, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,2,0) AS ridtipper, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_gastodebito.idcli,tes_cajaorigendet.idper) AS ridper, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_gastodebito.tipdoc,tes_cajaorigendet.tipdoc) AS rtipdoc, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_gastodebito.fchemi,tes_cajaorigendet.fchdoc) AS rfchope, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,IIf([vta_gastodebito].[numser] Is Null Or [vta_gastodebito].[numser]='','',[vta_gastodebito].[numser] & '-') & [vta_gastodebito].[numdoc],IIf([tes_cajaorigendet].[numser] Is Null Or [tes_cajaorigendet].[numser]='','',[tes_cajaorigendet].[numser] & '-') & [tes_cajaorigendet].[numdoc]) AS rnumerodoc, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,Left([vta_gastodebito].[numreg],2) & [mae_libros_1].[codsun] & Right([vta_gastodebito].[numreg],4),Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS rregistro, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_gastodebito.glosa,tes_cajaorigendet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_gastodebito.idmon,tes_caja.idmon) AS ridmon, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_gastodebito.idtipdocref,tes_cajaorigendet.idtipdocref) AS rtipdoc1, " _
                                    + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,vta_gastodebito.numerodocref,tes_cajaorigendet.numerodocref) AS rnumerodoc1, " _
                                    + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                    + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN ((tes_cajaorigendet LEFT JOIN vta_gastodebito ON tes_cajaorigendet.iddoc = vta_gastodebito.id) LEFT JOIN mae_libros AS mae_libros_1 ON vta_gastodebito.idlib = mae_libros_1.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes "
                                    ''WHERE   ((tes_cajaorigendet.idmod)=11 );

''                                    + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                    + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN ((tes_cajaorigendet LEFT JOIN vta_gastodebito ON tes_cajaorigendet.iddoc = vta_gastodebito.id) LEFT JOIN mae_libros AS mae_libros_1 ON vta_gastodebito.idlib = mae_libros_1.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes ) "


                            Case 19 '--planilla letras
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajaorigendet.iddoc, tes_cajaorigendet.corr AS correlativo, tes_cajaori.tc, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS impdebsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS impdebdol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS imphabsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS imphabdol, " _
                                        + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaorigendet.idmod, 0 AS iddocpro, 1 AS tipo, tes_caja.idmon,(tes_cajaorigendet.importe+tes_cajaorigendet.acuenta) as impreal, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,42, 6) AS ridlib, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,5, tes_cajaorigendet.idtipper) AS ridtipper, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,mae_banconumcta.idban, tes_cajaorigendet.idper) AS ridper, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_planilla.tipdoc, tes_cajaorigendet.tipdoc) AS rtipdoc, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_planilla.fchemi, tes_cajaorigendet.fchdoc) AS rfchope, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_planilla.numdoc, IIf([tes_cajaorigendet].[numser] Is Null Or [tes_cajaorigendet].[numser]='','',[tes_cajaorigendet].[numser] & '-') & [tes_cajaorigendet].[numdoc]) AS rnumerodoc, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,Left([let_planilla].[numreg],2) & [mae_libros_1].[codsun] & Right([let_planilla].[numreg],4),Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS rregistro, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_planilla.glosa, tes_cajaorigendet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,let_planilla.idmon, tes_caja.idmon) AS ridmon, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,0, tes_cajaorigendet.idtipdocref) AS rtipdoc1, " _
                                        + vbCr + " IIF(tes_cajaorigendet.iddoc <>0 ,'', tes_cajaorigendet.numerodocref) AS rnumerodoc1, " _
                                        + vbCr + nSQLRef & ", -1 as aplicatc, 0 as ajuste " _
                                        + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (mae_banconumcta RIGHT JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN (tes_cajaorigendet LEFT JOIN (mae_libros AS mae_libros_1 RIGHT JOIN let_planilla ON mae_libros_1.id = let_planilla.idlib) ON tes_cajaorigendet.iddoc = let_planilla.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON mae_banconumcta.id = let_planilla.idbcocta) ON tes_caja.id = tes_cajaori.idtes "
                                        '--WHERE    ((tes_cajaorigendet.idmod)=19);

''                                        + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                        + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (mae_banconumcta RIGHT JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN (tes_cajaorigendet LEFT JOIN (mae_libros AS mae_libros_1 RIGHT JOIN let_planilla ON mae_libros_1.id = let_planilla.idlib) ON tes_cajaorigendet.iddoc = let_planilla.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON mae_banconumcta.id = let_planilla.idbcocta) ON tes_caja.id = tes_cajaori.idtes ) "


                            Case Else
                                '--para los casos siguientes
                                '--3   Entregas a Rendir
                                '--5   Anticipos
                                '--6   Bancos
                                '--7   Fondo Fijo
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso')  AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                    + vbCr + " IIF(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajaorigendet.iddoc, tes_cajaorigendet.corr AS correlativo, tes_cajaori.tc, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=1,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS impdebsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIf(tes_caja.idmon=2,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) ) AS impdebdol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=1,tes_cajaorigendet.importe,0) ) AS imphabsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIf(tes_caja.idmon=2,tes_cajaorigendet.importe,0) ) AS imphabdol, " _
                                    + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaorigendet.idmod, 0 AS iddocpro, 1 AS tipo, tes_caja.idmon,(tes_cajaorigendet.importe+tes_cajaorigendet.acuenta) as impreal, " _
                                    + vbCr + " 6 AS ridlib, " _
                                    + vbCr + " tes_cajaorigendet.idtipper AS ridtipper, " _
                                    + vbCr + " tes_cajaorigendet.idper AS ridper, " _
                                    + vbCr + " tes_cajaorigendet.tipdoc AS rtipdoc, " _
                                    + vbCr + " tes_cajaorigendet.fchdoc AS rfchope, " _
                                    + vbCr + " IIf([tes_cajaorigendet].[numser] Is Null Or [tes_cajaorigendet].[numser]='','',[tes_cajaorigendet].[numser] & '-') & [tes_cajaorigendet].[numdoc] AS rnumerodoc, " _
                                    + vbCr + " Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4) AS rregistro, " _
                                    + vbCr + " tes_cajaorigendet.glosa AS rglosa,tes_caja.glosa AS rglosaope, " _
                                    + vbCr + " tes_caja.idmon AS ridmon, " _
                                    + vbCr + " tes_cajaorigendet.idtipdocref AS rtipdoc1, " _
                                    + vbCr + " tes_cajaorigendet.numerodocref AS rnumerodoc1, " _
                                    + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                    + vbCr + " FROM ((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idtes = tes_cajaorigendet.idtes) AND (tes_cajaori.idori = tes_cajaorigendet.idori)) ON tes_caja.id = tes_cajaori.idtes  "
                                    ''WHERE   ((tes_cajaorigendet.idmod) in (3,5,6,7) );

''                                    + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                    + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes  ) "


                            End Select
                                '--consulta adicional para mostrar la referencia de la operacion
                                '--
                                nSQL = nSQL + "" _
                                    + vbCr + " WHERE tes_cajaorigendet.idmod = " & xRst("idmod") & " AND tes_cajaorigendet.idtes = " & IdMovimiento

''                                nSQL = nSQL + "" _
                                    + vbCr + " LEFT JOIN " _
                                    + vbCr + " ( SELECT TOP 1 tes_caja.id AS idtes, tes_cajaorigendet.tipdoc AS rtipdoc2, tes_caja.fchope AS rfchope2, tes_cajaorigendet.numdoc AS rnumerodoc2, 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, mae_bancos.numruc, mae_bancos.descripcion " _
                                    + vbCr + " FROM (tes_caja INNER JOIN ((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) ON tes_cajaori.idbcocta = mae_banconumcta.id " _
                                    + vbCr + " WHERE  mae_banconumcta.id Is Not Null and tes_cajaorigendet.idmod in (5,6,7)and tes_caja.id = " & IdMovimiento & "" _
                                    + vbCr + " ) as ref2  ON tes_caja.id = ref2.idtes " _
                                    + vbCr + " WHERE tes_cajaorigendet.idmod = " & xRst("idmod") & " AND tes_caja.id = " & IdMovimiento


                                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, True
                                '===========================================================================================================================


                        '--ir al siguiente registro
                        xRst.MoveNext
                    Loop
                End If

                '***********************************************************************************************************************************
                '***********************************************************************************************************************************
                '--Destino con detalle

                '--limpiando el rst
                Set xRst = Nothing

                RST_Busq xRst, "SELECT tes_caja.id, tes_cajadestinodet.idmod " _
                        & " FROM tes_caja INNER JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
                        & " GROUP BY tes_caja.id, tes_cajadestinodet.idmod " _
                        & " HAVING (((tes_caja.id)=" & IdMovimiento & "));", xConn

                If xRst.RecordCount <> 0 Then
                    xRst.MoveFirst
                    Do While Not xRst.EOF
                        '--

                        Select Case xRst("idmod")
                            Case 1 '--Compras
                            '--asiento para compras y percepciones segun campo:tes_cajadestinodet.idori(1::compras; 2::Percepcion)
                            nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                    + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajadestinodet.iddoc, tes_cajadestinodet.corr AS correlativo, tes_cajadestino.tc, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebdol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabdol, " _
                                    + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod, 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,(tes_cajadestinodet.importe+tes_cajadestinodet.acuenta) as impreal, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, 1, 4), 6) AS ridlib, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, 1, con_percepcion.tipo) ,tes_cajadestinodet.idtipper) AS ridtipper, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, com_compras.idpro, con_percepcion.idcli) ,tes_cajadestinodet.idper) AS ridper, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, com_compras.tipdoc, con_percepcion.tipdoc) ,tes_cajadestinodet.tipdoc) AS rtipdoc, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, com_compras.fchdoc, con_percepcion.fchdoc), tes_cajadestinodet.fchdoc) AS rfchope, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, IIf(com_compras.numser Is Null Or com_compras.numser='','',com_compras.numser & '-') & com_compras.numdoc, IIf(con_percepcion.numser Is Null Or con_percepcion.numser='','',con_percepcion.numser & '-') & con_percepcion.numdoc), IIf(tes_cajadestinodet.numser Is Null Or tes_cajadestinodet.numser='','',tes_cajadestinodet.numser & '-') & tes_cajadestinodet.numdoc) AS rnumerodoc, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, Left(com_compras.numreg,2) & mae_libros_1.codsun & Right(com_compras.numreg,4), Left(con_percepcion.numreg,2) & mae_libros_2.codsun & Right(con_percepcion.numreg,4)), Left(tes_caja.numreg,2) & mae_libros.codsun & Right(tes_caja.numreg,4)) AS rregistro, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, com_compras.glosa, con_percepcion.glosa), tes_cajadestinodet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, com_compras.idmon, con_percepcion.idmon), tes_caja.idmon) AS ridmon, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, com_compras.idtipdocref, 0), tes_cajadestinodet.idtipdocref) AS rtipdoc1, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF(tes_cajadestinodet.idori <> 2, com_compras.numerodocref, ''), tes_cajadestinodet.numerodocref) AS rnumerodoc1, " _
                                    + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                    + vbCr + " FROM ((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN (com_compras LEFT JOIN mae_libros AS mae_libros_1 ON com_compras.idlib = mae_libros_1.id) ON tes_cajadestinodet.iddoc = com_compras.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) LEFT JOIN (con_percepcion LEFT JOIN mae_libros AS mae_libros_2 ON con_percepcion.idlib = mae_libros_2.id) ON tes_cajadestinodet.iddoc = con_percepcion.id) ON tes_caja.id = tes_cajadestino.idtes "

                                    ''WHERE   ((tes_cajadestinodet.idmod)=1 );

''                                    + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                    + vbCr + " FROM ( (((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) ON tes_caja.id = tes_cajadestino.idtes) INNER JOIN (tes_cajadestinodet LEFT JOIN (com_compras LEFT JOIN mae_libros AS mae_libros_1 ON com_compras.idlib = mae_libros_1.id) ON tes_cajadestinodet.iddoc = com_compras.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) LEFT JOIN (con_percepcion LEFT JOIN mae_libros AS mae_libros_2 ON con_percepcion.idlib = mae_libros_2.id) ON tes_cajadestinodet.iddoc = con_percepcion.id ) "


                            Case 2 '--ventas
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajadestinodet.iddoc, tes_cajadestinodet.corr AS correlativo, tes_cajadestino.tc, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebdol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabdol, " _
                                        + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod, 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,(tes_cajadestinodet.importe+tes_cajadestinodet.acuenta) as impreal, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,2, 6) AS ridlib, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,2, tes_cajadestinodet.idtipper) AS ridtipper, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_ventas.idcli, tes_cajadestinodet.idper) AS ridper, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_ventas.tipdoc, tes_cajadestinodet.tipdoc) AS rtipdoc, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_ventas.fchdoc, tes_cajadestinodet.fchdoc) AS rfchope, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIf([vta_ventas].[numser] Is Null Or [vta_ventas].[numser]='','',[vta_ventas].[numser] & '-') & [vta_ventas].[numdoc], IIF([tes_cajadestinodet].[numser] Is Null Or [tes_cajadestinodet].[numser]='','',[tes_cajadestinodet].[numser] & '-') & [tes_cajadestinodet].[numdoc]) AS rnumerodoc, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,Left([vta_ventas].[numreg],2) & [mae_libros_1].[codsun] & Right([vta_ventas].[numreg],4), Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS rregistro, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_ventas.glosa, tes_cajadestinodet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_ventas.idmon, tes_caja.idmon) AS ridmon, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_ventas.idtipdocref, tes_cajadestinodet.idtipdocref) AS rtipdoc1, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_ventas.numerodocref, tes_cajadestinodet.numerodocref) AS rnumerodoc1, " _
                                        + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                        + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN (vta_ventas LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas.idlib = mae_libros_1.id) ON tes_cajadestinodet.iddoc = vta_ventas.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes "
                                        ''WHERE ((tes_cajadestinodet.idmod)=2);

                            Case 4 '--letras
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajadestinodet.iddoc, tes_cajadestinodet.corr AS correlativo, tes_cajadestino.tc, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebdol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabdol, " _
                                        + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod, 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,(tes_cajadestinodet.importe+tes_cajadestinodet.acuenta) as impreal, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,37, 6) AS ridlib, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,2, tes_cajadestinodet.idtipper) AS ridtipper, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_letra.idclipro, tes_cajadestinodet.idper) AS ridper, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_letra.tipdoc, tes_cajadestinodet.tipdoc) AS rtipdoc, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_letra.fchemi, tes_cajadestinodet.fchdoc) AS rfchope, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,[let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser], [tes_cajadestinodet].[numser] & ' ' & [tes_cajadestinodet].[numdoc]) AS rnumerodoc, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,Left([let_letra].[numreg],2) & [mae_libros_1].[codsun] & Right([let_letra].[numreg],4), Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS rregistro, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_letra.glosa, tes_cajadestinodet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_letra.idmon, tes_caja.idmon) AS ridmon, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_letra.idtipdocref, tes_cajadestinodet.idtipdocref) AS rtipdoc1, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_letra.numerodocref, tes_cajadestinodet.numerodocref) AS rnumerodoc1, " _
                                        + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                        + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((let_letra LEFT JOIN mae_libros AS mae_libros_1 ON let_letra.idlib = mae_libros_1.id) RIGHT JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN let_letradet ON tes_cajadestinodet.iddoc = let_letradet.corr) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON let_letra.id = let_letradet.idlet) ON tes_caja.id = tes_cajadestino.idtes "
                                        '---WHERE  ((tes_cajadestinodet.idmod)=4);

''                                        + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                        + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((let_letra LEFT JOIN mae_libros AS mae_libros_1 ON let_letra.idlib = mae_libros_1.id) RIGHT JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN let_letradet ON tes_cajadestinodet.iddoc = let_letradet.corr) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON let_letra.id = let_letradet.idlet) ON tes_caja.id = tes_cajadestino.idtes ) "


                            Case 8 '--Planillas
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajadestinodet.iddoc, tes_cajadestinodet.corr AS correlativo, tes_cajadestino.tc, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebdol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabdol, " _
                                        + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod, 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,(tes_cajadestinodet.importe+tes_cajadestinodet.acuenta) as impreal, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,9, 6) AS ridlib, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,3, tes_cajadestinodet.idtipper) AS ridtipper, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,pla_boleta.idemp, tes_cajadestinodet.idper) AS ridper, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,pla_boleta.iddoc, tes_cajadestinodet.tipdoc) AS rtipdoc, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,pla_boleta.fchdoc, tes_cajadestinodet.fchdoc) AS rfchope, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIf([pla_boleta].[numser] Is Null Or [pla_boleta].[numser]='','',[pla_boleta].[numser] & '-') & [pla_boleta].[numdoc], IIF([tes_cajadestinodet].[numser] Is Null Or [tes_cajadestinodet].[numser]='','',[tes_cajadestinodet].[numser] & '-') & [tes_cajadestinodet].[numdoc]) AS rnumerodoc, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,Left([pla_boleta].[numreg],2) & [mae_libros_1].[codsun] & Right([pla_boleta].[numreg],4), Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS rregistro, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,pla_boleta.glosa, tes_cajadestinodet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,pla_boleta.idmon,tes_caja.idmon) AS ridmon, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,0, tes_cajadestinodet.idtipdocref) AS rtipdoc1, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,'', tes_cajadestinodet.numerodocref) AS rnumerodoc1, " _
                                        + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                        + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN (pla_boleta LEFT JOIN mae_libros AS mae_libros_1 ON pla_boleta.idlib = mae_libros_1.id) ON tes_cajadestinodet.iddoc = pla_boleta.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes "
                                        ''WHERE ((tes_cajadestinodet.idmod)=8);

''                                        + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                        + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN (pla_boleta LEFT JOIN mae_libros AS mae_libros_1 ON pla_boleta.idlib = mae_libros_1.id) ON tes_cajadestinodet.iddoc = pla_boleta.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes ) "


                            Case 9 '--Honorarios
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajadestinodet.iddoc, tes_cajadestinodet.corr AS correlativo, tes_cajadestino.tc, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebdol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabdol, " _
                                        + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod, 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,(tes_cajadestinodet.importe+tes_cajadestinodet.acuenta) as impreal, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,40, 6) AS ridlib, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,1, tes_cajadestinodet.idtipper) AS ridtipper, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_honorarios.idpro, tes_cajadestinodet.idper) AS ridper, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_honorarios.tipdoc, tes_cajadestinodet.tipdoc) AS rtipdoc, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_honorarios.fchdoc, tes_cajadestinodet.fchdoc) AS rfchope, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIF([com_honorarios].[numser] Is Null Or [com_honorarios].[numser]='','',[com_honorarios].[numser] & '-') & [com_honorarios].[numdoc], IIF([tes_cajadestinodet].[numser] Is Null Or [tes_cajadestinodet].[numser]='','',[tes_cajadestinodet].[numser] & '-') & [tes_cajadestinodet].[numdoc]) AS rnumerodoc, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,Left([com_honorarios].[numreg],2) & [mae_libros_1].[codsun] & Right([com_honorarios].[numreg],4), Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS rregistro, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_honorarios.glosa, tes_cajadestinodet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_honorarios.idmon, tes_caja.idmon) AS ridmon, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_honorarios.idtipdocref, tes_cajadestinodet.idtipdocref) AS rtipdoc1, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_honorarios.numerodocref, tes_cajadestinodet.numerodocref) AS rnumerodoc1, " _
                                        + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                        + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN (com_honorarios LEFT JOIN mae_libros AS mae_libros_1 ON com_honorarios.idlib = mae_libros_1.id) ON tes_cajadestinodet.iddoc = com_honorarios.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes "
                                        '---WHERE  ((tes_cajadestinodet.idmod)=9);

''                                        + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                        + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN (com_honorarios LEFT JOIN mae_libros AS mae_libros_1 ON com_honorarios.idlib = mae_libros_1.id) ON tes_cajadestinodet.iddoc = com_honorarios.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes ) "


                            Case 10 '--Reembolsables
                                '--la provision de los reembolsables no genera asiento contable, en banco se escribe en blanco el num registro de referencia al reembolsable
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                    + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajadestinodet.iddoc, tes_cajadestinodet.corr AS correlativo, tes_cajadestino.tc, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebdol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabdol, " _
                                    + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod, 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,(tes_cajadestinodet.importe+tes_cajadestinodet.acuenta) as impreal, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,999, 6) AS ridlib, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,1, tes_cajadestinodet.idtipper) AS ridtipper, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_reembolsables.idpro, tes_cajadestinodet.idper) AS ridper, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_reembolsables.tipdoc, tes_cajadestinodet.tipdoc) AS rtipdoc, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_reembolsables.fchdoc, tes_cajadestinodet.fchdoc) AS rfchope, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIf([com_reembolsables].[numser] Is Null Or [com_reembolsables].[numser]='','',[com_reembolsables].[numser] & '-') & [com_reembolsables].[numdoc], IIF([tes_cajadestinodet].[numser] Is Null Or [tes_cajadestinodet].[numser]='','',[tes_cajadestinodet].[numser] & '-') & [tes_cajadestinodet].[numdoc]) AS rnumerodoc, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,'', '') AS rregistro, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_reembolsables.glosa, tes_cajadestinodet.glosa) AS rglosa, tes_caja.glosa AS rglosaope, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_reembolsables.idmon, tes_caja.idmon) AS ridmon, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_reembolsables.idtipdocref, tes_cajadestinodet.idtipdocref) AS rtipdoc1, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,com_reembolsables.numerodocref, tes_cajadestinodet.numerodocref) AS rnumerodoc1, " _
                                    + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                    + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN com_reembolsables ON tes_cajadestinodet.iddoc = com_reembolsables.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes "
                                    ''WHERE (((tes_caja.id)=2194) AND ((tes_cajadestinodet.idmod)=10));

''                                    + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                    + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN com_reembolsables ON tes_cajadestinodet.iddoc = com_reembolsables.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes ) "


                            Case 11 '--Liquidación Gasto Debito
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                    + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajadestinodet.iddoc, tes_cajadestinodet.corr AS correlativo, tes_cajadestino.tc, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebdol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabdol, " _
                                    + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod, 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,(tes_cajadestinodet.importe+tes_cajadestinodet.acuenta) as impreal, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,41, 6) AS ridlib, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,2, tes_cajadestinodet.idtipper) AS ridtipper, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_gastodebito.idcli, tes_cajadestinodet.idper) AS ridper, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_gastodebito.tipdoc, tes_cajadestinodet.tipdoc) AS rtipdoc, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_gastodebito.fchemi, tes_cajadestinodet.fchdoc) AS rfchope, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,IIf([vta_gastodebito].[numser] Is Null Or [vta_gastodebito].[numser]='','',[vta_gastodebito].[numser] & '-') & [vta_gastodebito].[numdoc], IIF([tes_cajadestinodet].[numser] Is Null Or [tes_cajadestinodet].[numser]='','',[tes_cajadestinodet].[numser] & '-') & [tes_cajadestinodet].[numdoc]) AS rnumerodoc, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,Left([vta_gastodebito].[numreg],2) & [mae_libros_1].[codsun] & Right([vta_gastodebito].[numreg],4), Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS rregistro, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_gastodebito.glosa, tes_cajadestinodet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_gastodebito.idmon, tes_caja.idmon) AS ridmon, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_gastodebito.idtipdocref, tes_cajadestinodet.idtipdocref) AS rtipdoc1, " _
                                    + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,vta_gastodebito.numerodocref, tes_cajadestinodet.numerodocref) AS rnumerodoc1, " _
                                    + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                    + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN ((tes_cajadestinodet LEFT JOIN vta_gastodebito ON tes_cajadestinodet.iddoc = vta_gastodebito.id) LEFT JOIN mae_libros AS mae_libros_1 ON vta_gastodebito.idlib = mae_libros_1.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes "
                                    ''--WHERE  ((tes_cajadestinodet.idmod)=11);

''                                    + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc" _
''                                    + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN ((tes_cajadestinodet LEFT JOIN vta_gastodebito ON tes_cajadestinodet.iddoc = vta_gastodebito.id) LEFT JOIN mae_libros AS mae_libros_1 ON vta_gastodebito.idlib = mae_libros_1.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes ) "


                            Case 19 '--Planilla de letras
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajadestinodet.iddoc, tes_cajadestinodet.corr AS correlativo, tes_cajadestino.tc, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebdol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabsol, " _
                                        + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabdol, " _
                                        + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod, 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,(tes_cajadestinodet.importe+tes_cajadestinodet.acuenta) as impreal, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,42, 6) AS ridlib, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,5, tes_cajadestinodet.idtipper) AS ridtipper, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,mae_banconumcta.idban, tes_cajadestinodet.idper) AS ridper, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_planilla.tipdoc, tes_cajadestinodet.tipdoc) AS rtipdoc, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_planilla.fchemi, tes_cajadestinodet.fchdoc) AS rfchope, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_planilla.numdoc, IIf([tes_cajadestinodet].[numser] Is Null Or [tes_cajadestinodet].[numser]='','',[tes_cajadestinodet].[numser] & '-') & [tes_cajadestinodet].[numdoc]) AS rnumerodoc, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,Left([let_planilla].[numreg],2) & [mae_libros_1].[codsun] & Right([let_planilla].[numreg],4),Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS rregistro, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_planilla.glosa, tes_cajadestinodet.glosa) AS rglosa,tes_caja.glosa AS rglosaope, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,let_planilla.idmon, tes_caja.idmon) AS ridmon, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,0, tes_cajadestinodet.idtipdocref) AS rtipdoc1, " _
                                        + vbCr + " IIF(tes_cajadestinodet.iddoc <>0 ,'', tes_cajadestinodet.numerodocref) AS rnumerodoc1, " _
                                        + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                        + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (mae_banconumcta RIGHT JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN (mae_libros AS mae_libros_1 RIGHT JOIN let_planilla ON mae_libros_1.id = let_planilla.idlib) ON tes_cajadestinodet.iddoc = let_planilla.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON mae_banconumcta.id = let_planilla.idbcocta) ON tes_caja.id = tes_cajadestino.idtes "
                                        '--WHERE    ((tes_cajadestinodet.idmod)=19);

''                                        + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                        + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (mae_banconumcta RIGHT JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (tes_cajadestinodet LEFT JOIN (mae_libros AS mae_libros_1 RIGHT JOIN let_planilla ON mae_libros_1.id = let_planilla.idlib) ON tes_cajadestinodet.iddoc = let_planilla.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON mae_banconumcta.id = let_planilla.idbcocta) ON tes_caja.id = tes_cajadestino.idtes ) "

                            Case Else
                                '--para los casos presentados
                                '--3   Entregas a Rendir
                                '--5   Anticipos
                                '--6   Bancos
                                '--7   Fondo Fijo
                                nSQL = "SELECT IIF(tes_caja.tipmov=1,'Ingreso' ,'Egreso') AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                                    + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, tes_cajadestinodet.iddoc, tes_cajadestinodet.corr AS correlativo, tes_cajadestino.tc, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=1,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS impdebdol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabsol, " _
                                    + vbCr + " IIF(tes_caja.tipmov=2,0 , IIF(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) ) AS imphabdol, " _
                                    + vbCr + " tes_caja.tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod, 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,(tes_cajadestinodet.importe+tes_cajadestinodet.acuenta) as impreal, " _
                                    + vbCr + " 6 AS ridlib, " _
                                    + vbCr + " tes_cajadestinodet.idtipper AS ridtipper, " _
                                    + vbCr + " tes_cajadestinodet.idper AS ridper, " _
                                    + vbCr + " tes_cajadestinodet.tipdoc AS rtipdoc, " _
                                    + vbCr + " tes_cajadestinodet.fchdoc AS rfchope, " _
                                    + vbCr + " IIf([tes_cajadestinodet].[numser] Is Null Or [tes_cajadestinodet].[numser]='','',[tes_cajadestinodet].[numser] & '-') & [tes_cajadestinodet].[numdoc] AS rnumerodoc, " _
                                    + vbCr + " Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4) AS rregistro, " _
                                    + vbCr + " tes_cajadestinodet.glosa AS rglosa,tes_caja.glosa AS rglosaope, " _
                                    + vbCr + " tes_caja.idmon AS ridmon, " _
                                    + vbCr + " tes_cajadestinodet.idtipdocref AS rtipdoc1, " _
                                    + vbCr + " tes_cajadestinodet.numerodocref AS rnumerodoc1, " _
                                    + vbCr + nSQLRef & " , -1 as aplicatc, 0 as ajuste " _
                                    + vbCr + " FROM (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes "
                                    '--WHERE  ((tes_cajadestinodet.idmod) In (3,5,6,7));

''                                    + vbCr + " ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2, -1 as aplicatc, 0 as ajuste " _
''                                    + vbCr + " FROM ( (((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) ON tes_caja.id = tes_cajadestino.idtes ) "

                        End Select

                            nSQL = nSQL & " " _
                                + vbCr + " WHERE tes_cajadestinodet.idmod = " & xRst("idmod") & " AND tes_cajadestinodet.idtes = " & IdMovimiento

'                            nSQL = nSQL & " " _
                                + vbCr + " LEFT JOIN " _
                                + vbCr + " ( SELECT TOP 1 tes_caja.id AS idtes, tes_cajaorigendet.tipdoc AS rtipdoc2, tes_caja.fchope AS rfchope2, tes_cajaorigendet.numdoc AS rnumerodoc2, 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, mae_bancos.numruc, mae_bancos.descripcion " _
                                + vbCr + " FROM (tes_caja INNER JOIN ((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) ON tes_cajaori.idbcocta = mae_banconumcta.id " _
                                + vbCr + " WHERE mae_banconumcta.id Is Not Null and tes_cajaorigendet.idmod in (5,6,7) and tes_caja.id = " & IdMovimiento & "" _
                                + vbCr + " ) as ref2  ON tes_caja.id = ref2.idtes " _
                                + vbCr + " WHERE tes_cajadestinodet.idmod = " & xRst("idmod") & " AND tes_caja.id = " & IdMovimiento


                            GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, True
                            '===========================================================================================================================




                        xRst.MoveNext
                    Loop
                End If


                '***********************************************************************************************************************************
                '***********************************************************************************************************************************


        Case 666 '--caja y bancos
                'antiguo
                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE tes_caja SET tes_caja.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((tes_caja.id)=" & IdMovimiento & "))"
                '***********************************************************************************

                '--egresos
                If TipoMovimiento = 2 Then
                    '--origen sin detalle
                    nSQL = "SELECT 'Egreso' AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajaori.tc, " _
                        + vbCr + " 0 AS impdebsol, " _
                        + vbCr + " 0 AS impdebdol, " _
                        + vbCr + " IIf(tes_caja.idmon=1,tes_cajaori.importe,0) AS imphabsol, " _
                        + vbCr + " IIf(tes_caja.idmon=2,tes_cajaori.importe,0) AS imphabdol, " _
                        + vbCr + " 0 AS tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaori.idmod AS idmod, 0 AS iddocpro, 1 AS tipo, tes_caja.idmon,tes_cajaori.importe as impreal, " _
                        + vbCr + " 6 AS ridlib, 0 AS ridtipper, 0 AS ridper, 0 AS rtipdoc, tes_caja.fchope AS rfchope, '' AS rnumerodoc, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([tes_caja].[numreg],4) AS rregistro, NULL AS rglosa,tes_caja.glosa AS rglosaope ,tes_caja.idmon as ridmon, " _
                        + vbCr + " null as rtipdoc1, null as rnumerodoc1, null as ridtipper2,null as ridper2, null as rtipdoc2, null as rfchope2, null as rnumerodoc2 " _
                        + vbCr + " ,-1 as aplicatc, 0 as ajuste " _
                        + vbCr + " FROM ((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) INNER JOIN ((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id " _
                        + vbCr + " WHERE tes_caja.tipmov=2 and tes_cajaorigendet.idtes Is Null and tes_caja.id = " & IdMovimiento

                    GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, True
                    '===========================================================================================================================


                    '--origen con detalle
                    nSQL = " SELECT 'Egreso' AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajaori.tc, " _
                        + vbCr + " 0 AS impdebsol, " _
                        + vbCr + " 0 AS impdebdol, " _
                        + vbCr + " IIf(tes_caja.idmon=1,tes_cajaorigendet.importe,0) AS imphabsol, " _
                        + vbCr + " IIf(tes_caja.idmon=2,tes_cajaorigendet.importe,0) AS imphabdol, " _
                        + vbCr + " 0 AS tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaorigendet.idmod, tes_cajaorigendet.corr AS iddocpro, 1 AS tipo, tes_caja.idmon,tes_cajaorigendet.importe as impreal, " _
                        + vbCr + " 6 AS ridlib, iif(ref2.ridper2 is null,0,5) AS ridtipper, ref2.ridper2 as ridper , tes_cajaorigendet.tipdoc AS rtipdoc, tes_caja.fchope AS rfchope, IIf([tes_cajaorigendet].[numser] Is Null Or [tes_cajaorigendet].[numser]='','',[tes_cajaorigendet].[numser] & '-') & [tes_cajaorigendet].[numdoc] AS rnumerodoc, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([tes_caja].[numreg],4) AS rregistro, NULL AS rglosa,tes_caja.glosa AS rglosaope ,tes_caja.idmon as ridmon, " _
                        + vbCr + " null as rtipdoc1, null as rnumerodoc1, null AS ridtipper2, null as ridper2 , null as rtipdoc2, null as rfchope2, null as rnumerodoc2 " _
                        + vbCr + " ,-1 as aplicatc, 0 as ajuste " _
                        + vbCr + " FROM ((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) INNER JOIN (((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) " _
                        + vbCr + " LEFT JOIN " _
                        + vbCr + " ( SELECT TOP 1 tes_caja.id AS idtes, tes_cajaorigendet.tipdoc AS rtipdoc2, tes_caja.fchope AS rfchope2, tes_cajaorigendet.numdoc AS rnumerodoc2, 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, mae_bancos.numruc, mae_bancos.descripcion " _
                        + vbCr + " FROM (tes_caja INNER JOIN ((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) ON tes_cajaori.idbcocta = mae_banconumcta.id " _
                        + vbCr + " WHERE (((tes_caja.tipmov) = 2) And ((tes_cajaorigendet.idtes) Is Not Null) And ((mae_banconumcta.id) Is Not Null)) and tes_cajaori.idmod in (5,6,7)and tes_caja.id = " & IdMovimiento & "" _
                        + vbCr + " ) as ref2  ON tes_caja.id = ref2.idtes " _
                        + vbCr + " WHERE tes_caja.tipmov=2 and tes_cajaorigendet.idtes Is Not Null and tes_caja.id = " & IdMovimiento

                    GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, True
                    '===========================================================================================================================

                    '--destino sin detalle
                    nSQL = " SELECT 'Egreso' AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajadestino.tc, " _
                        + vbCr + " IIf(tes_caja.idmon=1,tes_cajadestino.importe,0) AS impdebsol, " _
                        + vbCr + " IIf(tes_caja.idmon=2,tes_cajadestino.importe,0) AS impdebdol, " _
                        + vbCr + " 0 AS imphabsol, " _
                        + vbCr + " 0 AS imphabdol, " _
                        + vbCr + " 0 AS tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestino.idmod , 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,tes_cajadestino.importe as impreal, " _
                        + vbCr + " 6 AS ridlib, 0 AS ridtipper, 0 AS ridper, 0 AS rtipdoc, tes_caja.fchope AS rfchope, '' AS rnumerodoc, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([tes_caja].[numreg],4) AS rregistro, NULL AS rglosa,tes_caja.glosa AS rglosaope ,tes_caja.idmon as ridmon, " _
                        + vbCr + " NULL as rtipdoc1, NULL as rnumerodoc1, ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2 " _
                        + vbCr + " ,-1 as aplicatc, 0 as ajuste " _
                        + vbCr + " FROM (((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) INNER JOIN ((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) " _
                        + vbCr + " LEFT JOIN " _
                        + vbCr + " ( SELECT TOP 1 tes_caja.id AS idtes, tes_cajaorigendet.tipdoc AS rtipdoc2, tes_caja.fchope AS rfchope2, tes_cajaorigendet.numdoc AS rnumerodoc2, 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, mae_bancos.numruc, mae_bancos.descripcion " _
                        + vbCr + " FROM (tes_caja INNER JOIN ((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) ON tes_cajaori.idbcocta = mae_banconumcta.id " _
                        + vbCr + " WHERE (((tes_caja.tipmov) = 2) And ((tes_cajaorigendet.idtes) Is Not Null) And ((mae_banconumcta.id) Is Not Null)) and tes_cajaori.idmod in (5,6,7) and tes_caja.id = " & IdMovimiento & "" _
                        + vbCr + " ) as ref2  ON tes_caja.id = ref2.idtes " _
                        + vbCr + " WHERE tes_caja.tipmov=2 and tes_cajadestinodet.idtes Is Null and tes_caja.id = " & IdMovimiento

                    GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, True
                    '===========================================================================================================================

                    '--destino con detalle
                    nSQL = " SELECT 'Egreso' AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajadestino.tc, " _
                        + vbCr + " IIf(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) AS impdebsol, " _
                        + vbCr + " IIf(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) AS impdebdol, " _
                        + vbCr + " 0 AS imphabsol, " _
                        + vbCr + " 0 AS imphabdol, " _
                        + vbCr + " 0 AS tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod , tes_cajadestinodet.corr AS iddocpro, 2 AS tipo, tes_caja.idmon,tes_cajadestinodet.acuenta as impreal, " _
                        + vbCr + " iif(tes_cajadestinodet.iddoc=0 or tes_cajadestinodet.idmod in (3,5,6,7),6, iif(tes_cajadestinodet.idmod in (1,2),tes_cajadestinodet.idmod, iif(tes_cajadestinodet.idmod=8,9, iif(tes_cajadestinodet.idmod=9,40, iif(tes_cajadestinodet.idmod=10,999, iif(tes_cajadestinodet.idmod=11,41,0) )) ) ) ) as ridlib, " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7),tes_cajadestinodet.idtipper, iif ( tes_cajadestinodet.idmod=1,iif(tes_cajadestinodet.idori=2,con_percepcion.tipo,1 ), iif(tes_cajadestinodet.idmod in (9,10,),1, iif(tes_cajadestinodet.idmod in (2,11),1, iif(tes_cajadestinodet.idmod=4,99,-1 ) ) ) ) ) as ridtipper, " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,tes_cajadestinodet.idper, iif(tes_cajadestinodet.idmod =1, iif(tes_cajadestinodet.idori=2,con_percepcion.idcli,com_compras.idpro ) , iif(tes_cajadestinodet.idmod=2,vta_ventas.idcli, iif(tes_cajadestinodet.idmod=4,99, iif(tes_cajadestinodet.idmod=8, pla_boleta.idemp, iif(tes_cajadestinodet.idmod=9, com_honorarios.idpro,iif(tes_cajadestinodet.idmod=10,com_reembolsables.idpro,iif(tes_cajadestinodet.idmod=11,vta_gastodebito.idcli,-1 ) ) ) ) ) ) ) ) as ridper, " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,tes_cajadestinodet.tipdoc, iif(tes_cajadestinodet.idmod =1, iif(tes_cajadestinodet.idori=2,con_percepcion.tipdoc,com_compras.tipdoc ), iif(tes_cajadestinodet.idmod=2,vta_ventas.tipdoc, iif(tes_cajadestinodet.idmod=4,99, iif(tes_cajadestinodet.idmod=8, pla_boleta.iddoc, iif(tes_cajadestinodet.idmod=9, com_honorarios.tipdoc,iif(tes_cajadestinodet.idmod=10,com_reembolsables.tipdoc,iif(tes_cajadestinodet.idmod=11,vta_gastodebito.tipdoc,-1 ) ) ) ) ) ) ) ) as rtipdoc , " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,tes_cajadestinodet.fchdoc, iif(tes_cajadestinodet.idmod =1, iif(tes_cajadestinodet.idori=2,con_percepcion.fchdoc,com_compras.fchdoc ), iif(tes_cajadestinodet.idmod=2,vta_ventas.fchdoc, iif(tes_cajadestinodet.idmod=4,99, iif(tes_cajadestinodet.idmod=8, pla_boleta.fchdoc, iif(tes_cajadestinodet.idmod=9, com_honorarios.fchdoc,iif(tes_cajadestinodet.idmod=10,com_reembolsables.fchdoc,iif(tes_cajadestinodet.idmod=11,vta_gastodebito.fchemi,-1 ) ) ) ) ) ) ) ) as rfchope, " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,IIf([tes_cajadestinodet].[numser] Is Null Or [tes_cajadestinodet].[numser]='','',[tes_cajadestinodet].[numser] & '-') & [tes_cajadestinodet].[numdoc], iif(tes_cajadestinodet.idmod =1, iif(tes_cajadestinodet.idori=2,[con_percepcion]![numser] & '-' & [con_percepcion]![numdoc] ,IIf([com_compras].[numser] Is Null Or [com_compras].[numser]='','',[com_compras].[numser] & '-') & [com_compras].[numdoc] ), iif(tes_cajadestinodet.idmod=2,IIf([vta_ventas].[numser] Is Null Or [vta_ventas].[numser]='','',[vta_ventas].[numser] & '-') & [vta_ventas].[numdoc], " _
                        + vbCr + " iif(tes_cajadestinodet.idmod=4,99,iif(tes_cajadestinodet.idmod=8, IIf([pla_boleta].[numser] Is Null Or [pla_boleta].[numser]='','',[pla_boleta].[numser] & '-') & [pla_boleta].[numdoc], iif(tes_cajadestinodet.idmod=9, IIf([com_honorarios].[numser] Is Null Or [com_honorarios].[numser]='','',[com_honorarios].[numser] & '-') & [com_honorarios].[numdoc],iif(tes_cajadestinodet.idmod=10,IIf([com_reembolsables].[numser] Is Null Or [com_reembolsables].[numser]='','',[com_reembolsables].[numser] & '-') & [com_reembolsables].[numdoc],iif(tes_cajadestinodet.idmod=11,IIf([vta_gastodebito].[numser] Is Null Or [vta_gastodebito].[numser]='','',[vta_gastodebito].[numser] & '-') & [vta_gastodebito].[numdoc],-1 ) ) ) ) ) ) ) ) as rnumerodoc, " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4), iif(tes_cajadestinodet.idmod =1, iif(tes_cajadestinodet.idori=2,Left([con_percepcion].[numreg],2) & [mae_libros_3].[codsun] & Right([con_percepcion].[numreg],4),Left([com_compras].[numreg],2) & [mae_libros_1].[codsun] & Right([com_compras].[numreg],4)), iif(tes_cajadestinodet.idmod=2,Left([vta_ventas].[numreg],2) & [mae_libros_5].[codsun] & Right([vta_ventas].[numreg],4), " _
                        + vbCr + " iif(tes_cajadestinodet.idmod=4,99, iif(tes_cajadestinodet.idmod=8, Left([pla_boleta].[numreg],2) & [mae_libros_4].[codsun] & Right([pla_boleta].[numreg],4), iif(tes_cajadestinodet.idmod=9, Left([com_honorarios].[numreg],2) & [mae_libros_2].[codsun] & Right([com_honorarios].[numreg],4),iif(tes_cajadestinodet.idmod=10,Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4),iif(tes_cajadestinodet.idmod=11,Left([vta_gastodebito].[numreg],2) & [mae_libros_6].[codsun] & Right([vta_gastodebito].[numreg],4),-1 ) ) ) ) ) ) ) ) as rregistro, " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,tes_cajadestinodet.glosa, iif(tes_cajadestinodet.idmod =1, iif(tes_cajadestinodet.idori=2,con_percepcion.glosa,com_compras.glosa ), iif(tes_cajadestinodet.idmod=2,vta_ventas.glosa, iif(tes_cajadestinodet.idmod=4,99, iif(tes_cajadestinodet.idmod=8, pla_boleta.glosa, iif(tes_cajadestinodet.idmod=9, com_honorarios.glosa,iif(tes_cajadestinodet.idmod=10,com_reembolsables.glosa,iif(tes_cajadestinodet.idmod=11,vta_gastodebito.glosa,-1 ) ) ) ) ) ) ) ) as rglosa,tes_caja.glosa AS rglosaope , " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,tes_caja.idmon, iif(tes_cajadestinodet.idmod =1, iif(tes_cajadestinodet.idori=2,con_percepcion.idmon,com_compras.idmon ), iif(tes_cajadestinodet.idmod=2,vta_ventas.idmon, iif(tes_cajadestinodet.idmod=4,99, iif(tes_cajadestinodet.idmod=8, pla_boleta.idmon, iif(tes_cajadestinodet.idmod=9, com_honorarios.idmon,iif(tes_cajadestinodet.idmod=10,com_reembolsables.idmon,iif(tes_cajadestinodet.idmod=11,vta_gastodebito.idmon,-1 ) ) ) ) ) ) ) ) as ridmon, " _
                        + vbCr + " iif(tes_cajadestinodet.numerodocref is null,0,108)  as rtipdoc1,  tes_cajadestinodet.numerodocref as rnumerodoc1, ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2 " _
                        + vbCr + " ,-1 as aplicatc, 0 as ajuste "

                    nSQL = nSQL _
                        + vbCr + " FROM (((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) INNER JOIN ((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((((((((((((tes_cajadestinodet LEFT JOIN com_compras ON tes_cajadestinodet.iddoc = com_compras.id) LEFT JOIN com_honorarios ON tes_cajadestinodet.iddoc = com_honorarios.id) LEFT JOIN con_percepcion ON tes_cajadestinodet.iddoc = con_percepcion.id) LEFT JOIN pla_boleta ON tes_cajadestinodet.iddoc = pla_boleta.id) " _
                        + vbCr + " LEFT JOIN mae_libros AS mae_libros_1 ON com_compras.idlib = mae_libros_1.id) LEFT JOIN mae_libros AS mae_libros_2 ON com_honorarios.idlib = mae_libros_2.id) LEFT JOIN mae_libros AS mae_libros_3 ON con_percepcion.idlib = mae_libros_3.id) LEFT JOIN mae_libros AS mae_libros_4 ON pla_boleta.idlib = mae_libros_4.id) LEFT JOIN vta_gastodebito ON tes_cajadestinodet.iddoc = vta_gastodebito.id) LEFT JOIN com_reembolsables ON tes_cajadestinodet.iddoc = com_reembolsables.id) LEFT JOIN (vta_ventas LEFT JOIN mae_libros AS mae_libros_5 ON vta_ventas.idlib = mae_libros_5.id) ON tes_cajadestinodet.iddoc = vta_ventas.id) LEFT JOIN mae_libros AS mae_libros_6 ON vta_gastodebito.idlib = mae_libros_6.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes))  " _
                        + vbCr + " LEFT JOIN " _
                        + vbCr + " ( SELECT TOP 1 tes_caja.id AS idtes, tes_cajaorigendet.tipdoc AS rtipdoc2, tes_caja.fchope AS rfchope2, tes_cajaorigendet.numdoc AS rnumerodoc2, 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, mae_bancos.numruc, mae_bancos.descripcion " _
                        + vbCr + " FROM (tes_caja INNER JOIN ((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) ON tes_cajaori.idbcocta = mae_banconumcta.id " _
                        + vbCr + " WHERE (((tes_caja.tipmov) = 2) And ((tes_cajaorigendet.idtes) Is Not Null) And ((mae_banconumcta.id) Is Not Null)) and tes_cajaori.idmod in (5,6,7) and tes_caja.id = " & IdMovimiento & "" _
                        + vbCr + " ) as ref2  ON tes_caja.id = ref2.idtes " _
                        + vbCr + " WHERE tes_caja.tipmov=2 AND tes_cajadestinodet.idtes is not null and tes_caja.id = " & IdMovimiento

                    GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, True
                    '===========================================================================================================================

                End If


                If TipoMovimiento = 1 Then            '--ingreso
                    '--origen sin detalle

                    nSQL = "SELECT 'Ingreso' AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajaori.tc, " _
                        + vbCr + " IIf(tes_caja.idmon=1,tes_cajaori.importe,0) AS impdebsol, " _
                        + vbCr + " IIf(tes_caja.idmon=2,tes_cajaori.importe,0) AS impdebdol, " _
                        + vbCr + " 0 AS imphabsol, " _
                        + vbCr + " 0 AS imphabdol, " _
                        + vbCr + " 0 AS tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaori.idmod AS idmod, 0 AS iddocpro, 1 AS tipo, tes_caja.idmon,tes_cajaori.importe as impreal, " _
                        + vbCr + " 6 AS ridlib, 0 AS ridtipper, 0 AS ridper, 0 AS rtipdoc, tes_caja.fchope AS rfchope, '' AS rnumerodoc, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([tes_caja].[numreg],4) AS rregistro, NULL AS rglosa,tes_caja.glosa AS rglosaope ,tes_caja.idmon as ridmon, " _
                        + vbCr + " NULL as rtipdoc1, null as rnumerodoc1, null as ridtipper2,null as ridper2, null as rtipdoc2, null as rfchope2, null as rnumerodoc2 " _
                        + vbCr + " ,-1 as aplicatc, 0 as ajuste " _
                        + vbCr + " FROM ((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) INNER JOIN ((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id " _
                        + vbCr + " WHERE tes_caja.tipmov=1 and tes_cajaorigendet.idtes Is Null and tes_caja.id = " & IdMovimiento


                    GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                    '===========================================================================================================================

                    '--origen con detalle

                    nSQL = " SELECT 'Ingreso' AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajaori.tc, " _
                        + vbCr + " IIf(tes_caja.idmon=1,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) AS impdebsol, " _
                        + vbCr + " IIf(tes_caja.idmon=2,iif(tes_cajaorigendet.acuenta = 0 ,tes_cajaorigendet.importe,tes_cajaorigendet.acuenta),0) AS impdebdol, " _
                        + vbCr + " 0 AS imphabsol, " _
                        + vbCr + " 0 AS imphabdol, " _
                        + vbCr + " 0 AS tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_origen.id AS idorides, tes_cajaorigendet.idmod, tes_cajaorigendet.corr AS iddocpro, 1 AS tipo, tes_caja.idmon,tes_cajaorigendet.importe as impreal, " _
                        + vbCr + " iif(tes_cajaorigendet.iddoc=0 or tes_cajaorigendet.idmod in (3,5,6,7),6, iif(tes_cajaorigendet.idmod in (2,4,11),tes_cajaorigendet.idmod, 6 ) ) as ridlib, " _
                        + vbCr + " iif(tes_cajaorigendet.idmod in (3,5,6,7),tes_cajaorigendet.idtipper, iif ( tes_cajaorigendet.idmod in (2,4,11),2,0) ) as ridtipper, " _
                        + vbCr + " iif(tes_cajaorigendet.idmod in (3,5,6,7) or tes_cajaorigendet.iddoc =0,tes_cajaorigendet.idper, iif(tes_cajaorigendet.idmod =2, vta_ventas.idcli , iif(tes_cajaorigendet.idmod=11, vta_gastodebito.idcli,iif(tes_cajaorigendet.idmod=4,88, -1 ) ) ) ) as ridper, " _
                        + vbCr + " iif(tes_cajaorigendet.idmod in (3,5,6,7) or tes_cajaorigendet.iddoc =0,tes_cajaorigendet.tipdoc, iif(tes_cajaorigendet.idmod =2, vta_ventas.tipdoc, iif(tes_cajaorigendet.idmod=11, vta_gastodebito.tipdoc,iif(tes_cajaorigendet.idmod=4,88, -1 ) ) ) ) as rtipdoc , " _
                        + vbCr + " iif(tes_cajaorigendet.idmod in (3,5,6,7) or tes_cajaorigendet.iddoc =0,tes_cajaorigendet.fchdoc, iif(tes_cajaorigendet.idmod =2, vta_ventas.fchdoc, iif(tes_cajaorigendet.idmod=11, vta_gastodebito.fchemi, iif(tes_cajaorigendet.idmod=4,88,NULL ) ) ) ) as rfchope , " _
                        + vbCr + " iif(tes_cajaorigendet.idmod in (3,5,6,7) or tes_cajaorigendet.iddoc =0,IIf([tes_cajaorigendet].[numser] Is Null Or [tes_cajaorigendet].[numser]='','',[tes_cajaorigendet].[numser] & '-') & [tes_cajaorigendet].[numdoc], iif(tes_cajaorigendet.idmod =2, IIf([vta_ventas].[numser] Is Null Or [vta_ventas].[numser]='','',[vta_ventas].[numser] & '-') & [vta_ventas].[numdoc], iif(tes_cajaorigendet.idmod=11, IIf([vta_gastodebito].[numser] Is Null Or [vta_gastodebito].[numser]='','',[vta_gastodebito].[numser] & '-') & [vta_gastodebito].[numdoc] , iif(tes_cajaorigendet.idmod=4,'--', '**' ) ) ) ) as rnumerodoc, " _
                        + vbCr + " iif(tes_cajaorigendet.idmod in (3,5,6,7) or tes_cajaorigendet.iddoc =0,Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4), iif(tes_cajaorigendet.idmod =2, Left([vta_ventas].[numreg],2) & [mae_libros_1].[codsun] & Right([vta_ventas].[numreg],4), iif(tes_cajaorigendet.idmod=11, Left([vta_gastodebito].[numreg],2) & [mae_libros_2].[codsun] & Right([vta_gastodebito].[numreg],4),iif(tes_cajaorigendet.idmod=4,88, null ) ) ) ) as rregistro, " _
                        + vbCr + " iif(tes_cajaorigendet.idmod in (3,5,6,7) or tes_cajaorigendet.iddoc =0,tes_cajaorigendet.glosa, iif(tes_cajaorigendet.idmod =2, vta_ventas.glosa, iif(tes_cajaorigendet.idmod=11, vta_gastodebito.glosa, iif(tes_cajaorigendet.idmod=4,88,'--' ) ) ) ) as rglosa , tes_caja.glosa AS rglosaope , " _
                        + vbCr + " iif(tes_cajaorigendet.idmod in (3,5,6,7) or tes_cajaorigendet.iddoc =0,tes_cajaorigendet.idmon, iif(tes_cajaorigendet.idmod =2, vta_ventas.idmon, iif(tes_cajaorigendet.idmod=11, vta_gastodebito.idmon, iif(tes_cajaorigendet.idmod=4,88,-1 ) ) ) ) as ridmon , " _
                        + vbCr + " iif(tes_cajaorigendet.numerodocref is null,0,108) as rtipdoc1, tes_cajaorigendet.numerodocref as rnumerodoc1, ridtipper2, ref2.ridper2 , ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2 " _
                        + vbCr + " ,-1 as aplicatc, 0 as ajuste "

                    nSQL = nSQL _
                        + vbCr + " FROM ( ((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) ON tes_caja.id = tes_cajaori.idtes) INNER JOIN ((((tes_cajaorigendet LEFT JOIN vta_ventas ON tes_cajaorigendet.iddoc = vta_ventas.id) LEFT JOIN vta_gastodebito ON tes_cajaorigendet.iddoc = vta_gastodebito.id) LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas.idlib = mae_libros_1.id) LEFT JOIN mae_libros AS mae_libros_2 ON vta_gastodebito.idlib = mae_libros_2.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes) ) " _
                        + vbCr + " LEFT JOIN " _
                        + vbCr + " ( SELECT TOP 1 tes_caja.id AS idtes, tes_cajaorigendet.tipdoc AS rtipdoc2, tes_caja.fchope AS rfchope2, tes_cajaorigendet.numdoc AS rnumerodoc2, 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, mae_bancos.numruc, mae_bancos.descripcion " _
                        + vbCr + " FROM (tes_caja INNER JOIN ((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) ON tes_cajaori.idbcocta = mae_banconumcta.id " _
                        + vbCr + " WHERE (((tes_caja.tipmov) = 1) And ((tes_cajaorigendet.idtes) Is Not Null) And ((mae_banconumcta.id) Is Not Null)) and tes_cajaori.idmod in (5,6,7) and tes_caja.id = " & IdMovimiento & "" _
                        + vbCr + " ) as ref2  ON tes_caja.id = ref2.idtes " _
                        + vbCr + " WHERE tes_caja.tipmov=1 and tes_cajaorigendet.idtes Is Not Null and tes_caja.id = " & IdMovimiento


                    GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                    '===========================================================================================================================

                    '--destino sin detalle
                    nSQL = " SELECT 'Ingreso' AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajadestino.tc, " _
                        + vbCr + " 0 AS impdebsol, " _
                        + vbCr + " 0 AS impdebdol, " _
                        + vbCr + " IIf(tes_caja.idmon=1,tes_cajadestino.importe,0) AS imphabsol, " _
                        + vbCr + " IIf(tes_caja.idmon=2,tes_cajadestino.importe,0) AS imphabdol, " _
                        + vbCr + " 0 AS tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestino.idmod , 0 AS iddocpro, 2 AS tipo, tes_caja.idmon,tes_cajadestino.importe as impreal, " _
                        + vbCr + " 6 AS ridlib, 0 AS ridtipper, 0 AS ridper, 0 AS rtipdoc, tes_caja.fchope AS rfchope, '' AS rnumerodoc, Left([tes_caja].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([tes_caja].[numreg],4) AS rregistro, NULL AS rglosa,tes_caja.glosa AS rglosaope ,tes_caja.idmon as ridmon, " _
                        + vbCr + " NULL as rtipdoc1, NULL as rnumerodoc1, ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2 " _
                        + vbCr + " ,-1 as aplicatc, 0 as ajuste " _
                        + vbCr + " FROM (((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) INNER JOIN ((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) LEFT JOIN tes_cajadestinodet ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) " _
                        + vbCr + " LEFT JOIN " _
                        + vbCr + " ( SELECT TOP 1 tes_caja.id AS idtes, tes_cajaorigendet.tipdoc AS rtipdoc2, tes_caja.fchope AS rfchope2, tes_cajaorigendet.numdoc AS rnumerodoc2, 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, mae_bancos.numruc, mae_bancos.descripcion " _
                        + vbCr + " FROM (tes_caja INNER JOIN ((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) ON tes_cajaori.idbcocta = mae_banconumcta.id " _
                        + vbCr + " WHERE (((tes_caja.tipmov) = 1) And ((tes_cajaorigendet.idtes) Is Not Null) And ((mae_banconumcta.id) Is Not Null)) and tes_cajaori.idmod in (5,6,7) and tes_caja.id = " & IdMovimiento & "" _
                        + vbCr + " ) as ref2  ON tes_caja.id = ref2.idtes " _
                        + vbCr + " WHERE tes_caja.tipmov=1 and tes_cajadestinodet.idtes Is Null and tes_caja.id = " & IdMovimiento


                    GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                    '===========================================================================================================================

                    '--destino con detalle
                    nSQL = " SELECT 'Ingreso' AS Tipo1, tes_caja.numreg, mae_moneda.simbolo, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                        + vbCr + " iif(tes_caja.fchreg is null,null,Year(tes_caja.fchreg)) AS anotra, iif(tes_caja.fchreg is null,null,Month(tes_caja.fchreg)) AS idmes, 6 AS idlib, iif(tes_caja.numreg is null,null,Right(tes_caja.numreg,4)) AS numasi, tes_caja.id as idmov, con_planctas.id AS idcue, 0 AS iddoc, 0 AS correlativo, tes_cajadestino.tc, " _
                        + vbCr + " 0 AS impdebsol, " _
                        + vbCr + " 0 AS impdebdol, " _
                        + vbCr + " IIf(tes_caja.idmon=1,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) AS imphabsol, " _
                        + vbCr + " IIf(tes_caja.idmon=2,iif(tes_cajadestinodet.acuenta = 0 ,tes_cajadestinodet.importe,tes_cajadestinodet.acuenta),0) AS imphabdol, " _
                        + vbCr + " 0 AS tipmov, 0 AS prodiv, tes_caja.fchreg AS fchasi, tes_caja.fchope AS fchdoc, tes_destino.id AS idorides, tes_cajadestinodet.idmod, tes_cajadestinodet.corr AS iddocpro, 2 AS tipo, tes_caja.idmon,(tes_cajadestinodet.importe+tes_cajadestinodet.acuenta) as impreal, " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,6, iif(tes_cajadestinodet.idmod=4,37, iif(tes_cajadestinodet.idmod in (1,2),tes_cajadestinodet.idmod, iif(tes_cajadestinodet.idmod=8,9, iif(tes_cajadestinodet.idmod=9,40, iif(tes_cajadestinodet.idmod=10,999, iif(tes_cajadestinodet.idmod=11,41,iif(tes_cajadestinodet.idmod=19,42,0) ) )) ) ) ) ) as ridlib,  " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,tes_cajadestinodet.idtipper, iif(tes_cajadestinodet.idmod = 19,5, iif(tes_cajadestinodet.idmod in (1,9,10),1, iif(tes_cajadestinodet.idmod in (2,11,4),2, 99 ) ) ) ) as ridtipper,  " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,tes_cajadestinodet.idper, iif(tes_cajadestinodet.idmod in (1,8,9,10),0, iif(tes_cajadestinodet.idmod=2,vta_ventas.idcli,iif(tes_cajadestinodet.idmod=11,vta_gastodebito.idcli, iif(tes_cajadestinodet.idmod=4,5555,iif(tes_cajadestinodet.idmod=19,222222,-1 ) ) ) ) ) ) as ridper,  " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,tes_cajadestinodet.tipdoc,iif(tes_cajadestinodet.idmod in (1,8,9,10),0, iif(tes_cajadestinodet.idmod=4,let_planilla.tipdoc, iif(tes_cajadestinodet.idmod=2,vta_ventas.tipdoc,iif(tes_cajadestinodet.idmod=11,vta_gastodebito.tipdoc,-1 ) ) ) ) ) as rtipdoc ,  " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,tes_cajadestinodet.fchdoc,iif(tes_cajadestinodet.idmod in (1,8,9,10),NULL, iif(tes_cajadestinodet.idmod=4,let_planilla.fchemi, iif(tes_cajadestinodet.idmod=2,vta_ventas.fchdoc,iif(tes_cajadestinodet.idmod=11,vta_gastodebito.fchemi,NULL ) ) ) ) ) as rfchope,  " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,IIf([tes_cajadestinodet].[numser] Is Null Or [tes_cajadestinodet].[numser]='','',[tes_cajadestinodet].[numser] & '-') & [tes_cajadestinodet].[numdoc], " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (1,8,9,10), NULL, iif(tes_cajadestinodet.idmod=4,let_planilla.numdoc, iif(tes_cajadestinodet.idmod=2,IIf([vta_ventas].[numser] Is Null Or [vta_ventas].[numser]='','',[vta_ventas].[numser] & '-') & [vta_ventas].[numdoc] ,iif(tes_cajadestinodet.idmod=11, IIf([vta_gastodebito].[numser] Is Null Or [vta_gastodebito].[numser]='','',[vta_gastodebito].[numser] & '-') & [vta_gastodebito].[numdoc] , '' ) ) ) ) ) as rnumerodoc,  " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4),iif(tes_cajadestinodet.idmod in (1,8,9,10),NULL, iif(tes_cajadestinodet.idmod=4,Left([let_planilla].[numreg],2) & [mae_libros_7].[codsun] & Right([let_planilla].[numreg],4), iif(tes_cajadestinodet.idmod=2,Left([vta_ventas].[numreg],2) & [mae_libros_5].[codsun] & Right([vta_ventas].[numreg],4),iif(tes_cajadestinodet.idmod=11,Left([vta_gastodebito].[numreg],2) & [mae_libros_6].[codsun] & Right([vta_gastodebito].[numreg],4),'' ) ) ) ) ) as rregistro,  " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.iddoc =0,tes_cajadestinodet.glosa,iif(tes_cajadestinodet.idmod in (1,8,9,10),'', iif(tes_cajadestinodet.idmod=4,let_planilla.glosa, iif(tes_cajadestinodet.idmod=2,vta_ventas.glosa,iif(tes_cajadestinodet.idmod=11,vta_gastodebito.glosa,'' ) ) ) ) ) as rglosa,tes_caja.glosa AS rglosaope , " _
                        + vbCr + " iif(tes_cajadestinodet.idmod in (3,5,6,7) or tes_cajadestinodet.idmod =0,tes_caja.idmon,iif(tes_cajadestinodet.idmod in (1,8,9,10),0, iif(tes_cajadestinodet.idmod=4,let_planilla.idmon, iif(tes_cajadestinodet.idmod=2,vta_ventas.idmon,iif(tes_cajadestinodet.idmod=11,vta_gastodebito.idmon,0 ) ) ) ) ) as ridmon, " _
                        + vbCr + " iif(tes_cajadestinodet.numerodocref is null,0,108) as rtipdoc1,  tes_cajadestinodet.numerodocref as rnumerodoc1, ref2.rtipdoc2, ref2.rfchope2, ref2.rnumerodoc2, ridtipper2, ref2.ridper2 " _
                        + vbCr + " ,-1 as aplicatc, 0 as ajuste "

                    nSQL = nSQL _
                        + vbCr + " FROM ( (((((tes_caja LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) INNER JOIN (((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes) INNER JOIN (((tes_cajadestinodet LEFT JOIN vta_gastodebito ON tes_cajadestinodet.iddoc = vta_gastodebito.id) LEFT JOIN (vta_ventas  " _
                        + vbCr + " LEFT JOIN mae_libros AS mae_libros_5 ON vta_ventas.idlib = mae_libros_5.id) ON tes_cajadestinodet.iddoc = vta_ventas.id) LEFT JOIN mae_libros AS mae_libros_6 ON vta_gastodebito.idlib = mae_libros_6.id) ON (tes_cajadestino.idtes = tes_cajadestinodet.idtes) AND (tes_cajadestino.iddes = tes_cajadestinodet.iddes)) ON tes_caja.id = tes_cajadestino.idtes) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) LEFT JOIN let_planilla ON tes_cajadestinodet.iddoc = let_planilla.id) LEFT JOIN mae_libros AS mae_libros_7 ON let_planilla.idlib = mae_libros_7.id ) " _
                        + vbCr + " LEFT JOIN " _
                        + vbCr + " ( SELECT TOP 1 tes_caja.id AS idtes, tes_cajaorigendet.tipdoc AS rtipdoc2, tes_caja.fchope AS rfchope2, tes_cajaorigendet.numdoc AS rnumerodoc2, 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, mae_bancos.numruc, mae_bancos.descripcion " _
                        + vbCr + " FROM (tes_caja INNER JOIN ((tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) INNER JOIN tes_cajaorigendet ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) ON tes_cajaori.idbcocta = mae_banconumcta.id " _
                        + vbCr + " WHERE (((tes_caja.tipmov) = 1) And ((tes_cajaorigendet.idtes) Is Not Null) And ((mae_banconumcta.id) Is Not Null)) and tes_cajaori.idmod in (5,6,7) and tes_caja.id = " & IdMovimiento & "" _
                        + vbCr + " ) as ref2  ON tes_caja.id = ref2.idtes " _
                        + vbCr + " WHERE tes_caja.tipmov=1 AND tes_cajadestinodet.idtes is not null and tes_caja.id = " & IdMovimiento

                    GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                    '===========================================================================================================================

                End If





                '***********************************************************************************************************************************
                '***********************************************************************************************************************************


            Case 8 '--Canjes de Facturas

                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE con_canjes SET con_canjes.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((con_canjes.id)=" & IdMovimiento & "))"
                '***********************************************************************************

                '---de compra
                nSQL = "SELECT con_canjes.idpro, 0 AS tipdoc, con_canjes.numser, con_canjes.numdoc, con_canjes.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " IIf(con_canjes.fchreg Is Null,Null,Year(con_canjes.fchreg)) AS anotra, IIf(con_canjes.fchreg Is Null,Null,Month(con_canjes.fchreg)) AS idmes, 8 AS idlib, con_canjes.id AS idmov, mae_documentocta.idcuen AS idcue, IIf(con_canjes.numreg Is Null,Null,Right(con_canjes.numreg,4)) AS numasi, com_compras.id AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                    + vbCr + " IIf(con_canjes.idmon=1,con_canjes.impcan,0) AS impdebsol, " _
                    + vbCr + " IIf(con_canjes.idmon=2,con_canjes.impcan,0) AS impdebdol, " _
                    + vbCr + " 0 AS imphabsol, " _
                    + vbCr + " 0 AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, con_canjes.fchreg AS fchasi, con_canjes.fchemi AS fchdoc, 0 AS idorides, 0 AS idmod, con_canjesdet.iddoc AS iddocpro, 0 AS tipo, con_canjes.idmon, " _
                    + vbCr + " 1 AS ridlib, 1 AS ridtipper, com_compras.idpro AS ridper, com_compras.tipdoc AS rtipdoc, com_compras.fchdoc AS rfchope, IIf(com_compras.numser Is Null Or com_compras.numser='','',com_compras.numser & '-') & com_compras.numdoc AS rnumerodoc, Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4) AS rregistro, com_compras.glosa AS rglosa, '' AS rglosaope, com_compras.idmon AS ridmon, " _
                    + vbCr + " 0 AS rtipdoc1, '' AS rnumerodoc1, 1 AS ridtipper2, con_canjes.idpro AS ridper2, 0 AS rtipdoc2, con_canjes.fchemi AS rfchope2, con_canjes.numser & '-' & con_canjes.numdoc AS rnumerodoc2 " _
                    + vbCr + " ,0 as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM ((con_canjes LEFT JOIN con_tc ON con_canjes.fchemi = con_tc.fecha) LEFT JOIN mae_libros AS mae_libros_1 ON con_canjes.idlib = mae_libros_1.id) INNER JOIN (con_canjesdet INNER JOIN ((com_compras INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (com_compras.tipdoc = mae_documentocta.iddoc) AND (com_compras.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON con_canjesdet.iddoc = com_compras.id) ON con_canjes.id = con_canjesdet.idcan " _
                    + vbCr + " WHERE (((con_canjesdet.tipo)=2) AND ((mae_documentocta.tipope)=0) AND ((con_canjes.id)=" & IdMovimiento & ")) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                '===========================================================================================================================

                '--de venta

                nSQL = "SELECT con_canjes.idpro, 0 AS tipdoc, con_canjes.numser, con_canjes.numdoc, con_canjes.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " IIf(con_canjes.fchreg Is Null,Null,Year(con_canjes.fchreg)) AS anotra, IIf(con_canjes.fchreg Is Null,Null,Month(con_canjes.fchreg)) AS idmes, 8 AS idlib, con_canjes.id AS idmov, mae_documentocta.idcuen AS idcue, IIf(con_canjes.numreg Is Null,Null,Right(con_canjes.numreg,4)) AS numasi, vta_ventas.id AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                    + vbCr + " 0 AS impdebsol, " _
                    + vbCr + " 0 AS impdebdol, " _
                    + vbCr + " IIf(con_canjes.idmon=1,con_canjes.impcan,0) AS imphabsol, " _
                    + vbCr + " IIf(con_canjes.idmon=2,con_canjes.impcan,0) AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, con_canjes.fchreg AS fchasi, con_canjes.fchemi AS fchdoc, 0 AS idorides, 0 AS idmod, con_canjesdet.iddoc AS iddocpro, 0 AS tipo, con_canjes.idmon, " _
                    + vbCr + " 2 AS ridlib, 2 AS ridtipper, vta_ventas.idcli AS ridper, vta_ventas.tipdoc AS rtipdoc, vta_ventas.fchdoc AS rfchope, IIf(vta_ventas.numser Is Null Or vta_ventas.numser='','',vta_ventas.numser & '-') & vta_ventas.numdoc AS rnumerodoc, Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4) AS rregistro, vta_ventas.glosa AS rglosa, '' AS rglosaope, vta_ventas.idmon AS ridmon, " _
                    + vbCr + " 0 AS rtipdoc1, '' AS rnumerodoc1, 2 AS ridtipper2, con_canjes.idcli AS ridper2, 0 AS rtipdoc2, con_canjes.fchemi AS rfchope2, con_canjes.numser & '-' & con_canjes.numdoc AS rnumerodoc2 " _
                    + vbCr + " ,0 as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM ((con_canjes LEFT JOIN con_tc ON con_canjes.fchemi = con_tc.fecha) LEFT JOIN mae_libros AS mae_libros_1 ON con_canjes.idlib = mae_libros_1.id) INNER JOIN (con_canjesdet INNER JOIN ((vta_ventas INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) ON con_canjesdet.iddoc = vta_ventas.id) ON con_canjes.id = con_canjesdet.idcan " _
                    + vbCr + " WHERE (((con_canjesdet.tipo)=1) AND ((mae_documentocta.tipope)=-1) AND ((con_canjes.id)=" & IdMovimiento & ")) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                '===========================================================================================================================

            Case 9 '--Planilla de Pago

                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE pla_boleta SET pla_boleta.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((pla_boleta.id)=" & IdMovimiento & "))"
                '***********************************************************************************

                '--cargar el documento debe

                nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion AS nomcta, 'H' AS Origen, 'Doc' AS Tipo1, " _
                    + vbCr + " Year(pla_boleta.fchreg) AS anotra, Month(pla_boleta.fchreg) AS idmes, pla_boleta.idlib, pla_boleta.id AS idmov, mae_documentocta.idcuen AS idcue, Right(pla_boleta.numreg,4) AS numasi, pla_boleta.id as iddoc , 0 AS correlativo, con_tc.impven AS tc, " _
                    + vbCr + " 0 AS impdebsol, " _
                    + vbCr + " 0 AS impdebdol, " _
                    + vbCr + " IIf(pla_boleta.idmon=2,0,pla_boleta.imptot) AS imphabsol, " _
                    + vbCr + " IIf(pla_boleta.idmon=1,0,pla_boleta.imptot) AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, pla_boleta.fchreg AS fchasi, pla_boleta.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, pla_boleta.idmon, 9 AS ridlib, " _
                    + vbCr + " 3 AS ridtipper, pla_boleta.idemp AS ridper, pla_boleta.iddoc AS rtipdoc, pla_boleta.fchdoc AS rfchope, pla_boleta.numser & '-' & pla_boleta.numdoc AS rnumerodoc, Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4) AS rregistro, '' AS rglosa, pla_boleta.glosa AS rglosaope, pla_boleta.idmon AS ridmon, 0 AS rtipdoc1, '' AS rnumerodoc1, 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, Null AS rfchope2, '' AS rnumerodoc2 " _
                    + vbCr + " , 0 as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM ((pla_boleta INNER JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (pla_boleta.idmon = mae_documentocta.idmon) AND (pla_boleta.iddoc = mae_documentocta.iddoc)) LEFT JOIN con_tc ON pla_boleta.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON pla_boleta.idlib = mae_libros.id " _
                    + vbCr + " WHERE (((pla_boleta.id)=" & IdMovimiento & " )) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

                '--si es apertura salir
                If IdMes = 0 Then GoTo Salto_Apertura:

                '===========================================================================================================================

                '--detalle conceptos debe
                nSQL = "SELECT con_planctas.id AS idcta, con_planctas.cuenta, con_planctas.descripcion AS nomcta, 'D' AS Origen, 'Detalle' AS Tipo1, Sum(pla_boletadet.imptot) AS total, " _
                    + vbCr + " Year(pla_boleta.fchreg) AS anotra, Month(pla_boleta.fchreg) AS idmes, pla_boleta.idlib, pla_boleta.id AS idmov, pla_concepto.idctadeb AS idcue, Right(pla_boleta.numreg,4) AS numasi, 0 AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                    + vbCr + " Sum(IIf(pla_boleta.idmon=2,0,pla_boletadet.imptot)) AS impdebsol, " _
                    + vbCr + " Sum(IIf(pla_boleta.idmon=1,0,pla_boletadet.imptot)) AS impdebdol, " _
                    + vbCr + " 0 AS imphabsol, " _
                    + vbCr + " 0 AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, pla_boleta.fchreg AS fchasi, pla_boleta.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, pla_boleta.idmon, 9 AS ridlib, " _
                    + vbCr + " 3 AS ridtipper, pla_boleta.idemp AS ridper, pla_boleta.iddoc AS rtipdoc, pla_boleta.fchdoc AS rfchope, pla_boleta.numser & '-' & pla_boleta.numdoc AS rnumerodoc, Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4) AS rregistro, '' AS rglosa, pla_boleta.glosa AS rglosaope, pla_boleta.idmon AS ridmon, 0 AS rtipdoc1, '' AS rnumerodoc1, 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, Null AS rfchope2, '' AS rnumerodoc2 " _
                    + vbCr + " , 0 as aplicatc, 0 as ajuste  " _
                    + vbCr + " FROM (con_planctas RIGHT JOIN pla_concepto ON con_planctas.id = pla_concepto.idctadeb) INNER JOIN (((pla_boleta LEFT JOIN mae_libros ON pla_boleta.idlib = mae_libros.id) LEFT JOIN con_tc ON pla_boleta.fchdoc = con_tc.fecha) INNER JOIN pla_boletadet ON pla_boleta.id = pla_boletadet.idbol) ON pla_concepto.id = pla_boletadet.idcpto " _
                    + vbCr + " WHERE (((pla_boleta.id)=" & IdMovimiento & " )) " _
                    + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion, 'Detalle', Year(pla_boleta.fchreg), Month(pla_boleta.fchreg), pla_boleta.idlib, pla_boleta.id, pla_concepto.idctadeb, Right(pla_boleta.numreg,4), con_tc.impven, pla_boleta.fchreg, pla_boleta.fchdoc, pla_boleta.idmon, pla_boleta.idemp, pla_boleta.iddoc, pla_boleta.fchdoc, pla_boleta.numser & '-' & pla_boleta.numdoc, Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4), pla_boleta.glosa, pla_boleta.idmon, pla_concepto.idctadeb, pla_concepto.aplanilla, pla_boleta.id " _
                    + vbCr + " HAVING (((pla_concepto.idctadeb) Is Not Null And (pla_concepto.idctadeb)<>0) AND ((pla_concepto.aplanilla)=-1) ) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                '===========================================================================================================================

               '--detalle concepto haber
               '--Modificar consulta para considerar la cuenta asignada al regimen de pensiones
            nSQL = "SELECT det.*, IIf([con_planctas].[id] Is Null,[det].[idcta1],[con_planctas].[id]) AS IdCta, IIf([con_planctas].[id] Is Null,[det].[cuenta1],[con_planctas].[cuenta]) AS cuenta, IIf([con_planctas].[id] Is Null,[det].[nomcta1],[con_planctas].[descripcion]) AS nomcta, [IdCta] AS idcue " _
                    + vbCr + " FROM " _
                        + vbCr + "( SELECT con_planctas.id AS idcta1, con_planctas.cuenta as cuenta1, con_planctas.descripcion AS nomcta1, 'H' AS Origen, 'Detalle' AS Tipo1, (pla_boletadet.imptot) AS total, " _
                        + vbCr + " Year(pla_boleta.fchreg) AS anotra, Month(pla_boleta.fchreg) AS idmes, pla_boleta.idlib, pla_boleta.id AS idmov, pla_concepto.idctahab AS idcue1, Right(pla_boleta.numreg,4) AS numasi, 0 AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                        + vbCr + " 0 AS impdebsol, " _
                        + vbCr + " 0 AS impdebdol, " _
                        + vbCr + " IIf([pla_boleta].[idmon]=2,0,[pla_boletadet].[imptot]) AS imphabsol, " _
                        + vbCr + " IIf([pla_boleta].[idmon]=1,0,[pla_boletadet].[imptot]) AS imphabdol, " _
                        + vbCr + " 0 AS tipmov, 0 AS prodiv, pla_boleta.fchreg AS fchasi, pla_boleta.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, pla_boleta.idmon, 9 AS ridlib, " _
                        + vbCr + " 3 AS ridtipper, pla_boleta.idemp AS ridper, pla_boleta.iddoc AS rtipdoc, pla_boleta.fchdoc AS rfchope, pla_boleta.numser & '-' & pla_boleta.numdoc AS rnumerodoc, Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4) AS rregistro, '' AS rglosa, pla_boleta.glosa AS rglosaope, pla_boleta.idmon AS ridmon, 0 AS rtipdoc1, '' AS rnumerodoc1, 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, Null AS rfchope2, '' AS rnumerodoc2 " _
                        + vbCr + " , 0 AS aplicatc, 0 as ajuste, " _
                        + vbCr + " pla_boleta.idregpen, pla_boletadet.idcpto " _
                        + vbCr + " FROM (pla_concepto LEFT JOIN con_planctas ON pla_concepto.idctahab = con_planctas.id) INNER JOIN (((pla_boleta LEFT JOIN con_tc ON pla_boleta.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON pla_boleta.idlib = mae_libros.id) INNER JOIN pla_boletadet ON pla_boleta.id = pla_boletadet.idbol) ON pla_concepto.id = pla_boletadet.idcpto " _
                        + vbCr + " WHERE (((pla_boleta.id)=" & IdMovimiento & ") AND ((pla_concepto.idctahab) Is Not Null And (pla_concepto.idctahab)<>0) AND ((pla_concepto.aplanilla)=-1)) " _
                        + vbCr + " ) as det " _
                        + vbCr + " LEFT JOIN (pla_conceptoregpen LEFT JOIN con_planctas ON pla_conceptoregpen.idcuenta = con_planctas.id) ON (det.idcpto = pla_conceptoregpen.idcpto) AND (det.idregpen = pla_conceptoregpen.idregpen) "


'               nSQL = "SELECT con_planctas.id AS idcta, con_planctas.cuenta, con_planctas.descripcion AS nomcta, 'H' AS Origen, 'Detalle' AS Tipo1, Sum(pla_boletadet.imptot) AS total, " _
                    + vbCr + " Year(pla_boleta.fchreg) AS anotra, Month(pla_boleta.fchreg) AS idmes, pla_boleta.idlib, pla_boleta.id AS idmov, pla_concepto.idctahab AS idcue, Right(pla_boleta.numreg,4) AS numasi, 0 AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                    + vbCr + " 0 AS impdebsol, " _
                    + vbCr + " 0 AS impdebdol, " _
                    + vbCr + " Sum(IIf([pla_boleta].[idmon]=2,0,[pla_boletadet].[imptot])) AS imphabsol, " _
                    + vbCr + " Sum(IIf([pla_boleta].[idmon]=1,0,[pla_boletadet].[imptot])) AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, pla_boleta.fchreg AS fchasi, pla_boleta.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, pla_boleta.idmon, 9 AS ridlib, " _
                    + vbCr + " 3 AS ridtipper, pla_boleta.idemp AS ridper, pla_boleta.iddoc AS rtipdoc, pla_boleta.fchdoc AS rfchope, pla_boleta.numser & '-' & pla_boleta.numdoc AS rnumerodoc, Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4) AS rregistro, '' AS rglosa, pla_boleta.glosa AS rglosaope, pla_boleta.idmon AS ridmon, 0 AS rtipdoc1, '' AS rnumerodoc1, 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, Null AS rfchope2, '' AS rnumerodoc2 " _
                    + vbCr + " , 0 as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM (pla_concepto LEFT JOIN con_planctas ON pla_concepto.idctahab = con_planctas.id) INNER JOIN (((pla_boleta LEFT JOIN con_tc ON pla_boleta.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON pla_boleta.idlib = mae_libros.id) INNER JOIN pla_boletadet ON pla_boleta.id = pla_boletadet.idbol) ON pla_concepto.id = pla_boletadet.idcpto " _
                    + vbCr + " WHERE (((pla_boleta.id)=" & IdMovimiento & " )) " _
                    + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion, 'Detalle', Year(pla_boleta.fchreg), Month(pla_boleta.fchreg), pla_boleta.idlib, pla_boleta.id, pla_concepto.idctahab, Right(pla_boleta.numreg,4), con_tc.impven, pla_boleta.fchreg, pla_boleta.fchdoc, pla_boleta.idmon, pla_boleta.idemp, pla_boleta.iddoc, pla_boleta.fchdoc, pla_boleta.numser & '-' & pla_boleta.numdoc, Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4), pla_boleta.glosa, pla_boleta.idmon, pla_concepto.idctahab, pla_concepto.aplanilla, pla_boleta.id " _
                    + vbCr + " HAVING (((pla_concepto.idctahab) Is Not Null And (pla_concepto.idctahab)<>0) AND ((pla_concepto.aplanilla)=-1)) "



                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                '===========================================================================================================================
            '---
            '--
            Case 37 '--Emision de Letras
                '--22/09/09:: Crear script para generar asiento
                '--16/10/09:: Considerar en asiento de la retencion automatica el campo let_letra.aplicaret = -1, cuando el cliente sea agente de retencion y el usuario lo considere necesario
                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE let_letra SET let_letra.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((let_letra.id)=" & IdMovimiento & "))"
                '***********************************************************************************

                '--DEBE (DETALLE DE LETRAS)
                nSQL = "SELECT let_letra.idclipro, let_letra.tipdoc, let_letradet.numser, let_letradet.numdoc, let_letra.idmon, let_letra.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc,  " _
                    + vbCr + " Year([let_letra].[fchreg]) AS anotra,let_letra.idmes, let_letra.idlib, let_letra.id AS idmov, mae_letra.idcuenven AS idcue,IIf([let_letra].[numreg] Is Null,Null,Right([let_letra].[numreg],4)) AS numasi, let_letradet.corr AS iddoc, let_letradet.corr AS correlativo, IIf([let_letra].[tc]=0,[con_tc].[impven],[let_letra].[tc]) AS tc, " _
                    + vbCr + " IIf(let_letra.idmon=1,let_letradet.implet,0) AS impdebsol, " _
                    + vbCr + " IIf(let_letra.idmon=2,let_letradet.implet,0) AS impdebdol, " _
                    + vbCr + " 0 AS imphabsol, " _
                    + vbCr + " 0 AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, let_letra.fchreg AS fchasi, let_letra.fchini as fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, " _
                    + vbCr + " 37 AS ridlib, 2 AS ridtipper, let_letra.idclipro AS ridper, let_letra.tipdoc AS rtipdoc, let_letra.fchini AS rfchope, " _
                    + vbCr + " [let_letra].[ano]  & ' ' &  [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS rnumerodoc, let_letra.glosa AS rglosa,'' as rglosaope, " _
                    + vbCr + " Left([let_letra].[numreg],2) & [mae_libros].[codsun] & Right([let_letra].[numreg],4) AS rregistro,let_letra.idmon as ridmon, " _
                    + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, let_letra.numerodocref AS rnumerodoc1, " _
                    + vbCr + " 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, NULL AS rfchope2, '' AS rnumerodoc2, " _
                    + vbCr + " IIF(let_letra.tc = 0 Or let_letra.tc Is Null, 0, -1) as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM (((((((let_letra LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id) LEFT JOIN mae_letra ON let_letra.idmon = mae_letra.idmon) LEFT JOIN con_planctas ON mae_letra.idcuenven = con_planctas.id) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) LEFT JOIN con_tc ON let_letra.fchini = con_tc.fecha) INNER JOIN let_letradet ON let_letra.id = let_letradet.idlet) LEFT JOIN mae_docreferencia ON let_letra.idtipdocref = mae_docreferencia.id " _
                    + vbCr + " WHERE (((let_letra.id)=" & IdMovimiento & ")) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                '===========================================================================================================================


            '--DEBE(SI ES AGENTE RETENEDOR
            '--el % de retencion sera fijo en esta consulta;
            '--la retencion se tomara por la suma de (interes+portes+igv) de la tabla let_letradet
            '--si hay retenciones parciales que hacen referencia a una nota de debito, solo se toma la primera retencion
            '--si no hay nota de debito =>> no hay retencion por lo tanto se creara con un numero pendiente de retencion
            '--el codigo del documento de retencion sera fijo en esta consulta con valor 20
            '--se considera agente retenedor cuando el campo "mae_cliente.ageret" tiene valor "-1"

            nSQL = "SELECT let_letra.idclipro, let_letra.tipdoc, ret.numser, ret.numdoc, let_letra.idmon, let_letra.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, let_letra.idmes, " _
                    + vbCr + " Year([let_letra].[fchreg]) AS anotra, let_letra.idlib, let_letra.id AS idmov, mae_documentocta.idcuen AS idcue,IIf([let_letra].[numreg] Is Null,Null,Right([let_letra].[numreg],4)) AS numasi, ret.idret AS iddoc, 0 AS correlativo,IIf([let_letra].[tc]=0,[con_tc].[impven],[let_letra].[tc]) AS tc, " _
                    + vbCr + " IIf([let_letra].[idmon]=1,detletnd.imptot*0.06,0) AS impdebsol, " _
                    + vbCr + " IIf([let_letra].[idmon]=2,detletnd.imptot*0.06,0) AS impdebdol, " _
                    + vbCr + " 0 AS imphabsol, " _
                    + vbCr + " 0 AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, let_letra.fchreg AS fchasi, let_letra.fchini AS fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, " _
                    + vbCr + " IIf(ret.idlet Is Null,37,5) AS ridlib, 2 AS ridtipper, IIf(ret.idlet Is Null,let_letra.idclipro,ret.idcli) AS ridper, 20 AS rtipdoc, IIf(ret.idlet Is Null,let_letra.fchini,ret.fchdoc) AS rfchope, " _
                    + vbCr + " IIf(ret.idlet Is Null,'PENDIENTE',ret.numerodoc) AS rnumerodoc, IIf(ret.idlet Is Null,let_letra.glosa,ret.glosa) AS rglosa, '' AS rglosaope, " _
                    + vbCr + " IIf(ret.idlet Is Null,Left([let_letra].[numreg],2) & [mae_libros].[codsun] & Right([let_letra].[numreg],4),ret.registro) AS rregistro, IIf(ret.idlet Is Null,let_letra.idmon,ret.idmon) AS ridmon, " _
                    + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, let_letra.numerodocref AS rnumerodoc1, " _
                    + vbCr + " IIf(ret.idlet Is Null,0,2) AS ridtipper2, IIf(ret.idlet Is Null,0,let_letra.idclipro) AS ridper2, IIf(ret.idlet Is Null,0,let_letra.tipdoc) AS rtipdoc2, IIf(ret.idlet Is Null,Null,let_letra.fchini) AS rfchope2, '' AS rnumerodoc2, " _
                    + vbCr + " IIf(let_letra.tc = 0 Or let_letra.tc Is Null, 0, -1) as aplicatc, 0 as ajuste "

                nSQL = nSQL _
                    + vbCr + " FROM (( (mae_cliente INNER JOIN (((((let_letra LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) LEFT JOIN con_tc ON let_letra.fchini = con_tc.fecha) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON let_letra.idmon = mae_documentocta.idmon) ON mae_cliente.id = let_letra.idclipro) LEFT JOIN mae_docreferencia ON let_letra.idtipdocref = mae_docreferencia.id ) " _
                    + vbCr + " LEFT JOIN " _
                    + vbCr + " (SELECT TOP 1 con_retencion.id as idret, let_letradoc.idlet, Left([con_retencion].[numreg],2) & [mae_libros].[codsun] & Right([con_retencion].[numreg],4) AS registro, con_retencion.idpro AS idcli, con_retencion.iddoc AS tipdoc, con_retencion.fchemi AS fchdoc, con_retenciondet.impret AS imptot, con_retencion.idmon, [con_retencion].[numser] & '-' & [con_retencion].[numdoc] AS numerodoc, con_retencion.glosa, con_retencion.numser, con_retencion.numdoc " _
                    + vbCr + " FROM (mae_libros INNER JOIN ((con_retencion INNER JOIN con_retenciondet ON con_retencion.id = con_retenciondet.id) INNER JOIN let_letradoc ON con_retenciondet.iddoc = let_letradoc.iddoc) ON mae_libros.id = con_retencion.idlib) INNER JOIN let_letra ON let_letradoc.idlet = let_letra.id " _
                    + vbCr + " WHERE let_letradoc.idlet=" & IdMovimiento & " and let_letradoc.idmod=2 ) AS ret  ON let_letra.id = ret.idlet ) " _
                    + vbCr + " INNER JOIN " _
                    + vbCr + " (SELECT let_letradet.idlet, Sum(let_letradet.impporte)+Sum(let_letradet.impinteres)+Sum(let_letradet.impigv) AS imptot FROM let_letradet GROUP BY let_letradet.idlet " _
                    + vbCr + " HAVING let_letradet.idlet = " & IdMovimiento & " And ((Sum(let_letradet.impporte) + Sum(let_letradet.impinteres) + Sum(let_letradet.impigv)) <> 0) " _
                    + vbCr + " ) AS detletnd ON let_letra.id = detletnd.idlet " _
                    + vbCr + " WHERE (((let_letra.id)=" & IdMovimiento & ") AND ((mae_documentocta.tipope)=-1) AND ((mae_documentocta.iddoc)=8) AND ((mae_cliente.ageret)=-1)) and let_letra.aplicaret=-1 "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                '===========================================================================================================================

            '--HABER
            '--Proforma ==>>> SE ESTA UTILIZANDO LA CUENTA FIJA

            '--ID      NUMERO  DESCRIPCION
            '--==      ======= ==================================
            Dim IdCtaMn As Integer
            Dim IdCtaMe As Integer
            '--185     1220103 Anticipos de Clientes en Letras MN
            '--1997    1220104 Anticipos de Clientes en Letras ME
            Select Case AnoTra
                Case "2010"
                    IdCtaMn = 185
                    IdCtaMe = 1997
                Case "2011"
                    IdCtaMn = 233
                    IdCtaMe = 234
                Case Else

            End Select

            nSQL = "SELECT let_letra.idclipro, let_letra.tipdoc, vta_proforma.numser, vta_proforma.numdoc, let_letra.idmon, let_letra.numreg, IIf(let_letra.idmon=1,'1220103','1220104') AS ctanum, IIf(let_letra.idmon=1,'Anticipos de Clientes en Letras MN','Anticipos de Clientes en Letras ME') AS ctadesc, let_letra.idmes, " _
                    + vbCr + " Year(let_letra.fchreg) AS anotra, let_letra.idlib, let_letra.id AS idmov, IIf(let_letra.idmon=1," & IdCtaMn & " , " & IdCtaMe & ") AS idcue, IIf(let_letra.numreg Is Null,Null,Right(let_letra.numreg,4)) AS numasi, vta_proforma.id AS iddoc, 0 AS correlativo, " _
                    + vbCr + " IIf(let_letra.tc=0,con_tc.impven,let_letra.tc) AS tc, " _
                    + vbCr + " 0 AS impdebsol, " _
                    + vbCr + " 0 AS impdebdol, " _
                    + vbCr + " IIf(let_letra.idmon=1,let_letradoc.impfin,0) AS imphabsol, " _
                    + vbCr + " IIf(let_letra.idmon=2,let_letradoc.impfin,0) AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, let_letra.fchreg AS fchasi, let_letra.fchini AS fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, " _
                    + vbCr + " 37 AS ridlib, 2 AS ridtipper, vta_proforma.idcli AS ridper, vta_proforma.tipdoc AS rtipdoc, let_letra.fchini AS rfchope, " _
                    + vbCr + " IIf(vta_proforma.numser Is Null,'',vta_proforma.numser & '') & vta_proforma.numdoc AS rnumerodoc, vta_proforma.glosa AS rglosa, let_letra.glosa AS rglosaope, Left(let_letra.numreg,2) & mae_libros.codsun & Right(let_letra.numreg,4) AS rregistro, let_letra.idmon AS ridmon, " _
                    + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, let_letra.numerodocref AS rnumerodoc1, " _
                    + vbCr + " 2 AS ridtipper2, let_letra.idclipro AS ridper2, let_letra.tipdoc AS rtipdoc2, let_letra.fchini AS rfchope2, '' AS rnumerodoc2, " _
                    + vbCr + " IIf(let_letra.tc = 0 Or let_letra.tc Is Null, 0, -1) as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM (((((let_letra LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) LEFT JOIN con_tc ON let_letra.fchini = con_tc.fecha) INNER JOIN (let_letradoc INNER JOIN vta_proforma ON let_letradoc.iddoc = vta_proforma.id) ON let_letra.id = let_letradoc.idlet) LEFT JOIN mae_docreferencia ON let_letra.idtipdocref = mae_docreferencia.id " _
                    + vbCr + " WHERE (((let_letra.id)=" & IdMovimiento & ") AND ((let_letradoc.idmod)=12)) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                '===========================================================================================================================
            '--HABER FACTURACION
                nSQL = "SELECT let_letra.idclipro, let_letra.tipdoc, vta_ventas.numser, vta_ventas.numdoc, let_letra.idmon, let_letra.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, let_letra.idmes, " _
                    + vbCr + " Year(let_letra.fchreg) AS anotra, let_letra.idlib, let_letra.id AS idmov, mae_documentocta.idcuen AS idcue,IIf(let_letra.numreg Is Null,Null,Right(let_letra.numreg,4)) AS numasi, vta_ventas.id AS iddoc, 0 AS correlativo,IIf(let_letra.tc=0,con_tc.impven,let_letra.tc) AS tc, " _
                    + vbCr + " 0 AS impdebsol, " _
                    + vbCr + " 0 AS impdebdol, " _
                    + vbCr + " IIf(let_letra.idmon=1,IIf(vta_ventas.idmon=1,let_letradoc.impfin,let_letradoc.impfin*tc)) AS imphabsol, " _
                    + vbCr + " IIf(let_letra.idmon=2,IIf(vta_ventas.idmon=2,let_letradoc.impfin,let_letradoc.impfin/tc)) AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, let_letra.fchreg AS fchasi, let_letra.fchini AS fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, " _
                    + vbCr + " 2 AS ridlib, 2 AS ridtipper, vta_ventas.idcli AS ridper, vta_ventas.tipdoc AS rtipdoc,vta_ventas.fchdoc AS rfchope, " _
                    + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS rnumerodoc, vta_ventas.glosa AS rglosa, '' AS rglosaope, " _
                    + vbCr + " Left(vta_ventas.numreg,2) & mae_libros_1.codsun & Right(vta_ventas.numreg,4) AS rregistro, vta_ventas.idmon AS ridmon, " _
                    + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, let_letra.numerodocref AS rnumerodoc1, " _
                    + vbCr + " let_letra.id AS ridtipper2, let_letra.idclipro AS ridper2, let_letra.tipdoc AS rtipdoc2, let_letra.fchini AS rfchope2, '' AS rnumerodoc2,IIf(let_letra.tc=0 Or let_letra.tc Is Null,0,-1) as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM ((((let_letra LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) LEFT JOIN con_tc ON let_letra.fchini = con_tc.fecha) INNER JOIN ((((vta_ventas LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas.idlib = mae_libros_1.id) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon)) INNER JOIN let_letradoc ON vta_ventas.id = let_letradoc.iddoc) LEFT JOIN mae_docreferencia ON vta_ventas.idtipdocref = mae_docreferencia.id) ON let_letra.id = let_letradoc.idlet " _
                    + vbCr + " WHERE (((let_letra.id)=" & IdMovimiento & ") AND ((mae_documentocta.tipope)=-1) AND ((let_letradoc.idmod)=2)) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                '===========================================================================================================================

            '--HABER NOTA DEBITO
                '--si letra no tiene referencia de nota de debito; => el asiento tendra un numero de nota debito como pendiente
                '--solo se agregara el asiento cuando no se registre manualmente la nota debito
                '--para obtener la cta contable se esta relacionando let_letra con mae_documentocta
                nSQL = "SELECT let_letra.idclipro, let_letra.tipdoc, '' as numser, '' as numdoc, let_letra.idmon, let_letra.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, let_letra.idmes, " _
                    + vbCr + " Year(let_letra.fchreg) AS anotra, let_letra.idlib, let_letra.id AS idmov, mae_documentocta.idcuen AS idcue,IIf(let_letra.numreg Is Null,Null,Right(let_letra.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo,IIf(let_letra.tc=0,con_tc.impven,let_letra.tc) AS tc, " _
                    + vbCr + " 0 AS impdebsol, " _
                    + vbCr + " 0 AS impdebdol, " _
                    + vbCr + " IIf(let_letra.idmon=1,detletnd.imptot) AS imphabsol, " _
                    + vbCr + " IIf(let_letra.idmon=2,detletnd.imptot) AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, let_letra.fchreg AS fchasi, let_letra.fchini AS fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, " _
                    + vbCr + " 37 AS ridlib, 2 AS ridtipper, let_letra.idclipro AS ridper, 8 AS rtipdoc, let_letra.fchini AS rfchope, " _
                    + vbCr + " 'PENDIENTE' AS rnumerodoc, let_letra.glosa AS rglosa, '' AS rglosaope, " _
                    + vbCr + " Left(let_letra.numreg,2) & mae_libros.codsun & Right(let_letra.numreg,4) AS rregistro, let_letra.idmon AS ridmon, " _
                    + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, let_letra.numerodocref AS rnumerodoc1, " _
                    + vbCr + " 2 AS ridtipper2, let_letra.idclipro AS ridper2, let_letra.tipdoc AS rtipdoc2, let_letra.fchini AS rfchope2, '' AS rnumerodoc2,IIf(let_letra.tc=0 Or let_letra.tc Is Null,0,-1) as aplicatc, 0 as ajuste "

                nSQL = nSQL _
                    + vbCr + " FROM (( (((((let_letra LEFT JOIN mae_moneda ON let_letra.idmon = mae_moneda.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) LEFT JOIN con_tc ON let_letra.fchini = con_tc.fecha) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON let_letra.idmon = mae_documentocta.idmon) LEFT JOIN mae_docreferencia ON let_letra.idtipdocref = mae_docreferencia.id  ) " _
                    + vbCr + " Left Join " _
                    + vbCr + " (SELECT vta_ventas.id as idvta,let_letradoc.idlet, vta_ventas.idlib, 2 AS ridtipper, vta_ventas.idcli, vta_ventas.tipdoc, vta_ventas.fchdoc, vta_ventas.numser, vta_ventas.numdoc, vta_ventas.numser & '-' & vta_ventas.numdoc AS numerodoc, vta_ventas.glosa, Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4) AS registro, vta_ventas.idmon, vta_ventas.imptotdoc , let_letradoc.impfin " _
                    + vbCr + " FROM (let_letradoc INNER JOIN vta_ventas ON let_letradoc.iddoc = vta_ventas.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
                    + vbCr + " Where (((let_letradoc.idmod) = 2)) And vta_ventas.tipdoc = 8 " _
                    + vbCr + " ) AS vta ON let_letra.id = vta.idlet ) " _
                    + vbCr + " Left Join " _
                    + vbCr + " (SELECT let_letradet.idlet, Sum(let_letradet.impporte)+Sum(let_letradet.impinteres)+Sum(let_letradet.impigv) AS imptot FROM let_letradet GROUP BY let_letradet.idlet HAVING (  (Sum(let_letradet.impporte)+Sum(let_letradet.impinteres)+Sum(let_letradet.impigv)) <>0) " _
                    + vbCr + " ) AS detletnd  ON let_letra.id = detletnd.idlet " _
                    + vbCr + " WHERE  (mae_documentocta.tipope=-1 and mae_documentocta.iddoc=8 and let_letra.id = " & IdMovimiento & ") and (  ((detletnd.imptot -  vta.impfin not between -1 and 1) and detletnd.idlet is not null)  or (vta.idlet is null and detletnd.idlet is not null) ) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto
                '===========================================================================================================================



            Case 38 '--Entregas a Rendir a Cuenta

            Case 39 '--Rendición de Cuentas

            Case 40 '--honorarios

                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE com_honorarios SET com_honorarios.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((com_honorarios.id)=" & IdMovimiento & "))"
                '***********************************************************************************


                '--encabezado
                nSQL = "SELECT com_honorarios.idpro, com_honorarios.tipdoc, com_honorarios.numser, com_honorarios.numdoc, com_honorarios.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " iif(com_honorarios.fchreg is null,null,Year(com_honorarios.fchreg)) AS anotra, iif(com_honorarios.fchreg is null,null,Month(com_honorarios.fchreg)) AS idmes, 1 AS idlib, com_honorarios.id AS idmov, mae_documentocta.idcuen AS idcue, iif(com_honorarios.numreg is null,null,Right(com_honorarios.numreg,4)) AS numasi, com_honorarios.id as iddoc, 0 AS correlativo, IIF(com_honorarios.tc =0 or com_honorarios.tc is null ,con_tc.impven,com_honorarios.tc) AS tc, " _
                    + vbCr + " 0 AS impdebsol, " _
                    + vbCr + " 0 AS impdebdol, " _
                    + vbCr + " IIf(com_honorarios.idmon=1,com_honorarios.imptot,0) AS imphabsol, " _
                    + vbCr + " IIf(com_honorarios.idmon=2,com_honorarios.imptot,0) AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, com_honorarios.fchreg AS fchasi, com_honorarios.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,com_honorarios.idmon,  " _
                    + vbCr + " 40 AS ridlib, 1 AS ridtipper, com_honorarios.idpro AS ridper, com_honorarios.tipdoc AS rtipdoc, com_honorarios.fchdoc AS rfchope, IIf([com_honorarios].[numser] Is Null Or [com_honorarios].[numser]='','',[com_honorarios].[numser] & '-') & [com_honorarios].[numdoc] AS rnumerodoc, com_honorarios.glosa AS rglosa,'' as rglosaope, " _
                    + vbCr + " Left([com_honorarios].[numreg],2) & [mae_libros].[codsun] & Right([com_honorarios].[numreg],4) AS rregistro,com_honorarios.idmon as ridmon, " _
                    + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, com_honorarios.numerodocref AS rnumerodoc1, " _
                    + vbCr + " 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, NULL AS rfchope2, '' AS rnumerodoc2, " _
                    + vbCr + " IIF(com_honorarios.tc =0 or com_honorarios.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM (((com_honorarios LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (com_honorarios.tipdoc = mae_documentocta.iddoc) AND (com_honorarios.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN mae_docreferencia ON com_honorarios.idtipdocref = mae_docreferencia.id " _
                    + vbCr + " WHERE (((mae_documentocta.tipope)=0)) and com_honorarios.id = " & IdMovimiento

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

                '--si es apertura salir
                If IdMes = 0 Then GoTo Salto_Apertura:
                '===========================================================================================================================

                '--igv
                nSQL = " SELECT  com_honorarios.idpro, com_honorarios.tipdoc, com_honorarios.numser, com_honorarios.numdoc, com_honorarios.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " iif(com_honorarios.fchreg is null,null,Year(com_honorarios.fchreg)) AS anotra, iif(com_honorarios.fchreg is null,null,Month(com_honorarios.fchreg)) AS idmes, 1 AS idlib, com_honorarios.id AS idmov, con_planctas.id AS idcue, iif(com_honorarios.numreg is null,null,Right(com_honorarios.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, IIF(com_honorarios.tc =0 or com_honorarios.tc is null ,con_tc.impven,com_honorarios.tc) AS tc, " _
                    + vbCr + " 0 AS impdebsol, " _
                    + vbCr + " 0 AS impdebdol, " _
                    + vbCr + " IIf(com_honorarios.idmon=1,com_honorarios.impigv,0) AS imphabsol, " _
                    + vbCr + " IIf(com_honorarios.idmon=2,com_honorarios.impigv,0) AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, com_honorarios.fchreg AS fchasi, com_honorarios.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,com_honorarios.idmon, " _
                    + vbCr + " 40 AS ridlib, 1 AS ridtipper, com_honorarios.idpro AS ridper, com_honorarios.tipdoc AS rtipdoc, com_honorarios.fchdoc AS rfchope, IIf([com_honorarios].[numser] Is Null Or [com_honorarios].[numser]='','',[com_honorarios].[numser] & '-') & [com_honorarios].[numdoc] AS rnumerodoc, com_honorarios.glosa AS rglosa,'' as rglosaope,  " _
                    + vbCr + " Left([com_honorarios].[numreg],2) & [mae_libros].[codsun] & Right([com_honorarios].[numreg],4) AS rregistro,com_honorarios.idmon as ridmon, " _
                    + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, com_honorarios.numerodocref AS rnumerodoc1, " _
                    + vbCr + " 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, NULL AS rfchope2, '' AS rnumerodoc2, " _
                    + vbCr + " IIF(com_honorarios.tc =0 or com_honorarios.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM (((mae_prov INNER JOIN ((mae_documento INNER JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) INNER JOIN (com_honorarios LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) ON mae_documento.id = com_honorarios.tipdoc) ON mae_prov.id = com_honorarios.idpro) INNER JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN mae_docreferencia ON com_honorarios.idtipdocref = mae_docreferencia.id " _
                    + vbCr + " WHERE com_honorarios.id = " & IdMovimiento

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

                '--detalle
                nSQL = " SELECT com_honorarios.idpro, com_honorarios.tipdoc, com_honorarios.numser, com_honorarios.numdoc, com_honorarios.numreg,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " iif(com_honorarios.fchreg is null,null,Year(com_honorarios.fchreg)) AS anotra, iif(com_honorarios.fchreg is null,null,Month(com_honorarios.fchreg)) AS idmes, 1 AS idlib, com_honorarios.id AS idmov, con_planctas.id AS idcue, iif(com_honorarios.numreg is null,null,Right(com_honorarios.numreg,4)) AS numasi, 0 AS iddoc, 0 AS correlativo, IIF(com_honorarios.tc =0 or com_honorarios.tc is null ,con_tc.impven,com_honorarios.tc) AS tc, " _
                    + vbCr + " Sum(IIf([com_honorarios].[idmon]=1,[com_honorariosdet].[imptot],0)) AS impdebsol, " _
                    + vbCr + " Sum(IIf([com_honorarios].[idmon]=2,[com_honorariosdet].[imptot],0)) AS impdebdol, " _
                    + vbCr + " 0 AS imphabsol, " _
                    + vbCr + " 0 AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, com_honorarios.fchreg AS fchasi, com_honorarios.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo,com_honorarios.idmon, " _
                    + vbCr + " 40 AS ridlib, 1 AS ridtipper, com_honorarios.idpro AS ridper, com_honorarios.tipdoc AS rtipdoc, com_honorarios.fchdoc AS rfchope, IIf([com_honorarios].[numser] Is Null Or [com_honorarios].[numser]='','',[com_honorarios].[numser] & '-') & [com_honorarios].[numdoc] AS rnumerodoc, com_honorarios.glosa AS rglosa,'' as rglosaope, " _
                    + vbCr + " Left([com_honorarios].[numreg],2) & [mae_libros].[codsun] & Right([com_honorarios].[numreg],4) AS rregistro,com_honorarios.idmon as ridmon, " _
                    + vbCr + " mae_docreferencia.iddoc AS rtipdoc1, com_honorarios.numerodocref AS rnumerodoc1, " _
                    + vbCr + " 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, NULL AS rfchope2, '' AS rnumerodoc2, " _
                    + vbCr + " IIF(com_honorarios.tc =0 or com_honorarios.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM (((com_honorarios LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) INNER JOIN ((alm_inventario INNER JOIN com_honorariosdet ON alm_inventario.id = com_honorariosdet.iditem) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id) ON com_honorarios.id = com_honorariosdet.idhon) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN mae_docreferencia ON com_honorarios.idtipdocref = mae_docreferencia.id " _
                    + vbCr + " WHERE com_honorarios.id = " & IdMovimiento & "" _
                    + vbCr + " GROUP BY com_honorarios.idpro, com_honorarios.tipdoc, com_honorarios.numser, com_honorarios.numdoc, com_honorarios.numreg, con_planctas.cuenta, con_planctas.descripcion, com_honorarios.id, con_planctas.id,   con_tc.impven,com_honorarios.tc, com_honorarios.fchreg, com_honorarios.fchdoc, com_honorarios.idmon, com_honorarios.glosa, mae_libros.codsun,  mae_docreferencia.iddoc,com_honorarios.numerodocref "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

            '===========================================================================================================================
        Case 41 '--Liquidacion Gasto Debito LGD

                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE vta_gastodebito SET vta_gastodebito.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((vta_gastodebito.id)=" & IdMovimiento & "))"
                '***********************************************************************************

                '--cabecera
                nSQL = "SELECT vta_gastodebito.idcli, vta_gastodebito.tipdoc, vta_gastodebito.numser, vta_gastodebito.numdoc, vta_gastodebito.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " IIf(vta_gastodebito.fchreg Is Null,Null,Year(vta_gastodebito.fchreg)) AS anotra, IIf(vta_gastodebito.fchreg Is Null,Null,Month(vta_gastodebito.fchreg)) AS idmes, 1 AS idlib, vta_gastodebito.id AS idmov, con_planctas.id AS idcue, IIf(vta_gastodebito.numreg Is Null,Null,Right(vta_gastodebito.numreg,4)) AS numasi, 0 as iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                    + vbCr + " IIf(vta_gastodebito.idmon=1,vta_gastodebito.imptot,0) AS impdebsol, " _
                    + vbCr + " IIf(vta_gastodebito.idmon=2,vta_gastodebito.imptot,0) AS impdebdol, " _
                    + vbCr + " 0 AS imphabsol, " _
                    + vbCr + " 0 AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, vta_gastodebito.fchreg AS fchasi, vta_gastodebito.fchemi AS fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, vta_gastodebito.idmon, " _
                    + vbCr + " 41 AS ridlib, 2 AS ridtipper, vta_gastodebito.idcli AS ridper, vta_gastodebito.tipdoc AS rtipdoc, vta_gastodebito.fchemi AS rfchope, IIf([vta_gastodebito].[numser] Is Null Or [vta_gastodebito].[numser]='','',[vta_gastodebito].[numser] & '-') & [vta_gastodebito].[numdoc] AS rnumerodoc, vta_gastodebito.glosa AS rglosa, vta_gastodebito.glosa AS rglosaope, Left([vta_gastodebito].[numreg],2) & [mae_libros].[codsun] & Right([vta_gastodebito].[numreg],4) AS rregistro,vta_gastodebito.idmon as ridmon, " _
                    + vbCr + " mae_docreferencia.iddoc as rtipdoc1, vta_gastodebito.numerodocref as rnumerodoc1, null as ridtipper2,null as ridper2, null as rtipdoc2, null as rfchope2, null as rnumerodoc2 " _
                    + vbCr + " ,IIF(vta_gastodebito.tc =0 or vta_gastodebito.tc is null ,0,-1) as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM (((vta_gastodebito LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) INNER JOIN (con_planctas INNER JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_gastodebito.tipdoc = mae_documentocta.iddoc) AND (vta_gastodebito.idmon = mae_documentocta.idmon)) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id) LEFT JOIN mae_docreferencia ON vta_gastodebito.idtipdocref = mae_docreferencia.id " _
                    + vbCr + " WHERE (((vta_gastodebito.id)=" & IdMovimiento & ")) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

                '--si es apertura salir
                If IdMes = 0 Then GoTo Salto_Apertura:

                    '===========================================================================================================================

                nSQL = " SELECT vta_gastodebito.id,vta_gastodebito.idcli, vta_gastodebito.tipdoc, vta_gastodebito.numser, vta_gastodebito.numdoc, vta_gastodebito.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " IIf(vta_gastodebito.fchreg Is Null,Null,Year(vta_gastodebito.fchreg)) AS anotra, IIf(vta_gastodebito.fchreg Is Null,Null,Month(vta_gastodebito.fchreg)) AS idmes, 1 AS idlib, vta_gastodebito.id AS idmov, con_planctas.id AS idcue, IIf(vta_gastodebito.numreg Is Null,Null,Right(vta_gastodebito.numreg,4)) AS numasi, vta_gastodebitodet.iddoc AS iddoc, 0 AS correlativo, con_tc.impven AS tc, " _
                    + vbCr + " 0 AS impdebsol, 0 AS impdebdol, " _
                    + vbCr + " IIf(vta_gastodebito.idmon=1,vta_gastodebitodet.impacue,0) AS imphabsol, " _
                    + vbCr + " IIf(vta_gastodebito.idmon=2,vta_gastodebitodet.impacue,0) AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, vta_gastodebito.fchreg AS fchasi, vta_gastodebito.fchemi AS fchdoc, 0 AS idorides, 0 AS idmod,  0 AS iddocpro, 0 AS tipo, vta_gastodebito.idmon, " _
                    + vbCr + " IIf(vta_gastodebitodet.idmod =10,999,IIf(vta_gastodebitodet.idmod=1,1,IIf(vta_gastodebitodet.idmod=2,2,IIf(vta_gastodebitodet.idmod=9,40,IIf(vta_gastodebitodet.idmod=11,41,-1))))) AS ridlib, " _
                    + vbCr + " IIF(vta_gastodebitodet.tipreg=-1,vta_gastodebitodet.idtipper,iif(vta_gastodebitodet.idmod In (1,10,9),1,IIf(vta_gastodebitodet.idmod In (2,11),2,-1)) ) AS ridtipper, " _
                    + vbCr + " IIf(vta_gastodebitodet.idmod=0 Or vta_gastodebitodet.tipreg=-1,vta_gastodebitodet.idper,IIf(vta_gastodebitodet.idmod=1,com_compras.idpro,IIf(vta_gastodebitodet.idmod=2,vta_ventas.idcli,IIf(vta_gastodebitodet.idmod=10,com_reembolsables.idpro,IIf(vta_gastodebitodet.idmod=9,com_honorarios.idpro,IIf(vta_gastodebitodet.idmod=11,vta_gastodebito.idcli,-1)))))) AS ridper, " _
                    + vbCr + " vta_gastodebitodet.tipdoc AS rtipdoc, " _
                    + vbCr + " IIf(vta_gastodebitodet.idmod=0 Or vta_gastodebitodet.tipreg=-1,vta_gastodebitodet.fchdoc,IIf(vta_gastodebitodet.idmod=1,com_compras.fchdoc,IIf(vta_gastodebitodet.idmod=2,vta_ventas.fchdoc,IIf(vta_gastodebitodet.idmod=10,com_reembolsables.fchdoc,IIf(vta_gastodebitodet.idmod=9,com_honorarios.fchdoc,IIf(vta_gastodebitodet.idmod=11,vta_gastodebito.fchemi,-1)))))) AS rfchope, " _
                    + vbCr + " iif(vta_gastodebitodet.idmod=0 or vta_gastodebitodet.tipreg=-1,IIf(vta_gastodebitodet.numser Is Null Or vta_gastodebitodet.numser='','',vta_gastodebitodet.numser & '-') & vta_gastodebitodet.numdoc , iif(vta_gastodebitodet.idmod=1,IIf(com_compras.numser Is Null Or com_compras.numser='','',com_compras.numser & '-') & com_compras.numdoc,iif(vta_gastodebitodet.idmod=2,IIf(vta_ventas.numser Is Null Or vta_ventas.numser='','',vta_ventas.numser & '-') & vta_ventas.numdoc, " _
                    + vbCr + " iif(vta_gastodebitodet.idmod=10,IIf(com_reembolsables.numser Is Null Or com_reembolsables.numser='','',com_reembolsables.numser & '-') & com_reembolsables.numdoc,iif(vta_gastodebitodet.idmod=9,IIf(com_honorarios.numser Is Null Or com_honorarios.numser='','',com_honorarios.numser & '-') & com_honorarios.numdoc ,iif(vta_gastodebitodet.idmod=11,IIf(vta_gastodebito_1.numser Is Null Or vta_gastodebito_1.numser='','',vta_gastodebito_1.numser & '-') & vta_gastodebito_1.numdoc,-1)))  ) ) ) AS rnumerodoc, " _
                    + vbCr + " iif(vta_gastodebitodet.idmod=0 or vta_gastodebitodet.tipreg=-1,vta_gastodebitodet.glosa, iif(vta_gastodebitodet.idmod=1,com_compras.glosa,iif(vta_gastodebitodet.idmod=2,vta_ventas.glosa,iif(vta_gastodebitodet.idmod=10,com_reembolsables.glosa,iif(vta_gastodebitodet.idmod=9,com_honorarios.glosa,iif(vta_gastodebitodet.idmod=11,vta_gastodebito_1.glosa,-1)))  ) ) ) AS rglosa, vta_gastodebito.glosa AS rglosaope, " _
                    + vbCr + " iif(vta_gastodebitodet.idmod=0 or vta_gastodebitodet.tipreg=-1,Left(vta_gastodebito.numreg,2) & mae_libros_4.codsun & Right(vta_gastodebito.numreg,4), iif(vta_gastodebitodet.idmod=1,Left(com_compras.numreg,2) & mae_libros_1.codsun & Right(com_compras.numreg,4),iif(vta_gastodebitodet.idmod=2,Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4),iif(vta_gastodebitodet.idmod=10,'',iif(vta_gastodebitodet.idmod=9,Left(com_honorarios.numreg,2) & mae_libros_2.codsun & Right(com_honorarios.numreg,4)  ,iif(vta_gastodebitodet.idmod=11,Left(vta_gastodebito_1.numreg,2) & mae_libros_3.codsun & Right(vta_gastodebito_1.numreg,4),-1)))  ) ) ) AS rregistro, " _
                    + vbCr + " iif(vta_gastodebitodet.idmod=0 or vta_gastodebitodet.tipreg=-1,vta_gastodebitodet.idmon, iif(vta_gastodebitodet.idmod=1,com_compras.idmon,iif(vta_gastodebitodet.idmod=2,vta_ventas.idmon,iif(vta_gastodebitodet.idmod=10,com_reembolsables.idmon,iif(vta_gastodebitodet.idmod=9,com_honorarios.idmon,iif(vta_gastodebitodet.idmod=11,vta_gastodebito_1.idmon,-1)))  ) ) ) AS ridmon, " _
                    + vbCr + " IIf(vta_gastodebitodet.idmod=0 Or vta_gastodebitodet.tipreg=-1,vta_gastodebitodet.idtipdocref,IIf(vta_gastodebitodet.idmod=1,com_compras.idtipdocref,IIf(vta_gastodebitodet.idmod=2,vta_ventas.idtipdocref,IIf(vta_gastodebitodet.idmod=10,108,IIf(vta_gastodebitodet.idmod=9,0,IIf(vta_gastodebitodet.idmod=11,vta_gastodebito.idtipdocref,108)))))) as rtipdoc1, " _
                    + vbCr + " IIf(vta_gastodebitodet.idmod=0 Or vta_gastodebitodet.tipreg=-1,vta_gastodebitodet.numerodocref,IIf(vta_gastodebitodet.idmod=1,com_compras.numerodocref,IIf(vta_gastodebitodet.idmod=2,vta_ventas.numerodocref,IIf(vta_gastodebitodet.idmod=10,com_reembolsables.numerodocref,IIf(vta_gastodebitodet.idmod=9,'',IIf(vta_gastodebitodet.idmod=11,vta_gastodebito.numerodocref,'--')))))) as rnumerodoc1, " _
                    + vbCr + " cab.ridtipper2,cab.ridper2, cab.rtipdoc2, cab.rfchope2, cab.rnumerodoc2  " _
                    + vbCr + " ,IIF(vta_gastodebito.tc =0 or vta_gastodebito.tc is null ,0,-1) as aplicatc, 0 as ajuste "

                    nSQL = nSQL _
                    + vbCr + " FROM ( ((((((((((((vta_gastodebito LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) INNER JOIN vta_gastodebitodet ON vta_gastodebito.id = vta_gastodebitodet.idlgd) LEFT JOIN vta_ventas ON vta_gastodebitodet.iddoc = vta_ventas.id) LEFT JOIN com_compras ON vta_gastodebitodet.iddoc = com_compras.id) LEFT JOIN com_honorarios ON vta_gastodebitodet.iddoc = com_honorarios.id) LEFT JOIN com_reembolsables ON vta_gastodebitodet.iddoc = com_reembolsables.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_libros AS mae_libros_1 ON com_compras.idlib = mae_libros_1.id)  " _
                    + vbCr + " LEFT JOIN mae_libros AS mae_libros_2 ON com_honorarios.idlib = mae_libros_2.id) LEFT JOIN vta_gastodebito AS vta_gastodebito_1 ON vta_gastodebitodet.iddoc = vta_gastodebito_1.id) LEFT JOIN mae_libros AS mae_libros_3 ON vta_gastodebito_1.idlib = mae_libros_3.id) LEFT JOIN mae_libros AS mae_libros_4 ON vta_gastodebito.idlib = mae_libros_4.id) LEFT JOIN con_planctas ON vta_gastodebitodet.idcuen = con_planctas.id  ) " _
                    + vbCr + " LEFT JOIN " _
                    + vbCr + " ( SELECT vta_gastodebito.id AS idlgd, 1 AS ridtipper2, vta_gastodebito.idcli AS ridper2, vta_gastodebito.tipdoc AS rtipdoc2, vta_gastodebito.fchemi AS rfchope2, [vta_gastodebito].[numser] & '-' & [vta_gastodebito].[numdoc] AS rnumerodoc2 " _
                    + vbCr + " FROM vta_gastodebito WHERE vta_gastodebito.id = " & IdMovimiento & ") as cab on cab.idlgd = vta_gastodebito.id " _
                    + vbCr + " WHERE (((vta_gastodebito.id)=" & IdMovimiento & ")) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

                '===========================================================================================================================
        Case 42 '--Planilla de Letras Presentadas al Banco
            '--13/10/09:: Crear script para generar asiento

                '***********************************************************************************
                '--acualizamos el numero de registro, y fecha de registro
                xConn.Execute "UPDATE let_planilla SET let_planilla.numreg = '" & Format(IdMes, "00") + xAsiento & "' WHERE (((let_planilla.id)=" & IdMovimiento & "))"
                '***********************************************************************************


            '--DEBE "CABECERA DE PLANILLA"
            nSQL = "SELECT mae_banconumcta.idban, let_planilla.tipdoc, let_planilla.numser, let_planilla.numdoc, let_planilla.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " year([let_planilla].[fchreg]) AS anotra, Month([let_planilla].[fchreg]) AS idmes, let_planilla.idlib, let_planilla.id AS idmov, let_modalidadctabco.idcuen AS idcue, IIf([let_planilla].[numreg] Is Null,Null,Right([let_planilla].[numreg],4)) AS numasi, 0 AS iddoc, 0 AS correlativo, IIf([let_planilla].[tc] Is Null,0,[let_planilla].[tc]) AS tc, " _
                    + vbCr + " IIf(let_planilla.idmon=1,let_planilla.imptot,0) AS impdebsol, " _
                    + vbCr + " IIf(let_planilla.idmon=2,let_planilla.imptot,0) AS impdebdol, " _
                    + vbCr + " 0 AS imphabsol, " _
                    + vbCr + " 0 AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, let_planilla.fchreg AS fchasi, let_planilla.fchemi AS fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, let_planilla.idmon, " _
                    + vbCr + " 42 AS ridlib, 5 AS ridtipper, mae_banconumcta.idban AS ridper, let_planilla.tipdoc AS rtipdoc, let_planilla.fchemi AS rfchope, let_planilla.numdoc AS rnumerodoc, '' AS rglosa, let_planilla.glosa AS rglosaope, " _
                    + vbCr + " Left([let_planilla].[numreg],2) & [mae_libros].[codsun] & Right([let_planilla].[numreg],4) AS rregistro, let_planilla.idmon AS ridmon, " _
                    + vbCr + " 0 AS rtipdoc1, '' AS rnumerodoc1, 0 AS ridtipper2, 0 AS ridper2, 0 AS rtipdoc2, Null AS rfchope2, '' AS rnumerodoc2, " _
                    + vbCr + " IIF(tc = 0, 0, -1) as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM ((((let_planilla LEFT JOIN let_modalidadctabco ON (let_planilla.idmon = let_modalidadctabco.idmon) AND (let_planilla.idbcocta = let_modalidadctabco.idbcocta) AND (let_planilla.idmod = let_modalidadctabco.idmod)) LEFT JOIN con_planctas ON let_modalidadctabco.idcuen = con_planctas.id) LEFT JOIN mae_banconumcta ON let_planilla.idbcocta = mae_banconumcta.id) LEFT JOIN mae_bancos ON mae_banconumcta.idban = mae_bancos.id) LEFT JOIN mae_libros ON let_planilla.idlib = mae_libros.id " _
                    + vbCr + " WHERE (((let_planilla.id)= " & IdMovimiento & " )) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

            '===========================================================================================================================


            '--HABER "DETALLE DE LA PLANILLA"

            nSQL = "SELECT let_letra.idclipro, let_letra.tipdoc, let_letradet.numser, let_letradet.numdoc, let_planilla.idmon, let_planilla.numreg, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
                    + vbCr + " year([let_planilla].[fchreg]) AS anotra, Month([let_planilla].[fchreg]) AS idmes, let_planilla.idlib, let_planilla.id AS idmov, mae_letra.idcuenven AS idcue, IIf([let_planilla].[numreg] Is Null,Null,Right([let_planilla].[numreg],4)) AS numasi, let_planilladet.idcorrlet AS iddoc, 0 AS correlativo, IIf([let_planilla].[tc] Is Null,0,[let_planilla].[tc]) AS tc, " _
                    + vbCr + " 0 AS impdebsol, " _
                    + vbCr + " 0 AS impdebdol, " _
                    + vbCr + " IIf([let_planilla].[idmon]=1,[let_planilladet].[implet],0) AS imphabsol, " _
                    + vbCr + " IIf([let_planilla].[idmon]=2,[let_planilladet].[implet],0) AS imphabdol, " _
                    + vbCr + " 0 AS tipmov, 0 AS prodiv, let_planilla.fchreg AS fchasi, let_planilla.fchemi AS fchdoc,0 AS idorides, 0 AS idmod, 0 AS iddocpro, 0 AS tipo, let_planilla.idmon, " _
                    + vbCr + " 37 AS ridlib, 2 AS ridtipper, let_letra.idclipro AS ridper, let_letra.tipdoc AS rtipdoc, let_letradet.fchemi AS rfchope, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS rnumerodoc, '' AS rglosa, let_letra.glosa AS rglosaope, " _
                    + vbCr + " Left([let_letra].[numreg],2) & [mae_libros_1].[codsun] & Right([let_letra].[numreg],4) AS rregistro, let_letra.idmon AS ridmon, " _
                    + vbCr + " 0 AS rtipdoc1, let_letra.numerodocref AS rnumerodoc1, " _
                    + vbCr + " 5 AS ridtipper2, mae_banconumcta.idban AS ridper2, let_planilla.tipdoc AS rtipdoc2, let_planilla.fchemi AS rfchope2, let_planilla.numdoc AS rnumerodoc2, " _
                    + vbCr + " IIf(tc = 0, 0, -1) as aplicatc, 0 as ajuste " _
                    + vbCr + " FROM (let_letra LEFT JOIN mae_libros AS mae_libros_1 ON let_letra.idlib = mae_libros_1.id) RIGHT JOIN ((mae_letra RIGHT JOIN ((((let_planilla LEFT JOIN mae_banconumcta ON let_planilla.idbcocta = mae_banconumcta.id) LEFT JOIN mae_bancos ON mae_banconumcta.idban = mae_bancos.id) LEFT JOIN mae_libros ON let_planilla.idlib = mae_libros.id) INNER JOIN (let_planilladet LEFT JOIN let_letradet ON (let_planilladet.idlet = let_letradet.idlet) AND (let_planilladet.idcorrlet = let_letradet.corr)) ON let_planilla.id = let_planilladet.idpla) ON mae_letra.idmon = let_planilla.idmon) LEFT JOIN con_planctas ON mae_letra.idcuenven = con_planctas.id) ON let_letra.id = let_letradet.idlet " _
                    + vbCr + " WHERE (((let_planilla.id)= " & IdMovimiento & " )) "

                GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, DestinoAuto

            '===========================================================================================================================


  End Select

  '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
  '--/DEBE  /si el importe esta en el haber, el importe del destino del debe se encontrara en el haber
  '--/HABER /si el importe esta en el debe, el importe del destino del haber se encontrara en el debe

  '--insertando los asientos autoamticos

  If DestinoAuto = 1 Then

      '--cta automatico debe
      nSQL = " SELECT  con_diario.año, con_diario.idmes, con_diario.idlib, con_diario.idmov, con_planctas.ctadesdeb AS idcue, con_diario.numasi, con_diario.iddoc, 0 AS correlativo, con_diario.tc, " _
          + vbCr + " IIf(con_diario.idmon=1,IIf(con_diario.impdebsol<>0,con_diario.impdebsol,0),0) AS impdebsol, " _
          + vbCr + " IIf(con_diario.idmon=2,IIf(con_diario.impdebdol<>0,con_diario.impdebdol,0),0) AS impdebdol, " _
          + vbCr + " IIf(con_diario.idmon=1,IIf(con_diario.impdebsol=0,con_diario.imphabsol,0),0) AS imphabsol, " _
          + vbCr + " IIf(con_diario.idmon=2,IIf(con_diario.impdebdol=0,con_diario.imphabdol,0),0) AS imphabdol, " _
          + vbCr + " con_diario.tipmov, iif(con_diario.idlib=3,-1,0) AS prodiv, con_diario.fchasi, con_diario.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, con_diario.tipo, con_diario.idmon, " _
          + vbCr + " con_diario.ridlib,con_diario.ridtipper,con_diario.ridper,con_diario.rtipdoc,con_diario.rfchope,con_diario.rnumerodoc,con_diario.rregistro,con_diario.rglosa,con_diario.rglosaope,con_diario.ridmon, " _
          + vbCr + " con_diario.rtipdoc1, con_diario.rnumerodoc1, con_diario.ridtipper2,con_diario.ridper2, con_diario.rtipdoc2, con_diario.rfchope2, con_diario.rnumerodoc2,con_diario.aplicatc, con_diario.ajuste " _
          + vbCr + " FROM con_planctas INNER JOIN con_diario ON con_planctas.id = con_diario.idcue " _
          + vbCr + " WHERE (((con_diario.idlib)=" & IdLibro & ") AND ((con_diario.idmov)=" & IdMovimiento & ") AND ((con_planctas.ctadesdeb)<>0)) "
      '--cta automatico haber
      nSQL = nSQL & vbCr & " Union All " _
          + vbCr + " SELECT con_diario.año, con_diario.idmes, con_diario.idlib, con_diario.idmov, con_planctas.ctadeshab AS idcue, con_diario.numasi, con_diario.iddoc, 0 AS correlativo, con_diario.tc, " _
          + vbCr + " IIf(con_diario.idmon=1,IIf(con_diario.impdebsol=0,con_diario.imphabsol,0),0) AS impdebsol, " _
          + vbCr + " IIf(con_diario.idmon=2,IIf(con_diario.impdebdol=0,con_diario.imphabdol,0),0) AS impdebdol, " _
          + vbCr + " IIf(con_diario.idmon=1,IIf(con_diario.impdebsol<>0,con_diario.impdebsol,0),0) AS imphabsol, " _
          + vbCr + " IIf(con_diario.idmon=2,IIf(con_diario.impdebdol<>0,con_diario.impdebdol,0),0) AS imphabdol, " _
          + vbCr + " con_diario.tipmov, iif(con_diario.idlib=3,-1,0) AS prodiv, con_diario.fchasi, con_diario.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, con_diario.tipo, con_diario.idmon, " _
          + vbCr + " con_diario.ridlib,con_diario.ridtipper,con_diario.ridper,con_diario.rtipdoc,con_diario.rfchope,con_diario.rnumerodoc,con_diario.rregistro,con_diario.rglosa,con_diario.rglosaope ,con_diario.ridmon,  " _
          + vbCr + " con_diario.rtipdoc1, con_diario.rnumerodoc1, con_diario.ridtipper2,con_diario.ridper2, con_diario.rtipdoc2, con_diario.rfchope2, con_diario.rnumerodoc2,con_diario.aplicatc, con_diario.ajuste " _
          + vbCr + " FROM con_planctas INNER JOIN con_diario ON con_planctas.id = con_diario.idcue " _
          + vbCr + " WHERE (((con_diario.idlib)=" & IdLibro & ") AND ((con_diario.idmov)=" & IdMovimiento & ") AND ((con_planctas.ctadeshab)<>0)) "

        GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, True
        '===========================================================================================================================

  ElseIf DestinoAuto = 2 Then
      '--cta automatico debe

      nSQL = " SELECT con_diario.año, con_diario.idmes, con_diario.idlib, con_diario.idmov, con_planctas.ctadesdeb AS idcue, con_diario.numasi, con_diario.iddoc, 0 AS correlativo, con_diario.tc, Sum(IIf(con_diario.idmon=1,IIf(con_diario.impdebsol<>0,con_diario.impdebsol,0),0)) AS impdebsol, " _
          + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(con_diario.impdebdol<>0,con_diario.impdebdol,0),0)) AS impdebdol, " _
          + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(con_diario.impdebsol=0,con_diario.imphabsol,0),0)) AS imphabsol, " _
          + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(con_diario.impdebdol=0,con_diario.imphabdol,0),0)) AS imphabdol, " _
          + vbCr + " con_diario.tipmov, iif(con_diario.idlib=3,-1,0) AS prodiv, con_diario.fchasi, con_diario.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, con_diario.tipo, con_diario.idmon, " _
          + vbCr + " con_diario.ridlib,con_diario.ridtipper,con_diario.ridper,con_diario.rtipdoc,con_diario.rfchope,con_diario.rnumerodoc,con_diario.rregistro,con_diario.rglosa,con_diario.ridmon,  " _
          + vbCr + " con_diario.rtipdoc1, con_diario.rnumerodoc1, con_diario.ridtipper2,con_diario.ridper2, con_diario.rtipdoc2, con_diario.rfchope2, con_diario.rnumerodoc2, con_diario.ajuste " _
          + vbCr + " FROM con_planctas INNER JOIN con_diario ON con_planctas.id = con_diario.idcue " _
          + vbCr + " GROUP BY con_diario.año, con_diario.idmes, con_diario.idlib, con_diario.idmov, con_planctas.ctadesdeb, con_diario.numasi, con_diario.iddoc, con_diario.tc, con_diario.tipmov, con_diario.fchasi, con_diario.fchdoc, con_diario.tipo, con_diario.idmon, con_diario.ridlib, con_diario.ridtipper, con_diario.ridper, con_diario.rtipdoc, con_diario.rfchope, con_diario.rnumerodoc, con_diario.rregistro, con_diario.rglosa, con_diario.ridmon, con_diario.rtipdoc1, con_diario.rnumerodoc1, con_diario.ridtipper2, con_diario.ridper2, con_diario.rtipdoc2, con_diario.rfchope2, con_diario.rnumerodoc2, con_diario.ajuste " _
          + vbCr + " HAVING (((con_diario.idlib)=" & IdLibro & ") AND ((con_diario.idmov)=" & IdMovimiento & ") AND ((con_planctas.ctadesdeb)<>0)) "

      '--cta automatico haber
      nSQL = nSQL & vbCr & " Union All " _
          + vbCr + " SELECT con_diario.año, con_diario.idmes, con_diario.idlib, con_diario.idmov, con_planctas.ctadeshab AS idcue, con_diario.numasi, con_diario.iddoc, 0 AS correlativo, con_diario.tc, Sum(IIf(con_diario.idmon=1,IIf(con_diario.impdebsol=0,con_diario.imphabsol,0),0)) AS impdebsol, " _
          + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(con_diario.impdebdol=0,con_diario.imphabdol,0),0)) AS impdebdol, " _
          + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(con_diario.impdebsol<>0,con_diario.impdebsol,0),0)) AS imphabsol, " _
          + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(con_diario.impdebdol<>0,con_diario.impdebdol,0),0)) AS imphabdol, " _
          + vbCr + " con_diario.tipmov, iif(con_diario.idlib=3,-1,0) AS prodiv, con_diario.fchasi, con_diario.fchdoc, 0 AS idorides, 0 AS idmod, 0 AS iddocpro, con_diario.tipo, con_diario.idmon, " _
          + vbCr + " con_diario.ridlib,con_diario.ridtipper,con_diario.ridper,con_diario.rtipdoc,con_diario.rfchope,con_diario.rnumerodoc,con_diario.rregistro,con_diario.rglosa,con_diario.ridmon,  " _
          + vbCr + " con_diario.rtipdoc1, con_diario.rnumerodoc1, con_diario.ridtipper2,con_diario.ridper2, con_diario.rtipdoc2, con_diario.rfchope2, con_diario.rnumerodoc2, con_diario.ajuste " _
          + vbCr + " FROM con_planctas INNER JOIN con_diario ON con_planctas.id = con_diario.idcue " _
          + vbCr + " GROUP BY con_diario.año, con_diario.idmes, con_diario.idlib, con_diario.idmov, con_planctas.ctadeshab, con_diario.numasi, con_diario.iddoc, con_diario.tc, con_diario.tipmov, con_diario.fchasi, con_diario.fchdoc, con_diario.tipo, con_diario.idmon, con_diario.ridlib, con_diario.ridtipper, con_diario.ridper, con_diario.rtipdoc, con_diario.rfchope, con_diario.rnumerodoc, con_diario.rregistro, con_diario.rglosa, con_diario.ridmon, con_diario.rtipdoc1, con_diario.rnumerodoc1, con_diario.ridtipper2, con_diario.ridper2, con_diario.rtipdoc2, con_diario.rfchope2, con_diario.rnumerodoc2, con_diario.ajuste " _
          + vbCr + " HAVING (((con_diario.idlib)=" & IdLibro & ") AND ((con_diario.idmov)=" & IdMovimiento & ") AND ((con_planctas.ctadeshab)<>0)) "

        GenerarAsientoInsert xConn, IdLibro, AnnoTrabajo, IdMes, xAsiento, nSQL, True
        '===========================================================================================================================

  End If

Salto_Apertura: '--Comodin para saltar linea de codigo solo cuando sea apertura.

  '----------------------------------------------------------------------------------------------------------------------------------------------------------------------

    GenerarAsiento = xNumRegistro
    Set dll_varios = Nothing

Exit Function
ERROR:
'''''    Resume
    MsgBox "Generar Asiento - No se pudo guardar el registro por el siguiente motivo: " + vbCr + Trim(Err.Description) & vbCr & Err.Number, vbCritical, xTitulo
    Err.Clear
    GenerarAsiento = ""
    Set dll_varios = Nothing

End Function


Private Sub GenerarAsientoInsert(xConn As ADODB.Connection, IdLibro As Integer, AnnoTrabajo As String, IdMes As Integer, xAsiento As String, nSQL As String, Optional DestinoAuto As Integer)
    '===================================================================================================
    'Creado : 05/02/09 Por: Johan Castro
    'Propósito: Insertar registro de asiento en el diario
    '
    'Entradas:  xConn = Conexion activa a la base de datos , antes xcon
    '           IdLibro = Codigo de Libro contable
    '           AnoTrabajo = Año de trabajo
    '           IdMes = Mes de trabajo
    '           xAsiento = Asiento asignado al movimiento
    '           nSQL = indica la cosulta a procesar
    '           DestinoAuto=Indica como se va generar los asientos de destino True:Detallado; False:Resumen
    '
    'Resultados:Registro en el diario
    'Modificado: 30/03/11 Johan Castro
    '           Cambiar nombre de parametro xcon a xconn por que hay incompatibilidad con variable publica de proyecto con el parametro xcon
    '           se procede a renombrar todas las lineas de codigo que usen parametro xcon a xconn
    '           21/02/12 Johan Castro
    '           Agregar campo ajuste= para guardar si aplica ajuste x dif cambio, para el select() e insert()
    '===================================================================================================

    If DestinoAuto = 10 Then
        nSQL = Replace(nSQL, "WHERE", "WHERE (con_tc.idmon=2 or con_tc.idmon is null) and ")
        nSQL = Replace(nSQL, "HAVING", "HAVING (con_tc.idmon=2 or con_tc.idmon is null) and ")
        nSQL = Replace(nSQL, "GROUP BY", "GROUP BY con_tc.idmon,")
    End If
  '--generando la consulta para insertar

    nSQL = "SELECT " & AnnoTrabajo & " as anotra, " & IdMes & " as idmes, " & IIf(IdMes = 0, 3, IdLibro) & " as idlib, diario.idmov, diario.idcue, '" & xAsiento & "' as numasi, diario.iddoc, diario.correlativo, diario.tc, diario.impdebsol, diario.impdebdol, diario.imphabsol, diario.imphabdol, diario.tipmov, diario.prodiv, " & IIf(IdMes = 0, "NULL", "diario.fchasi") & " AS fchasi, diario.fchdoc, diario.idorides, diario.idmod, diario.iddocpro, diario.tipo,diario.idmon, diario.ridlib,diario.ridtipper,diario.ridper,diario.rtipdoc,diario.rfchope,diario.rnumerodoc,diario.rregistro,diario.rglosa, diario.rglosaope,diario.ridmon,diario.rtipdoc1, diario.rnumerodoc1, diario.ridtipper2, diario.ridper2, diario.rtipdoc2, diario.rfchope2, diario.rnumerodoc2,diario.aplicatc,diario.ajuste " _
        + vbCr + " From ( " _
        + vbCr + nSQL _
        + vbCr + " )  AS diario "

'    --insertar el diario
    '
    xConn.Execute "INSERT INTO con_diario (año,idmes,idlib,idmov,idcue,numasi,iddoc,correlativo,tc,impdebsol,impdebdol,imphabsol,imphabdol,tipmov,prodiv,fchasi,fchdoc,idorides,idmod,iddocpro,tipo,idmon,ridlib,ridtipper,ridper,rtipdoc,rfchope,rnumerodoc,rregistro,rglosa,rglosaope,ridmon,rtipdoc1,rnumerodoc1,ridtipper2,ridper2,rtipdoc2,rfchope2,rnumerodoc2,aplicatc,ajuste ) " & _
                 nSQL
                 
        
'''''    nSQL = "SELECT " & AnnoTrabajo & " as anotra, " & IdMes & " as idmes, " & IIf(IdMes = 0, 3, IdLibro) & " as idlib, diario.idmov, diario.idcue, '" & xAsiento & "' as numasi, diario.iddoc, diario.correlativo, diario.tc, diario.impdebsol, diario.impdebdol, diario.imphabsol, diario.imphabdol, diario.tipmov, diario.prodiv, " & IIf(IdMes = 0, "NULL", "diario.fchasi") & " AS fchasi, diario.fchdoc, diario.idorides, diario.idmod, diario.iddocpro, diario.tipo,diario.idmon, diario.ridlib,diario.ridtipper,diario.ridper,diario.rtipdoc,diario.rfchope,diario.rnumerodoc,diario.rregistro,diario.rglosa, diario.rglosaope,diario.ridmon,diario.rtipdoc1, diario.rnumerodoc1, diario.ridtipper2, diario.ridper2, diario.rtipdoc2, diario.rfchope2, diario.rnumerodoc2,diario.aplicatc " _
'''''        + vbCr + " From ( " _
'''''        + vbCr + nSQL _
'''''        + vbCr + " )  AS diario "
'''''
'''''    xConn.Execute "INSERT INTO con_diario (año,idmes,idlib,idmov,idcue,numasi,iddoc,correlativo,tc,impdebsol,impdebdol,imphabsol,imphabdol,tipmov,prodiv,fchasi,fchdoc,idorides,idmod,iddocpro,tipo,idmon,ridlib,ridtipper,ridper,rtipdoc,rfchope,rnumerodoc,rregistro,rglosa,rglosaope,ridmon,rtipdoc1,rnumerodoc1,ridtipper2,ridper2,rtipdoc2,rfchope2,rnumerodoc2,aplicatc ) " & _
'''''                 nSQL

End Sub

Public Sub xls_UnirCelda(objExcel As Object, pRango As String)
    With objExcel
        .Range(pRango).Select
        With .Selection
            .WrapText = False
            .Orientation = 0
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Selection.Merge
    End With
End Sub




''Public Sub RstRegistroBuscar(Rst As ADODB.Recordset, nCampo As String, _
''                               nValor As String, fEsNumero As Boolean)
''    '--02/01/09
''
''    Dim Apostrofe As String
''    If Rst.State = 0 Then Exit Sub
''    Rst.Filter = ""
''    If Rst.RecordCount = 0 Then Exit Sub
''
''    If fEsNumero = False Then Apostrofe = "'"
''    If fEsNumero = True Then Apostrofe = ""
''
''    Rst.MoveFirst
''    '--buscando el registro
''    '--con filtro
''    Rst.Filter = nCampo & " = " & Apostrofe & nValor & Apostrofe
''    If Rst.EOF = False Or Rst.BOF = False Or Rst.RecordCount <> 0 Then Rst.MoveFirst
''    Do While Not Rst.EOF
''        Rst.Delete
''        Rst.Update
''        Rst.MoveNext
''    Loop
''    Rst.Filter = ""
''    If Rst.EOF = False Or Rst.BOF = False Or Rst.RecordCount <> 0 Then Rst.MoveFirst
''
''End Sub

Public Function fSetearCuadriculaColumna(Con As ADODB.Connection, mIdFormato As Integer, Optional mOrigen As Integer = 1, Optional mIdFormatoDet As Integer = 0) As String
    '===================================================================================================
    'Creado : xx/xx/09 Por: Johan Castro
    'Propósito: Muestra el orden de los campos en texto, solo los campos que se muestran en el reporte.
    '
    'Entradas:  xcon = Conexion activa a la base de datos
    '           mIdFormato = Codigo del formato a mostrar
    '           mOrigen = Indica el origen de donde se obtendra el formato del reporte
    '           1=Tabla con_formatos(x defecto); 2=Tabla con_analisis(Opcional)
    '           mIdFormatoDet = Indica el codigo para filtrar registros del detalle
    '
    'Resultados:Relación de campos en formato texto concatenados por ","
    '
    'Modificado  03/01/11 Por Johan Castro
    '            Cambiar el nombre del parametro mFormato a mIdFormato
    '            Agregar parametro mOrigen para agregar lineas que hagan referencia a la tabla con_analisis
    'Modificado  27/04/11 Johan castro
    '            Agregar parametro mIdFormatoDet, no se toma en cuenta el filtro x "defecto = -1"
    '===================================================================================================

    Dim nSQL As String
    Dim nSQLCampos As String
    Dim Rst As New ADODB.Recordset
    Dim nSQLFiltro As String

    If mIdFormatoDet = 0 Then
        nSQLFiltro = " and f.defecto = -1 "
    Else
        nSQLFiltro = " and f_det.idformatotipo=" & mIdFormatoDet & " "
    End If

    If mOrigen = 1 Then
        '--Obtener datos de con_formatos
        nSQL = "SELECT f_det.abrev, f_det.nomcampo " _
         & " FROM con_formatostipo AS f LEFT JOIN con_formatostipodet AS f_det ON (f.idformato = f_det.idformato) AND (f.id = f_det.idformatotipo) " _
         & " WHERE f.idformato=" & mIdFormato & " AND f_det.mostrar=-1 " & nSQLFiltro _
         & " ORDER BY f_det.orden; "
    Else
        '--Obtener datos de con_analisis
        nSQL = "SELECT f_det.abrev, f_det.nomcampo " _
         & " FROM con_analisistipo AS f LEFT JOIN con_analisistipodet AS f_det ON (f.idformato = f_det.idformato) AND (f.id = f_det.idformatotipo) " _
         & " WHERE f.idformato=" & mIdFormato & " AND f_det.mostrar=-1 " & nSQLFiltro _
         & " ORDER BY f_det.orden; "
    End If

    RST_Busq Rst, nSQL, Con

    If Rst.RecordCount = 0 Then
        MsgBox "El formato que desea Consultar no tiene columnas para visualizar" & vbCr & "Para continuar seleccione las columnas que desea visualizar", vbExclamation, xTitulo
        fSetearCuadriculaColumna = ""
        Exit Function
    End If

    Rst.MoveFirst
    nSQLCampos = ""
    Do While Not Rst.EOF
        nSQLCampos = nSQLCampos & NulosC(Rst("nomcampo")) & ","
        Rst.MoveNext
    Loop

    Set Rst = Nothing
    fSetearCuadriculaColumna = mID(nSQLCampos, 1, Len(nSQLCampos) - 1)



End Function


Public Sub CierrePeriodo(oToolBar As Object, _
                         IdFormulario As Integer, _
                         IdMes As Integer, _
                         fCierrePeriodo As Boolean, _
                         xCon As ADODB.Connection, _
                         Optional xIdUsuario As Integer)

    '===================================================================================================
    'Creado : 30/03/09 Por: Johan Castro
    'Propósito: Restringir el acceso al formulario segun mes seleccionado
    '
    'Entradas:  oToolBar = Control ToolBar
    '           IdFormulario = id del formulario activo
    '           IdMes = Mes de trabajo
    '           fCierrePeriodo = variable que permitira bloquear cuando el usuario intente editar con el teclado
    '           xcon = Conexion activa a la base de datos
    '           xIdUsuario = Codigo del usuario
    '
    'Resultados:formulario bloqueado o listo para editar
    '
    'modificado  04/01/10 por Johan Castro
    '            Se agrega "entrada" idusuario para bloquear los accesos por usuario, utilizara funcion OpcionesUsuario
    '            Controlar los errores
    '            17/01/11 por Johan Castro
    '            El parametro IdFormulario almacenara el codigo del menu de tabla mae_menu
    '            Modificar linea de codigo para el bloqueo de los accesos
    '===================================================================================================

    On Error GoTo LaCague

    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    Dim A&
    nSQL = "SELECT var_cierre.estado FROM var_cierre WHERE (((var_cierre.idmes)=" & IdMes & " ) AND ((var_cierre.idform)=" & IdFormulario & "));"
    RST_Busq Rst, nSQL, xCon
    If Rst.State = 1 Then
        If Rst.RecordCount <> 0 Then
            fCierrePeriodo = CBool(NulosN(Rst("estado")))
            For A = 1 To oToolBar.Buttons.Count
                Select Case LCase(Trim(oToolBar.Buttons(A).ToolTipText))
                    Case "agregar", "modificar", "eliminar"
                        oToolBar.Buttons(A).Enabled = CBool(NulosN(Rst("estado")))
                    Case Else

                End Select
            Next A
            If Not fCierrePeriodo Then
                MsgBox "El presente mes se encuentra cerrado para la opción actual" _
                        , vbInformation + vbOKOnly + vbDefaultButton1, "Cierre de Mes"
            End If
        Else
            fCierrePeriodo = False
        End If
    Else
        fCierrePeriodo = False
    End If
    Set Rst = Nothing
'
    '--bloquear opciones por usuario
    '--utilizar funcion creada por enrique para bloquear acceso por usuario
    Dim obj As Object
    Set obj = CreateObject("sgi2_funciones.varias")
    obj.OpcionesUsuario xIdUsuario, IdFormulario, oToolBar, xCon
    Set obj = Nothing
'    Set Rst = Nothing

    '--si agregar, modificar o elimiar esta oculto entonces los accesos directos no fucionara con el teclado
    If oToolBar.Buttons(1).Visible = False Or oToolBar.Buttons(2).Visible = False Or oToolBar.Buttons(3).Visible = False Then
        fCierrePeriodo = False
    End If

    Exit Sub

LaCague:
    fCierrePeriodo = False
    MsgBox Err.Description & vbCr & Err.Source, vbCritical, xTitulo
    Err.Clear

End Sub

Sub AperturaDocumento(Con As ADODB.Connection, IdUsuario As Integer, Idform As Integer, IdMenu As Integer)

    '===================================================================================================
    'Creado : 05/02/10 Por: Johan Castro
    'Propósito: Cargar los registros de la base de datos que servira como apertura, segun parametros
    '           seleccionado por el usuario
    '
    'Entradas:  Con=conexion a la base de datos
    '           IdUsuario = Código de usuario
    '           IdForm = Código de Libro
    '           IdMenu = Código del Menu
    '
    'Resultados: Registros que seleccionan seran agregados como apertura
    '
    'Nota:       1.- Buscar el documento
    '            2.- Activar con el check
    '            3.- Repetir pasos 1,2 tantas veces documentos tenga
    '            4.- Hacer clic en boton aceptar(se graba automaticamente en su tabla correspondiente)
    '                El saldo se graba en el campo total del documento
    '                El total del documento se graba en un campo temporal "imptotori"
    'Modificado  29/04/11 Johan Castro
    '            Considerar el campo imptotori como saldo final del cierre de periodo, este saldo sera como
    '            utilizado para analizar la cuenta en el siguiente ejercicio economico ( Compras, ventas, honorarios)
    'Modificado  04/01/12 Enrique Pollongo
    '            Agregar linea para grabar el saldo de honorarios.
    '            Comentar linea que graba el importe descontado
    'Modificado  24/01/12 Johan Castro
    '            Agregar parámetro IdMenu para escribir el log.
    '            Agregar variable xHoraInicio para la hora de inicio que se grabara en el Log
    '            Mostrar Documentos de LGD
    '            Cambiar paso de parámetro a evento GrabarOperacion() Menu::antes IdForm, ahora IdMenu; Quehace Antes 1=Nuevo, Ahora 6=Importado
    '===================================================================================================


    Dim nSQL As String
    Dim nSQLPeriodo As String '--filtrar por periodo
    Dim nTitulo As String '--titulo que se muestra en seleccionar registro
    Dim xHoraInicio As Date
    '--asiganar la conexion a la variable temporal
    Set xCon = Con

    '--cargar datos de reconocimiento de la base de datos
    Class_Initialize

    '--cargar datos de la empresa
    CargaDatosEmpresa

    xHoraInicio = Time()

    Select Case Idform
        Case 1, 40 '--1=compras, 40=Honorarios
            ReDim xCampos(8, 5) As String

            xCampos(0, 0) = "Nº.Reg.":       xCampos(0, 1) = "registro":       xCampos(0, 2) = "850":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
            xCampos(1, 0) = "Proveedor":     xCampos(1, 1) = "nombre":         xCampos(1, 2) = "3500":    xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
            xCampos(2, 0) = "T.D.":          xCampos(2, 1) = "abrev":          xCampos(2, 2) = "600":     xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
            xCampos(3, 0) = "Nº Documento":  xCampos(3, 1) = "numerodoc":      xCampos(3, 2) = "1500":    xCampos(3, 3) = "C":    xCampos(3, 4) = "S"
            xCampos(4, 0) = "Fch. Emi.":     xCampos(4, 1) = "fchdoc1":        xCampos(4, 2) = "1000":    xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
            xCampos(5, 0) = "M":             xCampos(5, 1) = "moneda":         xCampos(5, 2) = "450":     xCampos(5, 3) = "C":    xCampos(5, 4) = "N"
            xCampos(6, 0) = "Importe":       xCampos(6, 1) = "total":          xCampos(6, 2) = "1100":    xCampos(6, 3) = "N":    xCampos(6, 4) = "N"
            xCampos(7, 0) = "Saldo":         xCampos(7, 1) = "impsal":         xCampos(7, 2) = "1100":    xCampos(7, 3) = "N":    xCampos(7, 4) = "N"

        If Idform = 1 Then
            nSQL = "SELECT 0 as xsel, Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS registro, mae_prov.numruc as ruc, mae_prov.nombre, mae_documento.abrev, com_compras.numser & '-' & com_compras.numdoc as numerodoc,mae_moneda.simbolo as moneda, com_compras.imptot as total,com_compras.fchdoc & '' as fchdoc1,com_compras.* " _
                + vbCr + " FROM (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id " & nSQLPeriodo

            nTitulo = "Seleccionar Documento de Compras"

        ElseIf Idform = 40 Then
            nSQL = "SELECT 0 as xsel,Left([com_honorarios].[numreg],2) & [mae_libros].[codsun] & Right([com_honorarios].[numreg],4) AS registro, mae_prov.numruc as ruc, mae_prov.nombre, mae_documento.abrev, com_honorarios.numser & '-' & com_honorarios.numdoc as numerodoc,mae_moneda.simbolo as moneda, com_honorarios.imptot as total,com_honorarios.fchdoc & '' as fchdoc1,com_honorarios.* " _
                + vbCr + " FROM (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_honorarios ON mae_documento.id = com_honorarios.tipdoc) ON mae_moneda.id = com_honorarios.idmon) ON mae_prov.id = com_honorarios.idpro) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id " & nSQLPeriodo

            nTitulo = "Seleccionar Documento de Honorarios"
        Else
            Exit Sub
        End If

        Case 2, 41 '--ventas, LGD
            ReDim xCampos(8, 5) As String
            xCampos(0, 0) = "Nº.Reg.":       xCampos(0, 1) = "registro":       xCampos(0, 2) = "850":     xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
            xCampos(1, 0) = "Cliente":       xCampos(1, 1) = "nombre":         xCampos(1, 2) = "3500":    xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
            xCampos(2, 0) = "T.D.":          xCampos(2, 1) = "abrev":          xCampos(2, 2) = "600":     xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
            xCampos(3, 0) = "Nº Documento":  xCampos(3, 1) = "numerodoc":      xCampos(3, 2) = "1500":    xCampos(3, 3) = "C":    xCampos(3, 4) = "S"
            xCampos(4, 0) = "Fch. Emi.":     xCampos(4, 1) = "fchdoc1":         xCampos(4, 2) = "1000":    xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
            xCampos(5, 0) = "M":             xCampos(5, 1) = "moneda":         xCampos(5, 2) = "450":     xCampos(5, 3) = "C":    xCampos(5, 4) = "N"
            xCampos(6, 0) = "Importe":       xCampos(6, 1) = "total":          xCampos(6, 2) = "1100":    xCampos(6, 3) = "N":    xCampos(6, 4) = "N"
            xCampos(7, 0) = "Saldo":         xCampos(7, 1) = "impsal":         xCampos(7, 2) = "1100":    xCampos(7, 3) = "N":    xCampos(7, 4) = "N"

            If Idform = 2 Then
                nSQL = "SELECT 0 as xsel,Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc AS ruc, mae_cliente.nombre, mae_documento.abrev, vta_ventas.numser & '-' & vta_ventas.numdoc AS numerodoc, mae_moneda.simbolo AS moneda, vta_ventas.imptotdoc AS total,vta_ventas.fchdoc & '' as fchdoc1, vta_ventas.* " _
                    + vbCr + " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN vta_ventas ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
                    + vbCr + " WHERE (((vta_ventas.anulado)=0)) " & nSQLPeriodo

                '----------------------------------------------------
                nTitulo = "Seleccionar Documento de Ventas"

            ElseIf Idform = 41 Then
                nSQL = "SELECT 0 as xsel,Left([vta_gastodebito].[numreg],2) & [mae_libros].[codsun] & Right([vta_gastodebito].[numreg],4) AS registro, mae_cliente.numruc AS ruc, mae_cliente.nombre, mae_documento.abrev, vta_gastodebito.numser & '-' & vta_gastodebito.numdoc AS numerodoc, mae_moneda.simbolo AS moneda, vta_gastodebito.imptot AS total,vta_gastodebito.fchemi & '' AS fchdoc1,vta_gastodebito.* " _
                    + vbCr + " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN vta_gastodebito ON mae_documento.id = vta_gastodebito.tipdoc) ON mae_moneda.id = vta_gastodebito.idmon) ON mae_cliente.id = vta_gastodebito.idcli) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id " _
                    + vbCr + " WHERE (((vta_gastodebito.anulado)=0)) " & nSQLPeriodo

                '----------------------------------------------------
                nTitulo = "Seleccionar Documento de Liquidación de Gasto Debito"
            Else
                Exit Sub
            End If

'        Case 24 '--boleta de pago
'
'
'        Case 20 '--planilla de letras
'
'
'        Case 19 '--letras
        Case Else
            MsgBox "Falta implementar la opcion para importar  información del libro ", vbInformation, xTitulo
            Exit Sub
    End Select

    '-------------------------------------------------------------------------------------------------------------
    '--proceder a cargar los datos
    '-------------------------------------------------------------------------------------------------------------

    Dim nPathBD As String '--ruta de la base de datos de apertura

    Dim RstTmp As New ADODB.Recordset
    Dim xConTmp As New ADODB.Connection '--conexion temporal a la base de datos (maestro o de apertura)

    On Error GoTo LaCague

    DoEvents

    '------------------------------------
    '--verificar si la base de datos principal existe
    If ArchivoExiste(AP_RUTABD + "data.mdb") = False Then
        MsgBox "No existe la ruta a la Base de Datos Principal", vbInformation, xTitulo
        Exit Sub
    End If

    '--abriendo la conexion a la base de datos principal, para obtener la ruta de la data de apertura
    '--para esto buscamos por el num ruc de la empresa que se trabaja
    OPEN_CONEX_TMP xConTmp, AP_RUTABD + "data.mdb"
    If xConTmp.State = 0 Then Exit Sub
    RST_Busq RstTmp, "SELECT ruta,anotra FROM mae_empresa WHERE numruc = '" + NumRUC + "' and anotra = " & CInt(AnoTra) - 1, xConTmp

    '--verificar si existe la base de datos para proceder a consultar la informacion de apertura
    If RstTmp.RecordCount = 0 Then
        MsgBox "No existe el año de trabajo para buscar datos de apertura", vbInformation, xTitulo
        Exit Sub
    End If

    '--asignando la ruta de la base de datos de apertura a la variable
    nPathBD = NulosC(RstTmp("ruta"))
    '-------
    '--cerrando la conexion a la base de datos principal para conectar a la base de datos de apertura
    Set RstTmp = Nothing
    xConTmp.Close
    '------------------------------------------------
    '--verificar si la base de datos de apertura existe
    If ArchivoExiste(AP_RUTABD & nPathBD) = False Then
        MsgBox "No existe la ruta a la Base de Datos Año: " + AnoTra - 1, vbInformation, xTitulo
        GoTo salir
    End If
    '--abriendo la nueva conexion a la base de datos de apertura
    OPEN_CONEX_TMP xConTmp, AP_RUTABD + nPathBD
    If xConTmp.State = 0 Then Exit Sub
    '--cargando la data segun libro y periodo seleccionado
    CARGAR_DLL_EPSBUSCAR_SEL xConTmp, RstTmp, nSQL, xCampos(), nTitulo
    '-----------------------------------------------------------------------------------------------------------------------------------
    '--cerrando conexion a la base de datos apertura
    xConTmp.Close
    Dim RstCab As New ADODB.Recordset
    Dim IdPer As Double '--codigo de proveedor o cliente
    Dim xId As Double
    Dim RstBus As New ADODB.Recordset
    Dim GrabaOperacion As Boolean '--indica si agrega registro true::Graba; False:: No graba
    '--definir objeto para invocar libreria funciones
    Dim obj As Object
    Set obj = CreateObject("sgi2_funciones.varias")
    '-----------------
    '--proceder a grabar los datos en sus respectivas tablas
    If RstTmp.State = 0 Then GoTo salir

    '--si hay registros seleccionados, ubicar en el primer registro
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst

    '--inciar el bucle para grabar registro por registro
    Do While Not RstTmp.EOF
        '--incializar
        GrabaOperacion = False

        Set RstCab = Nothing
        Set RstBus = Nothing

        Select Case Idform
            Case 1 '--compras

            '******************
            RST_Busq RstCab, "SELECT TOP 1 * FROM com_compras", xCon

            '--buscar si ya esta ingresado el documento
            nSQL = "SELECT mae_prov.numruc AS ruc, com_compras.numser & '-' & com_compras.numdoc AS numerodoc, com_compras.tipdoc " _
                + vbCr + " FROM mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro " _
                + vbCr + " WHERE (((mae_prov.numruc)='" & NulosC(RstTmp("ruc")) & "') AND ((com_compras.numser & '-' & com_compras.numdoc)='" & RstTmp("numerodoc") & "') AND ((com_compras.tipdoc)=" & RstTmp("tipdoc") & "));"
            '--ejecutar la consulta
            RST_Busq RstBus, nSQL, xCon

            '--verificar si no se a registrado
            If RstBus.RecordCount = 0 Then
                '--buscar el codigo del proveedor
                IdPer = NulosN(Busca_Codigo(NulosC(RstTmp("ruc")), "numruc", "id", "mae_prov", "C", xCon))
                '--si el proveedor no esta registrado en el nuevo ejercicio, insertar el proveedor
                If IdPer = 0 Then
                    IdPer = HallaCodigoTabla("mae_prov", xCon, "id")
                    xCon.Execute "insert into mae_prov (id,tipper,idtipdoc,numruc,nombre) values (" & IdPer & ",2,5,'" & RstTmp("ruc") & "','" & RstTmp("nombre") & "')"
                End If

                '---------------------------------------------------------------------------
                '--buscar si el codigo del documento existe, sino utilizar el mismo codigo
                Set RstBus = Nothing
                nSQL = "select id from com_compras where id=" & NulosN(RstTmp("id"))
                RST_Busq RstBus, nSQL, xCon
                If RstBus.RecordCount = 0 Then
                    '--si el codigo no esta en uso =>> asignar el codigo
                    xId = NulosN(RstTmp("id"))
                Else
                    '--obtener el nuevo codigo cuando el codigo este en uso
                    xId = HallaCodigoTabla("com_compras", xCon, "id")
                End If
                '-----------------------------------------------------------------------------
                RstCab.AddNew
                RstCab("id") = xId
                RstCab("idlib") = 1
                RstCab("idtipo") = NulosN(RstTmp("idtipo"))
                RstCab("tipdoc") = NulosN(RstTmp("tipdoc"))
                RstCab("idpro") = IdPer
                RstCab("numser") = NulosC(RstTmp("numser"))
                RstCab("numdoc") = NulosC(RstTmp("numdoc"))
                If IsDate(RstTmp("fchdoc")) = True Then RstCab("fchdoc") = CDate(RstTmp("fchdoc"))
                If IsDate(RstTmp("fchven")) = True Then RstCab("fchven") = CDate(RstTmp("fchven"))
                RstCab("idconpag") = NulosN(RstTmp("idconpag"))
                RstCab("idmon") = NulosN(RstTmp("idmon"))
                RstCab("impbru") = NulosN(RstTmp("impbru"))
                RstCab("impbru2") = NulosN(RstTmp("impbru2"))
                RstCab("impbru3") = NulosN(RstTmp("impbru3"))
                RstCab("impina") = NulosN(RstTmp("impina"))
                RstCab("impigv") = NulosN(RstTmp("impigv"))
                RstCab("impigv2") = NulosN(RstTmp("impigv2"))
                RstCab("impigv3") = NulosN(RstTmp("impigv3"))
                RstCab("impisc") = NulosN(RstTmp("impisc"))
                'RstCab("impdesc") = NulosN(RstTmp("impdesc"))
                RstCab("otroscargos") = NulosN(RstTmp("otroscargos"))
                RstCab("imptot") = NulosN(RstTmp("imptot"))
                RstCab("imptotori") = NulosN(RstTmp("impsal")) '--saldo final
                RstCab("impsal") = NulosN(RstTmp("impsal"))
                RstCab("glosa") = NulosC(RstTmp("glosa"))
                RstCab("idalm") = NulosN(RstTmp("idalm"))
                RstCab("fchreg") = CDate("01/01" + "/" + AnoTra)
                RstCab("numreg") = "000001"
                '''RstCab("impred") = 0
                RstCab("tc") = 0
                RstCab("iddocref") = 0
                RstCab("idtipdocref") = 0
                RstCab("iddocref2") = 0
                RstCab("numerodocref") = NulosC(RstTmp("numerodocref"))
                RstCab.Update
                '--activar opcion de grabar operacion
                GrabaOperacion = True

            End If

            '******************
            Case 2 '--ventas
                RST_Busq RstCab, "SELECT TOP 1 * FROM vta_ventas", xCon

                '--buscar si ya esta ingresado el documento
                nSQL = "SELECT mae_cliente.numruc AS ruc, vta_ventas.numser & '-' & vta_ventas.numdoc AS numerodoc, vta_ventas.tipdoc " _
                    + vbCr + " FROM mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli " _
                    + vbCr + " WHERE (((mae_cliente.numruc)='" & NulosC(RstTmp("ruc")) & "') AND ((vta_ventas.numser & '-' & vta_ventas.numdoc)='" & RstTmp("numerodoc") & "') AND ((vta_ventas.tipdoc)=" & RstTmp("tipdoc") & "));"

                '--ejecutar la consulta
                RST_Busq RstBus, nSQL, xCon

                '--verificar si no se a registrado
                If RstBus.RecordCount = 0 Then
                    '--buscar el codigo del cliente
                    IdPer = NulosN(Busca_Codigo(NulosC(RstTmp("ruc")), "numruc", "id", "mae_cliente", "C", xCon))
                    '--si el cliente no esta registrado en el nuevo ejercicio, insertar el cliente
                    If IdPer = 0 Then
                        IdPer = HallaCodigoTabla("mae_cliente", xCon, "id")
                        xCon.Execute "insert into mae_cliente (id,tipper,idtipdoc,numruc,nombre) values (" & IdPer & ",2,5,'" & RstTmp("ruc") & "','" & RstTmp("nombre") & "')"
                    End If

                    '---------------------------------------------------------------------------
                    '--buscar si el codigo del documento existe, sino utilizar el mismo codigo
                    Set RstBus = Nothing
                    nSQL = "select id from vta_ventas where id=" & NulosN(RstTmp("id"))
                    RST_Busq RstBus, nSQL, xCon
                    If RstBus.RecordCount = 0 Then
                        '--si el codigo no esta en uso =>> asignar el codigo
                        xId = NulosN(RstTmp("id"))
                    Else
                        '--obtener el nuevo codigo cuando el codigo este en uso
                        xId = HallaCodigoTabla("vta_ventas", xCon, "id")
                    End If
                    '-----------------------------------------------------------------------------
                    '--agregando nuevo registro
                    RstCab.AddNew
                    '--colocando valores en los campos
                    RstCab("id") = xId
                    RstCab("idlib") = 2
                    RstCab("idtipo") = NulosN(RstTmp("idtipo"))
                    RstCab("tipdoc") = NulosN(RstTmp("tipdoc"))
                    RstCab("idcli") = IdPer
                    RstCab("numser") = NulosC(RstTmp("numser"))
                    RstCab("numdoc") = NulosC(RstTmp("numdoc"))
                    If IsDate(RstTmp("fchdoc")) = True Then RstCab("fchdoc") = CDate(RstTmp("fchdoc"))
                    If IsDate(RstTmp("fchven")) = True Then RstCab("fchven") = CDate(RstTmp("fchven"))
                    RstCab("idconpag") = NulosN(RstTmp("idconpag"))
                    RstCab("idmon") = NulosN(RstTmp("idmon"))
                    RstCab("impbru") = NulosN(RstTmp("impbru"))
                    RstCab("impbru2") = NulosN(RstTmp("impbru2"))
                    RstCab("impbru3") = NulosN(RstTmp("impbru3"))
                    RstCab("impinaf") = NulosN(RstTmp("impinaf"))
                    RstCab("impigv") = NulosN(RstTmp("impigv"))
                    RstCab("impisc") = NulosN(RstTmp("impisc"))
                    RstCab("impotr") = NulosN(RstTmp("impotr"))
                    'RstCab("impdesc") = NulosN(RstTmp("impdesc"))
                    RstCab("imptotdoc") = NulosN(RstTmp("imptotdoc"))
                    RstCab("impsal") = NulosN(RstTmp("impsal"))
                    RstCab("imptotori") = NulosN(RstTmp("impsal")) '--saldo final
                    RstCab("glosa") = NulosC(RstTmp("glosa"))
                    RstCab("idalm") = NulosN(RstTmp("idalm"))
                    RstCab("fchreg") = CDate("01/01" + "/" + AnoTra)
                    RstCab("numreg") = "000001"
                    RstCab("anulado") = 0
                    RstCab("iddocref") = 0
                    RstCab("idtipdocref") = 0
                    RstCab("iddocref2") = 0
                    RstCab("numerodocref") = NulosC(RstTmp("numerodocref"))
                    RstCab("tc") = 0
                    RstCab.Update
                    '--activar opcion de grabar operacion
                    GrabaOperacion = True

                End If

            Case 40 '--honorarios

                '******************
                RST_Busq RstCab, "SELECT TOP 1 * FROM com_honorarios", xCon

                '--buscar si ya esta ingresado el documento
                nSQL = "SELECT mae_prov.numruc AS ruc, com_honorarios.numser & '-' & com_honorarios.numdoc AS numerodoc, com_honorarios.tipdoc " _
                    + vbCr + " FROM mae_prov RIGHT JOIN com_honorarios ON mae_prov.id = com_honorarios.idpro " _
                    + vbCr + " WHERE (((mae_prov.numruc)='" & NulosC(RstTmp("ruc")) & "') AND ((com_honorarios.numser & '-' & com_honorarios.numdoc)='" & RstTmp("numerodoc") & "') AND ((com_honorarios.tipdoc)=" & RstTmp("tipdoc") & "));"
                '--ejecutar la consulta
                RST_Busq RstBus, nSQL, xCon

                '--verificar si no se a registrado
                If RstBus.RecordCount = 0 Then
                    '--buscar el codigo del proveedor
                    IdPer = NulosN(Busca_Codigo(NulosC(RstTmp("ruc")), "numruc", "id", "mae_prov", "C", xCon))
                    '--si el proveedor no esta registrado en el nuevo ejercicio, insertar el proveedor
                    If IdPer = 0 Then
                        IdPer = HallaCodigoTabla("mae_prov", xCon, "id")
                        xCon.Execute "insert into mae_prov (id,tipper,idtipdoc,numruc,nombre) values (" & IdPer & ",1,5,'" & RstTmp("ruc") & "','" & RstTmp("nombre") & "')"
                    End If

                    '---------------------------------------------------------------------------
                    '--buscar si el codigo del documento existe, sino utilizar el mismo codigo
                    Set RstBus = Nothing
                    nSQL = "select id from com_honorarios where id=" & NulosN(RstTmp("id"))
                    RST_Busq RstBus, nSQL, xCon
                    If RstBus.RecordCount = 0 Then
                        '--si el codigo no esta en uso =>> asignar el codigo
                        xId = NulosN(RstTmp("id"))
                    Else
                        '--obtener el nuevo codigo cuando el codigo este en uso
                        xId = HallaCodigoTabla("com_honorarios", xCon, "id")
                    End If
                    '-----------------------------------------------------------------------------
                    RstCab.AddNew
                    RstCab("id") = xId
                    RstCab("idlib") = 40
                    RstCab("idtipo") = NulosN(RstTmp("idtipo"))
                    RstCab("tipdoc") = NulosN(RstTmp("tipdoc"))
                    RstCab("idpro") = IdPer
                    RstCab("numser") = NulosC(RstTmp("numser"))
                    RstCab("numdoc") = NulosC(RstTmp("numdoc"))
                    If IsDate(RstTmp("fchdoc")) = True Then RstCab("fchdoc") = CDate(RstTmp("fchdoc"))
                    If IsDate(RstTmp("fchven")) = True Then RstCab("fchven") = CDate(RstTmp("fchven"))
                    RstCab("idconpag") = NulosN(RstTmp("idconpag"))
                    RstCab("idmon") = NulosN(RstTmp("idmon"))
                    RstCab("impbru") = NulosN(RstTmp("impbru"))
                    RstCab("impina") = NulosN(RstTmp("impina"))
                    RstCab("impigv") = NulosN(RstTmp("impigv"))
                    'RstCab("impdesc") = NulosN(RstTmp("impdesc"))
                    RstCab("imptotori") = NulosN(RstTmp("impsal")) '--saldo original
                    RstCab("glosa") = NulosC(RstTmp("glosa"))
                    RstCab("impsal") = NulosN(RstTmp("impsal"))
                    RstCab("idalm") = NulosN(RstTmp("idalm"))
                    RstCab("fchreg") = CDate("01/01" + "/" + AnoTra)
                    RstCab("numreg") = "000001"
                    RstCab("tc") = 0
                    RstCab("idtipdocref") = 0
                    RstCab("iddocref2") = 0
                    RstCab("numerodocref") = ""
                    RstCab.Update

                    '--activar opcion de grabar operacion
                    GrabaOperacion = True

                End If

            Case 41 '--Liquidacion de Gasto Debito

                RST_Busq RstCab, "SELECT TOP 1 * FROM vta_gastodebito", xCon

                '--buscar si ya esta ingresado el documento
                nSQL = "SELECT mae_cliente.numruc AS ruc, vta_gastodebito.numser & '-' & vta_gastodebito.numdoc AS numerodoc, vta_gastodebito.tipdoc " _
                    + vbCr + " FROM mae_cliente RIGHT JOIN vta_gastodebito ON mae_cliente.id = vta_gastodebito.idcli " _
                    + vbCr + " WHERE (((mae_cliente.numruc)='" & NulosC(RstTmp("ruc")) & "') AND ((vta_gastodebito.numser & '-' & vta_gastodebito.numdoc)='" & RstTmp("numerodoc") & "') AND ((vta_gastodebito.tipdoc)=" & RstTmp("tipdoc") & "));"

                '--ejecutar la consulta
                RST_Busq RstBus, nSQL, xCon

                '--verificar si no se a registrado
                If RstBus.RecordCount = 0 Then
                    '--buscar el codigo del cliente
                    IdPer = NulosN(Busca_Codigo(NulosC(RstTmp("ruc")), "numruc", "id", "mae_cliente", "C", xCon))
                    '--si el cliente no esta registrado en el nuevo ejercicio, insertar el cliente
                    If IdPer = 0 Then
                        IdPer = HallaCodigoTabla("mae_cliente", xCon, "id")
                        xCon.Execute "insert into mae_cliente (id,tipper,idtipdoc,numruc,nombre) values (" & IdPer & ",2,5,'" & RstTmp("ruc") & "','" & RstTmp("nombre") & "')"
                    End If

                    '---------------------------------------------------------------------------
                    '--buscar si el codigo del documento existe, sino utilizar el mismo codigo
                    Set RstBus = Nothing
                    nSQL = "select id from vta_gastodebito where id=" & NulosN(RstTmp("id"))
                    RST_Busq RstBus, nSQL, xCon
                    If RstBus.RecordCount = 0 Then
                        '--si el codigo no esta en uso =>> asignar el codigo
                        xId = NulosN(RstTmp("id"))
                    Else
                        '--obtener el nuevo codigo cuando el codigo este en uso
                        xId = HallaCodigoTabla("vta_gastodebito", xCon, "id")
                    End If
                    '-----------------------------------------------------------------------------

                    '--agregando nuevo registro
                    RstCab.AddNew
                    '--colocando valores en los campos
                    RstCab("id") = xId
                    RstCab("idlib") = 41
                    RstCab("tipdoc") = NulosN(RstTmp("tipdoc"))
                    RstCab("idcli") = IdPer
                    RstCab("numser") = NulosC(RstTmp("numser"))
                    RstCab("numdoc") = NulosC(RstTmp("numdoc"))
                    If IsDate(RstTmp("fchemi")) = True Then RstCab("fchemi") = CDate(RstTmp("fchemi"))
                    If IsDate(RstTmp("fchven")) = True Then RstCab("fchven") = CDate(RstTmp("fchven"))
                    RstCab("idmon") = NulosN(RstTmp("idmon"))
                    RstCab("impbru") = NulosN(RstTmp("impbru"))
                    RstCab("impina") = NulosN(RstTmp("impina"))
                    RstCab("igv") = NulosN(RstTmp("igv"))
                    RstCab("imptot") = NulosN(RstTmp("imptot"))
                    RstCab("impsal") = NulosN(RstTmp("impsal"))
''                    RstCab("imptotori") = NulosN(RstTmp("impsal")) '--saldo final
                    RstCab("glosa") = NulosC(RstTmp("glosa"))
                    RstCab("fchreg") = CDate("01/01" + "/" + AnoTra)
                    RstCab("numreg") = "000001"
                    RstCab("anulado") = 0
                    RstCab("idtipdocref") = 0
                    RstCab("iddocref2") = 0
                    RstCab("numerodocref") = NulosC(RstTmp("numerodocref"))

                    RstCab("tc") = 0
                    RstCab.Update

                    '--activar opcion de grabar operacion
                    GrabaOperacion = True

                End If

    '        Case 8 '--boleta de pago
    '        Case 42 '--planilla de letras
    '        Case 37 '--letras

            Case Else

                MsgBox "Falta implementar la opcion para importar  información del libro ", vbInformation, xTitulo

                Exit Sub
        End Select

        '-------------------------------------------------------------------------------
        'grabamos el movimiento en la tabla var_edicion
        If GrabaOperacion = True Then
            '--quehace=1::Agregar nuevo registro
            obj.GrabarOperacion IdUsuario, IdMenu, 6, xHoraInicio, Time, Date, xCon, CDbl(xId)
        End If
        '-------------------------------------------------------------------------------

        '--ir al sgte. registro
        RstTmp.MoveNext
    Loop


salir:
    '--limpiando los recodset
    Set RstBus = Nothing
    Set RstTmp = Nothing

    Set obj = Nothing

    Exit Sub
LaCague:
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"
    Err.Clear

    If xConTmp.State = 1 Then xConTmp.Close

    Set obj = Nothing

    Set RstBus = Nothing
    Set RstTmp = Nothing

End Sub

Public Function DirectorioPlanilla(Con As ADODB.Connection) As String
    '===================================================================================================
    'Creado : 03/05/10 Por: Johan Castro
    'Propósito: verificar si esta creado el directorio para las fotos y firmas
    '
    'Entradas:  Con= Conexion
    '
    'Resultados: Ruta del directorio
    '
    'Nota:       1.- Verificar si existe el directorio "\ARCHIVOS\RUC_EMPRESA\PERSONAL\"
    '            2.- Crear el directorio si no existe
    '===================================================================================================

    Dim xPath As String

    Set xCon = Con
    CargaDatosEmpresa

    '------------------------------------------------------------------------
    xPath = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTAAR", "RUTAS")
    '--verificar si existe la carpeta archivos
    If ArchivoExiste(xPath) = False Then
        MkDir (xPath)
    End If

    '--verificar si existe el directorio de la empresa
    xPath = xPath + NumRUC & "\"
    If ArchivoExiste(xPath) = False Then
        MkDir (xPath)
    End If

    '--verificar si existe el directorio de la empresa
    xPath = xPath + "PERSONAL\"
    If ArchivoExiste(xPath) = False Then
        MkDir (xPath)
    End If

    DirectorioPlanilla = xPath

End Function



Sub GrabarOperacionCtaCte(idlib As Integer, IdDoc As Double, Con As ADODB.Connection)
    '===================================================================================================
    'Creado : 17/09/10 Por: Johan Castro
    'Propósito: Grabar registro para analisis de cta cte
    '
    'Entradas:  IdLib= Codigo del libro
    '           IdDoc= Codigo del registro segun libro
    '           Con= Conexion a la base de datos
    '
    'Resultados: Registro en tabla var_analisisctacte
    '
    'Nota:       1.- Consultar el registro
    '            2.- Graba registro

    'Modificado: 20/09/10 Por: Johan Castro
    '           Agregar escritura de cta cte de lgd
    'Modificado: 05/10/10 Por:Johan Castro
    '           Agregar escritura de cta cte de Bancos (pendiente de correccion)
    'Modificado: 29/01/11 Johan Castro
    '           Modificar envio de parametro para ventas, compras, letras el tipo de cambio agregar funcion NulosN(xRst("tipcam"))
    '===================================================================================================

    '--
    Dim xRst As New ADODB.Recordset
    Dim Func As New SGI2_funciones.Varias
    Dim nSQL As String


    Set xCon = Con

    CargaDatosEmpresa


    Select Case idlib
        Case 2 '--Ventas
            '--Generar consulta
            nSQL = "SELECT mae_documentocta.idcuen as idcue, [vta_ventas].[numser] & '-' & [vta_ventas].[numdoc] AS numerodoc, mae_libros.codsun, Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) AS registro, IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc]) AS tipcam, vta_ventas.*  " _
                + vbCr + " FROM ((vta_ventas INNER JOIN mae_documentocta ON (vta_ventas.idmon = mae_documentocta.idmon) AND (vta_ventas.tipdoc = mae_documentocta.iddoc)) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
                + vbCr + " WHERE (((vta_ventas.id)=" & IdDoc & ") AND ((mae_documentocta.tipope)=-1) AND ((vta_ventas.numreg)<>'000001') AND ((vta_ventas.anulado)=0)); "

            RST_Busq xRst, nSQL, xCon

            If xRst.RecordCount <> 0 Then
                If xRst("tipdoc") <> 7 Then
                    If NulosN(xRst("idmon")) = 1 Then
                        Func.GrabarOperacionCtaCteDocRef 2, xRst("id"), NulosC(xRst("numerodocref")), xRst("idcli"), xRst("tipdoc"), xRst("numerodoc"), _
                            xRst("fchdoc"), NulosN(xRst("idmon")), NulosN(xRst("tipcam")), 0, xRst("imptotdoc"), 0, 0, xRst("registro"), xCon, 0, NulosC(xRst("glosa")), NulosN(xRst("idtipdocref")), NulosN(xRst("iddocref2")), 0, xRst("idcue")
                    Else
                        Func.GrabarOperacionCtaCteDocRef 2, xRst("id"), NulosC(xRst("numerodocref")), xRst("idcli"), xRst("tipdoc"), xRst("numerodoc"), _
                            xRst("fchdoc"), NulosN(xRst("idmon")), NulosN(xRst("tipcam")), 0, 0, 0, xRst("imptotdoc"), xRst("registro"), xCon, 0, NulosC(xRst("glosa")), NulosN(xRst("idtipdocref")), NulosN(xRst("iddocref2")), 0, xRst("idcue")
                    End If
                Else
                    If NulosN(xRst("idmon")) = 1 Then
                        Func.GrabarOperacionCtaCteDocRef 2, xRst("id"), NulosC(xRst("numerodocref")), xRst("idcli"), xRst("tipdoc"), xRst("numerodoc"), _
                            xRst("fchdoc"), NulosN(xRst("idmon")), NulosN(xRst("tipcam")), xRst("imptotdoc"), 0, 0, 0, xRst("registro"), xCon, 0, NulosC(xRst("glosa")), NulosN(xRst("idtipdocref")), NulosN(xRst("iddocref2")), 0, xRst("idcue")
                    Else
                        Func.GrabarOperacionCtaCteDocRef 2, xRst("id"), NulosC(xRst("numerodocref")), xRst("idcli"), xRst("tipdoc"), xRst("numerodoc"), _
                            xRst("fchdoc"), NulosN(xRst("idmon")), NulosN(xRst("tipcam")), 0, 0, xRst("imptotdoc"), 0, xRst("registro"), xCon, 0, NulosC(xRst("glosa")), NulosN(xRst("idtipdocref")), NulosN(xRst("iddocref2")), 0, xRst("idcue")
                    End If
                End If

            End If

        Case 6 '--bancos
''            Dim RstTes As New ADODB.Recordset
''            Dim RstTesdes As New ADODB.Recordset
''            Dim RstTesDesDet As New ADODB.Recordset
''            Dim xId As Double
''            Dim A, B As Integer
''            Dim xTc As Double
''            Dim xCodSunLib As String
''            Dim xIdCliente As Integer
''            Dim xImporte As Double
''
''            Dim xNumeroDocRef As String
''            Dim xIdPer As Integer
''
''            RstMov.MoveFirst
''        '    RstMov.Filter = "id =47169"
''            For B = 1 To RstMov.RecordCount
''                xId = RstMov("id")
''
''                If xId = 34011 Then
''                    MsgBox ""
''                End If
''                xCodSunLib = Busca_Codigo(6, "id", "codsun", "mae_libros", "N", xCon)
''
''                ' OBTENEMOS EL TC DE LA OPERACION
''                xTc = NulosN(Busca_Codigo(xId, "idtes", "tc", "tes_cajadestino", "N", xCon))
''                If xTc = 0 Then   'SI EL TIPO DE CAMBIO ES 0 OBTENEMOS EL TIPO DE CAMBIO DE LA TABLA con_tc
''                    xTc = HallaTipoCambio(RstMov("fchope"), 2, Venta, xCon)
''                End If
''
''                RST_Busq RstTesDesDet, "SELECT tes_cajadestinodet.*, tes_cajadestinodet.iddes From tes_cajadestinodet Where (((tes_cajadestinodet.idtes) = " & xId & ")) " _
''                    & " ORDER BY tes_cajadestinodet.iddes", xCon
''
''                If RstTesDesDet.RecordCount <> 0 Then
''                    RstTesDesDet.MoveFirst
''
''                    For A = 1 To RstTesDesDet.RecordCount
''                        Dim xTipDoc As Integer
''                        Dim xNumDoc As String
''                        xImporte = 0
''                        If RstTesDesDet("idmod") = 2 Then   ' SI ES MODULO DE VENTAS
''                            If RstTesDesDet("iddoc") <> 0 Then
''                                xTipDoc = NulosN(Busca_Codigo(RstTesDesDet("iddoc"), "id", "tipdoc", "vta_ventas", "N", xCon))
''                                xNumDoc = NulosC(Busca_Codigo(RstTesDesDet("iddoc"), "id", "numser", "vta_ventas", "N", xCon))
''                                xNumDoc = xNumDoc & "-" & NulosC(Busca_Codigo(RstTesDesDet("iddoc"), "id", "numdoc", "vta_ventas", "N", xCon))
''                                xNumeroDocRef = NulosN(Busca_Codigo(RstTesDesDet("iddoc"), "id", "numerodocref", "vta_ventas", "N", xCon))
''                                xIdPer = Busca_Codigo(RstTesDesDet("iddoc"), "id", "idcli", "vta_ventas", "N", xCon)
''                                xImporte = RstTesDesDet("acuenta")
''                            Else
''                                xTipDoc = 0
''                                xNumDoc = ""
''                                xNumeroDocRef = NulosC(RstTesDesDet("numerodocref"))
''                                xIdPer = NulosN(RstTesDesDet("idper"))
''                            End If
''                        End If
''                        If RstTesDesDet("idmod") = 11 Then   ' SI ES MODULO DE LIQUIDACION GASTO DEBITO
''                            xTipDoc = NulosN(Busca_Codigo(RstTesDesDet("iddoc"), "id", "tipdoc", "vta_gastodebito", "N", xCon))
''                            xNumDoc = Busca_Codigo(RstTesDesDet("iddoc"), "id", "numser", "vta_gastodebito", "N", xCon)
''                            xNumDoc = xNumDoc & "-" & Busca_Codigo(RstTesDesDet("iddoc"), "id", "numdoc", "vta_gastodebito", "N", xCon)
''
''                            xNumeroDocRef = NulosN(Busca_Codigo(RstTesDesDet("iddoc"), "id", "numerodocref", "vta_gastodebito", "N", xCon))
''                            xIdPer = NulosN(Busca_Codigo(RstTesDesDet("iddoc"), "id", "idcli", "vta_gastodebito", "N", xCon))
''
''
''                            xImporte = RstTesDesDet("acuenta")
''                        End If
''
''                        If RstTesDesDet("idmod") = 6 Then   ' SI ES MODULO DE LIQUIDACION GASTO DEBITO
''                            xTipDoc = NulosN(RstTesDesDet("tipdoc")) 'Busca_Codigo(RstTesDesDet("iddoc"), "id", "tipdoc", "vta_gastodebito", "N", xCon)
''                            xNumDoc = RstTesDesDet("numser") & "-" & RstTesDesDet("numdoc") 'Busca_Codigo(RstTesDesDet("iddoc"), "id", "numser", "vta_gastodebito", "N", xCon)
''                            xImporte = RstTesDesDet("importe")
''
''                            xNumeroDocRef = NulosC(RstTesDesDet("numerodocref"))
''                            xIdPer = NulosN(RstTesDesDet("idper"))
''                        End If
''
''                        If RstTesDesDet("idmod") = 19 Then
''                          '  MsgBox ""
''                        End If
''                        If RstTesDesDet("idmod") = 2 Or RstTesDesDet("idmod") = 11 Or RstTesDesDet("idmod") = 6 Then
''                            If NulosN(RstMov("idmon")) = 1 Then
''                                GrabarOperacionCtaCteDocRef 6, xId, xNumeroDocRef, xIdPer, xTipDoc, xNumDoc, _
''                                    RstMov("fchope"), RstMov("idmon"), xTc, xImporte, 0, 0, 0, Format(xCodSunLib, "00") & RstMov("numreg"), xCon, RstTesDesDet("corr"), NulosC(RstTesDesDet("docctacte"))
''                            Else
''                                Func.GrabarOperacionCtaCteDocRef 6, xId, xNumeroDocRef, xIdPer, xTipDoc, xNumDoc, _
''                                    RstMov("fchope"), RstMov("idmon"), xTc, 0, 0, xImporte, 0, Format(xCodSunLib, "00") & RstMov("numreg"), xCon, RstTesDesDet("corr"), NulosC(RstTesDesDet("docctacte"))
''                            End If
''                        End If
''                        RstTesDesDet.MoveNext
''                        If RstTesDesDet.EOF = True Then Exit For
''                    Next A
''                Else
''                    'MsgBox ""
''                End If
''
''                RstMov.MoveNext
''                If RstMov.EOF = True Then Exit For
''            Next B


        Case 41 '--Liquidacion de Gasto debito
            nSQL = "SELECT mae_documentocta.idcuen as idcue, [vta_gastodebito].[numser] & '-' & [vta_gastodebito].[numdoc] AS numerodoc, mae_libros.codsun, Left([vta_gastodebito].[numreg],2) & [mae_libros].[codsun] & Right([vta_gastodebito].[numreg],4) AS registro, IIf([vta_gastodebito].[tc]=0,[con_tc].[impven],[vta_gastodebito].[tc]) AS tipcam, vta_gastodebito.*  " _
                + vbCr + " FROM ((vta_gastodebito INNER JOIN mae_documentocta ON (vta_gastodebito.idmon = mae_documentocta.idmon) AND (vta_gastodebito.tipdoc = mae_documentocta.iddoc)) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha " _
                + vbCr + " WHERE (((vta_gastodebito.id)=" & IdDoc & ") AND ((mae_documentocta.tipope)=-1) AND ((vta_gastodebito.numreg)<>'000001') AND ((vta_gastodebito.anulado)=0)); "


            RST_Busq xRst, nSQL, xCon

            If xRst.RecordCount <> 0 Then
'
                If xRst("tipdoc") = 120 Then   ' LIQUIDACION GASTO DEBITO
                    If NulosN(xRst("idmon")) = 1 Then
                        Func.GrabarOperacionCtaCteDocRef 41, xRst("id"), NulosC(xRst("numerodocref")), xRst("idcli"), xRst("tipdoc"), xRst("numerodoc"), _
                            xRst("fchemi"), xRst("idmon"), NulosN(xRst("tipcam")), 0, 0, xRst("imptot"), 0, xRst("registro"), xCon, 0, NulosC(xRst("glosa")), xRst("idtipdocref"), xRst("iddocref2"), 0, xRst("idcue")
                    Else
                        Func.GrabarOperacionCtaCteDocRef 41, xRst("id"), NulosC(xRst("numerodocref")), xRst("idcli"), xRst("tipdoc"), xRst("numerodoc"), _
                            xRst("fchemi"), xRst("idmon"), NulosN(xRst("tipcam")), 0, 0, 0, xRst("imptot"), xRst("registro"), xCon, 0, NulosC(xRst("glosa")), xRst("idtipdocref"), xRst("iddocref2"), 0, xRst("idcue")
                    End If
                End If

                If xRst("tipdoc") = 126 Then   ' LIQUIDACION GASTO CREDITO
                    If NulosN(xRst("idmon")) = 1 Then
                        Func.GrabarOperacionCtaCteDocRef 41, xRst("id"), NulosC(xRst("numerodocref")), xRst("idcli"), xRst("tipdoc"), xRst("numerodoc"), _
                            xRst("fchemi"), xRst("idmon"), xRst("tipcam"), 0, xRst("imptot"), 0, 0, xRst("registro"), xCon, 0, NulosC(xRst("glosa")), xRst("idtipdocref"), xRst("iddocref2"), 0, xRst("idcue")
                    Else
                        Func.GrabarOperacionCtaCteDocRef 41, xRst("id"), NulosC(xRst("numerodocref")), xRst("idcli"), xRst("tipdoc"), xRst("numerodoc"), _
                            xRst("fchemi"), xRst("idmon"), xRst("tipcam"), 0, 0, xRst("imptot"), 0, xRst("registro"), xCon, 0, NulosC(xRst("glosa")), xRst("idtipdocref"), xRst("iddocref2"), 0, xRst("idcue")
                    End If
                End If

            End If

        Case 37 '--letras

            nSQL = "SELECT [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numerodoc, mae_letra.idcuenven AS idcue, Left([let_letra].[numreg],2) & [mae_libros].[codsun] & Right([let_letra].[numreg],4) AS registro, IIf([let_letra].[tc]=0,[con_tc].[impven],[let_letra].[tc]) AS tipcam, let_letra.idclipro AS idcli, let_letradet.fchemi, let_letradet.fchven, let_letradet.implet, let_letradet.idlet, let_letradet.corr, let_letra.idtipdocref, let_letra.iddocref2, let_letra.numerodocref, let_letra.glosa, let_letra.tipdoc, let_letra.idmon " _
                + vbCr + " FROM ((mae_libros RIGHT JOIN (let_letra INNER JOIN let_letradet ON let_letra.id = let_letradet.idlet) ON mae_libros.id = let_letra.idlib) LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha) LEFT JOIN mae_letra ON let_letra.idmon = mae_letra.idmon " _
                + vbCr + " WHERE (((let_letradet.idlet)=" & IdDoc & ")); "

            RST_Busq xRst, nSQL, xCon

            If xRst.RecordCount <> 0 Then
                xRst.MoveFirst
                '--recorrer cada letra
                Do While Not xRst.EOF

                    If NulosN(xRst("idmon")) = 1 Then
                        Func.GrabarOperacionCtaCteDocRef 37, xRst("idlet"), NulosC(xRst("numerodocref")), xRst("idcli"), xRst("tipdoc"), xRst("numerodoc"), _
                            xRst("fchemi"), xRst("idmon"), NulosN(xRst("tipcam")), xRst("implet"), 0, 0, 0, xRst("registro"), xCon, xRst("corr"), NulosC(xRst("glosa")), xRst("idtipdocref"), xRst("iddocref2"), 0, xRst("idcue")
                    Else
                        Func.GrabarOperacionCtaCteDocRef 37, xRst("idlet"), NulosC(xRst("numerodocref")), xRst("idcli"), xRst("tipdoc"), xRst("numerodoc"), _
                            xRst("fchemi"), xRst("idmon"), NulosN(xRst("tipcam")), 0, 0, xRst("implet"), 0, xRst("registro"), xCon, xRst("corr"), NulosC(xRst("glosa")), xRst("idtipdocref"), xRst("iddocref2"), 0, xRst("idcue")
                    End If

                    xRst.MoveNext
                Loop
            End If
        Case Else

    End Select


    Set Func = Nothing

'    MsgBox "se termino con exito"
End Sub






