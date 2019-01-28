VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "formularios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function CentroCostoSeleBloques(Coneccion As ADODB.Connection) As ADODB.Recordset
    Set xCon = Coneccion
    FrmSelCentroCosto2.Show vbModal
    If FrmSelCentroCosto2.EnviarRST = True Then
        Set CentroCostoSeleBloques = FrmSelCentroCosto2.RstDat
    Else
        Set CentroCostoSeleBloques = Nothing
    End If
    
    Unload FrmSelCentroCosto
End Function

Function SeleCentroCosto(Coneccion As ADODB.Connection) As ADODB.Recordset
    Set xCon = Coneccion
    FrmSelCentroCosto.Show vbModal
    Set SeleCentroCosto = FrmSelCentroCosto.RstSele
End Function

Function SelePlanCuentas(Coneccion As ADODB.Connection) As ADODB.Recordset
    Set xCon = Coneccion
    FrmSelPlanContable.Show vbModal
    If FrmSelPlanContable.RstSele.RecordCount = 0 Then
        Set FrmSelPlanContable.RstSele = Nothing
    End If
    Set SelePlanCuentas = FrmSelPlanContable.RstSele
End Function

Sub DocumentoPagarCobrar(Coneccion As ADODB.Connection)
    Set xCon = Coneccion
    FrmDocCobPag.Show vbModal
End Sub

Public Sub Imprimir_x_VSFlexGrid(GRID As Object, _
                TITULO As String, _
                Optional TITULO_1 As String = "", _
                Optional PERIODO As String = "", _
                Optional TITULO_EN_HOJAS As Boolean = True, _
                Optional ENCABEZADO_EN_HOJAS As Boolean = True)
    '--GRID = VSFlexGrid
    '--IMPRIMIR DATOS DE UN FLEXGRID
'    On Error GoTo ERROR
    
    If GRID.Rows - GRID.FixedRows = 0 Then
        MsgBox "No hay Registros a Imprimir", vbInformation, xTitulo
        Exit Sub
    End If
    
    
    Dim oFrm As New FrmVSFlexGrid_Imprimir
    oFrm.MousePointer = vbHourglass
    Err.Clear
    oFrm.Show
    oFrm.PONER_DATOS GRID, TITULO, TITULO_1, PERIODO, TITULO_EN_HOJAS, ENCABEZADO_EN_HOJAS
    oFrm.SetFocus
    Set oFrm = Nothing
    Exit Sub
ERROR:
    If Err.Number = 401 Then
        oFrm.PONER_DATOS GRID, TITULO, TITULO_1, PERIODO, TITULO_EN_HOJAS, ENCABEZADO_EN_HOJAS
        oFrm.Show 1
    Else
         MsgBox Err.Description + vbCr + "Nº.: " + Err.Number, vbCritical, "Error..."
    End If
    Err.Clear
End Sub

Public Sub Imprimir_x_VSFlexGrid_GRID_EN_RPT(GRID As Object)
    '--GRID = VSFlexGrid
    '--EL OBJETO SE COLOCA EN EL REPORTE
    FrmVSFlexGrid_Imprimir.MousePointer = vbHourglass
    FrmVSFlexGrid_Imprimir.PONER_GRID_EN_RPT GRID
    FrmVSFlexGrid_Imprimir.Show 1
    
End Sub

Public Sub VSFlexGrid_Exportar_MSExcel(Con As ADODB.Connection, _
                                        Optional GRID As Object, _
                                        Optional TITULO As String, _
                                        Optional PERIODO As String = "", _
                                        Optional TITULO_1 As String = "", _
                                        Optional NOMBRE_A_EXPORTAR As String = "", _
                                        Optional RstTmp As ADODB.Recordset, _
                                        Optional xCampos)
                
    '--GRID = VSFlexGrid
    '--IMPRIMIR DATOS DE UN FLEXGRID
    If IsArray(xCampos) = False Then
'    If RstTmp.State = 0 Then
        If GRID.Rows - GRID.FixedRows = 0 Then
            MsgBox "No hay Registros a exportar", vbInformation, xTitulo
            Exit Sub
        End If
    Else
        If RstTmp.RecordCount = 0 Then
            MsgBox "No hay Registros a exportar", vbInformation, xTitulo
            Exit Sub
        End If
        If RstTmp.Fields.Count - 1 < UBound(xCampos) Then
            MsgBox "La lista de campos a mostrar es incompatible" & vbCr & "Total Campos Array: " & UBound(xCampos) & vbCr & "Total Campos Rst: " & RstTmp.Fields.Count, vbInformation, xTitulo
            Exit Sub
        End If
    End If
    
    
    Set xCon = Con
    CargaDatosEmpresa
    Dim oFrm As New FrmVSFlexGrid_MSExcel
    Err.Clear
    On Error GoTo ERROR
    oFrm.Show
    oFrm.RECIBE_LINK_EXPORT GRID, TITULO, PERIODO, TITULO_1, NOMBRE_A_EXPORTAR, RstTmp, xCampos
    Set oFrm = Nothing
    Exit Sub
ERROR:
    If Err.Number = 401 Then
        oFrm.RECIBE_LINK_EXPORT GRID, TITULO, PERIODO, TITULO_1, NOMBRE_A_EXPORTAR
        oFrm.Show 1
    Else
         MsgBox Err.Description + vbCr + "Nº.: " + Err.Number, vbCritical, "Error..."
    End If
    Err.Clear
    Set oFrm = Nothing
End Sub


Public Sub VSFlexGrid_Buscar(hWndFrmPadre As Long, GRID As Object, xCampos() As String, Optional FILA_INICIO As Long = -1)
                
    '--GRID = VSFlexGrid
    '--IMPRIMIR DATOS DE UN FLEXGRID
    If GRID.Rows - GRID.FixedRows = 0 Then
        MsgBox "No hay Registros a Buscar ", vbInformation, xTitulo
        Exit Sub
    End If
    '--VALIDAR EL ARRAY
    Dim Q_ROW&
    Dim F_ESTADO As Boolean
    For Q_ROW = 0 To UBound(xCampos())
        If xCampos(Q_ROW, 0) = "" Then
            F_ESTADO = True: Exit For
        ElseIf xCampos(Q_ROW, 1) = "" Then
            F_ESTADO = True: Exit For
        ElseIf xCampos(Q_ROW, 2) = "" Then
            F_ESTADO = True: Exit For
        End If
    Next
    If F_ESTADO = True Then
        MsgBox "Error en el paso de parametros", vbCritical, "Buscar..."
        Exit Sub
    End If
    '--DEL INICIO DE BUSQUEDA
    If FILA_INICIO = -1 Then
        FILA_INICIO = GRID.FixedRows + 1
    End If
    
    FrmVSFlexGrid_Buscar.RECIBE_LINK_SEARCH hWndFrmPadre, GRID, xCampos(), FILA_INICIO
    FrmVSFlexGrid_Buscar.Show
    
End Sub

Public Sub FechaSeleccionar(obj As Object, mRow As Long, mCol As Long, nValor As String, Optional eTipo As e_TipoFrmFecha = e_Escribir)
    Dim pt As POINTAPI
    
    If TypeName(obj) = "VSFlexGrid" Then
        pt.X = obj.ColPos(mCol) \ Screen.TwipsPerPixelX
        pt.Y = (obj.RowPos(mRow) + obj.RowHeight(mRow)) \ Screen.TwipsPerPixelY
    ElseIf TypeName(obj) = "TextBox" Or TypeName(obj) = "DTPicker" Then
        pt.X = Screen.TwipsPerPixelX - 15
        pt.Y = Screen.TwipsPerPixelY
    Else
        MsgBox "Falta implementar...", vbCritical, xTitulo
        Exit Sub
        '--si agregan mas condiciones cambiar frmSelFecha.Form_Deactivate
    End If
    ClientToScreen obj.hwnd, pt
    
    If eTipo = e_Escribir Then
        With frmSelFecha
            '.Moveable = MoverForm
            .lblRow = mRow
            .lblCol = mCol
            
            If TypeName(obj) = "VSFlexGrid" Then
                .Tag = obj.Cell(flexcpText, mRow, mCol)
            ElseIf TypeName(obj) = "TextBox" Then
                .Tag = obj.Text
            ElseIf TypeName(obj) = "DTPicker" Then
                .Tag = obj.Value
            End If
            
            .Move pt.X * Screen.TwipsPerPixelX, pt.Y * Screen.TwipsPerPixelY
            .pRecibeLink obj
            .Show 1
        End With
    ElseIf eTipo = e_Seleccion Then
        With frmSelFecha1
            '.Moveable = MoverForm
            .lblRow = mRow
            .lblCol = mCol
            
            If TypeName(obj) = "VSFlexGrid" Then
                .Tag = obj.Cell(flexcpText, mRow, mCol)
            ElseIf TypeName(obj) = "TextBox" Then
                .Tag = obj.Text
            ElseIf TypeName(obj) = "DTPicker" Then
                .Tag = obj.Value
            End If
            
            .Move pt.X * Screen.TwipsPerPixelX, pt.Y * Screen.TwipsPerPixelY
            .pRecibeLink obj
            .Show 1
        End With
        
    End If
End Sub

Public Sub HoraSeleccionar(obj As Object, mRow As Long, mCol As Long, nValor As String)
    Dim pt As POINTAPI
    
    If TypeName(obj) = "VSFlexGrid" Then
        pt.X = obj.ColPos(mCol) \ Screen.TwipsPerPixelX
        pt.Y = (obj.RowPos(mRow) + obj.RowHeight(mRow)) \ Screen.TwipsPerPixelY
    ElseIf TypeName(obj) = "TextBox" Or TypeName(obj) = "DTPicker" Then
        pt.X = Screen.TwipsPerPixelX - 15
        pt.Y = Screen.TwipsPerPixelY
    Else
        MsgBox "Falta implementar...", vbCritical, xTitulo
        Exit Sub
        '--si agregan mas condiciones cambiar frmSelHora.Form_Deactivate
    End If
    ClientToScreen obj.hwnd, pt
    
    With frmSelHora
        '.Moveable = MoverForm
        .lblRow = mRow
        .lblCol = mCol
        .Tag = nValor
        .Move pt.X * Screen.TwipsPerPixelX, pt.Y * Screen.TwipsPerPixelY
        .pRecibeLink obj
        .Show 1
    End With
    Unload frmSelHora
End Sub


Public Sub PreciosHistoricos(Coneccion As ADODB.Connection, IdItem As String, TipoVentana As Boolean, Optional NumeroRuc As String = "")
    Set xCon = Coneccion
    FrmAnalizaPrecio.RECIBE_ID_ITEM IdItem, TipoVentana, NumeroRuc
    FrmAnalizaPrecio.Show
End Sub

Private Sub Class_Initialize()
    Nomsis = LeerLineaINI(Trim(App.Path) + "\seven.ini", "NOMBRE", "SOFTWARE")
    AP_RUTABD = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    AP_RUTASY = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS")
    AP_RUTABM = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")

End Sub



'----------
Public Sub AsientoVer(xCon As ADODB.Connection, NumReg As String)
    '--numreg debe ser completo mes + cod sunat + asiento
    '--23/02/09
    FrmAsientoVer.pRecibeLink xCon, NumReg
    FrmAsientoVer.Show 1
    Unload FrmAsientoVer
End Sub


Public Sub AsientoVerTmp(xCon As ADODB.Connection, xRst As ADODB.Recordset, idlib As Integer, Optional idmov As Double = 0)
    '--17/11/09
    FrmAsientoVer.pRecibeLinkTmp xCon, xRst, idlib, idmov
    FrmAsientoVer.Show 1
    Unload FrmAsientoVer
End Sub



Public Sub AsientoRstDefine(Rst As ADODB.Recordset)
    '--20/02/09
    FrmAsientoVer.fDefinirRst Rst
    Unload FrmAsientoVer
    
End Sub

Public Sub AsientoBuscar(Coneccion As ADODB.Connection)
    '--20/02/09
    Set xCon = Coneccion
    FrmAsientoBuscar.Show
    
    
End Sub




