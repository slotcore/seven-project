VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Punto de Venta - Control de Acceso"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      Height          =   420
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   3375
      Width           =   1590
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   420
      Index           =   0
      Left            =   1155
      TabIndex        =   6
      Top             =   3375
      Width           =   1590
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   -15
      TabIndex        =   9
      Top             =   15
      Width           =   5610
      Begin VB.CommandButton cmd 
         Caption         =   "&Verficar Usuario"
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         Left            =   3465
         TabIndex        =   2
         Top             =   525
         Width           =   1365
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   2
         Left            =   1155
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "txt(2)"
         Top             =   525
         Width           =   2205
      End
      Begin VB.CommandButton cb 
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   1845
         Picture         =   "FrmLongin.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1140
         Width           =   240
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   1
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txt(1)"
         Top             =   135
         Width           =   2205
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0000FFFF&
         Height          =   330
         Index           =   0
         Left            =   4245
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "txt(0)"
         Top             =   135
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox txt_cb 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1155
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "txt_cb(0)"
         Top             =   1110
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   645
         Width           =   810
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   5625
         Y1              =   1485
         Y2              =   1470
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   5580
         X2              =   5595
         Y1              =   0
         Y2              =   1950
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   30
         Y1              =   75
         Y2              =   1935
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   -30
         X2              =   5580
         Y1              =   30
         Y2              =   15
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   5430
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   1215
         Width           =   615
      End
      Begin VB.Label lbl_cb 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb(0)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   2175
         TabIndex        =   13
         Top             =   1110
         Width           =   3195
      End
      Begin VB.Label lbl_cb_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb_cod(0)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   3870
         TabIndex        =   12
         Top             =   1110
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   270
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   3645
         TabIndex        =   10
         Top             =   285
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   1725
      Left            =   -15
      TabIndex        =   5
      Top             =   1545
      Width           =   5610
      _cx             =   9895
      _cy             =   3043
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLongin.frx":0132
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'-----------------------------
'-----------------------------

Private Sub cb_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Select Case Index
        Case 0 '--ALMACEN
            ReDim xCampos(1, 3) As String
            xCampos(0, 0) = "Almacén":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            
            nSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion AS nombre, alm_almacenes.id AS cod " _
            + vbCr + " FROM alm_almacenes ORDER BY alm_almacenes.descripcion ;"
        
    End Select
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Almacén", "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
Salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--ACEPTAR
            
            '--almacenar el codigo del almacen, para hacer todas la operaciones en funcion al almacen seleccionado
            '---------------------
            mIdAlmacen = NulosN(lbl_cb_cod(0).Caption)
            mIdEmpleado = NulosN(txt(0).Text)

            '----------
            pCargarArray
            '**********************************
            FrmMenu.Show '--POR CAMBIAR
            FrmMenu.SetFocus
            '**********************************
            Unload Me
            Exit Sub
        Case 1 '--CANCELAR
            Unload Me
            Exit Sub
        Case 2 '--VALIDAR USUARIO
            If fControlAcceso() = False Then
                Exit Sub
            End If
            cmd(0).SetFocus
            
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    LimpiaText txt, True
    txt_cb(0).Text = ""
    '------
    GRID_COMBOLIST Fg1, 4
    GRID_COMBOLIST Fg1, 5
    '------
    OCULTAR_COL Fg1, 1, 2
    
End Sub

Private Sub txt_Change(Index As Integer)
    If txt(Index) = "" Then
        txt_cb(0).Text = ""
        Fg1.Rows = 1
        cmd(0).Enabled = False
    End If
    If Trim(txt(1).Text) <> "" And Trim(txt(2).Text) <> "" Then
        cmd(2).Enabled = True
    Else
        cmd(2).Enabled = False
    End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub


Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        lbl_cb(Index).Caption = ""
        lbl_cb_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    On Error GoTo error
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If

    If txt_cb(Index).Text = "" Then Exit Sub
    If KeyCode <> 13 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    Select Case Index
        Case 0 '--ALMACEN
            nSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion AS nombre, alm_almacenes.id AS cod " _
                + vbCr + " FROM alm_almacenes WHERE alm_almacenes.id = " + CStr(NulosN(Trim(txt_cb(Index).Text))) + " ;"
    End Select

    If xCon.State = 0 Then Exit Sub
    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.State = 0 Then Exit Sub
    If RstTmp.RecordCount > 0 Then
        txt_cb(Index) = RstTmp.Fields(0) & "" '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
    End If
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
    If KeyAscii = 46 Then KeyAscii = 0
End Sub



Private Function fControlAcceso() As Boolean
    Dim RstTmp As New ADODB.Recordset
    Dim IdUsuario As Integer
    Dim IdEmpleado As Integer
    Dim nSQL As String
    '---------
    If txt(1).Text = "" Then
        MsgBox "Ingrese el Login", vbExclamation, xTitulo
        txt(1).SetFocus
        Exit Function
    ElseIf txt(2).Text = "" Then
        MsgBox "Ingrese el Password", vbExclamation, xTitulo
        txt(2).SetFocus
        Exit Function
    End If
    txt_cb(0).Text = ""
    Fg1.Rows = 1
    '---------
    nSQL = "SELECT mae_usuarios.id AS userid, mae_usuarios.login, mae_usuarios.pass,mae_usuarios.idemp AS empid " _
        + vbCr + " FROM mae_usuarios " _
        + vbCr + " WHERE ((UCASE(mae_usuarios.login)='" + UCase(txt(1).Text) + "') AND (UCASE(mae_usuarios.pass)='" + UCase(txt(2).Text) + "'));"

    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.BOF = True And RstTmp.EOF = True And RstTmp.RecordCount = 0 Then

        LimpiaText txt
        MsgBox "El Usuario no Existe", vbInformation, xTitulo
        txt(1).SetFocus
        Exit Function
    End If
    IdUsuario = NulosN(RstTmp.Fields("userid"))
    IdEmpleado = NulosN(RstTmp.Fields("empid"))
    
    Set RstTmp = Nothing
    '---------
    nSQL = "SELECT pla_empleados.id AS empid, [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS empdesc, pvt_emp.id AS pvtid, pvt_emp.idalm AS almid, alm_almacenes.descripcion AS almdesc, pvt_emp.idemp " _
        + vbCr + " FROM (pvt_emp INNER JOIN pla_empleados ON pvt_emp.idemp = pla_empleados.id) LEFT JOIN alm_almacenes ON pvt_emp.idalm = alm_almacenes.id " _
        + vbCr + " WHERE (((pvt_emp.idemp)=" + CStr(IdEmpleado) + "));"
        
    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        '--DEL ALMACEN
        txt_cb(0).Enabled = True
        cb(0).Enabled = True
        '-----
        cmd(0).Enabled = True '--BOTON ACEPTAR
        '-----
        txt_cb(0).Text = RstTmp.Fields("almid") & ""
        lbl_cb(0).Caption = RstTmp.Fields("almdesc") & ""
        lbl_cb_cod(0).Caption = RstTmp.Fields("almid") & ""
        '------
        txt(0).Text = RstTmp.Fields("pvtid") & "" '--CODIGO DEL PERSONAL SEGUN => PVT_EMP.ID
        
        
        pCargarDatosGrid
        
        txt_cb(0).SetFocus '--ALMACEN
        
    Else
        
        Exit Function
        
    End If
    Set RstTmp = Nothing
    fControlAcceso = True
End Function


Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col <> 4 And Col <> 5 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nTitulo As String
    
    If NulosN(lbl_cb_cod(0).Caption) = 0 And Col = 3 Then
        MsgBox "Seleccione el Almacén" + vbCr + "Luego proceda a Seleccionar el Nº de Serie.", vbExclamation, xTitulo
        Fg1.TextMatrix(Row, Col) = ""
        txt_cb(0).SetFocus:     Exit Sub
    End If
    
    Select Case Col
        Case 4 '--DE LAS SERIES
            ReDim xCampos(1, 3) As String
            xCampos(0, 0) = "Número":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            
            nSQL = "SELECT Format([alm_numseries].[numser],'0000') AS nombre,alm_numseries.id ,alm_numseries.id AS cod " _
                + vbCr + " FROM alm_numseries " _
                + vbCr + " WHERE alm_numseries.idtipdoc=" + CStr(Fg1.TextMatrix(Row, 1)) + " AND alm_numseries.idalm=" + CStr(NulosN(lbl_cb_cod(0).Caption)) + ";"
            
            nTitulo = "Buscando Series"
        
        Case 5 '--PLANTILLA DE IMPRESION
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Nombre":           xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4000":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Tipo de Letra":    xCampos(1, 1) = "tipoletra": xCampos(1, 2) = "1800":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Tamaño":           xCampos(2, 1) = "zise":      xCampos(2, 2) = "1000":   xCampos(2, 3) = "N"
            
            nSQL = "SELECT var_plantilladoc.id, var_plantilladoc.descripcion as nombre ,var_plantilladoc.id as cod, var_plantilladoc.tipoletra, var_plantilladoc.tamañoletra AS zise " _
                + vbCr + " FROM var_plantilladoc " _
                + vbCr + " WHERE (((var_plantilladoc.tipdoc) = " + CStr(Fg1.TextMatrix(Row, 1)) + ")) " _
                + vbCr + " ORDER BY var_plantilladoc.descripcion;"
            
            nTitulo = "Buscando Plantillas de Impresión"
               
    End Select

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio, ""

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    Agregando = True
    Fg1.TextMatrix(Row, Col) = xRs.Fields("nombre") & ""
    Agregando = False
    Set xRs = Nothing
    Exit Sub
Salir:
    Set xRs = Nothing
    Agregando = False
End Sub


Private Sub Fg1_EnterCell()
    If Fg1.Col = 4 Or Fg1.Col = 5 Then
        
        Fg1.Editable = flexEDKbdMouse
    Else
        
        Fg1.Editable = flexEDNone
        
    End If
End Sub

Private Sub Fg1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then
'        cmd_item(0).SetFocus
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    Select Case Col
        Case 5, 6, 7
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
'        REGISTRO_ADD
    End If

End Sub


Private Sub pCargarDatosGrid()
    '--MOSTRAR DEMAS DATOS
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    '--nSQL=COTIZACION UNION  FATURA UNION BOLETA
    nSQL = " SELECT 12 AS tipodoc, alm_numseries.id AS serid,'Ticket'  as docdesc, alm_numseries.numser AS serdesc, var_plantilladoc.id AS impid, var_plantilladoc.descripcion AS impdesc " _
        + vbCr + " FROM (pvt_emp LEFT JOIN var_plantilladoc ON pvt_emp.idplan2 = var_plantilladoc.id) LEFT JOIN alm_numseries ON pvt_emp.idalmser2 = alm_numseries.id " _
        + vbCr + " WHERE pvt_emp.id = " + CStr(NulosN(txt(0).Text)) + " And (pvt_emp.idalmser2 <> 0 Or pvt_emp.idplan2 <> 0) " _
        + vbCr + " UNION " _
        + vbCr + "SELECT 1 AS tipodoc, alm_numseries.id AS serid,'Factura' as docdesc, alm_numseries.numser AS serdesc, var_plantilladoc.id AS impid, var_plantilladoc.descripcion AS impdesc " _
        + vbCr + " FROM (pvt_emp LEFT JOIN alm_numseries ON pvt_emp.idalmser = alm_numseries.id) LEFT JOIN var_plantilladoc ON pvt_emp.idplan = var_plantilladoc.id " _
        + vbCr + " WHERE pvt_emp.id = " + CStr(NulosN(txt(0).Text)) + " And (pvt_emp.idalmser <> 0 Or pvt_emp.idplan <> 0) " _
        + vbCr + " UNION " _
        + vbCr + " SELECT 3 AS tipodoc, alm_numseries.id AS serid,'Boleta de Venta' as docdesc, alm_numseries.numser AS serdesc, var_plantilladoc.id AS impid, var_plantilladoc.descripcion AS impdesc " _
        + vbCr + " FROM (pvt_emp LEFT JOIN alm_numseries ON pvt_emp.idalmser1 = alm_numseries.id) LEFT JOIN var_plantilladoc ON pvt_emp.idplan1 = var_plantilladoc.id " _
        + vbCr + " WHERE pvt_emp.id = " + CStr(NulosN(txt(0).Text)) + " And (pvt_emp.idalmser1 <> 0 Or pvt_emp.idplan1 <> 0) "
        
    
    RST_Busq RstTmp, nSQL, xCon
    RstTmp.Sort = " tipodoc desc"
    Fg1.Rows = 1
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = RstTmp.Fields("tipodoc") & ""
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = RstTmp.Fields("impid") & ""
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = RstTmp.Fields("docdesc") & ""
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = RstTmp.Fields("serdesc") & ""
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = RstTmp.Fields("impdesc") & ""
        RstTmp.MoveNext
    Loop
    If Fg1.Rows = 1 Then
        Fg1.Rows = 2
        Fg1.RowHeight(1) = 700
        UNIR_CELDAS Fg1, 1, 3, 1, 5, "Solicite al Administrador le configure" + vbCr + "un Nº. de Serie con su respectiva Plantilla de Impresión", flexAlignLeftCenter
    End If
    
    Set RstTmp = Nothing
End Sub

Private Sub pCargarArray()
    Dim mPos As Integer
    Erase ArrDocumento()
    With Fg1
        For mPos = 1 To .Rows - 1
            Select Case NulosN(.TextMatrix(mPos, 1))
                Case 12 '--TICKET
                    ArrDocumento(0, 0) = 0 '--IdDoc
                    ArrDocumento(0, 1) = .TextMatrix(mPos, 4) '--Nº Serie
                    ArrDocumento(0, 2) = NulosN(.TextMatrix(mPos, 2)) '--IdPlantilla
                    ArrDocumento(0, 3) = .TextMatrix(mPos, 5) '--Nombre Plantilla
                Case 1 '--FACTURA
                    ArrDocumento(1, 0) = 1
                    ArrDocumento(1, 1) = .TextMatrix(mPos, 4)
                    ArrDocumento(1, 2) = NulosN(.TextMatrix(mPos, 2))
                    ArrDocumento(1, 3) = .TextMatrix(mPos, 5)
                Case 3 '--BOLETA
                    ArrDocumento(2, 0) = 2
                    ArrDocumento(2, 1) = .TextMatrix(mPos, 4)
                    ArrDocumento(2, 2) = NulosN(.TextMatrix(mPos, 2))
                    ArrDocumento(2, 3) = .TextMatrix(mPos, 5)
                End Select
        Next mPos
    End With
End Sub

