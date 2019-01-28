VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCierreMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEVEN - Cierre de Mes"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5385
   Begin VB.CommandButton CmdBusTipDoc 
      Height          =   240
      Left            =   1545
      Picture         =   "FrmCierreMes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   390
      Width           =   240
   End
   Begin VB.TextBox TxtTipitem 
      Height          =   300
      Left            =   900
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "Txttipitem"
      Top             =   360
      Width           =   915
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7755
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":0132
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":0676
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":0A08
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":0B8C
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":0FE0
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":10F8
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":163C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":1B80
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":1C94
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":1DA8
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":21FC
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":2368
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCierreMes.frx":28B0
            Key             =   "IMG12"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Gasto"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Restaurar Gasto"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Saldo del Documento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anular Gasto"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Gasto"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Emitir Gasto Anulado"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir Gasto"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   4335
      Left            =   60
      TabIndex        =   2
      Top             =   690
      Width           =   5295
      _cx             =   9340
      _cy             =   7646
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
      ForeColorSel    =   -2147483634
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
      Rows            =   20
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmCierreMes.frx":2BCA
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
   Begin VB.Label Lbldesitem 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lbldesitem"
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
      Height          =   300
      Left            =   1860
      TabIndex        =   4
      Top             =   360
      Width           =   3435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Formulario"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   450
      Width           =   720
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar Item            "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Item                "
      End
      Begin VB.Menu menu1_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_5 
         Caption         =   "Ver Historico de Precios"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu menu2_1 
         Caption         =   "Agregar Documento"
      End
      Begin VB.Menu menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu2_3 
         Caption         =   "Eliminar Documento"
      End
   End
End
Attribute VB_Name = "FrmCierreMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : FRMMANOPCIONESUSUARIO
'* Tipo              : FORMULARIO
'* Descripcion       : CONTROLA EL INGRESO DE LOS USUARIOS AL SISTEMA, VALIDA EL USUARIO Y EL PASSWORD
'*                     , CARGANDO LAS OPCIONES DE ACCESO AL MENU DEL SISTEMA
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 04/09/09
'* VERSION           : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer          ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim SeEjecuto As Boolean        ' ESPECIFICA SE EJECUTO EL EVENTO ACTIVATE, SOLO ES USUADO EN ESTE EVENTO
Dim xHorIni As Date
Dim mIdRegistro&                ' IDENTIFICADO DEL REGISTRO ACTUAL
Dim mMesActivo As Integer       ' ESPECIFICA EL MES ACTIVO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO

'*****************************************************************************************************
'* Nombre Modulo  : Cancelar()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : CANCELA EL PROCESO DE INGRESO O MODIFICACION DE UN REGISTRO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Cancelar()
    Dim x As Integer
    ActivaTool
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : ActivaTool()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : ACTIVA DESACTIVA LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    Toolbar1.Buttons(11).Enabled = Not Toolbar1.Buttons(11).Enabled
    
    Toolbar1.Buttons(13).Enabled = Not Toolbar1.Buttons(13).Enabled
    Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub

'Sub Nuevo()
'    QueHace = 1
'    Blanquea
'
'    ActivaTool
'
'    Fg1.ColComboList(1) = "0 Seleccion|1 Manual"
'    Fg1.Editable = flexEDKbdMouse
'    Fg1.SelectionMode = flexSelectionFree
'
'    Fg1.Rows = 1
'    pGridConfigurar
'    xHorIni = Time
'End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Blanquea()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : INICIALIZA A VACIO LOS CONTROLES DEL FORMULARIO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Blanquea()
    Me.TxtTipitem = ""
    Me.Lbldesitem = ""
End Sub

Private Sub CmdBusTipDoc_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    Dim x As Integer
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    xCampos(2, 0) = "Modulo":         xCampos(2, 1) = "modulo":           xCampos(2, 2) = "3000":         xCampos(2, 3) = "C"
    
    'xform.SQLCad = " SELECT descripcion, id FROM mae_formularios WHERE  Proceso = -1 ORDER BY descripcion"
    
    ' 0 = Ninguno; 1 = Mantenimientos; 2 =Proceso; 3 = Consultas; 4 = Analisis
    xform.SQLCad = "SELECT mae_menu.id, mae_menu.descripcion, mae_modulo.descripcion AS modulo " _
        + vbCr + "FROM mae_menu " _
        + vbCr + "LEFT JOIN mae_modulo ON mae_modulo.idmodulo = mae_menu.idmodulo " _
        + vbCr + "WHERE (((mae_menu.categoria)=2)) " _
        + vbCr + "ORDER BY mae_menu.descripcion"
    
    xform.Titulo = "Buscando Formularios de proceso"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            '--almacenar temporalmente la hora si se pretende grabar los cambios
            xHorIni = Time
        
            TxtTipitem = NulosN(xRs!Id)
            Lbldesitem = NulosC(xRs!Descripcion)
        
            ' BUSCAMOS QUE EL FORMULARIO SELECCIONADO EXISTA EN LA TABLA VAR_CIERRE
'            RST_Busq xRs2, " SELECT mae_formularios.id, mae_formularios.descripcion, var_cierre.idmes, con_meses.descripcion as Mes, IIf(var_cierre.estado=0,'Bloqueado','Abierto') AS Estado " & _
'                " FROM con_meses RIGHT JOIN (mae_formularios INNER JOIN var_cierre ON mae_formularios.id = var_cierre.idform) ON con_meses.id = var_cierre.idmes " & _
'                " Where mae_formularios.id = " & NulosN(xRs!Id) & " ORDER BY var_cierre.idmes ", xCon
            
            
            RST_Busq xRs2, "SELECT mae_menu.id, mae_menu.descripcion, var_cierre.idmes, con_meses.descripcion AS Mes, IIf(var_cierre.estado=0,'Bloqueado','Abierto') AS Estado " _
                & " FROM (con_meses RIGHT JOIN var_cierre ON con_meses.id = var_cierre.idmes) INNER JOIN mae_menu ON var_cierre.idform = mae_menu.id " _
                & " Where (((mae_menu.Id) = " & NulosN(xRs!Id) & ")) " _
                & " ORDER BY var_cierre.idmes ", xCon

            With Fg1
                ' si existe el la tabla var_cierre
                If xRs2.RecordCount <> 0 Then
                    .Rows = 1
                    Do While Not xRs2.EOF
                        .AddItem ""
                        .Row = .Rows - 1
                        .TextMatrix(.Row, 1) = NulosC(xRs2!Mes)
                        .TextMatrix(.Row, 2) = IIf(xRs2!estado = "Bloqueado", -1, 0)
                        
                        .TextMatrix(.Row, 3) = NulosN(xRs2!IdMes)
                        .TextMatrix(.Row, 4) = NulosN(xRs2!Id)
                        xRs2.MoveNext
                    Loop
                Else
                    ' si se va cerrar o bloquear por primera vez
                    Set xRs2 = Nothing
                    RST_Busq xRs2, " SELECT id, descripcion  FROM con_meses order by id", xCon
                
                    If xRs2.RecordCount <> 0 Then
                        .Rows = 1
                    
                        Do While Not xRs2.EOF
                            .AddItem ""
                            .Row = .Rows - 1
                            .TextMatrix(.Row, 1) = NulosC(xRs2!Descripcion)
                            .TextMatrix(.Row, 2) = 0
                            .TextMatrix(.Row, 3) = NulosN(xRs2!Id)
                            .TextMatrix(.Row, 4) = NulosN(TxtTipitem)
                            xRs2.MoveNext
                        Loop
            
                    End If
                    End If
            End With
        End If
    End If
    
    Me.Fg1.SetFocus
    Set xform = Nothing
    Set xRs = Nothing
    Set xRs2 = Nothing
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case 1, 3
            KeyAscii = 0
    End Select
End Sub

Private Sub Form_Activate()
If SeEjecuto = False Then
    SeEjecuto = True
        '--Almacenar temporalmente el codigo del menu
    IdMenuActivo = 120

    '--bloquear accesos
    OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
    
    '--ocultar botones
    Toolbar1.Buttons(1).Visible = False '--agregar
    Toolbar1.Buttons(2).Visible = False '--modificar
    Toolbar1.Buttons(3).Visible = False '--eliminar
    Toolbar1.Buttons(6).Visible = False '--cancelar
End If
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Fg1.SelectionMode = flexSelectionByRow
    
    Blanquea
    Fg1.Rows = 1
    Fg1.Col = 2
    Fg1.Editable = flexEDKbd
    
    pGridConfigurar
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
        End If
    End If
    
    If Button.Index = 15 Then
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Grabar()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : GRABA LAS OPCIONES DE CIEERE
'* Paranetros     : NULL
'* Retorna        : LOGICO (VERDADERO = SE CULMINO EL PROCESO; FALSO = NO SE HIZO EL PROCESO
'*****************************************************************************************************
Function Grabar() As Boolean
    If NulosN(TxtTipitem.Text) = 0 Then
        MsgBox "Falta seleecionar el Formulario", vbInformation, xTitulo
        TxtTipitem.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para los bloqueos y cierres respectivos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
On Error GoTo LaCague
    Dim RstCab As New ADODB.Recordset
    Dim Rst As New ADODB.Recordset
    Dim xId As Double
    Dim x As Integer
    
    xCon.Execute "DELETE * FROM var_cierre WHERE idform =" & NulosN(TxtTipitem.Text) & ""
        
    xId = HallaCodigoTabla("var_cierre", xCon, "id")
  
    RST_Busq RstCab, "SELECT TOP 1 * FROM var_cierre ", xCon
    
    With Me.Fg1
        For x = 1 To Fg1.Rows - 1
            RstCab.AddNew
            RstCab!Id = xId
            RstCab!IdMes = NulosN(.TextMatrix(x, 3))
            RstCab!idForm = NulosN(TxtTipitem.Text)
                    
            If NulosN(.TextMatrix(x, 2)) = -1 Then
                RstCab!estado = 0
            Else
                RstCab!estado = -1
            End If
            RstCab.Update
            xId = xId + 1
        Next
    
    End With
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, 2, xHorIni, Time, Date, xCon, NulosN(TxtTipitem.Text)
    
    MsgBox "La operacion se grabó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Exit Function
    
LaCague:
    Set RstCab = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Exit Function
End Function

'*****************************************************************************************************
'* Nombre Modulo  : pGridConfigurar()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : CONFIGURA EL GRID DEL FORMULARIO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Private Sub pGridConfigurar()
    Fg1.ColWidth(1) = 1500
    Fg1.ColWidth(2) = 1200
    Fg1.ColWidth(3) = 0 'idmes
    Fg1.ColWidth(4) = 0 'idform
End Sub

Private Sub TxtTipItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipitem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And NulosN(TxtTipitem.Text) <> 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(TxtTipitem.Text), xCon
    End If
End Sub

Private Sub TxtTipitem_Validate(Cancel As Boolean)
    Dim xRs2 As New ADODB.Recordset
    
    If NulosN(TxtTipitem.Text) = 0 Then
        Exit Sub
    End If
    
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    
    RST_Busq xRs2, "SELECT mae_menu.id, mae_menu.descripcion FROM mae_menu WHERE mae_menu.id = " & NulosN(TxtTipitem.Text) & "  and (((mae_menu.categoria)=2)) ORDER BY mae_menu.descripcion", xCon
    If xRs2.RecordCount = 0 Then
        MsgBox "Código de menú incorrecto", vbInformation, xTitulo
        TxtTipitem.Text = 0
        Lbldesitem.Caption = ""
        Set xRs2 = Nothing
        Exit Sub
    End If
    
    Set xRs2 = Nothing
    
    ' BUSCAMOS QUE EL FORMULARIO SELECCIONADO EXISTA EN LA TABLA VAR_CIERRE
'    RST_Busq xRs2, " SELECT mae_formularios.id, mae_formularios.descripcion, var_cierre.idmes, con_meses.descripcion as Mes, IIf(var_cierre.estado=0,'Bloqueado','Abierto') AS Estado " & _
'        " FROM con_meses RIGHT JOIN (mae_formularios INNER JOIN var_cierre ON mae_formularios.id = var_cierre.idform) ON con_meses.id = var_cierre.idmes " & _
'        " Where mae_formularios.id = " & NulosN(Me.TxtTipitem) & " ORDER BY var_cierre.idmes ", xCon
            
    RST_Busq xRs2, "SELECT mae_menu.id, mae_menu.descripcion, var_cierre.idmes, con_meses.descripcion AS Mes, IIf(var_cierre.estado=0,'Bloqueado','Abierto') AS Estado " _
        & " FROM (con_meses RIGHT JOIN var_cierre ON con_meses.id = var_cierre.idmes) INNER JOIN mae_menu ON var_cierre.idform = mae_menu.id " _
        & " Where (((mae_menu.Id) = " & NulosN(Me.TxtTipitem) & ")) " _
        & " ORDER BY var_cierre.idmes ", xCon
    
    '--almacenar temporalmente la hora si se pretende grabar los cambios
    xHorIni = Time
    
    With Fg1
        ' si existe el la tabla var_cierre
        If xRs2.RecordCount <> 0 Then
            TxtTipitem = NulosN(xRs2!Id)
            Lbldesitem = NulosC(xRs2!Descripcion)
        
            .Rows = 1
            Do While Not xRs2.EOF
                .AddItem ""
                .Row = .Rows - 1
                .TextMatrix(.Row, 1) = NulosC(xRs2!Mes)
                .TextMatrix(.Row, 2) = IIf(xRs2!estado = "Bloqueado", -1, 0)
                .TextMatrix(.Row, 3) = NulosN(xRs2!IdMes)
                .TextMatrix(.Row, 4) = NulosN(xRs2!Id)
                xRs2.MoveNext
            Loop
        Else
            ' si se va cerrar o bloquear por primera vez
            Set xRs2 = Nothing
            RST_Busq xRs2, " SELECT id, descripcion  FROM con_meses order by id", xCon
        
            If xRs2.RecordCount <> 0 Then
                .Rows = 1
                Do While Not xRs2.EOF
                    .AddItem ""
                    .Row = .Rows - 1
                    .TextMatrix(.Row, 1) = NulosC(xRs2!Descripcion)
                    .TextMatrix(.Row, 2) = 0
                    
                    .TextMatrix(.Row, 3) = NulosN(xRs2!Id)
                    .TextMatrix(.Row, 4) = NulosN(TxtTipitem)
                    xRs2.MoveNext
                Loop
            End If
        End If
    End With
        
    Me.Fg1.SetFocus
    Set xRs2 = Nothing
End Sub
