VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmSelEmp2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresas Disponibles"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "FrmSelEmp2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAcepta 
      Caption         =   "&Aceptar"
      Height          =   350
      Left            =   1935
      TabIndex        =   1
      Top             =   1935
      Width           =   1300
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancelar"
      Height          =   350
      Left            =   3285
      TabIndex        =   0
      Top             =   1935
      Width           =   1300
   End
   Begin TrueOleDBGrid70.TDBGrid Dg1 
      Height          =   1845
      Left            =   15
      TabIndex        =   2
      Top             =   45
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   3254
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nº R.U.C."
      Columns(0).DataField=   "numruc"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Empresa"
      Columns(1).DataField=   "nomemp"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Año"
      Columns(2).DataField=   "anotra"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2275"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=6826"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6747"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1296"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1217"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1.5
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=-1,.fontsize=825"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80&"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Named:id=33:Normal"
      _StyleDefs(49)  =   ":id=33,.parent=0"
      _StyleDefs(50)  =   "Named:id=34:Heading"
      _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=34,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=37:Caption"
      _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(59)  =   "Named:id=38:HighlightRow"
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "FrmSelEmp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : FRMSELEMP2
'* Tipo              : FORMULARIO
'* Descripcion       : FORMULARIO PARA SELECCIONAR LA EMPRESA DE TRABAJO ACTUAL, ASI MISMO SIRVE PARA
'*                     CAMBIAR DE UNA EMPRESA A OTRA SIN NECESIDAD DE SALIR DEL PROGRAMA.ES AQUI DONDE
'*                     SE OBTIENE LA RUTA DE LA BASE DE DATOS DE TRABAJO
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 03/09/09
'* VERSION           : 1.0
'*****************************************************************************************************

Option Explicit
Dim RstEmp As New ADODB.Recordset               ' RECORSET PARA MOSTRAR LAS EMPRESAS REGISTRADAS EN EL SISTEM
Dim fOrdenLista As Boolean                      ' especfica el orden de la lista de la consulta

Private Sub CmdAcepta_Click()
    Dim xCad As String
    Dim xRutaData As String
    
    ' CARGAMOS EL ID DE LA EMPRESA
    xIdEmpresa = 0
    
    xIdEmpresa = RstEmp("id")
    
    'CARGAMOS EL NOMBRE DEL SISTEM  Y EL NOMBRE DEL LA EMPRESA EN EL CAPTION DEL FORMULARIO PRINCIPAL
    MDIPrincipal.Caption = Trim(AP_NOMSIS) + "                                              " + Trim(RstEmp("nomemp"))
    
    AP_AÑODAT = NulosC(RstEmp("anotra"))                    ' CARGAMOS EL AÑO DE TRABAJO DE LA EMPRESA SELECCIONADA
    AP_RUTDATTRA = NulosC(RstEmp("ruta"))                   ' CARGAMOS LA RUTA DE TRABAJO DE LA EMPRESA SELECCIONADA
    
    xRutaData = Trim(AP_RUTABD) + NulosC(RstEmp("ruta"))    'DEFINIMOS LA RUTA OBSOLUTA DE LA BASE DE DATOS DE LA EMPRESA SELECCIONADA
    
    'CARGAMOS EL NOMBRE DE LA EMPRES Y EL AÑO DE TRABAJO EN LA BARRA DE ESTADO DEL FORMULARIO PRINCIPAL
    MDIPrincipal.StatusBar1.Panels(1).Text = "Empresa  : " + NulosC(RstEmp("abrevia"))
    MDIPrincipal.StatusBar1.Panels(3).Text = "Año : " + Trim(AP_AÑODAT)
    
    MDIPrincipal.Toolbar1.Buttons(3).Enabled = True
    
    ' ABRIMOS LA CONEXION A LA BASE DE DATOS DE LA EMPRESA
    Dim xFun As New eps_librerias.FuncionesData
    
    xFun.F_BASEDATOS = xRutaData                            ' PASAMOS LA RUTA DE ALA BASE DE DATOS
    xFun.F_GRUPOTRABAJO = AP_RUTASY + "seven.mdw"           ' PASAMOS LA RUTA DEL ARCHIVO DE TRABAJO DE LA BASE DE DATOS
    xFun.F_PASSWORD = Eps_Pass                              ' PASAMOS EL PASSWORD DE LA BASE DE DATOS
    xFun.F_USUARIO = Eps_User                               ' PASAMOS EL USUARIO DE LA BASE DE DATOS
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"            ' PASAMOS EL NOMBRE DEL PROVEEDOR DE DATOS PARA ADO 2.5
    
    xDataSource = xRutaData
    Set xCon = xFun.AbrirConeccion                          ' ABRIMOS LA CONECCION
    Set xFun = Nothing

    CargaDatosEmpresa                                       ' CARGAMOS LOS DATOS DE LA EMPRESA
    Set RstEmp = Nothing
    Unload Me
    
'    If MDIPrincipal.menu11.Enabled = False Then
''        ActivarMenus
'    End If

    ' PREGUNTAMOS SI SE NECESITA SOLICITAR NUEVAMENTE EL USUARIO Y PASSWORD PARA ACCEDER A LA EMPRESA SELECCIONADA
    If PedirUsuario = True Then
        ' SI QUE REQUIERE PEDIR NUEVAMENTE EL USUARIO, LLAMAMOS AL FORMULARIO FRMINGRESOUSUARIO
        FrmIngresoUsuario.Show vbModal
        
    Else
        ' LE DAMOS LA BIENVENIDA A LA NUEVA EMPRESA
        MsgBox "Ha seleccionado la " + Trim(MDIPrincipal.StatusBar1.Panels(1).Text), vbInformation + vbOKOnly + vbDefaultButton1
        
    End If
    
    SetearMenus xIdUsuario
    
End Sub

Private Sub CmdCancel_Click()
    xIdEmpresa = 0
    Set RstEmp = Nothing
    MsgBox "No ha seleccionado ninguna empresa", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set xCon = Nothing
    Unload Me
    End
End Sub

Private Sub Dg1_DblClick()
    CmdAcepta_Click
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstEmp.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdAcepta_Click
    End If
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        CmdAcepta_Click
    End If
End Sub

Private Sub Form_Activate()
    ' CARGAMOS TODAS LAS EMPRESAS ACTIVAS AL RECORDSET PRINCIPAL
    RST_Busq RstEmp, "SELECT mae_empresa.* From mae_empresa WHERE activo = -1  ORDER BY mae_empresa.anotra DESC , mae_empresa.nomemp ", xCon

    ' MOSTRAMOS LAS EMPRESAS ACTIVAS EL DATA GRID DEL FORMULARIO
    Set Dg1.DataSource = RstEmp
    Dg1.SetFocus
End Sub

