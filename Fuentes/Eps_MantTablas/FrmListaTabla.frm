VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.CommandBars.v12.0.0.ocx"
Begin VB.Form FrmListaTabla 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin SizerOneLibCtl.ElasticOne EO 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   555
      Width           =   10425
      _cx             =   18389
      _cy             =   12726
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   4
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmListaTabla.frx":0000
      Begin VB.Frame Frame3 
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   480
         Left            =   90
         TabIndex        =   7
         Top             =   90
         Width           =   10245
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   450
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   6675
         Width           =   10245
         _cx             =   18071
         _cy             =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   1
         GridCols        =   3
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmListaTabla.frx":005A
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   270
            Left            =   5205
            TabIndex        =   5
            Top             =   90
            Width           =   2445
            Begin VB.Label LblNumReg 
               Alignment       =   2  'Center
               Caption         =   "LblNumReg"
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
               Height          =   225
               Left            =   1335
               TabIndex        =   8
               Top             =   30
               Width           =   1050
            End
            Begin VB.Label Label3 
               Caption         =   "Nº Registros :"
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
               Height          =   210
               Left            =   45
               TabIndex        =   6
               Top             =   30
               Width           =   1215
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   270
            Left            =   7710
            TabIndex        =   4
            Top             =   90
            Width           =   2445
         End
      End
      Begin TrueOleDBGrid70.TDBGrid Dg1 
         Height          =   5640
         Left            =   90
         TabIndex        =   2
         Top             =   975
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   9948
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Hola"
         Columns(0).DataField=   "Hola"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
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
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Named:id=33:Normal"
         _StyleDefs(41)  =   ":id=33,.parent=0"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(44)  =   ":id=34,.wraptext=-1"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(49)  =   "Named:id=37:Caption"
         _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(51)  =   "Named:id=38:HighlightRow"
         _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=39:EvenRow"
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "LblTitulo"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   630
         Width           =   10245
      End
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   1275
      Top             =   120
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmListaTabla.frx":00A8
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   0
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "FrmListaTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TopEO As Integer
Dim TAMAÑO_TOOL As TOOL_TAMAÑO_ICO
Dim RstLis As New ADODB.Recordset
Dim SeEjecuto As Boolean
Const BTN_NEW = 1  ' NUEVO
Const BTN_MOD = 2  ' MOFICAR
Const BTN_BUS = 3  ' BUSCAR
Const BTN_EXC = 4  ' EXPORTAR EXCEL
Const BTN_IMP = 5  ' IMPRIMIR
Const BTN_CAL = 6  ' CALENDARIO
Const BTN_CON = 7  ' CONFIGURAR
Const BTN_SAL = 8  ' SALIR

Public xSqlCad As String
Public xRstCampos As New ADODB.Recordset


Sub CrearTool()
    'CREAMOS EL TOOLBAR
    Dim Opciones(7, 3) As String
    
    Opciones(0, 0) = Str(BTN_NEW):    Opciones(0, 1) = "Nuevo Registro":              Opciones(0, 2) = "0":      Opciones(0, 3) = "Nuevo Registro"
    Opciones(1, 0) = Str(BTN_MOD):    Opciones(1, 1) = "Modificar Registro":          Opciones(1, 2) = "0":      Opciones(1, 3) = "Modificar Registro"
    Opciones(2, 0) = Str(BTN_BUS):    Opciones(2, 1) = "Buscar Registro":             Opciones(2, 2) = "1":      Opciones(2, 3) = "Buscar Registro"
    Opciones(3, 0) = Str(BTN_EXC):    Opciones(3, 1) = "Exportar Excel":              Opciones(3, 2) = "0":      Opciones(3, 3) = "Exportar Excel"
    Opciones(4, 0) = Str(BTN_IMP):    Opciones(4, 1) = "Imprimir":                    Opciones(4, 2) = "0":      Opciones(4, 3) = "Imprimir"
    Opciones(5, 0) = Str(BTN_CAL):    Opciones(5, 1) = "Calendario":                  Opciones(5, 2) = "0":      Opciones(5, 3) = "Calendario"
    Opciones(6, 0) = Str(BTN_CON):    Opciones(6, 1) = "Configurar Formulario":       Opciones(6, 2) = "1":      Opciones(6, 3) = "Configurar Formulario"
    Opciones(7, 0) = Str(BTN_SAL):    Opciones(7, 1) = "Salir":                       Opciones(7, 2) = "1":      Opciones(7, 3) = "Salir"
        
    Dim xFun As New eps_librerias.Codejock
    'PocisionarContenedor
    xFun.BORRARMENU = True
    TAMAÑO_TOOL = I24x24
    xFun.CrearToolBar Opciones, CommandBars1, ImageManager1, TAMAÑO_TOOL
    Set xFun = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        CrearCampos
        SeEjecuto = True
        RST_Busq RstLis, xSqlCad, xCon
        
        Set Dg1.DataSource = RstLis
        If RstLis.RecordCount = 0 Then
            LblNumReg.Caption = "0"
        Else
            LblNumReg.Caption = Str(RstLis.RecordCount)
        End If
    End If
End Sub

Sub CrearCampos()
    Dim C As TrueOleDBGrid70.Column
    ' INICIALIZAMOS CADA COLUMNA
    Dg1.Columns.Remove 0                                ' ELIMINAMOS LA COLUMNA INICIAL

    Dim A, xIndex As Integer
    xIndex = 0
    xRstCampos.MoveFirst
    For A = 1 To xRstCampos.RecordCount
        Set C = Dg1.Columns.Add(xIndex)                 ' AGREGAMOS UNA COLUMNA
        With C
            .Visible = True                             ' ESPECIFICA QUE LA COLUMNA SE PUEDE VER
            .HeadAlignment = dbgCenter                  ' ALINEACION DE LA CABECERA
            .DataField = xRstCampos("nomcampo")         ' NOMBRE DEL CAMPO PARA LA COLUMNA
            .Caption = xRstCampos("caption")            ' CAPTION PARA LA CABECECERA DE LA COLUMNA
            If xRstCampos("alineacion") = 1 Then .Alignment = dbgRight           ' ALINEACION DE TODA LA COLUMNA MENOS DE LA CABECERA
            If xRstCampos("alineacion") = 2 Then .Alignment = dbgLeft            ' ALINEACION DE TODA LA COLUMNA MENOS DE LA CABECERA
            If xRstCampos("alineacion") = 3 Then .Alignment = dbgCenter          ' ALINEACION DE TODA LA COLUMNA MENOS DE LA CABECERA
            .Width = xRstCampos("ancho")                ' ANCHO DE LA COLUMNA
            
        End With
    
        ' SUPRIMIMOS LAS LINEAS DE DIVICION DE LAS COLUMNAS
        Dg1.Splits(0).Columns(xIndex).DividerStyle = dbgNoDividers
        
        ' PONEMOS EL CHECK A LOS CAMPOS LOGICOS SIEMPRE Y CUANDO LA PRESENTACION SEA 2
        If xRstCampos("presentacion") = 2 Then Dg1.Splits(0).Columns(xIndex).ValueItems.Presentation = dbgCheckBox
        
        xRstCampos.MoveNext
        If xRstCampos.EOF = True Then Exit For
        xIndex = xIndex + 1
    Next A
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    LblTitulo.BackColor = &H8000000F
    CrearTool
    
    ' ESPECIFICAMOS LA POSICION INICIAL DEL CONTROL SIZER ONE EN FUNCION A LA BARRA DE HERRAMIENTAS
    If TAMAÑO_TOOL = I16x16 Then EO.Top = 400: TopEO = 400
    If TAMAÑO_TOOL = I24x24 Then EO.Top = 520: TopEO = 520
    If TAMAÑO_TOOL = I32x32 Then EO.Top = 640: TopEO = 640
    If TAMAÑO_TOOL = I48x48 Then EO.Top = 890: TopEO = 890
    
    Dg1.RecordSelectors = False                      ' OCULTAMOS LA COLUMNA DE SELECCION DE REGISTROS
    Dg1.Appearance = dbgTrack3D                      ' ESPECIFICAMOS LA APARIENCIA DEL CONTROL GRID
    Dg1.HeadLines = 1.5                              ' ESTABLE EL ALTO DE LAS CABECERAS 1.5
    Dg1.HeadFont.Bold = True                         ' PONE A NEGRITA LAS CABECERAS DE LAS COLUMNAS
    Dg1.HeadForeColor = &H800000                     ' COLOR PARA LAS CABECERAS DE LAS COLUMNAS
    Dg1.Splits(0).AlternatingRowStyle = True         ' ALTERNAMOS LOS COLORES DE LA CUADRICULA
    Dg1.Splits(0).EvenRowStyle.BackColor = &HDAFEFB  ' DEFINIMOS EL COLOR PARA ALTERNAR
    Dg1.Splits(0).MarqueeStyle = dbgHighlightRow     ' ESPECIFICAMOS QUE SE SELECCIONARAN TODAS LAS COLUMNAS
    Dg1.Splits(0).SelectedBackColor = &H80&          ' ESTABLECEMOS QUE EL COLOR DE SELECCION SEA ROJO
    Dg1.RowDividerStyle = dbgNoDividers              ' ELIMINAMOS LA LINEAS DE DIVICION HORIZONTALES
    Dg1.RowHeight = 250                              ' ESPECIFIVA EL ALTO DE CADA CELDA
    Dg1.SelectedBackColor = &H80&                    ' ESPECIFICA EL COLOR DE SELECCION DE FILA
    
    Dg1.Splits(0).MarqueeStyle = dbgHighlightRow      ' SELECCIONAMOS TODA LA FILA
    Dg1.Splits(0).HighlightRowStyle.BackColor = &H80& ' PONE LA SELECCION EN COLOR GRANATE
End Sub

Private Sub Form_Resize()
    ' RECONFIGURAMOS EL TAMAÑOS DE LOS CONTROLES CUANDO SE MODIFIQUE EL TAMAÑO DEL FORMULARIO
    EO.Width = Me.Width - 130
    
    If Me.Height <= (TopEO + 2375) Then
        Me.Height = (TopEO + 2375)
    Else
        EO.Height = (Me.Height - (TopEO + 400))
    End If
    Me.Visible = True
    Me.Refresh
End Sub
