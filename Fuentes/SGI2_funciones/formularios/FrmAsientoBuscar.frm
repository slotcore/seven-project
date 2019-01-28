VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmAsientoBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Asiento"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7Ctl.VSFlexGrid fg1 
      Height          =   4770
      Left            =   1800
      TabIndex        =   11
      Top             =   990
      Width           =   7530
      _cx             =   13282
      _cy             =   8414
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColor       =   14745342
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14745342
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
      Rows            =   5
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmAsientoBuscar.frx":0000
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
   Begin VB.Frame Frame3 
      Caption         =   "Ingresar Asiento"
      Height          =   600
      Left            =   5340
      TabIndex        =   5
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton Command2 
         Caption         =   "&Ver Asiento"
         Height          =   300
         Left            =   1380
         TabIndex        =   12
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox TxtAsiento 
         Height          =   300
         Left            =   120
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "TxtAsiento"
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Libro"
      Height          =   600
      Left            =   1740
      TabIndex        =   0
      Top             =   360
      Width           =   3555
      Begin VB.CommandButton CmdBusProv 
         Height          =   230
         Left            =   3240
         Picture         =   "FrmAsientoBuscar.frx":00D5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   210
      End
      Begin VB.TextBox TxtLibro 
         Height          =   300
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "TxtLibro"
         Top             =   240
         Width           =   3435
      End
      Begin VB.Label LblIdLibro 
         AutoSize        =   -1  'True
         Caption         =   "LblIdLibro"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2070
         TabIndex        =   3
         Top             =   90
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin TrueOleDBGrid70.TDBGrid Dg1 
      Height          =   4770
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Doble Clik para Ver el Asiento"
      Top             =   990
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   8414
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nº Reg."
      Columns(0).DataField=   "registro"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   265
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColMove=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2275"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   0
      ColumnFooters   =   -1  'True
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
      _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
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
   Begin VB.Frame Frame4 
      Caption         =   "Busca Periodo"
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   1755
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Left            =   1365
         Picture         =   "FrmAsientoBuscar.frx":0207
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   270
         Width           =   285
      End
      Begin VB.Label LblIdMes 
         Caption         =   "LblIdMes"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   690
         Width           =   1365
      End
      Begin VB.Label lbl_periodo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_periodo "
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
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Width           =   1620
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   556
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6855
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":0589
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":0ACD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":0E5F
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":0FE3
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":1437
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":154F
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":1A93
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":1FD7
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":20EB
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":21FF
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":2653
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":27BF
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":2D07
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":3021
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAsientoBuscar.frx":33B3
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmAsientoBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim SeEjecuto As Boolean

Dim SGI_JC As New SGI2_funciones.JC_Varios
Dim SGI_JC1 As New SGI2_funciones.JC_VSFlexGrid

Dim RstFrm As New ADODB.Recordset
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim NumRegistro As String

Public Sub pRecibeVerAsiento()
    '---------------------------
    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    
    Configurar_Grilla

    DoEvents
    If NumRegistro = "" Then Exit Sub
       
    DoEvents
    
    
'    nSQL = "SELECT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, mae_libros.descripcion AS libro, con_diario.fchdoc AS fchope, con_diario.rregistro AS registroref, IIf(con_diario.ridtipper=0,tes_documentos.abrev,mae_documento.abrev) AS tdocdesc, con_diario.rfchope AS fchdoc, con_diario.rnumerodoc AS numdoc, IIf([con_diario].[ridtipper]=1,[mae_prov].[nombre],IIf([con_diario].[ridtipper]=2,[mae_cliente].[nombre],IIf([con_diario].[ridtipper]=3,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ', ' & [pla_empleados].[nom],IIf([con_diario].[ridtipper]=5,[mae_bancos].[descripcion],'')))) AS apenom, con_tc.impven AS tc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebesol, " _
'        + vbCr + " IIf(con_diario.idmon=2,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabersol, " _
'        + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(con_tc.impven Is Null Or con_diario.impdebsol=0,0,(con_diario.impdebsol/con_tc.impven))) AS impdebedol, " _
'        + vbCr + " IIf(con_diario.idmon = 2, con_diario.imphabdol, IIf(con_tc.impven Is Null Or con_diario.imphabsol = 0, 0, (con_diario.imphabsol / con_tc.impven))) As imphaberdol " _
'        + vbCr + " FROM ((pla_empleados RIGHT JOIN (mae_cliente RIGHT JOIN (((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper) ON pla_empleados.id = con_diario.ridper) LEFT JOIN tes_documentos ON con_diario.rtipdoc = tes_documentos.id) LEFT JOIN mae_bancos ON con_diario.ridper = mae_bancos.id " _
'        + vbCr + " WHERE (((Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi])='" & NumRegistro & "')) " _
'        + vbCr + " ORDER BY con_planctas.cuenta;"
    
    
    nSQL = "SELECT Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi] AS registro, mae_libros.descripcion AS libro, iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tc, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, " _
        + vbCr + " iif(con_diario.ajuste=2,0, IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0,0,con_diario.impdebdol*(iif(con_diario.idlib in (3,6,44), con_diario.tc,con_tc.impven))),con_diario.impdebsol) ) AS impdebesol, " _
        + vbCr + " iif(con_diario.ajuste=2,0, IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0,0,con_diario.imphabdol*(iif(con_diario.idlib in (3,6,44), con_diario.tc,con_tc.impven))),con_diario.imphabsol) ) AS imphabersol, " _
        + vbCr + " iif(con_diario.ajuste=1,0, IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/(iif(con_diario.idlib in (3,6,44), con_diario.tc,con_tc.impven))))) ) AS impdebedol, " _
        + vbCr + " iif(con_diario.ajuste=1,0, IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/(iif(con_diario.idlib in (3,6,44), con_diario.tc,con_tc.impven))))) ) As imphaberdol " _
        + vbCr + " FROM ((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE (((Format([con_diario].[idmes],'00') & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & [con_diario].[numasi])='" & NumRegistro & "')) " _
        + vbCr + " ORDER BY con_planctas.cuenta; "

Me.MousePointer = vbHourglass
RST_Busq Rst, nSQL, xCon


If Rst.RecordCount <> 0 Then



Do While Not Rst.EOF
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("ctanum"))
    Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("ctadesc"))
    Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(Rst("tc"))
    Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosN(Rst("impdebesol")), SGI_JC1.FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(Rst("imphabersol")), SGI_JC1.FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(Rst("impdebedol")), SGI_JC1.FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(Rst("imphaberdol")), SGI_JC1.FORMAT_MONTO)
    
'    Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(rst("registroref"))
'    Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(rst("tdocdesc"))
'    Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(rst("numdoc"))
'    Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosC(rst("fchdoc"))
'    Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosC(rst("apenom"))
    
    Rst.MoveNext
Loop



    Fg1.Rows = Fg1.Rows + 1

'''    Fg1.TextMatrix(Fg1.Rows - 1, 4) = SGI_JC1.GRID_SUMAR_COL(Fg1, 4)
'''    Fg1.TextMatrix(Fg1.Rows - 1, 5) = SGI_JC1.GRID_SUMAR_COL(Fg1, 5)
'''    Fg1.TextMatrix(Fg1.Rows - 1, 6) = SGI_JC1.GRID_SUMAR_COL(Fg1, 6)
'''    Fg1.TextMatrix(Fg1.Rows - 1, 7) = SGI_JC1.GRID_SUMAR_COL(Fg1, 7)

    SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, &H800000, True, , "TOTAL =>"
    SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 4)
    SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 5)
    SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 6)
    SGI_JC1.FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H800000, True, , SGI_JC1.GRID_SUMAR_COL(Fg1, 7)

End If
   
    Me.MousePointer = vbDefault
    
    DoEvents
End Sub

Public Sub pRecibeLinkTmp(xCon As ADODB.Connection, xRst As ADODB.Recordset, idlib As Integer, idmov As Double)
    '---------------------------
    Set RstFrm = xRst
    
    DoEvents
End Sub

Public Sub fDefinirRst(xRst As ADODB.Recordset)
    Set xRst = Nothing
    Set xRst = New ADODB.Recordset
    
    xRst.Fields.Append "IdCue", adNumeric '--codigo de cuenta
    xRst.Fields.Append "Importe", adDouble '--importe de la cuenta
    xRst.Fields.Append "tipo", adVarChar, 2 '--
    
    xRst.Open

End Sub

Private Sub pCargarDatos()
'--cargar los datos de la grilla


End Sub


Private Sub cmd_periodo_Click()
    Dim xfrm As New SGI2_funciones.Varias
    LblIdMes.Caption = xfrm.SeleccionaMes(xCon)
    Set xfrm = Nothing
    lbl_periodo.Caption = Busca_Codigo(NulosN(LblIdMes.Caption), "id", "descripcion", "con_meses", "N", xCon)
    pCargarAsientos
End Sub




Private Sub Command2_Click()
    If NulosC(TxtAsiento.Text) = "" Then Exit Sub
    NumRegistro = NulosC(TxtAsiento.Text)
    pRecibeVerAsiento
End Sub

Private Sub Dg1_DblClick()
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.EOF = True Or RstFrm.RecordCount = 0 Then Exit Sub
    NumRegistro = NulosC(RstFrm("registro"))
    pRecibeVerAsiento
    
End Sub

Private Sub Dg1_FilterChange()
    
    Dim xObj As New SGI2_funciones.JC_TDBGrid
    xObj.TDB_FiltroGenerar Dg1, RstFrm
    Set xObj = Nothing
    
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    TxtAsiento.Text = ""
    TxtLibro.Text = ""
    lbl_periodo.Caption = ""
    
    LblIdMes.Caption = Month(Date)
    lbl_periodo.Caption = Busca_Codigo(NulosN(LblIdMes.Caption), "id", "descripcion", "con_meses", "N", xCon)
    
    Configurar_Grilla
    SeEjecuto = True
    
End Sub

Private Sub Form_Deactivate()
    
'    On Error Resume Next
'
'    Err.Clear
'    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()
    SeEjecuto = False
    SGI_JC.CentrarFrm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SGI_JC = Nothing
    Set SGI_JC1 = Nothing
    Set RstFrm = Nothing
    
End Sub





Private Sub Configurar_Grilla()

    With Fg1
        '-----
        .Rows = 2
        .FixedRows = 2
        .Cols = 13
        
        .ColWidth(0) = 200
        '--DATOS DE FILA
        
        SGI_JC1.GRID_COMBINAR Fg1, 0, 1, 0, 7, "DATOS DE LA OPERACIÓN", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        SGI_JC1.GRID_COMBINAR Fg1, 0, 8, 0, 12, "DATOS DE REFERENCIA", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        
''        .FrozenCols = 7
       
        .TextMatrix(1, 1) = "Nª Cuenta":                .ColWidth(1) = 1000:  .ColAlignment(1) = flexAlignLeftCenter:   .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Descripción de la Cuenta": .ColWidth(2) = 3200:  .ColAlignment(2) = flexAlignLeftCenter:   .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "T.C.":                     .ColWidth(3) = 600:   .ColAlignment(3) = flexAlignRightCenter:  .Row = 1: .Col = 3: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 4) = "Debe MN":                  .ColWidth(4) = 1100: .ColAlignment(4) = flexAlignRightCenter:   .Row = 1: .Col = 4: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 5) = "Haber MN":                 .ColWidth(5) = 1100: .ColAlignment(5) = flexAlignRightCenter:   .Row = 1: .Col = 5: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 6) = "Debe ME":                  .ColWidth(6) = 1100: .ColAlignment(6) = flexAlignRightCenter:   .Row = 1: .Col = 6: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 7) = "Haber ME":                 .ColWidth(7) = 1100: .ColAlignment(7) = flexAlignRightCenter:   .Row = 1: .Col = 7: .CellAlignment = flexAlignRightCenter
        
        
        .TextMatrix(1, 8) = "Num.Reg.":                 .ColWidth(8) = 900:   .ColAlignment(8) = flexAlignLeftCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 9) = "T.D.":                     .ColWidth(9) = 750:  .ColAlignment(9) = flexAlignLeftCenter:    .Row = 1: .Col = 9: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 10) = "Nº.Doc":                  .ColWidth(10) = 1000:  .ColAlignment(10) = flexAlignLeftCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 11) = "Fch.Doc":                 .ColWidth(11) = 800:  .ColAlignment(11) = flexAlignLeftCenter:   .Row = 1: .Col = 11: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 12) = "Proveedor/Cliente/Otros": .ColWidth(12) = 1800:  .ColAlignment(12) = flexAlignLeftCenter:   .Row = 1: .Col = 12: .CellAlignment = flexAlignLeftCenter
        
        Fg1.ColFormat(3) = "0.000"
        
        Fg1.ColFormat(4) = SGI_JC1.FORMAT_MONTO
        Fg1.ColFormat(5) = SGI_JC1.FORMAT_MONTO
        Fg1.ColFormat(6) = SGI_JC1.FORMAT_MONTO
        Fg1.ColFormat(7) = SGI_JC1.FORMAT_MONTO
                
        Fg1.ColFormat(11) = SGI_JC1.FORMAT_DATE
        
        SGI_JC1.OCULTAR_COL Fg1, 8, 12
        
    End With
    DoEvents
End Sub


'*******************************

Private Sub CmdBusProv_Click()
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
   
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    Dim xObj As New SGI2_funciones.JC_Varios
    
    xObj.CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_libros  where activo = -1 ORDER BY descripcion ", xCampos(), "Buscando Libro Contable", "descripcion", "descripcion", Principio
    Set xObj = Nothing
    If xRs.State = 1 Then
        TxtLibro.Text = NulosC(xRs("descripcion"))
        LblIdLibro.Caption = NulosC(xRs("id"))
        
        pCargarAsientos
        
        Dg1.SetFocus
    End If
    
    Set xRs = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.index = 3 Then pExportar
    If Button.index = 4 Then pImprimir
    If Button.index = 6 Then
        Unload Me
    End If
End Sub

Private Sub TxtLibro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtLibro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub


Private Sub pCargarAsientos()
    Dim nSQL As String
    
    Set RstFrm = Nothing
    Set Dg1.DataSource = Nothing
    
    If NulosN(LblIdMes.Caption) = 0 Then Exit Sub

    nSQL = "SELECT 0 AS sel, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & [con_diario].[numasi] AS registro " _
        + vbCr + " FROM con_diario LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id " _
        + vbCr + " Where (((con_diario.idlib) = " & NulosN(LblIdLibro.Caption) & ") And ((con_diario.IdMes) = " & NulosN(LblIdMes.Caption) & ")) " _
        + vbCr + " GROUP BY con_diario.numasi, mae_libros.codsun, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & [con_diario].[numasi] " _
        + vbCr + " ORDER BY con_diario.numasi;"
    
    RST_Busq RstFrm, nSQL, xCon

    Dg1.BatchUpdates = False
      
    Dg1.Columns(0).FooterText = "Tot.Reg. " & RstFrm.RecordCount
    
    Set Dg1.DataSource = RstFrm
    

End Sub




Private Sub pExportar()
    Dim xFun As New SGI2_funciones.formularios
    Dim Rst As New ADODB.Recordset
    
    If Fg1.Rows = Fg1.FixedRows Then
        MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "CONSULTA DE ASIENTO Nº. " & NumRegistro, "", "", "Consulta de Asiento"
    Set xFun = Nothing
    
End Sub


Private Sub pImprimir()

    Dim xPrint As New SGI2_funciones.formularios
    
    Me.MousePointer = vbHourglass
    xPrint.Imprimir_x_VSFlexGrid Fg1, "CONSULTA DE ASIENTO Nº. " & NumRegistro, " ", "", False, True
    Set xPrint = Nothing
    Me.MousePointer = vbDefault
  
    
End Sub

