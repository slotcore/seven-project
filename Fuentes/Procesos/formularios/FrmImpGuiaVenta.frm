VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmImpGuiaVenta 
   Caption         =   "Herramientas - Importar Guias de Remision de Ventas"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1650
      Left            =   0
      TabIndex        =   6
      Top             =   345
      Width           =   11985
      Begin VB.CommandButton CmdBusArch 
         Enabled         =   0   'False
         Height          =   240
         Left            =   7050
         Picture         =   "FrmImpGuiaVenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   225
         Width           =   240
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar"
         Height          =   615
         Left            =   10740
         TabIndex        =   18
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdBusBruto 
         Enabled         =   0   'False
         Height          =   240
         Left            =   2595
         Picture         =   "FrmImpGuiaVenta.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   870
         Width           =   240
      End
      Begin VB.CommandButton CmdBusIGV 
         Enabled         =   0   'False
         Height          =   240
         Left            =   2595
         Picture         =   "FrmImpGuiaVenta.frx":0264
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1185
         Width           =   240
      End
      Begin VB.CommandButton CmdBusISC 
         Enabled         =   0   'False
         Height          =   240
         Left            =   8475
         Picture         =   "FrmImpGuiaVenta.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   870
         Width           =   240
      End
      Begin VB.CommandButton CmdBusImpTot 
         Enabled         =   0   'False
         Height          =   240
         Left            =   8475
         Picture         =   "FrmImpGuiaVenta.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1185
         Width           =   240
      End
      Begin VB.CommandButton CmdBusArch2 
         Enabled         =   0   'False
         Height          =   240
         Left            =   7050
         Picture         =   "FrmImpGuiaVenta.frx":05FA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   540
         Width           =   240
      End
      Begin VB.CommandButton CmdBusMes 
         Height          =   240
         Left            =   10230
         Picture         =   "FrmImpGuiaVenta.frx":072C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   540
         Width           =   240
      End
      Begin VB.TextBox TxtMes 
         Enabled         =   0   'False
         Height          =   300
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "TxtMes"
         Top             =   510
         Width           =   1740
      End
      Begin VB.TextBox TxtCtaTot 
         Height          =   300
         Left            =   7185
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "TxtCtaTot"
         Top             =   1155
         Width           =   1560
      End
      Begin VB.TextBox TxtCtaISC 
         Height          =   300
         Left            =   7185
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "TxtCtaISC"
         Top             =   840
         Width           =   1560
      End
      Begin VB.TextBox TxtCtaIGV 
         Height          =   300
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "TxtCtaIGV"
         Top             =   1155
         Width           =   1560
      End
      Begin VB.TextBox TxtCtaBru 
         Height          =   300
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "TxtCtaBru"
         Top             =   840
         Width           =   1560
      End
      Begin VB.TextBox TxtArchivo 
         Height          =   300
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "TxtArchivo"
         Top             =   195
         Width           =   6015
      End
      Begin VB.TextBox TxtArchivo2 
         Height          =   300
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "TxtArchivo2"
         Top             =   510
         Width           =   6015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Arch. Cabecera"
         Height          =   195
         Left            =   135
         TabIndex        =   37
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe Bruto"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   870
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe IGV"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   1185
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Importe I.S.C."
         Height          =   195
         Left            =   6120
         TabIndex        =   34
         Top             =   870
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Importe Total"
         Height          =   195
         Left            =   6150
         TabIndex        =   33
         Top             =   1185
         Width           =   930
      End
      Begin VB.Label Label71 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label71"
         Height          =   300
         Left            =   2880
         TabIndex        =   32
         Top             =   840
         Width           =   3105
      End
      Begin VB.Label Label73 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label73"
         Height          =   300
         Left            =   2880
         TabIndex        =   31
         Top             =   1155
         Width           =   3105
      End
      Begin VB.Label Label72 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label72"
         Height          =   300
         Left            =   8760
         TabIndex        =   30
         Top             =   840
         Width           =   3105
      End
      Begin VB.Label Label74 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label74"
         Height          =   300
         Left            =   8760
         TabIndex        =   29
         Top             =   1155
         Width           =   3105
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Arch. Detalle"
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   540
         Width           =   915
      End
      Begin VB.Label LblIdMes 
         AutoSize        =   -1  'True
         Caption         =   "LblIdMes"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7380
         TabIndex        =   27
         Top             =   150
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   8265
         TabIndex        =   26
         Top             =   540
         Width           =   300
      End
      Begin VB.Label LblIdCtaISC 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCtaISC"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8370
         TabIndex        =   25
         Top             =   150
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label LblIdCtaImpTot 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCtaImpTot"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   9495
         TabIndex        =   24
         Top             =   150
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label LblIdCtaBru 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCtaBru"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7380
         TabIndex        =   23
         Top             =   615
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label LblCtaIGV 
         AutoSize        =   -1  'True
         Caption         =   "LblCtaIGV"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7380
         TabIndex        =   22
         Top             =   390
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1200
      Left            =   3090
      TabIndex        =   0
      Top             =   2940
      Visible         =   0   'False
      Width           =   5805
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   300
         Left            =   135
         TabIndex        =   1
         Top             =   615
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5820
         Y1              =   15
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   30
         Y2              =   1170
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   1
         X1              =   5790
         X2              =   5790
         Y1              =   15
         Y2              =   1155
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   5835
         Y1              =   1185
         Y2              =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Documentos : "
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   1935
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5490
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4890
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":085E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":0DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":1134
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":12B8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":170C
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":1824
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":1D68
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":22AC
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":23C0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":24D4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":2928
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpGuiaVenta.frx":2A94
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   2445
      Left            =   0
      TabIndex        =   3
      Top             =   2310
      Width           =   11985
      _cx             =   21140
      _cy             =   4313
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
      BackColorSel    =   4210816
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmImpGuiaVenta.frx":2FDC
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Inventario Inicial"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg2 
      Height          =   2445
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   11985
      _cx             =   21140
      _cy             =   4313
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
      BackColorSel    =   4210816
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmImpGuiaVenta.frx":31AB
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "[ Documentos de Venta ]"
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
      Height          =   195
      Left            =   45
      TabIndex        =   39
      Top             =   2070
      Width           =   2130
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "[ Items de la Venta ]"
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
      Height          =   195
      Left            =   45
      TabIndex        =   38
      Top             =   4800
      Width           =   1740
   End
End
Attribute VB_Name = "FrmImpGuiaVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
