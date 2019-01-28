VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmConsVenta_Administ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas - Consulta de Ventas"
   ClientHeight    =   8010
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11760
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "FrmConsVenta_Administ.frx":0000
   ScaleHeight     =   8010
   ScaleWidth      =   11760
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2730
      TabIndex        =   33
      Top             =   3630
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   34
         Top             =   345
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Ventas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   1185
         TabIndex        =   37
         Top             =   75
         Width           =   585
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Procesando:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   36
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label lbl 
         Caption         =   "Interrumpir = ESC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   4140
         TabIndex        =   35
         Top             =   75
         Width           =   1530
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5745
         X2              =   5745
         Y1              =   -90
         Y2              =   4800
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -60
         X2              =   6360
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Align           =   2  'Align Bottom
      Height          =   5445
      Left            =   0
      TabIndex        =   17
      Top             =   2565
      Width           =   11760
      _cx             =   20743
      _cy             =   9604
      _ConvInfo       =   1
      Appearance      =   1
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   20
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsVenta_Administ.frx":0C42
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
      TabIndex        =   24
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4860
         Top             =   90
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":0E53
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":1397
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":1729
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":18AD
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":1D01
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":1E19
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":235D
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":28A1
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":29B5
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":2AC9
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":2F1D
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":3089
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":35D1
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Administ.frx":38EB
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2220
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   11745
      Begin VB.CheckBox chkAnioPasados 
         Caption         =   "Considerar Años Anteriores"
         Height          =   195
         Left            =   2130
         TabIndex        =   26
         Top             =   750
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox ChkMostrarItem 
         Caption         =   "Mostrar item"
         Height          =   195
         Left            =   7455
         TabIndex        =   8
         Top             =   735
         Width           =   1275
      End
      Begin VB.Frame Frame6 
         Caption         =   "Ordenar Por"
         Height          =   810
         Left            =   11640
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1320
         Begin VB.OptionButton opt_orden 
            Caption         =   "Fch. Doc"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   570
            Width           =   1125
         End
         Begin VB.OptionButton opt_orden 
            Caption         =   "Nº Doc."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   375
            Width           =   1095
         End
         Begin VB.OptionButton opt_orden 
            Caption         =   "Num. Reg."
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   180
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Seleccionar"
         Height          =   585
         Left            =   3585
         TabIndex        =   18
         Top             =   120
         Width           =   1215
         Begin VB.OptionButton OptEmi 
            Caption         =   "Fch. Emi."
            Height          =   195
            Left            =   60
            TabIndex        =   4
            Top             =   180
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton OptVenc 
            Caption         =   "Fch. Venc."
            Height          =   195
            Left            =   60
            TabIndex        =   28
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Consulta"
         Height          =   585
         Left            =   2055
         TabIndex        =   13
         Top             =   120
         Width           =   1485
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Detallado"
            Height          =   195
            Left            =   135
            TabIndex        =   27
            Top             =   360
            Width           =   1155
         End
         Begin VB.OptionButton OptResum 
            Caption         =   "Resumen"
            Height          =   195
            Left            =   135
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Moneda"
         Height          =   810
         Left            =   4830
         TabIndex        =   12
         Top             =   105
         Width           =   1080
         Begin VB.OptionButton OptDol 
            Caption         =   "Dólares"
            Height          =   195
            Left            =   90
            TabIndex        =   30
            Top             =   570
            Width           =   840
         End
         Begin VB.OptionButton OptSol 
            Caption         =   "Soles"
            Height          =   195
            Left            =   90
            TabIndex        =   29
            Top             =   375
            Width           =   750
         End
         Begin VB.OptionButton OptMonTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   90
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.CommandButton CmdBusProducto 
         Height          =   240
         Left            =   7845
         Picture         =   "FrmConsVenta_Administ.frx":3C7D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   420
         Width           =   225
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg3 
         Height          =   1080
         Left            =   90
         TabIndex        =   10
         Top             =   1065
         Width           =   5790
         _cx             =   10213
         _cy             =   1905
         _ConvInfo       =   1
         Appearance      =   1
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsVenta_Administ.frx":3DAF
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
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   1080
         Left            =   5955
         TabIndex        =   9
         Top             =   1065
         Width           =   5715
         _cx             =   10081
         _cy             =   1905
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   0   'False
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsVenta_Administ.frx":3E13
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec1 
         Height          =   300
         Left            =   630
         TabIndex        =   1
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "25/09/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec2 
         Height          =   300
         Left            =   630
         TabIndex        =   2
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "25/09/2007"
      End
      Begin VB.Frame Frame5 
         Caption         =   "Seleccionar"
         Height          =   810
         Left            =   5970
         TabIndex        =   19
         Top             =   120
         Width           =   1410
         Begin VB.OptionButton OptTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   180
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton OptPend 
            Caption         =   "Pendientes"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   375
            Width           =   1095
         End
         Begin VB.OptionButton OptPag 
            Caption         =   "Pagados"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   570
            Width           =   1125
         End
      End
      Begin VB.TextBox TxtIdTipProd 
         Height          =   300
         Left            =   7455
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "TxtIdTipProd"
         Top             =   390
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Producto"
         Height          =   195
         Left            =   7455
         TabIndex        =   25
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   705
         Width           =   465
      End
      Begin VB.Label lblTipProducto 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTipProducto"
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
         Left            =   8100
         TabIndex        =   14
         Top             =   390
         Width           =   3510
      End
   End
End
Attribute VB_Name = "FrmConsVenta_Administ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vStrCons As String, vFormatString As String, vFormatStrGridItem As String, vFormatGridProv As String
Dim CaracteresNumericos As String

'-- ALMACENAR LOS TOTALES DE TODA LA CONSULTA
Dim Arr_Totales_grls() As Double
Dim Arr_Totales() As Double

Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
                                
Dim Q_POSICION_TOTAL  As Integer '--INDICA LA POCISION DE LA COLUMNA DONDE SE COLOCARA EL NOMBRE DEL TOTAL Y TOTAL_GRL
                                 '--OBTENDRA VALOR EN pGenerarConsulta()
                                
                                
Dim T_RPT_PERIODO As String
Dim T_RPT_TITULO As String

                                
Private Sub CmdBusProducto_Click()
    On Error GoTo error
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "800":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT id, descripcion FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdTipProd.Text = xRs("id")
        lblTipProducto.Caption = xRs("descripcion")
    End If

    Set xform = Nothing
    Set xRs = Nothing
    ChkMostrarItem.SetFocus
    Exit Sub
error:
    Set xform = Nothing
    Set xRs = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"
    
End Sub


Private Sub pConsultar()
    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
       
    Dim vStrSelect As String '--RECIBIR LA CONSULTA
    If Validar_Consulta() = False Then Exit Sub
    vStrSelect = pGenerarConsulta()  '--DEVUELVE LA CONSULTA
    If vStrSelect = "" Then Exit Sub
    BAND_INTERRUMPIR = False
    LimpiarGrid Me.Fg1
    pConfigurarGrilla
    Me.MousePointer = vbHourglass
    DoEvents
    RST_Busq rst_select, vStrSelect, xCon
    PosicionarProgBar
    CARGAR_DATOS_GRILLA rst_select
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Exit Sub
error:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"
End Sub

Private Function CARGAR_DATOS_GRILLA(RST_ORIGEN As ADODB.Recordset) As ADODB.Recordset
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim vStrCampo As String
    Dim vCampos As Long
    Dim BAND_ADD_REG As Boolean
    Dim i As Integer
    
    BAND_ADD_REG = True
    
    vCampos = RST_ORIGEN.Fields.Count
    '--Libera la memoria usada por la matriz.
    Erase Arr_Totales
    Erase Arr_Totales_grls
    
    '--ARRAY QUE ACUMULARA LOS TOTALES
    ReDim Arr_Totales(7, 0)
    ReDim Arr_Totales_grls(7, 0)
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    
    PgBar.Min = 0
    PgBar.Max = RST_ORIGEN.RecordCount
    While Not RST_ORIGEN.EOF
    
    DoEvents
        '--SI SE NTERRUMPE EL PROCESO
        If BAND_INTERRUMPIR = True Then Exit Function
        '------CREANDO LOS GRUPOS
        If ((Me.OptDetalle.Value = True) Or (Me.OptResum.Value = True And (Trim(Me.TxtIdTipProd.Text) <> "" Or Me.ChkMostrarItem.Value = 1))) And RST_ORIGEN.Bookmark = 1 Then
            ADD_REG Fg1
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 5, RST_ORIGEN.Fields("nomcliente") & "", flexAlignLeftCenter: FORMATO_CELDA Fg1, Fg1.Rows - 1, 1
        End If
    
        Comparar_Grupo RST_ORIGEN, BAND_ADD_REG
        ADD_REG Fg1
        '--ASIGNAR LOS DATOS AL RECORDSET TEMPORAL
        For i = 0 To vCampos - 1
            vStrCampo = RST_ORIGEN.Fields(i).Name
            '--OBS: SE VA LLENAR EL ARRAY "MONTOS DEL TOTAL" O "MONTOS DEL RESUMEN"
            Select Case LCase(vStrCampo)
                '--MONTOS DEL TOTAL
                Case "cantdoc":         Arr_Totales(0, 0) = Arr_Totales(0, 0) + NulosN(RST_ORIGEN.Fields("cantdoc"))
                Case "total_dol":       Arr_Totales(1, 0) = Arr_Totales(1, 0) + NulosN(RST_ORIGEN.Fields("total_dol"))
                Case "saldo_dol":       Arr_Totales(2, 0) = Arr_Totales(2, 0) + NulosN(RST_ORIGEN.Fields("saldo_dol"))
                Case "total_mn":        Arr_Totales(3, 0) = Arr_Totales(3, 0) + NulosN(RST_ORIGEN.Fields("total_mn"))
                Case "saldo_mn":        Arr_Totales(4, 0) = Arr_Totales(4, 0) + NulosN(RST_ORIGEN.Fields("saldo_mn"))
                Case "total_sol":       Arr_Totales(5, 0) = Arr_Totales(5, 0) + NulosN(RST_ORIGEN.Fields("total_sol"))
                Case "abono_sol":       Arr_Totales(6, 0) = Arr_Totales(6, 0) + NulosN(RST_ORIGEN.Fields("abono_sol"))
                Case "saldo_sol":       Arr_Totales(7, 0) = Arr_Totales(7, 0) + NulosN(RST_ORIGEN.Fields("saldo_sol"))
                '--MONTOS DEL RESUMEN
                Case "canpro":          Arr_Totales(0, 0) = Arr_Totales(0, 0) + NulosN(RST_ORIGEN.Fields("canpro"))
                Case "total_dol_d":     Arr_Totales(1, 0) = Arr_Totales(1, 0) + NulosN(RST_ORIGEN.Fields("total_dol_d"))
                Case "total_mn_d":      Arr_Totales(2, 0) = Arr_Totales(2, 0) + NulosN(RST_ORIGEN.Fields("total_mn_d"))
                Case "total_sol_d":     Arr_Totales(3, 0) = Arr_Totales(3, 0) + NulosN(RST_ORIGEN.Fields("total_sol_d"))
    '            '''
                Case "total_pu_dol":    Arr_Totales(4, 0) = Arr_Totales(4, 0) + NulosN(RST_ORIGEN.Fields("total_pu_dol"))
                Case "total_pu_mn":     Arr_Totales(5, 0) = Arr_Totales(5, 0) + NulosN(RST_ORIGEN.Fields("total_pu_mn"))
                
            End Select
            
            If Me.OptDetalle.Value = True And Me.ChkMostrarItem.Value = 1 And LCase(vStrCampo) = "total_pu_mn" Then
                '--PARA ACUMULAR LOS REGISTROS ENCONTRDOS POR CLIENTE Y A LA VEZ ACUMULAR LOS REGISTROS ENCONTRADO DE TODA LA CONSULTA
                '--NOS SERVIRA PARA CALCULAR EL PRE. PROM. POR CLIENTE Y PRE. PROM. GRAL
                Arr_Totales(6, 0) = Arr_Totales(6, 0) + 1
            End If
            '--
            Select Case LCase(vStrCampo)
                Case "total_dol", "total_dol", "saldo_dol", "total_mn", "saldo_mn", "total_sol", "abono_sol", "saldo_sol", "total_dol_d", "total_dol_d", "total_mn_d", "total_sol_d"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                Case "total_pu_dol", "total_pu_mn"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                Case "canpro"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_CANTIDAD)
                Case "impven"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_IMPUESTO)
                Case "fchdoc", "fchven"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_DATE)
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
            End Select
    
        Next
        RST_ORIGEN.MoveNext
        '--PONER TOTALES AL FINAL DE LA GRILLA
        If RST_ORIGEN.EOF Then
            If Me.OptDetalle.Value = True Or (Me.TxtIdTipProd.Text <> "" Or Me.ChkMostrarItem.Value = 1) Then
                CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:"
            End If
            If Verificar_Poner_Datos_Grls() = True Then CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", True, True
            '--DEL PRECIO PROMEDIO
            If VERIFICAR_PONER_PRECIO_PROMEDIO() = True Then
                CARGAR_DATOS_GRILLA_ADD_TOTALES True, "P. Prom"
                If Verificar_Poner_Datos_Grls() = True Then CARGAR_DATOS_GRILLA_ADD_TOTALES True, "P. Prom. Gen", True, True
            End If
        Else
            PgBar.Value = CLng(RST_ORIGEN.Bookmark)
        End If
        
    Wend
End Function

Private Sub pImprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO, " ", T_RPT_PERIODO, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR
End Sub
Private Sub ChkMostrarItem_Click()
    If Me.ChkMostrarItem.Value = 0 Then
        Fg2.Enabled = False
    Else
        '--LIMPIAR GRILLA
        Fg2.Enabled = True
        OptTodos.Value = True
        LimpiarGrid Fg2, True, 2
        GRID_COMBOLIST Fg2
    End If

End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim nSQLNotIn  As String
    On Error GoTo error
    If TxtIdTipProd.Text = "" Then
        MsgBox "Falta especificar el tipo de item...!", vbExclamation, xTitulo
        TxtIdTipProd.SetFocus
        Exit Sub
    End If
    If Col = 2 Then
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        
        Dim xCampos(3, 4) As String
        
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Cod. Prod.":    xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2000":         xCampos(1, 3) = "c"
        xCampos(2, 0) = "Id":            xCampos(2, 1) = "id":             xCampos(2, 2) = "600":          xCampos(2, 3) = "N"
                
        nSQLNotIn = GRID_GENERAR_SQL_ID(Fg2, 3, " AND alm_inventario.id", "NOT IN", True)
        
        '--si se ingresa algun filtro adicional
        If NulosC(Fg2.TextMatrix(Row, Col)) <> "" Then
            nSQLNotIn = nSQLNotIn & " AND (UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg2.TextMatrix(Row, Col))) & "%' OR UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg2.TextMatrix(Row, Col))) & "%' ) "
        End If
        
        Fg2.TextMatrix(Row, Col) = ""
        
        xform.SQLCad = "SELECT id, codpro, descripcion FROM alm_inventario WHERE tippro = " & NulosN(TxtIdTipProd.Text) & nSQLNotIn & ""
        
        xform.Titulo = "Buscando Tipo de Item"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 0 Then GoTo salir
        If xRs.RecordCount = 0 Then GoTo salir
        Fg2.TextMatrix(Row, 1) = NulosC(xRs("codpro"))
        Fg2.TextMatrix(Row, 2) = NulosC(xRs("descripcion"))
        Fg2.TextMatrix(Row, 3) = NulosN(xRs("id"))
        
        If Fg2.Row = Fg2.Rows - 1 Then Fg2.AddItem ""
        
        Fg2.Row = Fg2.Rows - 1: Fg2.Col = 2
        
salir:
        Set xform = Nothing
        Set xRs = Nothing
    End If
    Exit Sub
error:
        Set xform = Nothing
        Set xRs = Nothing
        MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub

Private Sub Fg2_DblClick()
    Fg2_CellButtonClick Fg2.Rows - 1, 2
End Sub

Private Sub Fg2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Fg2.Row = -2 Then Exit Sub
    Select Case KeyCode
        Case 45  'INSERTAR REGI
            Fg2.AddItem ""
            Fg2.Row = Fg2.Rows - 1: Fg2.Col = 2
        Case 46 'SUPRIMIR/DELETE
            If Fg2.Rows - 1 >= 2 Then
                Fg2.RemoveItem Fg2.Row
                Fg2.Row = Fg2.Rows - 1: Fg2.Col = 2
            Else
                LimpiarGrid Fg2, True, 2
                GRID_COMBOLIST Fg2
            End If
    End Select
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim nSQLNotIn As String
    On Error GoTo error
    If Col = 2 Then
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        
        Dim xCampos(3, 4) As String
        
        xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Ruc":   xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1500":   xCampos(1, 3) = "C"
        xCampos(2, 0) = "Id":   xCampos(2, 1) = "id":        xCampos(2, 2) = "800":   xCampos(2, 3) = "N"
        
        nSQLNotIn = GRID_GENERAR_SQL_ID(Fg3, 1, " WHERE mae_cliente.id", "NOT IN", True)
        
        '--si se ingresa algun filtro adicional
        If NulosC(Fg3.TextMatrix(Row, Col)) <> "" Then
            nSQLNotIn = IIf(nSQLNotIn = "", " WHERE ", nSQLNotIn & " AND ") & "  (UCASE(mae_cliente.nombre) LIKE '%" & UCase(NulosC(Fg3.TextMatrix(Row, Col))) & "%' OR UCASE(mae_cliente.nombre) LIKE '%" & UCase(NulosC(Fg3.TextMatrix(Row, Col))) & "%' ) "
        End If
        
        Fg3.TextMatrix(Row, Col) = ""
        
        xform.SQLCad = "SELECT * FROM mae_cliente " & nSQLNotIn & " order by nombre asc"
        
        xform.Titulo = "Buscando Clientes"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "nombre"
        xform.CampoBusca = "nombre"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 0 Then GoTo salir
        If xRs.RecordCount = 0 Then GoTo salir
    
        Fg3.TextMatrix(Row, 1) = Trim(xRs("id"))
        Fg3.TextMatrix(Row, 2) = xRs("nombre")
        If Fg3.Row = Fg3.Rows - 1 Then Fg3.AddItem ""
        Fg3.Row = Fg3.Rows - 1: Fg3.Col = 2

salir:
        Set xform = Nothing
        Set xRs = Nothing
    End If
    Exit Sub
error:
        Set xform = Nothing
        Set xRs = Nothing
        MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"
End Sub

Private Sub Fg3_DblClick()
    Fg3_CellButtonClick Fg3.Rows - 1, 2
End Sub

Private Sub Fg3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Fg3.Row = -2 Then Exit Sub
    Select Case KeyCode
        Case 45  'INSERTAR REGI
            Fg3.AddItem ""
            Fg3.Row = Fg3.Rows - 1: Fg3.Col = 2
        Case 46
            If Fg3.Rows - 1 >= 2 Then
                Fg3.RemoveItem Fg3.Row
                Fg3.Row = Fg3.Rows - 1: Fg3.Col = 2
            Else
                LimpiarGrid Fg3, True, 2
                GRID_COMBOLIST Fg3
            End If
    End Select
End Sub

Private Sub Fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        '--interrumpir
        BAND_INTERRUMPIR = True
    End If
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    
    GRID_COMBOLIST Fg2
    GRID_COMBOLIST Fg3
    
    vFormatString = Fg1.FormatString
    Fg2.Tag = Fg2.FormatString
    Fg3.Tag = Fg3.FormatString
 
    TxtIdTipProd.Text = ""
    lblTipProducto.Caption = ""
    CaracteresNumericos = "0123456789." & Chr(8)
    
    TxtFec1.Valor = CDate("01/01/" + CStr(AnoTra))
    TxtFec2.Valor = CDate("31/12/" + CStr(AnoTra))
    
    LimpiarGrid Me.Fg1
    pConfigurarGrilla
End Sub



Private Sub OptDetalle_Click()
    habilitar opt_orden, True
End Sub

Private Sub OptResum_Click()
    habilitar opt_orden, False
End Sub

Private Sub TxtIdTipProd_Change()
    If TxtIdTipProd.Text = "" Then
        lblTipProducto.Caption = ""
        If Me.ChkMostrarItem.Value = 1 Then ChkMostrarItem.Value = 0
        LimpiarGrid Fg2, True
    End If
End Sub

Private Sub TxtIdTipProd_KeyPress(KeyAscii As Integer)
    On Error GoTo error
    If KeyAscii = 13 Then
        Dim RsTipProd As New ADODB.Recordset
        RsTipProd.CursorLocation = adUseClient
        If TxtIdTipProd.Text <> "" Then
            Set RsTipProd = BuscaConCriterio("SELECT id, descripcion FROM mae_tipoproducto WHERE id =" & Val(TxtIdTipProd.Text) & "", xCon)
            If RsTipProd.RecordCount <> 0 Then
                lblTipProducto.Caption = RsTipProd("descripcion")
            Else
                lblTipProducto.Caption = ""
                TxtIdTipProd.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    Set RsTipProd = Nothing
    Exit Sub
error:
    Set RsTipProd = Nothing
    SHOW_ERROR

End Sub

Private Sub TxtIdTipProd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then  'TECHAL F5
        CmdBusProducto.Value = True
    End If
End Sub

'------
Private Function Validar_Consulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    If TxtFec1.Valor = "" Or TxtFec2.Valor = "" Then
        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
        If TxtFec1.Valor = "" Then TxtFec1.SetFocus Else TxtFec2.SetFocus
        Exit Function
    End If
    If CDate(TxtFec1.Valor) > CDate(TxtFec2.Valor) Then
        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
        TxtFec1.SetFocus
        Exit Function
    End If
        If (Year(TxtFec1.Valor) <> Year(TxtFec2.Valor)) Then
        MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
        TxtFec1.SetFocus
        Exit Function
    ElseIf Year(TxtFec1.Valor) <> CStr(AnoTra) Then
        MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
        TxtFec1.SetFocus
        Exit Function
    End If
    Validar_Consulta = True

End Function

Private Function pGenerarConsulta() As String
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--
    Dim vStrSelect As String        '--CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim vStrFiltro_ITEM As String   '--SOLO ITEM
    Dim vStrFiltro_CLI As String    '--SOLO CLIENTES
    Dim vStrTipoProducto As String
    Dim vStrFecha As String
    Dim vStrFiltro As String
    Dim vStrFiltro_1 As String      '--ESTE FILTRO SERVIRA PARA pConsultar EN EL SUB_SELECT
    Dim vFiltro As String
    Dim k  As Integer
    
    '--DE LA FECHA
    If CDate(TxtFec1.Valor) < CDate(TxtFec2.Valor) Then
        'vStrFecha = " vta_ventas.fchdoc >= cdate('" + Format(TxtFec1.Valor, "dd/mm/yyyy") + "') AND vta_ventas.fchdoc<= cdate('" + Format(TxtFec2.Valor, "dd/mm/yyyy") + "') "
        vStrFecha = " vta_ventas.fchdoc between cdate('" + Format(TxtFec1.Valor, "dd/mm/yyyy") + "') AND cdate('" + Format(TxtFec2.Valor, "dd/mm/yyyy") + "') "
        T_RPT_PERIODO = " Del: " + TxtFec1.Valor + " Al: " + TxtFec2.Valor
    Else
        vStrFecha = " vta_ventas.fchdoc = cdate('" + Format(TxtFec1.Valor, "dd/mm/yyyy") + "') "
        T_RPT_PERIODO = "Al: " + TxtFec2.Valor
    End If
    '--SI OPCION DE SELECCIONAR POR FECHA DE VENCIMIENTO
    If Me.OptVenc.Value = True Then vStrFecha = Replace(vStrFecha, "vta_ventas.fchdoc", "vta_ventas.fchven")
    
    '--DEL TIPO DE PRODUCTO
    If TxtIdTipProd.Text <> "" Then vFiltro = vFiltro + " AND alm_inventario.tippro = " + CStr(TxtIdTipProd.Text) + " "
    
    '--DEL ITEM
    vFiltro = vFiltro & GRID_GENERAR_SQL_ID(Fg2, 3, " AND alm_inventario.id", "IN")
    'If vStrFiltro_ITEM <> "" Then vFiltro = vFiltro + " AND " + vStrFiltro_ITEM
    
    '--DEL CLIENTE
    vFiltro = vFiltro & GRID_GENERAR_SQL_ID(Fg3, 1, " AND vta_ventas.idcli", "IN")
    'If vStrFiltro_CLI <> "" Then vFiltro = vFiltro + " AND " + vStrFiltro_CLI
 
    '--DE LA MONEDA
    If OptSol.Value = True Then vFiltro = vFiltro + " AND vta_ventas.idmon= 1 "       '--SOLES
    If Me.OptDol.Value = True Then vFiltro = vFiltro + " AND vta_ventas.idmon= 2 "    '--DOLARES
    '---------------
    
    If OptPag.Value = True Then         '---SI ES CANCELADO
        vFiltro = vFiltro + " AND vta_ventas.impsal = 0 "
        
    ElseIf OptPend.Value = True Then    '---SI ES PENDIENTE DE PAGO
        vFiltro = vFiltro + " AND vta_ventas.impsal > 0 "
        
    End If
        
    vStrFiltro = " vta_ventas.anulado = 0 AND " + vStrFecha + vFiltro
    
    If OptPag.Value = False Then
        If chkAnioPasados.Value = 1 Then
            vStrFiltro = "(" + vStrFiltro + ") OR ( vta_ventas.anulado = 0 AND year(vta_ventas.fchdoc)<> " + AnoTra + " " + vFiltro + " )"
        End If
    End If
    
    '------------------------------------------------------------------------------------
    vStrFiltro_1 = Replace(vStrFiltro, "vta_ventas.", "vta_ventas1.")
    vStrFiltro_1 = Replace(vStrFiltro_1, "alm_inventario.", "alm_inventario1.")
    
    If OptResum.Value = True Then '--RESUMEN
        If ChkMostrarItem.Value = 1 Or TxtIdTipProd.Text <> "" Then
            If TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0 Then
            '--MOSTRAR SOLO PRODUCTO
   
    
        T_RPT_TITULO = "REPORTE DE VENTAS RESUMIDO POR TIPO PRODUCTO"
        vStrSelect = "SELECT DISTINCT mae_cliente.numruc, mae_cliente.nombre AS nomcliente, mae_moneda.simbolo, mae_tipoproducto.descripcion AS desctipcom, " _
            & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=2,vta_ventasdet1.imptot,0)) AS total_dol_d " _
            & vbCr + " FROM (vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha) LEFT JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN vta_ventasdet AS vta_ventasdet1 ON alm_inventario1.id = vta_ventasdet1.iditem) ON vta_ventas1.id = vta_ventasdet1.idvta " _
            & vbCr + " Where " + vStrFiltro_1 _
            & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon, alm_inventario1.tippro HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon=vta_ventas.idmon AND alm_inventario1.tippro=alm_inventario.tippro ) as  total_dol_d, " _
            & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=1,vta_ventasdet1.imptot,0)) AS total_mn_d " _
            & vbCr + " FROM (vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha) INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN vta_ventasdet AS vta_ventasdet1 ON alm_inventario1.id = vta_ventasdet1.iditem) ON vta_ventas1.id = vta_ventasdet1.idvta " _
            & vbCr + " Where " + vStrFiltro_1 _
            & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon, alm_inventario1.tippro HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon=vta_ventas.idmon AND alm_inventario1.tippro=alm_inventario.tippro ) as total_mn_d, " _
            & vbCr + " (SELECT Sum((IIf(vta_ventas1.idmon=1,vta_ventasdet1.imptot,0)+IIf(vta_ventas1.idmon=1,0,IIf(con_tc1.impven Is Null,0,con_tc1.impven*vta_ventasdet1.imptot)))) AS total_mn_sol " _
            & vbCr + " FROM (vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha) INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN vta_ventasdet AS vta_ventasdet1 ON alm_inventario1.id = vta_ventasdet1.iditem) ON vta_ventas1.id = vta_ventasdet1.idvta " _
            & vbCr + " Where " + vStrFiltro_1 _
            & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon, alm_inventario1.tippro HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon=vta_ventas.idmon AND alm_inventario1.tippro=alm_inventario.tippro ) as total_sol_d "
            vStrSelect = vStrSelect _
            & vbCr + " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN vta_ventas ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) INNER JOIN (mae_tipoproducto RIGHT JOIN (alm_inventario RIGHT JOIN vta_ventasdet ON alm_inventario.id = vta_ventasdet.iditem) ON mae_tipoproducto.id = alm_inventario.tippro) ON vta_ventas.id = vta_ventasdet.idvta " _
            & vbCr + " Where " _
            & vStrFiltro _
            & vbCr + " ORDER BY mae_cliente.nombre, mae_moneda.simbolo, mae_tipoproducto.descripcion;"
    
                Q_POSICION_TOTAL = 4
                       
            Else
            '--MOSTRAR PRODUCTO Y ITEM
                T_RPT_TITULO = "REPORTE DE VENTAS RESUMIDO POR TIPO PRODUCTO CON ITEM"
                vStrSelect = "SELECT DISTINCT mae_cliente.numruc, mae_cliente.nombre AS nomcliente, mae_moneda.simbolo, mae_tipoproducto.descripcion AS desctipcom, alm_inventario.descripcion, mae_unidades.abrev, " _
                & vbCr + " (SELECT Sum(vta_ventasdet1.canpro) AS canpro " _
                & vbCr + " FROM (vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha) INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN vta_ventasdet AS vta_ventasdet1 ON alm_inventario1.id = vta_ventasdet1.iditem) ON vta_ventas1.id = vta_ventasdet1.idvta " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon, alm_inventario1.tippro, alm_inventario1.id, alm_inventario1.idunimed " _
                & vbCr + " HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon=vta_ventas.idmon AND alm_inventario1.tippro=alm_inventario.tippro and alm_inventario1.id =alm_inventario.id and  alm_inventario1.idunimed = alm_inventario.idunimed ) as canpro, " _
                & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=2,vta_ventasdet1.imptot,0)) AS total_dol_d " _
                & vbCr + " FROM (vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha) INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN vta_ventasdet AS vta_ventasdet1 ON alm_inventario1.id = vta_ventasdet1.iditem) ON vta_ventas1.id = vta_ventasdet1.idvta " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon, alm_inventario1.tippro,alm_inventario1.id HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon=vta_ventas.idmon AND alm_inventario1.tippro=alm_inventario.tippro and alm_inventario1.id =alm_inventario.id )  as  total_dol_d, " _
                & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=1,vta_ventasdet1.imptot,0)) AS total_mn_d " _
                & vbCr + " FROM (vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha) INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN vta_ventasdet AS vta_ventasdet1 ON alm_inventario1.id = vta_ventasdet1.iditem) ON vta_ventas1.id = vta_ventasdet1.idvta " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon, alm_inventario1.tippro,alm_inventario1.id " _
                & vbCr + " HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon=vta_ventas.idmon AND alm_inventario1.tippro=alm_inventario.tippro and alm_inventario1.id =alm_inventario.id ) as total_mn_d, " _
                & vbCr + " (SELECT Sum((IIf(vta_ventas1.idmon=1,vta_ventasdet1.imptot,0)+IIf(vta_ventas1.idmon=1,0,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*vta_ventasdet1.imptot))))) AS total_sol_d " _
                & vbCr + " FROM (vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha) INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN vta_ventasdet AS vta_ventasdet1 ON alm_inventario1.id = vta_ventasdet1.iditem) ON vta_ventas1.id = vta_ventasdet1.idvta " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon, alm_inventario1.tippro,alm_inventario1.id " _
                & vbCr + " HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon=vta_ventas.idmon AND alm_inventario1.tippro=alm_inventario.tippro and alm_inventario1.id =alm_inventario.id ) as  total_sol_d"
                vStrSelect = vStrSelect _
                & vbCr + " FROM mae_unidades RIGHT JOIN ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN vta_ventas ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) INNER JOIN (mae_tipoproducto RIGHT JOIN (alm_inventario RIGHT JOIN vta_ventasdet ON alm_inventario.id = vta_ventasdet.iditem) ON mae_tipoproducto.id = alm_inventario.tippro) ON vta_ventas.id = vta_ventasdet.idvta) ON mae_unidades.id = alm_inventario.idunimed " _
                & vbCr + " Where " _
                & vStrFiltro _
                & vbCr + " ORDER BY mae_cliente.nombre, mae_moneda.simbolo, mae_tipoproducto.descripcion, alm_inventario.descripcion;"
    
                Q_POSICION_TOTAL = 5
            End If
        Else '--GENERAL
                
                T_RPT_TITULO = "REPORTE DE VENTAS RESUMIDO POR CLIENTE"
                vStrSelect = "SELECT DISTINCT mae_cliente.numruc, mae_cliente.nombre AS nomcliente, " _
                & vbCr + " (SELECT Count(vta_ventas1.numdoc)  AS CuentaDeid " _
                & vbCr + " FROM vta_ventas AS vta_ventas1 " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon = vta_ventas.idmon) as cantdoc, " _
                & vbCr + " mae_moneda.simbolo, " _
                & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=2,vta_ventas1.imptotdoc,0)) " _
                & vbCr + " FROM vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon = vta_ventas.idmon) AS total_dol, " _
                & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=2,vta_ventas1.impsal,0)) " _
                & vbCr + " FROM vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon = vta_ventas.idmon)  AS saldo_dol, " _
                & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=1,vta_ventas1.imptotdoc,0)) AS total_mn1 " _
                & vbCr + " FROM vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon = vta_ventas.idmon) as total_mn, " _
                & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=1,vta_ventas1.impsal,0)) AS saldo_mn " _
                & vbCr + " FROM vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon = vta_ventas.idmon)  as  saldo_mn, "
                vStrSelect = vStrSelect _
                & vbCr + " (SELECT   Sum(IIf(vta_ventas1.idmon=1,vta_ventas1.imptotdoc,0)+IIf(vta_ventas1.idmon=1,0,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*vta_ventas1.imptotdoc)))) AS total_sol " _
                & vbCr + " FROM vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon = vta_ventas.idmon)  as total_sol, " _
                & vbCr + " (total_sol - saldo_sol) as abono_sol , " _
                & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=1,vta_ventas1.impsal,0)+IIf(vta_ventas1.idmon=1,0,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*vta_ventas1.impsal)))) AS saldo_sol " _
                & vbCr + " FROM vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon = vta_ventas.idmon) as saldo_sol " _
                & vbCr + " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN vta_ventas ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha "
                vStrSelect = vStrSelect _
                & vbCr + " Where " _
                & vStrFiltro _
                & vbCr + " ORDER BY mae_cliente.nombre, mae_moneda.simbolo;"
                
            Q_POSICION_TOTAL = 2
        End If
    
    
    Else '--DETALLADO
        If ChkMostrarItem.Value = 1 Or TxtIdTipProd.Text <> "" Then
            If TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0 Then '--MOSTRAR SOLO PRODUCTO
                
                T_RPT_TITULO = "REPORTE DETALLADO POR TIPO PRODUCTO"
                
                vStrSelect = "SELECT DISTINCT IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='',[vta_ventas].[numreg],Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4)) AS registro,mae_documento.abrev AS tdocabrev, mae_cliente.nombre AS nomcliente, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, IIf(vta_ventas.impsal<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, con_tc.impven, mae_tipoproducto.descripcion AS desctipcom, " _
                & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=2,vta_ventasdet1.imptot,0)) AS total_dol_d " _
                & vbCr + " FROM (vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha) INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN vta_ventasdet AS vta_ventasdet1 ON alm_inventario1.id = vta_ventasdet1.iditem) ON vta_ventas1.id = vta_ventasdet1.idvta " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon, alm_inventario1.tippro, alm_inventario1.id HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon= vta_ventas.idmon  AND alm_inventario1.tippro= alm_inventario.tippro  AND alm_inventario1.id=alm_inventario.id ) AS total_dol_d, " _
                & vbCr + " (SELECT Sum(IIf(vta_ventas1.idmon=1,vta_ventasdet1.imptot,0)) AS total_mn_d " _
                & vbCr + " FROM (vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha) INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN vta_ventasdet AS vta_ventasdet1 ON alm_inventario1.id = vta_ventasdet1.iditem) ON vta_ventas1.id = vta_ventasdet1.idvta " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon, alm_inventario1.tippro, alm_inventario1.id HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon= vta_ventas.idmon  AND  alm_inventario1.tippro= alm_inventario.tippro  AND alm_inventario1.id=alm_inventario.id )  AS total_mn_d, " _
                & vbCr + " (SELECT Sum((IIf(vta_ventas1.idmon=1,vta_ventasdet1.imptot,0)+IIf(vta_ventas1.idmon=1,0,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*vta_ventasdet1.imptot))))) AS total_sol_d " _
                & vbCr + " FROM (vta_ventas AS vta_ventas1 LEFT JOIN con_tc AS con_tc1 ON vta_ventas1.fchdoc = con_tc1.fecha) INNER JOIN (alm_inventario AS alm_inventario1 RIGHT JOIN vta_ventasdet AS vta_ventasdet1 ON alm_inventario1.id = vta_ventasdet1.iditem) ON vta_ventas1.id = vta_ventasdet1.idvta " _
                & vbCr + " Where " + vStrFiltro_1 _
                & vbCr + " GROUP BY vta_ventas1.idcli, vta_ventas1.idmon, alm_inventario1.tippro, alm_inventario1.id HAVING vta_ventas1.idcli=vta_ventas.idcli AND vta_ventas1.idmon= vta_ventas.idmon  AND alm_inventario1.tippro= alm_inventario.tippro  AND alm_inventario1.id=alm_inventario.id ) AS total_sol_d "
                vStrSelect = vStrSelect _
                & vbCr + " FROM (((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_condpago RIGHT JOIN vta_ventas ON mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) INNER JOIN ((alm_inventario RIGHT JOIN vta_ventasdet ON alm_inventario.id = vta_ventasdet.iditem) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
                & vbCr + " Where " _
                & vStrFiltro _
                & vbCr + " ORDER BY mae_cliente.nombre, vta_ventas.fchdoc "
                ', mae_moneda.simbolo, mae_tipoproducto.descripcion, vta_ventas.fchdoc;
                
                Q_POSICION_TOTAL = 6
            Else
            '--MOSTRAR PRODUCTO Y ITEM
                T_RPT_TITULO = "REPORTE DE VENTAS DETALLADO POR TIPO PRODUCTO CON ITEM"
                vStrSelect = "SELECT IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='',[vta_ventas].[numreg],Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4)) AS registro,mae_documento.abrev AS tdocabrev, mae_cliente.nombre AS nomcliente, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, " _
                           & vbCr + " IIf(vta_ventas.impsal<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, con_tc.impven, mae_tipoproducto.descripcion AS desctipcom, alm_inventario.descripcion, mae_unidades.abrev AS prodabrev, vta_ventasdet.canpro, IIf(vta_ventas.idmon=2,vta_ventasdet.preuni,0) AS total_pu_dol, IIf(vta_ventas.idmon=2,vta_ventasdet.imptot,0) AS total_dol_d, IIf(vta_ventas.idmon=1,vta_ventasdet.preuni,0) AS total_pu_mn, IIf(vta_ventas.idmon=1,vta_ventasdet.imptot,0) AS total_mn_d ,(IIf(vta_ventas.idmon=2,iif(con_tc.impven is null,0,vta_ventasdet.imptot * con_tc.impven),0) + total_mn_d) as total_sol_d " _
                           & vbCr + " FROM (((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_condpago RIGHT JOIN vta_ventas ON mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) INNER JOIN (mae_unidades RIGHT JOIN ((alm_inventario RIGHT JOIN vta_ventasdet ON alm_inventario.id = vta_ventasdet.iditem) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id  " _
                           & vbCr + " WHERE  " _
                           & vStrFiltro _
                           & vbCr + " ORDER BY mae_cliente.nombre,vta_ventas.fchdoc"
                           ' mae_moneda.simbolo, mae_tipoproducto.descripcion, alm_inventario.descripcion, vta_ventas.fchdoc;
                Q_POSICION_TOTAL = 6
            End If
        Else '--MOSTRAR SIN DETALLE
            T_RPT_TITULO = "REPORTE DE VENTAS DETALLADO POR CLIENTE"
            vStrSelect = "SELECT IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='',[vta_ventas].[numreg],Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4)) AS registro, mae_documento.abrev AS tdocabrev, mae_cliente.nombre AS nomcliente, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, IIf(vta_ventas.impsal<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, con_tc.impven, IIf(vta_ventas.idmon=2,vta_ventas.imptotdoc,0) AS total_dol, IIf(vta_ventas.idmon=2,vta_ventas.impsal,0) AS saldo_dol, IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc,0) AS total_mn, IIf(vta_ventas.idmon=1,vta_ventas.impsal,0) AS saldo_mn, IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc,0)+IIf(vta_ventas.idmon=1,0,IIf(con_tc.impven Is Null,0,(con_tc.impven*vta_ventas.imptotdoc))) AS total_sol, " _
                & vbCr + " ((IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc,0)+IIf(vta_ventas.idmon=1,0,IIf(con_tc.impven Is Null,0,(con_tc.impven*vta_ventas.imptotdoc))))-(IIf(vta_ventas.idmon=1,vta_ventas.impsal,0)+IIf(vta_ventas.idmon=1,0,IIf(con_tc.impven Is Null,0,(con_tc.impven*vta_ventas.impsal))))) AS abono_sol, IIf(vta_ventas.idmon=1,vta_ventas.impsal,0)+IIf(vta_ventas.idmon=1,0,IIf(con_tc.impven Is Null,0,(con_tc.impven*vta_ventas.impsal))) AS saldo_sol    " _
                & vbCr + " FROM ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_condpago RIGHT JOIN vta_ventas ON mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
                & vbCr + " WHERE " _
                & vStrFiltro _
                & vbCr + " ORDER BY mae_cliente.nombre asc ,vta_ventas.fchdoc "
                ',mae_moneda.simbolo
            Q_POSICION_TOTAL = 6
        End If
    End If
    '------------------------------------------------------------------------------------
    pGenerarConsulta = vStrSelect
End Function



'--011007
Private Sub Comparar_Grupo(RST_ORIGEN As ADODB.Recordset, BAND_ADD_REG As Boolean)
    '--FUNCION QUE NOS PERMITE ARMAR LOS GRUPOS POR EL CLIENTE
    '--CUANDO SE GENERA EL GRUPO SE ARGEGA EL NOMBRE DEL CLIENTE COMO CABECERA
    '--COMPARA CUANDO CAMBIAR DE GRUPO
    Dim RST_TEPM_1 As New ADODB.Recordset
    
    Set RST_TEPM_1 = RST_ORIGEN.Clone
    RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
    RST_TEPM_1.MovePrevious

    If RST_ORIGEN.Bookmark = 1 Then
        If OptDetalle.Value = False Then
            'ADD_REG Fg1
        End If
        Exit Sub
    End If
    
    '---------------------------------------------------------
    If RST_ORIGEN.Bookmark <> 1 Then
        If NulosC(RST_TEPM_1.Fields("nomcliente")) <> NulosC(RST_ORIGEN.Fields("nomcliente")) Then  '--CLIENTE
            If Me.OptResum.Value = True And (Trim(Me.TxtIdTipProd.Text) = "" And Me.ChkMostrarItem.Value = 0) Then Exit Sub
            CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:"
            '--DEL PRECIO PROMEDIO
            If VERIFICAR_PONER_PRECIO_PROMEDIO() = True Then
                CARGAR_DATOS_GRILLA_ADD_TOTALES True, "P. Prom"
            End If
            ADD_REG Fg1
            Limpiar_ARRAY_TOTAL
            If OptDetalle.Value = True Or (Me.OptResum.Value = True And (Trim(Me.TxtIdTipProd.Text) <> "" Or Me.ChkMostrarItem.Value = 1)) Then
                ADD_REG Fg1
                UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 5, NulosC(RST_ORIGEN.Fields("nomcliente")), flexAlignLeftCenter:      FORMATO_CELDA Fg1, Fg1.Rows - 1, 1
            End If
            Exit Sub
        End If
    End If
    Set RST_TEPM_1 = Nothing
End Sub

Private Sub Limpiar_ARRAY_TOTAL()
    Erase Arr_Totales
    ReDim Arr_Totales(7, 0) As Double
End Sub

Private Sub CARGAR_DATOS_GRILLA_ADD_TOTALES(BAND_ADD_TOTAL As Boolean, Nombre_total As String, Optional Band_Total_gral As Boolean = False, Optional band_forzar_suma As Boolean = False)
    '--AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL
    '--ACUMULA LOS TOTALES EN EL TOTAL GENERAL
    Dim X_ROW As Long
    Dim k As Integer
    
    'On Error Resume Next
    X_ROW = Fg1.Rows - 1
    If BAND_ADD_TOTAL = True Then
        ADD_REG Fg1
        X_ROW = Fg1.Rows - 1
        'PONIENDO LOS NOMBRES DE LOS TOTALES
        Fg1.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
        FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL
    End If
    
    '-----------------------------------------------------------------------------
    '--ACUMULANDO LOS TOTALES GRLES
    If Me.OptResum.Value = True Then    '--RESUMEN
        If Band_Total_gral = False And (Me.TxtIdTipProd.Text <> "" Or Me.ChkMostrarItem.Value = 1) Then
            For k = 0 To UBound(Arr_Totales())
                Arr_Totales_grls(k, 0) = Arr_Totales_grls(k, 0) + Arr_Totales(k, 0)
            Next k
        End If
    Else
        If Band_Total_gral = False Then     '--DETALLE
            For k = 0 To UBound(Arr_Totales())
                Arr_Totales_grls(k, 0) = Arr_Totales_grls(k, 0) + Arr_Totales(k, 0)
            Next k
        End If
    End If
    '-----------------------------------------------------------------------------
    
    '
    If Me.OptResum.Value = True Then
        '--RESUMEN
            With Fg1
            If Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0 Then '--PRODUCTO
                .TextMatrix(X_ROW, 5) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 5    '"Imp. $"
                .TextMatrix(X_ROW, 6) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 6   '"Imp. S/."
                .TextMatrix(X_ROW, 7) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 7    '"Total"
            ElseIf Me.ChkMostrarItem.Value = 1 Then '--PRODUCTO Y ITEM
                .TextMatrix(X_ROW, 7) = Format(IIf(Band_Total_gral = False, Arr_Totales(0, 0), Arr_Totales_grls(0, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 7    '"Imp. $"
                .TextMatrix(X_ROW, 8) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 8    '"Imp. $"
                .TextMatrix(X_ROW, 9) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 9    '"Imp. S/."
                .TextMatrix(X_ROW, 10) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 10    '"Total"
            Else
                .TextMatrix(X_ROW, 3) = Format(IIf(Band_Total_gral = False, Arr_Totales(0, 0), Arr_Totales_grls(0, 0)), FORMAT_CANTIDAD):: FORMATO_CELDA Fg1, X_ROW, 3    '"# Doc"
                .TextMatrix(X_ROW, 5) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 5   '"Imp. $"
                .TextMatrix(X_ROW, 6) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 6    '"Saldo $"
                .TextMatrix(X_ROW, 7) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 7    ' "Imp. S/."
                .TextMatrix(X_ROW, 8) = Format(IIf(Band_Total_gral = False, Arr_Totales(4, 0), Arr_Totales_grls(4, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 8   '"Saldo S/."
                .TextMatrix(X_ROW, 9) = Format(IIf(Band_Total_gral = False, Arr_Totales(5, 0), Arr_Totales_grls(5, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 9    '"Total"
                .TextMatrix(X_ROW, 10) = Format(IIf(Band_Total_gral = False, Arr_Totales(6, 0), Arr_Totales_grls(6, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 10    '"Abono"
                .TextMatrix(X_ROW, 11) = Format(IIf(Band_Total_gral = False, Arr_Totales(7, 0), Arr_Totales_grls(7, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 11    ' "Saldo"
            End If
        End With
    Else '-DETALLE
        With Fg1
            If Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0 Then '--PRODUCTO
                .TextMatrix(X_ROW, 12) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 12     '"Imp. $"
                .TextMatrix(X_ROW, 13) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 13     '"Imp. S/."
                .TextMatrix(X_ROW, 14) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 14    '"Total"
            ElseIf Me.ChkMostrarItem.Value = 1 Then '--PRODUCTO E ITEM
                If VERIFICAR_PONER_PRECIO_PROMEDIO() = True And (Nombre_total = "P. Prom" Or Nombre_total = "P. Prom. Gen") Then
                    'Calcular_Precio_Promedio (Band_Total_gral)
                    .TextMatrix(X_ROW, 15) = CALCULAR_PRECIO_PROMEDIO(Band_Total_gral, 4): FORMATO_CELDA Fg1, X_ROW, 15 '"PRECIO PROM DOL
                    .TextMatrix(X_ROW, 17) = CALCULAR_PRECIO_PROMEDIO(Band_Total_gral, 5): FORMATO_CELDA Fg1, X_ROW, 17 '"PRECIO PROM SOL
                    Exit Sub
                End If
    
            
                .TextMatrix(X_ROW, 14) = Format(IIf(Band_Total_gral = False, Arr_Totales(0, 0), Arr_Totales_grls(0, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 14 '"cantidad"
                .TextMatrix(X_ROW, 16) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 16   '"Imp. Total $"
                .TextMatrix(X_ROW, 18) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 18   '"Imp. Total S/."
                .TextMatrix(X_ROW, 19) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 19   '"Total "
                 
            Else '--SIN PRODUCTO E ITEM
                .TextMatrix(X_ROW, 11) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 11   '"Imp. $"
                .TextMatrix(X_ROW, 12) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 12    '"Saldo $"
                .TextMatrix(X_ROW, 13) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 13     '"Imp. S/."
                .TextMatrix(X_ROW, 14) = Format(IIf(Band_Total_gral = False, Arr_Totales(4, 0), Arr_Totales_grls(4, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 14    '"Saldo S/."
                .TextMatrix(X_ROW, 15) = Format(IIf(Band_Total_gral = False, Arr_Totales(5, 0), Arr_Totales_grls(5, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 15     '"Total"
                .TextMatrix(X_ROW, 16) = Format(IIf(Band_Total_gral = False, Arr_Totales(6, 0), Arr_Totales_grls(6, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 16    '"Abono"
                .TextMatrix(X_ROW, 17) = Format(IIf(Band_Total_gral = False, Arr_Totales(7, 0), Arr_Totales_grls(7, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 17     '"Saldo"
            End If
    
        End With
    End If
    'Err.Clear
End Sub

    

Private Sub pConfigurarGrilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
    If F_CONSERVAR_FORMATO = True Then
        Fg1.Clear
        Fg1.FormatString = vFormatString
    End If
    Fg1.FrozenCols = 0
    If Me.OptResum.Value = True Then '--RESUMEN
        With Fg1
            If Trim(Me.TxtIdTipProd.Text) <> "" Or Me.ChkMostrarItem.Value = 1 Then
                .ColWidth(1) = 0 'RUC
                .ColWidth(2) = 0 'CLIENTE
            Else
                .TextMatrix(1, 1) = "RUC":          .ColWidth(1) = 1200:    .ColAlignment(1) = flexAlignCenterBottom
                .TextMatrix(1, 2) = "Cliente":      .ColWidth(2) = 2500:    .ColAlignment(2) = flexAlignLeftBottom
            End If
            If Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0 Then
                '--SOLO PRODUCTO
                '.FrozenCols = 4
                .TextMatrix(1, 3) = "M":        .ColWidth(3) = 500:     .ColAlignment(3) = flexAlignLeftBottom
                .TextMatrix(1, 4) = "Producto": .ColWidth(4) = 3500:    .ColAlignment(4) = flexAlignLeftBottom
                .TextMatrix(1, 5) = "Imp. $":   .ColWidth(5) = 1100:    .ColAlignment(5) = flexAlignRightBottom
                .TextMatrix(1, 6) = "Imp. S/.": .ColWidth(6) = 1100:    .ColAlignment(6) = flexAlignRightBottom
                .TextMatrix(1, 7) = "Total S/.": .ColWidth(7) = 1400:   .ColAlignment(7) = flexAlignRightBottom
                '--SOLO DOLARES OCULTAR SOLES
                If Me.OptDol.Value = True Then .ColWidth(6) = 0
                '--SOLO SOLES OCULTAR DOLARES
                If Me.OptSol.Value = True Then .ColWidth(5) = 0
                UNIR_CELDAS Fg1, 0, 1, 0, 4, " "
                UNIR_CELDAS Fg1, 0, 5, 0, 7, "TOTAL"
                OCULTAR_COL Fg1, 8, 19
            ElseIf Me.ChkMostrarItem.Value = 1 Then
                '--CON PRODUCTO E ITEM
                .FrozenCols = 5
                .TextMatrix(1, 3) = "M":            .ColWidth(3) = 500:     .ColAlignment(3) = flexAlignLeftBottom
                .TextMatrix(1, 4) = "Producto":     .ColWidth(4) = 1000:    .ColAlignment(4) = flexAlignLeftBottom
                .TextMatrix(1, 5) = "Item":         .ColWidth(5) = 3500:    .ColAlignment(5) = flexAlignLeftBottom
                .TextMatrix(1, 6) = "U.M.":         .ColWidth(6) = 500:     .ColAlignment(6) = flexAlignLeftBottom
                .TextMatrix(1, 7) = "Cant.":        .ColWidth(7) = 800:     .ColAlignment(7) = flexAlignRightBottom
                
                .TextMatrix(1, 8) = "Imp. $":       .ColWidth(8) = 1100:    .ColAlignment(8) = flexAlignRightBottom
                .TextMatrix(1, 9) = "Imp. S/.":     .ColWidth(9) = 1100:    .ColAlignment(9) = flexAlignRightBottom
                .TextMatrix(1, 10) = "Total S/.":   .ColWidth(10) = 1400:   .ColAlignment(10) = flexAlignRightBottom
                 '--SOLO DOLARES OCULTAR SOLES
                If Me.OptDol.Value = True Then .ColWidth(9) = 0
                '--SOLO SOLES OCULTAR DOLARES
                If Me.OptSol.Value = True Then .ColWidth(8) = 0
                UNIR_CELDAS Fg1, 0, 1, 0, 5, " "
                UNIR_CELDAS Fg1, 0, 8, 0, 10, "TOTAL"
                OCULTAR_COL Fg1, 11, 19
            Else
                .FrozenCols = 4
                .TextMatrix(1, 3) = "# Doc":        .ColWidth(3) = 650:     .ColAlignment(3) = flexAlignLeftBottom
                .TextMatrix(1, 4) = "M":            .ColWidth(4) = 500:     .ColAlignment(4) = flexAlignLeftBottom
                .TextMatrix(1, 5) = "Imp.":         .ColWidth(5) = 1000:    .ColAlignment(5) = flexAlignRightBottom
                .TextMatrix(1, 6) = "Saldo":        .ColWidth(6) = 1000:    .ColAlignment(6) = flexAlignRightBottom
                .TextMatrix(1, 7) = "Imp.":         .ColWidth(7) = 1000:    .ColAlignment(7) = flexAlignRightBottom
                .TextMatrix(1, 8) = "Saldo":        .ColWidth(8) = 1000:    .ColAlignment(8) = flexAlignRightBottom
                .TextMatrix(1, 9) = "Total":        .ColWidth(9) = 1200:    .ColAlignment(9) = flexAlignRightBottom
                .TextMatrix(1, 10) = "Abono":       .ColWidth(10) = 1200:   .ColAlignment(10) = flexAlignRightBottom
                .TextMatrix(1, 11) = "Saldo":       .ColWidth(11) = 1200:   .ColAlignment(11) = flexAlignRightBottom
                '--SOLO PAGADO OCULTAR SALDOS
                If Me.OptPag.Value = True Then .ColWidth(6) = 0: .ColWidth(8) = 0: .ColWidth(11) = 0
                '--SOLO DOLARES OCULTAR SOLES
                If Me.OptDol.Value = True Then .ColWidth(7) = 0: .ColWidth(8) = 0
                '--SOLO SOLES OCULTAR DOLARES
                If Me.OptSol.Value = True Then .ColWidth(5) = 0: .ColWidth(6) = 0
                UNIR_CELDAS Fg1, 0, 1, 0, 4, " "
                UNIR_CELDAS Fg1, 0, 5, 0, 6, "DOLARES"
                UNIR_CELDAS Fg1, 0, 7, 0, 8, "SOLES"
                UNIR_CELDAS Fg1, 0, 9, 0, 11, "TOTALES EN S/."
                OCULTAR_COL Fg1, 12, 19
            End If
        End With
    Else '--DETALLE
        With Fg1
            .TextMatrix(1, 1) = "N°.Reg.":    .ColWidth(1) = 820:   .ColAlignment(1) = flexAlignLeftBottom
            .TextMatrix(1, 2) = "T.D.":       .ColWidth(2) = 420:   .ColAlignment(2) = flexAlignCenterBottom
            .TextMatrix(1, 3) = "Cliente":    .ColWidth(3) = 0:     .ColAlignment(3) = flexAlignLeftBottom
            
            .TextMatrix(1, 4) = "Num. Documento":   .ColWidth(4) = 1400:    .ColAlignment(4) = flexAlignCenterBottom
            .TextMatrix(1, 5) = "Fec.Doc.":         .ColWidth(5) = 840:     .ColAlignment(5) = flexAlignCenterBottom
            .TextMatrix(1, 6) = "Fec.Venc.":        .ColWidth(6) = 840:     .ColAlignment(6) = flexAlignCenterBottom
            .TextMatrix(1, 7) = "Cond. Pago":       .ColWidth(7) = 950:     .ColAlignment(7) = flexAlignRightBottom
            .TextMatrix(1, 8) = "Dias Atra..":      .ColAlignment(8) = flexAlignRightBottom
            If Me.OptPag.Value = True Then
                .ColWidth(8) = 0
            Else
                .ColWidth(8) = 800
            End If
            .TextMatrix(1, 9) = "M":   .ColWidth(9) = 450:      .ColAlignment(9) = flexAlignLeftBottom
            .TextMatrix(1, 10) = "T.C.": .ColWidth(10) = 550:   .ColAlignment(10) = flexAlignRightBottom
            
            If Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0 Then '--SOLO PRODUCTO
                .FrozenCols = 6
                .TextMatrix(1, 11) = "Producto":    .ColWidth(11) = 1200: .ColAlignment(11) = flexAlignLeftBottom
                .TextMatrix(1, 12) = "Imp. $":      .ColWidth(12) = 1000: .ColAlignment(12) = flexAlignRightBottom
                .TextMatrix(1, 13) = "Imp. S/.":    .ColWidth(13) = 1000: .ColAlignment(13) = flexAlignRightBottom
                .TextMatrix(1, 14) = "Total S/.":   .ColWidth(14) = 1000: .ColAlignment(14) = flexAlignRightBottom
                '--SOLO DOLARES OCULTAR SOLES
                If Me.OptDol.Value = True Then
                    .ColWidth(13) = 0:
                End If
                '--SOLO SOLES OCULTAR DOLARES
                If Me.OptSol.Value = True Then
                    .ColWidth(12) = 0
                End If
                '--SOLO PAGADO OCULTAR SALDOS
    '            If Me.OptPag.Value = True Then .ColWidth(11) = 0: .ColWidth(13) = 0: .ColWidth(16) = 0
                UNIR_CELDAS Fg1, 0, 1, 0, 11, " "
                UNIR_CELDAS Fg1, 0, 12, 0, 14, "TOTALES"
                OCULTAR_COL Fg1, 15, 19
            ElseIf Me.ChkMostrarItem.Value = 1 Then '--ITEM
                .FrozenCols = 6
                .TextMatrix(1, 11) = "Producto":    .ColWidth(11) = 900:    .ColAlignment(11) = flexAlignLeftBottom
                .TextMatrix(1, 12) = "Item":        .ColWidth(12) = 2800:   .ColAlignment(12) = flexAlignLeftBottom
                
                .TextMatrix(1, 13) = "U.M.":        .ColWidth(13) = 500:    .ColAlignment(13) = flexAlignLeftBottom
                .TextMatrix(1, 14) = "Cant.":       .ColWidth(14) = 700:    .ColAlignment(14) = flexAlignRightBottom
                .TextMatrix(1, 15) = "P/U":         .ColWidth(15) = 500:    .ColAlignment(15) = flexAlignRightBottom
                .TextMatrix(1, 16) = "Imp.Total":   .ColWidth(16) = 900:    .ColAlignment(16) = flexAlignRightBottom
                .TextMatrix(1, 17) = "P/U":         .ColWidth(17) = 500:    .ColAlignment(17) = flexAlignRightBottom
                .TextMatrix(1, 18) = "Imp.Total":   .ColWidth(18) = 900:    .ColAlignment(18) = flexAlignRightBottom
                .TextMatrix(1, 19) = "Total S/.":   .ColWidth(19) = 0:      .ColAlignment(19) = flexAlignRightBottom
                UNIR_CELDAS Fg1, 0, 1, 0, 14, " "
                UNIR_CELDAS Fg1, 0, 15, 0, 16, "DOLARES"
                UNIR_CELDAS Fg1, 0, 17, 0, 18, "SOLES"
                UNIR_CELDAS Fg1, 0, 19, 0, 19, "TOTAL"
            Else
                .FrozenCols = 6
                .TextMatrix(1, 11) = "Imp.":    .ColWidth(11) = 900: .ColAlignment(11) = flexAlignRightBottom
                .TextMatrix(1, 12) = "Saldo":   .ColWidth(12) = 900: .ColAlignment(12) = flexAlignRightBottom
                .TextMatrix(1, 13) = "Imp.":    .ColWidth(13) = 900: .ColAlignment(13) = flexAlignRightBottom
                .TextMatrix(1, 14) = "Saldo":   .ColWidth(14) = 900: .ColAlignment(14) = flexAlignRightBottom
                .TextMatrix(1, 15) = "Total":   .ColWidth(15) = 1100: .ColAlignment(15) = flexAlignRightBottom
                .TextMatrix(1, 16) = "Abono":   .ColWidth(16) = 1100: .ColAlignment(16) = flexAlignRightBottom
                .TextMatrix(1, 17) = "Saldo":   .ColWidth(17) = 1200: .ColAlignment(17) = flexAlignRightBottom
                '--SOLO DOLARES OCULTAR SOLES
                If Me.OptDol.Value = True Then
                    .ColWidth(13) = 0: .ColWidth(14) = 0
                    .ColWidth(11) = 1000: .ColWidth(12) = 1000
                    .ColWidth(15) = 1000: .ColWidth(16) = 1000: .ColWidth(17) = 1250
                End If
                '--SOLO SOLES OCULTAR DOLARES
                If Me.OptSol.Value = True Then
                    .ColWidth(11) = 0: .ColWidth(12) = 0
                    .ColWidth(13) = 1000: .ColWidth(14) = 1000
                    .ColWidth(15) = 1000: .ColWidth(16) = 1000: .ColWidth(17) = 1000
                End If
                '--SOLO PAGADO OCULTAR SALDOS
                If Me.OptPag.Value = True Then .ColWidth(12) = 0: .ColWidth(14) = 0: .ColWidth(17) = 0
    
                UNIR_CELDAS Fg1, 0, 1, 0, 10, " "
                UNIR_CELDAS Fg1, 0, 11, 0, 12, "DOLARES"
                UNIR_CELDAS Fg1, 0, 13, 0, 14, "SOLES"
                UNIR_CELDAS Fg1, 0, 15, 0, 17, "TOTAL EN S/."
                OCULTAR_COL Fg1, 16, 19
            End If
        End With
    End If
End Sub




Private Sub PosicionarProgBar()
    '--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
    '    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    FraProgreso.Visible = True
End Sub


Private Function VERIFICAR_PONER_PRECIO_PROMEDIO() As Boolean
    '--VERIFICAR SI INSERTARA EL PRECIO PROMEDIO
    '--SOLO ESTA ACTIVO CUANDO SELECCIONE UN ITEM, LA SELECCION DE CLIENTE
    '--SI INSERTA SERA EN LA SIGUIENTE FILA DE LOS TOTALES
    Dim k, M_CANTIDAD_REGI As Integer
    
    '--DEL ITEM: M_CANTIDAD_REGI_CLI = 0
    If Me.OptResum.Value = True Then Exit Function
    If Me.ChkMostrarItem.Value = 0 Then GoTo Salir_FUNC
    With Fg2
        For k = 0 To .Rows - 1
            If Me.ChkMostrarItem.Value = 0 Then GoTo Salir_FUNC '--SALIR SI NO SELECCIONA MOSTRAR ITEM
            If k + 1 = .Rows Then Exit For
            'M_CANTIDAD_REGI = M_CANTIDAD_REGI + 1
            If CStr(.TextMatrix(k + 1, 3)) <> "" Then M_CANTIDAD_REGI = M_CANTIDAD_REGI + 1
        Next k
    End With
    
    If M_CANTIDAD_REGI = 1 Then VERIFICAR_PONER_PRECIO_PROMEDIO = True
    Exit Function
Salir_FUNC:
End Function

Private Function Verificar_Poner_Datos_Grls() As Boolean
    '--VERIFICAR SI INSERTARA LOS DATOS GENERALES TANTO PARA LOS MONTOS Y PRECIO PROMEDIO
    '--NO INSERTARA CUANDO SELECCIONA UN CLIENTE
    Dim k, M_CANTIDAD_REGI_CLI As Integer
    '--DEL ITEM
    M_CANTIDAD_REGI_CLI = 0
    With Fg3
        For k = 0 To .Rows - 1
            If k + 1 = .Rows Then Exit For
            If CStr(.TextMatrix(k + 1, 1)) <> "" Then M_CANTIDAD_REGI_CLI = M_CANTIDAD_REGI_CLI + 1
        Next k
    End With
    '---
    If M_CANTIDAD_REGI_CLI = 1 Then Exit Function
    Verificar_Poner_Datos_Grls = True
End Function


Private Function CALCULAR_PRECIO_PROMEDIO(Band_Total_gral As Boolean, M_POS As Integer) As String
    '--M_POS = 4: PU DOL
    '--M_POS = 5: PU SOL
    '--M_POS = 6 CANTIDAD DE REGISTROS
    If (Arr_Totales(M_POS, 0) = 0) And (Arr_Totales_grls(M_POS, 0) = 0) Then
        CALCULAR_PRECIO_PROMEDIO = ""
        Exit Function
    End If
    If Band_Total_gral = False Then
        CALCULAR_PRECIO_PROMEDIO = Format(Val(Arr_Totales(M_POS, 0)) / Val(Arr_Totales(6, 0)), FORMAT_MONTO)
    Else
        CALCULAR_PRECIO_PROMEDIO = Format(Val(Arr_Totales_grls(M_POS, 0)) / Val(Arr_Totales_grls(6, 0)), FORMAT_MONTO)
    End If
    
End Function

'--------
Private Sub pExportarExcel()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO + " ", "", T_RPT_PERIODO, "Ventas"
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub



'************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 5 Then pExportarExcel
    If Button.Index = 6 Then pImprimir
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub

'************************************************

