VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPlaneProduccion 
   Caption         =   "Produccion - Planeacion de Produccion"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frm 
      Caption         =   "[ Mostrar ]"
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
      Height          =   1035
      Index           =   1
      Left            =   7950
      TabIndex        =   16
      Top             =   360
      Width           =   3390
      Begin VB.CheckBox ChkMostrar 
         Caption         =   "P. de Produccion"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   200
         Width           =   1725
      End
      Begin VB.CheckBox ChkMostrar 
         Caption         =   "Stock"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   19
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox ChkMostrar 
         Caption         =   "Programado"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   450
         Width           =   1275
      End
      Begin VB.CheckBox ChkMostrar 
         Caption         =   "P. de Pedidos"
         Height          =   225
         Index           =   3
         Left            =   1890
         TabIndex        =   17
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Selec. Mes ]"
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
      Height          =   1035
      Left            =   5940
      TabIndex        =   14
      Top             =   360
      Width           =   1995
      Begin VB.ListBox LbMes 
         Height          =   735
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   230
         Width           =   1860
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   4320
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   7350
      Visible         =   0   'False
      Width           =   7530
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   7260
         Picture         =   "FrmPlaneProduccion.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   9
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   3330
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   840
         Width           =   7320
         _cx             =   12912
         _cy             =   5874
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
         Rows            =   3
         Cols            =   7
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmPlaneProduccion.frx":02EC
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
         Editable        =   2
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
      Begin VB.Label LblItem 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblItem"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   765
         TabIndex        =   13
         Top             =   360
         Width           =   6420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   195
         Index           =   18
         Left            =   30
         TabIndex        =   12
         Top             =   375
         Width           =   645
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de Ítem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   28
         Left            =   105
         TabIndex        =   11
         Top             =   60
         Width           =   1305
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   5
         X1              =   30
         X2              =   7500
         Y1              =   4290
         Y2              =   4290
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   5
         X1              =   7500
         X2              =   7500
         Y1              =   0
         Y2              =   4290
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   30
         Top             =   30
         Width           =   7440
      End
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   3240
      TabIndex        =   3
      Top             =   3780
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   4
         Top             =   465
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape Shape1 
         Height          =   765
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   5805
      End
      Begin VB.Label LblProg 
         AutoSize        =   -1  'True
         Caption         =   "Pedidos"
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
         Left            =   1710
         TabIndex        =   7
         Top             =   180
         Width           =   660
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
         Left            =   225
         TabIndex        =   6
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "No Interrumpir"
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
         Index           =   2
         Left            =   4170
         TabIndex        =   5
         Top             =   180
         Width           =   1530
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   1470
      Width           =   11385
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   5700
         Index           =   2
         Left            =   30
         TabIndex        =   2
         Top             =   60
         Width           =   11325
         _cx             =   19976
         _cy             =   10054
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
         BackColorSel    =   -2147483635
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   20
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmPlaneProduccion.frx":03B9
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9480
      Top             =   90
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
            Picture         =   "FrmPlaneProduccion.frx":05E0
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":0B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":0EB6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":103A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":148E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":15A6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":1AEA
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":202E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":2142
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":2256
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":26AA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":2816
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":2D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":30F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlaneProduccion.frx":340A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   609
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
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar "
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Grabar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Recetas del producto"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Productos "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar Excel"
            ImageIndex      =   13
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
   Begin VSFlex7Ctl.VSFlexGrid fg 
      Height          =   945
      Index           =   1
      Left            =   60
      TabIndex        =   21
      Top             =   450
      Width           =   5835
      _cx             =   10292
      _cy             =   1667
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
      FormatString    =   $"FrmPlaneProduccion.frx":3724
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
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu00 
         Caption         =   "Insertar Item"
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu menu01 
         Caption         =   "Modificar Item"
      End
      Begin VB.Menu separador1 
         Caption         =   "-"
      End
      Begin VB.Menu menu02 
         Caption         =   "Eliminar Item"
      End
   End
End
Attribute VB_Name = "FrmPlaneProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPLANEPRODUCCION.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : MUESTRA LOS PRODUCTOS Y LAS CANTIDADES A PRODUCIR SEGUN EL MES, ADEMAS HACE LA PROGRAMACION MENSUAL
'* DISEÑADO POR     : jOSE CHACON MANRIQUE
'*****************************************************************************************************
Option Explicit

Dim Agregando As Boolean        ' INFORMA QUE SE ESTA AGREGANDO UNA FILA AL CONTROL FLEXGRID
Dim SeEjecuto As Boolean        ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim QueHace As Integer          ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim IdMenuActivo As Integer     ' INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date             ' ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO

Dim RstValores As New ADODB.Recordset
Dim RstPedidos As New ADODB.Recordset
Dim RstPedidoDet As New ADODB.Recordset
Dim RstPedidoFiltrado As New ADODB.Recordset
Dim RstPlan As New ADODB.Recordset
Dim RstStock As New ADODB.Recordset
Dim RstPrograma As New ADODB.Recordset
Dim cSQL As String

Dim MES_(1 To 12) As Boolean
Dim CLIENTE_ As Boolean ' Variable que identifica si esta seleccionado Mostrar columna Cliente
Dim FECHA_ As Boolean ' Variable que identifica si esta seleccionado Mostrar columna Fecha de Entrega
Dim PLAN_ As Boolean ' Variable que identifica si esta seleccionado Mostrar columna Plan de Produccion
Dim STOCK_ As Boolean ' Variable que identifica si esta seleccionado Mostrar columna Stock
Dim PROGRAMADO_ As Boolean ' Variable que identifica si esta seleccionado Mostrar columna Programado

Dim CANTIDADTOTALHORAS_ As Double  ' Indica el numero de horas que se va a necesitar en producir todos los productos

Dim DETECTOR_ As CalendarHitTestInfo ' Variable que detecta eventos seleccionados en el calendario
Dim EVENTO_ As CalendarEvent ' Variable que contiene al evento seleccionado en el calendario
Dim ARRASTRANDO_ As Boolean ' Variable que identifica si se esta cambiando de lugar un evento arrastrandolo

Dim RSTCRONO_ As New ADODB.Recordset            ' Estado de un producto en el cronograma
Dim RSTCRONODET_ As New ADODB.Recordset
Dim RSTCRONOTAR_ As New ADODB.Recordset
Dim RSTCRONODETAUX_ As New ADODB.Recordset
Dim RSTCRONOTARAUX_ As New ADODB.Recordset

Dim INDICE_ As Integer ' Variable que identifica cual es el Index del Grid seleccionado
Dim CORRELATIVO_ As Double
' ----------------------DEFINICION DE COLUMNAS
Private Enum COLUMNA_
    COLUMNAITEM_ = 1
    COLUMNATIPO_
    COLUMNASTOCK_
    COLUMNAPLANPROD_
    COLUMNAPLANPED_
    COLUMNATOTALPROD_
    COLUMNAPROGAPROBADO_
    COLUMNAPROGCUMPLIDO_
    COLUMNAPROGCANCELADO_
    COLUMNAPROGRESTANTE_
    COLUMNAPROGTOT_
    COLUMNAPROGDESF_
    COLUMNAPRODUCIDO_
    COLUMNARESTO_
    COLUMNARESTOPORC_
    COLUMNAUNIXHORA_
    COLUMNAHRSTRAB_
    COLUMNAHRSPERS_
    COLUMNAIDITEM_
End Enum


Dim OrigFX As Long
Dim OrigFY As Long
Dim SALIR_ As Boolean

Private Sub generarConsulta(MESATRABAJAR_ As Integer, PEDIDO_ As Boolean, PEDIDODET_ As Boolean, _
                                                        Optional PLAN_ As Boolean = False, _
                                                        Optional STOCK_ As Boolean = False, _
                                                        Optional PROGRAMA_ As Boolean = False)
    Dim c_PRODUCTOS As String
    Dim c_CLIENTES As String
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim B As Integer
    Dim IDPRO_ As Double
    
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Double
    
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    ' Se encuentra el primer dia del mes actual
    PRIMERDIAMES_ = CDate("01/" & MESATRABAJAR_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    ULTIMODIAMES_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    ' Si es que haya habido algun cambio se regresan a su estado inicial
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    
    If PEDIDO_ Then
        c_PRODUCTOS = GENERAR_SQL_ID(fg(1), 1, " AND alm_inventario.id", "IN", True)
        
        cSQL = "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion AS desitem, alm_inventario.idunimed, mae_unidades.abrev AS unimed, Sum(ped_pedidodet.canpro) AS cantot " _
            + vbCr + "FROM ((ped_pedido LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]))>=CDate('" & PRIMERDIAMES_ & "') And (IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]))<=CDate('" & ULTIMODIAMES_ & "')) AND ((ped_pedido.anulado)=0)) " & c_PRODUCTOS _
            + vbCr + "GROUP BY alm_inventario.id, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev " _
            + vbCr + "UNION " _
            + vbCr + "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion AS desitem, alm_inventario.idunimed, mae_unidades.abrev AS unimed, Sum(ped_pedidodetent.canpro) AS cantot " _
            + vbCr + "FROM ((ped_pedido LEFT JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodetent.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((ped_pedido.idtipped)=2) AND ((ped_pedidodetent.fchent)>=CDate('" & PRIMERDIAMES_ & "') And (ped_pedidodetent.fchent)<=CDate('" & ULTIMODIAMES_ & "')) AND ((ped_pedido.anulado)=0)) " & c_PRODUCTOS _
            + vbCr + "GROUP BY alm_inventario.id, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        If RstPedidos.State = 0 Then DEFINIR_RST_TMP RstPedidos, xRs
        limpiarRST RstPedidos
        CARGAR_RST_TMP RstPedidos, xRs
    End If
    
    If PEDIDODET_ Then
    End If
    
    If PLAN_ Then
        Dim CANMESACT_ As Double
        Dim CANMESANT_ As Double
        Dim CANMESPOS_ As Double
        Dim MESPOSTERIOR_ As Double
        Dim MESANTERIOR_ As Double
        
        CANMESACT_ = 0
        CANMESANT_ = 0
        CANMESPOS_ = 0
        
        If MESACTUAL_ = 12 Then MESPOSTERIOR_ = 1 Else MESPOSTERIOR_ = MESACTUAL_ + 1
        If MESACTUAL_ = 1 Then MESANTERIOR_ = 12 Else MESANTERIOR_ = MESACTUAL_ - 1
             
        c_PRODUCTOS = GENERAR_SQL_ID(fg(1), 1, " AND ges_plaproddet.codpro", "IN", True)
        
        ' CARGAMOS LOS TERMINADOS
        cSQL = "SELECT ges_plaproddet.codpro AS iditem, alm_inventario.descripcion AS desitem, ges_plaproddet.cantidad, alm_inventario.idunimed, mae_unidades.abrev, 'T' AS tipo " _
            + vbCr + "FROM ((ges_plaprod LEFT JOIN ges_plaproddet ON ges_plaprod.id = ges_plaproddet.idpv) LEFT JOIN alm_inventario ON ges_plaproddet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((ges_plaprod.activo)=-1) AND ((ges_plaproddet.idmes)=" & MESACTUAL_ & ")) " & c_PRODUCTOS
    
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        If RstPlan.State = 0 Then DEFINIR_RST_TMP RstPlan, xRs
        limpiarRST RstPlan
        CARGAR_RST_TMP RstPlan, xRs
             
        c_PRODUCTOS = GENERAR_SQL_ID(fg(1), 1, " AND ges_plaproddet2.codpro", "IN", True)
        
        ' CARGAMOS LOS INTERMEDIOS
        cSQL = "SELECT ges_plaproddet2.codpro AS iditem, alm_inventario.descripcion AS desitem, ges_plaproddet2.cantidad, alm_inventario.idunimed, mae_unidades.abrev, 'I' AS tipo " _
            + vbCr + "FROM ((ges_plaprod LEFT JOIN ges_plaproddet2 ON ges_plaprod.id = ges_plaproddet2.idpv) LEFT JOIN alm_inventario ON ges_plaproddet2.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((ges_plaprod.activo)=-1) AND ((ges_plaproddet2.idmes)=" & MESACTUAL_ & ")) " & c_PRODUCTOS
            
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        CentrarFrm FraProgreso
        FraProgreso.Visible = True
        LblProg.Caption = "Plan de Producción"
        PgBar.Min = 0
        PgBar.Max = xRs.RecordCount
        PgBar.Value = 0
        
        xRs.MoveFirst
        While Not xRs.EOF
            FraProgreso.Refresh
            PgBar.Value = PgBar.Value + 1
            
            RstPlan.Filter = adFilterNone
            RstPlan.Filter = "iditem=" & NulosN(xRs("iditem"))
            If RstPlan.RecordCount = 0 Then
                RstPlan.AddNew
                RstPlan("iditem") = NulosN(xRs("iditem"))
                RstPlan("desitem") = NulosC(xRs("desitem"))
                RstPlan("cantidad") = NulosN(xRs("cantidad"))
                RstPlan("idunimed") = NulosN(xRs("idunimed"))
                RstPlan("abrev") = NulosC(xRs("abrev"))
                RstPlan("tipo") = NulosC(xRs("tipo"))
                RstPlan.Update
            Else
                RstPlan("cantidad") = RstPlan("cantidad") + NulosN(xRs("cantidad"))
                RstPlan("tipo") = "A"
                RstPlan.Update
            End If
            xRs.MoveNext
        Wend
        
        RstPlan.Filter = adFilterNone
        If RstPedidos.State = 0 Then Exit Sub
        If RstPedidos.RecordCount = 0 Then Exit Sub
        
        CentrarFrm FraProgreso
        FraProgreso.Visible = True
        LblProg.Caption = "Plan de Pedidos"
        PgBar.Min = 0
        PgBar.Max = RstPedidos.RecordCount
        PgBar.Value = 0
        
        RstPedidos.MoveFirst
        While Not RstPedidos.EOF
            FraProgreso.Refresh
            PgBar.Value = PgBar.Value + 1
            
            RstPlan.Filter = "iditem=" & NulosN(RstPedidos("iditem"))
            If RstPlan.RecordCount = 0 Then
                RstPlan.AddNew
                RstPlan("iditem") = NulosN(RstPedidos("iditem"))
                RstPlan("desitem") = NulosC(RstPedidos("desitem"))
                RstPlan("cantidad") = 0
                RstPlan("idunimed") = NulosN(RstPedidos("idunimed"))
                RstPlan("abrev") = NulosC(RstPedidos("unimed"))
                RstPlan("tipo") = "O"
                RstPlan.Update
            End If
            RstPedidos.MoveNext
        Wend
        FraProgreso.Visible = False
    End If
    
    If STOCK_ Then
    End If
    
    If PROGRAMA_ Then
    End If
End Sub

Private Function consultarProgramado(MESATRABAJAR_ As Integer, IDPRODUCTO_ As Double, _
                                            Optional TIPO_ As Integer = 0) As Double
    Dim xRs As New ADODB.Recordset
    Dim PROGRAMADO_ As Double
    Dim CONSULTA_ As String
    Dim FECHAINICIO_ As String
    Dim FECHAFIN_ As String
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Double
    
    ' TIPO_=0:TOTAL PEDIENTE, TIPO_=1:TOTAL APROBADO, TIPO_=2:TOTAL CANCELADO
    
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    ' Se encuentra el primer dia del mes actual
    FECHAINICIO_ = CDate("01/" & MESATRABAJAR_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    FECHAFIN_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    
    Set xRs = Nothing
    Select Case TIPO_
        Case 0
            CONSULTA_ = "WHERE (((pro_cronogramadet.fchpro)>=CDate('" & FECHAINICIO_ & "') " _
                                    & "AND (pro_cronogramadet.fchpro)<=CDate('" & FECHAFIN_ & "')) " _
                                    & "AND ((pro_cronogramadet.iditem)=" & IDPRODUCTO_ & ") " _
                                    & "AND ((pro_cronogramadet.estado) In (1)));"
        Case 1
            CONSULTA_ = "WHERE (((pro_cronogramadet.fchpro)>=CDate('" & FECHAINICIO_ & "') " _
                                    & "AND (pro_cronogramadet.fchpro)<=CDate('" & FECHAFIN_ & "')) " _
                                    & "AND ((pro_cronogramadet.iditem)=" & IDPRODUCTO_ & ") " _
                                    & "AND ((pro_cronogramadet.estado) In (2)));"
        Case 2
            CONSULTA_ = "WHERE (((pro_cronogramadet.fchpro)>=CDate('" & FECHAINICIO_ & "') " _
                                    & "AND (pro_cronogramadet.fchpro)<=CDate('" & FECHAFIN_ & "')) " _
                                    & "AND ((pro_cronogramadet.iditem)=" & IDPRODUCTO_ & ") " _
                                    & "AND ((pro_cronogramadet.estado) In (4)));"
    End Select
    
    cSQL = "SELECT Sum(pro_cronogramadet.cantidad) AS cantot " _
        + vbCr + "FROM pro_cronogramadet " _
        + vbCr + CONSULTA_
        
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then PROGRAMADO_ = 0: GoTo SALIR
    If xRs.RecordCount = 0 Then PROGRAMADO_ = 0: GoTo SALIR
    PROGRAMADO_ = NulosN(xRs("cantot"))
SALIR:
    consultarProgramado = PROGRAMADO_
End Function

Private Function consultarProducido(MESATRABAJAR_ As Integer, IDITEM_ As Double) As Double
    Dim xRs As New ADODB.Recordset
    Dim PRODUCIDO_ As Double
    Dim CONSULTA_ As String
    Dim FECHAINICIO_ As String
    Dim FECHAFIN_ As String
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Integer
    
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    ' Se encuentra el primer dia del mes actual
    FECHAINICIO_ = CDate("01/" & MESATRABAJAR_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    FECHAFIN_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    
    cSQL = "SELECT Sum(pro_producciondet.cantidad) AS cantot " _
        + vbCr + "FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        + vbCr + "WHERE (((pro_produccion.dia)>=CDate('" & FECHAINICIO_ & "') " _
                & "AND (pro_produccion.dia)<=CDate('" & FECHAFIN_ & "')) " _
                & "AND ((pro_producciondet.iditem)=" & IDITEM_ & ") AND ((pro_producciondet.estado)>1));"
        
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then PRODUCIDO_ = 0: GoTo SALIR
    If xRs.RecordCount = 0 Then PRODUCIDO_ = 0: GoTo SALIR
    PRODUCIDO_ = NulosN(xRs("cantot"))
SALIR:
    consultarProducido = PRODUCIDO_
End Function

Private Function consultarSaldo(MESATRABAJAR_ As Integer, IDITEM_ As Double) As Double
    Dim xRs As New ADODB.Recordset
    Dim SALDO_ As Double
    Dim FECHAINICIO_ As String
    Dim FECHAFIN_ As String
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Integer
    
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    ' Se encuentra el primer dia del mes actual
    FECHAINICIO_ = CDate("01/" & MESATRABAJAR_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    FECHAFIN_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    
    SALDO_ = SaldoActual(IDITEM_, "01/01/" & ANIOACTUAL_, CDate(FECHAINICIO_) - 1, xCon)
    
    consultarSaldo = SALDO_
End Function

Private Sub ChkMostrar_Click(Index As Integer)
    Select Case Index
        Case 0
            If ChkMostrar(3).Value = 0 Then
                ChkMostrar(0).Value = 1
            End If
        Case 3
            If ChkMostrar(0).Value = 0 Then
                ChkMostrar(3).Value = 1
            End If
            
    End Select
End Sub

Private Sub fg_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    If Index = 2 Then
        With fg(Index)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COLUMNAUNIXHORA_) = "TOTAL"
            .Select .Rows - 1, COLUMNAUNIXHORA_
            .CellForeColor = &HFF&
            .CellFontBold = True
            .TextMatrix(.Rows - 1, COLUMNAHRSTRAB_) = Format(CANTIDADTOTALHORAS_, FORMAT_CANTIDAD)
            If .Rows > .FixedRows Then .Select .FixedRows, 1
        End With
    End If
End Sub

Private Sub fg_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    If Index = 2 Then
        Dim TOTALHORAS_ As Double
        With fg(Index)
            CANTIDADTOTALHORAS_ = .TextMatrix(.Rows - 1, COLUMNAHRSTRAB_)
            .Rows = .Rows - 1
        End With
    End If
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim nSQLId As String
    
    If Index = 1 Then
        ReDim xCampos(2, 3) As String
        Set xRs = Nothing
        
        xCampos(0, 0) = "Nombre":        xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "codpro":          xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"

        nSQLId = GENERAR_SQL_ID(fg(Index), 1, " AND alm_inventario.id", "NOT IN", True)
        
        cSQL = "SELECT alm_inventario.id AS idpro, alm_inventario.descripcion, alm_inventario.codpro " _
            + vbCr + "FROM alm_inventario " _
            + vbCr + "WHERE (((alm_inventario.activo)=-1) AND ((alm_inventario.tippro) In (1,3)) AND ((alm_inventario.idcuentaven)<>0)) " & nSQLId
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando Items", "descripcion", "descripcion", Principio
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("idpro"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("descripcion"))
        fg(Index).Rows = fg(Index).Rows + 1
        fg(Index).Select fg(Index).Rows - 1, 2
        fg(Index).TopRow = fg(Index).Rows - 1
    End If
End Sub

Private Sub fg_DblClick(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim FECHAINICIO_ As String
    Dim FECHAFIN_ As String
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Double
    Dim IDITEM_ As Integer
    Dim NUMEROFILAS_ As Integer
    Dim A As Integer
        
    If Not frm(0).Visible Then CentrarFrm frm(0)
    frm(0).Visible = True
    fg(0).Rows = fg(0).FixedRows
    fg(0).MergeCells = flexMergeFixedOnly
            
    For A = 1 To LbMes.ListCount - 1
        LbMes.ListIndex = A
        MESACTUAL_ = A + 1
        If LbMes.Selected(A) = False Then GoTo SIGUIENTE
        A = LbMes.ListCount - 1
SIGUIENTE:
    Next A
        
        
    ANIOACTUAL_ = AnoTra
    ' Se encuentra el primer dia del mes actual
    FECHAINICIO_ = CDate("01/" & MESACTUAL_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    FECHAFIN_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    
    IDITEM_ = NulosN(fg(2).TextMatrix(fg(2).Row, COLUMNAIDITEM_))
    lblItem.Caption = NulosC(fg(2).TextMatrix(fg(2).Row, COLUMNAITEM_))
    
    ' CONSULTA PROGRAMADO
    cSQL = "SELECT pro_cronogramadet.fchpro AS fecha, pro_cronogramadet.cantidad AS canprog, mae_estados.descripcion AS desestado " _
        + vbCr + "FROM pro_cronogramadet INNER JOIN mae_estados ON pro_cronogramadet.estado = mae_estados.id " _
        + vbCr + "WHERE (((pro_cronogramadet.fchpro)>=CDate('" & FECHAINICIO_ & "') And (pro_cronogramadet.fchpro)<=CDate('" & FECHAFIN_ & "')) AND ((pro_cronogramadet.iditem)=" & IDITEM_ & ")) " _
        + vbCr + "ORDER BY pro_cronogramadet.fchpro;"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    NUMEROFILAS_ = xRs.RecordCount
    If xRs.RecordCount = 0 Then GoTo LLENARPRODUCIDO
    xRs.MoveFirst
    While Not xRs.EOF
        fg(0).Rows = fg(0).Rows + 1
        fg(0).TextMatrix(fg(0).Rows - 1, 1) = Format(xRs("fecha"), FORMAT_DATE)
        fg(0).TextMatrix(fg(0).Rows - 1, 2) = Format(NulosN(xRs("canprog")), FORMAT_CANTIDAD)
        fg(0).TextMatrix(fg(0).Rows - 1, 3) = UCase(NulosC(xRs("desestado")))
        xRs.MoveNext
    Wend
    
LLENARPRODUCIDO:
    ' CANTIDAD PRODUCIDA
    cSQL = "SELECT pro_produccion.dia AS fecha, pro_producciondet.cantidad AS canprod, mae_estados.descripcion AS desestado " _
        + vbCr + "FROM (pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro) INNER JOIN mae_estados ON pro_producciondet.estado = mae_estados.id " _
        + vbCr + "WHERE (((pro_produccion.dia)>=CDate('" & FECHAINICIO_ & "') And (pro_produccion.dia)<=CDate('" & FECHAFIN_ & "')) AND ((pro_producciondet.iditem)=" & IDITEM_ & ")) " _
        + vbCr + "ORDER BY pro_produccion.dia;"

    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If NUMEROFILAS_ < xRs.RecordCount Then fg(0).Rows = xRs.RecordCount + 2
    If xRs.RecordCount = 0 Then GoTo SALIR_
    
    xRs.MoveFirst
    For A = 2 To fg(0).Rows - 1
        fg(0).TextMatrix(A, 4) = Format(xRs("fecha"), FORMAT_DATE)
        fg(0).TextMatrix(A, 5) = Format(NulosN(xRs("canprod")), FORMAT_CANTIDAD)
        fg(0).TextMatrix(A, 6) = UCase(NulosC(xRs("desestado")))
        xRs.MoveNext
        If xRs.EOF Then GoTo SALIR_
    Next A
SALIR_:
    fg(0).Rows = fg(0).Rows + 1
    fg(0).TextMatrix(fg(0).Rows - 1, 1) = "TOTAL"
    fg(0).TextMatrix(fg(0).Rows - 1, 2) = Format(GRID_SUMAR_COL(fg(0), 2), FORMAT_CANTIDAD)
    fg(0).TextMatrix(fg(0).Rows - 1, 4) = "TOTAL"
    fg(0).TextMatrix(fg(0).Rows - 1, 5) = Format(GRID_SUMAR_COL(fg(0), 5), FORMAT_CANTIDAD)
    
End Sub

Private Sub Fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    INDICE_ = Index
    If KeyCode = vbKeyInsert Then ' Agregar
        menu00_Click
    End If
    If KeyCode = vbKeyDelete Then ' Eliminar
        menu02_Click
    End If
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    Select Case Index
        Case 0, 1
            menu01.Enabled = False
            INDICE_ = Index
            PopupMenu menu
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub fg_RowColChange(Index As Integer)
    If frm(0).Visible Then fg_DblClick Index
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE LE FORMULARIO
    If SeEjecuto = False Then
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    iniciarCampos
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 8100 Then Me.Height = 8100
        
    ' Se dimensiona el contenido
    frm(2).Width = Me.Width - 20
    frm(2).Height = Me.Height - 1580
    
    fg(2).Width = frm(2).Width - 150
    fg(2).Height = frm(2).Height - 450
End Sub

Private Sub iniciarCampos()
    fg(1).AllowUserResizing = flexResizeColumns
    fg(1).AutoSearch = flexSearchFromTop
    fg(1).ExplorerBar = flexExSortShow
    fg(1).SelectionMode = flexSelectionByRow
    fg(1).ForeColorSel = &H80000005
    fg(1).BackColorSel = &H80&
    fg(1).Editable = flexEDKbdMouse
    GRID_COMBOLIST fg(1), 2
    fg(1).ColWidth(1) = 0
    fg(1).Rows = 2
    
    fg(2).AllowUserResizing = flexResizeColumns
    fg(2).AutoSearch = flexSearchFromTop
    fg(2).ExplorerBar = flexExSortShow
    fg(2).SelectionMode = flexSelectionByRow
    fg(2).ForeColorSel = &H80000005
    fg(2).BackColorSel = &H80&
    fg(2).ColWidth(0) = 0
    
    ChkMostrar(0).Value = 1
    ChkMostrar(1).Value = 1
    ChkMostrar(2).Value = 1
    ChkMostrar(3).Value = 1
    
    fg(2).Rows = 2
    fg(2).FixedRows = 2
    ' Se configura el Grid
    configurarGrid
    ' Se llena los meses
    Llenar_Mes LbMes
    ' Se selecciona el mes actual
    LbMes.Selected(Month(Date) - 1) = True
    
    'fg(3).Editable = flexEDKbdMouse
    
    CORRELATIVO_ = -666
    ARRASTRANDO_ = False
End Sub

Private Sub configurarGrid()
    fg(2).ColWidth(COLUMNAIDITEM_) = 0
    fg(2).ColWidth(COLUMNAHRSPERS_) = 0
    
    ' Plan de Produccion
    If ChkMostrar(0).Value = 0 Then
        fg(2).ColWidth(COLUMNAPLANPROD_) = 0
    Else
        fg(2).ColWidth(COLUMNAPLANPROD_) = 900
    End If
    
    ' Programado
    If ChkMostrar(1).Value = 0 Then
        fg(2).ColWidth(COLUMNAPROGAPROBADO_) = 0
        fg(2).ColWidth(COLUMNAPROGCUMPLIDO_) = 0
        fg(2).ColWidth(COLUMNAPROGCANCELADO_) = 0
        fg(2).ColWidth(COLUMNAPROGRESTANTE_) = 0
        fg(2).ColWidth(COLUMNAPROGTOT_) = 0
        fg(2).ColWidth(COLUMNAPROGDESF_) = 0
    Else
        fg(2).ColWidth(COLUMNAPROGAPROBADO_) = 900
        fg(2).ColWidth(COLUMNAPROGCUMPLIDO_) = 900
        fg(2).ColWidth(COLUMNAPROGCANCELADO_) = 900
        fg(2).ColWidth(COLUMNAPROGRESTANTE_) = 900
        fg(2).ColWidth(COLUMNAPROGTOT_) = 900
        fg(2).ColWidth(COLUMNAPROGDESF_) = 900
    End If
    
    ' Stock
    If ChkMostrar(2).Value = 0 Then
        fg(2).ColWidth(COLUMNASTOCK_) = 0
    Else
        fg(2).ColWidth(COLUMNASTOCK_) = 900
    End If
    
    ' Plan de Pedidos
    If ChkMostrar(3).Value = 0 Then
        fg(2).ColWidth(COLUMNAPLANPED_) = 0
    Else
        fg(2).ColWidth(COLUMNAPLANPED_) = 900
    End If
    
    
    GRID_COMBINAR fg(2), 0, COLUMNAITEM_, 1, COLUMNAITEM_, "Ítem", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 0, COLUMNATIPO_, 1, COLUMNATIPO_, "Tipo", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 0, COLUMNASTOCK_, 0, COLUMNATOTALPROD_, "Producción", flexAlignCenterCenter, True, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNASTOCK_, 1, COLUMNASTOCK_, "Stock Ini.", flexAlignCenterCenter, True, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNAPLANPROD_, 1, COLUMNAPLANPROD_, "Plan", flexAlignCenterCenter, True, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNAPLANPED_, 1, COLUMNAPLANPED_, "Pedido", flexAlignCenterCenter, True, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNATOTALPROD_, 1, COLUMNATOTALPROD_, "Total", flexAlignCenterCenter, True, , , &H8000000F, False
    
    GRID_COMBINAR fg(2), 0, COLUMNAPROGAPROBADO_, 0, COLUMNAPROGDESF_, "Programado", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNAPROGAPROBADO_, 1, COLUMNAPROGAPROBADO_, "Aprobado", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNAPROGCUMPLIDO_, 1, COLUMNAPROGCUMPLIDO_, "Cumplido", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNAPROGCANCELADO_, 1, COLUMNAPROGCANCELADO_, "Anulado", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNAPROGRESTANTE_, 1, COLUMNAPROGRESTANTE_, "Pendiente", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNAPROGTOT_, 1, COLUMNAPROGTOT_, "Total", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNAPROGDESF_, 1, COLUMNAPROGDESF_, "Desface", flexAlignCenterCenter, False, , , &H8000000F, False
    
    GRID_COMBINAR fg(2), 0, COLUMNAPRODUCIDO_, 1, COLUMNAPRODUCIDO_, "Producido", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 0, COLUMNARESTO_, 0, COLUMNARESTOPORC_, "Resto a Producir", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNARESTO_, 1, COLUMNARESTO_, "Cantidad", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 1, COLUMNARESTOPORC_, 1, COLUMNARESTOPORC_, "%", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 0, COLUMNAUNIXHORA_, 1, COLUMNAUNIXHORA_, "Unid.xHrs.", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 0, COLUMNAHRSTRAB_, 1, COLUMNAHRSTRAB_, "Hrs. Trab.", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(2), 0, COLUMNAHRSPERS_, 1, COLUMNAHRSPERS_, "Hrs. Pers.", flexAlignCenterCenter, False, , , &H8000000F, False
        
    GRID_COMBINAR fg(0), 0, 1, 0, 3, "Programado", flexAlignCenterCenter, True, , , &H8000000F, False
    GRID_COMBINAR fg(0), 1, 1, 1, 1, "Fecha", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(0), 1, 2, 1, 2, "Cantidad", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(0), 1, 3, 1, 3, "Estado", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(0), 0, 4, 0, 6, "Producido", flexAlignCenterCenter, True, , , &H8000000F, False
    GRID_COMBINAR fg(0), 1, 4, 1, 4, "Fecha", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(0), 1, 5, 1, 5, "Cantidad", flexAlignCenterCenter, False, , , &H8000000F, False
    GRID_COMBINAR fg(0), 1, 6, 1, 6, "Estado", flexAlignCenterCenter, False, , , &H8000000F, False
    
    fg(2).MergeCells = flexMergeFixedOnly
    If fg(2).Rows = fg(2).FixedRows Then Exit Sub
    
    ' Total a Producir
    fg(2).Select fg(2).FixedRows, COLUMNATOTALPROD_, fg(2).Rows - 1, COLUMNATOTALPROD_
    fg(2).FillStyle = flexFillRepeat
    fg(2).CellBackColor = &HDDFFFF
    ' Total Programado
    fg(2).Select fg(2).FixedRows, COLUMNAPROGTOT_, fg(2).Rows - 1, COLUMNAPROGTOT_
    fg(2).FillStyle = flexFillRepeat
    fg(2).CellBackColor = &HDDFFFF
    ' Desface
    fg(2).Select fg(2).FixedRows, COLUMNAPROGDESF_, fg(2).Rows - 1, COLUMNAPROGDESF_
    fg(2).FillStyle = flexFillRepeat
    fg(2).CellBackColor = &HE8E8FF
    ' Producido
    fg(2).Select fg(2).FixedRows, COLUMNAPRODUCIDO_, fg(2).Rows - 1, COLUMNAPRODUCIDO_
    fg(2).FillStyle = flexFillRepeat
    fg(2).CellBackColor = &HDDFFFF
    ' Resto
    fg(2).Select fg(2).FixedRows, COLUMNARESTO_, fg(2).Rows - 1, COLUMNARESTO_
    fg(2).FillStyle = flexFillRepeat
    fg(2).CellBackColor = &HE8E8FF
    fg(2).Select fg(2).FixedRows, 1, fg(2).FixedRows, 1
End Sub

Private Sub aplicarFiltrado()
    Dim A As Integer
    Dim INDICE_ As Integer
    Dim INDICETOPE_ As Integer
    Dim MES_ As Integer
        
    ' Se encuentran las caracteristicas del indice seleccionado
    INDICE_ = LbMes.ListIndex
    INDICETOPE_ = LbMes.TopIndex
        
    For A = 1 To LbMes.ListCount - 1
        LbMes.ListIndex = A
        MES_ = A + 1
        If LbMes.Selected(A) = False Then GoTo SIGUIENTE
        ' Se generan los Rst
        LblProg.Caption = "Buscando Pedidos"
        generarConsulta CDbl(MES_), True, False ' Pedidos
        
        LblProg.Caption = "Procesando Plan"
        generarConsulta MES_, False, False, True ' Plan
        
        LblProg.Caption = "Cargando Productos"
        llenarDatos MES_
        
        A = LbMes.ListCount - 1
SIGUIENTE:
    Next A
    
    LbMes.TopIndex = INDICETOPE_
    LbMes.ListIndex = INDICE_
End Sub

Private Sub llenarDatos(MESATRABAJAR_ As Integer)
    Dim xRs As New ADODB.Recordset
    Dim IDITEM_ As Double
    Dim VALOR_ As Double ' unid/hora de cada producto
    Dim TOTALHORAS_ As Double ' Tiempo en horas de cada producto
    
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Double
    Dim MESACTUAL_ As Double
    
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    ' Se encuentra el primer dia del mes actual
    PRIMERDIAMES_ = CDate("01/" & MESATRABAJAR_ & "/" & ANIOACTUAL_ & "")
    ' Se encuentra el ultimo dia del mes actual
    If MESACTUAL_ = 12 Then MESACTUAL_ = 0: ANIOACTUAL_ = ANIOACTUAL_ + 1
    ULTIMODIAMES_ = CDate("01/" & MESACTUAL_ + 1 & "/" & ANIOACTUAL_ & "") - 1
    ' Si es que haya habido algun cambio se regresan a su estado inicial
    ANIOACTUAL_ = AnoTra
    MESACTUAL_ = MESATRABAJAR_
    
    fg(2).Rows = fg(2).FixedRows
    If RstPlan.State = 0 Then Exit Sub
    
    VALOR_ = 0
    TOTALHORAS_ = 0
    
    With fg(2)
        RstPlan.Filter = adFilterNone
        If RstPlan.RecordCount = 0 Then Exit Sub
        
        CentrarFrm FraProgreso
        FraProgreso.Visible = True
        LblProg.Caption = "Llenando datos"
        PgBar.Min = 0
        PgBar.Max = RstPlan.RecordCount
        PgBar.Value = 0
        
        RstPlan.MoveFirst
        While Not RstPlan.EOF
            .Rows = .Rows + 1
            FraProgreso.Refresh
            PgBar.Value = PgBar.Value + 1
            
            IDITEM_ = RstPlan("iditem")
            .TextMatrix(.Rows - 1, COLUMNAIDITEM_) = IDITEM_
            .TextMatrix(.Rows - 1, COLUMNAITEM_) = NulosC(RstPlan("desitem"))
            .TextMatrix(.Rows - 1, COLUMNATIPO_) = NulosC(RstPlan("tipo"))
            
            ' Stock
            If ChkMostrar(2).Value = 1 Then
                .TextMatrix(.Rows - 1, COLUMNASTOCK_) = consultarSaldo(MESATRABAJAR_, IDITEM_)
            Else
                .TextMatrix(.Rows - 1, COLUMNASTOCK_) = 0
            End If
            .TextMatrix(.Rows - 1, COLUMNASTOCK_) = Format(.TextMatrix(.Rows - 1, COLUMNASTOCK_), FORMAT_CANTIDAD)
            pintarGrid fg(2), COLUMNASTOCK_, &H0&, &HFF&
            
            ' Produccion
            '**************************************************************************
            ' Plan de Produccion
            If ChkMostrar(0).Value = 1 Then
                .TextMatrix(.Rows - 1, COLUMNAPLANPROD_) = NulosN(RstPlan("cantidad"))
            Else
                .TextMatrix(.Rows - 1, COLUMNAPLANPROD_) = 0
            End If
            .TextMatrix(.Rows - 1, COLUMNAPLANPROD_) = Format(.TextMatrix(.Rows - 1, COLUMNAPLANPROD_), FORMAT_CANTIDAD)
            
            ' Plan de Pedido
            If ChkMostrar(3).Value = 1 Then
                RstPedidos.Filter = adFilterNone
                RstPedidos.Filter = "iditem=" & IDITEM_
                
                If RstPedidos.RecordCount = 0 Then
                    .TextMatrix(.Rows - 1, COLUMNAPLANPED_) = 0
                Else
                    .TextMatrix(.Rows - 1, COLUMNAPLANPED_) = NulosN(RstPedidos("cantot"))
                End If
            Else
                .TextMatrix(.Rows - 1, COLUMNAPLANPED_) = 0
            End If
            .TextMatrix(.Rows - 1, COLUMNAPLANPED_) = Format(.TextMatrix(.Rows - 1, COLUMNAPLANPED_), FORMAT_CANTIDAD)
            ' Total
            If (NulosN(.TextMatrix(.Rows - 1, COLUMNAPLANPROD_)) >= NulosN(.TextMatrix(.Rows - 1, COLUMNAPLANPED_))) Then
                .TextMatrix(.Rows - 1, COLUMNATOTALPROD_) = NulosN(.TextMatrix(.Rows - 1, COLUMNAPLANPROD_))
            Else
                .TextMatrix(.Rows - 1, COLUMNATOTALPROD_) = NulosN(.TextMatrix(.Rows - 1, COLUMNAPLANPED_))
            End If
            
            If NulosN(.TextMatrix(.Rows - 1, COLUMNATOTALPROD_)) = 0 And NulosN(.TextMatrix(.Rows - 1, COLUMNASTOCK_)) >= 0 Then
                .TextMatrix(.Rows - 1, COLUMNATOTALPROD_) = 0
            Else
                .TextMatrix(.Rows - 1, COLUMNATOTALPROD_) = NulosN(.TextMatrix(.Rows - 1, COLUMNATOTALPROD_)) - NulosN(.TextMatrix(.Rows - 1, COLUMNASTOCK_))
            End If
            
            .TextMatrix(.Rows - 1, COLUMNATOTALPROD_) = Format(.TextMatrix(.Rows - 1, COLUMNATOTALPROD_), FORMAT_CANTIDAD)
            '**************************************************************************
            
            ' Producido
            .TextMatrix(.Rows - 1, COLUMNAPRODUCIDO_) = consultarProducido(MESATRABAJAR_, IDITEM_)
            .TextMatrix(.Rows - 1, COLUMNAPRODUCIDO_) = Format(.TextMatrix(.Rows - 1, COLUMNAPRODUCIDO_), FORMAT_CANTIDAD)
            
            ' Programado
            '**************************************************************************
            If ChkMostrar(1).Value = 1 Then
                ' -----------APROBADO
                .TextMatrix(.Rows - 1, COLUMNAPROGAPROBADO_) = consultarProgramado(MESATRABAJAR_, IDITEM_, 1)
                .TextMatrix(.Rows - 1, COLUMNAPROGAPROBADO_) = Format(.TextMatrix(.Rows - 1, COLUMNAPROGAPROBADO_), FORMAT_CANTIDAD)
                ' -----------CUMPLIDO
                .TextMatrix(.Rows - 1, COLUMNAPROGCUMPLIDO_) = NulosN(.TextMatrix(.Rows - 1, COLUMNAPRODUCIDO_))
                .TextMatrix(.Rows - 1, COLUMNAPROGCUMPLIDO_) = Format(.TextMatrix(.Rows - 1, COLUMNAPROGCUMPLIDO_), FORMAT_CANTIDAD)
                ' -----------CANCELADO
                .TextMatrix(.Rows - 1, COLUMNAPROGCANCELADO_) = consultarProgramado(MESATRABAJAR_, IDITEM_, 2)
                .TextMatrix(.Rows - 1, COLUMNAPROGCANCELADO_) = Format(.TextMatrix(.Rows - 1, COLUMNAPROGCANCELADO_), FORMAT_CANTIDAD)
                ' -----------PEDIENTE
                .TextMatrix(.Rows - 1, COLUMNAPROGRESTANTE_) = consultarProgramado(MESATRABAJAR_, IDITEM_)
                .TextMatrix(.Rows - 1, COLUMNAPROGRESTANTE_) = Format(.TextMatrix(.Rows - 1, COLUMNAPROGRESTANTE_), FORMAT_CANTIDAD)
                ' -----------TOTAL
                .TextMatrix(.Rows - 1, COLUMNAPROGTOT_) = NulosN(.TextMatrix(.Rows - 1, COLUMNAPROGAPROBADO_)) + NulosN(.TextMatrix(.Rows - 1, COLUMNAPROGCANCELADO_)) + NulosN(.TextMatrix(.Rows - 1, COLUMNAPROGRESTANTE_))
                .TextMatrix(.Rows - 1, COLUMNAPROGTOT_) = Format(.TextMatrix(.Rows - 1, COLUMNAPROGTOT_), FORMAT_CANTIDAD)
                ' -----------DESFACE
                .TextMatrix(.Rows - 1, COLUMNAPROGDESF_) = (NulosN(.TextMatrix(.Rows - 1, COLUMNAPROGCUMPLIDO_)) + NulosN(.TextMatrix(.Rows - 1, COLUMNAPROGRESTANTE_))) - NulosN(.TextMatrix(.Rows - 1, COLUMNATOTALPROD_))
                .TextMatrix(.Rows - 1, COLUMNAPROGDESF_) = Format(.TextMatrix(.Rows - 1, COLUMNAPROGDESF_), FORMAT_CANTIDAD)
                pintarGrid fg(2), COLUMNAPROGDESF_, &H0&, &HFF&
            Else
                .TextMatrix(.Rows - 1, COLUMNAPROGAPROBADO_) = 0
                .TextMatrix(.Rows - 1, COLUMNAPROGCUMPLIDO_) = 0
                .TextMatrix(.Rows - 1, COLUMNAPROGCANCELADO_) = 0
                .TextMatrix(.Rows - 1, COLUMNAPROGRESTANTE_) = 0
                .TextMatrix(.Rows - 1, COLUMNAPROGTOT_) = 0
                .TextMatrix(.Rows - 1, COLUMNAPROGDESF_) = 0
            End If
            '**************************************************************************
            ' Resto
            .TextMatrix(.Rows - 1, COLUMNARESTO_) = NulosN(.TextMatrix(.Rows - 1, COLUMNAPRODUCIDO_)) - NulosN(.TextMatrix(.Rows - 1, COLUMNATOTALPROD_))
            pintarGrid fg(2), COLUMNARESTO_, &H0&, &HFF&
            
            ' %
            If NulosN(.TextMatrix(.Rows - 1, COLUMNARESTO_)) > 0 Then
                .TextMatrix(.Rows - 1, COLUMNARESTOPORC_) = 0
            Else
                If (Abs(NulosN(.TextMatrix(.Rows - 1, COLUMNARESTO_))) + NulosN(.TextMatrix(.Rows - 1, COLUMNAPRODUCIDO_))) = 0 Then
                    .TextMatrix(.Rows - 1, COLUMNARESTOPORC_) = 0
                Else
                    .TextMatrix(.Rows - 1, COLUMNARESTOPORC_) = (Abs(NulosN(.TextMatrix(.Rows - 1, COLUMNARESTO_))) / (Abs(NulosN(.TextMatrix(.Rows - 1, COLUMNARESTO_))) + NulosN(.TextMatrix(.Rows - 1, COLUMNAPRODUCIDO_)))) * 100
                End If
            End If
            .TextMatrix(.Rows - 1, COLUMNARESTOPORC_) = Format(.TextMatrix(.Rows - 1, COLUMNARESTOPORC_), FORMAT_PORCENTAJE)
            'pintarGrid Fg(2), COLUMNARESTOPORC_, &H0&, &HFF&
            
            ' unid/hora
            VALOR_ = calcularRendimiento(IDITEM_)
            .TextMatrix(.Rows - 1, COLUMNAUNIXHORA_) = Format(VALOR_, "0.00")
            If VALOR_ = 0 Then .Select .Rows - 1, COLUMNAUNIXHORA_: .CellForeColor = &HFF&: .CellFontBold = True
        
            ' Hrs Trab.
            If NulosN(.TextMatrix(.Rows - 1, COLUMNATOTALPROD_)) > 0 Then ' Si hay por Producir
                If VALOR_ = 0 Then ' Si no se encontro Linea
                    .TextMatrix(.Rows - 1, COLUMNAHRSTRAB_) = 0
                    .Select .Rows - 1, COLUMNAHRSTRAB_
                    .CellForeColor = &HFF&
                    .CellFontBold = True
                Else
                    .TextMatrix(.Rows - 1, COLUMNAHRSTRAB_) = NulosN(.TextMatrix(.Rows - 1, COLUMNATOTALPROD_)) / VALOR_
                End If
            Else ' Si no hay por Producir
                .TextMatrix(.Rows - 1, COLUMNAHRSTRAB_) = 0
            End If
            .TextMatrix(.Rows - 1, COLUMNAHRSTRAB_) = Format(.TextMatrix(.Rows - 1, COLUMNAHRSTRAB_), FORMAT_CANTIDAD)
                        
            
            TOTALHORAS_ = TOTALHORAS_ + NulosN(.TextMatrix(.Rows - 1, COLUMNAHRSTRAB_))
            
            RstPlan.MoveNext
        Wend
    End With
    
    FraProgreso.Visible = False
    configurarGrid
    ' Se llenan los Totales
    fg(2).Rows = fg(2).Rows + 1
    fg(2).TextMatrix(fg(2).Rows - 1, COLUMNAUNIXHORA_) = "TOTAL"
    fg(2).Select fg(2).Rows - 1, COLUMNAUNIXHORA_
    fg(2).CellFontBold = True
    fg(2).CellForeColor = &HFF&
    fg(2).TextMatrix(fg(2).Rows - 1, COLUMNAHRSTRAB_) = Format(TOTALHORAS_, "0.00")
    
    
End Sub

Private Sub pintarGrid(GRID_ As VSFlexGrid, COLUMNA_ As Integer, COLOR1_ As String, COLOR2_ As String)
    Dim A As Integer
    
    With GRID_
        For A = GRID_.FixedRows To .Rows - 1
            .Select A, COLUMNA_
            If NulosN(.TextMatrix(A, COLUMNA_)) >= 0 Then
                .CellForeColor = COLOR1_
            Else
                .CellForeColor = COLOR2_
            End If
        Next A
    End With
End Sub

Private Function calcularHorasCantidad(IDPRO_ As Double, VALOR_ As Variant, HORAS_ As Boolean, CANTIDAD_ As Boolean) As Variant
    Dim TIEMPO_ As Double
    Dim HOR_ As String
    Dim CANT_ As Double
    Dim H_() As String
    
    RSTCRONO_.Filter = "idpro = " & IDPRO_ & ""
    If HORAS_ Then
        If NulosN(RSTCRONO_("unixhora")) = 0 Then
            TIEMPO_ = 0
        Else
            TIEMPO_ = VALOR_ / NulosN(RSTCRONO_("unixhora"))
        End If
        
        HOR_ = Format(Int(TIEMPO_), "00")
        HOR_ = HOR_ & ":" & Format(((TIEMPO_ * 60) Mod 60), "00")
        
        calcularHorasCantidad = HOR_
    End If
    
    If CANTIDAD_ Then
        H_ = Split(VALOR_, ":")
        TIEMPO_ = (60 * Val(H_(0))) + Val(H_(1))
        TIEMPO_ = TIEMPO_ / 60
        
        CANT_ = TIEMPO_ * NulosN(RSTCRONO_("unixhora"))
        calcularHorasCantidad = CANT_
    End If
End Function

Private Function buscarLinea(IDPRO_ As Double) As Double
    Dim xRs As New ADODB.Recordset
    Dim VALOR_ As Double
    
    Set xRs = Nothing
    
    cSQL = "SELECT pro_linea.id AS idlineadet" _
    + vbCr + "FROM pro_linea LEFT JOIN pro_receta ON pro_linea.idrec = pro_receta.id " _
    + vbCr + "Where (((pro_receta.IDITEM) = " & IDPRO_ & ") And ((pro_linea.activo) = -1)) " _
    + vbCr + "GROUP BY pro_linea.id;"

    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then VALOR_ = 0: GoTo SALIR
    If xRs.RecordCount = 0 Then VALOR_ = 0: GoTo SALIR
    
    VALOR_ = NulosN(xRs("idlineadet"))
SALIR:
    buscarLinea = VALOR_
End Function

Private Function calcularRendimiento(IDPRO_ As Double) As Double
    Dim RENDIMIENTO_ As Double
    Dim xRs As New ADODB.Recordset
    
    Set xRs = Nothing
    
    cSQL = "SELECT pro_linea.kghora " _
        + vbCr + "FROM pro_linea LEFT JOIN pro_receta ON pro_linea.idrec = pro_receta.id " _
        + vbCr + "Where (((pro_receta.IDITEM) = " & IDPRO_ & ") And ((pro_linea.activo) = -1)) " _
        + vbCr + "GROUP BY pro_linea.kghora;"
        
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then RENDIMIENTO_ = 0: GoTo SALIR
    If xRs.RecordCount = 0 Then RENDIMIENTO_ = 0: GoTo SALIR
    
    RENDIMIENTO_ = NulosN(xRs("kghora"))
SALIR:
    calcularRendimiento = RENDIMIENTO_
End Function

Private Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    If Rst.State = 0 Then Exit Sub
    With Rst
        If TODO Then .Filter = adFilterNone
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        While Not .EOF
            .Delete
            .MoveNext
        Wend
    End With
End Sub

Private Sub LbMes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim A As Integer
    Dim ENCONTRO_ As Integer
    Dim INDICE_ As Double
    Dim INDICETOPE_ As Double
    
    If Button = 1 Then
        ' Se encuentran las caracteristicas del indice selecionado
        INDICE_ = LbMes.ListIndex
        INDICETOPE_ = LbMes.TopIndex
        
        ' Se verifica que indices estan seleccionados
        For A = 1 To LbMes.ListCount - 1
            LbMes.ListIndex = A
            If LbMes.Selected(A) = True Then
                ENCONTRO_ = ENCONTRO_ + 1
                If ENCONTRO_ = 2 Then A = LbMes.ListCount - 1
            End If
        Next
        
        ' Si hay mas de un seleccionado
        If ENCONTRO_ = 2 Then
            LbMes.Selected(INDICE_) = False
        End If
        
        ' Se seleccionan los indices del inicio
        LbMes.TopIndex = INDICETOPE_
        LbMes.ListIndex = INDICE_
    End If
End Sub

Private Sub menu00_Click() ' Agregar
    Dim FECHINI_ As Date
    Dim FECHFIN_ As Date
    Dim TODODIA_ As Boolean
    
    fg(INDICE_).Rows = fg(INDICE_).Rows + 1
    fg(INDICE_).Select fg(INDICE_).Rows - 1, 2
    fg(INDICE_).TopRow = fg(INDICE_).Rows - 1
    Fg_CellButtonClick INDICE_, fg(INDICE_).Rows - 1, 1
End Sub

Private Sub menu02_Click() ' Eliminar
    If fg(INDICE_).Row < fg(INDICE_).FixedRows Then Exit Sub
    fg(INDICE_).RemoveItem fg(INDICE_).Row
    
    If fg(INDICE_).Rows = fg(INDICE_).FixedRows Then fg(INDICE_).Rows = fg(INDICE_).Rows + 1
End Sub

Private Sub PbCerrar_Click(Index As Integer)
    frm(0).Visible = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 10 Then ' Buscar
        aplicarFiltrado
    End If
    
    If Button.Index = 13 Then ' Exportar Excel
        ExportarExcel fg(2)
    End If
    
    If Button.Index = 15 Then ' Salir
        Unload Me
    End If
End Sub

Sub ExportarExcel(ByRef GRID_ As VSFlexGrid)
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "REPORTE DE PRODUCCIÓN"

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, GRID_, TITULO_, "", "", TITULO_
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub

'Metodos para arrastrar el Frame
''''''''''''''''''''''''''''''''
Private Sub frm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    frm(Index).ZOrder 0
End Sub

Private Sub frm_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With frm(Index)
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub

