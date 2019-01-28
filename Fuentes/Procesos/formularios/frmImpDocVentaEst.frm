VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImpDocVentaEst 
   Caption         =   "Herraientas - Importar Documentos de Venta (Estudio)"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   12030
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
         Picture         =   "frmImpDocVentaEst.frx":0000
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
         Picture         =   "frmImpDocVentaEst.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   870
         Width           =   240
      End
      Begin VB.CommandButton CmdBusIGV 
         Enabled         =   0   'False
         Height          =   240
         Left            =   2595
         Picture         =   "frmImpDocVentaEst.frx":0264
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1185
         Width           =   240
      End
      Begin VB.CommandButton CmdBusISC 
         Enabled         =   0   'False
         Height          =   240
         Left            =   8475
         Picture         =   "frmImpDocVentaEst.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   870
         Width           =   240
      End
      Begin VB.CommandButton CmdBusImpTot 
         Enabled         =   0   'False
         Height          =   240
         Left            =   8475
         Picture         =   "frmImpDocVentaEst.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1185
         Width           =   240
      End
      Begin VB.CommandButton CmdBusArch2 
         Enabled         =   0   'False
         Height          =   240
         Left            =   7050
         Picture         =   "frmImpDocVentaEst.frx":05FA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   540
         Width           =   240
      End
      Begin VB.CommandButton CmdBusMes 
         Height          =   240
         Left            =   10230
         Picture         =   "frmImpDocVentaEst.frx":072C
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
            Picture         =   "frmImpDocVentaEst.frx":085E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":0DA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":1134
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":12B8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":170C
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":1824
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":1D68
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":22AC
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":23C0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":24D4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":2928
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImpDocVentaEst.frx":2A94
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmImpDocVentaEst.frx":2FDC
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
      Width           =   12030
      _ExtentX        =   21220
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmImpDocVentaEst.frx":3206
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
Attribute VB_Name = "frmImpDocVentaEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim RstTmp As New ADODB.Recordset
Dim QueHace As Integer

Sub ActivarTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(4).Enabled = Not Toolbar1.Buttons(4).Enabled
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(7).Enabled = Not Toolbar1.Buttons(7).Enabled
End Sub

Sub Blanquea()
    TxtArchivo.Text = ""
    TxtArchivo2.Text = ""
    TxtCtaBru.Text = ""
    TxtCtaISC.Text = ""
    TxtCtaIGV.Text = ""
    TxtCtaTot.Text = ""
    TxtMes.Text = ""
    
    Label71.Caption = ""
    Label72.Caption = ""
    Label73.Caption = ""
    Label74.Caption = ""
    
    LblIdMes.Caption = ""
    LblIdCtaISC.Caption = ""
    LblIdCtaImpTot.Caption = ""
    LblCtaIGV.Caption = ""
    LblIdCtaBru.Caption = ""
End Sub

Sub Bloquea()
    CmdBusArch.Enabled = Not CmdBusArch.Enabled
    CmdBusBruto.Enabled = Not CmdBusBruto.Enabled
    CmdBusISC.Enabled = Not CmdBusISC.Enabled
    CmdBusIGV.Enabled = Not CmdBusIGV.Enabled
    CmdBusImpTot.Enabled = Not CmdBusImpTot.Enabled
    CmdBusArch2.Enabled = Not CmdBusArch2.Enabled
    'CmdBusMes.Enabled = Not CmdBusMes.Enabled
    'CmdCargar.Enabled = Not CmdCargar.Enabled
End Sub

Private Sub CmdBusArch_Click()
    'CommonDialog1.CancelError = True
    'Especificar las extensiones a usar
    CommonDialog1.DefaultExt = "*.xls"
    'CommonDialog1.Filter = "Cardfile (*.crd)|*.crd|Textos (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    CommonDialog1.Filter = "Documentos de Excel (*.xls)|*.xls"
    CommonDialog1.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
    Else
        TxtArchivo.Text = CommonDialog1.FileName
    End If
End Sub

Sub CargaDocumentos()
    Dim xNumFilas As Integer
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    'Dim objExcel As New Excel.Application
    
    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo.Text)
    
    Frame2.Left = 3090
    Frame2.Top = 2910
    Frame2.Visible = True
    PreparaRstTmp
    
    xFilas = 3
    xNumFilas = 1
    Fg1.Rows = 1
    Fg2.Rows = 1
    
    Fg1.Rows = 1
    
    With objExcel.ActiveSheet
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        Label4.Caption = "Calculando numero de registros"
        ProgressBar2.Max = 32000
        For A = 2 To 32000
            ProgressBar2.Value = A
            If NulosC(.Cells(A, 1)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit For
            End If
        Next A
        
        xNumFilas = xNumFilas + 1
        Label4.Caption = "Cargando registros para la importacion"
        ProgressBar2.Max = xNumFilas
        
        For A = 2 To xNumFilas
            ProgressBar2.Value = A
            Frame2.Refresh
            
            If NulosC(.Cells(A, 1)) = "" Then Exit For
            Fg1.Rows = Fg1.Rows + 1
            For B = 1 To 17
                Fg1.TextMatrix(A - 1, B) = Trim(.Cells(A, B))
                If (B = 2) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(NulosN(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_documento", "N", xCon)
                If (B = 3) Or (B = 4) Then Fg1.TextMatrix(A - 1, B) = Format(CDate(Trim(.Cells(A, B))), "dd/mm/yy")
                
                If (B = 14) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(NulosN(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_moneda", "N", xCon)
                If (B = 15) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(NulosN(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_condpago", "N", xCon)
                If (B = 16) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(NulosN(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_tipoventa", "N", xCon)
                If (B = 17) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(NulosN(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_tipoproducto", "N", xCon)
            Next B
            
        Next A
    End With
    
    'CARGAMOS EL DETALLE DE LA VENTA
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo2.Text)
    
    Fg2.Rows = 1
    With objExcel.ActiveSheet
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        Label4.Caption = "Calculando numero de registros"
        ProgressBar2.Max = 32000
        For A = 2 To 32000
            ProgressBar2.Value = A
            If NulosC(.Cells(A, 1)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit For
            End If
        Next A
        
        xNumFilas = xNumFilas + 1
        ProgressBar2.Max = xNumFilas
        Label4.Caption = "Cargando Registros"
        
        Dim xIdPro As Integer
        
        For A = 2 To xNumFilas
            ProgressBar2.Value = A
            Frame2.Refresh
            
            If NulosC(.Cells(A, 1)) = "" Then Exit For
            Fg2.Rows = Fg2.Rows + 1
            RstTmp.AddNew
            For B = 1 To 7
                Fg2.TextMatrix(A - 1, B) = Trim(.Cells(A, B))
                
                If B = 1 Then RstTmp("numruc") = NulosC(.Cells(A, B))
                If B = 2 Then RstTmp("numdoc") = NulosC(.Cells(A, B))
                If B = 3 Then
                    xIdPro = NulosN(Busca_Codigo(NulosC(.Cells(A, B)), "codpro", "id", "alm_inventario", "C", xCon))
                    
                    If xIdPro = 0 Then
                        MsgBox "El item " + NulosC(.Cells(A, B + 1)) + " no esta registrado como item en el sistema", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                        Fg1.Rows = 1
                        Fg2.Rows = 1
                        Frame2.Visible = False
                        Set RstTmp = Nothing
                        Exit Sub
                    End If
                    
                    RstTmp("codite") = Trim(.Cells(A, B))
                End If
                If B = 4 Then RstTmp("descri") = NulosC(.Cells(A, B))
                If B = 5 Then RstTmp("unimed") = NulosN(.Cells(A, B))
                If (B = 5) Then Fg2.TextMatrix(A - 1, B) = Busca_Codigo(NulosN(Fg2.TextMatrix(A - 1, B)), "id", "descripcion", "mae_unidades", "N", xCon)
                If B = 6 Then RstTmp("cantid") = NulosN(.Cells(A, B))
                If B = 7 Then RstTmp("precio") = NulosN(.Cells(A, B))
                
            Next B
            RstTmp.Update
        Next A
    End With
    
    
    Frame2.Visible = False
    MsgBox "El proceso termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
    objExcel.WindowState = 2
    objExcel.Workbooks.Close
    
    Set objExcel = Nothing
    Exit Sub
End Sub

Private Sub CmdBusArch2_Click()
    'CommonDialog1.CancelError = True
    'Especificar las extensiones a usar
    CommonDialog1.DefaultExt = "*.xls"
    'CommonDialog1.Filter = "Cardfile (*.crd)|*.crd|Textos (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    CommonDialog1.Filter = "Documentos de Excel (*.xls)|*.xls"
    CommonDialog1.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
    Else
        TxtArchivo2.Text = CommonDialog1.FileName
    End If
End Sub

Private Sub CmdBusBruto_Click()
'    Dim xfrm As New SGI2_funciones.formularios
'    Dim Rst As New ADODB.Recordset
'    Set Rst = xfrm.SelePlanCuentas(xCon)
'    If Rst.State = 1 Then
'        If Rst.RecordCount <> 0 Then
'            TxtCtaBru.Text = Trim(Rst("cuenta"))
'            Label71.Caption = Trim(Rst("descripcion"))
'            LblIdCtaBru.Caption = Trim(Rst("id"))
'            TxtCtaISC.SetFocus
'        End If
'    End If
'    Set xfrm = Nothing
    
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nº Cuenta":    xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Cuenta Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtCtaBru.Text = Trim(xRs("cuenta"))
        Label71.Caption = Trim(xRs("descripcion"))
        LblIdCtaBru.Caption = Trim(xRs("id"))
        TxtCtaISC.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusIGV_Click()
'    Dim xfrm As New SGI2_funciones.formularios
'    Dim Rst As New ADODB.Recordset
'    Set Rst = xfrm.SelePlanCuentas(xCon)
'    If Rst.State = 1 Then
'        If Rst.RecordCount <> 0 Then
'            TxtCtaIGV.Text = Trim(Rst("cuenta"))
'            Label73.Caption = Trim(Rst("descripcion"))
'            LblCtaIGV.Caption = Trim(Rst("id"))
'            TxtCtaIGV.SetFocus
'        End If
'    End If
'    Set xfrm = Nothing

    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nº Cuenta":    xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Cuenta Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtCtaIGV.Text = Trim(xRs("cuenta"))
        Label73.Caption = Trim(xRs("descripcion"))
        LblCtaIGV.Caption = Trim(xRs("id"))
        TxtCtaIGV.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusImpTot_Click()
'    Dim xfrm As New SGI2_funciones.formularios
'    Dim Rst As New ADODB.Recordset
'    Set Rst = xfrm.SelePlanCuentas(xCon)
'    If Rst.State = 1 Then
'        If Rst.RecordCount <> 0 Then
'            TxtCtaTot.Text = Trim(Rst("cuenta"))
'            Label74.Caption = Trim(Rst("descripcion"))
'            LblIdCtaImpTot.Caption = Trim(Rst("id"))
'            CmdCargar.SetFocus
'        End If
'    End If
'    Set xfrm = Nothing

    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nº Cuenta":    xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Cuenta Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtCtaTot.Text = Trim(xRs("cuenta"))
        Label74.Caption = Trim(xRs("descripcion"))
        LblIdCtaImpTot.Caption = Trim(xRs("id"))
        CmdCargar.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusISC_Click()
'    Dim xfrm As New SGI2_funciones.formularios
'    Dim Rst As New ADODB.Recordset
'    Set Rst = xfrm.SelePlanCuentas(xCon)
'    If Rst.State = 1 Then
'        If Rst.RecordCount <> 0 Then
'            TxtCtaISC.Text = Trim(Rst("cuenta"))
'            Label72.Caption = Trim(Rst("descripcion"))
'            LblIdCtaISC.Caption = Trim(Rst("id"))
'        End If
'    End If
'    Set xfrm = Nothing

    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nº Cuenta":    xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Cuenta Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtCtaISC.Text = Trim(xRs("cuenta"))
        Label72.Caption = Trim(xRs("descripcion"))
        LblIdCtaISC.Caption = Trim(xRs("id"))
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusMes_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Documento":    xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Codigo":       xCampos2(1, 1) = "id":             xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"

    xform.SQLCad = "SELECT * FROM con_meses"
    xform.Titulo = "Buscando Mes de Trabajo"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtMes.Text = xRs("descripcion")
        LblIdMes.Caption = xRs("id")
        If QueHace = 2 Then
            RST_Busq xRs2, "SELECT * FROM var_importados WHERE idtabla = 1 AND idmes = " & xRs("id") & "", xCon
            If xRs2.RecordCount <> 0 Then
                If MsgBox("Ya se importaron datos para el mes seleccionado" & vbCr & "Desea seguir importando para este mes", vbInformation + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then
                    LblIdMes.Caption = ""
                    TxtMes.Text = ""
                    'TxtMes.SetFocus
                    Exit Sub
                End If
            End If
        End If
        TxtCtaBru.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCargar_Click()
    If QueHace = 2 Then
        If TxtArchivo.Text = "" Then
            MsgBox "No ha especificado el nombre del archivo cabecera ", vbInformation + vbOKCancel + vbCritical, xTitulo
            TxtArchivo.SetFocus
            Exit Sub
        End If
        
        If TxtArchivo2.Text = "" Then
            MsgBox "No ha especificado el nombre del archivo detalle ", vbInformation + vbOKCancel + vbCritical, xTitulo
            TxtArchivo2.SetFocus
            Exit Sub
        End If
        
        CargaDocumentos  'estamos cargando datos de un archivo de excel
    End If
    
    If QueHace = 3 Then
        If TxtMes.Text = "" Then
            MsgBox "No ha especificado el mes a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtMes.SetFocus
            Exit Sub
        End If
        MostrarDocumentos  'mostramos los registros ya guardados
    End If
    
End Sub

Sub MostrarDocumentos()
    Dim rst As New ADODB.Recordset
    Dim A As Integer
    
    'CARGAMOS LA CABECERA DE LAS FACTURAS
    RST_Busq rst, "SELECT var_importados.idtabla, var_importados.idmes, mae_cliente.numruc, mae_documento.abrev, vta_ventas.fchdoc, " _
        & " vta_ventas.fchven, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.impbru, vta_ventas.impigv, " _
        & " vta_ventas.impisc, vta_ventas.imptotdoc, vta_ventas.impsal, mae_moneda.simbolo AS descmon, mae_condpago.descripcion AS descconpag, " _
        & " mae_tipoventa.descripcion AS tipven, mae_tipoproducto.descripcion AS tippro " _
        & " FROM ((((((var_importados LEFT JOIN vta_ventas ON var_importados.iddoc = vta_ventas.id) LEFT JOIN mae_documento " _
        & " ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_condpago " _
        & " ON vta_ventas.idconpag = mae_condpago.id) LEFT JOIN mae_tipoventa ON vta_ventas.idtipven = mae_tipoventa.id) " _
        & " LEFT JOIN mae_tipoproducto ON vta_ventas.idtipo = mae_tipoproducto.id) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id " _
        & " Where (((var_importados.idtabla) = 1) And ((var_importados.idmes) = " & Val(LblIdMes.Caption) & ")) " _
        & " ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]", xCon
    
    Fg1.Rows = 1
    If rst.RecordCount = 0 Then
        MsgBox "No se han importado ventas en el mes especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtMes.Text = ""
        LblIdMes.Caption = ""
        Exit Sub
    Else
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(rst("numruc"))
            Fg1.TextMatrix(A, 2) = NulosC(rst("abrev"))
            Fg1.TextMatrix(A, 3) = Format(rst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(A, 4) = Format(rst("fchven"), "dd/mm/yy")
            Fg1.TextMatrix(A, 5) = NulosC(rst("numdoc"))
            Fg1.TextMatrix(A, 6) = Format(rst("impbru"), "0.00")
            Fg1.TextMatrix(A, 7) = Format(rst("impigv"), "0.00")
            Fg1.TextMatrix(A, 9) = Format(rst("imptotdoc"), "0.00")
            Fg1.TextMatrix(A, 8) = Format(rst("impisc"), "0.00")
            Fg1.TextMatrix(A, 10) = Format(rst("impsal"), "0.00")
            Fg1.TextMatrix(A, 11) = NulosC(rst("descmon"))
            Fg1.TextMatrix(A, 12) = NulosC(rst("descconpag"))
            Fg1.TextMatrix(A, 13) = NulosN(rst("tipven"))
            Fg1.TextMatrix(A, 14) = NulosN(rst("tippro"))
            rst.MoveNext
            If rst.EOF = True Then Exit For
            
        Next A
    End If
    
    'CARGAMOS EL DETALLE DE LA FACTURA
    Set rst = Nothing
    Fg2.Rows = 1
    RST_Busq rst, "SELECT var_importados.idtabla, var_importados.idmes, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, " _
        & " alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, vta_ventasdet.canpro, vta_ventasdet.preuni, " _
        & " vta_ventasdet.imptot FROM (var_importados LEFT JOIN (vta_ventas LEFT JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta) " _
        & " ON var_importados.iddoc = vta_ventas.id) LEFT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) " _
        & " ON vta_ventasdet.iditem = alm_inventario.id Where (((var_importados.idtabla) = 1) And ((var_importados.idmes) = " & Val(LblIdMes.Caption) & ") " _
        & " And ((alm_inventario.codpro) Is Not Null)) ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]", xCon

    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = NulosC(rst("numdoc"))
            Fg2.TextMatrix(A, 2) = rst("codpro")
            Fg2.TextMatrix(A, 3) = rst("descripcion")
            Fg2.TextMatrix(A, 4) = rst("abrev")
            Fg2.TextMatrix(A, 5) = Format(rst("canpro"), "0.00")
            Fg2.TextMatrix(A, 6) = Format(rst("preuni"), "0.00")
            
            rst.MoveNext
            If rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Fg1.Rows = 1 Then Exit Sub
    If KeyCode = 57 Then
        Fg1.RemoveItem Fg1.Row
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then
        QueHace = True
        SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = False
    Blanquea
    QueHace = 3
    Fg1.Rows = 1
    Fg2.Rows = 1
End Sub

Sub Modificar()
    ActivarTool
    QueHace = 2
    Bloquea
    TxtArchivo.SetFocus
End Sub

Sub Eliminar()
    If TxtMes.Text = "" Then
        MsgBox "No ha especificado el mes", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim Rpta As Integer
    Dim xMes As Integer
    'xMes = SeleccionaMes(xCon)
    If Val(LblIdMes.Caption) <> 0 Then
        Rpta = MsgBox("¿Esta seguro de eliminar los datos importados del mes seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            BorrarDatos Val(LblIdMes.Caption)
            MsgBox "Los datos importados se eliminaron con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Fg1.Rows = 1
            Fg2.Rows = 1
            TxtMes.Text = ""
            LblIdMes.Caption = ""
        End If
    End If
End Sub

Sub BorrarDatos(MesBorrar As Integer)
    Dim A As Integer
    Dim rst As New ADODB.Recordset
    
    RST_Busq rst, "SELECT * FROM var_importados WHERE idtabla = 1 AND idmes = " & MesBorrar & "", xCon
    
    If rst.RecordCount <> 0 Then
        Frame2.Left = 3090
        Frame2.Top = 2910
        Label4.Caption = "Eliminando Datos"
        ProgressBar2.Max = rst.RecordCount
        Frame2.Visible = True
        
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            ProgressBar2.Value = A
            Frame2.Refresh
            'borramos el diario
            xCon.Execute "DELETE con_diario.idmes, con_diario.idlib, con_diario.idmov From con_diario " _
                & " WHERE (((con_diario.idmes)=" & MesBorrar & ") AND ((con_diario.idlib)=2) AND ((con_diario.idmov)=" & rst("iddoc") & "))"
            
            xCon.Execute "DELETE * FROM vta_ventas WHERE id = " & rst("iddoc") & ""

            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
        Frame2.Visible = False
        
        xCon.Execute "DELETE * FROM var_importados WHERE idtabla = 1 AND idmes = " & MesBorrar & ""
    End If
End Sub

Sub Cancelar()
    Blanquea
    Bloquea
    ActivarTool
    QueHace = 3
End Sub

Function Grabar() As Boolean
    If TxtCtaBru.Text = "" Then
        MsgBox "No ha especificado la cuenta contable para el importe bruto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCtaBru.SetFocus
        Grabar = False
        Exit Function
    End If
    
    If TxtCtaISC.Text = "" Then
        MsgBox "No ha especificado la cuenta contable para el impuesto Selectivo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCtaISC.SetFocus
        Grabar = False
        Exit Function

    End If
    
    If TxtCtaIGV.Text = "" Then
        MsgBox "No ha especificado la cuenta contable para el impuesto I.G.V.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCtaIGV.SetFocus
        Grabar = False
        Exit Function
    End If
    
    If TxtCtaTot.Text = "" Then
        MsgBox "No ha especificado la cuenta contable para el total del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCtaTot.SetFocus
        Grabar = False
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No se han cargado documentos de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Grabar = False
        Exit Function
    End If
    
    Dim A As Double
    Dim RstCab As New ADODB.Recordset
    Dim Rstdet As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim RstImp As New ADODB.Recordset
    Dim rst As New ADODB.Recordset
    Dim RstTC As New ADODB.Recordset
    
    Dim xId As Double
    Dim xNumAsiento As String
    Dim xIdItem As Integer
    Dim xCodMon As Integer
    Dim xMes As Integer
    
    Dim nSQL  As String
    
    Frame2.Left = 3090
    Frame2.Top = 2910
    Label4.Caption = "Verificando Datos"
    Frame2.Visible = True
    ProgressBar2.Max = Fg1.Rows - 1
    
    For A = 1 To Fg1.Rows - 1
        ProgressBar2.Value = A
        Frame2.Refresh

        RST_Busq rst, "SELECT * FROM mae_cliente WHERE numruc = '" & Fg1.TextMatrix(A, 1) & "'", xCon
        If rst.RecordCount = 0 Then
            MsgBox "El Nº R.U.C. " + NulosC(Fg1.TextMatrix(A, 1)) + " no existe en el maestro de cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo

            Frame2.Visible = False
            Set rst = Nothing
            Grabar = False
            Exit Function
        End If
        Set rst = Nothing
    Next A
    
    Label4.Caption = "Importando registros"
   
    ProgressBar2.Max = Fg1.Rows - 1
    
    RST_Busq RstCab, "SELECT TOP 1 * FROM vta_ventas", xCon
    RST_Busq Rstdet, "SELECT TOP 1 * FROM vta_ventasdet", xCon
    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    RST_Busq RstImp, "SELECT TOP 1 * FROM var_importados", xCon
    
    xMes = LblIdMes.Caption
    Dim xTipDoc As Integer
    Dim xTotalxxx As Double
    
    For A = 1 To Fg1.Rows - 1
        ProgressBar2.Value = A
        Frame2.Refresh
        'grabamos el documeto de venta
        xId = HallaCodigoTabla("vta_ventas", xCon, "id")
        xTipDoc = 0
        xTipDoc = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 2)), "descripcion", "id", "mae_documento", "C", xCon)
        RstCab.AddNew
        RstCab("id") = xId
        RstCab("idlib") = 2
        RstCab("idtipo") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 17)), "descripcion", "id", "mae_tipoproducto", "C", xCon)
        RstCab("idcli") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 1)), "numruc", "id", "mae_cliente", "C", xCon)
        RstCab("tipdoc") = xTipDoc
        RstCab("numser") = Format(Mid(Fg1.TextMatrix(A, 5), 1, 4), "0000")
        RstCab("numdoc") = Format(Mid(Fg1.TextMatrix(A, 5), 6, 10), "0000000000")
        RstCab("fchreg") = CDate("01/" + Format(LblIdMes.Caption, "00") + "/" + Trim(Str(Val(AnoTra))))
        RstCab("fchdoc") = CDate(Fg1.TextMatrix(A, 3))
        RstCab("fchven") = CDate(Fg1.TextMatrix(A, 4))
        RstCab("idconpag") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 15)), "descripcion", "id", "mae_condpago", "C", xCon)
        
        xCodMon = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 14)), "descripcion", "id", "mae_moneda", "C", xCon)
        RstCab("idmon") = xCodMon
        RstCab("impbru") = NulosN(Fg1.TextMatrix(A, 6))
        RstCab("impbru2") = NulosN(Fg1.TextMatrix(A, 7))
        RstCab("impbru3") = NulosN(Fg1.TextMatrix(A, 8))
        RstCab("impinaf") = NulosN(Fg1.TextMatrix(A, 9))
        
        
        RstCab("impigv") = NulosN(Fg1.TextMatrix(A, 10))
        RstCab("impisc") = NulosN(Fg1.TextMatrix(A, 11))
        RstCab("imptotdoc") = NulosN(Fg1.TextMatrix(A, 12))
        RstCab("impsal") = NulosN(Fg1.TextMatrix(A, 13))
        RstCab("idtipven") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 16)), "descripcion", "id", "mae_tipoventa", "C", xCon)
        RstCab("importado") = -1
        RstCab("oriitem") = 1 'especificamos que la venta ha sido directa, luego se actualizara este flag cuando se importen las guias
        xTotalxxx = 0
        xTotalxxx = NulosN((Fg1.TextMatrix(A, 6))) + NulosN(Val(Fg1.TextMatrix(A, 7))) + NulosN(Val(Fg1.TextMatrix(A, 8))) + NulosN(Val(Fg1.TextMatrix(A, 9)))
        
        
        If xTotalxxx = 0 Then
            RstCab("anulado") = -1
        End If
        RstCab.Update
        
        Dim H As Integer
        
        If xTotalxxx <> 0 Then
            'grabamos el detalle del documento de venta
            RstTmp.Filter = adFilterNone
            RstTmp.Filter = "numruc='" & Fg1.TextMatrix(A, 1) & "' and numdoc = '" & Fg1.TextMatrix(A, 5) & "' "
            If RstTmp.RecordCount <> 0 Then
                Dim C As Integer
                RstTmp.MoveFirst
                For C = 1 To RstTmp.RecordCount
                    Rstdet.AddNew
                    Rstdet("idvta") = xId
                    
                    xIdItem = Busca_Codigo(NulosC(RstTmp("codite")), "codpro", "id", "alm_inventario", "C", xCon)
                    
                    Rstdet("iditem") = xIdItem
                    Rstdet("idunimed") = NulosN(RstTmp("unimed")) 'Busca_Codigo(NulosN(RstTmp("unimed")), "id", "descripcion", "mae_unidades", "N", xCon)
                    Rstdet("canpro") = NulosN(RstTmp("cantid"))
                    Rstdet("preuni") = NulosN(RstTmp("precio"))
                    Rstdet("imptot") = NulosN(RstTmp("precio")) * NulosN(RstTmp("cantid"))
                    Rstdet.Update
                    RstTmp.MoveNext
                    If RstTmp.EOF = True Then Exit For
                Next C
            End If
        
            xNumAsiento = NuevoNumAsiento(2, xMes, xCon)
            xCon.Execute "UPDATE vta_ventas SET vta_ventas.numreg = '" & Format(xMes, "00") + xNumAsiento & "' WHERE (((vta_ventas.id)=" & xId & "))"
    
            'grabamos el diario
            'grabamos las cuentas debe
            'importe bruto
            
            RstTmp.Filter = adFilterNone
            RstTmp.Filter = "numruc='" & Fg1.TextMatrix(A, 1) & "' and numdoc = '" & Fg1.TextMatrix(A, 5) & "' "
            
            If RstTmp.RecordCount <> 0 Then
                H = 0
                RstTmp.MoveFirst
                For H = 1 To RstTmp.RecordCount
                    RstDia.AddNew
                    RstDia("año") = AnoTra
                    RstDia("idmes") = xMes
                    RstDia("idlib") = 2
                    RstDia("idmov") = xId
                    RstDia("numasi") = xNumAsiento
                    RstDia("idcue") = Busca_Codigo(NulosC(RstTmp("codite")), "codpro", "idcuentaven", "alm_inventario", "C", xCon)
                    RstDia("iddoc") = NulosN(xTipDoc)
                    RstDia("fchasi") = CDate("01/" + Format(LblIdMes.Caption, "00") + "/" + Trim(Str(Val(AnoTra))))
                    RstDia("fchdoc") = CDate(Fg1.TextMatrix(A, 3))
    
                    If xTipDoc <> 7 Then   'si es diferente a la nota de credito
                        If xCodMon = 1 Then
                            RstDia("imphabsol") = NulosN(RstTmp("precio"))
                            RstDia("imphabdol") = 0
                        Else
                            Set RstTC = Nothing
                            Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 3) & "')", xCon)
                            If RstTC.RecordCount <> 0 Then
                                RstDia("tc") = RstTC("impven")
                                'RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 6)) * RstTC("impven")
                                RstDia("imphabsol") = NulosN(RstTmp("precio")) * RstTC("impven")
                                RstDia("imphabdol") = NulosN(RstTmp("precio"))
                            Else
                                RstDia("tc") = 0
                                RstDia("imphabdol") = NulosN(RstTmp("precio"))
                            End If
                        End If
                    Else
                        If xCodMon = 1 Then
                            RstDia("impdebsol") = NulosN(RstTmp("precio"))
                            RstDia("impdebdol") = 0
                        Else
                            Set RstTC = Nothing
                            Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 3) & "')", xCon)
                            If RstTC.RecordCount <> 0 Then
                                RstDia("tc") = RstTC("impven")
                                'RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 6)) * RstTC("impven")
                                RstDia("impdebsol") = NulosN(RstTmp("precio")) * RstTC("impven")
                                RstDia("impdebdol") = NulosN(RstTmp("precio"))
                            Else
                                RstDia("tc") = 0
                                RstDia("impdebdol") = NulosN(RstTmp("precio"))
                            End If
                        End If
                    End If
                    RstDia.Update
                
                    RstTmp.MoveNext
                    If RstTmp.EOF = True Then Exit For
                Next H
            End If
        
            'IGV
''            If NulosN(fg1.TextMatrix(A, 10)) <> 0 Then
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = xMes
                RstDia("idlib") = 2
                RstDia("idmov") = xId
                RstDia("numasi") = xNumAsiento
                RstDia("idcue") = Val(LblCtaIGV.Caption)
                RstDia("iddoc") = NulosN(xTipDoc)
                RstDia("fchasi") = CDate("01/" + Format(LblIdMes.Caption, "00") + "/" + Trim(Str(Val(AnoTra))))
                RstDia("fchdoc") = CDate(Fg1.TextMatrix(A, 3))
                
                If xTipDoc <> 7 Then   'si es diferente a la nota de credito
                    If xCodMon = 1 Then
                        RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 10))
                        RstDia("imphabdol") = 0
                    Else
                        Set RstTC = Nothing
                        Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 3) & "')", xCon)
                        If RstTC.RecordCount <> 0 Then
                            RstDia("tc") = RstTC("impven")
                            RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 10)) * RstTC("impven")
                            RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 10))
                        Else
                            RstDia("tc") = 0
                            RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 10))
                        End If
                    End If
                Else
                    If xCodMon = 1 Then
                        RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 10))
                        RstDia("impdebdol") = 0
                    Else
                        Set RstTC = Nothing
                        Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 3) & "')", xCon)
                        If RstTC.RecordCount <> 0 Then
                            RstDia("tc") = RstTC("impven")
                            RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 10)) * RstTC("impven")
                            RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 10))
                        Else
                            RstDia("tc") = 0
                            RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 10))
                        End If
                    End If
                End If
                RstDia.Update
''            End If
        
            'ISC
            If NulosN(Fg1.TextMatrix(A, 11)) <> 0 Then
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = xMes
                RstDia("idlib") = 2
                RstDia("idmov") = xId
                RstDia("numasi") = xNumAsiento
                RstDia("idcue") = Val(LblIdCtaISC.Caption)
                RstDia("iddoc") = NulosN(xTipDoc)
                RstDia("fchasi") = CDate("01/" + Format(LblIdMes.Caption, "00") + "/" + Trim(Str(Val(AnoTra))))
                RstDia("fchdoc") = CDate(Fg1.TextMatrix(A, 3))
                
                If xTipDoc <> 7 Then   'si es diferente a la nota de credito
                    If xCodMon = 1 Then
                        RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 11))
                        RstDia("imphabdol") = 0
                    Else
                        Set RstTC = Nothing
                        Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 3) & "')", xCon)
                        If RstTC.RecordCount <> 0 Then
                            RstDia("tc") = RstTC("impven")
                            RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 11)) * RstTC("impven")
                            RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 11))
                        Else
                            RstDia("tc") = 0
                            RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 11))
                        End If
                    End If
                Else
                    If xCodMon = 1 Then
                        RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 11))
                        RstDia("impdebdol") = 0
                    Else
                        Set RstTC = Nothing
                        Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 3) & "')", xCon)
                        If RstTC.RecordCount <> 0 Then
                            RstDia("tc") = RstTC("impven")
                            RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 11)) * RstTC("impven")
                            RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 11))
                        Else
                            RstDia("tc") = 0
                            RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 11))
                        End If
                    End If
                End If
                RstDia.Update
            End If
        
            'Importe total
            If NulosN(Fg1.TextMatrix(A, 12)) <> 0 Then
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = xMes
                RstDia("idlib") = 2
                RstDia("idmov") = xId
                RstDia("numasi") = xNumAsiento
                'If xCodMon = 1 Then
                '    RstDia("idcue") = 65
                'Else
                RstDia("idcue") = Val(LblIdCtaImpTot.Caption)
                'End If
                
                RstDia("iddoc") = NulosN(xTipDoc)
                RstDia("fchasi") = CDate("01/" + Format(LblIdMes.Caption, "00") + "/" + Trim(Str(Val(AnoTra))))
                RstDia("fchdoc") = CDate(Fg1.TextMatrix(A, 3))
                
                If xTipDoc <> 7 Then   'si es diferente a la nota de credito
                    If xCodMon = 1 Then
                        RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 12))
                        RstDia("impdebdol") = 0
                    Else
                        Set RstTC = Nothing
                        Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 3) & "')", xCon)
                        If RstTC.RecordCount <> 0 Then
                            RstDia("tc") = RstTC("impven")
                            RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 12)) * RstTC("impven")
                            RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 12))
                        Else
                            RstDia("tc") = 0
                            RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 12))
                        End If
                    End If
                Else
                    If xCodMon = 1 Then
                        RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 12))
                        RstDia("imphabdol") = 0
                    Else
                        Set RstTC = Nothing
                        Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 3) & "')", xCon)
                        If RstTC.RecordCount <> 0 Then
                            RstDia("tc") = RstTC("impven")
                            RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 12)) * RstTC("impven")
                            RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 12))
                        Else
                            RstDia("tc") = 0
                            RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 12))
                        End If
                    End If
                End If
                RstDia.Update
            End If
        Else
            xNumAsiento = NuevoNumAsiento(2, xMes, xCon)
            xCon.Execute "UPDATE vta_ventas SET vta_ventas.numreg = '" & Format(xMes, "00") + xNumAsiento & "' WHERE (((vta_ventas.id)=" & xId & "))"
        
            'importe bruto
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = xMes
            RstDia("idlib") = 2
            RstDia("idmov") = xId
            RstDia("numasi") = xNumAsiento
            RstDia("idcue") = Val(LblIdCtaBru.Caption)

            RstDia("iddoc") = NulosN(xTipDoc)
            RstDia("fchasi") = CDate("01/" + Format(LblIdMes.Caption, "00") + "/" + Trim(Str(Val(AnoTra))))
            RstDia("fchdoc") = CDate(Fg1.TextMatrix(A, 3))
            RstDia("impdebsol") = 0
            RstDia("impdebdol") = 0
            RstDia.Update
        
            'importe total
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = xMes
            RstDia("idlib") = 2
            RstDia("idmov") = xId
            RstDia("numasi") = xNumAsiento
            RstDia("idcue") = Val(LblIdCtaImpTot.Caption)

            RstDia("iddoc") = NulosN(xTipDoc)
            RstDia("fchasi") = CDate("01/" + Format(LblIdMes.Caption, "00") + "/" + Trim(Str(Val(AnoTra))))
            RstCab("fchdoc") = CDate(Fg1.TextMatrix(A, 3))
            RstDia("imphabsol") = 0
            RstDia("imphabdol") = 0
            RstDia.Update
        
        End If
        
        nSQL = "UPDATE (vta_ventas INNER JOIN con_diario ON vta_ventas.id = con_diario.idmov) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            + vbCr + " SET con_diario.fchdoc=vta_ventas.fchdoc, con_diario.idmon=vta_ventas.idmon, con_diario.ridlib = 2, con_diario.ridtipper = 2, con_diario.ridper = [vta_ventas].[idcli], con_diario.rtipdoc = [vta_ventas].[tipdoc], con_diario.rfchope = [vta_ventas].[fchdoc], con_diario.rnumerodoc = IIf([vta_ventas].[numser] Is Null Or [vta_ventas].[numser]='','',[vta_ventas].[numser] & '-') & [vta_ventas].[numdoc], con_diario.rglosaope = [vta_ventas].[glosa] & '', con_diario.rregistro = Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) " _
            + vbCr + " WHERE con_diario.idlib=2 and con_diario.idmov= " & xId & " ; "
        xCon.Execute nSQL
        
        RstImp.AddNew
        RstImp("iddoc") = xId
        RstImp("idtabla") = 1
        RstImp("idmes") = xMes
        RstImp.Update
    Next A
    
    Frame2.Visible = False
    MsgBox "Los datos se importaron con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        Modificar
    End If
    If Button.Index = 2 Then
        Eliminar
    End If
    If Button.Index = 4 Then
        Cancelar
    End If
    If Button.Index = 5 Then
        If Grabar = True Then
            Fg1.Rows = 1
            Fg2.Rows = 1
            Cancelar
        End If
    End If
    If Button.Index = 7 Then
        Unload Me
    End If
End Sub


Sub PreparaRstTmp()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(7, 3) As String

    xCampos(0, 0) = "numdoc":        xCampos(0, 1) = "C":      xCampos(0, 2) = "15"
    xCampos(1, 0) = "codite":        xCampos(1, 1) = "C":      xCampos(1, 2) = "50"
    xCampos(2, 0) = "descri":        xCampos(2, 1) = "C":      xCampos(2, 2) = "100"
    xCampos(3, 0) = "unimed":        xCampos(3, 1) = "N":      xCampos(3, 2) = "8"
    xCampos(4, 0) = "cantid":        xCampos(4, 1) = "D":      xCampos(4, 2) = "8"
    xCampos(5, 0) = "precio":        xCampos(5, 1) = "D":      xCampos(5, 2) = "15"
    
    xCampos(6, 0) = "numruc":        xCampos(6, 1) = "C":      xCampos(6, 2) = "11"

    
    Set RstTmp = xFun.CrearRstTMP(xCampos)
    RstTmp.Open
End Sub


