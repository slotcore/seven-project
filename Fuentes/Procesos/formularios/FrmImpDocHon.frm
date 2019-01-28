VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmImpDocHon 
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6015
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6735
      Top             =   -45
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
            Picture         =   "FrmImpDocHon.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDocHon.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1200
      Left            =   2745
      TabIndex        =   0
      Top             =   2775
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Documentos : "
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   1935
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
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   1
         X1              =   5790
         X2              =   5790
         Y1              =   15
         Y2              =   1155
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
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5820
         Y1              =   15
         Y2              =   0
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg2 
      Height          =   1950
      Left            =   0
      TabIndex        =   3
      Top             =   5415
      Width           =   11985
      _cx             =   21140
      _cy             =   3440
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmImpDocHon.frx":277E
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
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   2925
      Left            =   0
      TabIndex        =   36
      Top             =   2115
      Width           =   11985
      _cx             =   21140
      _cy             =   5159
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
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmImpDocHon.frx":284F
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
      Height          =   345
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   609
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
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   4
      Top             =   270
      Width           =   11985
      Begin VB.CommandButton CmdBusMes 
         Height          =   240
         Left            =   10230
         Picture         =   "FrmImpDocHon.frx":2A4E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   540
         Width           =   240
      End
      Begin VB.CommandButton CmdBusArch2 
         Height          =   240
         Left            =   7050
         Picture         =   "FrmImpDocHon.frx":2B80
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   540
         Width           =   240
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar"
         Height          =   615
         Left            =   10740
         TabIndex        =   10
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdBusArch 
         Height          =   240
         Left            =   7050
         Picture         =   "FrmImpDocHon.frx":2CB2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   240
      End
      Begin VB.CommandButton CmdBusImpTot 
         Enabled         =   0   'False
         Height          =   240
         Left            =   8475
         Picture         =   "FrmImpDocHon.frx":2DE4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1230
         Width           =   240
      End
      Begin VB.CommandButton CmdBusISC 
         Enabled         =   0   'False
         Height          =   240
         Left            =   2595
         Picture         =   "FrmImpDocHon.frx":2F16
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   915
         Width           =   240
      End
      Begin VB.CommandButton CmdBusIGV 
         Enabled         =   0   'False
         Height          =   240
         Left            =   2595
         Picture         =   "FrmImpDocHon.frx":3048
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1230
         Width           =   240
      End
      Begin VB.CommandButton CmdBusBruto 
         Enabled         =   0   'False
         Height          =   240
         Left            =   8460
         Picture         =   "FrmImpDocHon.frx":317A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   915
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox TxtArchivo 
         Height          =   300
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "TxtArchivo"
         Top             =   195
         Width           =   6015
      End
      Begin VB.TextBox TxtArchivo2 
         Height          =   300
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "TxtArchivo2"
         Top             =   510
         Width           =   6015
      End
      Begin VB.TextBox TxtCtaISC 
         Height          =   300
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "TxtCtaISC"
         Top             =   885
         Width           =   1560
      End
      Begin VB.TextBox TxtCtaIGV 
         Height          =   300
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "TxtCtaIGV"
         Top             =   1200
         Width           =   1560
      End
      Begin VB.TextBox TxtCtaTot 
         Height          =   300
         Left            =   7185
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "TxtCtaTot"
         Top             =   1200
         Width           =   1560
      End
      Begin VB.TextBox TxtCtaBru 
         Height          =   300
         Left            =   7170
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "TxtCtaBru"
         Top             =   885
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.TextBox TxtMes 
         Enabled         =   0   'False
         Height          =   300
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "TxtMes"
         Top             =   510
         Width           =   1740
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   8265
         TabIndex        =   35
         Top             =   540
         Width           =   300
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Arch. Cuentas"
         Height          =   195
         Left            =   135
         TabIndex        =   34
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Arch. Cabecera"
         Height          =   195
         Left            =   135
         TabIndex        =   33
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label LblIdMes 
         AutoSize        =   -1  'True
         Caption         =   "LblIdMes"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8175
         TabIndex        =   32
         Top             =   105
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label74 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label74"
         Height          =   300
         Left            =   8760
         TabIndex        =   31
         Top             =   1200
         Width           =   3105
      End
      Begin VB.Label Label72 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label72"
         Height          =   300
         Left            =   2880
         TabIndex        =   30
         Top             =   885
         Width           =   3105
      End
      Begin VB.Label Label73 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label73"
         Height          =   300
         Left            =   2880
         TabIndex        =   29
         Top             =   1200
         Width           =   3105
      End
      Begin VB.Label Label71 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label71"
         Height          =   300
         Left            =   8895
         TabIndex        =   28
         Top             =   885
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Importe Total"
         Height          =   195
         Left            =   6150
         TabIndex        =   27
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Importe Bruto"
         Height          =   195
         Left            =   6195
         TabIndex        =   26
         Top             =   915
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe IGV"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1230
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Importe I.S.C."
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   915
         Width           =   960
      End
      Begin VB.Label LblCtaIGV 
         AutoSize        =   -1  'True
         Caption         =   "LblCtaIGV"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7350
         TabIndex        =   23
         Top             =   465
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label LblIdCtaBru 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCtaBru"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7350
         TabIndex        =   22
         Top             =   690
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label LblIdCtaImpTot 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCtaImpTot"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   9465
         TabIndex        =   21
         Top             =   225
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label LblIdCtaISC 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCtaISC"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8340
         TabIndex        =   20
         Top             =   285
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "[ Documentos de Compra ]"
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
      Top             =   1875
      Width           =   2265
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "[ Asientos Contables ]"
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
      TabIndex        =   37
      Top             =   5175
      Width           =   1875
   End
End
Attribute VB_Name = "FrmImpDocHon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim RstTmp As New ADODB.Recordset
Dim QueHace As Integer

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

Sub CargaDocumentos()
    Dim xNumFilas As Integer
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro cabecera
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo.Text)
    
    Frame2.Left = 3090
    Frame2.Top = 2910
    Label4.Caption = "Cargando registros para la importacion"
    Frame2.Visible = True
    
    xFilas = 3
    xNumFilas = 1
    
    Fg1.Rows = 1
    Fg2.Rows = 1
    With objExcel.ActiveSheet
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        For A = 2 To 10000
            If NulosC(.Cells(A, 1)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit For
            End If
        Next A
        
        xNumFilas = xNumFilas + 1
        ProgressBar2.Max = xNumFilas
        
        For A = 2 To xNumFilas
            ProgressBar2.Value = A
            Frame2.Refresh
            
            If NulosC(.Cells(A, 1)) = "" Then Exit For
            Fg1.Rows = Fg1.Rows + 1
            For B = 1 To 14
                Fg1.TextMatrix(A - 1, B) = Trim(.Cells(A, B))
                If B = 4 Then
                    If Len(Trim(.Cells(A - 1, B))) > 25 Then
                        MsgBox "El numero de caracteres del documento de compra no debe de exceder los 14 digitos, error en fila :" & Str(A) & " columna : " & Str(B), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                        Set RstTmp = Nothing
                        Fg1.Rows = 1
                        Fg2.Rows = 1
                        Frame2.Visible = False
                        Exit Sub
                    End If
                    'MsgBox Str(Len(Trim(.Cells(A - 1, B))))
                End If
                If NulosC(.Cells(A, B)) = "" Then
                    MsgBox "La celda de a fila " + Trim(Str(A)) + " en la columna " + Trim(Str(B)) + " no contiene datos en el archivo " & Chr(13) _
                        & Trim(TxtArchivo.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Frame2.Visible = False
                    Exit Sub
                End If
                If (B = 9) Or (B = 10) Or (B = 11) Or (B = 12) Or (B = 13) Or (B = 14) Then Fg1.TextMatrix(A - 1, B) = Format(Trim(.Cells(A, B)), "0.00")
                If (B = 2) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(Val(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_documento", "N", xCon)
                If (B = 3) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(Val(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_tipoproducto", "N", xCon)
                If (B = 5) Or (B = 6) Then Fg1.TextMatrix(A - 1, B) = Format(CDate(Trim(.Cells(A, B))), "dd/mm/yy")
                If (B = 7) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(Val(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_condpago", "N", xCon)
                If (B = 8) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(Val(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_moneda", "N", xCon)
                'If (B = 15) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(Val(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_moneda", "N", xCon)
            Next B
            
        Next A
    End With
    
    
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo2.Text)

    xFilas = 3
    xNumFilas = 1
    PreparaRstTmp

    Fg2.Rows = 1
    With objExcel.ActiveSheet
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        Label4.Caption = "Calculando Nº Registros"
        Frame2.Refresh
        ProgressBar2.Max = 32000
        For A = 2 To 32000
            ProgressBar2.Value = A
            Frame2.Refresh
            If NulosC(.Cells(A, 1)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit For
            End If
        Next A

        xNumFilas = xNumFilas + 1
        Label4.Caption = "Cargando datos contables de la compra"
        Frame2.Refresh
        ProgressBar2.Max = xNumFilas

        For A = 2 To xNumFilas
            ProgressBar2.Value = A
            Frame2.Refresh

            If NulosC(.Cells(A, 1)) = "" Then Exit For
            Fg2.Rows = Fg2.Rows + 1

            RstTmp.AddNew
            For B = 1 To 5
                If B <> 3 Then
                    If NulosC(.Cells(A, B)) = "" Then
                        MsgBox "La celda de a fila " + Trim(Str(A)) + " en la columna " + Trim(Str(B)) + " no contiene datos en el archivo " & Chr(13) _
                            & Trim(TxtArchivo2.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                        Frame2.Visible = False
                        Exit Sub
                    End If
                End If

                Fg2.TextMatrix(A - 1, B) = Trim(.Cells(A, B))
                If B = 3 Then Fg2.TextMatrix(A - 1, B) = Busca_Codigo(Fg2.TextMatrix(A - 1, B - 1), "cuenta", "descripcion", "con_planctas", "C", xCon)
                If (B = 4) Or (B = 5) Then Fg2.TextMatrix(A - 1, B) = Format(Trim(.Cells(A, B)), "0.00")
                If B = 1 Then
                    If Len(Trim(.Cells(A - 1, B))) > 25 Then
                        MsgBox "El numero de caracteres del documento de compra no debe de exceder los 14 digitos, error en  fila " & Str(A - 1) & " columna " & Str(B), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                        Set RstTmp = Nothing
                        Fg1.Rows = 1
                        Fg2.Rows = 1
                        Frame2.Visible = False
                        Exit Sub
                    End If
                End If
                If B = 1 Then RstTmp("numdoc") = Trim(.Cells(A - 1, B))
                If B = 2 Then RstTmp("numcue") = Trim(.Cells(A - 1, B))
                If B = 4 Then RstTmp("impdeb") = Val(.Cells(A - 1, B))
                If B = 5 Then RstTmp("imphab") = Val(.Cells(A - 1, B))
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

Private Sub CmdBusBruto_Click()
    Dim xfrm As New SGI2_funciones.formularios
    Dim rst As New ADODB.Recordset
    Set rst = xfrm.SelePlanCuentas(xCon)
    If rst.State = 1 Then
        If rst.RecordCount <> 0 Then
            TxtCtaBru.Text = Trim(rst("cuenta"))
            Label71.Caption = Trim(rst("descripcion"))
            LblIdCtaBru.Caption = Trim(rst("id"))
            TxtCtaISC.SetFocus
        End If
    End If
    Set xfrm = Nothing
End Sub

Private Sub CmdBusIGV_Click()
    Dim xfrm As New SGI2_funciones.formularios
    Dim rst As New ADODB.Recordset
    Set rst = xfrm.SelePlanCuentas(xCon)
    If rst.State = 1 Then
        If rst.RecordCount <> 0 Then
            TxtCtaIGV.Text = Trim(rst("cuenta"))
            Label73.Caption = Trim(rst("descripcion"))
            LblCtaIGV.Caption = Trim(rst("id"))
            TxtCtaIGV.SetFocus
        End If
    End If
    Set xfrm = Nothing
End Sub

Private Sub CmdBusImpTot_Click()
    Dim xfrm As New SGI2_funciones.formularios
    Dim rst As New ADODB.Recordset
    Set rst = xfrm.SelePlanCuentas(xCon)
    If rst.State = 1 Then
        If rst.RecordCount <> 0 Then
            TxtCtaTot.Text = Trim(rst("cuenta"))
            Label74.Caption = Trim(rst("descripcion"))
            LblIdCtaImpTot.Caption = Trim(rst("id"))
            CmdCargar.SetFocus
        End If
    End If
    Set xfrm = Nothing
End Sub

Private Sub CmdBusISC_Click()
    Dim xfrm As New SGI2_funciones.formularios
    Dim rst As New ADODB.Recordset
    Set rst = xfrm.SelePlanCuentas(xCon)
    If rst.State = 1 Then
        If rst.RecordCount <> 0 Then
            TxtCtaISC.Text = Trim(rst("cuenta"))
            Label72.Caption = Trim(rst("descripcion"))
            LblIdCtaISC.Caption = Trim(rst("id"))
        End If
    End If
    Set xfrm = Nothing
End Sub

Private Sub CmdBusMes_Click()
    'Dim xform As New eps_librerias.FormBuscar
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
            RST_Busq xRs2, "SELECT * FROM var_importados WHERE idtabla = 2 AND idmes = " & xRs("id") & "", xCon
            If xRs2.RecordCount <> 0 Then
                MsgBox "Ya se importaron datos para el mes seleccionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                LblIdMes.Caption = ""
                TxtMes.Text = ""
                Exit Sub
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCargar_Click()
    If QueHace = 2 Then
        If TxtArchivo.Text = "" Then
            MsgBox "No ha especificado el archivo cabecera de la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtArchivo.SetFocus
            Exit Sub
        End If
        
        If TxtArchivo2.Text = "" Then
            MsgBox "No ha especificado el archivo detalle de la compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtArchivo2.SetFocus
            Exit Sub
        End If
        
        CargaDocumentos
    End If
    
    If QueHace = 3 Then
        If NulosC(TxtMes.Text) = "" Then
            MsgBox "No ha especificado el mes a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        CargarGrabado
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
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

Sub Blanquea()
    TxtArchivo.Text = ""
    TxtArchivo2.Text = ""
    TxtMes.Text = ""
    
    TxtCtaBru.Text = ""
    Label71.Caption = ""
    Label72.Caption = ""
    Label73.Caption = ""
    Label74.Caption = ""
    
    TxtCtaISC.Text = ""
    TxtCtaIGV.Text = ""
    TxtCtaTot.Text = ""
    
    LblIdMes.Caption = ""
End Sub

Sub Cancelar()
    Blanquea
    Bloquea
    ActivarTool
    QueHace = 3
End Sub

Sub ActivarTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(4).Enabled = Not Toolbar1.Buttons(4).Enabled
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(7).Enabled = Not Toolbar1.Buttons(7).Enabled
End Sub

Sub Bloquea()
    CmdBusArch.Enabled = Not CmdBusArch.Enabled
    'CmdBusArch2.Enabled = Not CmdBusArch2.Enabled
    
    CmdBusBruto.Enabled = Not CmdBusBruto.Enabled
    CmdBusISC.Enabled = Not CmdBusISC.Enabled
    CmdBusIGV.Enabled = Not CmdBusIGV.Enabled
    CmdBusImpTot.Enabled = Not CmdBusImpTot.Enabled
    
    'TxtCtaBru.Enabled = Not TxtCtaBru.Enabled
    'TxtCtaISC.Enabled = Not TxtCtaISC.Enabled
    'TxtCtaIGV.Enabled = Not TxtCtaIGV.Enabled
    'TxtCtaTot.Enabled = Not TxtCtaTot.Enabled
End Sub

Sub CargarGrabado()
    Dim rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq rst, "SELECT mae_prov.numruc, mae_tipoproducto.descripcion AS desctippro, mae_documento.descripcion AS desdoc, com_compras.fchdoc, " _
        & " com_compras.fchven, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, mae_condpago.abrev AS conpag, mae_moneda.simbolo, " _
        & " com_compras.impbru, com_compras.impina, com_compras.impisc, com_compras.impigv, com_compras.imptot, com_compras.impsal " _
        & " FROM mae_tipoproducto RIGHT JOIN (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_condpago RIGHT JOIN (mae_documento RIGHT JOIN " _
        & " (var_importados LEFT JOIN com_compras ON var_importados.iddoc = com_compras.id) ON mae_documento.id = com_compras.tipdoc) " _
        & " ON mae_condpago.id = com_compras.idconpag) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) ON " _
        & " mae_tipoproducto.id = com_compras.idtipo Where (((var_importados.idtabla) = 2) And ((var_importados.idmes) = " & Val(LblIdMes.Caption) & "))" _
        & " ORDER BY [com_compras]![numser]+'-'+[com_compras]![numdoc]", xCon
    
    Fg1.Rows = 1
    Fg2.Rows = 1
    
    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = rst("numruc")
            Fg1.TextMatrix(A, 2) = rst("desdoc")
            Fg1.TextMatrix(A, 3) = rst("desctippro")
            Fg1.TextMatrix(A, 4) = rst("numdoc")
            Fg1.TextMatrix(A, 5) = Format(rst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(A, 6) = Format(rst("fchven"), "dd/mm/yy")
            Fg1.TextMatrix(A, 7) = rst("conpag")
            Fg1.TextMatrix(A, 8) = rst("simbolo")
            Fg1.TextMatrix(A, 9) = Format(rst("impbru"), "0.00")
            Fg1.TextMatrix(A, 10) = Format(rst("impina"), "0.00")
            Fg1.TextMatrix(A, 11) = Format(rst("impigv"), "0.00")
            Fg1.TextMatrix(A, 12) = Format(rst("impisc"), "0.00")
            Fg1.TextMatrix(A, 13) = Format(rst("imptot"), "0.00")
            Fg1.TextMatrix(A, 14) = Format(rst("impsal"), "0.00")
            
            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
    End If
    
    'cargamos las cuentas cargadas
    RST_Busq rst, "SELECT DISTINCT con_diario.idlib, con_diario.idmes, con_planctas.cuenta, con_planctas.descripcion, con_diario.impdebsol, " _
        & " con_diario.imphabsol, con_diario.impdebdol, con_diario.imphabdol, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.idmon " _
        & " FROM con_planctas RIGHT JOIN (com_compras RIGHT JOIN con_diario ON com_compras.id = con_diario.idmov) ON con_planctas.id = con_diario.idcue " _
        & " Where (((con_diario.idlib) = 1) And ((con_diario.idmes) = " & Val(LblIdMes.Caption) & ")) ORDER BY [com_compras]![numser]+'-'+[com_compras]![numdoc]", xCon

    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = rst("numdoc")
            Fg2.TextMatrix(A, 2) = rst("cuenta")
            Fg2.TextMatrix(A, 3) = rst("descripcion")
            If rst("idmon") = 1 Then
                Fg2.TextMatrix(A, 4) = Format(rst("impdebsol"), "0.00")
                Fg2.TextMatrix(A, 5) = Format(rst("imphabsol"), "0.00")
            Else
                Fg2.TextMatrix(A, 4) = Format(rst("impdebdol"), "0.00")
                Fg2.TextMatrix(A, 5) = Format(rst("imphabdol"), "0.00")
            End If
            
            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
    End If
End Sub

Sub PreparaRstTmp()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(4, 3) As String

    xCampos(0, 0) = "numdoc":        xCampos(0, 1) = "C":      xCampos(0, 2) = "25"
    xCampos(1, 0) = "numcue":        xCampos(1, 1) = "C":      xCampos(1, 2) = "50"
    xCampos(2, 0) = "impdeb":        xCampos(2, 1) = "D":      xCampos(2, 2) = "8"
    xCampos(3, 0) = "imphab":        xCampos(3, 1) = "D":      xCampos(3, 2) = "8"
    Set RstTmp = xFun.CrearRstTMP(xCampos)
    RstTmp.Open
End Sub

Sub Modificar()
    ActivarTool
    QueHace = 2
    Bloquea
    TxtArchivo.SetFocus
End Sub

Function Grabar() As Boolean
'    If TxtCtaBru.Text = "" Then
'        MsgBox "No ha especificado la cuenta contable para el importe bruto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtCtaBru.SetFocus
'        Grabar = False
'        Exit Function
'    End If
    
'    If TxtCtaISC.Text = "" Then
'        MsgBox "No ha especificado la cuenta contable para el impuesto Selectivo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtCtaISC.SetFocus
'        Grabar = False
'        Exit Function
'
'    End If
'
'    If TxtCtaIGV.Text = "" Then
'        MsgBox "No ha especificado la cuenta contable para el impuesto I.G.V.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtCtaIGV.SetFocus
'        Grabar = False
'        Exit Function
'    End If
'
'    If TxtCtaTot.Text = "" Then
'        MsgBox "No ha especificado la cuenta contable para el total del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtCtaTot.SetFocus
'        Grabar = False
'        Exit Function
'    End If
    
    If TxtMes.Text = "" Then
        MsgBox "No ha especificado a que mes se cargaran los documentos a importar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        'TxtMes.SetFocus
        Grabar = False
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado los documentos que se van a importar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmdCargar.SetFocus
        Grabar = False
        Exit Function
    End If
    
'    If Fg2.Rows = 1 Then
'        MsgBox "No ha especificado los documentos que se van a importar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        CmdCargar.SetFocus
'        Grabar = False
'        Exit Function
'    End If
    
    'procesamos que los proveedores estan registrados en la tabla de proveedores
    Dim A, B, xId, xCodMon As Integer
    Dim xNom, xNumAsiento As String
    
    Frame2.Left = 3090
    Frame2.Top = 2910
    Label4.Caption = "Verificando Datos"
    ProgressBar2.Max = Fg1.Rows - 1
    Frame2.Visible = True
    
    For A = 1 To Fg1.Rows - 1
        ProgressBar2.Value = A
        Frame2.Refresh
        
        xNom = Busca_Codigo(Fg1.TextMatrix(A, 1), "numruc", "nombre", "mae_prov", "C", xCon)
        If xNom = "" Then
            MsgBox "El Nº de R.U.C. :" + Fg1.TextMatrix(A, 1) + " no existe no se puede importar los registros", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Grabar = False
            Frame2.Visible = False
            Exit Function
        End If
    Next A
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim RstImp As New ADODB.Recordset
    Dim RstTC As New ADODB.Recordset
    Dim xNumDocErr As String
    Dim RstCta As New ADODB.Recordset
    
On Error GoTo LaCague

    xCon.BeginTrans
    
    Label4.Caption = "Importando Datos"
    ProgressBar2.Max = Fg1.Rows - 1
    Frame2.Refresh
    
    xMes = Val(LblIdMes.Caption)
    RST_Busq RstCab, "SELECT * FROM com_compras", xCon
    RST_Busq RstDia, "SELECT * FROM con_diario", xCon
    RST_Busq RstImp, "SELECT * FROM var_importados", xCon
    
    For A = 1 To Fg1.Rows - 1
        ProgressBar2.Value = A
        Frame2.Refresh
        xNumDocErr = Trim(Fg1.TextMatrix(A, 4))
        xId = HallaCodigoTabla("com_compras", xCon, "id")
        xNumAsiento = NuevoNumAsiento(1, xMes, xCon)
        
        RstCab.AddNew
        RstCab("id") = xId
        RstCab("idlib") = 1
        RstCab("idtipo") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 3)), "descripcion", "id", "mae_tipoproducto", "C", xCon) 'ok
        RstCab("tipdoc") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 2)), "descripcion", "id", "mae_documento", "C", xCon)    'ok
        RstCab("idpro") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 1)), "numruc", "id", "mae_prov", "C", xCon)               'ok
        
        'modifcamos para que no grabe el numero de serie
        RstCab("numser") = "" 'Mid(Fg1.TextMatrix(A, 4), 1, 4)                                                'ok
        RstCab("numdoc") = NulosC(Fg1.TextMatrix(A, 4)) 'Mid(Fg1.TextMatrix(A, 4), 6, 10)                                         'ok
        
        RstCab("fchreg") = CDate("01/" + Format(xMes, "00") + "/" + Trim(Str(Val(AnoTra))))                  'ok
        RstCab("fchdoc") = CDate(Fg1.TextMatrix(A, 5))                                                                    'ok
        RstCab("fchven") = CDate(Fg1.TextMatrix(A, 6))                                                                    'ok
        RstCab("idconpag") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 7)), "descripcion", "id", "mae_condpago", "C", xCon)   'ok
        
        xCodMon = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 8)), "descripcion", "id", "mae_moneda", "C", xCon)                'ok
        RstCab("idmon") = xCodMon
        
        RstCab("impbru") = Val(Fg1.TextMatrix(A, 9))             'ok
        RstCab("impina") = Val(Fg1.TextMatrix(A, 10))            'ok
        RstCab("impigv") = NulosN(Fg1.TextMatrix(A, 11))         'ok
        RstCab("impisc") = NulosN(Fg1.TextMatrix(A, 12))         'ok
        RstCab("imptot") = NulosN(Fg1.TextMatrix(A, 13))         'ok
        RstCab("impsal") = NulosN(Fg1.TextMatrix(A, 14))         'ok
        
        If NulosN(Fg1.TextMatrix(A, 9)) = 0 Then
            RstCab("afecto") = 0
        Else
            RstCab("afecto") = -1
        End If
        RstCab("numreg") = Format(xMes, "00") + Trim(xNumAsiento)
        RstCab("importado") = -1
        RstCab.Update
        
        
        Dim xTC As Double
        Set RstCta = Nothing
        RST_Busq RstCta, "SELECT con_planctas.cuenta, con_planctas.id, con_planctas.ctadesdeb, con_planctas.ctadeshab" _
            & " From con_planctas WHERE (((con_planctas.cuenta)='" & Fg1.TextMatrix(A, 15) & "'))", xCon

        Set RstTC = Nothing
       
        Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 5) & "')", xCon)
        If RstTC.RecordCount <> 0 Then
            xTC = RstTC("impven")
        End If
        
'        If RstCta.RecordCount <> 0 Then
'            'grabamos el importer bruto
'            RstDia.AddNew
'            RstDia("año") = AnoTra
'            RstDia("idmes") = xMes
'            RstDia("idlib") = 1
'            RstDia("idmov") = xId
'            RstDia("numasi") = xNumAsiento
'            RstDia("idcue") = NulosN(Busca_Codigo(Fg1.TextMatrix(A, 15), "cuenta", "id", "con_planctas", "C", xCon))   'Val(LblIdCtaBru.Caption)
'            RstDia("fchasi") = CDate("01/" + Format(xMes, "00") + "/" + Trim(Str(Val(AnoTra))))
'
'            If xCodMon = 1 Then
'                RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 9)):            RstDia("imphabsol") = 0
'                RstDia("impdebdol") = 0:                                       RstDia("imphabdol") = 0
'            Else
'                RstDia("impdebsol") = (NulosN(Fg1.TextMatrix(A, 9)) * xTC):    RstDia("imphabsol") = 0
'                RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 9)):            RstDia("imphabdol") = 0
'            End If
'
'            'grabamos el importer IGV
'            RstDia.AddNew
'            RstDia("año") = AnoTra
'            RstDia("idmes") = xMes
'            RstDia("idlib") = 1
'            RstDia("idmov") = xId
'            RstDia("numasi") = xNumAsiento
'            RstDia("idcue") = NulosN(LblCtaIGV.Caption)
'            RstDia("fchasi") = CDate("01/" + Format(xMes, "00") + "/" + Trim(Str(Val(AnoTra))))
'
'            If xCodMon = 1 Then
'                RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 11)):            RstDia("imphabsol") = 0
'                RstDia("impdebdol") = 0:                                        RstDia("imphabdol") = 0
'            Else
'                RstDia("impdebsol") = (NulosN(Fg1.TextMatrix(A, 11)) * xTC):    RstDia("imphabsol") = 0
'                RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 11)):            RstDia("imphabdol") = 0
'            End If
'
'            'grabamos el importer Total
'            RstDia.AddNew
'            RstDia("año") = AnoTra
'            RstDia("idmes") = xMes
'            RstDia("idlib") = 1
'            RstDia("idmov") = xId
'            RstDia("numasi") = xNumAsiento
'            RstDia("idcue") = NulosN(LblIdCtaImpTot.Caption)
'            RstDia("fchasi") = CDate("01/" + Format(xMes, "00") + "/" + Trim(Str(Val(AnoTra))))
'
'            If xCodMon = 1 Then
'                RstDia("impdebsol") = 0:                                       RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 13))
'                RstDia("impdebdol") = 0:                                       RstDia("imphabdol") = 0
'            Else
'                RstDia("impdebsol") = 0:                                       RstDia("imphabsol") = (NulosN(Fg1.TextMatrix(A, 13)) * xTC)
'                RstDia("impdebdol") = 0:                                       RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 13))
'            End If
'
'            'grabamos el destino DEBE
'            If NulosN(RstCta("ctadesdeb")) <> 0 Then
'                RstDia.AddNew
'                RstDia("año") = AnoTra
'                RstDia("idmes") = xMes
'                RstDia("idlib") = 1
'                RstDia("idmov") = xId
'                RstDia("numasi") = xNumAsiento
'                RstDia("idcue") = NulosN(RstCta("ctadesdeb"))
'                RstDia("fchasi") = CDate("01/" + Format(xMes, "00") + "/" + Trim(Str(Val(AnoTra))))
'
'                If xCodMon = 1 Then
'                    RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 9)):            RstDia("imphabsol") = 0
'                    RstDia("impdebdol") = 0:                                       RstDia("imphabdol") = 0
'                Else
'                    RstDia("impdebsol") = (NulosN(Fg1.TextMatrix(A, 9)) * xTC):    RstDia("imphabsol") = 0
'                    RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 9)):            RstDia("imphabdol") = 0
'                End If
'            End If
'            'grabamos el destino HABER
'            If NulosN(RstCta("ctadeshab")) <> 0 Then
'                RstDia.AddNew
'                RstDia("año") = AnoTra
'                RstDia("idmes") = xMes
'                RstDia("idlib") = 1
'                RstDia("idmov") = xId
'                RstDia("numasi") = xNumAsiento
'                RstDia("idcue") = NulosN(RstCta("ctadeshab"))
'                RstDia("fchasi") = CDate("01/" + Format(xMes, "00") + "/" + Trim(Str(Val(AnoTra))))
'
'                If xCodMon = 1 Then
'                    RstDia("impdebsol") = 0:                                       RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 9))
'                    RstDia("impdebdol") = 0:                                       RstDia("imphabdol") = 0
'                Else
'                    RstDia("impdebsol") = 0:                                       RstDia("imphabsol") = (NulosN(Fg1.TextMatrix(A, 9)) * xTC)
'                    RstDia("impdebdol") = 0:                                       RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 9))
'                End If
'            End If
'
''                RstDia.AddNew
''                RstDia("año") = AnoTra
''                RstDia("idmes") = xMes
''                RstDia("idlib") = 1
''                RstDia("idmov") = xId
''                RstDia("numasi") = xNumAsiento
''                RstDia("idcue") = NulosN(Busca_Codigo(RstTmp("numcue"), "cuenta", "id", "con_planctas", "C", xCon)) 'Val(LblIdCtaBru.Caption)
''                RstDia("fchasi") = CDate("01/" + Format(xMes, "00") + "/" + Trim(Str(Val(AnoTra))))
''
''                If xCodMon = 1 Then
''                    RstDia("impdebsol") = RstTmp("impdeb")
''                    RstDia("imphabsol") = RstTmp("imphab")
''                    RstDia("impdebdol") = 0
''                    RstDia("imphabdol") = 0
''                Else
''                    Set RstTC = Nothing
''                    Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 5) & "')", xCon)
''                    If RstTC.RecordCount <> 0 Then
''                        RstDia("tc") = RstTC("impven")
''                        RstDia("impdebsol") = NulosN(RstTmp("impdeb")) * RstTC("impven")
''                        RstDia("imphabsol") = NulosN(RstTmp("imphab")) * RstTC("impven")
''                        RstDia("impdebdol") = RstTmp("impdeb")
''                        RstDia("imphabdol") = RstTmp("imphab")
''                    Else
''                        RstDia("tc") = 0
''                        RstDia("impdebsol") = 0
''                        RstDia("imphabsol") = 0
''                        RstDia("impdebdol") = RstTmp("impdeb")
''                        RstDia("imphabdol") = RstTmp("imphab")
''                    End If
''                End If
''                RstDia.Update
'
'        End If

        'GRABAMOS EL DIARIO DEL MOVIMIENTO
        RstTmp.Filter = adFilterNone
        RstTmp.Filter = "numdoc = '" & Fg1.TextMatrix(A, 4) & "'"
        If RstTmp.RecordCount <> 0 Then
            RstTmp.MoveFirst
            For B = 1 To RstTmp.RecordCount
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = xMes
                RstDia("idlib") = 1
                RstDia("idmov") = xId
                RstDia("numasi") = xNumAsiento
                RstDia("idcue") = NulosN(Busca_Codigo(RstTmp("numcue"), "cuenta", "id", "con_planctas", "C", xCon)) 'Val(LblIdCtaBru.Caption)
                RstDia("fchasi") = CDate("01/" + Format(xMes, "00") + "/" + Trim(Str(Val(AnoTra))))

                If xCodMon = 1 Then
                    RstDia("impdebsol") = RstTmp("impdeb")
                    RstDia("imphabsol") = RstTmp("imphab")
                    RstDia("impdebdol") = 0
                    RstDia("imphabdol") = 0
                Else
                    Set RstTC = Nothing
                    Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 5) & "')", xCon)
                    If RstTC.RecordCount <> 0 Then
                        RstDia("tc") = RstTC("impven")
                        RstDia("impdebsol") = NulosN(RstTmp("impdeb")) * RstTC("impven")
                        RstDia("imphabsol") = NulosN(RstTmp("imphab")) * RstTC("impven")
                        RstDia("impdebdol") = RstTmp("impdeb")
                        RstDia("imphabdol") = RstTmp("imphab")
                    Else
                        RstDia("tc") = 0
                        RstDia("impdebsol") = 0
                        RstDia("imphabsol") = 0
                        RstDia("impdebdol") = RstTmp("impdeb")
                        RstDia("imphabdol") = RstTmp("imphab")
                    End If
                End If
                RstDia.Update
                RstTmp.MoveNext
                If RstTmp.EOF = True Then Exit For
            Next B
        Else
            MsgBox "El Documento Nº " + Trim(Fg1.TextMatrix(A, 4)) + " del R.U.C. Nº" + Trim(Fg1.TextMatrix(A, 1)) + " no tiene asiento contable", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Frame2.Visible = False
            xCon.RollbackTrans
            Exit Function
        End If
        
        RstImp.AddNew
        RstImp("iddoc") = xId
        RstImp("idtabla") = 2
        RstImp("idmes") = xMes
        RstImp.Update
    Next A
    
    xCon.CommitTrans
    
    Frame2.Visible = False
    Fg1.Rows = 1
    Fg2.Rows = 1
    
    Set RstCab = Nothing
    Set RstDia = Nothing
    Set RstImp = Nothing
    MsgBox "Los documentos de compra se importaron con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
'    Resume
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el documento por el siguiente motivo: " + Trim(Err.Description) + Chr(13) _
        & "El documento con problemas es : " + Trim(xNumDocErr), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Frame2.Visible = False
    Set RstCab = Nothing
    Set RstDia = Nothing
    Set RstImp = Nothing
End Function

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
    
    RST_Busq rst, "SELECT * FROM var_importados WHERE idtabla = 2 AND idmes = " & MesBorrar & "", xCon
    
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
                & " WHERE (((con_diario.idmes)=" & MesBorrar & ") AND ((con_diario.idlib)=1) AND ((con_diario.idmov)=" & rst("iddoc") & "))"
            
            xCon.Execute "DELETE * FROM com_compras WHERE id = " & rst("iddoc") & ""

            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
        Frame2.Visible = False
        
        xCon.Execute "DELETE * FROM var_importados WHERE idtabla = 2 AND idmes = " & MesBorrar & ""
    End If
End Sub

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
            Cancelar
            CargarGrabado
        End If
    End If
    If Button.Index = 7 Then
        Unload Me
    End If
End Sub


