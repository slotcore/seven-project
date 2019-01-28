VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmPlanillas 
   Caption         =   "Sistema de Planillas - Ingreso de Planillas"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4590
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":0B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":0C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":16BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":1D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanillas.frx":1EA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Reportes"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s1"
                  Object.Tag             =   "1"
                  Text            =   "Lista de Productos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s2"
                  Object.Tag             =   "2"
                  Text            =   "Lista de Precios"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "s3"
                  Object.Tag             =   "3"
                  Text            =   "Productos sin Stock"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   6
      Top             =   375
      Width           =   11700
      _cx             =   20637
      _cy             =   12726
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   12632256
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   8421504
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "   Consulta   |   Detalles   "
      Align           =   0
      CurrTab         =   1
      FirstTab        =   0
      Style           =   0
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6795
         Left            =   45
         TabIndex        =   10
         Top             =   375
         Width           =   11610
         Begin VB.TextBox TxtTotPla 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   9360
            TabIndex        =   26
            Text            =   "TxtTotPla"
            Top             =   6315
            Width           =   1095
         End
         Begin VB.TextBox TxtTotApo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10485
            TabIndex        =   22
            Text            =   "TxtTotApo"
            Top             =   6315
            Width           =   1095
         End
         Begin VB.TextBox TxtTotDsct 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   8235
            TabIndex        =   21
            Text            =   "TxtTotDsct"
            Top             =   6315
            Width           =   1095
         End
         Begin VB.TextBox TxtTotBas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   7110
            TabIndex        =   20
            Text            =   "TxtTotBas"
            Top             =   6315
            Width           =   1095
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   4350
            Left            =   45
            TabIndex        =   19
            Top             =   1665
            Width           =   11535
            _cx             =   20346
            _cy             =   7673
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
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
         Begin VB.CommandButton CmdBusGrupo 
            Height          =   225
            Left            =   5100
            Picture         =   "FrmPlanillas.frx":23EC
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   735
            Width           =   240
         End
         Begin VB.TextBox TxtBusTipPla 
            Height          =   300
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   2
            Text            =   "TxtBusTipPla"
            Top             =   705
            Width           =   1230
         End
         Begin AspaTextBoxFecha.TextBoxFecha txtFchPro 
            Height          =   300
            Left            =   1320
            TabIndex        =   1
            Top             =   705
            Width           =   1230
            _ExtentX        =   2170
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
            Locked          =   -1  'True
            Valor           =   "10/10/2005"
         End
         Begin VB.TextBox TxtNumPla 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "TxtNumPla"
            Top             =   405
            Width           =   1230
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
            Height          =   300
            Left            =   1320
            TabIndex        =   3
            Top             =   1020
            Width           =   1230
            _ExtentX        =   2170
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
            Locked          =   -1  'True
            Valor           =   "10/10/2005"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
            Height          =   300
            Left            =   4140
            TabIndex        =   4
            Top             =   1020
            Width           =   1230
            _ExtentX        =   2170
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
            Locked          =   -1  'True
            Valor           =   "10/10/2005"
         End
         Begin VB.Frame Frame3 
            Height          =   810
            Left            =   45
            TabIndex        =   28
            Top             =   5970
            Width           =   6960
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Aportaciones del Emp."
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4455
               TabIndex        =   31
               Top             =   195
               Width           =   2205
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Descuentos"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1350
               TabIndex        =   30
               Top             =   480
               Width           =   1050
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Remuneracion"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1350
               TabIndex        =   29
               Top             =   195
               Width           =   1260
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H00C0C0FF&
               BackStyle       =   1  'Opaque
               Height          =   255
               Left            =   3300
               Top             =   180
               Width           =   945
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               Height          =   255
               Left            =   225
               Top             =   480
               Width           =   945
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00E7FEFC&
               BackStyle       =   1  'Opaque
               Height          =   255
               Left            =   225
               Top             =   180
               Width           =   945
            End
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Total Planilla"
            Height          =   195
            Left            =   9375
            TabIndex        =   27
            Top             =   6090
            Width           =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tot. Apor. Emp"
            Height          =   195
            Left            =   10500
            TabIndex        =   25
            Top             =   6090
            Width           =   1065
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total Dsct."
            Height          =   195
            Left            =   8280
            TabIndex        =   24
            Top             =   6090
            Width           =   780
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Basico"
            Height          =   195
            Left            =   7110
            TabIndex        =   23
            Top             =   6090
            Width           =   885
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Planilla"
            Height          =   195
            Left            =   3090
            TabIndex        =   17
            Top             =   750
            Width           =   855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Nomina de Empleados"
            Height          =   225
            Left            =   60
            TabIndex        =   16
            Top             =   1410
            Width           =   1590
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Final"
            Height          =   195
            Left            =   3090
            TabIndex        =   15
            Top             =   1050
            Width           =   690
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Inicio"
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   1050
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Proceso"
            Height          =   195
            Left            =   60
            TabIndex        =   13
            Top             =   750
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Planilla"
            Height          =   195
            Left            =   60
            TabIndex        =   12
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Planilla"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   90
            TabIndex        =   11
            Top             =   60
            Width           =   11400
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   -12255
         TabIndex        =   7
         Top             =   375
         Width           =   11610
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6405
            Left            =   45
            TabIndex        =   8
            Top             =   390
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   11298
            _LayoutType     =   4
            _RowHeight      =   14
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Planilla"
            Columns(0).DataField=   "numpla"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tipo Planilla"
            Columns(1).DataField=   "destippla"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Proceso"
            Columns(2).DataField=   "fchpro"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Inicio"
            Columns(3).DataField=   "fchini"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Final"
            Columns(4).DataField=   "fchfin"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Tot. Basico"
            Columns(5).DataField=   "totbas"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Tot. Dsct."
            Columns(6).DataField=   "totdes"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Apo. Emp."
            Columns(7).DataField=   "totapoemp"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2328"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2249"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2143"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2064"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2275"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2196"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2223"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2143"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2196"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2117"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=2064"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1984"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE7FEFC&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.fgcolor=&HFF0000&,.bold=-1"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38,.bgcolor=&H80&"
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
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14,.alignment=2"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=2"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14,.alignment=2"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14,.alignment=2"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14,.alignment=2"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14,.alignment=2"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Named:id=33:Normal"
            _StyleDefs(69)  =   ":id=33,.parent=0"
            _StyleDefs(70)  =   "Named:id=34:Heading"
            _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   ":id=34,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=35:Footing"
            _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=36:Selected"
            _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=37:Caption"
            _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(79)  =   "Named:id=38:HighlightRow"
            _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=39:EvenRow"
            _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(83)  =   "Named:id=40:OddRow"
            _StyleDefs(84)  =   ":id=40,.parent=33"
            _StyleDefs(85)  =   "Named:id=41:RecordSelector"
            _StyleDefs(86)  =   ":id=41,.parent=34"
            _StyleDefs(87)  =   "Named:id=42:FilterBar"
            _StyleDefs(88)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Planillas"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   90
            TabIndex        =   9
            Top             =   60
            Width           =   11400
         End
      End
   End
End
Attribute VB_Name = "FrmPlanillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstCab As New ADODB.Recordset
Dim SeEjecuto As Boolean

Dim Rpta As Integer
Dim QueHace As Integer

Dim TipoPlanilla As Integer
Dim RstNomina As New ADODB.Recordset
Dim ColTotaBon As Integer

Private Sub CmdBusGrupo_Click()
    If QueHace = 3 Then Exit Sub
    
    'Dim xform As New eps_librerias.FormBuscar
    Dim xform As New eps_librerias.FormBuscar
    
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "Descripcion":    xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT * FROM mae_tipopersonal"
    xform.Titulo = "Buscando Tipo de Planilla"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "Descripcion"
    xform.CampoBusca = "Descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtBusTipPla.Text = xRs("descripcion")
        TipoPlanilla = xRs("id")
        MuestraNomina
        TxtFchIni.SetFocus
    End If
    Set xRs = Nothing
    Set xform = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        TabOne1.CurrTab = 0
        SeEjecuto = True
        
        RST_Busq RstCab, "SELECT pla_planillas.*, pla_tipopersonal.descripcion AS destippla " _
            & " FROM pla_tipopersonal RIGHT JOIN pla_planillas ON pla_tipopersonal.id = pla_planillas.tippla" _
            & " ORDER BY fchini DESC", xCon
        
        If RstCab.RecordCount = 0 Then
            Rpta = MsgBox("El registro de planillas esta vacio ¿Desea agregar una planilla ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                nuevo
                Exit Sub
            Else
                Set RstCab = Nothing
                Unload Me
                Exit Sub
            End If
        End If
        Dg1.DataSource = RstCab
    End If
End Sub


Sub MuestraPlanilla()
    TxtNumPla.Text = RstCab("numpla")
    txtFchPro.Valor = RstCab("fchpro")
    TxtBusTipPla.Text = Busca_Codigo(RstCab("tippla"), "id", "descripcion", "m_tipopersonal", "N", xCon)
    TxtFchIni.Valor = RstCab("fchini")
    TxtFchFin.Valor = RstCab("fchfin")
    
    TxtTotBas.Text = Format(RstCab("totbas"), "0.00")
    TxtTotDsct.Text = Format(RstCab("totdes"), "0.00")
    TxtTotApo.Text = Format(RstCab("totapoemp"), "0.00")
    TxtTotPla.Text = Format(RstCab("totpla"), "0.00")
    
    TipoPlanilla = RstCab("tippla")
    MostrarPlanillaGuardadas
    
End Sub

Sub MostrarPlanillaGuardadas()
    Dim RstBon As New ADODB.Recordset   'BONIFICACION DEL EMPLEADO
    Dim RstDsc As New ADODB.Recordset   'DESCUENTOS DEL EMPLEADO
    Dim RstEmp As New ADODB.Recordset   'APORTES DEL EMPLEADOR
    
    Dim RstNom As New ADODB.Recordset
    Dim A As Integer
    Dim RstCom As New ADODB.Recordset
    
    
    'MOSTRAMOS LA NOMINA DE LA PLANILLA SELECCIONADA
    RST_Busq RstNomina, "SELECT DISTINCT pl_planillasdet.idpla, pl_planillasdet.idnom, " _
        & " UCase([pl_nomina]![ape])+', '+[pl_nomina]![nom] AS apenom, pl_nomina.id" _
        & " FROM pl_planillasdet LEFT JOIN pl_nomina ON pl_planillasdet.idnom = pl_nomina.id " _
        & " WHERE (((pl_planillasdet.idpla)=" & RstCab("id") & ")) ORDER BY pl_planillasdet.idnom", xCon
        
    LlenaGrid
    
    RST_Busq RstCom, "SELECT DISTINCT pl_planillasdet.idpla, pl_planillasdet.idcom, pl_planillasdet.tipo, " _
        & " pl_conceptos.abrev FROM pl_planillasdet LEFT JOIN pl_conceptos ON " _
        & " pl_planillasdet.idcom = pl_conceptos.id", xCon

   
'    RST_Busq RstBon, "SELECT pl_planillasdet.idpla, pl_planillasdet.idnom, pl_conceptos.orden, " _
'        & " UCase([pl_nomina]![ape])+', '+[pl_nomina]![nom] AS apenom, pl_conceptos.abrev, " _
'        & " pl_planillasdet.import, pl_planillasdet.porcen, pl_planillasdet.tipo, pl_planillasdet.idcom " _
'        & " FROM (pl_planillasdet LEFT JOIN pl_nomina ON pl_planillasdet.idnom = pl_nomina.id) " _
'        & " LEFT JOIN pl_conceptos ON pl_planillasdet.idcom = pl_conceptos.id " _
'        & " Where (((pl_planillasdet.idpla) = " & RstCab("id") & ") And ((pl_planillasdet.tipo) = 1)) " _
'        & " ORDER BY pl_planillasdet.idnom, pl_conceptos.orden", xcon
    
    RST_Busq RstBon, "SELECT pl_planillasdet.idpla, pl_planillasdet.idnom, pl_conceptos.orden, " _
        & " UCase([pl_nomina]![ape])+', '+[pl_nomina]![nom] AS apenom, pl_conceptos.abrev, " _
        & " pl_planillasdet.import, pl_planillasdet.porcen, pl_planillasdet.tipo, pl_planillasdet.idcom " _
        & " FROM (pl_planillasdet LEFT JOIN pl_nomina ON pl_planillasdet.idnom = pl_nomina.id) " _
        & " LEFT JOIN pl_conceptos ON pl_planillasdet.idcom = pl_conceptos.id " _
        & " Where (((pl_planillasdet.idpla) = " & RstCab("id") & ")) " _
        & " ORDER BY pl_planillasdet.idnom, pl_conceptos.orden", xCon
    
    Dim B As Integer
    
    'recorremos todos las remuneraciones del los empleados y creamos sus columnas
    RstCom.Filter = "tipo = 1"
    RstCom.MoveFirst
    For A = 1 To RstCom.RecordCount
        Fg1.Cols = Fg1.Cols + 1
        Fg1.ColWidth(Fg1.Cols - 1) = 800
        Fg1.TextMatrix(0, Fg1.Cols - 1) = RstCom("abrev")
        
        RstCom.MoveNext
        If RstCom.EOF = True Then
            Exit For
        End If
    Next A
    
    'PINTAMOS LAS REMUNERACIONES
    With Fg1
        .Select 1, 2, Fg1.Rows - 1, Fg1.Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &HDDFFFF
    End With
    
    'AGREGAMOS LA COLUMNA PARA EL TOTAL DE BONIFICACIONES
    Fg1.Cols = Fg1.Cols + 1
    Fg1.ColWidth(Fg1.Cols - 1) = 800
    Fg1.TextMatrix(0, Fg1.Cols - 1) = "Tot. Rem."
        
    With Fg1
        .Select 1, Fg1.Cols - 1, Fg1.Rows - 1, Fg1.Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &HEBD7BC
    End With
    
    RstCom.MoveFirst
    Dim xNumCol As Integer
    Dim Total As Double
    Dim SubTot As Double
    xNumCol = 2
    
    'filtramos las remuneraciones
    RstBon.Filter = "tipo = 1"
    
    For A = 1 To RstCom.RecordCount
        'filtramos solo el concepto actual
        RstBon.Filter = "idcom = " & RstCom("idcom") & ""
        If RstBon.RecordCount <> 0 Then
            RstBon.MoveFirst
            SubTot = 0
            For B = 1 To RstBon.RecordCount
                Fg1.TextMatrix(B, xNumCol) = Format(RstBon("import"), "0.00")
                Fg1.TextMatrix(B, Fg1.Cols - 1) = Val(Fg1.TextMatrix(B, Fg1.Cols - 1)) + RstBon("import")
                Fg1.TextMatrix(B, Fg1.Cols - 1) = Format(Fg1.TextMatrix(B, Fg1.Cols - 1), "0.00")
                SubTot = SubTot + RstBon("import")
                RstBon.MoveNext
                If RstBon.EOF = True Then
                    Exit For
                End If
            Next B
        End If
        RstCom.MoveNext
        xNumCol = xNumCol + 1
        Total = Total + SubTot
        If RstCom.EOF = True Then
            Exit For
        End If
    Next A
    TxtTotBas.Text = Format(Total, "0.00")
    Dim UltimaCol As Integer
    UltimaCol = Fg1.Cols - 1
    
    
    '**********************************************************************
    'recorremos todos los descuento del los empleado y creamos sus columnas
    Total = 0
    RstCom.Filter = "tipo = 2"
    RstCom.MoveFirst
    For A = 1 To RstCom.RecordCount
        Fg1.Cols = Fg1.Cols + 1
        Fg1.ColWidth(Fg1.Cols - 1) = 800
        Fg1.TextMatrix(0, Fg1.Cols - 1) = RstCom("abrev")
        
        RstCom.MoveNext
        If RstCom.EOF = True Then
            Exit For
        End If
    Next A
    
    'PINTAMOS LOS DESCUENTOS
    With Fg1
        .Select 1, UltimaCol + 1, Fg1.Rows - 1, Fg1.Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &HC0E0FF
    End With
    
    'AGREGAMOS LA COLUMNA PARA EL TOTAL DE DESCUENTOS
    Fg1.Cols = Fg1.Cols + 1
    Fg1.ColWidth(Fg1.Cols - 1) = 800
    Fg1.TextMatrix(0, Fg1.Cols - 1) = "Tot. Dsct."
        
    With Fg1
        .Select 1, Fg1.Cols - 1, Fg1.Rows - 1, Fg1.Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &HEBD7BC
    End With
    
    RstCom.MoveFirst

    xNumCol = UltimaCol + 1
    'filtramos las remuneraciones
    RstBon.Filter = adFilterNone
    RstBon.Filter = "tipo = 2"
    
    For A = 1 To RstCom.RecordCount
        'filtramos solo el concepto actual
        RstBon.Filter = "idcom = " & RstCom("idcom") & ""
        If RstBon.RecordCount <> 0 Then
            RstBon.MoveFirst
            SubTot = 0
            For B = 1 To RstBon.RecordCount
                Fg1.TextMatrix(B, xNumCol) = Format(RstBon("import"), "0.00")
                Fg1.TextMatrix(B, Fg1.Cols - 1) = Val(Fg1.TextMatrix(B, Fg1.Cols - 1)) + RstBon("import")
                Fg1.TextMatrix(B, Fg1.Cols - 1) = Format(Fg1.TextMatrix(B, Fg1.Cols - 1), "0.00")
                SubTot = SubTot + RstBon("import")
                RstBon.MoveNext
                If RstBon.EOF = True Then
                    Exit For
                End If
            Next B
        End If
        RstCom.MoveNext
        xNumCol = xNumCol + 1
        Total = Total + SubTot
        If RstCom.EOF = True Then
            Exit For
        End If
    Next A
    TxtTotDsct.Text = Format(Total, "0.00")


    UltimaCol = Fg1.Cols - 1
    
    
    '**********************************************************************
    'recorremos todos los aportes del empleador y creamos sus columnas
    Total = 0
    RstCom.Filter = "tipo = 3"
    RstCom.MoveFirst
    For A = 1 To RstCom.RecordCount
        Fg1.Cols = Fg1.Cols + 1
        Fg1.ColWidth(Fg1.Cols - 1) = 800
        Fg1.TextMatrix(0, Fg1.Cols - 1) = RstCom("abrev")
        
        RstCom.MoveNext
        If RstCom.EOF = True Then
            Exit For
        End If
    Next A
    
    'PINTAMOS LOS APORTES DEL EMPLEADOR
    With Fg1
        .Select 1, UltimaCol + 1, Fg1.Rows - 1, Fg1.Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &HC0C0FF
    End With
    
    'AGREGAMOS LA COLUMNA PARA EL TOTAL DE APORTES DEL EMPLEADOR
    Fg1.Cols = Fg1.Cols + 1
    Fg1.ColWidth(Fg1.Cols - 1) = 800
    Fg1.TextMatrix(0, Fg1.Cols - 1) = "Tot. Apo. Emp."
        
    With Fg1
        .Select 1, Fg1.Cols - 1, Fg1.Rows - 1, Fg1.Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = &HEBD7BC
    End With
    
    RstCom.MoveFirst

    xNumCol = UltimaCol + 1
    'filtramos las remuneraciones
    RstBon.Filter = adFilterNone
    RstBon.Filter = "tipo = 3"
    
    For A = 1 To RstCom.RecordCount
        'filtramos solo el concepto actual
        RstBon.Filter = "idcom = " & RstCom("idcom") & ""
        If RstBon.RecordCount <> 0 Then
            RstBon.MoveFirst
            SubTot = 0
            For B = 1 To RstBon.RecordCount
                Fg1.TextMatrix(B, xNumCol) = Format(RstBon("import"), "0.00")
                Fg1.TextMatrix(B, Fg1.Cols - 1) = Val(Fg1.TextMatrix(B, Fg1.Cols - 1)) + RstBon("import")
                Fg1.TextMatrix(B, Fg1.Cols - 1) = Format(Fg1.TextMatrix(B, Fg1.Cols - 1), "0.00")
                SubTot = SubTot + RstBon("import")
                RstBon.MoveNext
                If RstBon.EOF = True Then
                    Exit For
                End If
            Next B
        End If
        RstCom.MoveNext
        xNumCol = xNumCol + 1
        Total = Total + SubTot
        If RstCom.EOF = True Then
            Exit For
        End If
    Next A
    TxtTotApo.Text = Format(Total, "0.00")
    TxtTotPla.Text = Val(TxtTotBas.Text) - Val(TxtTotDsct.Text)
    TxtTotPla.Text = Format(TxtTotPla.Text, "0.00")
    
    Set RstBon = Nothing
    Set RstCom = Nothing
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    QueHace = 3
    IniciarGrid
    Toolbar1.Buttons(2).Visible = False
End Sub

Sub IniciarGrid()
    With Fg1
        .Rows = 1
        .Cols = 2
        .ColWidth(0) = 400
        .ColWidth(1) = 3000
        .RowHeight(0) = 300
        .TextMatrix(0, 1) = "    Apellidos y Nombres"
    End With
End Sub

Sub nuevo()
    QueHace = 1
    ActivaTool
    IniciarGrid
    TabOne1.CurrTab = 1
    Blanquea
    Bloquea
    TxtNumPla.SetFocus
End Sub

Sub MuestraNomina()
    RST_Busq RstNomina, "SELECT pla_empleados.*, pla_cargos.descripcion AS descarg, [ape]+', '+[nom] AS apenom" _
        & " FROM pla_cargos INNER JOIN pla_empleados ON pla_cargos.id = pla_empleados.idcargo " _
        & " WHERE ((pla_empleados.cesado = 0) AND (pla_empleados.tiptra = " & TipoPlanilla & ") AND (tippla = 1))" _
        & " ORDER BY pla_empleados.id", xCon
    
    LlenaGrid
    ConfigurarGrid
End Sub

Sub LlenaGrid()
    Dim A As Integer
    If RstNomina.RecordCount <> 0 Then
        RstNomina.MoveFirst
            
        For A = 1 To RstNomina.RecordCount
            Fg1.AddItem Fg1.Rows + 1
            Fg1.TextMatrix(A, 0) = RstNomina("id")
            Fg1.TextMatrix(A, 1) = RstNomina("apenom")
            RstNomina.MoveNext
            
            If RstNomina.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Sub ConfigurarGrid()
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim xNumCol As Integer
    Dim A As Integer
    Dim xTotal As Double
    Dim xColDesc As Integer

    'MOSTRAMOS LOS APORTES DE LA PLANILLA
    RST_Busq Rst, "SELECT pla_conceptos.id, pla_conceptos.abrev, pla_conceptos.porcen, pla_conceptos.tipo, " _
        & " pla_conceptos.orden From pla_conceptos Where (((pla_conceptos.tipo) = 1) AND activo = -1) ORDER BY pla_conceptos.orden", xCon
    
    xNumCol = 2
    Rst.MoveFirst
    
    Dim B As Integer
    Dim ColBasico As Double
    
    'ColTotaBon = 2 + Rst.RecordCount
    For A = 1 To Rst.RecordCount
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, xNumCol) = Rst("abrev")
        
            For B = 1 To Fg1.Rows - 1
                RST_Busq Rst2, "SELECT pla_empleados.id, pla_nominacarg.idcom, pla_nominacarg.imptot, pla_nominacarg.porcen " _
                    & " FROM pla_empleados LEFT JOIN pla_nominacarg ON pla_empleados.id = pla_nominacarg.idnom " _
                    & " WHERE (((pla_empleados.id) = " & Fg1.TextMatrix(B, 0) & ") AND ((pla_nominacarg.idcom) = " & Rst("id") & "))", xCon
                If Rst2.RecordCount <> 0 Then
                    If Rst2("imptot") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format(Format(Rst2("imptot"), "0.00"), "0.00")
                    If Rst2("porcen") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format((Val(Fg1.TextMatrix(B, 2)) * (Rst2("porcen") / 100)), "0.00")
                End If
            Next B
        
'        'AGREGAMOS EL SUELDO BASICO DEL TRABAJADOR
'        If Rst("id") = 1 Then
'            For B = 1 To Fg1.Rows - 1
'                RST_Busq Rst2, "SELECT pl_nomina.id, pl_nominacarg.idcom, pl_nominacarg.imptot " _
'                    & " FROM pl_nomina LEFT JOIN pl_nominacarg ON pl_nomina.id = pl_nominacarg.idnom " _
'                    & " WHERE (((pl_nomina.id) = " & Fg1.TextMatrix(B, 0) & ") AND ((pl_nominacarg.idcom) = 1))", xcon
'                If Rst2.RecordCount <> 0 Then
'                    Fg1.TextMatrix(B, xNumCol) = Format(Rst2("imptot"), "0.00")
'                End If
'            Next B
'        End If
'
'        'AGREGAMOS LAS BONIFICACIONES ESPECIFALES
'        If Rst("id") = 8 Then
'
'            For B = 1 To Fg1.Rows - 1
'                RST_Busq Rst2, "SELECT pl_nomina.id, pl_nominacarg.idcom, pl_nominacarg.imptot, pl_nominacarg.porcen " _
'                    & " FROM pl_nomina LEFT JOIN pl_nominacarg ON pl_nomina.id = pl_nominacarg.idnom " _
'                    & " WHERE (((pl_nomina.id) = " & Fg1.TextMatrix(B, 0) & ") AND ((pl_nominacarg.idcom) = 8))", xcon
'                If Rst2.RecordCount <> 0 Then
'                    If Rst2("imptot") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format(Format(Rst2("imptot"), "0.00"), "0.00")
'                    If Rst2("porcen") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format((Val(Fg1.TextMatrix(B, 2)) * (Rst2("porcen") / 100)), "0.00")
'                End If
'            Next B
'        End If
'
'        'AGREGAMOS LA ASIGNACION FAMILIAR
'        If Rst("id") = 3 Then
'            For B = 1 To Fg1.Rows - 1
'                RST_Busq Rst2, "SELECT pl_nomina.id, pl_nominacarg.idcom, pl_nominacarg.imptot, pl_nominacarg.porcen " _
'                    & " FROM pl_nomina LEFT JOIN pl_nominacarg ON pl_nomina.id = pl_nominacarg.idnom " _
'                    & " WHERE (((pl_nomina.id) = " & Fg1.TextMatrix(B, 0) & ") AND ((pl_nominacarg.idcom) = 3))", xcon
'                If Rst2.RecordCount <> 0 Then
'                    If Rst2("imptot") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format(Format(Rst2("imptot"), "0.00"), "0.00")
'                    If Rst2("porcen") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format((Val(Fg1.TextMatrix(B, 2)) * (Rst2("porcen") / 100)), "0.00")
'                End If
'            Next B
'        End If
        
        
        Fg1.ColWidth(xNumCol) = 800
        'Fg1.TextMatrix(2, xNumCol) = "25000.00"
        Fg1.ColAlignment(xNumCol) = flexAlignRightCenter
        Rst.MoveNext
        
        If Rst.EOF = True Then
            Exit For
        End If
        xNumCol = xNumCol + 1
    Next A
    
    'PINTAMOS DE COLOR LAS COLUMNAS DE LA REMUNERACION
    With Fg1
        .Select 1, 2, Fg1.Rows - 1, xNumCol
        .FillStyle = flexFillRepeat
        .CellBackColor = &HDDFFFF
    End With
    
    'AGREGAMOS LA COLUMNA PARA EL TOTAL DE BONIFICACIONES
    xNumCol = xNumCol + 1
    Fg1.Cols = Fg1.Cols + 1
    Fg1.TextMatrix(0, xNumCol) = "Tot. Rem."
    ColTotaBon = Fg1.Cols - 1
    With Fg1
        .Select 1, xNumCol, Fg1.Rows - 1, xNumCol
        .FillStyle = flexFillRepeat
        .CellBackColor = &HEBD7BC
    End With
    
    'HALLAMOS EL TOTAL DE LAS BONIFICACIONES POR EMPLEADO
    For A = 1 To Fg1.Rows - 1
        xTotal = 0
        For B = 2 To xNumCol
            xTotal = xTotal + Val(Fg1.TextMatrix(A, B))
        Next B
        Fg1.TextMatrix(A, xNumCol) = Format(xTotal, "0.00")
    Next A
        
        
    '**************************************
    'MOSTRAMOS LOS DESCUENTOS QUE SE APLICA
    Set Rst = Nothing
    RST_Busq Rst, "SELECT pla_conceptos.id, pla_conceptos.abrev, pla_conceptos.porcen, pla_conceptos.tipo, " _
        & " pla_conceptos.orden From pla_conceptos Where (((pla_conceptos.tipo) = 2) AND activo = -1) ORDER BY pla_conceptos.orden", xCon

    Rst.MoveFirst
    
    xNumCol = xNumCol + 1
    xColDesc = xNumCol
    For A = 1 To Rst.RecordCount
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, xNumCol) = Rst("abrev")
        
        'AGREGAMOS LA O.N.P.
'        If Rst("id") = 14 Then
'            For B = 1 To Fg1.Rows - 1
'                RST_Busq Rst2, "SELECT pl_nomina.id, pl_nominacarg.idcom, pl_nominacarg.imptot, pl_nominacarg.porcen " _
'                    & " FROM pl_nomina LEFT JOIN pl_nominacarg ON pl_nomina.id = pl_nominacarg.idnom " _
'                    & " WHERE (((pl_nomina.id) = " & Fg1.TextMatrix(B, 0) & ") AND ((pl_nominacarg.idcom) = 14))", xcon
'
'                If Rst2.RecordCount <> 0 Then
'                    If Rst2("imptot") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format(Format(Rst2("imptot"), "0.00"), "0.00")
'                    If Rst2("porcen") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format((Val(Fg1.TextMatrix(B, 2)) * (Rst2("porcen") / 100)), "0.00")
'                End If
'            Next B

            For B = 1 To Fg1.Rows - 1
               ' RST_Busq Rst2, "SELECT pl_nomina.id, pl_nominacarg.idcom, pl_nominacarg.imptot, pl_nominacarg.porcen " _
                    & " FROM pl_nomina LEFT JOIN pl_nominacarg ON pl_nomina.id = pl_nominacarg.idnom " _
                    & " WHERE (((pl_nomina.id) = " & Fg1.TextMatrix(B, 0) & ") AND ((pl_nominacarg.idcom) = " & Rst("id") & "))", xCon
                
                RST_Busq Rst2, "SELECT pla_empleados.id, pla_nominacarg.idcom, pla_nominacarg.imptot, pla_nominacarg.porcen " _
                    & " FROM pla_empleados LEFT JOIN pla_nominacarg ON pla_empleados.id = pla_nominacarg.idnom " _
                    & " WHERE (((pla_empleados.id) = " & Fg1.TextMatrix(B, 0) & ") AND ((pla_nominacarg.idcom) = " & Rst("id") & "))", xCon
                

                If Rst2.RecordCount <> 0 Then
                    'If Rst2("imptot") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format(Format(Rst2("imptot"), "0.00"), "0.00")
                    'If Rst2("porcen") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format((Val(Fg1.TextMatrix(B, 2)) * (Rst2("porcen") / 100)), "0.00")
                    If Rst2("imptot") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format(Format(Rst2("imptot"), "0.00"), "0.00")
                    If Rst2("porcen") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format((Val(Fg1.TextMatrix(B, ColTotaBon)) * (Rst2("porcen") / 100)), "0.00")
                End If
            Next B
'        End If
'
'        'AGREGAMOS LA AFP
'        If Rst("id") = 18 Then
'            For B = 1 To Fg1.Rows - 1
'                RST_Busq Rst2, "SELECT pl_nomina.id, pl_nominacarg.idcom, pl_nominacarg.imptot, pl_nominacarg.porcen " _
'                    & " FROM pl_nomina LEFT JOIN pl_nominacarg ON pl_nomina.id = pl_nominacarg.idnom " _
'                    & " WHERE (((pl_nomina.id) = " & Fg1.TextMatrix(B, 0) & ") AND ((pl_nominacarg.idcom) = 18))", xcon
'
'                If Rst2.RecordCount <> 0 Then
'                    If Rst2("imptot") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format(Format(Rst2("imptot"), "0.00"), "0.00")
'                    If Rst2("porcen") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format((Val(Fg1.TextMatrix(B, 2)) * (Rst2("porcen") / 100)), "0.00")
'                End If
'            Next B
'        End If
        
        Fg1.ColWidth(xNumCol) = 800
        Fg1.ColAlignment(xNumCol) = flexAlignRightCenter
        Rst.MoveNext
        
        If Rst.EOF = True Then
            Exit For
        End If
        xNumCol = xNumCol + 1
    Next A
    
    With Fg1
        .Select 1, xColDesc, Fg1.Rows - 1, xNumCol
        .FillStyle = flexFillRepeat
        .CellBackColor = &HC0E0FF
    End With
    

    'AGREGAMOS LA COLUMNA PARA EL TOTAL DE RETENCIONES
    xNumCol = xNumCol + 1
    Fg1.Cols = Fg1.Cols + 1
    Fg1.TextMatrix(0, xNumCol) = "Tot. Ret."

    With Fg1
        .Select 1, xNumCol, Fg1.Rows - 1, xNumCol
        .FillStyle = flexFillRepeat
        .CellBackColor = &HEBD7BC
    End With

    'HALLAMOS EL TOTAL DE LAS BONIFICACIONES POR EMPLEADO
    For A = 1 To Fg1.Rows - 1
        xTotal = 0
        For B = xColDesc To xNumCol
            xTotal = xTotal + Val(Fg1.TextMatrix(A, B))
        Next B
        Fg1.TextMatrix(A, xNumCol) = Format(xTotal, "0.00")
    Next A


    
    '********************************************
    'MOSTRAMOS LOS APORTES QUE APLICA EL EMPLEADO
    Set Rst = Nothing
    RST_Busq Rst, "SELECT pla_conceptos.id, pla_conceptos.abrev, pla_conceptos.porcen, pla_conceptos.tipo, " _
        & " pla_conceptos.orden From pla_conceptos Where (((pla_conceptos.tipo) = 3)) ORDER BY pla_conceptos.orden", xCon

    Rst.MoveFirst
    
    xNumCol = xNumCol + 1
    xColDesc = xNumCol
    
    For A = 1 To Rst.RecordCount
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, xNumCol) = Rst("abrev")
        
        'AGREGAMOS EL ESSALUD
        If Rst("id") = 11 Then
            For B = 1 To Fg1.Rows - 1
                RST_Busq Rst2, "SELECT pla_empleados.id, pla_nominacarg.idcom, pla_nominacarg.imptot, pla_nominacarg.porcen " _
                    & " FROM pla_empleados LEFT JOIN pla_nominacarg ON pla_empleados.id = pla_nominacarg.idnom " _
                    & " WHERE (((pla_empleados.id) = " & Fg1.TextMatrix(B, 0) & ") AND ((pla_nominacarg.idcom) = 11))", xCon
                
                If Rst2.RecordCount <> 0 Then
'                    If Rst2("imptot") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format(Format(Rst2("imptot"), "0.00"), "0.00")
'                    If Rst2("porcen") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format((Val(Fg1.TextMatrix(B, 2)) * (Rst2("porcen") / 100)), "0.00")
                    If Rst2("imptot") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format(Format(Rst2("imptot"), "0.00"), "0.00")
                    If Rst2("porcen") <> "0" Then Fg1.TextMatrix(B, xNumCol) = Format((Val(Fg1.TextMatrix(B, ColTotaBon)) * (Rst2("porcen") / 100)), "0.00")
                End If
            Next B
        End If
        
        Fg1.ColWidth(xNumCol) = 800
        Fg1.ColAlignment(xNumCol) = flexAlignRightCenter
        Rst.MoveNext
        
        If Rst.EOF = True Then
            Exit For
        End If
        xNumCol = xNumCol + 1
    Next A

    With Fg1
        .Select 1, xColDesc, Fg1.Rows - 1, xNumCol
        .FillStyle = flexFillRepeat
        .CellBackColor = &HC0C0FF
    End With


    'AGREGAMOS LA COLUMNA PARA EL TOTAL DE RETENCIONES
    xNumCol = xNumCol + 1
    Fg1.Cols = Fg1.Cols + 1
    Fg1.TextMatrix(0, xNumCol) = "Tot. Ret."

    With Fg1
        .Select 1, xNumCol, Fg1.Rows - 1, xNumCol
        .FillStyle = flexFillRepeat
        .CellBackColor = &HEBD7BC
    End With

    'HALLAMOS EL TOTAL DE LAS BONIFICACIONES POR EMPLEADO
    For A = 1 To Fg1.Rows - 1
        xTotal = 0
        For B = xColDesc To xNumCol
            xTotal = xTotal + Val(Fg1.TextMatrix(A, B))
        Next B
        Fg1.TextMatrix(A, xNumCol) = Format(xTotal, "0.00")
    Next A

End Sub

Sub Blanquea()
    TxtNumPla.Text = ""
    txtFchPro.Valor = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    TxtBusTipPla.Text = ""
    
    TxtTotBas.Text = ""
    TxtTotDsct.Text = ""
    TxtTotApo.Text = ""
    
End Sub

Sub Bloquea()
    TxtNumPla.Locked = Not TxtNumPla.Locked
    txtFchPro.Locked = Not txtFchPro.Locked
    TxtFchIni.Locked = Not TxtFchIni.Locked
    TxtFchFin.Locked = Not TxtFchFin.Locked
    TxtBusTipPla.Locked = Not TxtBusTipPla.Locked
End Sub


Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then
            IniciarGrid
            MuestraPlanilla
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        nuevo
    End If
    
    If Button.Index = 3 Then
        Dim Rpta As Integer
        Rpta = MsgBox("¿Esta seguro de eliminar la planilla especificada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            xCon.Execute "DELETE * FROM pl_planillas WHERE id = " & RstCab("id") & ""
            MsgBox "La planilla se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    
    If Button.Index = 5 Then
        Cancelar
    End If
    
    If Button.Index = 6 Then
        If Grabar = True Then
            Cancelar
        End If
    End If
    If Button.Index = 14 Then
        Set RstCab = Nothing
        Unload Me
    End If
End Sub

Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
    
End Sub

Sub Cancelar()
    TabOne1.TabEnabled(0) = True
    ActivaTool
    QueHace = 3
    Bloquea
    TabOne1.CurrTab = 0
End Sub

Function Grabar() As Boolean
    If TxtNumPla.Text = "" Then
        MsgBox "No ha especificado el numero de planilla", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumPla.SetFocus
        Exit Function
    End If
    
    If TxtBusTipPla.Text = "" Then
        MsgBox "No ha especificado el tipo de planilla que se va a emitir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtBusTipPla.SetFocus
        Exit Function
    End If
    
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de incio de la planilla", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If
    
    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha final para la planilla", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Function
    End If
    
    Dim xRango As Integer
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Integer
    Dim A As Integer
    Dim B As Integer
    Dim xRangoRemu As Integer
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("pl_planillas", xCon, "id")
        RST_Busq RstCab, "SELECT * FROM pl_planillas", xCon
        RST_Busq RstDet, "SELECT * FROM pl_planillasdet", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
    End If
    
    RstCab("numpla") = TxtNumPla.Text
    If NulosC(txtFchPro.Valor) <> "" Then RstCab("fchpro") = txtFchPro.Valor
    RstCab("fchini") = TxtFchIni.Valor
    RstCab("fchfin") = TxtFchFin.Valor
    RstCab("tippla") = Busca_Codigo(TxtBusTipPla.Text, "descripcion", "id", "m_tipopersonal", "C", xCon)
    RstCab("totapo") = Val(TxtTotBas.Text)
    RstCab("totdes") = Val(TxtTotDsct.Text)
    RstCab("totpag") = (Val(TxtTotBas.Text) - Val(TxtTotDsct.Text))
    RstCab.Update
    
    xRango = HallarRangoRemu()
    
    For A = 1 To Fg1.Rows - 1
        For B = 2 To xRango
            RstDet.AddNew
            RstDet("idpla") = xId
            RstDet("idnom") = Val(Fg1.TextMatrix(A, 0))
            RstDet("idcom") = Busca_Codigo(Fg1.TextMatrix(0, B), "abrev", "id", "pl_conceptos", "C", xCon)
            RstDet("porcen") = Busca_Codigo(Fg1.TextMatrix(0, B), "abrev", "porcen", "pl_conceptos", "C", xCon)
            RstDet("import") = Val(Fg1.TextMatrix(A, B))
            RstDet("tipo") = 1
            RstDet.Update
        Next B
        
    Next A
    
    'GURADAMOS LOS DESCUENTOS AL TRABAJADOR
    xRango = HallarRangoDsct()
    xRangoRemu = HallarRangoRemu() + 2
    
    For A = 1 To Fg1.Rows - 1
        For B = xRangoRemu To xRango
            RstDet.AddNew
            RstDet("idpla") = xId
            RstDet("idnom") = Val(Fg1.TextMatrix(A, 0))
            RstDet("idcom") = Busca_Codigo(Fg1.TextMatrix(0, B), "abrev", "id", "pl_conceptos", "C", xCon)
            RstDet("porcen") = Busca_Codigo(Fg1.TextMatrix(0, B), "abrev", "porcen", "pl_conceptos", "C", xCon)
            RstDet("import") = Val(Fg1.TextMatrix(A, B))
            RstDet("tipo") = 2
            RstDet.Update
        Next B
    Next A
    
    'GURADAMOS LOS APORTES DEL EMPLEADOR
    xRango = HallarRangoEmp()
    xRangoRemu = HallarRangoDsct() + 2
    
    For A = 1 To Fg1.Rows - 1
        For B = xRangoRemu To xRango
            RstDet.AddNew
            RstDet("idpla") = xId
            RstDet("idnom") = Val(Fg1.TextMatrix(A, 0))
            RstDet("idcom") = Busca_Codigo(Fg1.TextMatrix(0, B), "abrev", "id", "pl_conceptos", "C", xCon)
            RstDet("porcen") = Busca_Codigo(Fg1.TextMatrix(0, B), "abrev", "porcen", "pl_conceptos", "C", xCon)
            RstDet("import") = Val(Fg1.TextMatrix(A, B))
            RstDet("tipo") = 3
            RstDet.Update
        Next B
    Next A
    
    xCon.CommitTrans
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Function HallarRangoRemu() As Integer
    Dim A As Integer
    For A = 2 To Fg1.Cols
        Fg1.Select 1, A
        If Fg1.CellBackColor <> 14548991 Then
            '.CellBackColor = &HEBD7BC
            HallarRangoRemu = A - 1
            Exit Function
        End If
    Next A
End Function

Function HallarRangoDsct() As Integer
    Dim A As Integer
    Dim xCol As Integer
    xCol = HallarRangoRemu() + 2
    For A = xCol To Fg1.Cols
        Fg1.Select 1, A
        If Fg1.CellBackColor <> 12640511 Then
            '.CellBackColor = &HEBD7BC
            HallarRangoDsct = A - 1
            Exit Function
        End If
    Next A
End Function

Function HallarRangoEmp() As Integer
    Dim A As Integer
    Dim xCol As Integer
    xCol = HallarRangoDsct() + 2
    For A = xCol To Fg1.Cols
        Fg1.Select 1, A
        If Fg1.CellBackColor <> 12632319 Then
            '.CellBackColor = &HEBD7BC
            HallarRangoEmp = A - 1
            Exit Function
        End If
    Next A
End Function


Private Sub TxtBusTipPla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtBusTipPla_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusGrupo_Click
    End If
End Sub

Private Sub TxtNumPla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub VSFlexGrid2_Click()

End Sub
