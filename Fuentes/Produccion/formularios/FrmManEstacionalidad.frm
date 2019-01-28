VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManEstacionalidad 
   Caption         =   "Produccion - Estacionalidad de la Materia Prima"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6780
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   11595
      _cx             =   20452
      _cy             =   11959
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   8388608
      Caption         =   "  &Consulta  |   &Detalle  "
      Align           =   0
      CurrTab         =   0
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6360
         Left            =   45
         TabIndex        =   3
         Top             =   375
         Width           =   11505
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5580
            Left            =   30
            TabIndex        =   6
            Top             =   345
            Width           =   11445
            _cx             =   20188
            _cy             =   9842
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   8454016
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16777215
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
            Rows            =   2
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManEstacionalidad.frx":0000
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
         Begin VB.Frame Frame4 
            Height          =   480
            Left            =   30
            TabIndex        =   12
            Top             =   5865
            Width           =   11460
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Abundancia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3285
               TabIndex        =   15
               Top             =   195
               Width           =   1020
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Escaces"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1185
               TabIndex        =   14
               Top             =   195
               Width           =   735
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H0000C000&
               BackStyle       =   1  'Opaque
               Height          =   225
               Left            =   2445
               Top             =   180
               Width           =   720
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H000000FF&
               BackStyle       =   1  'Opaque
               Height          =   225
               Left            =   375
               Top             =   180
               Width           =   720
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Regular"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   5325
               TabIndex        =   13
               Top             =   195
               Width           =   675
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H0080FFFF&
               BackStyle       =   1  'Opaque
               Height          =   225
               Left            =   4515
               Top             =   180
               Width           =   720
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   15
            TabIndex        =   4
            Top             =   30
            Width           =   11460
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6360
         Left            =   12240
         TabIndex        =   1
         Top             =   375
         Width           =   11505
         Begin VB.Frame Frame5 
            Height          =   705
            Left            =   2737
            TabIndex        =   16
            Top             =   5160
            Width           =   6120
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   240
               Left            =   4335
               TabIndex        =   22
               Top             =   300
               Width           =   135
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   240
               Left            =   2520
               TabIndex        =   21
               Top             =   300
               Width           =   135
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "1"
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
               Height          =   240
               Left            =   945
               TabIndex        =   20
               Top             =   300
               Width           =   135
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Abundancia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2790
               TabIndex        =   19
               Top             =   315
               Width           =   1020
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Escaces"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1185
               TabIndex        =   18
               Top             =   315
               Width           =   735
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Regular"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   4590
               TabIndex        =   17
               Top             =   315
               Width           =   675
            End
         End
         Begin VB.Frame Frame3 
            Height          =   5415
            Left            =   727
            TabIndex        =   7
            Top             =   675
            Width           =   10140
            Begin VB.CommandButton CmdBusMatPri 
               Height          =   240
               Left            =   3345
               Picture         =   "FrmManEstacionalidad.frx":018D
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   540
               Width           =   240
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   3150
               Left            =   2700
               TabIndex        =   10
               Top             =   1170
               Width           =   4845
               _cx             =   8546
               _cy             =   5556
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16777215
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
               Rows            =   13
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManEstacionalidad.frx":02BF
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
            Begin VB.TextBox TxtDescripcion 
               Height          =   300
               Left            =   180
               Locked          =   -1  'True
               TabIndex        =   8
               Text            =   "TxtDescripcion"
               Top             =   1830
               Visible         =   0   'False
               Width           =   2040
            End
            Begin VB.TextBox TxtIdMatPri 
               Height          =   300
               Left            =   2700
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   24
               Text            =   "TxtidMatPri"
               Top             =   510
               Width           =   915
            End
            Begin VB.Label LblMatPri 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblMatPri"
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
               Left            =   3645
               TabIndex        =   26
               Top             =   510
               Width           =   4170
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Distribucón x meses"
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
               Left            =   2700
               TabIndex        =   11
               Top             =   930
               Width           =   1710
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion producto"
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   9
               Top             =   1620
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Materia Prima"
               Height          =   195
               Index           =   25
               Left            =   1560
               TabIndex        =   25
               ToolTipText     =   "Materia Prima Principal"
               Top             =   600
               Width           =   960
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Estacionalidad"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   15
            TabIndex        =   2
            Top             =   30
            Width           =   11460
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   30
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
            Picture         =   "FrmManEstacionalidad.frx":0320
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEstacionalidad.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEstacionalidad.frx":09E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEstacionalidad.frx":0E3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEstacionalidad.frx":0F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEstacionalidad.frx":1498
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEstacionalidad.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEstacionalidad.frx":1AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEstacionalidad.frx":1C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEstacionalidad.frx":2058
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManEstacionalidad.frx":21C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   1005
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
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
            Style           =   5
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmManEstacionalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANCOSTO.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : MUESTRA LA ESTACIONALIDAD DE LA MATERIA PRIMA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 05/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim Agregando As Boolean        ' INFORMA QUE SE ESTA AGREGANDO UNA FILA AL CONTROL FLEXGRID
Dim SeEjecuto As Boolean        ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim Rst As New ADODB.Recordset  ' RECORDSET PARA ALAMCENAR LOS DATOS DE LA TABLA pro_estacionalidad
Dim QueHace As Integer          ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO


Private Sub Fg2_EnterCell()
    If QueHace = 3 Then
        Fg2.Editable = flexEDNone
    Else
        If Fg2.Col = 2 Then
            Fg2.Editable = flexEDKbdMouse
        Else
            Fg2.Editable = flexEDNone
        End If
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE LE FORMULARIO
    If SeEjecuto = False Then
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        VerTodaEstacionalidad
        SeEjecuto = True
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : VerTodaEstacionalidad
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LA ESTACIONALIDAD DE TODOS LOS PRODUCTOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub VerTodaEstacionalidad()
    Dim A As Integer
    RST_Busq Rst, "SELECT pro_estacionalidad.* , alm_inventario.descripcion as nommatpri " _
                + vbCr + " FROM alm_inventario INNER JOIN pro_estacionalidad ON alm_inventario.id = pro_estacionalidad.iditem " _
                + vbCr + " ORDER BY alm_inventario.descripcion " _
                , xCon
    
    Rst.MoveFirst
    Fg1.Rows = 1
    Agregando = True
    
    For A = 1 To Rst.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(A, 1) = NulosC(Rst("nommatpri"))
        Fg1.TextMatrix(A, 14) = Trim(Rst("id"))
        Fg1.TextMatrix(A, 15) = Trim(Rst("iditem"))
        MuestraEstacionalidad A
        Rst.MoveNext
        
        If Rst.EOF = True Then
            Exit For
        End If
    Next A
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraEstacionalidad
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LA ESTACIONALIDAD DE UN PRODUCTO
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    xFila     |  Integer          |
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraEstacionalidad(xFila As Integer)
    Dim A As Integer
    Dim xCol As Integer
    Dim NumMeses As Integer

    With Fg1
        If Rst("ene") = 1 Then .Select xFila, 2, xFila, 2: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 2) = "1"
        If Rst("feb") = 1 Then .Select xFila, 3, xFila, 3: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 3) = "1"
        If Rst("mar") = 1 Then .Select xFila, 4, xFila, 4: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 4) = "1"
        If Rst("abr") = 1 Then .Select xFila, 5, xFila, 5: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 5) = "1"
        If Rst("may") = 1 Then .Select xFila, 6, xFila, 6: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 6) = "1"
        If Rst("jun") = 1 Then .Select xFila, 7, xFila, 7: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 7) = "1"
        If Rst("jul") = 1 Then .Select xFila, 8, xFila, 8: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 8) = "1"
        If Rst("ago") = 1 Then .Select xFila, 9, xFila, 9: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 9) = "1"
        If Rst("set") = 1 Then .Select xFila, 10, xFila, 10: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 10) = "1"
        If Rst("oct") = 1 Then .Select xFila, 11, xFila, 11: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 11) = "1"
        If Rst("nov") = 1 Then .Select xFila, 12, xFila, 12: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 12) = "1"
        If Rst("dic") = 1 Then .Select xFila, 13, xFila, 13: .FillStyle = flexFillRepeat: .CellBackColor = &HC0&: .CellForeColor = &HC0&: .TextMatrix(xFila, 13) = "1"
        
        If Rst("ene") = 2 Then .Select xFila, 2, xFila, 2: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 2) = "2"
        If Rst("feb") = 2 Then .Select xFila, 3, xFila, 3: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 3) = "2"
        If Rst("mar") = 2 Then .Select xFila, 4, xFila, 4: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 4) = "2"
        If Rst("abr") = 2 Then .Select xFila, 5, xFila, 5: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 5) = "2"
        If Rst("may") = 2 Then .Select xFila, 6, xFila, 6: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 6) = "2"
        If Rst("jun") = 2 Then .Select xFila, 7, xFila, 7: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 7) = "2"
        If Rst("jul") = 2 Then .Select xFila, 8, xFila, 8: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 8) = "2"
        If Rst("ago") = 2 Then .Select xFila, 9, xFila, 9: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 9) = "2"
        If Rst("set") = 2 Then .Select xFila, 10, xFila, 10: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 10) = "2"
        If Rst("oct") = 2 Then .Select xFila, 11, xFila, 11: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 11) = "2"
        If Rst("nov") = 2 Then .Select xFila, 12, xFila, 12: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 12) = "2"
        If Rst("dic") = 2 Then .Select xFila, 13, xFila, 13: .FillStyle = flexFillRepeat: .CellBackColor = &HC000&: .CellForeColor = &HC000&: .TextMatrix(xFila, 13) = "2"
        
        If Rst("ene") = 3 Then .Select xFila, 2, xFila, 2: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 2) = "3"
        If Rst("feb") = 3 Then .Select xFila, 3, xFila, 3: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 3) = "3"
        If Rst("mar") = 3 Then .Select xFila, 4, xFila, 4: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 4) = "3"
        If Rst("abr") = 3 Then .Select xFila, 5, xFila, 5: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 5) = "3"
        If Rst("may") = 3 Then .Select xFila, 6, xFila, 6: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 6) = "3"
        If Rst("jun") = 3 Then .Select xFila, 7, xFila, 7: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 7) = "3"
        If Rst("jul") = 3 Then .Select xFila, 8, xFila, 8: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 8) = "3"
        If Rst("ago") = 3 Then .Select xFila, 9, xFila, 9: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 9) = "3"
        If Rst("set") = 3 Then .Select xFila, 10, xFila, 10: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 10) = "3"
        If Rst("oct") = 3 Then .Select xFila, 11, xFila, 11: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 11) = "3"
        If Rst("nov") = 3 Then .Select xFila, 12, xFila, 12: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 12) = "3"
        If Rst("dic") = 3 Then .Select xFila, 13, xFila, 13: .FillStyle = flexFillRepeat: .CellBackColor = &H80FFFF: .CellForeColor = &H80FFFF: .TextMatrix(xFila, 13) = "3"
    End With
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
    Fg1.ColWidth(14) = 0 '--id
    Fg1.ColWidth(15) = 0 '--iditem
    
    Fg2.TextMatrix(1, 1) = "Enero"
    Fg2.TextMatrix(2, 1) = "Febrero"
    Fg2.TextMatrix(3, 1) = "Marzo"
    Fg2.TextMatrix(4, 1) = "Abril"
    Fg2.TextMatrix(5, 1) = "Mayo"
    Fg2.TextMatrix(6, 1) = "Junio"
    Fg2.TextMatrix(7, 1) = "Julio"
    Fg2.TextMatrix(8, 1) = "Agosto"
    Fg2.TextMatrix(9, 1) = "Setiembre"
    Fg2.TextMatrix(10, 1) = "Octubre"
    Fg2.TextMatrix(11, 1) = "Noviembre"
    Fg2.TextMatrix(12, 1) = "Diciembre"
    
    Fg1.Editable = flexEDNone
    
    
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.ExplorerBar = flexExSortShowAndMove
    Fg1.ForeColorSel = &H0&
    Fg1.BackColorSel = &HC0E0FF
    
    Fg2.Editable = flexEDNone
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.ForeColorSel = &H0&
    Fg2.BackColorSel = &HC0E0FF
    Fg2.ColComboList(2) = "#1;Escaces|#2;Abundancia|#3;Regular"
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    Label1.Caption = "Agregando Estacionalidad"
    Blanquea
    ActivaTool
    
    Fg2.Editable = flexEDKbdMouse
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    TxtIdMatPri.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    xHorIni = Time
    Label1.Caption = "Modificando Estacionalidad"
'    Blanquea
    ActivaTool
    Fg2.Editable = flexEDKbdMouse
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    TxtIdMatPri.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESO O MODIFICACION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    ActivaTool
    Label1.Caption = "Detalle de la Estacionalidad"
    Fg2.Editable = flexEDNone
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_estacionalidad, ESTA FUNCION DEVUELVE VERDADERO
'*                    CUANDO TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    '--verificar si la materia prima esta seleccionada
    If NulosN(TxtIdMatPri.Text) = 0 Then
        MsgBox "No ha especificado la descripcion de la materia prima", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMatPri.SetFocus
        Exit Function
    End If
    
    Dim A As Integer
    Dim xId As Double
    
    For A = 1 To 12
        If Fg2.TextMatrix(A, 2) = "" Then
            MsgBox "No ha especificado el estado para el mes " & Trim(Fg1.TextMatrix(A, 1)), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    Next A
    
    Dim RstGraba As New ADODB.Recordset
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("pro_estacionalidad", xCon, "id")
        RST_Busq RstGraba, "SELECT * FROM pro_estacionalidad", xCon
        RstGraba.AddNew
        RstGraba("id") = xId
    Else
        xId = NulosN(Fg1.TextMatrix(Fg1.Row, 14))
        RST_Busq RstGraba, "SELECT * FROM pro_estacionalidad WHERE id = " & xId & "", xCon
    End If
    
    RstGraba("iditem") = NulosN(TxtIdMatPri.Text)
    
    RstGraba("descripcion") = NulosC(LblMatPri.Caption)
    RstGraba("ene") = NulosN(Fg2.TextMatrix(1, 2))
    RstGraba("feb") = NulosN(Fg2.TextMatrix(2, 2))
    RstGraba("mar") = NulosN(Fg2.TextMatrix(3, 2))
    RstGraba("abr") = NulosN(Fg2.TextMatrix(4, 2))
    RstGraba("may") = NulosN(Fg2.TextMatrix(5, 2))
    RstGraba("jun") = NulosN(Fg2.TextMatrix(6, 2))
    RstGraba("jul") = NulosN(Fg2.TextMatrix(7, 2))
    RstGraba("ago") = NulosN(Fg2.TextMatrix(8, 2))
    RstGraba("set") = NulosN(Fg2.TextMatrix(9, 2))
    RstGraba("oct") = NulosN(Fg2.TextMatrix(10, 2))
    RstGraba("nov") = NulosN(Fg2.TextMatrix(11, 2))
    RstGraba("dic") = NulosN(Fg2.TextMatrix(12, 2))
    
    RstGraba.Update
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    
    xCon.CommitTrans
    MsgBox "La estacionalidad del " + Trim(LblMatPri.Caption) + " se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstGraba = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : LIMPIA LAS CELDAS DEL CONTROL Fg2
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    
    TxtIdMatPri.Text = ""
    LblMatPri.Caption = ""
    Fg2.TextMatrix(1, 2) = ""
    Fg2.TextMatrix(2, 2) = ""
    Fg2.TextMatrix(3, 2) = ""
    Fg2.TextMatrix(4, 2) = ""
    Fg2.TextMatrix(5, 2) = ""
    Fg2.TextMatrix(6, 2) = ""
    Fg2.TextMatrix(7, 2) = ""
    Fg2.TextMatrix(8, 2) = ""
    Fg2.TextMatrix(9, 2) = ""
    Fg2.TextMatrix(10, 2) = ""
    Fg2.TextMatrix(11, 2) = ""
    Fg2.TextMatrix(12, 2) = ""
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraDataFruta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LOS DATOS DE LA ESTACIONALIDAD EN EL CONTROL Fg2
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraDataFruta()
    Dim RstEst As New ADODB.Recordset
    
    RST_Busq RstEst, "SELECT * FROM pro_estacionalidad WHERE id = " & NulosN(Fg1.TextMatrix(Fg1.Row, 14)) & "", xCon
    
    LblMatPri.Caption = Fg1.TextMatrix(Fg1.Row, 1)
    TxtIdMatPri.Text = Fg1.TextMatrix(Fg1.Row, 15)
    
    Fg2.TextMatrix(1, 2) = NulosN(RstEst("ene"))
    Fg2.TextMatrix(2, 2) = NulosN(RstEst("feb"))
    Fg2.TextMatrix(3, 2) = NulosN(RstEst("mar"))
    Fg2.TextMatrix(4, 2) = NulosN(RstEst("abr"))
    Fg2.TextMatrix(5, 2) = NulosN(RstEst("may"))
    Fg2.TextMatrix(6, 2) = NulosN(RstEst("jun"))
    Fg2.TextMatrix(7, 2) = NulosN(RstEst("jul"))
    Fg2.TextMatrix(8, 2) = NulosN(RstEst("ago"))
    Fg2.TextMatrix(9, 2) = NulosN(RstEst("set"))
    Fg2.TextMatrix(10, 2) = NulosN(RstEst("oct"))
    Fg2.TextMatrix(11, 2) = NulosN(RstEst("nov"))
    Fg2.TextMatrix(12, 2) = NulosN(RstEst("dic"))
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraDataFruta
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA pro_estacionalidad
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Dim xId As Double
    
    If Fg1.Row < Fg1.FixedRows Then
        MsgBox "Seleccione un registro correcto", vbInformation, xTitulo
        Exit Sub
    End If
    xId = NulosN(Fg1.TextMatrix(Fg1.Row, 14))
    Rpta = MsgBox("Esta seguro de eliminar la estacionalidad seleccionadade la materia prima " & Fg1.TextMatrix(Fg1.Row, 1), vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_estacionalidad WHERE id = " & xId & " "
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo
        
        MsgBox "La estacionalidad se elimino con exito", vbInformation + vbYesNo + vbDefaultButton1, xTitulo
        Fg1.RemoveItem Fg1.Row
        Exit Sub
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            VerTodaEstacionalidad
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 14 Then
        Set Rst = Nothing
        Unload Me
    End If
End Sub





Private Sub CmdBusMatPri_Click()
    'BUSCAMOS LA MATERIA PRIMA PRINCIPAL DEL PRODUCTO QUE SE LE ASIGNARA AL ITEM
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripción":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Código":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    'xform.SQLCad = "SELECT pro_estacionalidad.* FROM pro_estacionalidad"
    xform.SQLCad = "SELECT alm_inventario.id, alm_inventario.descripcion " _
        + vbCr + " FROM alm_inventario LEFT JOIN pro_estacionalidad ON alm_inventario.id = pro_estacionalidad.iditem " _
        + vbCr + " WHERE (((alm_inventario.tippro)=1) AND ((pro_estacionalidad.iditem) Is Null) AND ((alm_inventario.activo)=-1)) "
 
    
    xform.titulo = "Buscando Materia Prima Principal"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMatPri.Text = xRs("id")
        LblMatPri.Caption = NulosC(xRs("descripcion"))
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub TxtidMatPri_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtidMatPri_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMatPri_Click
    End If
End Sub

Private Sub TxtIdMatPri_Validate(Cancel As Boolean)

    If NulosN(TxtIdMatPri.Text) <> 0 Then
        Dim RstTem As New ADODB.Recordset
        Set RstTem = BuscaConCriterio("SELECT alm_inventario.id, alm_inventario.descripcion " _
                                     & " FROM alm_inventario LEFT JOIN pro_estacionalidad ON alm_inventario.id = pro_estacionalidad.iditem " _
                                     & " WHERE (((alm_inventario.id)=" & NulosN(TxtIdMatPri.Text) & ") AND ((alm_inventario.tippro)=1) AND ((pro_estacionalidad.iditem) Is Null) AND ((alm_inventario.activo)=-1));", _
                                     xCon)
                                     
        If RstTem.RecordCount <> 0 Then
            LblMatPri.Caption = NulosC(RstTem("descripcion"))
            
        Else
            TxtIdMatPri.Text = ""
            LblMatPri.Caption = ""
        End If
        Set RstTem = Nothing
    Else
        TxtIdMatPri.Text = ""
        LblMatPri.Caption = ""
    End If

End Sub
