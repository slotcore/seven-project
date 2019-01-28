VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManPlanCtas 
   Caption         =   "Contabilidad - Plan de Cuentas"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":08D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":0C68
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":0DEC
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":1240
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":1358
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":189C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":1DE0
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":1EF4
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":2008
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":245C
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPlanCtas.frx":25C8
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   12
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   12726
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
      Appearance      =   1
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Detalle de la Cuenta"
         Height          =   6795
         Left            =   12525
         TabIndex        =   16
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame3 
            Height          =   5430
            Left            =   795
            TabIndex        =   17
            Top             =   645
            Width           =   10125
            Begin VB.Frame fra 
               Caption         =   "Cuentas del Balance destinados a:"
               Enabled         =   0   'False
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
               Height          =   600
               Index           =   1
               Left            =   900
               TabIndex        =   48
               Top             =   3420
               Width           =   8670
               Begin VB.CheckBox ChkActivo 
                  Caption         =   "Activo"
                  Height          =   255
                  Left            =   690
                  TabIndex        =   50
                  Top             =   270
                  Width           =   1095
               End
               Begin VB.CheckBox ChkPasivo 
                  Caption         =   "Pasivo y Patrimonio"
                  Height          =   285
                  Left            =   2940
                  TabIndex        =   49
                  Top             =   240
                  Width           =   1755
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "[ Datos para la Apertura ]"
               Enabled         =   0   'False
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
               Height          =   1065
               Left            =   900
               TabIndex        =   43
               Top             =   4140
               Width           =   8670
               Begin VB.CommandButton CmdBusModulo 
                  Height          =   240
                  Left            =   2490
                  Picture         =   "FrmManPlanCtas.frx":2B10
                  Style           =   1  'Graphical
                  TabIndex        =   45
                  Top             =   615
                  Width           =   240
               End
               Begin VB.OptionButton OptNO1 
                  Caption         =   "No"
                  Height          =   195
                  Left            =   4230
                  TabIndex        =   9
                  Top             =   345
                  Width           =   675
               End
               Begin VB.OptionButton OptSI1 
                  Caption         =   "Si"
                  Height          =   195
                  Left            =   3330
                  TabIndex        =   8
                  Top             =   345
                  Width           =   675
               End
               Begin VB.TextBox TxtIdModulo 
                  Height          =   300
                  Left            =   1800
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   10
                  Text            =   "TxtIdModulo"
                  Top             =   585
                  Width           =   960
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Módulo"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   47
                  Top             =   615
                  Width           =   525
               End
               Begin VB.Label LblDescModulo 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblDescModulo"
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
                  Left            =   2805
                  TabIndex        =   46
                  Top             =   585
                  Width           =   5565
               End
               Begin VB.Label Label6 
                  Caption         =   "Documentar la cuenta"
                  Height          =   180
                  Left            =   165
                  TabIndex        =   44
                  Top             =   315
                  Width           =   2670
               End
            End
            Begin VB.Frame fra 
               Caption         =   "[Naturaleza de la Cuenta ]"
               Enabled         =   0   'False
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
               Height          =   600
               Index           =   0
               Left            =   900
               TabIndex        =   42
               Top             =   2745
               Width           =   8670
               Begin VB.OptionButton opt_saldo 
                  Caption         =   "Haber: (H)"
                  Height          =   195
                  Index           =   1
                  Left            =   2940
                  TabIndex        =   7
                  Top             =   270
                  Width           =   1335
               End
               Begin VB.OptionButton opt_saldo 
                  Caption         =   "Debe: (D)"
                  Height          =   195
                  Index           =   0
                  Left            =   720
                  TabIndex        =   6
                  Top             =   270
                  Width           =   1335
               End
            End
            Begin VB.CommandButton cb 
               Height          =   240
               Index           =   2
               Left            =   3645
               Picture         =   "FrmManPlanCtas.frx":2C42
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   2790
               Visible         =   0   'False
               Width           =   225
            End
            Begin VB.CommandButton cb 
               Height          =   240
               Index           =   0
               Left            =   3645
               Picture         =   "FrmManPlanCtas.frx":2D74
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   1995
               Width           =   240
            End
            Begin VB.CommandButton cb 
               Height          =   240
               Index           =   1
               Left            =   3645
               Picture         =   "FrmManPlanCtas.frx":2EA6
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   2400
               Width           =   240
            End
            Begin VB.CommandButton CmdBusDesHaber 
               Height          =   240
               Left            =   3645
               Picture         =   "FrmManPlanCtas.frx":2FD8
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   1635
               Width           =   240
            End
            Begin VB.CommandButton CmdBusDesDebe 
               Height          =   240
               Left            =   3645
               Picture         =   "FrmManPlanCtas.frx":310A
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   1275
               Width           =   240
            End
            Begin VB.TextBox TxtHaber 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   3
               Text            =   "TxtHaber"
               Top             =   1605
               Width           =   1455
            End
            Begin VB.TextBox TxtDebe 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   2
               Text            =   "TxtDebe"
               Top             =   1245
               Width           =   1455
            End
            Begin VB.TextBox TxtNumCta 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   0
               Text            =   "TxtNumCta"
               Top             =   525
               Width           =   1455
            End
            Begin VB.TextBox TxtDescripcion 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   1
               Text            =   "TxtDescripcion"
               Top             =   885
               Width           =   7050
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   1
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   5
               Text            =   "txt_cb(1)"
               Top             =   2370
               Width           =   1455
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   4
               Text            =   "txt_cb(0)"
               Top             =   1965
               Width           =   1455
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   2
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   11
               Text            =   "txt_cb(2)"
               ToolTipText     =   "Ingrese DNI del Supervisor"
               Top             =   2745
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label lbl_cb_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb_cod(2)"
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
               Height          =   300
               Index           =   2
               Left            =   8055
               TabIndex        =   41
               Top             =   2745
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.Label lbl_cb_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb_cod(0)"
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
               Height          =   300
               Index           =   0
               Left            =   8040
               TabIndex        =   40
               Top             =   1965
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.Label lbl_cb_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb_cod(1)"
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
               Height          =   300
               Index           =   1
               Left            =   8055
               TabIndex        =   39
               Top             =   2370
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(2)"
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
               Index           =   2
               Left            =   3945
               TabIndex        =   38
               Top             =   2745
               Visible         =   0   'False
               Width           =   5565
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "3ra Distribución"
               Height          =   195
               Index           =   2
               Left            =   915
               TabIndex        =   37
               Top             =   2775
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(0)"
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
               Index           =   0
               Left            =   3945
               TabIndex        =   36
               Top             =   1965
               Width           =   5565
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Distribución"
               Height          =   195
               Index           =   0
               Left            =   915
               TabIndex        =   35
               Top             =   2010
               Width           =   825
            End
            Begin VB.Label lbl_cb 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb(1)"
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
               Index           =   1
               Left            =   3945
               TabIndex        =   34
               Top             =   2370
               Width           =   5565
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "2da Distribución"
               Height          =   195
               Index           =   1
               Left            =   915
               TabIndex        =   33
               Top             =   2400
               Width           =   1140
            End
            Begin VB.Label LblIdCtaDeb 
               AutoSize        =   -1  'True
               Caption         =   "LblIdCtaDeb"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   8295
               TabIndex        =   32
               Top             =   1305
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.Label LblIdCtaHab 
               AutoSize        =   -1  'True
               Caption         =   "LblIdCtaHab"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   8325
               TabIndex        =   31
               Top             =   1650
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.Label LblDescHaber 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDescHaber"
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
               Left            =   3945
               TabIndex        =   30
               Top             =   1605
               Width           =   5565
            End
            Begin VB.Label LblDescDebe 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDescDebe"
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
               Left            =   3945
               TabIndex        =   29
               Top             =   1245
               Width           =   5565
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Cta Destino Haber"
               Height          =   195
               Left            =   915
               TabIndex        =   23
               Top             =   1635
               Width           =   1305
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cta Destino Debe"
               Height          =   195
               Left            =   915
               TabIndex        =   22
               Top             =   1275
               Width           =   1260
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Descripción"
               Height          =   195
               Index           =   1
               Left            =   915
               TabIndex        =   19
               Top             =   915
               Width           =   840
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Nº Cuenta"
               Height          =   195
               Left            =   915
               TabIndex        =   18
               Top             =   555
               Width           =   735
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Cuenta"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   60
            TabIndex        =   20
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   13
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   14
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "ID"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nº Cuenta"
            Columns(1).DataField=   "cuenta"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripción"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Dest. Debe"
            Columns(3).DataField=   "ctadeb"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Dest. Haber"
            Columns(4).DataField=   "ctahab"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Distribución"
            Columns(5).DataField=   "des"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "2da Distribución"
            Columns(6).DataField=   "des2"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   16
            Columns(7)._MaxComboItems=   5
            Columns(7).ValueItems(0)._DefaultItem=   0
            Columns(7).ValueItems(0).Value=   "1"
            Columns(7).ValueItems(0).Value.vt=   8
            Columns(7).ValueItems(0).DisplayValue=   "No"
            Columns(7).ValueItems(0).DisplayValue.vt=   8
            Columns(7).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(7).ValueItems(1)._DefaultItem=   0
            Columns(7).ValueItems(1).Value=   "0"
            Columns(7).ValueItems(1).Value.vt=   8
            Columns(7).ValueItems(1).DisplayValue=   "Si"
            Columns(7).ValueItems(1).DisplayValue.vt=   8
            Columns(7).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(7).ValueItems.Count=   2
            Columns(7).Caption=   "Es Divisionaria"
            Columns(7).DataField=   "tipo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2275"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2196"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=7594"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=7514"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1773"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1693"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1693"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1614"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2223"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2143"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=2143"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2064"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=1958"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1879"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
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
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Plan de Cuentas General Revisado"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   15
            Top             =   30
            Width           =   11610
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
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
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   9
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar Excel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmManPlanCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstCta As New ADODB.Recordset
Dim QueHace As Integer
Private SeEjecuto As Boolean
Dim xIdCtaAct As Integer
Dim xHorIni As Date

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO



Sub Filtrar()
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(5, 4) As String
   
    xCampos(0, 0) = "Nº Cuenta":            xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "C":         xCampos(0, 3) = "4200"
    xCampos(1, 0) = "Descripcion":          xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Cta. Destino Debe":    xCampos(2, 1) = "ctadeb":        xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Destino Haber":        xCampos(3, 1) = "ctahab":        xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Distribución":         xCampos(4, 1) = "des":           xCampos(4, 2) = "C":         xCampos(4, 3) = "1500"
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstCta, xCampos(), Dg1

End Sub

Private Sub CmdBusDesDebe_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nº Cuenta":    xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xform.SqlCad = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Cuenta Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtDebe.Text = xRs("cuenta")
        LblDescDebe.Caption = xRs("descripcion")
        LblIdCtaDeb.Caption = xRs("id")
        TxtHaber.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDesHaber_Click()
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nº Cuenta":    xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xform.SqlCad = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Cuenta Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtHaber.Text = xRs("cuenta")
        LblDescHaber.Caption = xRs("descripcion")
        LblIdCtaHab.Caption = xRs("id")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub CmdBusModulo_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1500":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xform.SqlCad = "SELECT * FROM tes_modulos"
    
    xform.Titulo = "Buscando Modulos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "id"
    xform.CampoBusca = "id"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtIdModulo.Text = xRs("id")
        LblDescModulo.Caption = xRs("descripcion")
        TxtNumCta.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstCta
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)

    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstCta.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear

End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstCta("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
'Modificado: 10/01/11 Johan Castro
'            Agregar linea de codigo para bloquear accesos de usuarios
'            Se elimina linea de codigo: CierrePeriodo Toolbar1, 26, 0, False, xCon, xIdUsuario


    On Error GoTo error
    If SeEjecuto = False Then
'        Dim Rpta As Integer
        
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '----------------------------------------------

        
        RST_Busq RstCta, "SELECT con_planctas.*, con_planctas_1.cuenta AS ctadeb, con_planctas_1.descripcion AS desctadeb, con_planctas_2.cuenta AS ctahab, con_planctas_2.descripcion AS desctahab, con_planctasdes.descripcion AS des, con_planctasdes_1.descripcion AS des2, con_planctasdes_2.descripcion AS des3, IIf([con_planctas].[tipo]=0,'Si','No') AS esdiv,IIF(con_planctas.desctabal =1,'A',IIF(con_planctas.desctabal=2,'P',IIF(con_planctas.desctabal IN (1,2),'A - P',''))) AS desbal " _
                + vbCr + " FROM ((((con_planctas LEFT JOIN con_planctas AS con_planctas_1 ON con_planctas.ctadesdeb = con_planctas_1.id) LEFT JOIN con_planctas AS con_planctas_2 ON con_planctas.ctadeshab = con_planctas_2.id) LEFT JOIN con_planctasdes ON con_planctas.iddes = con_planctasdes.id) LEFT JOIN con_planctasdes AS con_planctasdes_1 ON con_planctas.iddes2 = con_planctasdes_1.id) LEFT JOIN con_planctasdes AS con_planctasdes_2 ON con_planctas.iddes3 = con_planctasdes_2.id  " _
                + vbCr + " WHERE con_planctas.id not in (0) " _
                + vbCr + " ORDER BY con_planctas.cuenta ASC", xCon

        Set Dg1.DataSource = RstCta
    
    End If
    Exit Sub
error:
    SHOW_ERROR Me.Name, "Form_Activate"
End Sub

Sub Nuevo()
    Bloquea
    Blanquea
    ActivaTool
    QueHace = 1
    Label5.Caption = "Agregando Cuenta Contable"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    xHorIni = Time
    TxtNumCta.SetFocus
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

Sub Blanquea()
    TxtNumCta.Text = ""
    txtdescripcion.Text = ""
    TxtDebe.Text = ""
    TxtHaber.Text = ""
    TxtIdModulo.Text = ""
    LblDescDebe.Caption = ""
    LblDescHaber.Caption = ""
    LblDescModulo.Caption = ""
    
    ChkActivo.Value = 0
    ChkPasivo.Value = 0
    
    LimpiaText lbl_cb_cod
    LimpiaText txt_cb
End Sub

Sub Bloquea()
    TxtNumCta.Locked = Not TxtNumCta.Locked
    txtdescripcion.Locked = Not txtdescripcion.Locked
    TxtDebe.Locked = Not TxtDebe.Locked
    TxtHaber.Locked = Not TxtHaber.Locked
    TxtIdModulo.Locked = Not TxtIdModulo.Locked
    
    habilitar_Locked txt_cb, Not txt_cb(0).Locked
    habilitar fra, Not fra(0).Enabled
    
    Frame4.Enabled = Not Frame4.Enabled
    
End Sub

Private Sub Form_Load()
    QueHace = 3
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
End Sub

Sub MuestraSegundoTab()
    Blanquea
    If RstCta.State = 0 Then Exit Sub
    If RstCta.EOF = True Or RstCta.BOF = True Or RstCta.RecordCount = 0 Then Exit Sub
    TxtNumCta.Text = RstCta("cuenta") & ""
    txtdescripcion.Text = RstCta("descripcion") & ""
    
    If NulosN(RstCta("idmodulo")) = 0 Then
        TxtIdModulo.Text = ""
        LblDescModulo.Caption = ""
    Else
        TxtIdModulo.Text = RstCta("idmodulo")
        LblDescModulo.Caption = Busca_Codigo(TxtIdModulo.Text, "id", "descripcion", "tes_modulos", "N", xCon)
    End If
    
    If RstCta("documentar") = True Then
        OptSI1.Value = True
        OptNO1.Value = False
    Else
        OptSI1.Value = False
        OptNO1.Value = True
    End If
    
    If NulosN(RstCta("ctadesdeb")) <> 0 Then
        TxtDebe.Text = RstCta("ctadeb") & ""
        LblDescDebe.Caption = RstCta("desctadeb") & ""
        LblIdCtaDeb.Caption = RstCta("ctadesdeb") & ""
    End If
    
    If NulosN(RstCta("ctadeshab")) <> 0 Then
        TxtHaber.Text = RstCta("ctahab") & ""
        LblDescHaber.Caption = RstCta("desctahab") & ""
        LblIdCtaHab.Caption = RstCta("ctadeshab") & ""
    End If
    
    '---
    If RstCta.Fields("iddes") & "" <> "0" Then
        txt_cb(0).Text = RstCta.Fields("iddes") & ""
        lbl_cb(0).Caption = RstCta.Fields("des") & ""
        lbl_cb_cod(0).Caption = RstCta.Fields("iddes") & ""
    End If
    If RstCta.Fields("iddes2") & "" <> "0" Then
        txt_cb(1).Text = RstCta.Fields("iddes2") & ""
        lbl_cb(1).Caption = RstCta.Fields("des2") & ""
        lbl_cb_cod(1).Caption = RstCta.Fields("iddes2") & ""
    End If
    If RstCta.Fields("iddes3") & "" <> "0" Then
        txt_cb(2).Text = RstCta.Fields("iddes3") & ""
        lbl_cb(2).Caption = RstCta.Fields("des3") & ""
        lbl_cb_cod(2).Caption = RstCta.Fields("iddes3") & ""
    End If
    If UCase(RstCta.Fields("tipsal") & "") = "D" Then
        opt_saldo(0).Value = True
    ElseIf UCase(RstCta.Fields("tipsal") & "") = "H" Then
        opt_saldo(1).Value = True
    Else
        opt_saldo(0).Value = False
        opt_saldo(1).Value = False
    End If
    
    If RstCta.Fields("desctabal") = 3 Then
        ChkActivo.Value = 1
        ChkPasivo.Value = 1
        
    ElseIf RstCta.Fields("desctabal") = 1 Then
        ChkActivo.Value = 1
        
    ElseIf RstCta.Fields("desctabal") = 2 Then
        ChkPasivo.Value = 1
    End If
    

    
End Sub

Sub Cancelar()
    QueHace = 3
    Bloquea
    ActivaTool
    Label5.Caption = "Detalle de la Cuenta"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario mientras este ingresando o modificando una cuenta contable", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    Else
        Set RstCta = Nothing
        SeEjecuto = False
    End If
End Sub


Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar

    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstCta.Requery
            
            TDB_FiltroLimpiar Dg1
            RstCta.Filter = adFilterNone
            
            Dg1.Refresh
            
            RstCta.MoveFirst
            RstCta.Find "id = " & xIdCtaAct & ""
            If RstCta.EOF = True Then
                RstCta.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then
        If RstCta.State = 0 Then Exit Sub
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstCta.Filter = adFilterNone
        RstCta.Requery
        Dg1.Refresh
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 12 Then Exportar
    
    If Button.Index = 15 Then
        Set RstCta = Nothing
        Unload Me
    End If
End Sub

Sub Exportar()
    Dim oExport As New SGI2_funciones.formularios
    Dim Rst As New ADODB.Recordset
    
    Dim xCampos(11, 3) As String
    
    TabOne1.CurrTab = 0
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Id":                 xCampos(0, 1) = "id":           xCampos(0, 2) = 2:  xCampos(0, 3) = "450"
    xCampos(1, 0) = "Cuenta":             xCampos(1, 1) = "cuenta":       xCampos(1, 2) = 0:  xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Descripcion":        xCampos(2, 1) = "descripcion":  xCampos(2, 2) = 0:  xCampos(2, 3) = "4000"
    xCampos(3, 0) = "Cta. Dest. Debe":    xCampos(3, 1) = "ctadeb":       xCampos(3, 2) = 0:  xCampos(3, 3) = "1000"
    xCampos(4, 0) = "Descripción":        xCampos(4, 1) = "desctadeb":    xCampos(4, 2) = 0:  xCampos(4, 3) = "3500"
    xCampos(5, 0) = "Cta. Dest. Haber":   xCampos(5, 1) = "ctahab":       xCampos(5, 2) = 0:  xCampos(5, 3) = "1000"
    xCampos(6, 0) = "Descripción":        xCampos(6, 1) = "desctadeb":    xCampos(6, 2) = 0:  xCampos(6, 3) = "3500"
        
    xCampos(7, 0) = "1ra Distribución":   xCampos(7, 1) = "des":          xCampos(7, 2) = 0:  xCampos(7, 3) = "1500"
    xCampos(8, 0) = "2da Distribución":   xCampos(8, 1) = "des2":         xCampos(8, 2) = 0:  xCampos(8, 3) = "1500"
    xCampos(9, 0) = "Nat. Saldo":         xCampos(9, 1) = "tipsal":       xCampos(9, 2) = 1:  xCampos(9, 3) = "1000"
    xCampos(10, 0) = "Divisionaria":       xCampos(10, 1) = "esdiv":        xCampos(10, 2) = 1:  xCampos(10, 3) = "1200"
    
    xCampos(11, 0) = "Destino Balance":     xCampos(11, 1) = "desbal":       xCampos(11, 2) = 1:  xCampos(11, 3) = "1000"
        
    Set Rst = RstCta.Clone
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Plan de Cuentas", "", "", "Plan de Cuentas", Rst, xCampos
    Set oExport = Nothing
    Set Rst = Nothing
    Dg1.Refresh
End Sub

Function Grabar() As Boolean

    If VALIDAR_DATOS() = False Then Exit Function
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modficar") + " la Cuenta Contable", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim xId As Double

On Error GoTo LaCague

    xCon.BeginTrans
    If QueHace = 1 Then
        xId = HallaCodigoTabla("con_planctas", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_planctas", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstCta("id")
        RST_Busq RstCab, "SELECT * FROM con_planctas WHERE id = " & RstCta("id") & "", xCon
    End If
    
    xIdCtaAct = xId
    
    RstCab("cuenta") = Trim(TxtNumCta.Text)
    RstCab("descripcion") = Trim(txtdescripcion.Text)
    
    If NulosN(LblIdCtaHab.Caption) <> 0 Then
        RstCab("ctadesdeb") = Val(LblIdCtaDeb.Caption)
    Else
        RstCab("ctadesdeb") = 0
    End If
    If NulosN(LblIdCtaDeb.Caption) <> 0 Then
        RstCab("ctadeshab") = Val(LblIdCtaHab.Caption)
    Else
        RstCab("ctadeshab") = 0
    End If
    '-----
    RstCab("iddes") = NulosN(lbl_cb_cod(0).Caption)
    RstCab("iddes2") = NulosN(lbl_cb_cod(1).Caption)
    RstCab("iddes3") = NulosN(lbl_cb_cod(2).Caption)
    RstCab("idmodulo") = NulosN(TxtIdModulo.Text)
    
    '----
    If OptSI1.Value = True Then RstCab("documentar") = -1
    If OptNO1.Value = True Then RstCab("documentar") = 0
       
    
    '--Distribucion de la cuenta cuando la distribucion es balance general
    If NulosN(txt_cb(0).Text) = 1 Or NulosN(txt_cb(1).Text) = 1 Then
        If ChkActivo.Value = 1 And ChkPasivo.Value = 1 Then
            RstCab("desctabal") = 3
        ElseIf ChkActivo.Value = 1 Then
            RstCab("desctabal") = 1
        ElseIf ChkPasivo.Value = 1 Then
            RstCab("desctabal") = 2
        Else
            RstCab("desctabal") = 0
        End If
    Else
        RstCab("desctabal") = 0
    End If
    '----
    
    
    '-----DEL TIPO   1 = cuentas; 0 = registro
    '--SI DEPENDE DE OTRA CUENTA
    Dim xRs As New ADODB.Recordset
    Dim N_CTA As String
    Dim Q_POS As Integer
    N_CTA = StrReverse(Trim(TxtNumCta.Text))
    Q_POS = InStr(N_CTA, "-")
    If Q_POS <> 0 Then
        N_CTA = StrReverse(Mid(N_CTA, Q_POS + 1))
        RST_Busq xRs, "SELECT id FROM con_planctas WHERE (((cuenta)= '" + N_CTA + "'));", xCon
        If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
            xCon.Execute "UPDATE con_planctas SET tipo = 1 WHERE id = " + CStr(xRs.Fields("id")) '--CUENTA
        End If
    End If
    '--SI TIENEN CUENTAS QUE DEPENDEN DE ESTE
    Set xRs = Nothing
    N_CTA = Trim(TxtNumCta.Text)
    RST_Busq xRs, "SELECT con_planctas.id, con_planctas.cuenta FROM con_planctas WHERE (((con_planctas.id)<>" + CStr(xId) + ") AND ((con_planctas.cuenta) Like '" + N_CTA + "%'));", xCon
    If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
        RstCab("tipo") = 1 '--ES CUENTA
    Else
        RstCab("tipo") = 0 '--ES REGISTRO
    End If
    '-----
    RstCab("tipsal") = IIf(opt_saldo(0).Value = True, "D", IIf(opt_saldo(1).Value = True, "H", ""))
    
    RstCab.Update
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    Set RstCab = Nothing:       Set xRs = Nothing
    MsgBox "La Cuenta Contable se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    xCon.CommitTrans
    Grabar = True
    Exit Function
        
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing:       Set xRs = Nothing
    MsgBox "No se pudo guardar la cuenta contable por el siguiente motivo: " + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
End Function


Sub Modificar()
    Bloquea
    Blanquea
    ActivaTool
    QueHace = 2
    Label5.Caption = "Modificando Cuenta Contable"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    MuestraSegundoTab
    xHorIni = Time
    TxtNumCta.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim xRs  As New ADODB.Recordset
    '--SI TIENE DIVISIONARIAS NO ELIMINAR
    RST_Busq xRs, "SELECT id, cuenta FROM con_planctas WHERE (((id)<>" + CStr(RstCta.Fields("id")) + ") AND ((cuenta) Like '" + RstCta.Fields("cuenta") + "%'));", xCon
    If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
        MsgBox "El registro tiene Divisionaria" + vbCr + "Elimine las Divisionarias primero", vbExclamation, xTitulo
        Set xRs = Nothing
        Exit Sub
    End If
    '
    '--SI ESTA ASOCIADO A LIBRO DIARIO NO ELIMINAR
    RST_Busq xRs, "SELECT con_diario.idcue FROM con_diario WHERE (((con_diario.idcue)=" + CStr(RstCta("id")) + "));", xCon
    If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
        MsgBox "El registro no se puede eliminar" + vbCr + "Esta asociado a Libro Diario", vbExclamation, xTitulo
        Set xRs = Nothing
        Exit Sub
    End If
    Set xRs = Nothing
    '-------------
    Rpta = MsgBox("Esta seguro de eliminar la cuenta contable seleccionada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM con_planctas WHERE id =" & RstCta("id") & " "
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstCta("id") & " AND idform = " & IdMenuActivo
        
        MsgBox "La cuenta contable se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstCta.Requery
        Dg1.Refresh
    End If
End Sub

Private Sub txt_cb_LostFocus(Index As Integer)
    txt_cb_KeyDown Index, 13, 0
End Sub

Private Sub TxtDebe_Change()
    If TxtDebe.Text = "" Then
        Me.LblDescDebe.Caption = ""
        Me.LblIdCtaDeb.Caption = ""
    End If
End Sub

Private Sub TxtDebe_KeyDown(KeyCode As Integer, Shift As Integer)
    If TxtDebe.Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        CmdBusDesDebe_Click
        Exit Sub
    End If
    If TxtDebe.Text = "" Then Exit Sub
    If KeyCode <> 13 Then Exit Sub
    If xCon.State = 0 Then Exit Sub
    Dim RST_TMP As New ADODB.Recordset
    Dim N_SQL As String
    On Error GoTo error

    N_SQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id " _
        + vbCr + " FROM con_planctas " _
        + vbCr + " WHERE (((con_planctas.cuenta)='" + Trim(TxtDebe.Text) + "'));"

    RST_Busq RST_TMP, N_SQL, xCon
    
    If RST_TMP.State = 0 Then GoTo SALIR
    If RST_TMP.RecordCount > 0 Then
        TxtDebe.Text = RST_TMP.Fields(0) & ""  '--TEXTO A MOSTRAR
        LblDescDebe.Caption = RST_TMP.Fields(1) & "" '--NOMBRE
        LblIdCtaDeb.Caption = RST_TMP.Fields(2) & "" '--CODIGO
    Else
        TxtDebe.Text = "":    LblDescDebe.Caption = "":    LblIdCtaDeb.Caption = ""
    End If
SALIR:
    Set RST_TMP = Nothing
    Exit Sub
error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "TxtDebe_KeyDown"
End Sub

Private Sub TxtDebe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDebe_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        TxtDebe.Text = ""
        LblDescDebe.Caption = ""
        LblIdCtaDeb.Caption = ""
    End If
End Sub

Private Sub TxtDebe_LostFocus()
    TxtDebe_KeyDown 13, 0
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtHaber_Change()
    If TxtHaber.Text = "" Then
        Me.LblDescHaber.Caption = ""
        Me.LblIdCtaHab.Caption = ""
    End If
End Sub

Private Sub TxtHaber_KeyDown(KeyCode As Integer, Shift As Integer)
    If TxtHaber.Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        CmdBusDesHaber_Click
        Exit Sub
    End If
    If TxtHaber.Text = "" Then Exit Sub
    If KeyCode <> 13 Then Exit Sub
    If xCon.State = 0 Then Exit Sub
    Dim RST_TMP As New ADODB.Recordset
    Dim N_SQL As String
    On Error GoTo error

    N_SQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id " _
        + vbCr + " FROM con_planctas " _
        + vbCr + " WHERE (((con_planctas.cuenta)='" + Trim(TxtHaber.Text) + "'));"

    RST_Busq RST_TMP, N_SQL, xCon
    
    If RST_TMP.State = 0 Then GoTo SALIR
    If RST_TMP.RecordCount > 0 Then
        TxtHaber.Text = RST_TMP.Fields(0) & ""  '--TEXTO A MOSTRAR
        LblDescHaber.Caption = RST_TMP.Fields(1) & "" '--NOMBRE
        LblIdCtaHab.Caption = RST_TMP.Fields(2) & "" '--CODIGO
    Else
        TxtHaber.Text = "":    LblDescHaber.Caption = "":    LblIdCtaHab.Caption = ""
    End If
SALIR:
    Set RST_TMP = Nothing
    Exit Sub
error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "TxtDebe_KeyDown"
End Sub

Private Sub TxtHaber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtHaber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        TxtHaber.Text = ""
        LblDescHaber.Caption = ""
        LblIdCtaHab.Caption = ""
    End If
End Sub

Private Sub TxtHaber_LostFocus()
    TxtHaber_KeyDown 13, 0
End Sub

Private Sub TxtIdModulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdModulo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusModulo_Click
    End If
End Sub

Private Sub TxtIdModulo_Validate(Cancel As Boolean)
    If NulosN(TxtIdModulo.Text) = 0 Then
        LblDescModulo.Caption = ""
        Exit Sub
    End If
    
    LblDescModulo.Caption = Busca_Codigo(TxtIdModulo.Text, "id", "descripcion", "tes_modulos", "N", xCon)
    If LblDescModulo.Caption = "" Then
        TxtIdModulo.Text = ""
    End If
End Sub

Private Sub TxtNumCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Sub Buscar()
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Nº Cuenta":   xCampos2(0, 1) = "cuenta":          xCampos2(0, 2) = "1500":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Descripcion": xCampos2(1, 1) = "descripcion":     xCampos2(1, 2) = "6500":         xCampos2(1, 3) = "C"
    
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT con_planctas.* From con_planctas ORDER BY con_planctas.descripcion", xCampos2(), "Buscando Cuenta Contable", "cuenta", "cuenta", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    RstCta.MoveFirst
    RstCta.Find "id = " & xRs("id") & ""
SALIR:
    Set xRs = Nothing
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub



'---MODIFICADO AL 04/12/07  03:10 PM

Private Sub cb_Click(Index As Integer)
    On Error GoTo error
    Dim N_SQL As String
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    
    If QueHace = 3 Then Exit Sub
    
    ReDim xCampos(1, 3) As String
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    
    N_SQL = "SELECT con_planctasdes.id, con_planctasdes.descripcion AS nombre, con_planctasdes.id AS cod " _
        + vbCr + " FROM con_planctasdes " _
        + vbCr + " ORDER BY con_planctasdes.descripcion;"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, N_SQL, xCampos(), lbl_cb_capt(Index).Caption, "nombre", "nombre", Principio
    If xRs.BOF = True Or xRs.EOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    If xRs.State = 0 Then GoTo SALIR
    
    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
   
SALIR:
    Set xRs = Nothing
Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click (" + CStr(Index) + ")"
End Sub



Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cb_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
    If txt_cb(Index).Text = "" Then Exit Sub
    If KeyCode <> 13 Then Exit Sub
    Dim RST_TMP As New ADODB.Recordset
    Dim N_SQL As String
    On Error GoTo error

    N_SQL = "SELECT con_planctasdes.id, con_planctasdes.descripcion AS nombre, con_planctasdes.id AS cod " _
        + vbCr + " FROM con_planctasdes " _
        + vbCr + " WHERE con_planctasdes.id=" + CStr(Trim(txt_cb(Index).Text))
        
    If xCon.State = 0 Then Exit Sub
    RST_Busq RST_TMP, N_SQL, xCon
    
    If RST_TMP.State = 0 Then GoTo SALIR
    If RST_TMP.RecordCount > 0 Then
        txt_cb(Index) = RST_TMP.Fields(0) & "" '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RST_TMP.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = RST_TMP.Fields(2) & "" '--CODIGO
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
    End If
SALIR:
    Set RST_TMP = Nothing
    Exit Sub
error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown (" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Or KeyAscii = 46 Then KeyAscii = 0
End Sub


Private Function VALIDAR_DATOS() As Boolean
    '--VALIDAR QUE LA GRILLA DE ACTIVO Y PASIVO TENGAN VALORES TANTO DE ORDEN Y DESCRIPCION
    
    If TxtNumCta.Text = "" Then
        MsgBox "No ha especificado el número de la cuenta", vbInformation, xTitulo
        TxtNumCta.SetFocus
        Exit Function
    End If

    If txtdescripcion.Text = "" Then
        MsgBox "No ha especificado la descripción de la cuenta", vbInformation, xTitulo
        txtdescripcion.SetFocus
        Exit Function
    End If
    '--------------------------------
    '--VALIDAR QUE EL REGISTRO NO ESTE REGISTRADO
    Dim RstTmp As New ADODB.Recordset
    If QueHace = 1 Then
        RST_Busq RstTmp, "SELECT descripcion FROM con_planctas WHERE ucase(cuenta)='" + UCase(Trim(TxtNumCta.Text)) + "';", xCon
    Else
        RST_Busq RstTmp, "SELECT descripcion FROM con_planctas WHERE ucase(cuenta)='" + UCase(Trim(TxtNumCta.Text)) + "' AND id <> " + CStr(RstCta.Fields("id")) + ";", xCon
    End If
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        MsgBox "La Cuenta Contable " + IIf(QueHace = 1, " ya fue ingresado", "ya existe"), vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    
    If QueHace <> 1 And RstCta.Fields("cuenta") & "" <> Trim(Me.TxtNumCta.Text) Then
        RST_Busq RstTmp, "SELECT con_planctas.id, con_planctas.cuenta FROM con_planctas WHERE (((con_planctas.id)<>" + CStr(RstCta.Fields("id")) + ") AND ((con_planctas.cuenta) Like '" + RstCta.Fields("cuenta") + "%'));", xCon
        If RstTmp.EOF = False Or RstTmp.BOF = False Or RstTmp.RecordCount <> 0 Then
            MsgBox "La Cuenta Contable tiene Divisionaria" + vbCr + "Modifique las Divisionarias primero " + vbCr + "Nº de Cuenta Contable: " + RstCta.Fields("cuenta") & "", vbExclamation, xTitulo
            Me.TxtNumCta.Text = RstCta.Fields("cuenta") & ""
            Set RstTmp = Nothing
            Exit Function
        End If
    End If
    Set RstTmp = Nothing
    '-----
    VALIDAR_DATOS = True
End Function
 


