VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmControlDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punto de Venta - Control de Documentos"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6570
      Top             =   45
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
            Picture         =   "FrmControlDoc.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDoc.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
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
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
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
            Object.ToolTipText     =   "Imprimir Listado"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6165
      Left            =   15
      TabIndex        =   10
      Top             =   375
      Width           =   9840
      _cx             =   17357
      _cy             =   10874
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5745
         Left            =   10485
         TabIndex        =   14
         Top             =   375
         Width           =   9750
         Begin VB.Frame fra 
            Height          =   3780
            Index           =   1
            Left            =   555
            TabIndex        =   16
            Top             =   660
            Width           =   8340
            Begin VB.Frame fra 
               Caption         =   "[Nº Documento]"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   1755
               Index           =   2
               Left            =   180
               TabIndex        =   24
               Top             =   1050
               Width           =   7860
               Begin VB.CommandButton cb 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   2
                  Left            =   2535
                  Picture         =   "FrmControlDoc.frx":277E
                  Style           =   1  'Graphical
                  TabIndex        =   5
                  ToolTipText     =   "Seleccione el Nº el Número de Documento"
                  Top             =   270
                  Width           =   225
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   2
                  Left            =   1500
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   4
                  Text            =   "txt_cb(2)"
                  ToolTipText     =   "Ingrese el Nº de Documento"
                  Top             =   240
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº Serie"
                  Height          =   195
                  Index           =   5
                  Left            =   3030
                  TabIndex        =   38
                  Top             =   345
                  Width           =   585
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(1)"
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
                  Left            =   1500
                  TabIndex        =   37
                  Top             =   930
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "Num.Reg."
                  Height          =   195
                  Index           =   1
                  Left            =   225
                  TabIndex        =   36
                  Top             =   1035
                  Width           =   720
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(4)"
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
                  Index           =   4
                  Left            =   3750
                  TabIndex        =   35
                  Top             =   1275
                  Width           =   1890
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(3)"
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
                  Index           =   3
                  Left            =   1500
                  TabIndex        =   34
                  Top             =   1275
                  Width           =   1290
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(2)"
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
                  Left            =   3750
                  TabIndex        =   33
                  Top             =   930
                  Width           =   1890
               End
               Begin VB.Label lbl_dato 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_dato(0)"
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
                  Left            =   1500
                  TabIndex        =   32
                  Top             =   585
                  Width           =   6120
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "Importe"
                  Height          =   195
                  Index           =   4
                  Left            =   3030
                  TabIndex        =   31
                  Top             =   1380
                  Width           =   525
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "Moneda"
                  Height          =   195
                  Index           =   3
                  Left            =   240
                  TabIndex        =   30
                  Top             =   1380
                  Width           =   585
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch.Doc"
                  Height          =   195
                  Index           =   2
                  Left            =   3030
                  TabIndex        =   29
                  Top             =   1035
                  Width           =   615
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "Cliente"
                  Height          =   195
                  Index           =   0
                  Left            =   225
                  TabIndex        =   28
                  Top             =   690
                  Width           =   480
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº Documento"
                  Height          =   195
                  Index           =   2
                  Left            =   255
                  TabIndex        =   27
                  Top             =   345
                  Width           =   1050
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
                  Left            =   3750
                  TabIndex        =   25
                  Top             =   240
                  Width           =   1890
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
                  Left            =   6375
                  TabIndex        =   26
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1290
               End
            End
            Begin VB.CommandButton cb 
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   2640
               Picture         =   "FrmControlDoc.frx":28B0
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Seleccione el Documento"
               Top             =   660
               Width           =   225
            End
            Begin VB.CommandButton cb 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   2640
               Picture         =   "FrmControlDoc.frx":29E2
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Seleccione el Almancén"
               Top             =   285
               Width           =   225
            End
            Begin VB.Frame fra 
               Caption         =   "[ Acción a Tomar ]"
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
               ForeColor       =   &H00400000&
               Height          =   585
               Index           =   0
               Left            =   180
               TabIndex        =   17
               Top             =   2895
               Width           =   7860
               Begin VB.OptionButton opt_evento 
                  Caption         =   "&Anular"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Index           =   2
                  Left            =   6000
                  TabIndex        =   8
                  Top             =   255
                  Width           =   1530
               End
               Begin VB.OptionButton opt_evento 
                  Caption         =   "&Eliminar"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Index           =   1
                  Left            =   3127
                  TabIndex        =   7
                  Top             =   255
                  Width           =   1530
               End
               Begin VB.OptionButton opt_evento 
                  Caption         =   "&Modificar"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   240
                  Index           =   0
                  Left            =   255
                  TabIndex        =   6
                  Top             =   255
                  Value           =   -1  'True
                  Width           =   1530
               End
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   0
               Text            =   "txt_cb(0)"
               ToolTipText     =   "Ingrese el Código del Almacen"
               Top             =   255
               Width           =   1215
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   1
               Left            =   1680
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   2
               Text            =   "txt_cb(1)"
               ToolTipText     =   "Ingrese el código del documento"
               Top             =   630
               Width           =   1215
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
               Left            =   2910
               TabIndex        =   22
               Top             =   630
               Width           =   4905
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
               Left            =   2910
               TabIndex        =   19
               Top             =   255
               Width           =   4905
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
               Left            =   6705
               TabIndex        =   23
               Top             =   630
               Visible         =   0   'False
               Width           =   1290
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Documento"
               Height          =   195
               Index           =   1
               Left            =   210
               TabIndex        =   21
               Top             =   720
               Width           =   1185
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Almacén"
               Height          =   195
               Index           =   0
               Left            =   210
               TabIndex        =   20
               Top             =   360
               Width           =   615
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
               Left            =   6705
               TabIndex        =   18
               Top             =   255
               Visible         =   0   'False
               Width           =   1290
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Control de Doccumento"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   90
            TabIndex        =   15
            Top             =   45
            Width           =   9075
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5745
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   9750
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   5400
            Left            =   30
            TabIndex        =   12
            Top             =   345
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   9525
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Num.Reg."
            Columns(0).DataField=   "numreg"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "TD"
            Columns(1).DataField=   "abrev"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numerodoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch.Emi."
            Columns(3).DataField=   "fchdoc"
            Columns(3).NumberFormat=   "Long Date"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Cliente"
            Columns(4).DataField=   "clidesc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   16
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "M"
            Columns(5).DataField=   "simbolo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Importe"
            Columns(6).DataField=   "imptotdoc"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   16
            Columns(7)._MaxComboItems=   5
            Columns(7).ValueItems(0)._DefaultItem=   0
            Columns(7).ValueItems(0).Value=   "1"
            Columns(7).ValueItems(0).Value.vt=   8
            Columns(7).ValueItems(0).DisplayValue=   "Modificar"
            Columns(7).ValueItems(0).DisplayValue.vt=   8
            Columns(7).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(7).ValueItems(1)._DefaultItem=   0
            Columns(7).ValueItems(1).Value=   "2"
            Columns(7).ValueItems(1).Value.vt=   8
            Columns(7).ValueItems(1).DisplayValue=   "Eliminar"
            Columns(7).ValueItems(1).DisplayValue.vt=   8
            Columns(7).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(7).ValueItems(2)._DefaultItem=   0
            Columns(7).ValueItems(2).Value=   "3"
            Columns(7).ValueItems(2).Value.vt=   8
            Columns(7).ValueItems(2).DisplayValue=   "Anular"
            Columns(7).ValueItems(2).DisplayValue.vt=   8
            Columns(7).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
            Columns(7).ValueItems.Count=   3
            Columns(7).Caption=   "Pendiente de."
            Columns(7).DataField=   "evento"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=794"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=714"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2514"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2434"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1588"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1508"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=3043"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2963"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=953"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=873"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2514"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2434"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=2963"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2884"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
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
            HeadLines       =   1.25
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Control de Documento"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   13
            Top             =   45
            Width           =   9075
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
   End
End
Attribute VB_Name = "FrmControlDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim RstFrm As New ADODB.Recordset
Dim vStrSql As String
Dim Mostrando As Boolean
Dim SeEjecuto As Boolean

'-----------------------------------
Dim dFechaBusqueda As String '--INDICA LA FECHA ACTIVA


Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(1, 4) As String
    
    xCampos(0, 0) = "Empleado":     xCampos(0, 1) = "nomemp":     xCampos(0, 2) = "4000":    xCampos(0, 3) = "C"
        
    nSQL = "SELECT pvt_emp.id, pvt_emp.idemp, pla_empleados.numdoc, pla_empleados.numdoc, [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nomemp, pvt_emp.ven, pvt_emp.caj, pvt_emp.sup, pvt_emp.codigo, pvt_emp.idalm, alm_almacenes.descripcion AS almdesc, pvt_emp.idalm AS almcod " _
        + vbCr + " FROM (pla_empleados INNER JOIN pvt_emp ON pla_empleados.id = pvt_emp.idemp) LEFT JOIN alm_almacenes ON pvt_emp.idalm = alm_almacenes.id; "

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Documento", "nomemp", "nomemp", Principio
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub


Sub MuestraSegundoTab()
    If RstFrm.RecordCount = 0 Then Exit Sub
    Mostrando = True
    Blanquea
    '--DEL ALMACEN
    txt_cb(0).Text = RstFrm.Fields("idalm") & ""
    lbl_cb_cod(0).Caption = RstFrm("idalm") & ""
    lbl_cb(0).Caption = RstFrm("almdesc") & ""
    '--DEL TIPO DE DOCUMENTO
    txt_cb(1).Locked = False
    txt_cb(1).Text = RstFrm.Fields("tipdoc") & ""
    txt_cb_KeyDown 1, 13, 0
    txt_cb(1).Locked = True
    '--NUMERO DE DOCUMENTO
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    nSQL = "SELECT vta_ventas.numdoc, vta_ventas.numser, vta_ventas.id AS cod, [vta_ventas].[numser] & '-' & [vta_ventas].[numdoc] AS nombre, mae_documento.abrev, vta_ventas.numreg, vta_ventas.fchdoc, mae_cliente.nombre AS clidesc, mae_moneda.simbolo, vta_ventas.imptotdoc, vta_ventas.idalm " _
        + vbCr + " FROM ((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id " _
        + vbCr + " WHERE vta_ventas.id = " + CStr(NulosN(RstFrm.Fields("id")))
    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.RecordCount <> 0 Then
        txt_cb(2).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(2).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(2).Caption = RstTmp.Fields(2) & "" '--CODIGO

        lbl_dato(0).Caption = RstTmp.Fields("clidesc") & ""
        lbl_dato(1).Caption = RstTmp.Fields("numreg") & ""
        lbl_dato(2).Caption = RstTmp.Fields("fchdoc") & ""
        lbl_dato(3).Caption = RstTmp.Fields("simbolo") & ""
        lbl_dato(4).Caption = Format(NulosN(RstTmp.Fields("imptotdoc")), FORMAT_MONTO)
    End If
    Set RstTmp = Nothing
    
    '------
    Select Case NulosN(RstFrm.Fields("evento"))
        Case 1 '--MODIFICAR
            opt_evento(0).Value = True
        Case 2 '--ELIMINAR
            opt_evento(1).Value = True
        Case 3 '--ANULAR
            opt_evento(2).Value = True
    End Select
    Mostrando = False
    
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

Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked txt_cb, Not band
    habilitar Me.cb, band
    habilitar fra, band
    
    TabOne1.CurrTab = IIf(band = False, 0, 1)
    TabOne1.TabEnabled(0) = Not band
    
End Sub

Sub Blanquea()
    LimpiaText txt_cb
    opt_evento(0).Value = False:    opt_evento(1).Value = False:    opt_evento(2).Value = False
    
End Sub

Sub Cancelar()
    QueHace = 3
    Habilitar_Obj False
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
    Label5.Caption = "Detalle del Control de Documento"
End Sub

Function Grabar() As Boolean
    Dim xId, A As Integer
    
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modficar") + " el Registro", vbQuestion + vbYesNo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
            
    On Error GoTo LaCague
    xCon.BeginTrans
    Dim mAccion As Integer
    If opt_evento(0).Value = True Then mAccion = 1
    If opt_evento(1).Value = True Then mAccion = 2
    If opt_evento(2).Value = True Then mAccion = 3
    If QueHace = 1 Then
        xId = NulosN(lbl_cb_cod(2).Caption)
    Else 'MODIFICAR
        xId = RstFrm("id")
    End If
    RST_Busq RstCab, "SELECT * FROM vta_ventas WHERE id = " & xId & "", xCon
    
    RstCab("evento") = mAccion
    
    RstCab.Update
    
    xCon.CommitTrans
    Grabar = True
    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    Set RstCab = Nothing
    Label5.Caption = "Detalle del Control de Documento"
    Exit Function

LaCague:
    Set RstCab = Nothing
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo "
End Function

Sub Eliminar()
    On Error GoTo error
    Dim Rpta As Integer
    Rpta = MsgBox("¿Esta seguro de eliminar el registro seleccionado?", vbQuestion + vbYesNo, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "UPDATE vta_ventas SET evento =0 WHERE id = " & Val(RstFrm("id")) & ""
        RstFrm.Requery
        Dg3.Refresh
        MsgBox "Registro fue eliminado con éxito", vbInformation + vbOKOnly, xTitulo
    End If
    TabOne1.CurrTab = 0
    Exit Sub
error:
    SHOW_ERROR Me.Name, "Eliminar", True, "Error al eliminar..."
End Sub

Sub Modificar()
    
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    QueHace = 2
    ActivaTool
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    Habilitar_Obj True
    
    txt_cb(0).SetFocus
        
End Sub

Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Control de Documento"
    Habilitar_Obj True
    Blanquea
    txt_cb(0).SetFocus
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        
        vStrSql = "SELECT vta_ventas.id, vta_ventas.tipdoc, mae_documento.abrev, vta_ventas.numreg,vta_ventas.numdoc, vta_ventas.numser & '-' & vta_ventas.numdoc AS numerodoc, vta_ventas.fchdoc, mae_cliente.nombre AS clidesc, mae_moneda.simbolo, vta_ventas.imptotdoc, vta_ventas.evento, vta_ventas.idalm, alm_almacenes.descripcion AS almdesc " _
            + vbCr + " FROM (((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN alm_almacenes ON vta_ventas.idalm = alm_almacenes.id " _
            + vbCr + " WHERE (((vta_ventas.tipdoc) In (1,3)) AND ((vta_ventas.fchdoc)=CDate('" + dFechaBusqueda + "')) AND ((vta_ventas.evento) In (1,2,3)) AND ((vta_ventas.anulado)=0)) " _
            + vbCr + " ORDER BY vta_ventas.numser & '-' & vta_ventas.numdoc;"
           
        RST_Busq RstFrm, vStrSql, xCon
        
        Set Dg3.DataSource = RstFrm
        
        If RstFrm.RecordCount = 0 Then
            Dim Rpta As Integer
            Rpta = MsgBox("El registro esta vacio, ¿Desea agregar la el control a algún documento?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Blanquea

            End If
        End If
        
    End If
    
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Dg3.Columns("fchdoc").NumberFormat = FORMAT_DATE
    Mostrando = False
    dFechaBusqueda = "02/01/07" 'CStr(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set RstFrm = Nothing
End Sub



Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstFrm.Requery
            Dg3.Refresh
            Cancelar
        End If
    End If
    If Button.Index = 6 Then
        Cancelar
    End If
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 12 Then TDB_IMPRIMIR Dg3, "IMPRESIÓN", "LISTADO DE Documento"
        
    If Button.Index = 14 Then
        Unload Me
        Set RstFrm = Nothing
    End If
End Sub

Private Function fValidarDatos() As Boolean
    If Trim(lbl_cb_cod(0).Caption) = "" Then
        MsgBox "Seleccione el Almacén.", vbInformation, xTitulo
        cb(0).SetFocus
        Exit Function
    End If
    '----
    If Trim(lbl_cb_cod(1).Caption) = "" Then
        MsgBox "Seleccione Tipo de Documento.", vbInformation, xTitulo
        cb(1).SetFocus
        Exit Function
    End If
    If Trim(lbl_cb_cod(2).Caption) = "" Then
        MsgBox "Seleccione Número de Documento.", vbInformation, xTitulo
        cb(2).SetFocus
        Exit Function
    End If
    '--------------------------------
    If opt_evento(0).Value = False And opt_evento(1).Value = False And opt_evento(2).Value = False Then
        MsgBox "Seleccione que Acción va tomar con el documento?" + vbCr + "Modificar, Anular, Eliminar", vbExclamation, xTitulo
        Exit Function
    End If
    fValidarDatos = True
End Function


'-----------------------------
'-----------------------------

Private Sub cb_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nTitulo As String
    Dim mIdDoc As Integer '--INDICA EL DOCUMENTO (FACTURA O BOLETA)
    Select Case Index
        Case 0 '--ALMACEN
            ReDim xCampos(1, 3) As String
            xCampos(0, 0) = "Almacén":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            
            nSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion AS nombre, alm_almacenes.id AS cod " _
            + vbCr + " FROM alm_almacenes ORDER BY alm_almacenes.descripcion ;"
            
            nTitulo = "Buscando Almacén"
            
        Case 1 '--BUSCANDO EL TIPO DE DOCUMENTO
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "abrev":     xCampos(1, 2) = "550":    xCampos(1, 3) = "C"

            nSQL = "SELECT mae_documento.id, mae_documento.descripcion as nombre , mae_documento.id AS cod, mae_documento.abrev " _
                + vbCr + " FROM mae_documento " _
                + vbCr + " WHERE (((mae_documento.id) In (1,3))) " _
                + vbCr + " ORDER BY mae_documento.descripcion;"
                
            nTitulo = "Buscando Tipos de Docuentos"
            
        Case 2 '--DE LOS FORMATOS DE IMPRESION
            If NulosN(lbl_cb_cod(0).Caption) = 0 Then
                MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                cb(0).SetFocus
                Exit Sub
            End If
            If NulosN(lbl_cb_cod(1).Caption) = 0 Then
                MsgBox "Seleccione el Tipo de Documento", vbExclamation, xTitulo
                cb(1).SetFocus
                Exit Sub
            End If
            
            ReDim xCampos(6, 3) As String
            
            xCampos(0, 0) = "Nº Documento":     xCampos(0, 1) = "nombre":   xCampos(0, 2) = "1500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Num.Reg.":         xCampos(1, 1) = "numreg":   xCampos(1, 2) = "1000":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Fch.Doc.":         xCampos(2, 1) = "fchdoc":   xCampos(2, 2) = "900":    xCampos(2, 3) = "F"
            xCampos(3, 0) = "Cliente":          xCampos(3, 1) = "clidesc":  xCampos(3, 2) = "3380":   xCampos(3, 3) = "C"
            xCampos(4, 0) = "M":                xCampos(4, 1) = "simbolo":  xCampos(4, 2) = "500":    xCampos(4, 3) = "C"
            xCampos(5, 0) = "Importe":          xCampos(5, 1) = "imptotdoc":  xCampos(5, 2) = "800":    xCampos(5, 3) = "N"
            
            nSQL = "SELECT vta_ventas.numdoc, vta_ventas.numser, vta_ventas.id AS cod, [vta_ventas].[numser] & '-' & [vta_ventas].[numdoc] AS nombre, mae_documento.abrev, vta_ventas.numreg, format(vta_ventas.fchdoc,'dd/mm/yy') as fchdoc, mae_cliente.nombre AS clidesc, mae_moneda.simbolo, vta_ventas.imptotdoc, vta_ventas.idalm " _
                + vbCr + " FROM ((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id " _
                + vbCr + " WHERE (((vta_ventas.tipdoc)=" + CStr(lbl_cb_cod(1).Caption) + ") AND ((vta_ventas.fchdoc)=CDate('" + dFechaBusqueda + "')) AND ((vta_ventas.evento) Not In (1,2,3)) AND ((vta_ventas.idalm)=" + CStr(lbl_cb_cod(0).Caption) + ") AND ((vta_ventas.anulado)=0)) " _
                + vbCr + " ORDER BY [vta_ventas].[numser] & '-' & [vta_ventas].[numdoc];"

            
            nTitulo = "Buscando Documentos"
        
    End Select
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    If Index = 2 Then
        LimpiaText lbl_dato
        lbl_dato(0).Caption = xRs.Fields("clidesc") & ""
        lbl_dato(1).Caption = xRs.Fields("numreg") & ""
        lbl_dato(2).Caption = xRs.Fields("fchdoc") & ""
        lbl_dato(3).Caption = xRs.Fields("simbolo") & ""
        lbl_dato(4).Caption = Format(NulosN(xRs.Fields("imptotdoc")), FORMAT_MONTO)
    End If
    
    
Salir:
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub


Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cb_cod(Index).Caption = ""
        If Index = 0 Then txt_cb(1).Text = ""
        If Index = 1 Then txt_cb(2).Text = ""
        If Index = 2 Then LimpiaText lbl_dato
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If

    If txt_cb(Index).Text = "" Then Exit Sub
    If KeyCode <> 13 Then Exit Sub
    Dim RST_TMP As New ADODB.Recordset
    Dim nSQL As String
    Dim mIdDoc As Integer
    Select Case Index
        Case 0 '--ALMACEN
            nSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion AS nombre, alm_almacenes.id AS cod " _
                + vbCr + " FROM alm_almacenes WHERE alm_almacenes.id = " + CStr(NulosN(txt_cb(Index).Text)) + " ;"
                
        Case 1 '--TIPO DE DOCUMENTO
            nSQL = "SELECT mae_documento.id, mae_documento.descripcion , mae_documento.id AS cod, mae_documento.abrev " _
                + vbCr + " FROM mae_documento " _
                + vbCr + " WHERE mae_documento.id =" + CStr(NulosN(txt_cb(Index).Text)) + " AND  mae_documento.id In (1,3) "
                
        Case 2 '--NUMERO DE DOCUMENTO
            If NulosN(lbl_cb_cod(0).Caption) = 0 Then
                MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                cb(0).SetFocus
                Exit Sub
            End If
            If NulosN(lbl_cb_cod(1).Caption) = 0 Then
                MsgBox "Seleccione el Tipo de Documento", vbExclamation, xTitulo
                cb(1).SetFocus
                Exit Sub
            End If
                      
            nSQL = "SELECT vta_ventas.numdoc, vta_ventas.numser, vta_ventas.id AS cod, [vta_ventas].[numser] & '-' & [vta_ventas].[numdoc] AS nombre, mae_documento.abrev, vta_ventas.numreg, vta_ventas.fchdoc, mae_cliente.nombre AS clidesc, mae_moneda.simbolo, vta_ventas.imptotdoc, vta_ventas.idalm " _
                + vbCr + " FROM ((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id " _
                + vbCr + " WHERE FORMAT(vta_ventas.numdoc,'0000000000') = '" + Format(NulosN(txt_cb(Index).Text), "0000000000") + "' AND (((vta_ventas.tipdoc)=" + CStr(lbl_cb_cod(1).Caption) + ") AND ((vta_ventas.fchdoc)=CDate('" + dFechaBusqueda + "')) AND ((vta_ventas.evento) Not In (1,2,3)) AND ((vta_ventas.idalm)=" + CStr(lbl_cb_cod(0).Caption) + ") AND ((vta_ventas.anulado)=0)) "

            
    End Select
    If xCon.State = 0 Then Exit Sub
    RST_Busq RST_TMP, nSQL, xCon
    
    If RST_TMP.State = 0 Then Exit Sub
    If RST_TMP.RecordCount > 0 Then
        txt_cb(Index) = RST_TMP.Fields(0) & "" '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RST_TMP.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = RST_TMP.Fields(2) & "" '--CODIGO
        If Index = 2 Then '--DATOS DEL DOCUMENTO
            LimpiaText lbl_dato
            lbl_dato(0).Caption = RST_TMP.Fields("clidesc") & ""
            lbl_dato(1).Caption = RST_TMP.Fields("numreg") & ""
            lbl_dato(2).Caption = RST_TMP.Fields("fchdoc") & ""
            lbl_dato(3).Caption = RST_TMP.Fields("simbolo") & ""
            lbl_dato(4).Caption = Format(NulosN(RST_TMP.Fields("imptotdoc")), FORMAT_MONTO)
        End If
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
    End If
    Set RST_TMP = Nothing
    Exit Sub
error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    Select Case Index
        Case 0: If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case 1: If validar_numero(KeyAscii) = False Then KeyAscii = 0
        
    End Select
    
End Sub
'-----------------------------
'-----------------------------
Private Sub txt_cb_LostFocus(Index As Integer)
    txt_cb_KeyDown Index, 13, 0
End Sub

