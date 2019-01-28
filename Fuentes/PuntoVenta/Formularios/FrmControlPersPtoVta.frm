VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmControlPersPtoVta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punto de Venta - Control de Personal"
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
            Picture         =   "FrmControlPersPtoVta.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlPersPtoVta.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
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
      TabIndex        =   17
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
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5745
         Left            =   45
         TabIndex        =   21
         Top             =   375
         Width           =   9750
         Begin VB.Frame fra 
            Height          =   5205
            Index           =   1
            Left            =   465
            TabIndex        =   23
            Top             =   375
            Width           =   8520
            Begin VB.Frame fra 
               Caption         =   "[ Ticket de Venta ]"
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
               Height          =   1005
               Index           =   4
               Left            =   135
               TabIndex        =   46
               Top             =   4140
               Width           =   8175
               Begin VB.CommandButton cb 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   7
                  Left            =   2565
                  Picture         =   "FrmControlPersPtoVta.frx":277E
                  Style           =   1  'Graphical
                  TabIndex        =   48
                  ToolTipText     =   "Seleccione la Plantilla de Impresión"
                  Top             =   660
                  Width           =   225
               End
               Begin VB.CommandButton cb 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   6
                  Left            =   2565
                  Picture         =   "FrmControlPersPtoVta.frx":28B0
                  Style           =   1  'Graphical
                  TabIndex        =   47
                  ToolTipText     =   "Seleccione el Nº de Serie de la Boleta de Venta"
                  Top             =   300
                  Width           =   225
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   6
                  Left            =   1605
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   49
                  Text            =   "txt_cb(6)"
                  ToolTipText     =   "Ingrese el Nº de Serie de la Boleta de Venta"
                  Top             =   270
                  Width           =   1215
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   7
                  Left            =   1605
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   50
                  Text            =   "txt_cb(7)"
                  ToolTipText     =   "Ingrese el Código de la Plantilla de Impresión"
                  Top             =   615
                  Width           =   1215
               End
               Begin VB.Label lbl_cb_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb_cod(7)"
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
                  Index           =   7
                  Left            =   6345
                  TabIndex        =   56
                  Top             =   615
                  Visible         =   0   'False
                  Width           =   1290
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(7)"
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
                  Index           =   7
                  Left            =   2850
                  TabIndex        =   55
                  Top             =   615
                  Width           =   4905
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Plantilla de Impresión"
                  Height          =   195
                  Index           =   7
                  Left            =   45
                  TabIndex        =   54
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label lbl_cb_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb_cod(6)"
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
                  Index           =   6
                  Left            =   6345
                  TabIndex        =   53
                  Top             =   315
                  Visible         =   0   'False
                  Width           =   1290
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(6)"
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
                  Index           =   6
                  Left            =   2850
                  TabIndex        =   52
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   4905
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº Serie:"
                  Height          =   195
                  Index           =   6
                  Left            =   45
                  TabIndex        =   51
                  Top             =   360
                  Width           =   630
               End
            End
            Begin VB.Frame fra 
               Caption         =   "[ Boleta de Venta a Cargo ]"
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
               Height          =   1005
               Index           =   3
               Left            =   135
               TabIndex        =   39
               Top             =   3060
               Width           =   8175
               Begin VB.CommandButton cb 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   4
                  Left            =   2565
                  Picture         =   "FrmControlPersPtoVta.frx":29E2
                  Style           =   1  'Graphical
                  TabIndex        =   13
                  ToolTipText     =   "Seleccione el Nº de Serie de la Boleta de Venta"
                  Top             =   300
                  Width           =   225
               End
               Begin VB.CommandButton cb 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   5
                  Left            =   2565
                  Picture         =   "FrmControlPersPtoVta.frx":2B14
                  Style           =   1  'Graphical
                  TabIndex        =   15
                  ToolTipText     =   "Seleccione la Plantilla de Impresión"
                  Top             =   660
                  Width           =   225
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   4
                  Left            =   1605
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   12
                  Text            =   "txt_cb(4)"
                  ToolTipText     =   "Ingrese el Nº de Serie de la Boleta de Venta"
                  Top             =   270
                  Width           =   1215
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   5
                  Left            =   1605
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   14
                  Text            =   "txt_cb(5)"
                  ToolTipText     =   "Ingrese el Código de la Plantilla de Impresión"
                  Top             =   615
                  Width           =   1215
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº Serie:"
                  Height          =   195
                  Index           =   4
                  Left            =   45
                  TabIndex        =   45
                  Top             =   360
                  Width           =   630
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(4)"
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
                  Left            =   2850
                  TabIndex        =   44
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   4905
               End
               Begin VB.Label lbl_cb_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb_cod(4)"
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
                  Index           =   4
                  Left            =   6345
                  TabIndex        =   43
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   1290
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Plantilla de Impresión"
                  Height          =   195
                  Index           =   5
                  Left            =   45
                  TabIndex        =   42
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(5)"
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
                  Index           =   5
                  Left            =   2850
                  TabIndex        =   41
                  Top             =   615
                  Width           =   4905
               End
               Begin VB.Label lbl_cb_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb_cod(5)"
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
                  Index           =   5
                  Left            =   6345
                  TabIndex        =   40
                  Top             =   615
                  Visible         =   0   'False
                  Width           =   1290
               End
            End
            Begin VB.Frame fra 
               Caption         =   "[ Factura a Cargo ]"
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
               Height          =   1005
               Index           =   2
               Left            =   135
               TabIndex        =   32
               Top             =   1950
               Width           =   8175
               Begin VB.CommandButton cb 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   3
                  Left            =   2565
                  Picture         =   "FrmControlPersPtoVta.frx":2C46
                  Style           =   1  'Graphical
                  TabIndex        =   11
                  ToolTipText     =   "Seleccione la Plantilla de Impresión"
                  Top             =   660
                  Width           =   225
               End
               Begin VB.CommandButton cb 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   2
                  Left            =   2565
                  Picture         =   "FrmControlPersPtoVta.frx":2D78
                  Style           =   1  'Graphical
                  TabIndex        =   9
                  ToolTipText     =   "Seleccione el Nº de Serie de la Factura"
                  Top             =   300
                  Width           =   225
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   2
                  Left            =   1605
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   8
                  Text            =   "txt_cb(2)"
                  ToolTipText     =   "Ingrese el Nº de Serie de la Factura"
                  Top             =   270
                  Width           =   1215
               End
               Begin VB.TextBox txt_cb 
                  Height          =   300
                  Index           =   3
                  Left            =   1605
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   10
                  Text            =   "txt_cb(3)"
                  ToolTipText     =   "Ingrese el Código de la Plantilla de Impresión"
                  Top             =   615
                  Width           =   1215
               End
               Begin VB.Label lbl_cb_cod 
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb_cod(3)"
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
                  Index           =   3
                  Left            =   6360
                  TabIndex        =   38
                  Top             =   615
                  Visible         =   0   'False
                  Width           =   1290
               End
               Begin VB.Label lbl_cb 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbl_cb(3)"
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
                  Left            =   2850
                  TabIndex        =   37
                  Top             =   615
                  Width           =   4905
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Plantilla de Impresión"
                  Height          =   195
                  Index           =   3
                  Left            =   45
                  TabIndex        =   36
                  Top             =   720
                  Width           =   1485
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
                  Left            =   6360
                  TabIndex        =   35
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   1290
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
                  Left            =   2850
                  TabIndex        =   34
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   4905
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº Serie:"
                  Height          =   195
                  Index           =   2
                  Left            =   45
                  TabIndex        =   33
                  Top             =   360
                  Width           =   630
               End
            End
            Begin VB.CommandButton cb 
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   2070
               Picture         =   "FrmControlPersPtoVta.frx":2EAA
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Seleccione el Almancén"
               Top             =   600
               Width           =   225
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   1
               Left            =   1110
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   4
               Text            =   "txt(1)"
               ToolTipText     =   "Ingrese el Código del Personal"
               Top             =   915
               Width           =   1215
            End
            Begin VB.CommandButton cb 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   2070
               Picture         =   "FrmControlPersPtoVta.frx":2FDC
               Style           =   1  'Graphical
               TabIndex        =   1
               Top             =   240
               Width           =   225
            End
            Begin VB.Frame fra 
               Caption         =   "[ Funciones ]"
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
               TabIndex        =   24
               Top             =   1305
               Width           =   8175
               Begin VB.CheckBox chk 
                  Caption         =   "&Supervisor"
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
                  Left            =   5625
                  TabIndex        =   7
                  Top             =   255
                  Width           =   1575
               End
               Begin VB.CheckBox chk 
                  Caption         =   "&Vendedor"
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
                  Left            =   1095
                  TabIndex        =   5
                  Top             =   255
                  Width           =   1515
               End
               Begin VB.CheckBox chk 
                  Caption         =   "&Cajero"
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
                  Left            =   3435
                  TabIndex        =   6
                  Top             =   255
                  Width           =   1305
               End
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   1110
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   0
               Text            =   "txt_cb(0)"
               ToolTipText     =   "Ingrese DNI del Programador"
               Top             =   210
               Width           =   1215
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   1
               Left            =   1110
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   2
               Text            =   "txt_cb(1)"
               ToolTipText     =   "Ingrese el Código del Almacen"
               Top             =   570
               Width           =   1215
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Almacén"
               Height          =   195
               Index           =   1
               Left            =   270
               TabIndex        =   31
               Top             =   675
               Width           =   615
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
               Left            =   2325
               TabIndex        =   30
               Top             =   570
               Width           =   5730
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
               Index           =   1
               Left            =   6930
               TabIndex        =   29
               Top             =   555
               Visible         =   0   'False
               Width           =   1290
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Código"
               Height          =   195
               Index           =   1
               Left            =   270
               TabIndex        =   28
               Top             =   1020
               Width           =   495
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
               Left            =   6930
               TabIndex        =   27
               Top             =   195
               Visible         =   0   'False
               Width           =   1290
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
               Left            =   2325
               TabIndex        =   26
               Top             =   210
               Width           =   5730
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Empeado"
               Height          =   195
               Index           =   0
               Left            =   270
               TabIndex        =   25
               Top             =   315
               Width           =   855
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Control de Personal"
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
            TabIndex        =   22
            Top             =   45
            Width           =   9075
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5745
         Left            =   -10395
         TabIndex        =   18
         Top             =   375
         Width           =   9750
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   5400
            Left            =   30
            TabIndex        =   19
            Top             =   345
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   9525
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Código"
            Columns(0).DataField=   "codigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Empleado"
            Columns(1).DataField=   "nomemp"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   4
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Vendedor"
            Columns(2).DataField=   "ven"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   4
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cajero"
            Columns(3).DataField=   "caj"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   20
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Supervisor"
            Columns(4).DataField=   "sup"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Almacén"
            Columns(5).DataField=   "almdesc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1296"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1217"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=7170"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7091"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1588"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1508"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1111"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1032"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1746"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1667"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=3201"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=3122"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Control de Personal "
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
            TabIndex        =   20
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
Attribute VB_Name = "FrmControlPersPtoVta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim RstFrm As New ADODB.Recordset
Dim Mostrando As Boolean
Dim SeEjecuto As Boolean
'------------------------------------------------
Dim mIdCodigoTmp As Integer     '--UTIL PARA ALMACENAR LOS VALORES ANTES DE SELECCIONAR UN REGISTRO
                                '--SE USA PARA LIMPIAR LOS CAMPOS QUE DEPENDEN DEL CAMPO SELECCIONADO
                                '--CUANDO EL VALOR YA SELECCIONADO SEA DIFERENTE AL NUEVO VALOR SELECCIONADO

Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(1, 4) As String
    
    xCampos(0, 0) = "Empleado":     xCampos(0, 1) = "nomemp":     xCampos(0, 2) = "4000":    xCampos(0, 3) = "C"
        
    nSQL = "SELECT pvt_emp.id, pvt_emp.idemp, pla_empleados.numdoc, pla_empleados.numdoc, [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nomemp, pvt_emp.ven, pvt_emp.caj, pvt_emp.sup, pvt_emp.codigo, pvt_emp.idalm, alm_almacenes.descripcion AS almdesc, pvt_emp.idalm AS almcod " _
        + vbCr + " FROM (pla_empleados INNER JOIN pvt_emp ON pla_empleados.id = pvt_emp.idemp) LEFT JOIN alm_almacenes ON pvt_emp.idalm = alm_almacenes.id; "

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Personal", "nomemp", "nomemp", Principio
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
    '--DEL EMPLEADO
    txt_cb(0).Text = RstFrm.Fields("numdoc") & ""
    lbl_cb_cod(0).Caption = RstFrm("idemp") & ""
    lbl_cb(0).Caption = RstFrm("nomemp") & ""
    '--DEL ALMACEN
    txt_cb(1).Text = RstFrm.Fields("idalm") & ""
    lbl_cb_cod(1).Caption = RstFrm("almcod") & ""
    lbl_cb(1).Caption = RstFrm("almdesc") & ""
    '------
    chk(0).Value = Abs(Val(NulosN(RstFrm("ven"))))
    chk(1).Value = Abs(Val(NulosN(RstFrm("caj"))))
    chk(2).Value = Abs(Val(NulosN(RstFrm("sup"))))
    txt(1).Text = RstFrm("codigo") & ""
    '---------
    '--MOSTRAR DEMAS DATOS
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    nSQL = "SELECT pvt_emp.id, alm_numseries.id AS serfacid, alm_numseries.numser AS serfacdesc, var_plantilladoc.id AS impfacid, var_plantilladoc.descripcion AS impfacdesc, alm_numseries_1.id AS serbolid, alm_numseries_1.numser AS serboldesc, var_plantilladoc_1.id AS impbolid, var_plantilladoc_1.descripcion AS impboldesc, alm_numseries_2.id AS sertickid, alm_numseries_2.numser AS sertickdesc, var_plantilladoc_2.id AS imptickid, var_plantilladoc_2.descripcion AS imptickdesc " _
        + vbCr + " FROM (((((pvt_emp LEFT JOIN alm_numseries ON pvt_emp.idalmser = alm_numseries.id) LEFT JOIN alm_numseries AS alm_numseries_1 ON pvt_emp.idalmser1 = alm_numseries_1.id) LEFT JOIN var_plantilladoc ON pvt_emp.idplan = var_plantilladoc.id) LEFT JOIN var_plantilladoc AS var_plantilladoc_1 ON pvt_emp.idplan1 = var_plantilladoc_1.id) LEFT JOIN alm_numseries AS alm_numseries_2 ON pvt_emp.idalmser2 = alm_numseries_2.id) LEFT JOIN var_plantilladoc AS var_plantilladoc_2 ON pvt_emp.idplan2 = var_plantilladoc_2.id " _
        + vbCr + " WHERE (((pvt_emp.id)=" + CStr(RstFrm.Fields("id")) + "));"
    
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        '--DEL SERIE DE FATURA
        txt_cb(2).Text = RstTmp.Fields("serfacdesc") & ""
        lbl_cb_cod(2).Caption = RstTmp.Fields("serfacid") & ""
        '--DEL FORMATO DE IMPRESION DE FACTURA
        txt_cb(3).Text = RstTmp.Fields("impfacid") & ""
        lbl_cb_cod(3).Caption = RstTmp.Fields("impfacid") & ""
        lbl_cb(3).Caption = RstTmp.Fields("impfacdesc") & ""
        '--------------------------------------------------------------------------------
        '--DE SERIE DE BOLETA
        txt_cb(4).Text = RstTmp.Fields("serboldesc") & ""
        lbl_cb_cod(4).Caption = RstTmp.Fields("serbolid") & ""
        '--DEL FORMATO DE IMPRESION DE LA BOLETA
        txt_cb(5).Text = RstTmp.Fields("impbolid") & ""
        lbl_cb_cod(5).Caption = RstTmp.Fields("impbolid") & ""
        lbl_cb(5).Caption = RstTmp.Fields("impboldesc") & ""
        '--------------------------------------------------------------------------------
        '--DE SERIE DE BOLETA
        txt_cb(6).Text = RstTmp.Fields("sertickdesc") & ""
        lbl_cb_cod(6).Caption = RstTmp.Fields("sertickid") & ""
        '--DEL FORMATO DE IMPRESION DE LA BOLETA
        txt_cb(7).Text = RstTmp.Fields("imptickid") & ""
        lbl_cb_cod(7).Caption = RstTmp.Fields("imptickid") & ""
        lbl_cb(7).Caption = RstTmp.Fields("imptickdesc") & ""
    End If
    Set RstTmp = Nothing
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
    habilitar_Locked txt, Not band
    habilitar Me.cb, band
    habilitar fra, band
    
    TabOne1.CurrTab = IIf(band = False, 0, 1)
    TabOne1.TabEnabled(0) = Not band
    
End Sub

Sub Blanquea()

    LimpiaText txt_cb
    LimpiaText txt
    chk(0).Value = 0:           chk(1).Value = 0
End Sub

Sub Cancelar()
    QueHace = 3
    Habilitar_Obj False
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
    Label5.Caption = "Detalle del Control de Personal"
End Sub

Function Grabar() As Boolean
    Dim xId, A As Integer
    
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modficar") + " el Registro", vbQuestion + vbYesNo) = vbNo Then Exit Function

    
    Dim RstCab As New ADODB.Recordset
            
    On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then 'NUEVO
        xId = HallaCodigoTabla("pvt_emp", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pvt_emp", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else 'MODIFICAR
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pvt_emp WHERE id = " & xId & "", xCon
    End If
    RstCab("idemp") = NulosN(lbl_cb_cod(0).Caption)
    RstCab("ven") = chk(0).Value
    RstCab("caj") = chk(1).Value
    RstCab("sup") = chk(2).Value
    RstCab("codigo") = Trim(txt(1).Text)
    '--DEL ALMACEN
    RstCab("idalm") = NulosN(lbl_cb_cod(1).Caption)
    '--DE LA FACTURA
    RstCab("idalmser") = NulosN(lbl_cb_cod(2).Caption)
    RstCab("idplan") = NulosN(lbl_cb_cod(3).Caption)
    '--DE LA BOLETA
    RstCab("idalmser1") = NulosN(lbl_cb_cod(4).Caption)
    RstCab("idplan1") = NulosN(lbl_cb_cod(5).Caption)
    '--DEL TICKET
    RstCab("idalmser2") = NulosN(lbl_cb_cod(6).Caption)
    RstCab("idplan2") = NulosN(lbl_cb_cod(7).Caption)
    RstCab.Update
    
    xCon.CommitTrans
    Grabar = True
    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    Set RstCab = Nothing
    Label5.Caption = "Detalle del Control de Personal"
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
        xCon.Execute "DELETE * FROM pvt_emp WHERE id = " & Val(RstFrm("id")) & ""
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
'    TabOne1.CurrTab = 1
'    ActivaTool
'    TabOne1.TabEnabled(0) = False
'    Label5.Caption = "Modificando Control de Personal"
'    QueHace = 2
'    Habilitar_Obj True
'    Blanquea
'    MuestraSegundoTab
'    txt_cb(0).SetFocus
    
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
    Label5.Caption = "Agregando Control de Personal"
    Habilitar_Obj True
    Blanquea
    txt_cb(0).SetFocus
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Form_Activate()
    Dim nSQL As String
    If SeEjecuto = False Then
        SeEjecuto = True
        nSQL = "SELECT pvt_emp.id, pvt_emp.idemp, pla_empleados.numdoc, pla_empleados.numdoc, [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nomemp, pvt_emp.ven, pvt_emp.caj, pvt_emp.sup, pvt_emp.codigo, pvt_emp.idalm, alm_almacenes.descripcion AS almdesc, pvt_emp.idalm AS almcod " _
            + vbCr + " FROM (pla_empleados INNER JOIN pvt_emp ON pla_empleados.id = pvt_emp.idemp) LEFT JOIN alm_almacenes ON pvt_emp.idalm = alm_almacenes.id; "
           
        RST_Busq RstFrm, nSQL, xCon
        Set Dg3.DataSource = RstFrm
        If RstFrm.RecordCount = 0 Then
            Dim Rpta As Integer
            Rpta = MsgBox("El registro esta vacio, ¿Desea agregar la función del empleado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstFrm = Nothing
                Unload Me
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
    Mostrando = False
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
    
    If Button.Index = 12 Then TDB_IMPRIMIR Dg3, "IMPRESIÓN", "LISTADO DE PERSONAL"
        
    If Button.Index = 14 Then
        Unload Me
        Set RstFrm = Nothing
    End If
End Sub

Private Function fValidarDatos() As Boolean
    If Trim(lbl_cb_cod(0).Caption) = "" Then
        MsgBox "Falta especificar el empleado.", vbInformation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    '----
    If Trim(txt(1).Text) = "" Then
        MsgBox "Falta ingresar el Código del personal", vbExclamation, xTitulo
        txt(1).SetFocus
        Exit Function
    End If
    '--------------------------------
    '--VALIDAR QUE EL REGISTRO NO ESTE REGISTRADO
    Dim RstTmp As New ADODB.Recordset
    If QueHace = 1 Then
        RST_Busq RstTmp, "SELECT [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nomemp FROM pla_empleados INNER JOIN pvt_emp ON pla_empleados.id = pvt_emp.idemp WHERE ucase(pvt_emp.codigo)='" + UCase(Trim(txt(1).Text)) + "';", xCon
    Else
        RST_Busq RstTmp, "SELECT [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nomemp FROM pla_empleados INNER JOIN pvt_emp ON pla_empleados.id = pvt_emp.idemp WHERE ucase(pvt_emp.codigo)='" + UCase(Trim(txt(1).Text)) + "' AND pvt_emp.id <> " + CStr(RstFrm.Fields("id")) + ";", xCon
    End If
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        MsgBox "El código ya fue asignado a : " + RstTmp.Fields("nomemp") & "" + vbCr + "Cambie de código, luego continue", vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    Set RstTmp = Nothing
    '--------------------------------
    
    
    fValidarDatos = True
End Function


'-----------------------------
'-----------------------------

Private Sub cb_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nTitulo As String
    Dim mIdDoc As Integer '--INDICA EL DOCUMENTO (FACTURA O BOLETA)
    If Index = 2 Or Index = 3 Then mIdDoc = 1 '--FACTURA
    If Index = 4 Or Index = 5 Then mIdDoc = 3 '--BOLETA DE VENTA
    If Index = 6 Or Index = 7 Then mIdDoc = 12 '--TICKET DE MAQUINA REGISTRADORA
    
    Select Case Index
        Case 0 '--EMPLEADO
            ReDim xCampos(2, 4) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "DNI":      xCampos(1, 1) = "numdoc":   xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
                
            nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id " _
                + vbCr + " FROM mae_usuarios INNER JOIN pla_empleados ON mae_usuarios.idemp = pla_empleados.id " _
                + vbCr + " WHERE (((pla_empleados.id) Not In (select idemp from pvt_emp))) " _
                + vbCr + " ORDER BY [pla_empleados].[ape] & ' ' & [pla_empleados].[nom];"
                
             nTitulo = "Buscando Personal"
             
        Case 1 '--ALMACEN
            ReDim xCampos(1, 3) As String
            xCampos(0, 0) = "Almacén":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            
            nSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion AS nombre, alm_almacenes.id AS cod " _
            + vbCr + " FROM alm_almacenes ORDER BY alm_almacenes.descripcion ;"
            
            nTitulo = "Buscando Almacén"
        Case 2, 4, 6 '--DE LOS NUMEROS DE SERIE
        
            If NulosN(lbl_cb_cod(1).Caption) = 0 Then
                MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                txt_cb(1).SetFocus
                Exit Sub
            End If
            ReDim xCampos(1, 3) As String
            xCampos(0, 0) = "Número":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            
            nSQL = "SELECT Format([alm_numseries].[numser],'0000') AS nombre,alm_numseries.id ,alm_numseries.id AS cod " _
                + vbCr + " FROM alm_numseries " _
                + vbCr + " WHERE alm_numseries.idtipdoc=" + CStr(mIdDoc) + " AND alm_numseries.idalm=" + CStr(NulosN(lbl_cb_cod(1).Caption)) + ";"
            
            nTitulo = "Buscando Series"
        
        
        Case 3, 5, 7 '--DE LOS FORMATOS DE IMPRESION
            If NulosN(lbl_cb_cod(1).Caption) = 0 Then
                MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                cb(1).SetFocus
                Exit Sub
            End If
            If NulosN(lbl_cb_cod(Index - 1).Caption) = 0 Then
                MsgBox "Seleccione La Serie donde se procederá a configurar la plantilla de impresión", vbExclamation, xTitulo
                cb(Index - 1).SetFocus
                Exit Sub
            End If
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Nombre":           xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4000":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Tipo de Letra":    xCampos(1, 1) = "tipoletra": xCampos(1, 2) = "1800":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Tamaño":           xCampos(2, 1) = "zise":      xCampos(2, 2) = "1000":   xCampos(2, 3) = "N"
            
            nSQL = "SELECT var_plantilladoc.id, var_plantilladoc.descripcion as nombre ,var_plantilladoc.id as cod, var_plantilladoc.tipoletra, var_plantilladoc.tamañoletra AS zise " _
                + vbCr + " FROM var_plantilladoc " _
                + vbCr + " WHERE (((var_plantilladoc.tipdoc) = " + CStr(mIdDoc) + ")) " _
                + vbCr + " ORDER BY var_plantilladoc.descripcion;"
            
            nTitulo = "Buscando Plantillas de Impresión"
        
    End Select
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    mIdCodigoTmp = NulosN(txt_cb(Index).Text)
    txt_cb(Index).Text = xRs.Fields(0) & ""  '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    
    If mIdCodigoTmp <> 0 And mIdCodigoTmp <> NulosN(txt_cb(Index).Text) And Index = 1 Then
        txt_cb(2).Text = ""
        txt_cb(4).Text = ""
        txt_cb(6).Text = ""
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
        If Index = 1 Then
            txt_cb(2).Text = ""
            txt_cb(4).Text = ""
            txt_cb(6).Text = ""
        End If
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
    If Index = 2 Or Index = 3 Then mIdDoc = 1 '--FACTURA
    If Index = 4 Or Index = 5 Then mIdDoc = 3 '--BOLETA DE VENTA
    If Index = 6 Or Index = 7 Then mIdDoc = 12 '--TICKET DE MAQUINA REGISTRADORA
    Select Case Index
        Case 0 '--PERSONAL
            nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[ape] & ' ' & [pla_empleados].[nom] AS nomemp, pla_empleados.id " _
                + vbCr + " FROM mae_usuarios INNER JOIN pla_empleados ON mae_usuarios.idemp = pla_empleados.id " _
                + vbCr + " WHERE (((pla_empleados.id) Not In (select idemp from pvt_emp))) AND pla_empleados.numdoc ='" + Trim(txt_cb(Index).Text) + "'" _
                + vbCr + " ORDER BY [pla_empleados].[ape] & ' ' & [pla_empleados].[nom];"
        Case 1 '--ALMACEN
            nSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion AS nombre, alm_almacenes.id AS cod " _
                + vbCr + " FROM alm_almacenes WHERE alm_almacenes.id = " + Trim(txt_cb(Index).Text) + " ;"
        Case 2, 4, 6 '--NUMERO DE SERIE
            If NulosN(lbl_cb_cod(1).Caption) = 0 Then
                MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                txt_cb(1).SetFocus
                Exit Sub
            End If

            nSQL = "SELECT Format([alm_numseries].[numser],'0000') AS nombre, alm_numseries.id, alm_numseries.id AS cod  " _
            + vbCr + " FROM alm_numseries " _
            + vbCr + " WHERE Format(alm_numseries.numser,'0000') ='" + Format(NulosN(txt_cb(Index).Text), "0000") + "' AND alm_numseries.idtipdoc=" + CStr(mIdDoc) + " AND alm_numseries.idalm=" + CStr(NulosN(lbl_cb_cod(1).Caption)) + " ;"
        
        Case 3, 5, 7 'PLANTILLA DE IMPRESION
            If NulosN(lbl_cb_cod(1).Caption) = 0 Then
                MsgBox "Seleccione el Almacén donde se hará la Operación", vbExclamation, xTitulo
                txt_cb(1).SetFocus
                Exit Sub
            End If
            If NulosN(lbl_cb_cod(Index - 1).Caption) = 0 Then
                MsgBox "Seleccione La Serie donde se procederá a configurar la plantilla de impresión", vbExclamation, xTitulo
                txt_cb(Index - 1).SetFocus
                Exit Sub
            End If
            
            nSQL = "SELECT var_plantilladoc.id, var_plantilladoc.descripcion as nombre,var_plantilladoc.id as cod, var_plantilladoc.tipoletra, var_plantilladoc.tamañoletra AS zise " _
                + vbCr + " FROM var_plantilladoc " _
                + vbCr + " WHERE (((var_plantilladoc.tipdoc) = " + CStr(mIdDoc) + ")) " _
                + vbCr + " ORDER BY var_plantilladoc.descripcion;"

            
    End Select
    If xCon.State = 0 Then Exit Sub
    RST_Busq RST_TMP, nSQL, xCon
    
    If RST_TMP.State = 0 Then Exit Sub
    If RST_TMP.RecordCount > 0 Then
        txt_cb(Index) = RST_TMP.Fields(0) & "" '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RST_TMP.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = RST_TMP.Fields(2) & "" '--CODIGO
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
'    txt_cb_KeyDown Index, 13, 0
End Sub
