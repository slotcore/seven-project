VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#12.0#0"; "Codejock.Calendar.v12.0.0.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCronoPedidos2 
   Caption         =   "Ventas - Cronograma de Entregas"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H80000013&
      Caption         =   "[ Detalle ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3255
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "Clic para arrastrar"
      Top             =   2040
      Visible         =   0   'False
      Width           =   8955
      Begin VB.TextBox TxtTipPed 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   6210
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   14
         Text            =   "TxtTipPed"
         Top             =   2415
         Width           =   705
      End
      Begin VB.TextBox TxtConPag 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   13
         Text            =   "TxtConPag"
         Top             =   2415
         Width           =   915
      End
      Begin VB.TextBox TxtPtoVta 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "TxtPtoVta"
         Top             =   1425
         Width           =   915
      End
      Begin VB.TextBox TxtNumRuc 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   10
         Text            =   "TxtNumRuc"
         Top             =   1080
         Width           =   1770
      End
      Begin VB.TextBox TxtTipDoc 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "TxtTipDoc"
         Top             =   1755
         Width           =   915
      End
      Begin VB.TextBox TxtNumSer 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "TxtNumSer"
         Top             =   2085
         Width           =   915
      End
      Begin VB.TextBox TxtNumDoc 
         Height          =   300
         Left            =   2730
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "TxtNumDoc"
         Top             =   2085
         Width           =   1440
      End
      Begin VB.TextBox txtglosa 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         Text            =   "TxtGlosa"
         Top             =   2745
         Width           =   7200
      End
      Begin VB.TextBox TxtOC 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   6210
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "TxtOC"
         Top             =   2085
         Width           =   2595
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
         Height          =   300
         Left            =   1620
         TabIndex        =   11
         Top             =   735
         Width           =   1260
         _ExtentX        =   2223
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
         Valor           =   "03/01/2004"
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   8910
         X2              =   30
         Y1              =   3210
         Y2              =   3210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   8910
         X2              =   8910
         Y1              =   90
         Y2              =   3210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "[ X ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   8475
         TabIndex        =   33
         ToolTipText     =   "Cerrar Detalle"
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Pedido"
         Height          =   195
         Index           =   5
         Left            =   5325
         TabIndex        =   32
         Top             =   2490
         Width           =   855
      End
      Begin VB.Label LblTipPed 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTipPed"
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
         Left            =   6915
         TabIndex        =   31
         Top             =   2415
         Width           =   1890
      End
      Begin VB.Label LblIdCliente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LblIdCliente"
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   2925
         TabIndex        =   30
         Top             =   795
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         Height          =   50
         Left            =   2580
         Top             =   2200
         Width           =   105
      End
      Begin VB.Label LblNomDoc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblNomDoc"
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
         Left            =   2565
         TabIndex        =   29
         Top             =   1755
         Width           =   3615
      End
      Begin VB.Label LblNomCli 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblNomCli"
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
         Left            =   3405
         TabIndex        =   28
         Top             =   1080
         Width           =   5430
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de Pedido"
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
         Left            =   210
         TabIndex        =   27
         Top             =   300
         Width           =   8505
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Emisión"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   26
         Top             =   810
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   25
         Top             =   1155
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   24
         Top             =   1830
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Documento"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   23
         Top             =   2175
         Width           =   1275
      End
      Begin VB.Label lblglosa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Glosa"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   2865
         Width           =   405
      End
      Begin VB.Label LblCantidad 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCantidad"
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
         Left            =   6255
         TabIndex        =   21
         Top             =   735
         Width           =   2565
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         Height          =   195
         Index           =   11
         Left            =   5350
         TabIndex        =   20
         Top             =   810
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Punto Venta"
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   19
         Top             =   1500
         Width           =   885
      End
      Begin VB.Label LblPtoVta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblPtoVta"
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
         Left            =   2565
         TabIndex        =   18
         Top             =   1425
         Width           =   6255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condición de Pago"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   17
         Top             =   2490
         Width           =   1350
      End
      Begin VB.Label LblCondPag 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblCondPag"
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
         Left            =   2565
         TabIndex        =   16
         Top             =   2415
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Compra"
         Height          =   195
         Left            =   4935
         TabIndex        =   15
         Top             =   2160
         Width           =   1245
      End
   End
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":0544
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":06C8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":0B1C
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":0C34
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":1178
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":16BC
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":17D0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":18E4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":1D38
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoPedidos2.frx":1EA4
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12585
      _ExtentX        =   22199
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Plan de Ventas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar Plan de Ventas"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Plan de Ventas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Desactivar Plan de Ventas"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
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
   Begin SizerOneLibCtl.ElasticOne EO1 
      Height          =   7365
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   345
      Width           =   12630
      _cx             =   22278
      _cy             =   12991
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   3
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   -1  'True
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin XtremeCalendarControl.CalendarControl CalendarControl1 
         Height          =   7185
         Left            =   30
         TabIndex        =   2
         Top             =   60
         Width           =   9915
         _Version        =   786432
         _ExtentX        =   17489
         _ExtentY        =   12674
         _StockProps     =   64
         ViewType        =   2
         ShowCaptionBar  =   -1  'True
         ShowSwitchViewButtons=   0   'False
      End
      Begin XtremeCalendarControl.DatePicker DatePicker1 
         Height          =   7215
         Left            =   9960
         TabIndex        =   3
         Top             =   30
         Width           =   2580
         _Version        =   786432
         _ExtentX        =   4551
         _ExtentY        =   12726
         _StockProps     =   64
         ShowNoneButton  =   0   'False
         ShowWeekNumbers =   -1  'True
         RowCount        =   3
         TextNoneButton  =   "Ninguno"
         TextTodayButton =   "Hoy"
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "menus"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar               "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmCronoPedidos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim RstFuente As New ADODB.Recordset
Dim RstPlan As New ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim xConTMP As ADODB.Connection

Dim OrigFX As Long
Dim OrigFY As Long

Private Sub CalendarControl1_BeforeEditOperation(ByVal OpParams As XtremeCalendarControl.CalendarEditOperationParameters, bCancelOperation As Boolean)
    bCancelOperation = True
End Sub

Private Sub CalendarControl1_DblClick()
    Dim Rst As New ADODB.Recordset
    Dim cSQL As String
On Error Resume Next
    Dim HitTest As CalendarHitTestInfo
    Set HitTest = CalendarControl1.ActiveView.HitTest
    Dim m_pEditingEvent As CalendarEvent
    Set m_pEditingEvent = HitTest.ViewEvent.Event
    
    If m_pEditingEvent.Body = "" Then Exit Sub
    
    cSQL = "SELECT ped_pedido.id, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, ped_pedido.fchent, ped_pedido.anulado, ped_pedido.numreg, ped_pedido.idlib, ped_pedido.oc, ped_pedido.proceso, IIf(ped_pedido.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc) AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev AS conpagabre, ped_pedido.fchemi & '' AS fchemi1, vta_puntoVenta.descripcion AS ptovta, ped_tipo.descripcion AS tipped, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.canpro, ped_pedidodet.iditem " _
    + vbCr + "FROM (((ped_tipo RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN (vta_puntoVenta RIGHT JOIN (ped_pedido LEFT JOIN mae_condpago ON ped_pedido.idconpag = mae_condpago.id) ON vta_puntoVenta.id = ped_pedido.idpunvecli) ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) ON ped_tipo.id = ped_pedido.idtipped) LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
    + vbCr + "WHERE (((ped_pedido.id)=" & m_pEditingEvent.Body & ") AND ((ped_pedidodet.iditem)=" & m_pEditingEvent.ScheduleID & ") AND ((ped_pedido.fchent)=CDate('" & Mid(m_pEditingEvent.StartTime, 1, 10) & "')) AND ((ped_pedido.idtipped)=1) AND ((ped_pedido.anulado)=0)); " _
    + vbCr + "Union " _
    + vbCr + "SELECT ped_pedido.id, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, ped_pedidodetent.fchent, ped_pedido.anulado, ped_pedido.numreg, ped_pedido.idlib, ped_pedido.oc, ped_pedido.proceso, IIf(ped_pedido.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc) AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev AS conpagabre, ped_pedido.fchemi & '' AS fchemi1, vta_puntoVenta.descripcion AS ptovta, ped_tipo.descripcion AS tipped, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodetent.canpro, ped_pedidodetent.iditem " _
    + vbCr + "FROM (((ped_tipo RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN (vta_puntoVenta RIGHT JOIN (ped_pedido LEFT JOIN mae_condpago ON ped_pedido.idconpag = mae_condpago.id) ON vta_puntoVenta.id = ped_pedido.idpunvecli) ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) ON ped_tipo.id = ped_pedido.idtipped) LEFT JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodetent.idunimed = mae_unidades.id " _
    + vbCr + "WHERE (((ped_pedido.id)=" & m_pEditingEvent.Body & ") AND ((ped_pedidodetent.iditem)=" & m_pEditingEvent.ScheduleID & ") AND ((ped_pedidodetent.fchent)=CDate('" & Mid(m_pEditingEvent.StartTime, 1, 10) & "')) AND ((ped_pedido.idtipped)=2) AND ((ped_pedido.anulado)=0));"
    
    RST_Busq Rst, cSQL, xCon
    
    If Rst.RecordCount = 0 Then Exit Sub
    If Rst.EOF = True Or Rst.BOF = True Then Exit Sub
    
    If IsDate(Rst("fchemi")) = True Then TxtFchDoc.Valor = CDate(Rst("fchemi"))
    LblCantidad.Caption = NulosN(Rst("canpro")) & " " & NulosC(Rst("mae_unidades.abrev"))
    TxtTipDoc.Text = NulosN(Rst("tipdoc"))
    LblNomDoc.Caption = NulosC(Rst("nomdoc"))
    TxtNumRuc.Text = NulosC(Rst("numruc"))
    LblNomCli.Caption = NulosC(Rst("nombre"))
    LblIdCliente.Caption = Rst("idcli")
    TxtPtoVta.Text = NulosC(Rst("idpunvecli"))
    LblPtoVta.Caption = NulosC(Rst("ptovta"))
    TxtNumSer.Text = NulosC(Rst("numser"))
    TxtNumDoc.Text = NulosC(Rst("numdoc"))
    TxtConPag.Text = NulosC(Rst("idconpag"))
    LblCondPag.Caption = NulosC(Rst("desccond"))
    TxtTipPed.Text = NulosN(Rst("idtipped"))
    LblTipPed.Caption = NulosC(Rst("tipped"))
    TxtOC.Text = NulosC(Rst("oc"))
    txtglosa.Text = NulosC(Rst("glosa"))
    
    Set m_pEditingEvent = Nothing
    
    CalendarControl1.DataProvider.AddEvent m_pEditingEvent
    Err.Clear
    FrmDetalle.Visible = True
End Sub

Private Sub Form_Resize()
    EO1.Width = Me.Width - 125
    EO1.Height = Me.Height - 835
    FrmDetalle.Top = Me.Height - 4000
    FrmDetalle.Left = Me.Width - 10000
End Sub

Private Sub DatePicker1_SelectionChanged()
    DatePicker1.AttachToCalendar CalendarControl1
    DatePicker1.Select CalendarControl1.ActiveView.Selection.End
    If CalendarControl1.ViewType = xtpCalendarDayView Then CalendarControl1.ViewType = xtpCalendarWeekView
End Sub

Sub LlenarDatos()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim cSQL As String
    
    cSQL = "SELECT DISTINCT ped_pedido.id As idPed, alm_inventario.id As idProd, alm_inventario.descripcion, mae_cliente.nombre, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedidodetent.estado, ped_pedido.idtipped " _
    + vbCr + "FROM mae_cliente RIGHT JOIN (alm_inventario RIGHT JOIN (ped_pedido LEFT JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) ON alm_inventario.id = ped_pedidodetent.iditem) ON mae_cliente.id = ped_pedido.idcli " _
    + vbCr + "WHERE (((alm_inventario.descripcion) Is Not Null) AND ((ped_pedido.idtipped)=2)); " _
    + vbCr + "Union " _
    + vbCr + "SELECT DISTINCT ped_pedido.id As idPed, alm_inventario.id As idProd, alm_inventario.descripcion, mae_cliente.nombre, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedidodetent.estado, ped_pedido.idtipped " _
    + vbCr + "FROM (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) LEFT JOIN (alm_inventario RIGHT JOIN ped_pedidodetent ON alm_inventario.id = ped_pedidodetent.iditem) ON ped_pedido.id = ped_pedidodetent.idped " _
    + vbCr + "WHERE (((alm_inventario.descripcion) Is Not Null) AND ((ped_pedido.idtipped)=1));"

    RST_Busq Rst, cSQL, xCon
    
    Set xConTMP = Nothing
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Dim xHorIni, HorFin As String
        xHorIni = "08:00:00"
        Dim NewEvent As CalendarEvent, Recurrence As CalendarRecurrencePattern
        Set NewEvent = CalendarControl1.DataProvider.CreateEvent
        
        For A = 1 To Rst.RecordCount
            NewEvent.Body = NulosC(Rst("idPed"))
            NewEvent.ScheduleID = NulosN(Rst("idProd"))
            NewEvent.Subject = NulosC(Rst("descripcion"))
            NewEvent.Location = NulosC(Rst("nombre"))
            NewEvent.StartTime = Format(Rst("fchent"), "dd/mm/yyyy") & " " & Format(xHorIni, "hh:mm:ss")
            NewEvent.EndTime = Format(Rst("fchent"), "dd/mm/yyyy") & " " & Format(xHorIni, "hh:mm:ss")
            NewEvent.Importance = xtpCalendarImportanceHigh
            NewEvent.AllDayEvent = True
            
            Rst.MoveNext
            CalendarControl1.DataProvider.AddEvent NewEvent
        Next A
    End If
End Sub

Private Sub Label2_Click()
    FrmDetalle.Visible = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 14 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LlenarDatos
    DatePicker1_SelectionChanged
End Sub

Private Sub FrmDetalle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    FrmDetalle.Drag
End Sub

'Metodos para arrastrar el Frame
''''''''''''''''''''''''''''''''
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X - OrigFX, Y - OrigFY
End Sub

Private Sub Toolbar1_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X - OrigFX, Y - OrigFY
End Sub

Private Sub CalendarControl1_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X - OrigFX, Y - OrigFY
End Sub

Private Sub DatePicker1_DragDrop(Source As Control, X As Single, Y As Single)
    DatePicker1.RedrawControl
End Sub
