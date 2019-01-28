VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmManControl 
   Caption         =   "Control de Registros"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
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
      Height          =   2180
      Left            =   2100
      TabIndex        =   7
      Top             =   4830
      Visible         =   0   'False
      Width           =   5610
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   5310
         Picture         =   "FrmManControl.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   8
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid Fgr 
         Height          =   1605
         Index           =   3
         Left            =   90
         TabIndex        =   9
         Top             =   390
         Width           =   5400
         _cx             =   9525
         _cy             =   2831
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
         Rows            =   5
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManControl.frx":02EC
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
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de Registro"
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
         Left            =   45
         TabIndex        =   10
         Top             =   75
         Width           =   1650
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   1
         X1              =   5580
         X2              =   5580
         Y1              =   0
         Y2              =   2160
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   5580
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   0
         Top             =   45
         Width           =   5540
      End
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   3750
      TabIndex        =   3
      Top             =   3660
      Visible         =   0   'False
      Width           =   4740
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ESPERE POR FAVOR ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   1470
         TabIndex        =   6
         Top             =   480
         Width           =   1770
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
         Left            =   435
         TabIndex        =   5
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label LblProg 
         AutoSize        =   -1  'True
         Caption         =   "CONTROL DE REGISTROS"
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
         Left            =   1920
         TabIndex        =   4
         Top             =   180
         Width           =   2025
      End
      Begin VB.Shape Shape1 
         Height          =   765
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   4605
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg 
      Height          =   6975
      Left            =   30
      TabIndex        =   0
      Top             =   630
      Width           =   12075
      _cx             =   21299
      _cy             =   12303
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   21
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmManControl.frx":0372
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11460
      Top             =   240
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
            Picture         =   "FrmManControl.frx":05B0
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":0AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":0E86
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":100A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":145E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":1576
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":1ABA
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":1FFE
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":2112
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":2226
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":267A
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":27E6
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControl.frx":2D2E
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Materiales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Linea"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Control de Registros"
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
      Height          =   375
      Left            =   -30
      TabIndex        =   1
      Top             =   390
      Width           =   12180
   End
End
Attribute VB_Name = "FrmManControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANCONTROL.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE VISUALIZAR EL CONTROL DE REGISTROS DEL SISTEMA
'* DISEÑADO POR     : JOSE CHACON MANRIQUE
'* ULTIMA REVISION  : 08/09/2011
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer                         ' INDICA EN QUE MODO SE ENCUENTRA EL FORMULARUI
Dim SeEjecuto As Boolean                       ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim Agregando As Boolean
Dim mIdRegistro&                               ' identificador del registro
Dim IdMenuActivo As Integer                    'INDICA EL CODIGO DEL MENU ACTIVO
Dim cSQL As String
Dim RstControl As New ADODB.Recordset

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        'OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '--ocultar el boton a agregar
        'Toolbar1.Buttons(1).Visible = False
    End If
End Sub

Private Sub pCargarGrid()
    Dim A As Integer
    Dim B As Integer
    Dim CHONORARIOS_ As String
    Dim CBOLETA_ As String
    Dim CCOMPRAS_ As String
    Dim CVENTAS_ As String
    Dim CPRODUCCION_ As String
    Dim CALMACEN_ As String
    Dim CINGRESOS_ As String
    Dim CEGRESOS_ As String
    Dim CRETENCION_ As String
    Dim CPERCEPCION_ As String
    Dim CCANJE_ As String
    Dim CGUIA_ As String
    Dim CTAREAS_ As String
    Dim CCONCILIACION_ As String
    Dim CPROVISIONES_ As String
    Dim CPEDIDOS_ As String
    Dim xAnoInicio As String
    
    xAnoInicio = "CDate('01/01/" & AnoTra & "')"
    
    
    
    FraProgreso.Visible = True
    FraProgreso.Refresh
    
    
    CBOLETA_ = "SELECT 14 AS Orden, 'Boleta Pago' AS Area, Date() AS FechaActual, Max(pla_boleta.fchdoc) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(pla_boleta.fchreg) AS Mes, 105 AS idTabla " _
        + vbCr + "From pla_boleta " _
        + vbCr + "WHERE (((pla_boleta.numreg)<>'000001')) " _
        + vbCr + "GROUP BY 3, 'Boleta Pago', Date(), Month(pla_boleta.fchreg), 105,pla_boleta.ano " _
        + vbCr + "HAVING (((Max(pla_boleta.fchreg))>= " & xAnoInicio & ")) AND pla_boleta.ano = " & AnoTra
        
    CCOMPRAS_ = "SELECT 2 AS Orden, 'Compras' AS Area, Date() AS FechaActual, Max(com_compras.fchdoc) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month([com_compras].[fchreg]) AS Mes, 218 AS idTabla " _
        + vbCr + "From com_compras " _
        + vbCr + "WHERE (((com_compras.numreg)<>'000001')) " _
        + vbCr + "GROUP BY 2, 'Compras', Date(), Month(com_compras.fchreg), 218 " _
        + vbCr + "HAVING (((Max(com_compras.fchreg))>= " & xAnoInicio & "))"
    
    CHONORARIOS_ = "SELECT 3 AS Orden, 'Honorarios' AS Area, Date() AS FechaActual, Max(com_honorarios.fchdoc) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(com_honorarios.fchreg) AS Mes, 219 AS idTabla " _
        + vbCr + "From com_honorarios " _
        + vbCr + "WHERE (((com_honorarios.numreg)<>'000001')) " _
        + vbCr + "GROUP BY 2, 'Compras', Date(), Month(com_honorarios.fchreg), 219 " _
        + vbCr + "HAVING (((Max(com_honorarios.fchreg))>= " & xAnoInicio & "))"
    
    
    CVENTAS_ = "SELECT 7 AS Orden, 'Ventas' AS Area, Date() AS FechaActual, Max(vta_ventas.fchdoc) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(vta_ventas.fchreg) AS Mes, 18 AS idTabla " _
        + vbCr + "From vta_ventas " _
        + vbCr + "WHERE (((vta_ventas.numreg)<>'000001')) " _
        + vbCr + "GROUP BY 7, 'Ventas', Date(), Month(vta_ventas.fchreg), 18 " _
        + vbCr + "HAVING (((Max(vta_ventas.fchreg))>= " & xAnoInicio & "))"

    CPRODUCCION_ = "SELECT 4 AS Orden, 'Produccion' AS Area, Date() AS FechaActual, Max(pro_produccion.dia) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(pro_produccion.dia) AS Mes, 92 AS idTabla " _
        + vbCr + "From pro_produccion " _
        + vbCr + "GROUP BY 4, 'Produccion', Date(), Month(pro_produccion.dia), 92 " _
        + vbCr + "HAVING (((Max(pro_produccion.dia))>= " & xAnoInicio & "))"
        
    CALMACEN_ = "SELECT 1 AS Orden, 'Almacen' AS Area, Date() AS FechaActual, Max(alm_ingreso.fching) AS Ultimoregistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(alm_ingreso.fching) AS Mes, 8 AS idTabla " _
        + vbCr + "From alm_ingreso " _
        + vbCr + "GROUP BY 1, 'Almacen', Date(), Month(alm_ingreso.fching), 8 " _
        + vbCr + "HAVING (((Max(alm_ingreso.fching))>= " & xAnoInicio & "))"
    
    CINGRESOS_ = "SELECT 8 AS Orden, 'Ingresos' AS Area, Date() AS FechaActual, Max(tes_caja.fchope) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(tes_caja.fchreg) AS Mes, 43 AS idTabla " _
        + vbCr + "From tes_caja " _
        + vbCr + "WHERE (((tes_caja.tipmov)=1) AND ((tes_caja.numreg)<>'000001')) " _
        + vbCr + "GROUP BY 8, 'Ingresos', Date(), Month(tes_caja.fchreg), 43 " _
        + vbCr + "HAVING (((Max(tes_caja.fchreg))>= " & xAnoInicio & "))"
        
    CEGRESOS_ = "SELECT 9 AS Orden, 'Egresos' AS Area, Date() AS FechaActual, Max(tes_caja.fchope) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(tes_caja.fchreg) AS Mes, 44 AS idTabla " _
        + vbCr + "From tes_caja " _
        + vbCr + "WHERE (((tes_caja.tipmov)=2) AND ((tes_caja.numreg)<>'000001')) " _
        + vbCr + "GROUP BY 9, 'Egresos', Date(), Month(tes_caja.fchreg), 44 " _
        + vbCr + "HAVING (((Max(tes_caja.fchreg))>= " & xAnoInicio & "))"
        
    CRETENCION_ = "SELECT 10 AS Orden, 'Retencion' AS Area, Date() AS FechaActual, Max(con_retencion.fchemi) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(con_retencion.fchreg) AS Mes, 31 AS idTabla " _
        + vbCr + "From con_retencion " _
        + vbCr + "WHERE (((con_retencion.numreg)<>'000001')) " _
        + vbCr + "GROUP BY 10, 'Retencion', Date(), Month(con_retencion.fchreg), 31 " _
        + vbCr + "HAVING (((Max(con_retencion.fchreg))>= " & xAnoInicio & "))"

    CPERCEPCION_ = "SELECT 11 AS Orden, 'Percepcion' AS Area, Date() AS FechaActual, Max(con_percepcion.fchdoc) AS UltimoRegistro, Sum(1) AS CantidadRegistro, '' As Cuenta, Month(con_percepcion.fchreg) AS Mes, 30 AS idTabla " _
        + vbCr + "From con_percepcion " _
        + vbCr + "WHERE (((con_percepcion.numreg)<>'000001')) " _
        + vbCr + "GROUP BY 11, 'Percepcion', Date(), Month(con_percepcion.fchreg), 30 " _
        + vbCr + "HAVING (((Max(con_percepcion.fchreg))>= " & xAnoInicio & "))"

    CCANJE_ = "SELECT 12 As Orden, 'Canje de Documentos' AS Area, Date() AS FechaActual, Max(con_canjes.fchemi) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(con_canjes.fchreg) AS Mes, 136 AS idTabla " _
        + vbCr + "From con_canjes " _
        + vbCr + "WHERE (((con_canjes.numreg)<>'000001')) " _
        + vbCr + "GROUP BY 12, 'Canje de Documentos', Date(), Month(con_canjes.fchreg), 136 " _
        + vbCr + "HAVING (((Max(con_canjes.fchreg))>= " & xAnoInicio & "))"

    CGUIA_ = "SELECT 6 As Orden, 'Guias de Remision' AS Area, Date() AS FechaActual, Max(vta_guia.fchemiord) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(vta_guia.fchemiord) AS Mes, 17 AS idTabla " _
        + vbCr + "From vta_guia " _
        + vbCr + "GROUP BY 6, 'Guias de Remision', Date(), Month(vta_guia.fchemiord), 17 " _
        + vbCr + "HAVING (((Max(vta_guia.fchemiord))>= " & xAnoInicio & "))" _
    
    CTAREAS_ = "SELECT 5 As Orden, 'Registro de Tareas' AS Area, Date() AS FechaActual, Max(pro_controltar.fchtra) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' As Cuenta, Month(pro_controltar.fchtra) AS Mes, 179 AS idTabla " _
        + vbCr + "From pro_controltar " _
        + vbCr + "GROUP BY 5, 'Registro de Tareas', Date(), Month(pro_controltar.fchtra), 179 " _
        + vbCr + "HAVING (((Max(pro_controltar.fchtra))>= " & xAnoInicio & "))"
        
    CCONCILIACION_ = "SELECT 13 AS Orden, 'Conciliacion' AS Area, Date() AS FechaActual, Max(tes_conci.fchini) AS UltimoRegistro, Sum(1) AS CantidadRegistros, mae_bancos.abrev & ' ' & mae_banconumcta.numcue & ' ' & mae_moneda.simbolo  AS Cuenta, Month(tes_conci.fchini) AS Mes, 123 AS idTabla " _
        + vbCr + "FROM ((tes_conci LEFT JOIN mae_banconumcta ON tes_conci.idbcocta = mae_banconumcta.id) LEFT JOIN mae_bancos ON mae_banconumcta.idban = mae_bancos.id) LEFT JOIN mae_moneda ON tes_conci.idmon = mae_moneda.id " _
        + vbCr + "GROUP BY 'Conciliacion', Date(), mae_banconumcta.numcue, mae_moneda.simbolo, mae_bancos.abrev, Month(tes_conci.fchini), 123 " _
        + vbCr + "HAVING (((Max(tes_conci.fchini)) Is Not Null))"

    CPROVISIONES_ = "SELECT 15 AS Orden, 'Provisiones' AS Area, Date() AS FechaActual, Max(con_proviciones.fchdoc) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' AS Cuenta, Month([con_proviciones].[fchreg]) AS Mes, 122 AS idTabla " _
        + vbCr + "From con_proviciones " _
        + vbCr + "GROUP BY 15, 'Provisiones', Date(), '', Month([con_proviciones].[fchreg]), 122 " _
        + vbCr + "HAVING (((Max(con_proviciones.fchreg))>= " & xAnoInicio & "))"
    
    CPEDIDOS_ = "SELECT 16 AS Orden, 'Pedidos' AS Area, Date() AS FechaActual, Max(ped_pedido.fchemi) AS UltimoRegistro, Sum(1) AS CantidadRegistros, '' AS Cuenta, Month([ped_pedido].[fchemi]) AS Mes, 224 AS idTabla " _
        + vbCr + "From ped_pedido " _
        + vbCr + "GROUP BY 16, 'Pedidos', Date(), '', Month([ped_pedido].[fchemi]), 224, Year([ped_pedido].[fchemi]) " _
        + vbCr + "HAVING (((Max(ped_pedido.fchemi))>= " & xAnoInicio & ")) AND Year([ped_pedido].[fchemi])=" & AnoTra & " "
    
    cSQL = CHONORARIOS_ + vbCr + "UNION" + vbCr + CCOMPRAS_ + vbCr + "UNION" + vbCr + CVENTAS_ + vbCr + "UNION" _
        + vbCr + CPRODUCCION_ + vbCr + "UNION" + vbCr + CALMACEN_ + vbCr + "UNION" + vbCr + CINGRESOS_ + vbCr + "UNION" _
        + vbCr + CEGRESOS_ + vbCr + "UNION" + vbCr + CRETENCION_ + vbCr + "UNION" + vbCr + CPERCEPCION_ + vbCr + "UNION" _
        + vbCr + CCANJE_ + vbCr + "UNION" + vbCr + CGUIA_ + vbCr + "UNION" + vbCr + CTAREAS_ + vbCr + "UNION" _
        + vbCr + CCONCILIACION_ + vbCr + "UNION" + vbCr + CBOLETA_ + vbCr + "UNION" + vbCr + CPROVISIONES_ + vbCr + "UNION" _
        + vbCr + CPEDIDOS_
        
           
    cSQL = "TRANSFORM Sum(miConsulta.CantidadRegistros) AS SumaDeCantidadRegistros " _
        + vbCr + "SELECT miConsulta.Orden, miConsulta.Area, miConsulta.FechaActual, miConsulta.Cuenta, miConsulta.idTabla, Max(miConsulta.UltimoRegistro) AS MaxUltimoRegistro, Sum(miConsulta.CantidadRegistros) AS TotalRegistros " _
        + vbCr + "FROM " _
        + vbCr + "(" _
        + vbCr + cSQL _
        + vbCr + ") AS miConsulta " _
        + vbCr + "GROUP BY miConsulta.Orden, miConsulta.Area, miConsulta.FechaActual, miConsulta.Cuenta, miConsulta.idTabla " _
        + vbCr + "Pivot miConsulta.Mes;"

    
    RST_Busq RstControl, cSQL, xCon
    
    Fg.Rows = 1
    
    If RstControl.State = 0 Then Exit Sub
    If RstControl.RecordCount = 0 Then Exit Sub
        
    RstControl.MoveFirst
    
    For A = 1 To RstControl.RecordCount
        FraProgreso.Refresh
        Fg.Refresh
        Fg.Rows = Fg.Rows + 1
        Fg.TextMatrix(A, 1) = NulosN(RstControl("Orden"))
        Fg.TextMatrix(A, 2) = NulosC(RstControl("Area"))
        Fg.TextMatrix(A, 3) = NulosC(RstControl("MaxUltimoRegistro"))
        Fg.TextMatrix(A, 4) = NulosC(RstControl("FechaActual"))
        Fg.TextMatrix(A, 5) = DateDiff("d", CDate(RstControl("MaxUltimoRegistro")), CDate(RstControl("FechaActual")))
        Fg.TextMatrix(A, 6) = NulosC(RstControl("Cuenta"))
        For B = 1 To 12
            On Error Resume Next
            Fg.TextMatrix(A, B + 6) = NulosN(RstControl("" & B & ""))
        Next B
        
        Fg.TextMatrix(A, B + 6) = NulosN(RstControl("TotalRegistros"))
        Fg.TextMatrix(A, B + 7) = NulosN(RstControl("idTabla"))

        RstControl.MoveNext
    Next A
    
    FraProgreso.Visible = False
    
    Agregando = True
    Agregando = False
End Sub

Sub EXPORTAR()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "CONTROL DE REGISTROS"

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg, TITULO_, "", "", TITULO_
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub

Private Sub iniciarCampos()
    Fg.AllowUserResizing = flexResizeColumns
    Fg.AutoSearch = flexSearchFromTop
    Fg.ExplorerBar = flexExSortShow
    Fg.SelectionMode = flexSelectionByRow
    Fg.ForeColorSel = &H80000005
    Fg.BackColorSel = &H80&
    Fg.Editable = flexEDKbdMouse
    Fg.Rows = 1
    Fg.Cols = 21
    Fg.FrozenCols = 3
'    Fg.ColWidth(4) = 0
    Fg.ColWidth(20) = 0
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO QUE SE EJECUTARA CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    iniciarCampos
    QueHace = 3
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 8100 Then Me.Height = 8100
    
    ' Se dimensiona la cabecera
    Label2.Width = Me.Width - 100
    Fg.Width = Me.Width - 250
    Fg.Height = Me.Height - 1530
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 10 Then ' Buscar
        pCargarGrid
    End If
    
    If Button.Index = 14 Then ' Exportar Excel
        EXPORTAR
    End If
    
    If Button.Index = 17 Then ' Salir
        Set RstControl = Nothing
        Unload Me
    End If
End Sub
