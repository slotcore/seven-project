VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmBalanceGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Estados Financieros"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6750
      Left            =   15
      TabIndex        =   1
      Top             =   825
      Width           =   11850
      _cx             =   20902
      _cy             =   11906
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   " Balance General |Resultados x Funcion|Resultados x Naturaleza|Cambios en el Patrimonio"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   0
      Position        =   1
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   6330
         Left            =   12795
         TabIndex        =   4
         Top             =   45
         Width           =   11760
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   930
            Left            =   75
            TabIndex        =   24
            Top             =   60
            Width           =   11625
            Begin VB.Label LblPeriodo3 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "LblPeriodo"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   105
               TabIndex        =   27
               Top             =   360
               Width           =   11415
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "(Expresado en Nuevos Soles)"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   105
               TabIndex        =   26
               Top             =   600
               Width           =   11415
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Estado de Gan. y Perdidas x Naturaleza"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   105
               TabIndex        =   25
               Top             =   30
               Width           =   11415
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   5250
            Index           =   2
            Left            =   75
            TabIndex        =   8
            Top             =   990
            Width           =   11625
            _cx             =   20505
            _cy             =   9260
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   16777215
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmBalanceGeneral.frx":0000
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6330
         Left            =   12495
         TabIndex        =   3
         Top             =   45
         Width           =   11760
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   930
            Left            =   75
            TabIndex        =   20
            Top             =   60
            Width           =   11625
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Estado de Gan. y Perdidas x Funcion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   105
               TabIndex        =   23
               Top             =   30
               Width           =   11415
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "(Expresado en Nuevos Soles)"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   105
               TabIndex        =   22
               Top             =   600
               Width           =   11415
            End
            Begin VB.Label LblPeriodo2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "LblPeriodo"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   105
               TabIndex        =   21
               Top             =   360
               Width           =   11415
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   5250
            Index           =   1
            Left            =   75
            TabIndex        =   7
            Top             =   990
            Width           =   11625
            _cx             =   20505
            _cy             =   9260
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   16777215
            TreeColor       =   16777215
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmBalanceGeneral.frx":0078
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6330
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Width           =   11760
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   930
            Left            =   75
            TabIndex        =   16
            Top             =   60
            Width           =   11625
            Begin VB.Label LblPeriodo 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "LblPeriodo"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   105
               TabIndex        =   18
               Top             =   360
               Width           =   11415
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "(Expresado en Nuevos Soles)"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   105
               TabIndex        =   19
               Top             =   600
               Width           =   11415
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Balance General"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   105
               TabIndex        =   17
               Top             =   30
               Width           =   11415
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5265
            Left            =   75
            TabIndex        =   5
            Top             =   990
            Width           =   11625
            _cx             =   20505
            _cy             =   9287
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
            BackColorFixed  =   16777215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   16777215
            TreeColor       =   16777215
            FloodColor      =   192
            SheetBorder     =   16777215
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
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmBalanceGeneral.frx":00F1
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
   End
   Begin VB.Frame Frame1 
      Height          =   810
      Left            =   15
      TabIndex        =   0
      Top             =   -30
      Width           =   11865
      Begin VB.CommandButton CmdSalir 
         Height          =   570
         Left            =   11010
         Picture         =   "FrmBalanceGeneral.frx":01A3
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Exportar a Excel"
         Top             =   165
         Width           =   630
      End
      Begin VB.CommandButton CmdImprimir 
         Height          =   570
         Left            =   9570
         Picture         =   "FrmBalanceGeneral.frx":04AD
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   165
         Width           =   630
      End
      Begin VB.CommandButton CmdExp 
         Height          =   570
         Left            =   10230
         Picture         =   "FrmBalanceGeneral.frx":07B7
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Exportar a Excel"
         Top             =   165
         Width           =   630
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   1020
         TabIndex        =   9
         Top             =   300
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
      End
      Begin VB.CommandButton CmdBus 
         Height          =   570
         Left            =   8910
         Picture         =   "FrmBalanceGeneral.frx":12C1
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   165
         Width           =   630
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   3285
         TabIndex        =   10
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
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
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Final"
         Height          =   195
         Index           =   1
         Left            =   2415
         TabIndex        =   12
         Top             =   345
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Inicio"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   345
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmBalanceGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : ESTADOSFINANCIEROS.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : MUESTRA LOS 4 ESTADOS FINANCIEROS, EN FUNCION A CRITERIOS ESPECIFICADOS POR EL
'*                    USUARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 27/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstSal As New ADODB.Recordset
Dim RstTmp As New ADODB.Recordset
Dim RstEst As New ADODB.Recordset
Dim SeEjecuto As Boolean                  ' VARIABLE PARA VERIFICAR QUE EL EVENTO ACTIVATE SE EJECUTA UNA SOLA VEZ


'*****************************************************************************************************
'* Nombre           : CrearRSTTMPBal
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CREA UN RECORDSET TEMPORAL PARA ALMACENAR DATOS PARA EL BALANCE
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CrearRSTTMPBal()
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "idbal":         xCampos(0, 1) = "N":      xCampos(0, 2) = "8"
    xCampos(1, 0) = "descri":        xCampos(1, 1) = "C":      xCampos(1, 2) = "100"
    xCampos(2, 0) = "salinideb":     xCampos(2, 1) = "D":      xCampos(2, 2) = "2"
    xCampos(3, 0) = "salinihab":     xCampos(3, 1) = "D":      xCampos(3, 2) = "2"
    xCampos(4, 0) = "movperdeb":     xCampos(4, 1) = "D":      xCampos(4, 2) = "2"
    xCampos(5, 0) = "movperhab":     xCampos(5, 1) = "D":      xCampos(5, 2) = "2"
    xCampos(6, 0) = "saldeb":        xCampos(6, 1) = "D":      xCampos(6, 2) = "2"
    xCampos(7, 0) = "salhab":        xCampos(7, 1) = "D":      xCampos(7, 2) = "2"
    xCampos(8, 0) = "resultado":     xCampos(8, 1) = "D":      xCampos(8, 2) = "2"
    Set RstTmp = xFun.CrearRstTMP(xCampos)
    RstTmp.Open
End Sub

'*****************************************************************************************************
'* Nombre           : EstadoFuncion
'* Tipo             : FUNCCION
'* Descripcion      : MUESTRA LOS ESTADOS FINANCIEROS POR FUNCION Y NATURALEZA
'* Paranetros       : NOMBRE    |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Cual      |  Integer     |  ESPECIFICA QUE BALANCE SE PRESENTARA
'*                                                1 = POR FUNCION
'*                                                2 = POR NATURALEZA
'* Devuelve         :
'*****************************************************************************************************
Sub EstadoFuncion(Cual As Integer)
    '1 = Por funcion
    '2 = Por naturalesa
    Dim xTipo As Integer
    
    If Cual = 1 Then
        LblPeriodo2.Caption = "Al " + Format(CDate(TxtFchFin.Valor), "dd") + " de " + Format(CDate(TxtFchFin.Valor), "mmmm") + " del " + Format(CDate(TxtFchFin.Valor), "yyyy")
    Else
        LblPeriodo3.Caption = "Al " + Format(CDate(TxtFchFin.Valor), "dd") + " de " + Format(CDate(TxtFchFin.Valor), "mmmm") + " del " + Format(CDate(TxtFchFin.Valor), "yyyy")
    End If
    
    Dim RstMiE  As New ADODB.Recordset
    Dim TotAct, TotPas As Double
    Dim A, xFila As Integer
    Dim RstEst As New ADODB.Recordset
    
    CrearRSTTMPBal
    If Cual = 1 Then xTipo = 1
    If Cual = 2 Then xTipo = 2
    
    ' PREPARAMOS LAS SENTENCIAS SQL PARA PROCESAR LOS DATOS
    RST_Busq RstMiE, "SELECT con_estadoscab.* FROM con_estados LEFT JOIN con_estadoscab ON con_estados.id = con_estadoscab.idcab Where (((con_estados.tipo) = " & xTipo & ") " _
        & " And ((con_estados.activo) = -1)) ORDER BY con_estadoscab.orden", xCon

    RST_Busq RstEst, "TRANSFORM Sum(con_diario.impdebsol) AS SumaDeimpdebsol SELECT con_estadoscab.idcab, con_estadoscab.id, con_estadoscab.descripcion, " _
        & " Sum(IIf([impdebdol]<>0,[impdebdol]*[con_tc].[impven],[impdebsol])) AS Totdebsol, Sum(IIf([imphabdol]<>0,[imphabdol]*[con_tc].[impven],[imphabsol])) AS tothabsol, " _
        & " [Totdebsol]-[Tothabsol] AS saldo " _
        & " FROM (con_estadoscab LEFT JOIN (con_estadosdet LEFT JOIN con_diario ON con_estadosdet.idcuenta = con_diario.idcue) ON (con_estadoscab.id = con_estadosdet.idest) " _
        & " AND (con_estadoscab.idcab = con_estadosdet.idcab)) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha WHERE (((con_estadoscab.idcab)=" & RstMiE("idcab") & ") " _
        & " AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))) GROUP BY con_estadoscab.idcab, con_estadoscab.id, " _
        & " con_estadoscab.descripcion PIVOT con_estadosdet.idcuenta", xCon
    
    ' Cargamos al recorset temporal todos los movimientos
    Dim xTotal As Double
    If RstMiE.RecordCount <> 0 Then
        RstMiE.MoveFirst
        For A = 1 To RstMiE.RecordCount
            RstTmp.AddNew
            RstTmp("idbal") = RstMiE("id")
            RstTmp("descri") = RstMiE("descripcion")
            
            If NulosC(RstMiE("formula")) <> "" Then
                xTotal = EjecutarFormula(RstMiE("formula"), RstTmp)
                If RstEst.RecordCount <> 0 Then
                    RstTmp.MoveFirst
                    RstTmp.Find "idbal = " & RstMiE("id") & ""
                    If RstTmp.EOF = False Then
                        RstTmp("resultado") = xTotal
                    End If
                End If
            Else
                If RstEst.RecordCount <> 0 Then
                    RstEst.MoveFirst
                    RstEst.Find "id = " & RstMiE("id") & ""
                    If RstEst.EOF = False Then
                        RstTmp("movperdeb") = RstEst("totdebsol")
                        RstTmp("movperhab") = RstEst("tothabsol")
                        RstTmp("resultado") = (RstEst("totdebsol") - RstEst("tothabsol"))
                        RstTmp("resultado") = Abs(RstTmp("resultado"))
                    Else
                        RstTmp("resultado") = 0
                    End If
                End If
            End If
            RstMiE.MoveNext
            If RstMiE.EOF = True Then Exit For
        Next A
    End If
    
    Fg2(Cual).Rows = 1
    
    ' Calculamos el salto de lineas adicionales
    RstMiE.Filter = "sallin = -1"
    A = RstMiE.RecordCount
    RstMiE.Filter = adFilterNone
    Fg2(Cual).Rows = (RstMiE.RecordCount + A) + 5
    
    xFila = 1
    For A = 1 To RstMiE.RecordCount
        xFila = xFila + 1
        Fg2(Cual).TextMatrix(xFila, 1) = ""
        Fg2(Cual).TextMatrix(xFila, 2) = RstMiE("descripcion")
        
        RstTmp.MoveFirst
        RstTmp.Find "idbal = " & RstMiE("id") & ""
        If RstTmp.EOF = False Then
            Fg2(Cual).TextMatrix(xFila, 3) = NulosN(RstTmp("resultado"))
            If RstTmp("resultado") > 0 Then Fg2(Cual).TextMatrix(xFila, 3) = Format(Fg2(Cual).TextMatrix(xFila, 3), "###,###0")
            If RstTmp("resultado") < 0 Then Fg2(Cual).TextMatrix(xFila, 3) = Format(Abs(NulosN(Fg2(Cual).TextMatrix(xFila, 3))), "(###,###0)")
        Else
            Fg2(Cual).TextMatrix(xFila, 3) = ""
        End If
        
        If RstMiE("negrita") = -1 Then
            ' Pintamos de negrita la linea
            With Fg2(Cual)
                .Select xFila, 1, xFila, 4
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
        End If
        
        If RstMiE("sallin") = -1 Then
            ' Agregamos un fila
            xFila = xFila + 1
        End If
        RstMiE.MoveNext
        If RstMiE.EOF = True Then Exit For
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : BalanceGeneral
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL BALANCE GENERAL EN FUNCION ACRITERIOS APLICADOS POR EL USUARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub BalanceGeneral()
    LblPeriodo.Caption = "Al " + Format(CDate(TxtFchFin.Valor), "dd") + " de " + Format(CDate(TxtFchFin.Valor), "mmmm") + " del " + Format(CDate(TxtFchFin.Valor), "yyyy")
    Dim RstMiB As New ADODB.Recordset
    Dim TotAct, TotPas As Double
    Dim A, B As Integer
    
    On Error GoTo LaCague
    
    Fg1.Rows = Fg1.FixedRows
    Fg1.Rows = 50
    DoEvents
    Me.MousePointer = vbHourglass
    
    RST_Busq RstMiB, "SELECT con_balance.* From con_balance Where (((con_balance.activo) = -1))", xCon
    
    Set RstTmp = Nothing
    CrearRSTTMPBal
    
    Set RstTmp = CargarTMPBalance(TxtFchIni.Valor, TxtFchFin.Valor, RstMiB("id"), RstTmp)
    
    Dim xId As Integer
    If RstMiB.RecordCount = 0 Then
        MsgBox "No se ha especificado un balance por defecto, vaya el modulo Contabilidada, " & Chr(13) _
            & " Opcion [Configuracion de Estados Financieros] y selecciones el sub menu [Balance]", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstMiB = Nothing
        Exit Sub
    End If
    xId = RstMiB("id")
    Set RstMiB = Nothing
    RST_Busq RstMiB, "SELECT con_balancecab.* From con_balancecab Where ((con_balancecab.idcab =" & xId & ") AND (con_balancecab.tipo = 1)) ORDER BY con_balancecab.orden", xCon
    
    ' IMPRIMIMOS LA COLUMNA DEL ACTIVO
    Dim xFila As Integer
    Dim xImpFor As Double
    
    xFila = 0
    For A = 1 To RstMiB.RecordCount
        xFila = xFila + 1
        Fg1.TextMatrix(xFila, 1) = NulosC(RstMiB("descripcion"))
        If RstTmp.State = 1 Then
            RstTmp.MoveFirst
            RstTmp.Find "idbal = " & NulosN(RstMiB("id")) & ""
            
            If RstTmp.EOF = False Then
                Fg1.TextMatrix(xFila, 2) = Format(NulosN(RstTmp("resultado")), "###,###0")
                TotAct = TotAct + NulosN(RstTmp("resultado"))
            End If
        End If
        
        If NulosC(RstMiB("formula")) <> "" Then
            ' Ejecutamos una formula
            xImpFor = EjecutarFormula(RstMiB("formula"), RstTmp)
            If xImpFor >= 0 Then Fg1.TextMatrix(xFila, 2) = Format(xImpFor, "###,###0")
            If xImpFor < 0 Then Fg1.TextMatrix(xFila, 2) = Format(Abs(xImpFor), "(###,###0)")
        End If
        
        If RstMiB("negrita") = -1 Then
            ' Pintamos de negrita la linea
            With Fg1
                .Select xFila, 1, xFila, 2
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
        End If
        
        If RstMiB("sallin") = -1 Then
            ' Hacemos un saltgo de pagina
            xFila = xFila + 1
        End If
        
        RstMiB.MoveNext
        
        If RstMiB.EOF = True Then
            Exit For
        End If
    Next A
        
    RST_Busq RstMiB, "SELECT con_balancecab.* From con_balancecab Where ((con_balancecab.idcab =" & xId & ") AND (con_balancecab.tipo = 2)) ORDER BY con_balancecab.orden", xCon
    
    ' IMPRIMIMOS LA COLUMNA DEL PASIVO
    xFila = 0
    For A = 1 To RstMiB.RecordCount
        xImpFor = 0
        xFila = xFila + 1
        Fg1.TextMatrix(xFila, 5) = NulosC(RstMiB("descripcion"))
        
        If RstTmp.State <> 0 Then
            If RstTmp.RecordCount <> 0 Then
                RstTmp.MoveFirst
                RstTmp.Find "idbal = " & NulosN(RstMiB("id")) & ""
                
                If RstTmp.EOF = False Then
                    Fg1.TextMatrix(xFila, 6) = Format(Abs(RstTmp("resultado")), "###,###0")
                    TotPas = TotPas + Abs(RstTmp("resultado"))
                End If
            End If
        End If
        
        If NulosC(RstMiB("formula")) <> "" Then
            ' Ejecutamos una formula
            xImpFor = EjecutarFormula(RstMiB("formula"), RstTmp)
            If xImpFor >= 0 Then Fg1.TextMatrix(xFila, 6) = Format(xImpFor, "###,###0")
            If xImpFor < 0 Then Fg1.TextMatrix(xFila, 6) = Format(Abs(xImpFor), "(###,###0)")
        End If
        
        If NulosN(RstMiB("negrita")) = -1 Then
            ' Pintamos de negrita la linea
            With Fg1
                .Select xFila, 5, xFila, 6
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
        End If
                
        If NulosN(RstMiB("sallin")) = -1 Then
            xFila = xFila + 1
        End If
        
        RstMiB.MoveNext
        
        If RstMiB.EOF = True Then
            Exit For
        End If
    Next A
    
    ' Imprimimos la suma de ambas columnas
    xFila = xFila + 1
    Fg1.TextMatrix(xFila, 5) = "Resultado del Ejercicio"
    Fg1.TextMatrix(xFila, 6) = (TotAct - TotPas)
    
    If (TotAct - TotPas) >= 0 Then
        Fg1.TextMatrix(xFila, 6) = Format(NulosN(Fg1.TextMatrix(xFila, 6)), "#,###,###0")
    Else
        Fg1.TextMatrix(xFila, 6) = Format(NulosN(Fg1.TextMatrix(xFila, 6)), "(#,###,###0)")
    End If
          
    TotPas = TotPas + Abs(NulosN(Fg1.TextMatrix(xFila, 6)))
    
    xFila = xFila + 2
    Fg1.TextMatrix(xFila, 1) = "TOTAL ACTIVO"
    If TotAct > 0 Then Fg1.TextMatrix(xFila, 2) = Format(TotAct, "#,###,###0")
    If TotAct < 0 Then Fg1.TextMatrix(xFila, 2) = Format(Abs(TotAct), "(#,###,###0)")
    
    Fg1.TextMatrix(xFila, 5) = "TOTAL PASIVO"
    If TotPas > 0 Then Fg1.TextMatrix(xFila, 6) = Format(TotPas, "#,###,###0")
    If TotPas < 0 Then Fg1.TextMatrix(xFila, 6) = Format(Abs(TotPas), "(#,###,###0)")
    
    With Fg1
        .Select xFila, 1, xFila, 6
        .FillStyle = flexFillRepeat
        .CellFontBold = True
        .Select 1, 1, 1, 1
    End With
    
    Me.MousePointer = vbDefault
    Exit Sub

LaCague:
    Me.MousePointer = vbDefault
End Sub

Private Sub CmdBus_Click()
    ' EJECUTA LA CONSULTA DE LOS ESTADOS FINANCIEROS EN FUNCION A LA PESTAÑA SELECCIONADA
    If TabOne1.CurrTab = 0 Then BalanceGeneral
    If TabOne1.CurrTab = 1 Then EstadoFuncion 1
    If TabOne1.CurrTab = 2 Then EstadoFuncion 2
End Sub

Private Sub CmdSalir_Click()
    Unload Me
    Set RstEst = Nothing
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE LE FORMULARIOS
    If SeEjecuto = False Then
        SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE LE FORMULARIO
    SeEjecuto = False
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    Frame4.BackColor = &H8000000F
    TxtFchIni.Valor = CDate("01/01/" & Trim(AnoTra))
    TxtFchFin.Valor = Date
    LblPeriodo.Caption = ""
    LblPeriodo2.Caption = ""
    LblPeriodo3.Caption = ""
    TabOne1.CurrTab = 0
End Sub
