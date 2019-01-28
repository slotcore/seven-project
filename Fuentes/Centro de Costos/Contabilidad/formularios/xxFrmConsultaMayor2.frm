VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaMayor2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Mayor Auxiliar"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_msg 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   270
      Left            =   6015
      TabIndex        =   25
      Top             =   7095
      Visible         =   0   'False
      Width           =   5730
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Obs: Se recomienda minimizar la ventana para agilizar el proceso"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   30
         TabIndex        =   26
         Top             =   15
         Width           =   5535
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   795
      Left            =   3045
      TabIndex        =   18
      Top             =   2730
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   105
         TabIndex        =   19
         Top             =   330
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         Caption         =   "Interrumpir = ESC"
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
         Height          =   255
         Index           =   2
         Left            =   4365
         TabIndex        =   24
         Top             =   90
         Width           =   1530
      End
      Begin VB.Label Label3 
         Caption         =   "Procesando Asientos"
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
         Height          =   255
         Left            =   165
         TabIndex        =   20
         Top             =   90
         Width           =   4020
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   960
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   5925
         X2              =   5925
         Y1              =   15
         Y2              =   945
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   5940
         Y1              =   15
         Y2              =   30
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5925
         Y1              =   780
         Y2              =   765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11805
      Begin VB.CheckBox chk 
         Caption         =   "Procesar Todas las Cuentas"
         Height          =   360
         Left            =   9675
         TabIndex        =   27
         Top             =   855
         Width           =   1950
      End
      Begin VB.CommandButton Command1 
         Height          =   570
         Left            =   11115
         Picture         =   "FrmConsultaMayor2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exportar a Excel"
         Top             =   210
         Width           =   630
      End
      Begin VB.Frame Frame4 
         Caption         =   "[ Ordenado por ]"
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
         Height          =   1080
         Left            =   7860
         TabIndex        =   17
         Top             =   120
         Width           =   1725
         Begin VB.OptionButton opt 
            Caption         =   "Nº Registro"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   28
            Top             =   825
            Width           =   1425
         End
         Begin VB.OptionButton opt 
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   22
            Top             =   225
            Value           =   -1  'True
            Width           =   1560
         End
         Begin VB.OptionButton opt 
            Caption         =   "Nº Documento"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   21
            Top             =   525
            Width           =   1440
         End
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Eliminar"
         Height          =   435
         Left            =   7005
         TabIndex        =   6
         Top             =   750
         Width           =   750
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Agregar "
         Height          =   435
         Left            =   7005
         TabIndex        =   5
         Top             =   195
         Width           =   750
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   990
         Left            =   2355
         TabIndex        =   4
         Top             =   195
         Width           =   4620
         _cx             =   8149
         _cy             =   1746
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
         BackColorSel    =   128
         ForeColorSel    =   16777215
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
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsultaMayor2.frx":0B0A
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
      Begin VB.CommandButton CmdMuestra 
         Height          =   570
         Left            =   9690
         Picture         =   "FrmConsultaMayor2.frx":0B8F
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Buscar"
         Top             =   210
         Width           =   630
      End
      Begin VB.CommandButton CmdImprimir 
         Height          =   570
         Left            =   10402
         Picture         =   "FrmConsultaMayor2.frx":0FD1
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir"
         Top             =   210
         Width           =   630
      End
      Begin VB.OptionButton OptDolares 
         Caption         =   "Dolares"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1185
         TabIndex        =   3
         Top             =   990
         Width           =   900
      End
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   990
         Width           =   900
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   1005
         TabIndex        =   0
         Top             =   240
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   1005
         TabIndex        =   1
         Top             =   555
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
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   7800
         X2              =   7800
         Y1              =   195
         Y2              =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Inicio"
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Final"
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   645
         Width           =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   2295
         X2              =   2295
         Y1              =   195
         Y2              =   1200
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6090
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   11805
      _cx             =   20823
      _cy             =   10742
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
      FrontTabForeColor=   -2147483630
      Caption         =   "   Detalle   |   Resumen   "
      Align           =   0
      CurrTab         =   1
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
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5670
         Left            =   45
         TabIndex        =   14
         Top             =   45
         Width           =   11715
         Begin VSFlex7Ctl.VSFlexGrid Fg3 
            Height          =   5550
            Left            =   15
            TabIndex        =   15
            Top             =   60
            Width           =   11700
            _cx             =   20637
            _cy             =   9790
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
            Rows            =   1
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsultaMayor2.frx":12DB
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
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5670
         Left            =   -12360
         TabIndex        =   13
         Top             =   45
         Width           =   11715
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5550
            Left            =   15
            TabIndex        =   16
            Top             =   60
            Width           =   11700
            _cx             =   20637
            _cy             =   9790
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
            BackColorSel    =   128
            ForeColorSel    =   16777215
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
            Rows            =   1
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsultaMayor2.frx":13E1
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
End
Attribute VB_Name = "FrmConsultaMayor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstTmp As New ADODB.Recordset
Dim RstTmp2 As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE



Private Sub chk_Click()
    If Fg2.Rows = 1 Then Exit Sub
    Fg2.Rows = 1
End Sub

Private Sub CmdAdd_Click()
    On Error GoTo error
    Dim xfrm As New SGI2_funciones.formularios
    Dim Rst As New ADODB.Recordset
    Dim k As Integer
    Dim MSG_CUENTA As String    '--MUSTRA EL MENSAJE SI DESEA AGREGAR UNA CUENTA, CUANDO YA EXISTE UNA CUENTA DE NIVEL SUPERIOR O NIVEL INFERIOR
                                '--NO MOSTRAR MENSAJE SOLO CUANDO LAS CUENTAS SEA DEL MISMO NIVEL
    
    If chk.Value = 1 Then chk.Value = 0
    
    Set Rst = xfrm.SelePlanCuentas(xCon)
    If Rst.State = 1 Then
        If Rst.RecordCount <> 0 Then
            If GRID_BUSCAR_VALOR(Fg2, 3, CStr(Rst("id") & ""), False) <> "-1" Then
                MsgBox "La cuenta contable Nº " + Trim(Rst("cuenta") & "") + " ya fue seleccionada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            For k = 1 To Fg2.Rows - 1
                If Len(Trim(Rst.Fields("cuenta") & "")) < Len(Trim(Fg2.TextMatrix(k, 1))) Then
                    
                    If Trim(Rst.Fields("cuenta") & "") = Mid(Trim(Fg2.TextMatrix(k, 1)), 1, Len(Trim(Rst.Fields("cuenta") & ""))) Then
                        MSG_CUENTA = "Ya agregó la cuenta Nª: " + Trim(Fg2.TextMatrix(k, 1)) + " cuyo nivel es Inferior a la cuenta Nº: " + Trim(Rst.Fields("cuenta") & "") + " que desea agregar" _
                                    + vbCr + "Sólo puede agregar Cuentas del mismo nivel " _
                                    + vbCr + "Si desea continuar elimine la fila que contenga la Cuenta Nº: " + Trim(Fg2.TextMatrix(k, 1))
                        Exit For
                    End If
                    
                Else
                    If Trim(Fg2.TextMatrix(k, 1)) = Mid(Trim(Rst.Fields("cuenta") & ""), 1, Len(Trim(Fg2.TextMatrix(k, 1)))) Then
                        MSG_CUENTA = "Ya agregó la cuenta Nª: " + Trim(Fg2.TextMatrix(k, 1)) + " cuyo nivel es Superior a la cuenta Nº: " + Trim(Rst.Fields("cuenta") & "") + " que desea agregar" _
                                    + vbCr + "Sólo puede agregar Cuentas del mismo nivel " _
                                    + vbCr + "Si desea continuar elimine la fila que contenga la Cuenta Nº: " + Trim(Fg2.TextMatrix(k, 1))
                        Exit For
                    End If
                    
                End If
            Next k
            If MSG_CUENTA <> "" Then
                MsgBox MSG_CUENTA, vbExclamation, xTitulo
                GoTo Salir
            End If
            
            
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = Rst("cuenta") & ""
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = Rst("descripcion") & ""
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Rst("id") & ""
        End If
    End If
Salir:
    Set xfrm = Nothing
    Set Rst = Nothing
    Exit Sub
error:
    Set xfrm = Nothing
    Set Rst = Nothing
    SHOW_ERROR Me.Name, "CmdAdd_Click"
End Sub

Private Sub CmdDel_Click()
    If Fg2.Row <= 0 Then Exit Sub
    If Fg2.Rows <= 1 Then
        MsgBox "No hay cuentas seleccionadas para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    Else
        Fg2.RemoveItem Fg2.Row
        Fg2.Refresh
    End If
End Sub

Private Sub CmdImprimir_Click()
    Dim xMes, xMoneda As String
    Dim X_PRINT As New SGI2_funciones.formularios

    On Error GoTo error
    Me.MousePointer = vbHourglass
    If Me.TabOne1.CurrTab = 0 Then
        FrmPrintMayor.Show vbModal
    Else
        xMes = Format(TxtFchIni.Valor, "mmmm")
        xMoneda = "Nuevos Soles"
        If MsgBox("Desea conservar el formato de la consulta", vbQuestion + vbYesNo, "Imprimir...") = vbNo Then Configurar_Grilla False
'        X_PRINT.Imprimir_x_VSFlexGrid_GRID_EN_RPT Fg1

        X_PRINT.Imprimir_x_VSFlexGrid Fg3, "LIBRO MAYOR ", "", "(Expresado en " + xMoneda + ")" + " ", False, True

        Set X_PRINT = Nothing
        
    End If
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "CmdImprimir_Click"
End Sub

Private Sub CmdMuestra_Click()
    If Validar_Consulta() = False Then Exit Sub

    BAND_INTERRUMPIR = False
    Me.ProgressBar1.Value = 1
    Me.TabOne1.CurrTab = 0
    Configurar_Grilla True
    MuestraMayor
    If BAND_INTERRUMPIR = True Then Exit Sub
    Me.TabOne1.CurrTab = 1
    DoEvents
    CargarResumen
    
End Sub

Sub CargarResumen()
    On Error GoTo error
    Dim RstRes As New ADODB.Recordset
    Dim A As Integer
    Dim xTotal1, xTotal2, xTotal3, xTotal4 As Double
Dim xAcumulado(7) As Double
   
'''    PreparaRST_Tmp
    
    Frame5.Left = 3413
    Frame5.Top = 2685
    Me.ProgressBar1.Value = 1
    Frame5.Visible = True
    Label3.Caption = "Procesando Resumen"
    DoEvents
    
    Dim N_SQL As String
    Dim N_SQL_WHERE As String
    If TxtFchIni.Valor <> TxtFchFin.Valor Then
        N_SQL_WHERE = " (con_diario.fchasi >=CDATE ('" + TxtFchIni.Valor + "') AND con_diario.fchasi <= CDATE('" + TxtFchFin.Valor + "')) "
    Else
        N_SQL_WHERE = " con_diario.fchasi = CDATE('" + TxtFchIni.Valor + "') "
    End If
    
    N_SQL = "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion AS descri , con_planctas.tipsal , "
    
    If Me.OptSoles = True Then
        N_SQL = N_SQL _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol=0,con_diario1.impdebsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.impdebdol)))) AS saldebesol  FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue  WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING con_diario1.idcue = con_diario.idcue ) AS saldebesol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol=0,con_diario1.imphabsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.imphabdol)))) AS salhabersol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING con_diario1.idcue = con_diario.idcue ) AS salhabersol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol=0,con_diario1.impdebsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.impdebdol))))  AS debesol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07') GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue  ) AS  debesol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol=0,con_diario1.imphabsol,IIf(con_tc1.impven Is Null,0,(con_tc1.impven*con_diario1.imphabdol))))  AS habersol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')  GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue  ) AS  habersol, " _
            + vbCr + " IIf(debesol Is Null,0+IIf(saldebesol Is Null,0,saldebesol),debesol+IIf(saldebesol Is Null,0,saldebesol)) AS maydebesol, " _
            + vbCr + " IIf(habersol Is Null,0+IIf(salhabersol Is Null,0,salhabersol),habersol+IIf(salhabersol Is Null,0,salhabersol)) AS mayhabersol, " _
            + vbCr + " (IIF (con_planctas.tipsal='D' OR con_planctas.tipsal IS NULL OR con_planctas.tipsal ='', (maydebesol -  mayhabersol), (mayhabersol - maydebesol))) as saldosol, " _
            + vbCr + " IIf(maydebesol>mayhabersol,(maydebesol-mayhabersol),0) AS deudorsol, " _
            + vbCr + " IIf(mayhabersol>maydebesol,(mayhabersol-maydebesol),0) AS acreedorsol "
    End If
                        
    If OptDolares.Value = True Then
        N_SQL = N_SQL _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol<>0,con_diario1.impdebdol,IIf(con_tc1.impven Is Null Or con_diario1.impdebsol=0,0,(con_diario1.impdebsol/con_tc1.impven)))) AS saldebedol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING (((con_diario1.idcue)=con_diario.idcue))) AS saldebedol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol<>0,con_diario1.imphabdol,IIf(con_tc1.impven Is Null Or con_diario1.imphabsol=0,0,(con_diario1.imphabsol/con_tc1.impven)))) AS salhaberdol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE con_diario1.fchasi < CDate('01/01/07') Or con_diario1.fchasi Is Null GROUP BY con_diario1.idcue  HAVING (((con_diario1.idcue)=con_diario.idcue))) AS salhaberdol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.impdebdol<>0,con_diario1.impdebdol,IIf(con_tc1.impven Is Null Or con_diario1.impdebsol=0,0,(con_diario1.impdebsol/con_tc1.impven)))) AS debedol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')  GROUP BY con_diario1.idcue HAVING con_diario1.idcue=con_diario.idcue ) AS  debedol, " _
            + vbCr + " (SELECT Sum(IIf(con_diario1.imphabdol<>0,con_diario1.imphabdol,IIf(con_tc1.impven Is Null Or con_diario1.imphabsol=0,0,(con_diario1.imphabsol/con_tc1.impven)))) AS haberdol FROM con_planctas AS con_planctas1 RIGHT JOIN (con_diario AS con_diario1 LEFT JOIN con_tc AS con_tc1 ON con_diario1.fchdoc = con_tc1.fecha) ON con_planctas1.id = con_diario1.idcue WHERE  con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')  GROUP BY con_diario1.idcue  HAVING con_diario1.idcue=con_diario.idcue ) AS  haberdol, " _
            + vbCr + " IIf(debedol Is Null,0+IIf(saldebedol Is Null,0,saldebedol),debedol+IIf(saldebedol Is Null,0,saldebedol)) AS maydebedol, " _
            + vbCr + " IIf(haberdol Is Null,0+IIf(salhaberdol Is Null,0,salhaberdol),haberdol+IIf(salhaberdol Is Null,0,salhaberdol)) AS mayhaberdol, " _
            + vbCr + " (IIF (con_planctas.tipsal='D' OR con_planctas.tipsal IS NULL OR con_planctas.tipsal ='', (maydebedol -  mayhaberdol), (mayhaberdol - maydebedol))) as saldodol, " _
            + vbCr + " IIf(maydebedol>mayhaberdol,(maydebedol-mayhaberdol),0) AS deudordol, " _
            + vbCr + " IIf(mayhaberdol > maydebedol, (mayhaberdol - maydebedol), 0) As acreedordol "
    End If
        
    N_SQL = N_SQL _
        + vbCr + " FROM con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue " _
        + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion,con_planctas.tipsal " _
        + vbCr + " ORDER BY con_planctas.cuenta, con_planctas.descripcion;"

    '--UNIFICANDO LOS NOMBRES DE LOS CAMPOS TANTO PARA DOLARES Y SOLES
    N_SQL = Replace(N_SQL, "saldebesol", "saldeb")
    N_SQL = Replace(N_SQL, "salhabersol", "salhab")
    N_SQL = Replace(N_SQL, "maydebesol", "maydeb")
    N_SQL = Replace(N_SQL, "mayhabersol", "mayhab")
    N_SQL = Replace(N_SQL, "debesol", "movdeb")
    N_SQL = Replace(N_SQL, "habersol", "movhab")
    N_SQL = Replace(N_SQL, "deudorsol", "deudor")
    N_SQL = Replace(N_SQL, "acreedorsol", "acreedor")
    

    N_SQL = Replace(N_SQL, "saldebedol", "saldeb")
    N_SQL = Replace(N_SQL, "salhaberdol", "salhab")
    N_SQL = Replace(N_SQL, "maydebedol", "maydeb")
    N_SQL = Replace(N_SQL, "mayhaberdol", "mayhab")
    N_SQL = Replace(N_SQL, "debedol", "movdeb")
    N_SQL = Replace(N_SQL, "haberdol", "movhab")
    N_SQL = Replace(N_SQL, "deudordol", "deudor")
    N_SQL = Replace(N_SQL, "acreedordol", "acreedor")
    
    '--REEMPLAZANDO EL INTERVALO DE FECHA
    N_SQL = Replace(N_SQL, "con_diario1.fchasi >= CDate('01/01/07') and con_diario1.fchasi <= cdate('31/12/07')", " ( con_diario1.fchasi >=CDate('" + Me.TxtFchIni.Valor + "') And con_diario1.fchasi <= CDate('" + Me.TxtFchFin.Valor + "') ) ")
    '--REEMPLAZANDO LA FECHA DE INICIO PARA OBTENER LOS SALDOS
    N_SQL = Replace(N_SQL, "con_diario1.fchasi < CDate('01/01/07')", " ( con_diario1.fchasi < CDate('" + Me.TxtFchIni.Valor + "') ) ")
    
    RST_Busq RstTmp, N_SQL, xCon

    If RstTmp.State = 0 Then GoTo Salir:
    RstTmp.Filter = adFilterNone
    RstTmp.Sort = "cuenta"
    fra_msg.Visible = True
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    ProgressBar1.Max = RstTmp.RecordCount
    Label3.Caption = "Procesando Resumen"
    For A = 1 To RstTmp.RecordCount
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo Salir:
        '-----------------------------------------------
        ProgressBar1.Value = A
        Fg3.Rows = Fg3.Rows + 1
        Fg3.TextMatrix(A + 1, 1) = RstTmp("cuenta") & ""
        
        Fg3.TextMatrix(A + 1, 2) = RstTmp("descri") & ""
        
        Fg3.TextMatrix(A + 1, 3) = Format(NulosN(RstTmp("saldeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 4) = Format(NulosN(RstTmp("salhab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 5) = Format(NulosN(RstTmp("movdeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 6) = Format(NulosN(RstTmp("movhab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 7) = Format(NulosN(RstTmp("maydeb")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 8) = Format(NulosN(RstTmp("mayhab")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 9) = Format(NulosN(RstTmp("deudor")), FORMAT_MONTO)
        Fg3.TextMatrix(A + 1, 10) = Format(NulosN(RstTmp("acreedor")), FORMAT_MONTO)
        
        
'''        If NulosN(RstTmp("totdeb")) <> 0 Then
'''            Fg3.TextMatrix(A + 1, 8) = Format(NulosN(Fg3.TextMatrix(A + 1, 5)) + NulosN(Fg3.TextMatrix(A + 1, 6)), FORMAT_MONTO)
'''            If NulosN(RstTmp("tothab")) <> 0 Then
'''                Fg3.TextMatrix(A + 1, 8) = NulosN(Fg3.TextMatrix(A + 1, 8)) - NulosN(RstTmp("tothab"))
'''                Fg3.TextMatrix(A + 1, 8) = Format(Fg3.TextMatrix(A + 1, 8), FORMAT_MONTO)
'''            End If
'''        Else
'''            Fg3.TextMatrix(A + 1, 8) = Format(NulosN(Fg3.TextMatrix(A + 1, 5)) - NulosN(Fg3.TextMatrix(A + 1, 7)), FORMAT_MONTO)
'''        End If
'''        Fg3.TextMatrix(A + 1, 8) = Format(Fg3.TextMatrix(A + 1, 8), FORMAT_MONTO)
        
'''        xTotal1 = xTotal1 + NulosN(RstTmp("saldeb"))
'''        xTotal2 = xTotal2 + NulosN(RstTmp("salhab"))
'''        xTotal3 = xTotal3 + NulosN(RstTmp("totdeb"))
'''        xTotal4 = xTotal4 + NulosN(RstTmp("tothab"))
        
        xAcumulado(0) = xAcumulado(0) + NulosN(Fg3.TextMatrix(A + 1, 3)) '--saldeb
        xAcumulado(1) = xAcumulado(1) + NulosN(Fg3.TextMatrix(A + 1, 4)) '--salhab
        xAcumulado(2) = xAcumulado(2) + NulosN(Fg3.TextMatrix(A + 1, 5)) '--movdeb
        xAcumulado(3) = xAcumulado(3) + NulosN(Fg3.TextMatrix(A + 1, 6)) '--movhab
        xAcumulado(4) = xAcumulado(4) + NulosN(Fg3.TextMatrix(A + 1, 7)) '--maydeb
        xAcumulado(5) = xAcumulado(5) + NulosN(Fg3.TextMatrix(A + 1, 8)) '--mayhab
        xAcumulado(6) = xAcumulado(6) + NulosN(Fg3.TextMatrix(A + 1, 9)) '--deudor
        xAcumulado(7) = xAcumulado(7) + NulosN(Fg3.TextMatrix(A + 1, 10)) '--acreedor
        
        RstTmp.MoveNext
        If RstTmp.EOF = True Then Exit For
    Next A
    
    Fg3.Rows = Fg3.Rows + 2
    
    Fg3.TextMatrix(Fg3.Rows - 1, 2) = "TOTAL ==>"
    
'''    Fg3.TextMatrix(Fg3.Rows - 1, 3) = Format(xTotal1, FORMAT_MONTO)
'''    Fg3.TextMatrix(Fg3.Rows - 1, 4) = Format(xTotal2, FORMAT_MONTO)
'''    Fg3.TextMatrix(Fg3.Rows - 1, 5) = Format(xTotal1 - xTotal2, FORMAT_MONTO)
'''    Fg3.TextMatrix(Fg3.Rows - 1, 6) = Format(xTotal3, FORMAT_MONTO)
'''    Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(xTotal4, FORMAT_MONTO)
'''    FORMATO_CELDA Fg3, Fg3.Rows - 1, 2, , True
'''    FORMATO_CELDA Fg3, Fg3.Rows - 1, 3, , True
'''    FORMATO_CELDA Fg3, Fg3.Rows - 1, 4, , True
'''    FORMATO_CELDA Fg3, Fg3.Rows - 1, 5, , True
'''    FORMATO_CELDA Fg3, Fg3.Rows - 1, 6, , True
'''    FORMATO_CELDA Fg3, Fg3.Rows - 1, 7, , True
'''    FORMATO_CELDA Fg3, Fg3.Rows - 1, 8, , True
    Dim Col As Integer
    For A = 0 To UBound(xAcumulado())
        Fg3.TextMatrix(Fg3.Rows - 1, 3 + A) = Format(xAcumulado(A), FORMAT_MONTO)
        FORMATO_CELDA Fg3, Fg3.Rows - 1, 3 + A, , True
    Next A
    
    Erase xAcumulado()
    
    GRID_COLOR_FONDO Fg3, 2, 3, Fg3.Rows - 3, 4, RGB(255, 255, 236)
    GRID_COLOR_FONDO Fg3, 2, 7, Fg3.Rows - 3, 8, RGB(255, 255, 236)
    
    GRID_COLOR_FONDO Fg3, Fg3.Rows - 2, 1, Fg3.Rows - 1, Fg3.Cols - 1, RGB(231, 254, 224)
    
    
Salir:
    Set RstRes = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    
    MsgBox "El Mayor se terminó de procesar con éxito", vbInformation
    
    Exit Sub
error:
    Frame5.Visible = False
    Set RstRes = Nothing
    fra_msg.Visible = False
    SHOW_ERROR Me.Name, "CargarResumen"
End Sub

Private Sub Command1_Click()
    If TabOne1.CurrTab = 0 Then
        If Fg1.Rows = 1 Then
            MsgBox "No hay registros para exportar", vbExclamation, xTitulo
            Exit Sub
        End If
    End If
    
    If TabOne1.CurrTab = 1 Then
        If Fg3.Rows = 1 Then
            MsgBox "No hay registros para exportar", vbExclamation, xTitulo
            Exit Sub
        End If
    End If
    
    If TabOne1.CurrTab = 0 Then ExportarExcelDetalle
    If TabOne1.CurrTab = 1 Then EXPORTAR        'ExportarExcelRes
    
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 45 Then
        CmdAdd_Click
    End If
    If KeyCode = 46 Then
        CmdDel_Click
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        TxtFchIni.SetFocus
        SeEjecuto = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    ElseIf KeyCode = vbKeyF3 And Shift = 0 Then
        BUSCAR_VSFlexGrid
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Fg2.Rows = 1
    Fg1.Rows = 1
    Fg3.Rows = 1
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg3.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
    Fg2.Editable = flexEDNone
    Fg3.Editable = flexEDNone
    
    Fg2.Tag = Fg2.FormatString
    
    Fg2.ColWidth(3) = 0
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    Configurar_Grilla True
    
    TabOne1.CurrTab = 0
    opt(0).Value = True
End Sub

Sub PreparaRST_Tmp()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(7, 3) As String

    xCampos(0, 0) = "id":            xCampos(0, 1) = "N":      xCampos(0, 2) = "8"
    xCampos(1, 0) = "cuenta":        xCampos(1, 1) = "C":      xCampos(1, 2) = "15"
    xCampos(2, 0) = "descri":        xCampos(2, 1) = "C":      xCampos(2, 2) = "100"
    xCampos(3, 0) = "totdeb":        xCampos(3, 1) = "D":      xCampos(3, 2) = "8"
    xCampos(4, 0) = "tothab":        xCampos(4, 1) = "D":      xCampos(4, 2) = "8"
    xCampos(5, 0) = "saldeb":        xCampos(5, 1) = "D":      xCampos(5, 2) = "8"
    xCampos(6, 0) = "salhab":        xCampos(6, 1) = "D":      xCampos(6, 2) = "8"
    Set RstTmp = xFun.CrearRstTMP(xCampos)
    RstTmp.Open
End Sub

Sub MuestraMayor()
    Dim RstMay As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstSal As New ADODB.Recordset
    
    Dim A, B, C As Integer
    Dim N_SQL As String
'    On Error GoTo error
''    'llenamos el recorset temporal
'''''    PreparaRST_Tmp2
''

    Frame5.Left = 3413
    Frame5.Top = 2685
    Frame5.Visible = True
    DoEvents
    Dim SQL_CUENTA As String
    SQL_CUENTA = ""
    '--SI AGREGA CUENTAS AS GRID, GENERAR EL FILTRO A CONCATENAR A LA CONSULTA
    For A = 1 To Fg2.Rows - 1
        If Trim(Fg2.TextMatrix(A, 1)) <> "" Then
            SQL_CUENTA = SQL_CUENTA + " con_planctas.cuenta Like '" & Trim(Fg2.TextMatrix(A, 1)) & "%' OR "
        End If
    Next A
    If SQL_CUENTA <> "" Then SQL_CUENTA = " AND (" + Left(SQL_CUENTA, Len(SQL_CUENTA) - 3) + ") "
    '---------
    '--ESTABLECER EL CAMPO A TOTALIZAR EN FUNCION DEL RECORDSET TMP (RstTmp2) , TANTO A SOLES Y DOLARES
    Dim CAMPO_DEBE, CAMPO_HABER, CAMPO_SALDO As String
    If OptSoles.Value = True Then
        CAMPO_DEBE = "impdebsol":  CAMPO_HABER = "imphabsol":   CAMPO_SALDO = "saldosol"
    End If
    If OptDolares.Value = True Then
        CAMPO_DEBE = "impdebdol":  CAMPO_HABER = "imphabdol":   CAMPO_SALDO = "saldodol"
    End If
    '-----------------------------------------------------------------------
    '--SI SE NTERRUMPE EL PROCESO => SALIR
     If BAND_INTERRUMPIR = True Then GoTo Salir:
     '-----------------------------------------------
    Set RstTmp2 = Nothing
    N_SQL = "SELECT con_diario.idcue AS id, con_planctas.id as idcuenta ,con_planctas.cuenta, con_planctas.descripcion AS descri, con_planctas.tipsal, con_diario.idlib, con_diario.idmov, Format(con_diario!idmes,'00')+Trim(con_diario!numasi) AS numreg, mae_libros.descripcion AS nomlib, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.fchdoc,IIf(con_diario.idlib=2,vta_ventas.fchdoc,IIf(con_diario.idlib=3,con_proviciones.fchdoc,IIf(con_diario.idlib=4,con_percepcion.fchdoc,IIf(con_diario.idlib=5,con_retencion.fchemi,IIf(con_diario.idlib=6,con_cajabanco.fchope,IIf(con_diario.idlib=8,con_canjes.fchemi,IIf(con_diario.idlib=9,' ','OTROS LIBROS'))))))))) AS fchemi, " _
        + vbCr + " IIf(con_diario.idlib=0 Or con_diario.idlib Is Null,' ',IIf(con_diario.idlib=1,com_compras.numser & '-' & com_compras.numdoc,IIf(con_diario.idlib=2,vta_ventas.numser & '-' & vta_ventas.numdoc,IIf(con_diario.idlib=3,con_proviciones.numser & '-' & con_proviciones.numdoc,IIf(con_diario.idlib=4,con_percepcion.numser & '-' & con_percepcion.numdoc,IIf(con_diario.idlib=5,con_retencion.numser & '-' & con_retencion.numdoc,IIf(con_diario.idlib=6,con_cajabanco.numdoc,IIf(con_diario.idlib=8,' ',IIf(con_diario.idlib=9,' ','OTROS LIBROS'))))))))) AS numdoc, " _
        + vbCr + " IIf(con_diario.impdebdol<>0,IIf(con_tc.impven Is Null,0,con_diario.impdebdol*con_tc.impven),con_diario.impdebsol) AS impdebsol,IIf(con_diario.imphabdol<>0,IIf(con_tc.impven Is Null,0,con_diario.imphabdol*con_tc.impven),con_diario.imphabsol) AS imphabsol, " _
        + vbCr + " IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol],IIf([con_tc].[impven] Is Null Or [con_diario].[impdebsol]=0,0,([con_diario].[impdebsol]/[con_tc].[impven]))) AS impdebdol, IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol],IIf([con_tc].[impven] Is Null Or [con_diario].[imphabsol]=0,0,([con_diario].[imphabsol]/[con_tc].[impven]))) AS imphabdol, " _
        + vbCr + " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebsol-imphabsol),(imphabsol-impdebsol))) AS saldosol, " _
        + vbCr + " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebdol-imphabdol),(imphabdol-impdebdol))) AS saldodol " _
        + vbCr + " FROM ((((mae_libros RIGHT JOIN (con_planctas RIGHT JOIN ((((con_diario LEFT JOIN con_tc ON con_diario.fchdoc=con_tc.fecha) LEFT JOIN com_compras ON con_diario.idmov=com_compras.id) LEFT JOIN vta_ventas ON con_diario.idmov=vta_ventas.id) LEFT JOIN con_retencion ON con_diario.idmov=con_retencion.id) ON con_planctas.id=con_diario.idcue) ON mae_libros.id=con_diario.idlib) LEFT JOIN con_percepcion ON con_diario.idmov=con_percepcion.id) LEFT JOIN con_proviciones ON con_diario.idmov=con_proviciones.id) LEFT JOIN con_cajabanco ON con_diario.idmov=con_cajabanco.id) LEFT JOIN con_canjes ON con_diario.idmov=con_canjes.id " _
        + vbCr + " WHERE (con_diario.fchasi >=CDate('" + TxtFchIni.Valor + "') And con_diario.fchasi<=CDate('" + TxtFchFin.Valor + "')) " _
        + vbCr + " AND ( con_diario.fchasi >=CDate('01/01/" + AnoTra + "') And con_diario.fchasi <= CDate('31/12/" + AnoTra + "') ) " _
        + SQL_CUENTA _
        + vbCr + " ORDER BY con_planctas.cuenta ASC"


    RST_Busq RstTmp2, N_SQL, xCon

    'HACEMOS UNA CONSULTA DE LOS REGISTROS UNICOS DE LA CONSULTA ANTERIOR, PARA PODER TOTALIZARLA

    Fg1.Rows = 1
    Dim xFila As Integer
    Dim xSaldo As Double
    Dim xTotal1, xTotal2 As Double
    xFila = 1

    DoEvents
    '--SI SE NTERRUMPE EL PROCESO => SALIR
     If BAND_INTERRUMPIR = True Then GoTo Salir:
    '-----------------------------------------------
    N_SQL = "SELECT con_diario.idcue as idcuenta, con_planctas.cuenta, con_planctas.descripcion, Sum(con_diario.impdebsol) AS SumaDeimpdebsol, " _
         + vbCr + " Sum(con_diario.imphabsol) AS SumaDeimphabsol, Sum(con_diario.impdebdol) AS SumaDeimpdebdol, Sum(con_diario.imphabdol) AS SumaDeimphabdol " _
         + vbCr + " FROM con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue " _
         + vbCr + " WHERE (con_diario.fchasi >=CDate('" & TxtFchIni.Valor & "') And con_diario.fchasi <=CDate('" & TxtFchFin.Valor & "')) " _
         + vbCr + " AND ( con_diario.fchasi >=CDate('01/01/" + AnoTra + "') And con_diario.fchasi <= CDate('31/12/" + AnoTra + "') ) " _
         + vbCr + SQL_CUENTA _
         + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion ORDER BY con_planctas.cuenta ASC "
    
    RST_Busq RstMay, N_SQL, xCon

    xSaldo = 0
    '---
    If RstMay.RecordCount <> 0 Then
    
        fra_msg.Visible = True
        
        RstMay.MoveFirst
        
        ProgressBar1.Max = RstMay.RecordCount
        
        'Label3.Caption = "Procesando Cta Nº  :  " + Trim(RstMay("cuenta") & "") + " - " + RstMay.Fields("descripcion") & ""
        DoEvents
        For A = 1 To RstMay.RecordCount
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo Salir:
            '-----------------------------------------------
            ProgressBar1.Value = A
    
            xSaldo = 0
            Fg1.Rows = Fg1.Rows + 1
            
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 2, Fg1.Rows - 1, 7, "Cta Nº  :  " + RstMay("cuenta") & "   - " + RstMay.Fields("descripcion") & "", flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, , True
            
            'hallamos el saldo anterior de la cuenta
            Set RstSal = Nothing
            
            N_SQL = "SELECT con_diario.idcue as idcuenta, con_planctas.cuenta,con_planctas.tipsal, " _
                + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[impdebdol]*[con_tc].[impven]),[con_diario].[impdebsol])) AS impdebsol, " _
                + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[imphabdol]*[con_tc].[impven]),[con_diario].[imphabsol])) AS imphabsol, " _
                + vbCr + " Sum(IIf([con_diario]![impdebdol]<>0,[con_diario]![impdebdol],IIf([con_diario]![impdebsol]=0 Or [con_tc].[impven] Is Null,0,[con_diario]![impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
                + vbCr + " Sum(IIf([con_diario]![imphabdol]<>0,[con_diario]![imphabdol], IIf([con_diario]![imphabsol] = 0 Or [con_diario]![imphabsol] Is Null Or [con_tc].[impven] Is Null, 0, [con_diario]![imphabsol] / [con_tc].[impven]))) As imphabdol, " _
                + vbCr + " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebsol-imphabsol),(imphabsol-impdebsol))) AS saldosol, " _
                + vbCr + " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebdol-imphabdol),(imphabdol-impdebdol))) AS saldodol " _
                + vbCr + " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
                + vbCr + " WHERE (((con_diario.fchasi) Is Null Or (con_diario.fchasi) < CDate('" & TxtFchIni.Valor & "'))) " _
                + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.tipsal, con_tc.idmon,(IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebsol-imphabsol),(imphabsol-impdebsol))), (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebdol-imphabdol),(imphabdol-impdebdol))) " _
                + vbCr + " HAVING con_planctas.cuenta ='" & RstMay("cuenta") & "' "
                
            RST_Busq RstSal, N_SQL, xCon
            
            If RstSal.RecordCount <> 0 Then
'                xSaldo = (RstSal(CAMPO_DEBE) - RstSal(CAMPO_HABER))
'                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xSaldo, FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstSal.Fields(CAMPO_SALDO) & ""), FORMAT_MONTO)
                xSaldo = NulosN(RstSal.Fields(CAMPO_SALDO))
            Else
                xSaldo = 0
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = "0.00"
            End If
            
            
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, , True
            
            RstTmp2.Filter = adFilterNone
            RstTmp2.Filter = "idcuenta = '" & RstMay("idcuenta") & "'"
            Label3.Caption = "Procesando Cta Nº  :  " + Trim(RstMay("cuenta") & "") '+ " - " + RstMay.Fields("descripcion") & ""
            DoEvents
            xFila = xFila + 1
            If RstTmp2.RecordCount <> 0 Then
            
                If opt(0).Value = True Then
                    RstTmp2.Sort = "fchemi"
                ElseIf opt(1).Value = True Then
                    RstTmp2.Sort = "numdoc"
                Else
                    RstTmp2.Sort = "numreg"
                End If
                
                For B = 1 To RstTmp2.RecordCount
                    DoEvents
                    '--SI SE NTERRUMPE EL PROCESO => SALIR
                    If BAND_INTERRUMPIR = True Then GoTo Salir
                    '-----------------------------------------------
                    Fg1.Rows = Fg1.Rows + 1
                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = RstTmp2("numreg") & ""
                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = RstTmp2("nomlib") & ""
                    If IsDate(RstTmp2("fchemi")) = True Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(RstTmp2("fchemi"), FORMAT_DATE)
                    End If
                    Fg1.TextMatrix(Fg1.Rows - 1, 5) = RstTmp2("numdoc") & ""
                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(RstTmp2(CAMPO_DEBE), FORMAT_MONTO)
                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(RstTmp2(CAMPO_HABER), FORMAT_MONTO)
                    
                    If UCase(RstTmp2.Fields("tipsal") & "") = "D" Or RstTmp2.Fields("tipsal") = "" Then
                        xSaldo = xSaldo + (RstTmp2(CAMPO_DEBE) - RstTmp2(CAMPO_HABER))
                    Else
                        xSaldo = xSaldo + (RstTmp2(CAMPO_HABER) - RstTmp2(CAMPO_DEBE))
                    End If
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(xSaldo, FORMAT_MONTO)
                    
                    xTotal1 = xTotal1 + RstTmp2(CAMPO_DEBE)
                    xTotal2 = xTotal2 + RstTmp2(CAMPO_HABER)
                    RstTmp2.MoveNext
                    If RstTmp2.EOF = True Then
                        Fg1.Rows = Fg1.Rows + 1
                        Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(xTotal1, FORMAT_MONTO)
                        Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(xTotal2, FORMAT_MONTO)
                        
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, , True
                        FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, , True
                        
                        Fg1.Rows = Fg1.Rows + 1
                        Exit For
                    End If
                    xFila = xFila + 1
                Next B
            End If
            
            RstMay.MoveNext
            If RstMay.EOF = True Then Exit For
        Next A
    Else
        xSaldo = 0

        'hallamos el saldo anterior de la cuenta

        N_SQL = "SELECT con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion ,con_planctas.tipsal, " _
            + vbCr + " Sum(IIf([con_diario].[impdebdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[impdebdol]*[con_tc].[impven]),[con_diario].[impdebsol])) AS impdebsol, " _
            + vbCr + " Sum(IIf([con_diario].[imphabdol]<>0,IIf([con_tc].[impven] Is Null,0,[con_diario].[imphabdol]*[con_tc].[impven]),[con_diario].[imphabsol])) AS imphabsol, " _
            + vbCr + " Sum(IIf([con_diario]![impdebdol]<>0,[con_diario]![impdebdol],IIf([con_diario]![impdebsol]=0 Or [con_tc].[impven] Is Null,0,[con_diario]![impdebsol]/[con_tc].[impven]))) AS impdebdol, " _
            + vbCr + " Sum(IIf([con_diario]![imphabdol]<>0,[con_diario]![imphabdol], IIf([con_diario]![imphabsol] = 0 Or [con_diario]![imphabsol] Is Null Or [con_tc].[impven] Is Null, 0, [con_diario]![imphabsol] / [con_tc].[impven]))) As imphabdol, " _
            + vbCr + " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebsol-imphabsol),(imphabsol-impdebsol))) AS saldosol, " _
            + vbCr + " (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebdol-imphabdol),(imphabdol-impdebdol))) AS saldodol " _
            + vbCr + " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            + vbCr + " WHERE (((con_diario.fchasi) Is Null Or (con_diario.fchasi) < CDate('" & TxtFchIni.Valor & "'))) " + SQL_CUENTA _
            + vbCr + " GROUP BY con_diario.idcue, con_planctas.cuenta, con_planctas.descripcion,con_planctas.tipsal, con_tc.idmon, " _
            + vbCr + "  (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebsol-imphabsol),(imphabsol-impdebsol))), (IIf(con_planctas.tipsal='D' Or con_planctas.tipsal Is Null Or con_planctas.tipsal='',(impdebdol-imphabdol),(imphabdol-impdebdol))) "
            
        RST_Busq RstSal, N_SQL, xCon
        
        If RstSal.EOF = False Or RstSal.BOF = False Or RstSal.RecordCount <> 0 Then
            Fg1.Rows = Fg1.Rows + 1
            RstSal.MoveFirst
        End If
        
        Do While Not RstSal.EOF
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 2, Fg1.Rows - 1, 5, "Cta Nº  :  " + RstSal("cuenta") & "   - " + RstSal.Fields("descripcion") & "", flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, , True
            
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(RstSal(CAMPO_SALDO)), FORMAT_MONTO)
            Fg1.Rows = Fg1.Rows + 1
            RstSal.MoveNext
        Loop
        Set RstSal = Nothing
        '----------------------------
        
    End If

Salir:
    Set RstMay = Nothing:     Set RstDet = Nothing:     Set RstSal = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    Exit Sub
error:
    Set RstMay = Nothing:     Set RstDet = Nothing:     Set RstSal = Nothing
    Frame5.Visible = False
    fra_msg.Visible = False
    SHOW_ERROR Me.Name, "MuestraMayor"
End Sub

Sub PreparaRST_Tmp2()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(13, 3) As String

    xCampos(0, 0) = "id":            xCampos(0, 1) = "N":      xCampos(0, 2) = "8"
    xCampos(1, 0) = "cuenta":        xCampos(1, 1) = "C":      xCampos(1, 2) = "15"
    xCampos(2, 0) = "descri":        xCampos(2, 1) = "C":      xCampos(2, 2) = "100"
    xCampos(3, 0) = "impdebsol":     xCampos(3, 1) = "D":      xCampos(3, 2) = "8"
    xCampos(4, 0) = "imphabsol":     xCampos(4, 1) = "D":      xCampos(4, 2) = "8"
    xCampos(5, 0) = "impdebdol":     xCampos(5, 1) = "D":      xCampos(5, 2) = "8"
    xCampos(6, 0) = "imphabdol":     xCampos(6, 1) = "D":      xCampos(6, 2) = "8"
    xCampos(7, 0) = "idlib":         xCampos(7, 1) = "N":      xCampos(7, 2) = "8"
    xCampos(8, 0) = "idmov":         xCampos(8, 1) = "N":      xCampos(8, 2) = "8"
    xCampos(9, 0) = "fchemi":        xCampos(9, 1) = "F":      xCampos(9, 2) = "8"
    xCampos(10, 0) = "numdoc":       xCampos(10, 1) = "C":      xCampos(10, 2) = "20"
    xCampos(11, 0) = "numreg":       xCampos(11, 1) = "C":      xCampos(11, 2) = "20"
    xCampos(12, 0) = "nomlib":       xCampos(12, 1) = "C":      xCampos(12, 2) = "20"
    Set RstTmp2 = xFun.CrearRstTMP(xCampos)
    RstTmp2.Open
End Sub

Sub ExportarExcelDetalle()
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    
    On Error GoTo error

    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.Workbooks.Add  'Trim(App.Path) + "\RegCompras.xls"
    
    objExcel.WindowState = 1
    
    With objExcel.ActiveSheet
        
        .Cells(1, 2) = NomEmp
        .Cells(1, 11) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        

        If Fg1.ColWidth(1) <> 0 Then .Columns(2).ColumnWidth = Fg1.ColWidth(1) / 100
        If Fg1.ColWidth(2) <> 0 Then .Columns(3).ColumnWidth = Fg1.ColWidth(2) / 100
        If Fg1.ColWidth(3) <> 0 Then .Columns(4).ColumnWidth = Fg1.ColWidth(3) / 100
        If Fg1.ColWidth(4) <> 0 Then .Columns(5).ColumnWidth = Fg1.ColWidth(4) / 100
        If Fg1.ColWidth(5) <> 0 Then .Columns(6).ColumnWidth = Fg1.ColWidth(5) / 100
        If Fg1.ColWidth(6) <> 0 Then .Columns(7).ColumnWidth = Fg1.ColWidth(6) / 100
        If Fg1.ColWidth(7) <> 0 Then .Columns(8).ColumnWidth = Fg1.ColWidth(7) / 100

        
        xFilas = 7
        For B = 1 To Fg1.Cols - 1
            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(0, B)
        Next B
        
        .Range("B6:K7").Font.Bold = True
    
        xFilas = xFilas + 1
        For A = 1 To Fg1.Rows - 1
            DoEvents
            For B = 1 To Fg1.Cols - 1
                
                If B <= 5 Then
                    If B = 2 Then
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                        If InStr(Fg1.TextMatrix(A, B), "Cta Nº  :") <> 0 Then
                            .Cells(xFilas, 2) = "'" + Fg1.TextMatrix(A, B)
                            GoTo SIG_FIL
                        End If
                    Else
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                    End If
                    
                    
                Else
                    If IsNumeric(Fg1.TextMatrix(A, B)) = True Then
                        .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                    End If
                End If

            Next B
SIG_FIL:
            xFilas = xFilas + 1
        Next A
    End With
    
    MsgBox "El proceso de exportación terminó con éxito", vbInformation, xTitulo
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
error:
    SHOW_ERROR Me.Name, "ExportarExcelDetalle", , IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
    
End Sub


Sub ExportarExcelRes()
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    On Error GoTo error
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add  'Trim(App.Path) + "\RegCompras.xls"
    
    Frame4.Left = 2940
    Frame4.Top = 1890
    Label3.Caption = "Exportando Documentos"
    Frame4.Visible = True
    
    
    ProgressBar1.Max = Fg3.Rows - 1
    
    With objExcel.ActiveSheet
        
        .Cells(1, 2) = NomEmp
        .Cells(1, 11) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        .Cells(6, 2) = "'" + Fg3.TextMatrix(0, 1)
        .Cells(6, 4) = "'" + "Saldos Iniciales" + " Al " + CStr(CDate(TxtFchIni.Valor) - 1)
        .Cells(6, 6) = "'" + Fg3.TextMatrix(0, 5)
        .Cells(6, 8) = "'" + Fg3.TextMatrix(0, 7)
        .Cells(6, 10) = "'" + "Saldos Finales" + " Al " + CStr(CDate(TxtFchFin.Valor) - 1)
        
        If Fg3.ColWidth(1) <> 0 Then .Columns(2).ColumnWidth = Fg3.ColWidth(1) / 100
        If Fg3.ColWidth(2) <> 0 Then .Columns(3).ColumnWidth = Fg3.ColWidth(2) / 100
        If Fg3.ColWidth(3) <> 0 Then .Columns(4).ColumnWidth = Fg3.ColWidth(3) / 100
        If Fg3.ColWidth(4) <> 0 Then .Columns(5).ColumnWidth = Fg3.ColWidth(4) / 100
        If Fg3.ColWidth(5) <> 0 Then .Columns(6).ColumnWidth = Fg3.ColWidth(5) / 100
        If Fg3.ColWidth(6) <> 0 Then .Columns(7).ColumnWidth = Fg3.ColWidth(6) / 100
        If Fg3.ColWidth(7) <> 0 Then .Columns(8).ColumnWidth = Fg3.ColWidth(7) / 100
        If Fg3.ColWidth(8) <> 0 Then .Columns(9).ColumnWidth = Fg3.ColWidth(8) / 100
        If Fg3.ColWidth(9) <> 0 Then .Columns(10).ColumnWidth = Fg3.ColWidth(9) / 100
        If Fg3.ColWidth(10) <> 0 Then .Columns(11).ColumnWidth = Fg3.ColWidth(10) / 100
        
        
        xFilas = 7
        For B = 1 To Fg3.Cols - 1
            .Cells(xFilas, B + 1) = "'" + Fg3.TextMatrix(1, B)
        Next B
        .Range("B6:K7").Font.Bold = True
        
        xFilas = xFilas + 1
        For A = 2 To Fg3.Rows - 1
            ProgressBar1.Value = A
            If Fg3.TextMatrix(A, 1) <> CStr(Fila_en_Blanco) Then
                For B = 1 To Fg3.Cols - 1
                    If B <= 2 Then
                        .Cells(xFilas, B + 1) = "'" + Fg3.TextMatrix(A, B)
                    Else
                        If IsNumeric(Fg3.TextMatrix(A, B)) = True Then
                            .Cells(xFilas, B + 1) = NulosN(Fg3.TextMatrix(A, B))
                        End If
                    End If
                Next B
            End If
            xFilas = xFilas + 1
        Next A
        
    End With
    
    Frame4.Visible = False
    MsgBox "El proceso de exportación terminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "ExportExcelRes", , IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
End Sub





Private Sub Configurar_Grilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
    Dim M_ANCHO_COL As Integer '--DEPENDERA DEL TIPO DE CONSULTA
                                   
    Dim k, j As Integer
    Dim T_CONSULTA As Integer
    
        
    With Fg1
        '-----
        If F_CONSERVAR_FORMATO = True Then LimpiarGrid Fg1, , 1
        .FrozenCols = 0
        Fg1.Cols = 9
                 
        .ColWidth(0) = 200
        '--DATOS DE FILA
        .TextMatrix(0, 1) = "Descripción":  .ColWidth(1) = 0:       .ColAlignment(1) = flexAlignLeftBottom
        .TextMatrix(0, 2) = "Num.Reg.":     .ColWidth(2) = 800:     .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Libro":        .ColWidth(3) = 1200:    .ColAlignment(3) = flexAlignLeftBottom
        .TextMatrix(0, 4) = "Fch. Doc":     .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignCenterBottom
        .TextMatrix(0, 5) = "Nº Documento": .ColWidth(5) = 1500:    .ColAlignment(5) = flexAlignLeftBottom
        .TextMatrix(0, 6) = "Debe":         .ColWidth(6) = 1300:    .ColAlignment(6) = flexAlignRightBottom
        .TextMatrix(0, 7) = "Haber":        .ColWidth(7) = 1300:    .ColAlignment(7) = flexAlignRightBottom
        .TextMatrix(0, 8) = "Saldo":        .ColWidth(8) = 1300:    .ColAlignment(8) = flexAlignRightBottom
                
    End With
    
    With Fg3
        '-----
        If F_CONSERVAR_FORMATO = True Then LimpiarGrid Fg3, , 2
        
        .Cols = 11
        .FixedRows = 2
        .FrozenCols = 2
        .RowHeight(0) = 500
        
        UNIR_CELDAS Fg3, 0, 1, 0, 2, "Datos de la Cuenta", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 3, 0, 4, "Saldos Iniciales", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 5, 0, 6, "Movimiento del Periodo", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 7, 0, 8, "Sumas del Mayor", flexAlignCenterCenter
        UNIR_CELDAS Fg3, 0, 9, 0, 10, "Saldos Finales", flexAlignCenterCenter
        
'        .ColWidth(0) = 200
'        '--DATOS DE FILA
'
        .TextMatrix(1, 1) = "Nº. Cuenta":       .ColWidth(1) = 1100:       .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(1, 2) = "Descripción":      .ColWidth(2) = 3000:       .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(1, 3) = "Debe":       .ColWidth(3) = 1300:       .ColAlignment(3) = flexAlignRightCenter
        .TextMatrix(1, 4) = "Haber":      .ColWidth(4) = 1300:       .ColAlignment(4) = flexAlignRightCenter
        .TextMatrix(1, 5) = "Debe":       .ColWidth(5) = 1320:       .ColAlignment(5) = flexAlignRightCenter
        .TextMatrix(1, 6) = "Haber":      .ColWidth(6) = 1320:       .ColAlignment(6) = flexAlignRightCenter
        .TextMatrix(1, 7) = "Debe":       .ColWidth(7) = 1320:       .ColAlignment(7) = flexAlignRightCenter
        .TextMatrix(1, 8) = "Haber":      .ColWidth(8) = 1320:       .ColAlignment(8) = flexAlignRightCenter
        .TextMatrix(1, 9) = "Debe":       .ColWidth(9) = 1200:       .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(1, 10) = "Haber":     .ColWidth(10) = 1200:      .ColAlignment(10) = flexAlignRightCenter
        
        '--AGREGANDO LAS FECHAS EN LA CABECERA
        If IsDate(TxtFchIni.Valor) = True Then UNIR_CELDAS Fg3, 0, 3, 0, 4, "Saldos Iniciales" + vbCr + " Al " + CStr(CDate(TxtFchIni.Valor) - 1), flexAlignCenterCenter
        If IsDate(TxtFchFin.Valor) = True Then UNIR_CELDAS Fg3, 0, 9, 0, 10, "Saldos Finales" + vbCr + " Al " + CStr(CDate(TxtFchFin.Valor) - 1), flexAlignCenterCenter
        
    End With
    
    DoEvents
End Sub


Private Function Validar_Consulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Function
    End If
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Function
    End If
    
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If
    
    If (Year(TxtFchIni.Valor) <> Year(TxtFchFin.Valor)) Then
        MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    ElseIf Year(TxtFchIni.Valor) <> CStr(AnoTra) Then
        MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If
    
    If chk.Value = 0 Then
        If Fg2.Rows = 1 Then
            MsgBox "No ha especificado una cuenta contable a mayorizar" + vbCr + "Si desea ver todas las cuentas, Active la opción: Procesar Todas las Cuentas...", vbExclamation, xTitulo
            CmdAdd.SetFocus
            Exit Function
        End If
    Else
        If MsgBox("Seguro desea procesar todas las cuentas", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function
    End If
    
    Validar_Consulta = True
End Function



Private Sub EXPORTAR()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim T_PERIODO As String
    Dim T_TITULO1 As String
    If MsgBox("Desea conservar el formato de la consulta", vbQuestion + vbYesNo, "Exportar...") = vbNo Then Configurar_Grilla False
    If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
        T_PERIODO = " Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
    Else
        T_PERIODO = "Al: " + CStr(TxtFchIni.Valor)
    End If
    If Me.OptSoles.Value = True Then
        T_TITULO1 = "(Expresado en Nuevos Soles)"
    Else
        T_TITULO1 = "(Expresado en Dolares Americanos)"
    End If
    
    
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg3, "RESUMEN DEL MAYOR", T_PERIODO, T_TITULO1, "Resumen del Mayor"
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub


Private Sub BUSCAR_VSFlexGrid()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    
    Dim xCampos(6, 3) As String
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "Num.Reg.":     xCampos(0, 1) = "2":    xCampos(0, 2) = "C":    xCampos(0, 3) = "0"
    xCampos(1, 0) = "Libro":        xCampos(1, 1) = "3":    xCampos(1, 2) = "C":    xCampos(1, 3) = "0"
    xCampos(2, 0) = "Fch. Doc   ":  xCampos(2, 1) = "4":    xCampos(2, 2) = "F":    xCampos(2, 3) = "0"
    xCampos(3, 0) = "Nº Documento": xCampos(3, 1) = "5":    xCampos(3, 2) = "C":    xCampos(3, 3) = "-1"
    xCampos(4, 0) = "Debe":         xCampos(4, 1) = "6":    xCampos(4, 2) = "N":    xCampos(4, 3) = "0"
    xCampos(5, 0) = "Haber":        xCampos(5, 1) = "7":    xCampos(5, 2) = "N":    xCampos(5, 3) = "0"
    xCampos(6, 0) = "Saldo":        xCampos(6, 1) = "8":    xCampos(6, 2) = "N":    xCampos(6, 3) = "0"
    
    X_EXPORT.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos(), Fg1.Row
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "BUSCAR_VSFlexGrid"
End Sub


