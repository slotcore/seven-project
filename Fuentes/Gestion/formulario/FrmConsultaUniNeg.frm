VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.CommandBars.v12.0.0.ocx"
Begin VB.Form FrmConsultaUniNeg 
   Caption         =   "Gestion - Analisis por Unidad de Negocio"
   ClientHeight    =   8160
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.ElasticOne EO 
      Height          =   7275
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   810
      Width           =   11835
      _cx             =   20876
      _cy             =   12832
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
      Align           =   0
      AutoSizeChildren=   8
      BorderWidth     =   2
      ChildSpacing    =   2
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
      ResizeFonts     =   0   'False
      GridRows        =   2
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmConsultaUniNeg.frx":0000
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   5550
         Left            =   30
         TabIndex        =   1
         Top             =   1695
         Width           =   11775
         _cx             =   20770
         _cy             =   9790
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
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsultaUniNeg.frx":0045
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
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   1635
         Left            =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   11775
         _cx             =   20770
         _cy             =   2884
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
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   2
         ChildSpacing    =   2
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
         ResizeFonts     =   0   'False
         GridRows        =   1
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmConsultaUniNeg.frx":011A
         Begin VB.Frame Frame1 
            Caption         =   "[ Opciones ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   30
            TabIndex        =   6
            Top             =   30
            Width           =   2895
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Height          =   570
               Left            =   1530
               TabIndex        =   13
               Top             =   960
               Width           =   1305
               Begin VB.OptionButton OptDol 
                  Caption         =   "Dolares"
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
                  Left            =   0
                  TabIndex        =   15
                  Top             =   300
                  Width           =   1125
               End
               Begin VB.OptionButton OptSoles 
                  Caption         =   "Soles"
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
                  Left            =   0
                  TabIndex        =   14
                  Top             =   30
                  Width           =   1125
               End
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Resumido"
               Height          =   195
               Left            =   135
               TabIndex        =   8
               Top             =   990
               Width           =   1320
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Detallado"
               Height          =   195
               Left            =   135
               TabIndex        =   7
               Top             =   1260
               Width           =   1320
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   1155
               TabIndex        =   9
               Top             =   285
               Width           =   1200
               _ExtentX        =   2117
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
               Left            =   1155
               TabIndex        =   10
               Top             =   600
               Width           =   1200
               _ExtentX        =   2117
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
               Caption         =   "Fch. Inicio"
               Height          =   195
               Left            =   135
               TabIndex        =   12
               Top             =   315
               Width           =   735
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Final"
               Height          =   195
               Left            =   135
               TabIndex        =   11
               Top             =   645
               Width           =   690
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "[                              ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   2955
            TabIndex        =   4
            Top             =   30
            Width           =   4845
            Begin VB.CheckBox Check1 
               Caption         =   "Aplicar Clientes"
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
               Height          =   195
               Left            =   285
               TabIndex        =   17
               Top             =   15
               Width           =   1665
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   1230
               Left            =   75
               TabIndex        =   5
               Top             =   285
               Width           =   4605
               _cx             =   8123
               _cy             =   2170
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmConsultaUniNeg.frx":0176
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
            Caption         =   "[ Agrupar Por ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   7830
            TabIndex        =   3
            Top             =   30
            Visible         =   0   'False
            Width           =   3555
            Begin VB.OptionButton Option4 
               Caption         =   "Cliente"
               Height          =   255
               Left            =   180
               TabIndex        =   19
               Top             =   540
               Width           =   2115
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Unidad de Negocio"
               Height          =   255
               Left            =   180
               TabIndex        =   18
               Top             =   270
               Width           =   2115
            End
         End
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4410
      TabIndex        =   16
      Top             =   330
      Width           =   1485
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   1860
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   2280
      Top             =   30
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "FrmConsultaUniNeg.frx":01C6
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu_1 
         Caption         =   "Agregar Cliente"
      End
      Begin VB.Menu menu_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_3 
         Caption         =   "Eliminar Cliente"
      End
   End
End
Attribute VB_Name = "FrmConsultaUniNeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SeEjecuto As Boolean
Dim TopEO As Integer

Dim TAMAÑO_TOOL As TOOL_TAMAÑO_ICO

Const BTN_BUS = 1
Const BTN_EXP = 2
Const BTN_IMP = 3
Const BTN_CON = 4
Const BTN_SAL = 5
Dim xCadWhere As String
Dim xCadWhere2 As String
Dim xNomUnidad As String
Dim xTotC, xTotV As Double
Dim GTotC, GTotV As Double

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Option3.Value = True
        'Option3_Click
        Frame3.Visible = True
        Fg2.Editable = flexEDKbdMouse
        'If Option1.Value = True Then SetearCuadricula Fg1, 8, xCon, 2, 3, False
        'If Option2.Value = True Then SetearCuadricula Fg1, 8, xCon, 2, 4, False
        OptSoles.Visible = False
        OptDol.Visible = False
    Else
        Option3.Value = False
        Option4.Value = False
        
        Frame3.Visible = False
        Fg2.Rows = 0
        Fg2.Rows = Fg2.Rows + 1
        Fg2.Editable = flexEDNone
        If Option1.Value = True Then SetearCuadricula Fg1, 8, xCon, 2, 2, False
        If Option2.Value = True Then SetearCuadricula Fg1, 8, xCon, 2, 1, False
        OptSoles.Visible = True
        OptDol.Visible = True
    End If
End Sub

Private Sub Check2_Click()
'    If Check2.Value = 1 Then
'        Frame4.Visible = False
'    Else
'        Frame4.Visible = True
'    End If
'    Option1_Click
End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.ID = 1 Then Procesar
    
    If Control.ID = 2 Then pExportar
    If Control.ID = 5 Then
        Unload Me
    End If
End Sub

Private Sub pExportar()
    Dim xFun As New SGI2_funciones.Formularios
    Dim Rst As New ADODB.Recordset
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "Analisis x Documento de Referencia - DETALLADO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en : Ambas Monedas"
    
    Set xFun = Nothing
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":           xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":            xCampos(1, 2) = "800":          xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Documento":  xCampos(2, 1) = "numruc":           xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Tipo Empresa":  xCampos(3, 1) = "tipemp":           xCampos(3, 2) = "1500":         xCampos(3, 3) = "C"
    
    xform.SQLCad = "SELECT mae_cliente.nombre, mae_dociden.abrev, mae_tipoempresa.descripcion AS tipemp, mae_cliente.numruc, " _
        & " mae_cliente.id FROM (mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) LEFT JOIN mae_tipoempresa " _
        & " ON mae_cliente.tipper = mae_tipoempresa.id Where (((mae_cliente.activo) = -1)) ORDER BY mae_cliente.nombre"
    
    xform.Titulo = "Buscando Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = xRs("nombre")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = xRs("id")
            
            If Fg2.TextMatrix(Fg2.Rows - 1, 1) <> "" Then
                Fg2.Rows = Fg2.Rows + 1
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If NulosN(Check1.Value) = 0 Then Exit Sub
    If Button = 2 Then
        PopupMenu Menu
    End If
End Sub

Private Sub Form_Activate()
If SeEjecuto = False Then
    SeEjecuto = True
    TxtFchIni.SetFocus
End If
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    SeEjecuto = False
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
        
    SetearCuadricula Fg1, 8, xCon, 2, 2, False
    Fg1.BackColor = &HE2FEFB
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H80&
    Fg1.Editable = flexEDNone
    Frame4.BackColor = &H8000000F
    Fg2.BackColor = &HE2FEFB
    Fg2.Rows = 0
    Fg2.Rows = Fg2.Rows + 1
    Fg2.ColComboList(1) = "|..."
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.ColWidth(2) = 0
    'Fg2.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDNone
    
    Option1.Value = True
    OptSoles.Value = True
    
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    CrearTool
    
    
    
    If TAMAÑO_TOOL = I16x16 Then EO.Top = 400: TopEO = 400
    If TAMAÑO_TOOL = I24x24 Then EO.Top = 520: TopEO = 520
    If TAMAÑO_TOOL = I32x32 Then EO.Top = 640: TopEO = 640
    If TAMAÑO_TOOL = I48x48 Then EO.Top = 890: TopEO = 890

End Sub

Sub CrearTool()
    'CREAMOS EL TOOLBAR
    Dim Opciones(4, 3) As String
    
    Opciones(0, 0) = Str(BTN_BUS):    Opciones(0, 1) = "Buscar":                      Opciones(0, 2) = "0":      Opciones(0, 3) = "Ejecutar Busqueda"
    Opciones(1, 0) = Str(BTN_EXP):    Opciones(1, 1) = "Exportar Excel":              Opciones(1, 2) = "0":      Opciones(1, 3) = "Exportar Excel"
    Opciones(2, 0) = Str(BTN_IMP):    Opciones(2, 1) = "Imprimir":                    Opciones(2, 2) = "0":      Opciones(2, 3) = "Imprimir"
    Opciones(3, 0) = Str(BTN_CON):    Opciones(3, 1) = "Configurar":                  Opciones(3, 2) = "0":      Opciones(3, 3) = "Configurar"
    Opciones(4, 0) = Str(BTN_SAL):    Opciones(4, 1) = "Salir":                       Opciones(4, 2) = "1":      Opciones(4, 3) = "Salir"
        
    Dim xFun As New eps_librerias.Codejock
    'PocisionarContenedor
    xFun.BORRARMENU = True
    TAMAÑO_TOOL = I24x24
    xFun.CrearToolBar Opciones, CommandBars1, ImageManager1, TAMAÑO_TOOL
    Set xFun = Nothing
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    EO.Width = Me.Width - 130
    If Me.Height <= (TopEO + 2375) Then
        Me.Height = (TopEO + 2375)
    Else
        EO.Height = (Me.Height - (TopEO + 400))
    End If
    
    Me.Refresh
End Sub

Sub Procesar()
    Dim A As Integer
    'Dim xCadWhere As String
    
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "Rango de fechas ingresado no valido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    ' ELIMINAMOS LAS FILAS EN BLANCO
    If Fg2.Rows <> 0 Then
        For A = 0 To Fg2.Rows - 1
            If NulosN(Fg2.TextMatrix(A, 2)) = 0 Then
                Fg2.RemoveItem A
            End If
        Next A
    End If
    
    If Fg2.Rows <> 0 Then
        xCadWhere = ""
        ' MOSTRAMOS SOLOS LOS CLIENTES ESPECIFICADOS

        xCadWhere = "HAVING ("
        For A = 0 To Fg2.Rows - 1
        '   HAVING (((vta_ventas.idcli)=1 Or (vta_ventas.idcli)=2))
            xCadWhere = xCadWhere & "(vta_ventas.idcli = " & Fg2.TextMatrix(A, 2) & ")"
            If A = Fg2.Rows - 1 Then
                Exit For
            End If
            xCadWhere = xCadWhere & " OR "
        Next A
        xCadWhere = xCadWhere & ")"
        
        '-------------------------------------------
        xCadWhere2 = ""
        ' MOSTRAMOS SOLOS LOS CLIENTES ESPECIFICADOS

        xCadWhere2 = "("
        For A = 0 To Fg2.Rows - 1
        '   HAVING (((vta_ventas.idcli)=1 Or (vta_ventas.idcli)=2))
            xCadWhere2 = xCadWhere2 & "(vta_ventas.idcli = " & Fg2.TextMatrix(A, 2) & ")"
            If A = Fg2.Rows - 1 Then
                Exit For
            End If
            xCadWhere2 = xCadWhere2 & " OR "
        Next A
        xCadWhere2 = xCadWhere2 & ") AND "
    Else
        xCadWhere = ""
        xCadWhere2 = ""
    End If
    
    If Check1.Value = 0 Then
        If Option1.Value = True Then
            VerResumen
        Else
            VerDetalle
        End If
    Else
        If Option3.Value = True Then
            If Option1.Value = True Then
                VerResumenUniNegCliente
            Else
                VerDetalleUniNegCliente
            End If
        End If
        
        If Option4.Value = True Then
            If Option1.Value = True Then
                VerResumenCliente
            Else
                VerDetalleCliente
            End If
        End If
    End If
End Sub

Sub VerResumenDetCli()
    Dim Rst As New ADODB.Recordset
    Dim xCursor, xSQL As String
    Dim A As Integer
    
    xCursor = "SELECT vta_ventas.idcli, mae_cliente.numruc, mae_cliente.nombre, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion, " _
        & " IIf([vta_ventas].[idmon]=1,[vta_ventasunieg]![importe], " _
        & " IIf([vta_ventas].[tc]<>0,[vta_ventasunieg]![importe]*[vta_ventas].[tc],[vta_ventasunieg]![importe]*[con_tc].[impven])) AS impsol, " _
        & " IIf([vta_ventas].[idmon]=2,[vta_ventasunieg]![importe]," _
        & " IIf([vta_ventas].[tc]<>0,[vta_ventasunieg]![importe]/[vta_ventas].[tc],[vta_ventasunieg]![importe]/[con_tc].[impven])) AS impdol " _
        & " FROM (((vta_ventasunieg LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) LEFT JOIN con_unidadnegocio " _
        & " ON vta_ventasunieg.iduneg = con_unidadnegocio.id) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN con_tc " _
        & " ON vta_ventas.fchdoc = con_tc.fecha " _
        & " Where ( " _
        & " (vta_ventas.idcli = 530 Or vta_ventas.idcli = 348) And " _
        & " ((vta_ventas.tipdoc) <> 7)) ORDER BY mae_cliente.nombre, con_unidadnegocio.descripcion " _
        & " Union " _
        & " SELECT vta_ventas.idcli, mae_cliente.numruc, mae_cliente.nombre, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion, " _
        & " IIf([vta_ventas].[idmon]=1,0-[vta_ventasunieg]![importe]," _
        & " IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg]![importe]*[vta_ventas].[tc]),0-([vta_ventasunieg]![importe]*[con_tc].[impven]))) AS impsol, " _
        & " IIf([vta_ventas].[idmon]=2,0-([vta_ventasunieg]![importe]), " _
        & " IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg]![importe]/[vta_ventas].[tc]),0-([vta_ventasunieg]![importe]/[con_tc].[impven]))) AS impdol " _
        & " FROM (((vta_ventasunieg LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) LEFT JOIN con_unidadnegocio " _
        & " ON vta_ventasunieg.iduneg = con_unidadnegocio.id) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) " _
        & " LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
        & " Where (" _
        & " (vta_ventas.idcli = 530 Or vta_ventas.idcli = 348) And " _
        & " ((vta_ventas.tipdoc) = 7)) ORDER BY mae_cliente.nombre, con_unidadnegocio.descripcion"
        
    xSQL = "SELECT aa_union.iduneg, aa_union.descripcion, aa_union.numruc, aa_union.nombre, Sum(aa_union.impsol) AS SumaDeimpsol, " _
        & " Sum(aa_union.impdol) AS SumaDeimpdol " _
        & " From " _
        & " (" & xCursor _
        & " ) AS aa_union " _
        & " GROUP BY aa_union.iduneg, aa_union.descripcion, aa_union.numruc, aa_union.nombre ORDER BY aa_union.descripcion, aa_union.nombre"

    RST_Busq Rst, xSQL, xCon
    Fg1.Rows = 2
    If Rst.RecordCount <> 0 Then
        Dim xNumRuc As String
        Dim xTotCliSol, xTotCliDol, xGranTotCliSol, xGranTotCliDol As Double
        
        Rst.MoveFirst
        xNumRuc = Rst("numruc")
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("numruc")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("nombre")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(Rst("sumadeimpsol"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(Rst("sumadeimpdol"), FORMAT_MONTO)
            
            xTotCliSol = xTotCliSol + Rst("sumadeimpsol")
            xTotCliDol = xTotCliDol + Rst("sumadeimpdol")
            
            xGranTotCliSol = xGranTotCliSol + Rst("sumadeimpsol")
            xGranTotCliDol = xGranTotCliDol + Rst("sumadeimpdol")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Fg1.Rows = Fg1.Rows + 1
                ' IMPRIMIMOS EL TOTAL POR CLIENTE
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &H800000, True, &HE2FEFB, "TOTAL CLIENTE==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H80000012, True, &HE2FEFB, Format(xTotCliSol, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(xTotCliDol, FORMAT_MONTO)
                
                xTotCliSol = 0
                xTotCliDol = 0
                
                Exit For
            End If
            
            If xNumRuc <> Rst("numruc") Then
                Fg1.Rows = Fg1.Rows + 1
                ' IMPRIMIMOS EL TOTAL POR CLIENTE
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &H800000, True, &HE2FEFB, "TOTAL CLIENTE==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H80000012, True, &HE2FEFB, Format(xTotCliSol, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(xTotCliDol, FORMAT_MONTO)
                Fg1.Rows = Fg1.Rows + 1
                xTotCliSol = 0
                xTotCliDol = 0
                
                xNumRuc = Rst("numruc")
            End If
        Next A
        
        ' IMPRIMIMOS EL GRAN TOTAL
        Fg1.Rows = Fg1.Rows + 2
                
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &H800000, True, &HE2FEFB, "TOTAL GENERAL==>"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H80000012, True, &HE2FEFB, Format(xGranTotCliSol, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(xGranTotCliDol, FORMAT_MONTO)
    End If
End Sub

Sub VerDetalleUniNegCliente()
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    
    RST_Busq Rst, "SELECT vta_ventas.idcli, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion AS descunineg, mae_cliente.nombre, alm_inventario.codpro, " _
        & " alm_inventario.descripcion AS descitem, 'V' AS tipo, [vta_ventas]![numser] & '-' & [vta_ventas]![numdoc] AS numdoc, mae_documento.abrev, " _
        & " mae_moneda.simbolo, vta_ventas.fchdoc, vta_ventasunieg.importe, vta_ventas.tc, " _
        & " IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=1,[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]*[vta_ventas].[tc],[vta_ventasunieg].[importe]*[con_tc].[impven])), " _
        & " IIf([vta_ventas].[idmon]=1,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]*[vta_ventas].[tc]),0-([vta_ventasunieg].[importe]*[con_tc].[impven])))) AS impsol, " _
        & " IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=2,[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]/[vta_ventas].[tc],[vta_ventasunieg].[importe]/[con_tc].[impven])), " _
        & " IIf([vta_ventas].[idmon]=2,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]/[vta_ventas].[tc]),0-([vta_ventasunieg].[importe]/[con_tc].[impven])))) AS impdol " _
        & " FROM mae_moneda RIGHT JOIN (((mae_cliente RIGHT JOIN ((((vta_ventasunieg LEFT JOIN con_unidadnegocio ON vta_ventasunieg.iduneg = con_unidadnegocio.id) " _
        & " LEFT JOIN alm_inventario ON vta_ventasunieg.iditem = alm_inventario.id) LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) " _
        & " LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
        & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_moneda.id = vta_ventas.idmon " _
        & " WHERE ( " & xCadWhere2 _
        & " ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " ORDER BY con_unidadnegocio.descripcion, mae_cliente.nombre, alm_inventario.descripcion, vta_ventas.fchdoc", xCon

    'RST_Busq Rst, "SELECT vta_ventas.idcli, mae_cliente.nombre, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion AS descunineg, alm_inventario.codpro, " _
        & " alm_inventario.descripcion AS descitem, 'V' AS tipo, [vta_ventas]![numser] & '-' & [vta_ventas]![numdoc] AS numdoc, mae_documento.abrev, " _
        & " mae_moneda.simbolo, vta_ventas.fchdoc, vta_ventasunieg.importe, vta_ventas.tc, " _
        & " IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=1,[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]*[vta_ventas].[tc],[vta_ventasunieg].[importe]*[con_tc].[impven]))," _
        & " IIf([vta_ventas].[idmon]=1,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]*[vta_ventas].[tc]),0-([vta_ventasunieg].[importe]*[con_tc].[impven])))) AS impsol, " _
        & " IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=2,[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]/[vta_ventas].[tc],[vta_ventasunieg].[importe]/[con_tc].[impven]))," _
        & " IIf([vta_ventas].[idmon]=2,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]/[vta_ventas].[tc]),0-([vta_ventasunieg].[importe]/[con_tc].[impven])))) AS impdol " _
        & " FROM mae_moneda RIGHT JOIN (((mae_cliente RIGHT JOIN ((((vta_ventasunieg LEFT JOIN con_unidadnegocio ON vta_ventasunieg.iduneg = con_unidadnegocio.id) " _
        & " LEFT JOIN alm_inventario ON vta_ventasunieg.iditem = alm_inventario.id) LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) " _
        & " LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
        & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_moneda.id = vta_ventas.idmon " _
        & " WHERE (" & xCadWhere2 _
        & " ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " ORDER BY mae_cliente.nombre, con_unidadnegocio.descripcion, alm_inventario.descripcion, vta_ventas.fchdoc ", xCon

    Fg1.Rows = 2
    Dim xgTotC, xgTotV As Double
    
    xTotC = 0
    xTotV = 0
    xgTotC = 0
    xgTotV = 0
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xTotV = 0
        xTotC = 0
        Dim xIdCli As Integer
        xIdCli = Rst("iduneg")
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("descunineg")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("nombre")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("descitem")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Rst("abrev")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Rst("numdoc")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Rst("fchdoc")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Rst("simbolo")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Rst("importe")
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Rst("tc")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(Rst("impsol"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(Rst("impdol"), FORMAT_MONTO)
            
            xTotC = xTotC + NulosN(Rst("impsol"))
            xTotV = xTotV + NulosN(Rst("impdol"))
            
            xgTotC = xgTotC + NulosN(Rst("impsol"))
            xgTotV = xgTotV + NulosN(Rst("impdol"))
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL UNIDAD==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xTotC, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xTotV, FORMAT_MONTO)
                Exit For
            End If
            
            If xIdCli <> Rst("iduneg") Then
                Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL UNIDAD==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xTotC, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xTotV, FORMAT_MONTO)
                Fg1.Rows = Fg1.Rows + 1
                xTotV = 0
                xTotC = 0
                xIdCli = Rst("iduneg")
            End If
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL GENERAL ==>"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xgTotC, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xgTotV, FORMAT_MONTO)
    End If
End Sub

Sub VerDetalleCliente()
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    RST_Busq Rst, "SELECT vta_ventas.idcli, mae_cliente.nombre, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion AS descunineg, alm_inventario.codpro, " _
        & " alm_inventario.descripcion AS descitem, 'V' AS tipo, [vta_ventas]![numser] & '-' & [vta_ventas]![numdoc] AS numdoc, mae_documento.abrev, " _
        & " mae_moneda.simbolo, vta_ventas.fchdoc, vta_ventasunieg.importe, vta_ventas.tc, " _
        & " IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=1,[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]*[vta_ventas].[tc],[vta_ventasunieg].[importe]*[con_tc].[impven]))," _
        & " IIf([vta_ventas].[idmon]=1,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]*[vta_ventas].[tc]),0-([vta_ventasunieg].[importe]*[con_tc].[impven])))) AS impsol, " _
        & " IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=2,[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]/[vta_ventas].[tc],[vta_ventasunieg].[importe]/[con_tc].[impven]))," _
        & " IIf([vta_ventas].[idmon]=2,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]/[vta_ventas].[tc]),0-([vta_ventasunieg].[importe]/[con_tc].[impven])))) AS impdol " _
        & " FROM mae_moneda RIGHT JOIN (((mae_cliente RIGHT JOIN ((((vta_ventasunieg LEFT JOIN con_unidadnegocio ON vta_ventasunieg.iduneg = con_unidadnegocio.id) " _
        & " LEFT JOIN alm_inventario ON vta_ventasunieg.iditem = alm_inventario.id) LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) " _
        & " LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
        & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) ON mae_moneda.id = vta_ventas.idmon " _
        & " WHERE (" & xCadWhere2 _
        & " ((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " ORDER BY mae_cliente.nombre, con_unidadnegocio.descripcion, alm_inventario.descripcion, vta_ventas.fchdoc ", xCon

    Fg1.Rows = 2
    Dim xgTotC, xgTotV As Double
    
    xTotC = 0
    xTotV = 0
    xgTotC = 0
    xgTotV = 0
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xTotV = 0
        xTotC = 0
        Dim xIdCli As Integer
        xIdCli = Rst("idcli")
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("nombre")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("descunineg")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("descitem")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Rst("abrev")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Rst("numdoc")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Rst("fchdoc")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Rst("simbolo")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Rst("importe")
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Rst("tc")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(Rst("impsol"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(Rst("impdol"), FORMAT_MONTO)
            
            xTotC = xTotC + NulosN(Rst("impsol"))
            xTotV = xTotV + NulosN(Rst("impdol"))
            
            xgTotC = xgTotC + NulosN(Rst("impsol"))
            xgTotV = xgTotV + NulosN(Rst("impdol"))
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL CLIENTE==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xTotC, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xTotV, FORMAT_MONTO)
                Exit For
            End If
            
            If xIdCli <> Rst("idcli") Then
                Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL CLIENTE==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xTotC, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xTotV, FORMAT_MONTO)
                Fg1.Rows = Fg1.Rows + 1
                xTotV = 0
                xTotC = 0
                xIdCli = Rst("idcli")
            End If
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL CLIENTE==>"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xgTotC, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xgTotV, FORMAT_MONTO)
        
    End If

End Sub

Sub VerResumenCliente()
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    
    RST_Busq Rst, "SELECT vta_ventas.idcli, mae_cliente.nombre, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion AS descunineg, " _
        & " alm_inventario.codpro, alm_inventario.descripcion, 'V' AS tipo, Sum(IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=1,[vta_ventasunieg].[importe]," _
        & " IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]*[vta_ventas].[tc],[vta_ventasunieg].[importe]*[con_tc].[impven]))," _
        & " IIf([vta_ventas].[idmon]=1,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]*[vta_ventas].[tc])," _
        & " 0-([vta_ventasunieg].[importe]*[con_tc].[impven]))))) AS impsol, Sum(IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=2,[vta_ventasunieg].[importe]," _
        & " IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]/[vta_ventas].[tc],[vta_ventasunieg].[importe]/[con_tc].[impven])), " _
        & " IIf([vta_ventas].[idmon]=2,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]/[vta_ventas].[tc])," _
        & " 0-([vta_ventasunieg].[importe]/[con_tc].[impven]))))) AS impdol" _
        & " FROM (mae_cliente RIGHT JOIN (((vta_ventasunieg LEFT JOIN con_unidadnegocio ON vta_ventasunieg.iduneg = con_unidadnegocio.id) " _
        & " LEFT JOIN alm_inventario ON vta_ventasunieg.iditem = alm_inventario.id) LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) " _
        & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
        & " WHERE (((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')))" _
        & " GROUP BY vta_ventas.idcli, mae_cliente.nombre, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion, alm_inventario.codpro, alm_inventario.descripcion, 'V' " _
        & xCadWhere _
        & " ORDER BY mae_cliente.nombre, con_unidadnegocio.descripcion, alm_inventario.descripcion", xCon

    Fg1.Rows = 2
    Dim xgTotC, xgTotV As Double
    
    xTotC = 0
    xTotV = 0
    xgTotC = 0
    xgTotV = 0
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xTotV = 0
        xTotC = 0
        Dim xIdCli As Integer
        xIdCli = Rst("idcli")
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("nombre")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("descunineg")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(Rst("impsol"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(Rst("impdol"), FORMAT_MONTO)
            xTotC = xTotC + NulosN(Rst("impsol"))
            xTotV = xTotV + NulosN(Rst("impdol"))
            
            xgTotC = xgTotC + NulosN(Rst("impsol"))
            xgTotV = xgTotV + NulosN(Rst("impdol"))
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL CLIENTE==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(xTotC, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xTotV, FORMAT_MONTO)
                Exit For
            End If
            
            If xIdCli <> Rst("idcli") Then
                Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL CLIENTE==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(xTotC, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xTotV, FORMAT_MONTO)
                Fg1.Rows = Fg1.Rows + 1
                xTotV = 0
                xTotC = 0
                xIdCli = Rst("idcli")
            End If
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL CLIENTE==>"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(xgTotC, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xgTotV, FORMAT_MONTO)
        
    End If

End Sub

Sub VerResumen()
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    
    Dim nSQL As String
    Dim nSQLSub As String
    
    If OptSoles.Value = True Then
'        RST_Busq Rst, "TRANSFORM Sum(Consulta5.importe2) AS SumaDeimporte2 " _
            & " SELECT Consulta5.iduneg, Consulta5.descunimed, Sum(Consulta5.importe2) AS [Total de importe2] " _
            & " FROM " _
            & " ( " _
            & " SELECT com_comprasuneg.iduneg, con_unidadnegocio.descripcion AS descunimed, alm_inventario.codpro, alm_inventario.descripcion AS descitem, " _
            & " Sum(IIf([com_compras]![idmon]=1,IIf([com_compras]![tipdoc]<>7,[com_comprasuneg]![importe],0-[com_comprasuneg]![importe])," _
            & " IIf([com_compras]![tipdoc]<>7,[com_comprasuneg]![importe]*[con_tc]![impven],(0-[com_comprasuneg]![importe])*[con_tc]![impven]))) AS importe2, " _
            & " 'C' AS tipo FROM (((com_comprasuneg LEFT JOIN con_unidadnegocio ON com_comprasuneg.iduneg = con_unidadnegocio.id) LEFT JOIN alm_inventario " _
            & " ON com_comprasuneg.iditem = alm_inventario.id) LEFT JOIN com_compras ON com_comprasuneg.idcom = com_compras.id) LEFT JOIN con_tc " _
            & " ON com_compras.fchdoc = con_tc.fecha WHERE (((com_compras.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (com_compras.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " _
            & " GROUP BY com_comprasuneg.iduneg, con_unidadnegocio.descripcion, alm_inventario.codpro, alm_inventario.descripcion, 'C' " _
            & " ORDER BY con_unidadnegocio.descripcion, alm_inventario.descripcion " _
            & " UNION " _
            & " SELECT vta_ventasunieg.iduneg, con_unidadnegocio.descripcion AS descunineg, alm_inventario.codpro, alm_inventario.descripcion AS descitem, " _
            & " Sum(IIf([vta_ventas]![idmon]=1,IIf([vta_ventas]![tipdoc]<>7,[importe],0-[importe]),IIf([vta_ventas]![tipdoc]<>7, " _
            & " [importe]*[con_tc]![impven],(0-[importe])*[con_tc]![impven]))) AS importe2, 'V' AS tipo FROM (((vta_ventasunieg LEFT JOIN con_unidadnegocio " _
            & " ON vta_ventasunieg.iduneg = con_unidadnegocio.id) LEFT JOIN alm_inventario ON vta_ventasunieg.iditem = alm_inventario.id) " _
            & " LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
            & " WHERE (((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " _
            & " GROUP BY vta_ventasunieg.iduneg, con_unidadnegocio.descripcion, alm_inventario.codpro, alm_inventario.descripcion, 'V' " _
            & " ) AS Consulta5 " _
            & " GROUP BY Consulta5.iduneg, Consulta5.descunimed PIVOT Consulta5.tipo", xCon
    Else
    End If
    
    '--Generar la subconsulta
    nSQLSub = GenerarConsulta()

    '--expresado en moneda nacional
    If OptSoles.Value = True Then
        nSQL = "TRANSFORM Sum(vista.impexmn) AS SumaDeimpexmn " _
                & " SELECT vista.iduneg, vista.descunimed, Sum(vista.impexmn) AS xTotal " _
                & " FROM ( " & nSQLSub & ") AS vista " _
                & " GROUP BY vista.iduneg, vista.descunimed " _
                & " PIVOT vista.tipo "
    Else
    '--expresado en moneda extranjera
        nSQL = "TRANSFORM Sum(vista.impexme) AS SumaDeimpexme " _
                & " SELECT vista.iduneg, vista.descunimed, Sum(vista.impexme) AS xTotal " _
                & " FROM ( " & nSQLSub & ") AS vista " _
                & " GROUP BY vista.iduneg, vista.descunimed " _
                & " PIVOT vista.tipo "
        
    End If
    
    '--ejecutar la consulta
    RST_Busq Rst, nSQL, xCon
    
    '--validar si la conexion es correcta
    If Rst.State = 0 Then Exit Sub
    
    '-----------

    Fg1.Rows = 2
    xTotC = 0
    xTotV = 0
    DoEvents
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xTotV = 0
        xTotC = 0
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("descunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Format(Rst("v"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(Rst("c"), FORMAT_MONTO)
            xTotC = xTotC + NulosN(Rst("c"))
            xTotV = xTotV + NulosN(Rst("v"))
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 1, &H800000, True, &HE2FEFB, "TOTAL ==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, &H80000012, True, &HE2FEFB, Format(xTotV, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &H80000012, True, &HE2FEFB, Format(xTotC, FORMAT_MONTO)
                Exit For
            End If
            
        Next A
    End If
End Sub

Sub VerDetalle()
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
        
    Dim nSQL As String
    Dim nSQLSub As String
        
        
'    If OptSoles.Value = True Then
'        RST_Busq Rst, "TRANSFORM Sum(Consulta5.importe2) AS SumaDeimporte2" _
            & " SELECT Consulta5.iduneg, Consulta5.descunimed, Consulta5.codpro, Consulta5.descitem, Sum(Consulta5.importe2) AS [Total de importe2] " _
            & " From " _
            & " ( " _
            & " SELECT com_comprasuneg.iduneg, con_unidadnegocio.descripcion AS descunimed, alm_inventario.codpro, alm_inventario.descripcion AS descitem, " _
            & " Sum(IIf([com_compras]![idmon]=1,IIf([com_compras]![tipdoc]<>7,[com_comprasuneg]![importe],0-[com_comprasuneg]![importe]), " _
            & " IIf([com_compras]![tipdoc]<>7,[com_comprasuneg]![importe]*[con_tc]![impven],(0-[com_comprasuneg]![importe])*[con_tc]![impven]))) AS importe2, " _
            & " 'C' AS tipo  " _
            & " FROM (((com_comprasuneg LEFT JOIN con_unidadnegocio ON com_comprasuneg.iduneg = con_unidadnegocio.id) LEFT JOIN alm_inventario " _
            & " ON com_comprasuneg.iditem = alm_inventario.id) LEFT JOIN com_compras ON com_comprasuneg.idcom = com_compras.id) LEFT JOIN con_tc " _
            & " ON com_compras.fchdoc = con_tc.fecha WHERE (((com_compras.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (com_compras.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " _
            & " GROUP BY com_comprasuneg.iduneg, con_unidadnegocio.descripcion, alm_inventario.codpro, alm_inventario.descripcion, 'C' " _
            & " ORDER BY con_unidadnegocio.descripcion, alm_inventario.descripcion " _
            & " UNION " _
            & " SELECT vta_ventasunieg.iduneg, con_unidadnegocio.descripcion AS descunineg, alm_inventario.codpro, alm_inventario.descripcion AS descitem, " _
            & " Sum(IIf([vta_ventas]![idmon]=1,IIf([vta_ventas]![tipdoc]<>7,[importe],0-[importe]),IIf([vta_ventas]![tipdoc]<>7,[importe]*[con_tc]![impven], " _
            & " (0-[importe])*[con_tc]![impven]))) AS importe2, 'V' AS tipo " _
            & " FROM (((vta_ventasunieg LEFT JOIN con_unidadnegocio ON vta_ventasunieg.iduneg = con_unidadnegocio.id) LEFT JOIN alm_inventario " _
            & " ON vta_ventasunieg.iditem = alm_inventario.id) LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) LEFT JOIN con_tc " _
            & " ON vta_ventas.fchdoc = con_tc.fecha WHERE (((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " _
            & " GROUP BY vta_ventasunieg.iduneg, con_unidadnegocio.descripcion, alm_inventario.codpro, alm_inventario.descripcion, 'V' " _
            & " ) AS Consulta5 " _
            & " GROUP BY Consulta5.iduneg, Consulta5.descunimed, Consulta5.codpro, Consulta5.descitem PIVOT Consulta5.tipo", xCon
'    Else
'        RST_Busq Rst, "TRANSFORM Sum(Consulta5.importe2) AS SumaDeimporte2" _
            & " SELECT Consulta5.iduneg, Consulta5.descunimed, Consulta5.codpro, Consulta5.descitem, Sum(Consulta5.importe2) AS [Total de importe2] " _
            & " From " _
            & " ( " _
            & " SELECT com_comprasuneg.iduneg, con_unidadnegocio.descripcion AS descunimed, alm_inventario.codpro, alm_inventario.descripcion AS descitem, " _
            & " Sum(IIf([com_compras]![idmon]=2,IIf([com_compras]![tipdoc]<>7,[com_comprasuneg]![importe],0-[com_comprasuneg]![importe]), " _
            & " IIf([com_compras]![tipdoc]<>7,[com_comprasuneg]![importe]/[con_tc]![impven],(0-[com_comprasuneg]![importe])/[con_tc]![impven]))) AS importe2, " _
            & " 'C' AS tipo  " _
            & " FROM (((com_comprasuneg LEFT JOIN con_unidadnegocio ON com_comprasuneg.iduneg = con_unidadnegocio.id) LEFT JOIN alm_inventario " _
            & " ON com_comprasuneg.iditem = alm_inventario.id) LEFT JOIN com_compras ON com_comprasuneg.idcom = com_compras.id) LEFT JOIN con_tc " _
            & " ON com_compras.fchdoc = con_tc.fecha WHERE (((com_compras.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (com_compras.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " _
            & " GROUP BY com_comprasuneg.iduneg, con_unidadnegocio.descripcion, alm_inventario.codpro, alm_inventario.descripcion, 'C' " _
            & " ORDER BY con_unidadnegocio.descripcion, alm_inventario.descripcion " _
            & " UNION " _
            & " SELECT vta_ventasunieg.iduneg, con_unidadnegocio.descripcion AS descunineg, alm_inventario.codpro, alm_inventario.descripcion AS descitem, " _
            & " Sum(IIf([vta_ventas]![idmon]=2,IIf([vta_ventas]![tipdoc]<>7,[importe],0-[importe]),IIf([vta_ventas]![tipdoc]<>7,[importe]/[con_tc]![impven], " _
            & " (0-[importe])/[con_tc]![impven]))) AS importe2, 'V' AS tipo " _
            & " FROM (((vta_ventasunieg LEFT JOIN con_unidadnegocio ON vta_ventasunieg.iduneg = con_unidadnegocio.id) LEFT JOIN alm_inventario " _
            & " ON vta_ventasunieg.iditem = alm_inventario.id) LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) LEFT JOIN con_tc " _
            & " ON vta_ventas.fchdoc = con_tc.fecha WHERE (((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " _
            & " GROUP BY vta_ventasunieg.iduneg, con_unidadnegocio.descripcion, alm_inventario.codpro, alm_inventario.descripcion, 'V' " _
            & " ) AS Consulta5 " _
            & " GROUP BY Consulta5.iduneg, Consulta5.descunimed, Consulta5.codpro, Consulta5.descitem PIVOT Consulta5.tipo", xCon
 '   End If
    
    '--Generar la subconsulta
    nSQLSub = GenerarConsulta()

    '--expresado en moneda nacional
    If OptSoles.Value = True Then
        nSQL = "TRANSFORM Sum(vista.impexmn) AS SumaDeimpexmn " _
                & " SELECT vista.iduneg, vista.descunimed, vista.codpro, vista.descitem, Sum(vista.impexmn) AS xTotal " _
                & " FROM ( " & nSQLSub & ") AS vista " _
                & " GROUP BY vista.iduneg, vista.descunimed, vista.codpro, vista.descitem " _
                & " ORDER BY vista.descunimed,vista.codpro " _
                & " PIVOT vista.tipo "
    Else
    '--expresado en moneda extranjera
        nSQL = "TRANSFORM Sum(vista.impexme) AS SumaDeimpexme " _
                & " SELECT vista.iduneg, vista.descunimed, vista.codpro, vista.descitem, Sum(vista.impexme) AS xTotal " _
                & " FROM ( " & nSQLSub & ") AS vista " _
                & " GROUP BY vista.iduneg, vista.descunimed, vista.codpro, vista.descitem " _
                & " ORDER BY vista.descunimed,vista.codpro " _
                & " PIVOT vista.tipo "
        
    End If
    
    '--ejecutar la consulta
    RST_Busq Rst, nSQL, xCon
    
    '--validar si la conexion es correcta
    If Rst.State = 0 Then Exit Sub
    
    
    
    Fg1.Rows = 2
        
    GTotV = 0
    GTotC = 0
    xTotC = 0
    xTotV = 0
    
    If Rst.RecordCount <> 0 Then
        
        Rst.MoveFirst
        xNomUnidad = Rst("descunimed")
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("descunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("descitem")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(Rst("v"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(Rst("c"), FORMAT_MONTO)
            xTotC = xTotC + NulosN(Rst("c"))
            xTotV = xTotV + NulosN(Rst("v"))
            
            Rst.MoveNext
            If Rst.EOF = True Then

                Totalizar xNomUnidad
                
                Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, &H800000, True, &HE2FEFB, "TOTAL ==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H80000012, True, &HE2FEFB, Format(GTotV, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(GTotC, FORMAT_MONTO)
                
                Fg1.Rows = Fg1.Rows + 1
                Exit For
            End If
            
            If xNomUnidad <> Rst("descunimed") Then
                Totalizar xNomUnidad
                xNomUnidad = Rst("descunimed")
            End If
        Next A
    End If
End Sub

Sub Totalizar(NomUnidadNegocio As String)
    Fg1.Rows = Fg1.Rows + 1
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 2, &H800000, True, &HE2FEFB, "TOTAL ==>"
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &H800000, True, &HE2FEFB, UCase(NomUnidadNegocio)
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H80000012, True, &HE2FEFB, Format(xTotV, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(xTotC, FORMAT_MONTO)
    
    GTotC = GTotC + xTotC
    GTotV = GTotV + xTotV
    
    xTotC = 0
    xTotV = 0
    Fg1.Rows = Fg1.Rows + 1
End Sub

Private Sub menu_01_Click()
    If Fg2.Rows = 0 Then
        Fg2.Rows = Fg2.Rows + 1
        Exit Sub
    End If
    
    If NulosC(Fg2.TextMatrix(Fg2.Rows - 1, 1)) <> "" Then
        Fg2.Rows = Fg2.Rows + 1
    End If
End Sub

Private Sub menu_03_Click()
    If Fg2.Rows <> 0 Then
        Fg2.RemoveItem Fg2.Row
    End If
    If Fg2.Rows = 0 Then
        Fg2.Rows = Fg2.Rows + 1
    End If
End Sub

Private Sub menu_1_Click()
    If Fg2.Rows = 0 Then
        Fg2.Rows = Fg2.Rows + 1
        Fg2_CellButtonClick Fg2.Rows - 1, 1
    Else
        If NulosC(Fg2.TextMatrix(Fg2.Rows - 1, 0)) = "" Then Exit Sub
        Fg2.Rows = Fg2.Rows + 1
        Fg2_CellButtonClick Fg2.Rows - 1, 1
    End If
End Sub

Private Sub menu_3_Click()
    If Fg2.Rows = 0 Then Exit Sub
    Fg2.RemoveItem Fg2.Row
End Sub

Private Sub Option1_Click()
    If Check1.Value = 0 Then
        SetearCuadricula Fg1, 8, xCon, 2, 2, False
    Else
        If Option3.Value = True Then
            SetearCuadricula Fg1, 8, xCon, 2, 5, False
        End If
        
        If Option4.Value = True Then
            SetearCuadricula Fg1, 8, xCon, 2, 3, False
        End If
    End If
End Sub

Private Sub Option2_Click()
    If Check1.Value = 0 Then
        SetearCuadricula Fg1, 8, xCon, 2, 1, False
    Else
        If Option3.Value = True Then
            SetearCuadricula Fg1, 8, xCon, 2, 6, False
        End If
        If Option4.Value = True Then
            SetearCuadricula Fg1, 8, xCon, 2, 4, False
        End If
    End If
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then
        If Option1.Value = True Then Option1_Click
        If Option2.Value = True Then Option2_Click
    End If
End Sub

Private Sub Option4_Click()
    If Option4.Value = True Then
        If Option1.Value = True Then Option1_Click
        If Option2.Value = True Then Option2_Click
    End If
End Sub

Sub VerResumenUniNegCliente()
    Dim Rst As New ADODB.Recordset
    Dim A, B As Integer
    
    RST_Busq Rst, "SELECT vta_ventas.idcli, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion AS descunineg, mae_cliente.nombre, " _
        & " alm_inventario.codpro, alm_inventario.descripcion, 'V' AS tipo, Sum(IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=1,[vta_ventasunieg].[importe], " _
        & " IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]*[vta_ventas].[tc],[vta_ventasunieg].[importe]*[con_tc].[impven])), " _
        & " IIf([vta_ventas].[idmon]=1,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]*[vta_ventas].[tc]), " _
        & " 0-([vta_ventasunieg].[importe]*[con_tc].[impven]))))) AS impsol, Sum(IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=2,[vta_ventasunieg].[importe], " _
        & " IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]/[vta_ventas].[tc],[vta_ventasunieg].[importe]/[con_tc].[impven])), " _
        & " IIf([vta_ventas].[idmon]=2,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]/[vta_ventas].[tc]), " _
        & " 0-([vta_ventasunieg].[importe]/[con_tc].[impven]))))) AS impdol " _
        & " FROM (mae_cliente RIGHT JOIN (((vta_ventasunieg LEFT JOIN con_unidadnegocio ON vta_ventasunieg.iduneg = con_unidadnegocio.id) " _
        & " LEFT JOIN alm_inventario ON vta_ventasunieg.iditem = alm_inventario.id) LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id)  " _
        & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
        & " WHERE (((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " _
        & " GROUP BY vta_ventas.idcli, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion, mae_cliente.nombre, alm_inventario.codpro, alm_inventario.descripcion, 'V' " _
        & xCadWhere _
        & " ORDER BY con_unidadnegocio.descripcion, mae_cliente.nombre, alm_inventario.descripcion", xCon
    
    'SELECT vta_ventas.idcli, mae_cliente.nombre, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion AS descunineg, " _
        & " alm_inventario.codpro, alm_inventario.descripcion, 'V' AS tipo, Sum(IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=1,[vta_ventasunieg].[importe]," _
        & " IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]*[vta_ventas].[tc],[vta_ventasunieg].[importe]*[con_tc].[impven]))," _
        & " IIf([vta_ventas].[idmon]=1,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]*[vta_ventas].[tc])," _
        & " 0-([vta_ventasunieg].[importe]*[con_tc].[impven]))))) AS impsol, Sum(IIf([vta_ventas].[tipdoc]<>7,IIf([vta_ventas].[idmon]=2,[vta_ventasunieg].[importe]," _
        & " IIf([vta_ventas].[tc]<>0,[vta_ventasunieg].[importe]/[vta_ventas].[tc],[vta_ventasunieg].[importe]/[con_tc].[impven])), " _
        & " IIf([vta_ventas].[idmon]=2,0-[vta_ventasunieg].[importe],IIf([vta_ventas].[tc]<>0,0-([vta_ventasunieg].[importe]/[vta_ventas].[tc])," _
        & " 0-([vta_ventasunieg].[importe]/[con_tc].[impven]))))) AS impdol" _
        & " FROM (mae_cliente RIGHT JOIN (((vta_ventasunieg LEFT JOIN con_unidadnegocio ON vta_ventasunieg.iduneg = con_unidadnegocio.id) " _
        & " LEFT JOIN alm_inventario ON vta_ventasunieg.iditem = alm_inventario.id) LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) " _
        & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
        & " WHERE (((vta_ventas.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchFin.Valor & "')))" _
        & " GROUP BY vta_ventas.idcli, mae_cliente.nombre, vta_ventasunieg.iduneg, con_unidadnegocio.descripcion, alm_inventario.codpro, alm_inventario.descripcion, 'V' " _
        & xCadWhere _
        & " ORDER BY mae_cliente.nombre, con_unidadnegocio.descripcion, alm_inventario.descripcion", xCon

    Fg1.Rows = 2
    Dim xgTotC, xgTotV As Double
    
    xTotC = 0
    xTotV = 0
    xgTotC = 0
    xgTotV = 0
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        xTotV = 0
        xTotC = 0
        Dim xIdCli As Integer
        xIdCli = Rst("iduneg")
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("descunineg")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("nombre")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(Rst("impsol"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(Rst("impdol"), FORMAT_MONTO)
            xTotC = xTotC + NulosN(Rst("impsol"))
            xTotV = xTotV + NulosN(Rst("impdol"))
            
            xgTotC = xgTotC + NulosN(Rst("impsol"))
            xgTotV = xgTotV + NulosN(Rst("impdol"))
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL UNIDAD==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(xTotC, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xTotV, FORMAT_MONTO)
                Exit For
            End If
            
            If xIdCli <> Rst("iduneg") Then
                Fg1.Rows = Fg1.Rows + 1
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL UNIDAD==>"
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(xTotC, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xTotV, FORMAT_MONTO)
                Fg1.Rows = Fg1.Rows + 1
                xTotV = 0
                xTotC = 0
                xIdCli = Rst("iduneg")
            End If
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 4, &H800000, True, &HE2FEFB, "TOTAL GENERAL==>"
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 5, &H80000012, True, &HE2FEFB, Format(xgTotC, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xgTotV, FORMAT_MONTO)
        
    End If

End Sub




Private Function GenerarConsulta() As String
    '===================================================================================================
    'Creado : 10/06/11 Por: Johan Castro
    'Propósito: Generar la consulta de seleccion para mostrar en pantalla
    '
    'Entradas:  Ninguno
    '
    'Resultados: Setencia SQL listo a usar
    '
    '===================================================================================================

    Dim nSQL As String

    '--consulta de compras
    nSQL = "SELECT com_compras.id AS iddoc, com_comprasuneg.iduneg, Left([com_compras].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([com_compras].[numreg],4) AS registro, mae_prov.numruc, mae_prov.nombre, " _
        + vbCr + " [com_compras].[numser] & '-' & [com_compras].[numdoc] AS numerodoc, mae_documento.abrev, com_compras.fchdoc, mae_moneda.simbolo,  " _
        + vbCr + " IIf(com_compras.tc=0 Or com_compras.tc Is Null,con_tc.impven,com_compras.tc) AS tipcam, " _
        + vbCr + " IIf(com_compras.tipdoc=7,(-1)*com_comprasuneg.importe,com_comprasuneg.importe) AS impreal, " _
        + vbCr + " IIf(com_compras.idmon=1,impreal,impreal*tipcam) AS impexmn, " _
        + vbCr + " IIf(com_compras.idmon=2,impreal,IIf(tipcam=0,0,impreal/tipcam)) AS impexme, " _
        + vbCr + " 'C' AS tipo, con_unidadnegocio.descripcion AS descunimed, alm_inventario.codpro, alm_inventario.descripcion AS descitem " _
        + vbCr + " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (((((com_comprasuneg LEFT JOIN con_unidadnegocio ON com_comprasuneg.iduneg = con_unidadnegocio.id) LEFT JOIN alm_inventario ON com_comprasuneg.iditem = alm_inventario.id) LEFT JOIN com_compras ON com_comprasuneg.idcom = com_compras.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
        + vbCr + " WHERE (((com_compras.fchdoc) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) "
    
    '--consulta de honorarios
    nSQL = nSQL _
        + vbCr + " UNION " _
        + vbCr + " SELECT com_honorarios.id AS iddoc, com_honorariosuneg.iduneg, Left([com_honorarios].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([com_honorarios].[numreg],4) AS registro, mae_prov.numruc, mae_prov.nombre," _
        + vbCr + " [com_honorarios].[numser] & '-' & [com_honorarios].[numdoc] AS numerodoc, mae_documento.abrev, com_honorarios.fchdoc, mae_moneda.simbolo, " _
        + vbCr + " IIf(com_honorarios.tc=0 Or com_honorarios.tc Is Null,con_tc.impven,com_honorarios.tc) AS tipcam, " _
        + vbCr + " IIf(com_honorarios.tipdoc=7,(-1)*com_honorariosuneg.importe,com_honorariosuneg.importe) AS impreal, " _
        + vbCr + " IIf(com_honorarios.idmon=1,impreal,impreal*tipcam) AS impexmn, " _
        + vbCr + " IIf(com_honorarios.idmon=2,impreal,IIf(tipcam=0,0,impreal/tipcam)) AS impexme, " _
        + vbCr + " 'C' AS tipo, con_unidadnegocio.descripcion AS descunimed, alm_inventario.codpro, alm_inventario.descripcion AS descitem " _
        + vbCr + " FROM (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((com_honorariosuneg LEFT JOIN con_unidadnegocio ON com_honorariosuneg.iduneg = con_unidadnegocio.id) LEFT JOIN alm_inventario ON com_honorariosuneg.iditem = alm_inventario.id) LEFT JOIN com_honorarios ON com_honorariosuneg.idhon = com_honorarios.id) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) ON mae_documento.id = com_honorarios.tipdoc) ON mae_moneda.id = com_honorarios.idmon) ON mae_prov.id = com_honorarios.idpro) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id " _
        + vbCr + " WHERE (((com_honorarios.fchdoc) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) "
            
    '--consulta de ventas
    nSQL = nSQL _
        + vbCr + " UNION " _
        + vbCr + " SELECT vta_ventas.id AS iddoc, vta_ventasunieg.iduneg, Left([vta_ventas].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre, " _
        + vbCr + " mae_documento.abrev, [vta_ventas].[numser] & '-' & [vta_ventas].[numdoc] AS numerodoc, vta_ventas.fchdoc, mae_moneda.simbolo, " _
        + vbCr + " IIf(vta_ventas.tc=0 Or vta_ventas.tc Is Null,con_tc.impven,vta_ventas.tc) AS tipcam, " _
        + vbCr + " IIf(vta_ventas.tipdoc=7,(-1)*vta_ventasunieg.importe,vta_ventasunieg.importe) AS impreal, " _
        + vbCr + " IIf(vta_ventas.idmon=1,impreal,impreal*tipcam) AS impexmn, " _
        + vbCr + " IIf(vta_ventas.idmon=2,impreal,IIf(tipcam=0,0,impreal/tipcam)) AS impexme, " _
        + vbCr + " 'V' AS tipo, con_unidadnegocio.descripcion AS descunineg, alm_inventario.codpro, alm_inventario.descripcion AS descitem " _
        + vbCr + " FROM mae_libros RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN ((((vta_ventasunieg LEFT JOIN con_unidadnegocio ON vta_ventasunieg.iduneg = con_unidadnegocio.id) LEFT JOIN alm_inventario ON vta_ventasunieg.iditem = alm_inventario.id) LEFT JOIN vta_ventas ON vta_ventasunieg.idvta = vta_ventas.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_cliente.id = vta_ventas.idcli) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_libros.id = vta_ventas.idlib " _
        + vbCr + " WHERE (((vta_ventas.fchdoc) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) "

GenerarConsulta = nSQL

End Function
