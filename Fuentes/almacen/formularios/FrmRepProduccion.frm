VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmRepProduccion 
   Caption         =   "Produccion  -  Reporte de Producción"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   13320
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   3570
      TabIndex        =   4
      Top             =   3810
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   5
         Top             =   465
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape Shape1 
         Height          =   765
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   5805
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "lbl(1)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1395
         TabIndex        =   8
         Top             =   180
         Width           =   2670
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
         Left            =   225
         TabIndex        =   7
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   4170
         TabIndex        =   6
         Top             =   180
         Width           =   1530
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   11070
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepProduccion.frx":2A98
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   13350
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   795
         Index           =   1
         Left            =   5370
         TabIndex        =   14
         ToolTipText     =   "Buscar Linea"
         Top             =   90
         Width           =   3345
         _cx             =   5900
         _cy             =   1402
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepProduccion.frx":2E2A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
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
      Begin VB.Frame Frame2 
         Caption         =   "[ Criterios de Selección]"
         Height          =   885
         Left            =   1920
         TabIndex        =   15
         Top             =   0
         Width           =   3405
         Begin VB.OptionButton Opt 
            Caption         =   "xLote"
            Height          =   195
            Index           =   3
            Left            =   1320
            TabIndex        =   19
            Top             =   570
            Width           =   765
         End
         Begin VB.OptionButton Opt 
            Caption         =   "xTipo de Ítem"
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   18
            Top             =   330
            Width           =   1335
         End
         Begin VB.OptionButton Opt 
            Caption         =   "xAlmacén"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   17
            Top             =   570
            Width           =   1305
         End
         Begin VB.OptionButton Opt 
            Caption         =   "xÍtem"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   16
            Top             =   330
            Width           =   765
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "[Fech. Ing.]"
         Height          =   885
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1875
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDesde 
            Height          =   300
            Left            =   555
            TabIndex        =   10
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchHasta 
            Height          =   300
            Left            =   555
            TabIndex        =   11
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
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
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   45
            TabIndex        =   13
            Top             =   255
            Width           =   510
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   45
            TabIndex        =   12
            Top             =   585
            Width           =   465
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   6900
         Index           =   0
         Left            =   30
         TabIndex        =   1
         Top             =   930
         Width           =   13245
         _cx             =   23363
         _cy             =   12171
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepProduccion.frx":2E87
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
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO_01"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   1365
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu00 
         Caption         =   "Insertar Item"
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu menu01 
         Caption         =   "Eliminar Item"
      End
   End
End
Attribute VB_Name = "FrmRepProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CARGO As Boolean
Dim cSQL As String
Dim RstResumido As New ADODB.Recordset
Dim RstNumper As New ADODB.Recordset
Dim RstTareas As New ADODB.Recordset
Dim INDICE_ As Integer
Dim AGREGANDO_ As Boolean
Dim INTERRUMPIR_ As Boolean

Private Sub Buscar()
    generarConsulta
End Sub

Private Function verificarDatos() As Boolean
    Dim VERIFICO_ As Boolean
    Dim MENSAJE_ As String
    
    VERIFICO_ = True
    If (Not IsDate(TxtFchDesde.Valor) Or Not IsDate(TxtFchHasta.Valor)) Then
        MENSAJE_ = "Ingrese un valor adecuado para la Fecha de Produccion"
        VERIFICO_ = False
        GoTo SALIR
    End If
    
    If (CDate(TxtFchHasta.Valor) < CDate(TxtFchDesde.Valor)) Then
        MENSAJE_ = "La fecha Hasta no puede ser mayor que la fecha Desde"
        VERIFICO_ = False
    End If
    
SALIR:
    If Not VERIFICO_ Then MsgBox MENSAJE_, vbCritical + vbOKOnly, xTitulo
    verificarDatos = VERIFICO_
End Function

Private Function GENERAR_SQL_ID_RST(Rst As ADODB.Recordset, nDesc As String, _
                            nCampo As String, Optional nTipoIn As String = "IN", _
                            Optional fEsNumero As Boolean = True) As String
    Dim nSQL As String
    Dim k&
    nSQL = ""
    
    If Rst.RecordCount = 0 Then Exit Function Else Rst.MoveFirst
    While Not Rst.EOF
        If Trim(CStr(Rst("" & nDesc & ""))) <> "" Then
            If fEsNumero = True Then
                nSQL = nSQL & NulosN(Rst("" & nDesc & "")) & ","
            Else
                nSQL = nSQL & "'" & NulosC(Rst("" & nDesc & "")) & "',"
            End If
        End If
        Rst.MoveNext
    Wend
    
    If nSQL <> "" Then nSQL = " " & nCampo & " " & nTipoIn & " (" + Left(nSQL, Len(nSQL) - 1) & ") "
        
    GENERAR_SQL_ID_RST = nSQL
End Function

Private Sub generarConsulta()
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim CONSULTA_ As String
    
    Me.MousePointer = vbHourglass
         
    CONSULTA_ = GENERAR_SQL_ID(fg(1), 1, " AND pro_cronogramadet.iditem", "IN", True)
    
    With fg(0)
    
        .Rows = 2
        
        cSQL = "SELECT alm_inventariolotedet.idlote, alm_inventariolote.descripcion AS deslote, alm_inventariolote.fching, alm_inventariolote.iditem, alm_inventario.descripcion AS desitem, alm_inventariolotedet.idalm, alm_almacenes.descripcion AS desalm, Sum(alm_inventariolotedet.cantidad) AS cantot, alm_ingreso.idpro AS idprov, alm_ingreso.nombre AS desprov " _
            + vbCr + "FROM alm_ingreso RIGHT JOIN (alm_ingresodet RIGHT JOIN (((alm_inventariolote RIGHT JOIN alm_inventariolotedet ON alm_inventariolote.id = alm_inventariolotedet.idlote) LEFT JOIN alm_inventario ON alm_inventariolote.iditem = alm_inventario.id) LEFT JOIN alm_almacenes ON alm_inventariolotedet.idalm = alm_almacenes.id) ON alm_ingresodet.idlotedet = alm_inventariolotedet.id) ON alm_ingreso.id = alm_ingresodet.id " _
            + vbCr + "WHERE (((alm_ingreso.tipmov)=-1) AND ((alm_inventariolote.iditem)=130) AND ((alm_inventariolotedet.idalm)=1) AND ((alm_ingreso.idpro)=1602) AND ((alm_inventariolotedet.idlote)=1) AND ((alm_inventariolote.fching)>=CDate('01/01/2012') And (alm_inventariolote.fching)<=CDate('24/04/2012'))) " _
            + vbCr + "GROUP BY alm_inventariolotedet.idlote, alm_inventariolote.descripcion, alm_inventariolote.fching, alm_inventariolote.iditem, alm_inventario.descripcion, alm_inventariolotedet.idalm, alm_almacenes.descripcion, alm_ingreso.idpro, alm_ingreso.nombre " _
            + vbCr + "ORDER BY alm_inventariolote.fching, alm_inventario.descripcion;"
        
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Me.MousePointer = vbDefault: Exit Sub
        
        cCRITERIOS = GENERAR_SQL_ID_RST(xRs, "idprocorr", " AND pro_producciondet.corr", "IN", True)
        
        xRs.MoveFirst
    
        While Not xRs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = NulosN(xRs("idlote"))
            .TextMatrix(.Rows - 1, 2) = NulosN(xRs("iditem"))
            .TextMatrix(.Rows - 1, 3) = NulosN(xRs("idalm"))
            .TextMatrix(.Rows - 1, 4) = NulosN(xRs("idprov"))
            .TextMatrix(.Rows - 1, 5) = NulosC(xRs("deslote"))
            .TextMatrix(.Rows - 1, 6) = NulosC(xRs("desitem"))
            .TextMatrix(.Rows - 1, 7) = NulosC(xRs("fching"))
            .TextMatrix(.Rows - 1, 8) = NulosC(xRs("desalm"))
            .TextMatrix(.Rows - 1, 9) = NulosC(xRs("desprov"))
            .TextMatrix(.Rows - 1, 10) = NulosN(xRs("canttot"))
                        
            xRs.MoveNext
        Wend
    End With
    
    Me.MousePointer = vbDefault
    Set xRs = Nothing
End Sub

Private Sub configurarGrid()
    fg(0).ColWidth(1) = 0
    fg(0).ColWidth(2) = 0
    fg(0).ColWidth(3) = 0
    fg(0).ColWidth(4) = 0
    
    ' Fecha
    If ck(0).Value = 1 Then
        fg(0).ColWidth(11) = 850
        fg(0).ColWidth(12) = 850
        fg(0).ColWidth(13) = 600
        fg(0).ColWidth(14) = 500
    Else
        fg(0).ColWidth(11) = 0
        fg(0).ColWidth(12) = 0
        fg(0).ColWidth(13) = 0
        fg(0).ColWidth(14) = 0
    End If
    ' Cantidad
    If ck(1).Value = 1 Then
        fg(0).ColWidth(15) = 800
        fg(0).ColWidth(16) = 800
        fg(0).ColWidth(17) = 800
        fg(0).ColWidth(18) = 500
    Else
        fg(0).ColWidth(15) = 0
        fg(0).ColWidth(16) = 0
        fg(0).ColWidth(17) = 0
        fg(0).ColWidth(18) = 0
    End If
    ' Hora de Inicio
    If ck(2).Value = 1 Then
        fg(0).ColWidth(19) = 850
        fg(0).ColWidth(20) = 850
        fg(0).ColWidth(21) = 600
        fg(0).ColWidth(22) = 500
    Else
        fg(0).ColWidth(19) = 0
        fg(0).ColWidth(20) = 0
        fg(0).ColWidth(21) = 0
        fg(0).ColWidth(22) = 0
    End If
    ' Hora de Fin
    If ck(3).Value = 1 Then
        fg(0).ColWidth(23) = 850
        fg(0).ColWidth(24) = 850
        fg(0).ColWidth(25) = 600
        fg(0).ColWidth(26) = 500
    Else
        fg(0).ColWidth(23) = 0
        fg(0).ColWidth(24) = 0
        fg(0).ColWidth(25) = 0
        fg(0).ColWidth(26) = 0
    End If
    
    GRID_COMBINAR fg(0), 0, 5, 1, 5, "Item", flexAlignCenterCenter, False, flexMergeFixedOnly, &H0&, &H8000000F
    GRID_COMBINAR fg(0), 0, 6, 1, 6, "Supervisor", flexAlignCenterCenter, False, flexMergeFixedOnly, &H0&, &H8000000F
    GRID_COMBINAR fg(0), 0, 7, 1, 7, "Linea", flexAlignCenterCenter, False, flexMergeFixedOnly, &H0&, &H8000000F
    GRID_COMBINAR fg(0), 0, 8, 1, 8, "Receta", flexAlignCenterCenter, False, flexMergeFixedOnly, &H0&, &H8000000F
    GRID_COMBINAR fg(0), 0, 9, 1, 9, "Nº Reg Prod", flexAlignCenterCenter, False, flexMergeFixedOnly, &H0&, &H8000000F
    GRID_COMBINAR fg(0), 0, 10, 1, 10, "Efic(%)", flexAlignCenterCenter, False, flexMergeFixedOnly, &H0&, &H8000000F
    GRID_COMBINAR fg(0), 0, 11, 0, 14, "Fch. Prod.", flexAlignCenterCenter, True, flexMergeFixedOnly, &H0&, &H8000000F
    fg(0).TextMatrix(1, 11) = "PL"
    fg(0).TextMatrix(1, 12) = "PD"
    fg(0).TextMatrix(1, 13) = "Desv."
    fg(0).TextMatrix(1, 14) = "UM"
    GRID_COMBINAR fg(0), 0, 15, 0, 18, "Cantidad", flexAlignCenterCenter, True, flexMergeFixedOnly, &H0&, &H8000000F
    fg(0).TextMatrix(1, 15) = "PL"
    fg(0).TextMatrix(1, 16) = "PD"
    fg(0).TextMatrix(1, 17) = "Desv."
    fg(0).TextMatrix(1, 18) = "UM"
    GRID_COMBINAR fg(0), 0, 19, 0, 22, "Hor. Ini.", flexAlignCenterCenter, True, flexMergeFixedOnly, &H0&, &H8000000F
    fg(0).TextMatrix(1, 19) = "PL"
    fg(0).TextMatrix(1, 20) = "PD"
    fg(0).TextMatrix(1, 21) = "Desv."
    fg(0).TextMatrix(1, 22) = "UM"
    GRID_COMBINAR fg(0), 0, 23, 0, 26, "Hor. Fin", flexAlignCenterCenter, True, flexMergeFixedOnly, &H0&, &H8000000F
    fg(0).TextMatrix(1, 23) = "PL"
    fg(0).TextMatrix(1, 24) = "PD"
    fg(0).TextMatrix(1, 25) = "Desv."
    fg(0).TextMatrix(1, 26) = "UM"
    
End Sub

Sub EXPORTAR()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "REPORTE DE PRODUCCION"

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, fg(0), TITULO_, "", "", TITULO_
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub

Private Sub iniciarCampos()
    Dim MES_ As Integer
    Dim ANIO_ As Integer
    
    CARGO = False
    
    Set fg(1).DataSource = Nothing
    Set fg(2).DataSource = Nothing
    Set fg(3).DataSource = Nothing
    'Se inicializa:
    fg(0).Rows = 2
    'datos para clientes
    GRID_COMBOLIST fg(1), 2
    fg(1).Editable = flexEDKbdMouse
    'datos para productos
    GRID_COMBOLIST fg(2), 2
    fg(2).Editable = flexEDKbdMouse
    'datos para Ordenes de Compra
    GRID_COMBOLIST fg(3), 2
    fg(3).Editable = flexEDKbdMouse
    'datos para fechas
    TxtFchDesde.Valor = Date
    TxtFchHasta.Valor = Date
    ' datos para el reporte Simple
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).AutoSearch = flexSearchFromTop
    fg(0).ExplorerBar = flexExSortShow
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).ForeColorSel = &H80000005
    fg(0).BackColorSel = &H80&
    
    fg(1).ColWidth(1) = 0
    fg(2).ColWidth(1) = 0
    fg(3).ColWidth(1) = 0
    
    ck(0).Value = 1
    ck(1).Value = 1
    ck(2).Value = 1
    ck(3).Value = 1
    
    AGREGANDO_ = False
    INTERRUMPIR_ = False
    
    configurarGrid
End Sub

Private Sub fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    
    If Index = 1 Then ' Lineas
        ReDim xCampos(1, 4) As String
        Dim nTitulo As String
        Dim xRsAux As New ADODB.Recordset
        
        Set xRs = Nothing
        
        nSQLId = GENERAR_SQL_ID(fg(Index), 1, " AND alm_inventario.id", "NOT IN", True)

        cSQL = "SELECT alm_inventario.descripcion, pro_receta.iditem " _
            + vbCr + "FROM pro_receta LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id " _
            + vbCr + "WHERE (((alm_inventario.activo)=-1) AND ((pro_receta.prirec)=1)) " & nSQLId
        
        RST_Busq xRs, cSQL, xCon
        
        'descripcion                        'campo                           'tamaño                    'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6000":    xCampos(0, 3) = "C"
                
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "descripcion", "descripcion", Principio

        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub

        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("iditem"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("descripcion"))
    End If
    
    If Index = 2 Then ' Supervisores
        ReDim xCampos(2, 4) As String
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "id":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                
        nSQLId = GENERAR_SQL_ID(fg(2), 1, " AND pla_empleados.id", "NOT IN", True)
        
        cSQL = "SELECT pla_empleados.id, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
            + vbCr + "FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
            + vbCr + "Where (((pro_empdet.idfun) = 3) AND ((pla_empleados.nombre) Is Not Null)) " & nSQLId _
            
        nTitulo = "Buscando Personal Encargado"
                
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        If AGREGANDO_ Then
            fg(Index).Rows = fg(Index).Rows + 1
            fg(Index).Select fg(Index).Rows - 1, 1
        End If
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("id"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("nombre"))
    End If
    
    If Index = 3 Then ' Num reg Produccion
            ReDim xCampos(6, 4) As String
            
        'descripcion                        'campo                              'tamaño                         'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Num. Prod.":       xCampos(0, 1) = "numparte":          xCampos(0, 2) = "1200":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Descripcion":      xCampos(1, 1) = "despro":       xCampos(1, 2) = "3500":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Fech. Pro.":       xCampos(2, 1) = "dia":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
        xCampos(3, 0) = "Hor. Pro.":        xCampos(3, 1) = "horini":       xCampos(3, 2) = "900":          xCampos(3, 3) = "C"
        xCampos(4, 0) = "U.M":              xCampos(4, 1) = "abrev":        xCampos(4, 2) = "500":          xCampos(4, 3) = "C"
        xCampos(5, 0) = "Cantidad":         xCampos(5, 1) = "cantidad":     xCampos(5, 2) = "1000":         xCampos(5, 3) = "N"
            
        cSQL = "SELECT pro_produccion.dia, pro_receta.iditem, alm_inventario.descripcion AS despro, pro_producciondet.idrec, pro_receta.codrec, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.numparte, pro_producciondet.cantidad, pro_receta.idunimed, mae_unidades.abrev, pro_emp.idemp AS idresp, pla_empleados.nombre, pro_producciondet.corr AS idregprod " _
                + vbCr + "FROM pro_produccion LEFT JOIN (((((pro_producciondet LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) LEFT JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) ON pro_produccion.id = pro_producciondet.idpro;"

        nTitulo = "Buscando Reg. Prod."
    
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", CualquierParte
                      
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        If AGREGANDO_ Then
            fg(Index).Rows = fg(Index).Rows + 1
            fg(Index).Select fg(Index).Rows - 1, 1
        End If
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("idregprod"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("numparte"))
    End If
        
    If fg(Index).Row = fg(Index).Rows - 1 Then
        fg(Index).Rows = fg(Index).Rows + 1
        fg(Index).Select fg(Index).Rows - 1, 2
        fg(Index).TopRow = fg(Index).Rows - 1
    End If
        
    Set xRs = Nothing
End Sub

Private Sub fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 1, 2, 3
            If KeyCode = vbKeyInsert Then ' Agregar
                menu00_Click
            End If
            
            If KeyCode = vbKeyDelete Then ' Eliminar
                menu01_Click
            End If
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    INDICE_ = Index
    If Button <> 2 Then Exit Sub
    Select Case Index
        Case 1, 2, 3
            PopupMenu menu
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        INTERRUMPIR_ = True ' interrumpir
    End If
End Sub

Private Sub Form_Load()
    iniciarCampos
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Width <= 13200 Then Me.Width = 13200
    If Me.Height <= 2850 Then Me.Height = 2850
        
    ' Se dimensiona el contenido
    Frame6.Width = Me.Width - 90
    Frame6.Height = Me.Height - 795
    
    fg(0).Width = Frame6.Width - 105
    fg(0).Height = Frame6.Height - 975
End Sub

Private Sub menu00_Click()
    If fg(INDICE_).Rows > 2 Then fg(INDICE_).TopRow = fg(INDICE_).Rows - 2
    AGREGANDO_ = True
    fg_CellButtonClick INDICE_, fg(INDICE_).Rows - 1, 1
End Sub

Private Sub menu01_Click()
    If fg(INDICE_).Row < fg(INDICE_).FixedRows Then Exit Sub
    fg(INDICE_).RemoveItem fg(INDICE_).Row
    
    If fg(INDICE_).Rows = fg(INDICE_).FixedRows Then fg(INDICE_).Rows = fg(INDICE_).Rows + 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If verificarDatos Then
            Buscar
        End If
    End If
    
    If Button.Index = 5 Then
        EXPORTAR
    End If
    
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub
