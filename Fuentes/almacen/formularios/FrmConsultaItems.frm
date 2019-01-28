VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmConsultaItems 
   Caption         =   "Almacen - Reporte de Items"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.ElasticOne EO 
      Height          =   5685
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10335
      _cx             =   18230
      _cy             =   10028
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
      GridRows        =   3
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmConsultaItems.frx":0000
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   1785
         Left            =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   10275
         _cx             =   18124
         _cy             =   3149
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
         GridCols        =   3
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmConsultaItems.frx":0050
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1725
            Left            =   7155
            TabIndex        =   10
            Top             =   30
            Width           =   3090
            Begin VB.CommandButton Command4 
               Height          =   660
               Left            =   2265
               Picture         =   "FrmConsultaItems.frx":00A0
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   480
               Width           =   700
            End
            Begin VB.CommandButton Command3 
               Height          =   660
               Left            =   1530
               Picture         =   "FrmConsultaItems.frx":03AA
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   480
               Width           =   700
            End
            Begin VB.CommandButton Command1 
               Height          =   660
               Left            =   60
               Picture         =   "FrmConsultaItems.frx":0EB4
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   480
               Width           =   700
            End
            Begin VB.CommandButton Command2 
               Height          =   660
               Left            =   795
               Picture         =   "FrmConsultaItems.frx":12F6
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   480
               Width           =   700
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   1725
            Left            =   30
            TabIndex        =   3
            Top             =   30
            Width           =   6660
            Begin VB.CommandButton CmdBusTipiTem 
               Height          =   240
               Left            =   4350
               Picture         =   "FrmConsultaItems.frx":1600
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   105
               Width           =   225
            End
            Begin VB.TextBox txtTipIte 
               Height          =   300
               Left            =   1140
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   5
               Top             =   75
               Width           =   3450
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   1260
               Left            =   1155
               TabIndex        =   4
               Top             =   375
               Width           =   5385
               _cx             =   9499
               _cy             =   2222
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
               FormatString    =   $"FrmConsultaItems.frx":1732
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
            Begin VB.Label lblIdItem 
               AutoSize        =   -1  'True
               Caption         =   "lblIdItem"
               Height          =   195
               Left            =   4635
               TabIndex        =   8
               Top             =   120
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Item"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   7
               Top             =   105
               Width           =   885
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Reportes"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   6
               Top             =   405
               Width           =   645
            End
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   3450
         Left            =   30
         TabIndex        =   1
         Top             =   2205
         Width           =   10275
         _cx             =   18124
         _cy             =   6085
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsultaItems.frx":1782
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Label1"
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
         Height          =   330
         Left            =   30
         TabIndex        =   15
         Top             =   1845
         Width           =   10275
      End
   End
End
Attribute VB_Name = "FrmConsultaItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rst As New ADODB.Recordset
Dim A As Integer
Dim xId As Integer
Dim SeEjecuto  As Boolean

Private Sub CmdBusTipiTem_Click()
    'BUSCAMOS EL TIPO DE PRODUCTO QUE SE VA A LISTAR
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        txtTipIte.Text = xRs("descripcion")
        lblIdItem.Caption = xRs("id")
        'txtFam.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Command1_Click()
    If NulosC(txtTipIte.Text) = "" Then
        MsgBox "No ha especificado el tipo de item a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Label1.Caption = ""
        Fg1.Rows = 1
        txtTipIte.SetFocus
        Exit Sub
    End If
    If Fg2.Row = 0 Then ImprimirStock
    If Fg2.Row = 3 Then ProcesarStockMini
    If Fg2.Row = 4 Then ProcesarStockMax
End Sub

Sub ImprimirStock()
    Label1.Caption = Fg2.TextMatrix(Fg2.Row, 1)
    RST_Busq Rst, "SELECT alm_inventario.codpro AS Codigo, alm_inventario.descripcion AS Item,  mae_unidades.abrev AS UM, alm_inventario.preuni AS Precio, " _
        & " alm_inventario.stckact AS StockActual,  alm_inventario.tippro, alm_inventario.idfam FROM mae_familia RIGHT JOIN (mae_tipoproducto RIGHT JOIN " _
        & " (mae_unidades RIGHT JOIN alm_inventario  ON mae_unidades.id = alm_inventario.idunimed) ON mae_tipoproducto.id = alm_inventario.tippro) " _
        & " ON mae_familia.id = alm_inventario.idfam WHERE alm_inventario.tippro= " & nulosn(lblIdItem.Caption) & "  AND alm_inventario.activo = -1 ORDER BY alm_inventario.descripcion", xCon
    
    Fg1.Cols = 6
    Fg1.Rows = 1
    Fg1.TextMatrix(0, 1) = "Codigo":      Fg1.ColWidth(1) = 1400:  Fg1.ColAlignment(1) = flexAlignLeftCenter
    Fg1.TextMatrix(0, 2) = "Descripcion": Fg1.ColWidth(2) = 5500:  Fg1.ColAlignment(2) = flexAlignLeftCenter
    Fg1.TextMatrix(0, 3) = "Uni. Med.":   Fg1.ColWidth(3) = 900:   Fg1.ColAlignment(3) = flexAlignCenterCenter
    Fg1.TextMatrix(0, 4) = "Stock Act.":  Fg1.ColWidth(4) = 1000:  Fg1.ColAlignment(4) = flexAlignRightCenter
    Fg1.TextMatrix(0, 5) = "Precio":      Fg1.ColWidth(5) = 1000:  Fg1.ColAlignment(5) = flexAlignRightCenter
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("Codigo")
            Fg1.TextMatrix(A, 2) = Rst("Item")
            Fg1.TextMatrix(A, 3) = Rst("UM")
            Fg1.TextMatrix(A, 4) = Format(Rst("StockActual"), "0.00")
            Fg1.TextMatrix(A, 5) = Format(Rst("Precio"), "0.00")
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
End Sub

Sub ProcesarStockMax()
    Label1.Caption = Fg2.TextMatrix(Fg2.Row, 1)
    
    RST_Busq Rst, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckmin, " _
        & " alm_inventario.stckmax, alm_inventario.stckact, alm_inventario.tippro, [alm_inventario]![stckmax]-[alm_inventario]![stckact] AS diferencia, " _
        & " IIf([alm_inventario]![stckact]<=[alm_inventario]![stckmin],1,0) AS critico, alm_inventario.activo " _
        & " FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.stckmin) <> 0) " _
        & " And ((alm_inventario.tippro) = " & nulosn(lblIdItem.Caption) & ") And (([alm_inventario]![stckmax] - [alm_inventario]![stckact]) < 0) " _
        & " And ((IIf([alm_inventario]![stckact] <= [alm_inventario]![stckmin], 1, 0)) = 1) And ((alm_inventario.activo) = -1)) " _
        & " ORDER BY alm_inventario.descripcion", xCon

    Fg1.Cols = 7
    Fg1.Rows = 1
    Fg1.TextMatrix(0, 1) = "Codigo":      Fg1.ColWidth(1) = 1400:  Fg1.ColAlignment(1) = flexAlignLeftCenter
    Fg1.TextMatrix(0, 2) = "Descripcion": Fg1.ColWidth(2) = 4000:  Fg1.ColAlignment(2) = flexAlignLeftCenter
    Fg1.TextMatrix(0, 3) = "Uni. Med.":   Fg1.ColWidth(3) = 900:   Fg1.ColAlignment(3) = flexAlignCenterCenter
    Fg1.TextMatrix(0, 4) = "Stock Max.":  Fg1.ColWidth(4) = 1000:  Fg1.ColAlignment(4) = flexAlignRightCenter
    Fg1.TextMatrix(0, 5) = "Stock Act.":  Fg1.ColWidth(5) = 1000:  Fg1.ColAlignment(5) = flexAlignRightCenter
    Fg1.TextMatrix(0, 6) = "Diferencia":  Fg1.ColWidth(6) = 1000:  Fg1.ColAlignment(6) = flexAlignRightCenter
        
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("codpro")
            Fg1.TextMatrix(A, 2) = Rst("descripcion")
            Fg1.TextMatrix(A, 3) = Rst("abrev")
            Fg1.TextMatrix(A, 4) = Format(Rst("stckmax"), "0.00")
            Fg1.TextMatrix(A, 5) = Format(Rst("stckact"), "0.00")
            Fg1.TextMatrix(A, 6) = Format(Abs(Rst("diferencia")), "0.00")
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If

End Sub
Sub ProcesarStockMini()
    Label1.Caption = Fg2.TextMatrix(Fg2.Row, 1)
    RST_Busq Rst, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckmin, alm_inventario.stckmax, " _
        & " alm_inventario.stckact, alm_inventario.tippro, [alm_inventario]![stckact]-[alm_inventario]![stckmin] AS diferencia, " _
        & " IIf([alm_inventario]![stckact]<=[alm_inventario]![stckmin],1,0) AS critico, alm_inventario.activo FROM mae_unidades RIGHT JOIN alm_inventario " _
        & " ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.stckmin) <> 0) And ((alm_inventario.tippro) = " & nulosn(lblIdItem.Caption) & ") " _
        & " And ((IIf([alm_inventario]![stckact] <= [alm_inventario]![stckmin], 1, 0)) = 1) And ((alm_inventario.activo) = -1)) " _
        & " ORDER BY alm_inventario.descripcion", xCon

    Fg1.Cols = 7
    Fg1.Rows = 1
    Fg1.TextMatrix(0, 1) = "Codigo":      Fg1.ColWidth(1) = 1400:  Fg1.ColAlignment(1) = flexAlignLeftCenter
    Fg1.TextMatrix(0, 2) = "Descripcion": Fg1.ColWidth(2) = 4000:  Fg1.ColAlignment(2) = flexAlignLeftCenter
    Fg1.TextMatrix(0, 3) = "Uni. Med.":   Fg1.ColWidth(3) = 900:   Fg1.ColAlignment(3) = flexAlignCenterCenter
    Fg1.TextMatrix(0, 4) = "Stock Min.":  Fg1.ColWidth(4) = 1000:  Fg1.ColAlignment(4) = flexAlignRightCenter
    Fg1.TextMatrix(0, 5) = "Stock Act.":  Fg1.ColWidth(5) = 1000:  Fg1.ColAlignment(5) = flexAlignRightCenter
    Fg1.TextMatrix(0, 6) = "Diferencia":  Fg1.ColWidth(6) = 1000:  Fg1.ColAlignment(6) = flexAlignRightCenter
        
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("codpro")
            Fg1.TextMatrix(A, 2) = Rst("descripcion")
            Fg1.TextMatrix(A, 3) = Rst("abrev")
            Fg1.TextMatrix(A, 4) = Format(Rst("stckmin"), "0.00")
            Fg1.TextMatrix(A, 5) = Format(Rst("stckact"), "0.00")
            Fg1.TextMatrix(A, 6) = Format(Rst("diferencia"), "0.00")
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
End Sub

Private Sub Command2_Click()
    Dim xFrm As New eps_librerias.ImprimirFlexGrid
    
    xFrm.ImprimirFlex Fg1, UCase(Fg2.TextMatrix(Fg2.Row, 1)), "TIPO ITEMS  : " & UCase(txtTipIte.Text), NomEmp, xNumRuc
    Set xFrm = Nothing
End Sub

Private Sub Command3_Click()
    Dim xFun As New SGI2_funciones.formularios
    xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, UCase(Fg2.TextMatrix(Fg2.Row, 1)), "TIPO ITEM   : " & UCase(txtTipIte.Text), "", "ALMACEN"    ', Rst, ""
    Set xFun = Nothing
End Sub

Private Sub Command4_Click()
    Set Rst = Nothing
    Unload Me
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        txtTipIte.SetFocus
        lblIdItem.Caption = 3
        txtTipIte.Text = BUSCA_CODIGO(3, "id", "descripcion", "mae_tipoproducto", "N", xCon)
        
        Fg2.Rows = 0
        Fg2.Rows = 5
        Fg2.TextMatrix(0, 1) = "Lista de Precios y Stock"
        Fg2.TextMatrix(1, 1) = "Listar Stock"
        Fg2.TextMatrix(2, 1) = "Listar Inventario"
        Fg2.TextMatrix(3, 1) = "Items con Stock Minimo Critico"
        Fg2.TextMatrix(4, 1) = "Items con Stock Maximo Critico"
        Fg1.Rows = 1
        Fg1.Select 0, 0
        Command1_Click
    End If
End Sub

Private Sub Form_Load()
    Fg2.ColWidth(2) = 0
    SeEjecuto = False
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    EO.Left = 0
    EO.Top = 0
    CargaDatosEmpresa
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.BackColorSel = &H80&

    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H80&
    
End Sub

Private Sub Form_Resize()
    EO.Width = Me.Width - 120
    EO.Height = Me.Height - 400
End Sub

Private Sub txtTipIte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtTipIte_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipiTem_Click
    End If
End Sub
