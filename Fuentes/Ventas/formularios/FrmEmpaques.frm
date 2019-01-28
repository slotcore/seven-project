VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmEmpaques 
   Caption         =   "Ventas - Empaques para Despacho"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FrmEmpaques.frx":0000
   ScaleHeight     =   6750
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
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
      Height          =   855
      Left            =   3360
      TabIndex        =   4
      Top             =   0
      Width           =   8085
      Begin VB.CommandButton CmdImp 
         Height          =   630
         Left            =   5760
         Picture         =   "FrmEmpaques.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   165
         Width           =   660
      End
      Begin VB.CommandButton CmdExp 
         Height          =   630
         Left            =   6450
         Picture         =   "FrmEmpaques.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   165
         Width           =   660
      End
      Begin VB.CommandButton CmdBuscar 
         Height          =   630
         Left            =   5070
         Picture         =   "FrmEmpaques.frx":1256
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   165
         Width           =   660
      End
      Begin VB.CommandButton CmdSalir 
         Height          =   630
         Left            =   7200
         Picture         =   "FrmEmpaques.frx":1698
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   165
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Fecha de Emision de las Guias ]"
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
      Height          =   855
      Left            =   15
      TabIndex        =   1
      Top             =   0
      Width           =   3240
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Left            =   1020
         TabIndex        =   2
         Top             =   360
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
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   180
         Left            =   150
         TabIndex        =   3
         Top             =   390
         Width           =   615
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   3150
      Left            =   15
      TabIndex        =   0
      Top             =   885
      Width           =   11430
      _cx             =   20161
      _cy             =   5556
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmEmpaques.frx":19A2
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
   Begin VSFlex7Ctl.VSFlexGrid Fg2 
      Height          =   2430
      Left            =   15
      TabIndex        =   7
      Top             =   4320
      Width           =   11430
      _cx             =   20161
      _cy             =   4286
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmEmpaques.frx":1A5C
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
   Begin VB.Label Label2 
      Caption         =   "Clientes a Despachar"
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
      Height          =   195
      Left            =   45
      TabIndex        =   8
      Top             =   4110
      Width           =   2250
   End
End
Attribute VB_Name = "FrmEmpaques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Agregando As Boolean

Sub MuestraPedidos()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    
    RST_Busq Rst, "SELECT vta_guia.fecgiro, vta_guiadet.iditem, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, Sum(vta_guiadet.canpro) AS total " _
        & " FROM vta_guia LEFT JOIN ((vta_guiadet LEFT JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON vta_guiadet.idunimed = mae_unidades.id) ON vta_guia.id = vta_guiadet.idgui GROUP BY vta_guia.fecgiro, vta_guiadet.iditem, alm_inventario.codpro, " _
        & " alm_inventario.descripcion, mae_unidades.abrev HAVING (((vta_guia.fecgiro)=CDate('" & TxtFecha.Valor & "'))) ORDER BY alm_inventario.descripcion", xCon

    Fg1.Rows = 1
    Fg2.Rows = 1
    
    If Rst.RecordCount <> 0 Then
        Agregando = True
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("codpro")
            Fg1.TextMatrix(A, 2) = Rst("descripcion")
            Fg1.TextMatrix(A, 3) = Rst("abrev")
            Fg1.TextMatrix(A, 4) = Format(Rst("total"), FORMAT_MONTO)
            Fg1.TextMatrix(A, 5) = Rst("iditem")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        Agregando = False
        MuestraDetalle Fg1.TextMatrix(1, 5)
    Else
        MsgBox "No se han emitido guias en el periodo establecido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    Set Rst = Nothing
End Sub

Private Sub CmdBuscar_Click()
    MuestraPedidos
End Sub

Private Sub CmdExp_Click()
    Dim xFun As New SGI2_funciones.formularios
    Agregando = True
    xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "RESUMEN DE PRODUCTOS A REPARTIR", "Dia   : " & TxtFecha.Valor, "", "empaque.xls"
    Agregando = False
    Set xFun = Nothing
End Sub

Private Sub CmdImp_Click()
    Dim xFun As New SGI2_funciones.formularios
    Agregando = True
    xFun.Imprimir_x_VSFlexGrid Fg1, "RESUMEN DE PRODUCTOS A ENTREGAR", "", "Dia  : " & TxtFecha.Valor, True, True
    Agregando = False
    Set xFun = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Fg1_RowColChange()
    If Agregando = True Then Exit Sub
    MuestraDetalle Fg1.TextMatrix(Fg1.Row, 5)
End Sub

Private Sub Form_Load()
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg1.ColWidth(5) = 0
    Fg1.Rows = 1
    Fg2.Rows = 1
    Agregando = False
End Sub

Sub MuestraDetalle(IdProducto As Integer)
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq Rst, "SELECT mae_cliente.nombre, vta_puntoVenta.descripcion, [vta_guia]![numser]+'-'+[vta_guia]![numdoc] AS numdoc, alm_inventario.descripcion, " _
        & " vta_guiadet.canpro, vta_guia.fecgiro FROM (((vta_guia LEFT JOIN vta_guiadet ON vta_guia.id = vta_guiadet.idgui) LEFT JOIN mae_cliente " _
        & " ON vta_guia.idcli = mae_cliente.id) LEFT JOIN vta_puntoVenta ON vta_guia.idpunven = vta_puntoVenta.id) LEFT JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id " _
        & " WHERE (((vta_guia.fecgiro)=CDate('" & TxtFecha.Valor & "')) AND ((vta_guiadet.iditem)=" & IdProducto & "))", xCon

    Fg2.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = Rst("nombre")
            Fg2.TextMatrix(A, 2) = Rst("vta_puntoVenta.descripcion")
            Fg2.TextMatrix(A, 3) = Rst("numdoc")
            Fg2.TextMatrix(A, 4) = Format(Rst("canpro"), FORMAT_MONTO)
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub
