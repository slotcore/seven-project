VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmEmisionGuias 
   Caption         =   "Ventas - Emision Guias (Teleproceso)"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   12765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   915
      Left            =   3630
      TabIndex        =   0
      Top             =   3900
      Visible         =   0   'False
      Width           =   5910
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Procesando"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   330
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   5895
         X2              =   5880
         Y1              =   15
         Y2              =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   30
         X2              =   5925
         Y1              =   895
         Y2              =   895
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5895
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Guias de Remision"
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
         Left            =   105
         TabIndex        =   28
         Top             =   45
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   280
         Left            =   0
         Top             =   0
         Width           =   5850
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   7905
      Left            =   15
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   15
      Width           =   12750
      _cx             =   22490
      _cy             =   13944
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
      BorderWidth     =   4
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
      ResizeFonts     =   0   'False
      GridRows        =   3
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   2
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmEmisionGuias.frx":0000
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   705
         Left            =   60
         TabIndex        =   20
         Top             =   7140
         Width           =   12630
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   585
            Left            =   1065
            TabIndex        =   22
            Top             =   60
            Width           =   10485
            Begin VB.CommandButton CmdProcesar 
               Caption         =   "Crear Guias - Todos los Pedidos"
               Height          =   525
               Left            =   8715
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   30
               Width           =   1740
            End
            Begin VB.CommandButton CmdHojTra 
               Caption         =   "Hoja de Trabajo"
               Height          =   525
               Left            =   3510
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   30
               Width           =   1740
            End
            Begin VB.CommandButton CmdSalir 
               Caption         =   "Salir"
               Height          =   525
               Left            =   5265
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   30
               Width           =   1980
            End
            Begin VB.CommandButton CmdGenGuiaUno 
               Caption         =   "Crear Guias - Guia Actual"
               Height          =   525
               Left            =   15
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   30
               Width           =   1740
            End
            Begin VB.CommandButton CmeExp 
               Height          =   525
               Left            =   1785
               Picture         =   "FrmEmisionGuias.frx":004F
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   30
               Width           =   555
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1350
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   12630
         Begin VB.CommandButton CmdBusNumSer 
            Height          =   240
            Left            =   2145
            Picture         =   "FrmEmisionGuias.frx":0B59
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   60
            Width           =   225
         End
         Begin VB.CommandButton CmdEmpTra 
            Height          =   240
            Left            =   6915
            Picture         =   "FrmEmisionGuias.frx":0C8B
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   375
            Width           =   225
         End
         Begin VB.CommandButton CmdBusMotivo 
            Height          =   240
            Left            =   6915
            Picture         =   "FrmEmisionGuias.frx":0DBD
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   690
            Width           =   225
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Left            =   1260
            TabIndex        =   5
            Top             =   975
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
         Begin VB.TextBox TxtEmpTra 
            Height          =   300
            Left            =   1260
            TabIndex        =   8
            Text            =   "TxtEmpTra"
            Top             =   345
            Width           =   5910
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   3870
            TabIndex        =   10
            Text            =   "TxtNumDoc"
            Top             =   30
            Width           =   1605
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1260
            TabIndex        =   11
            Text            =   "TxtNumSer"
            Top             =   30
            Width           =   1140
         End
         Begin VB.TextBox TxtMotivo 
            Height          =   300
            Left            =   1260
            TabIndex        =   12
            Text            =   "TxtMotivo"
            Top             =   660
            Width           =   5910
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Serie  Guia"
            Height          =   195
            Left            =   15
            TabIndex        =   19
            Top             =   75
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Left            =   2670
            TabIndex        =   18
            Top             =   75
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Empresa Trans."
            Height          =   195
            Left            =   15
            TabIndex        =   17
            Top             =   375
            Width           =   1110
         End
         Begin VB.Label LblIdEmpTra 
            AutoSize        =   -1  'True
            Caption         =   "LblIdEmpTra"
            Height          =   195
            Left            =   7245
            TabIndex        =   16
            Top             =   390
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Motivo Trans."
            Height          =   195
            Left            =   15
            TabIndex        =   15
            Top             =   690
            Width           =   975
         End
         Begin VB.Label LblIdMotivo 
            AutoSize        =   -1  'True
            Caption         =   "LblIdMotivo"
            Height          =   195
            Left            =   7245
            TabIndex        =   14
            Top             =   705
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label Label7 
            Caption         =   "Fch. Emision"
            Height          =   195
            Left            =   15
            TabIndex        =   13
            Top             =   1005
            Width           =   1020
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   5610
         Left            =   60
         TabIndex        =   21
         Top             =   1470
         Width           =   12630
         _cx             =   22278
         _cy             =   9895
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
         BackColor       =   14745342
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   8388608
         BackColorSel    =   128
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14745342
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
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmEmisionGuias.frx":0EEF
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
Attribute VB_Name = "FrmEmisionGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstPunVen As New ADODB.Recordset
Dim ocurrioError As Boolean
Dim SeEjecuto As Boolean
Dim cSQL As String

Private Sub CmdBusMotivo_Click()
    'Dim xform As New eps_librerias.FormBuscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Cliente":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":   xCampos(1, 1) = "id":            xCampos(1, 2) = "1500":    xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_mottra.id, mae_mottra.descripcion From mae_mottra " _
        & " ORDER BY mae_mottra.descripcion"

    xform.Titulo = "Buscando Motivos de Transporte"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtMotivo.Text = NulosC(xRs("descripcion"))
        LblIdMotivo.Caption = NulosN(xRs("id"))
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusNumSer_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Almacen":     xCampos(0, 1) = "desalm":     xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Documento":   xCampos(1, 1) = "desdoc":     xCampos(1, 2) = "2000":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Serie":    xCampos(2, 1) = "numser":     xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"
    
    xform.SQLCad = "SELECT alm_numseries.id, alm_almacenes.descripcion AS desalm, mae_documento.descripcion AS desdoc, alm_numseries.numser " _
        & " FROM (alm_numseries LEFT JOIN mae_documento ON alm_numseries.idtipdoc = mae_documento.id) LEFT JOIN alm_almacenes " _
        & " ON alm_numseries.idalm = alm_almacenes.id WHERE (((alm_numseries.idtipdoc)=9))"

    xform.Titulo = "Buscando Series de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtNumSer.Text = NulosC(xRs("numser"))
        TxtNumDoc.Text = HallaNumGuia(Trim(TxtNumSer.Text), xCon)
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdEmpTra_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nombre":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":   xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_emptra.id, mae_emptra.nombre, mae_emptra.numruc " _
        & " From mae_emptra ORDER BY mae_emptra.nombre"

    xform.Titulo = "Buscando Empresa de Transporte"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtEmpTra.Text = NulosC(xRs("nombre"))
        LblIdEmpTra.Caption = NulosN(xRs("id"))
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdHojTra_Click()
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    Dim B, C As Integer
    Dim xFila As Integer
    Dim xCad As String
    
    'Printer.ScaleMode = vbCharacters
    Printer.ScaleMode = 6
    Printer.FontName = "Courier New"
    Printer.FontSize = 9
    'printer.PaperSize=
    xFila = 6
    For A = 1 To Fg1.Rows - 1
        'Printer.Font = "Curier New"
        Printer.CurrentX = 20: Printer.CurrentY = xFila: Printer.Print "Cliente        : " + Fg1.TextMatrix(A, 11)
        
        'buscamos todas las ordenes de compra
'        RST_Busq Rst, "SELECT pedidos.numcen, pedidos.fchemi, pedidos.fchent From pedidos WHERE (((pedidos.idpunvecli)=" & Val(Fg1.TextMatrix(A, 7)) & ") " _
'            & " AND ((pedidos.anulado)=0) AND ((pedidos.proceso)=0))", xCon

        RST_Busq Rst, "SELECT ped_pedido.numcen, ped_pedido.fchemi, ped_pedido.fchent From ped_pedido WHERE (((ped_pedido.idpunvecli)=" & Val(Fg1.TextMatrix(A, 7)) & ") " _
            & " AND ((ped_pedido.anulado)=0) AND ((ped_pedido.proceso)=0))", xCon

        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            xCad = ""
            For C = 1 To Rst.RecordCount
                xCad = xCad + Right(Rst("numcen"), 11)
                Rst.MoveNext
                
                If Rst.EOF = True Then
                    Rst.MovePrevious
                    'xFchGir = Rst("fchemi")
                    'xFchEnt = Rst("fchent")
                    Exit For
                End If
                xCad = xCad + ", "
            Next C
        End If
        
        Printer.CurrentX = 110: Printer.CurrentY = xFila: Printer.Print "Orden Compra  : " + Trim(xCad)
        
        xFila = xFila + 4
        Printer.CurrentX = 20:  Printer.CurrentY = xFila: Printer.Print "Punto de Venta : " + Trim(Fg1.TextMatrix(A, 2))
        Printer.CurrentX = 110: Printer.CurrentY = xFila: Printer.Print "Fch. Emision : " + Trim(Fg1.TextMatrix(A, 8))
        Printer.CurrentX = 160: Printer.CurrentY = xFila: Printer.Print "Fch. Entrega : " + Trim(Fg1.TextMatrix(A, 9))
        xFila = xFila + 4
        Printer.CurrentX = 20: Printer.CurrentY = xFila:
        Printer.Print "================================================================================================"
        xFila = xFila + 4
        Printer.CurrentX = 20: Printer.CurrentY = xFila:
        Printer.Print "COD. PROD.        DESCRIPCION                                   UNIDAD   CANTIDAD"
        xFila = xFila + 4
        Printer.CurrentX = 20: Printer.CurrentY = xFila:
        Printer.Print "================================================================================================"
        xFila = xFila + 4
                      'xxxxxxxxxxxxxxxx  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx   xxxx   XXXXX.XX
                      '123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
                      '         1         2         3         4         5
        RST_Busq Rst, "SELECT ped_pedidodet.codpro AS codcen, mae_productoscen.codpro AS codsgi, mae_producto.Descripcion, " _
            & " MAE_Unid_Med.DescAbrevia, Sum(ped_pedidodet.canpro) AS SumaDecanpro FROM ((cenproductos RIGHT JOIN " _
            & " (ped_pedido LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) ON cenproductos.codcen = ped_pedidodet.codpro) " _
            & " LEFT JOIN MAE_Producto ON cenproductos.codpro = MAE_Producto.Cod_Item) LEFT JOIN MAE_Unid_Med ON " _
            & " MAE_Producto.Cod_Unidad = MAE_Unid_Med.Cod_Unidad GROUP BY ped_pedidodet.codpro, cenproductos.codpro, " _
            & " MAE_Producto.Descripcion, MAE_Unid_Med.DescAbrevia, ped_pedido.idpunvecli, ped_pedido.anulado, ped_pedido.proceso " _
            & " Having (((ped_pedido.idpunvecli) = " & Val(Fg1.TextMatrix(A, 7)) & ") And ((ped_pedido.anulado) = 0) And ((ped_pedido.proceso) = 0)) " _
            & " ORDER BY MAE_Producto.Descripcion", xCon
        
        'xFila = 51
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            
            For B = 1 To Rst.RecordCount
                Printer.CurrentX = 20:  Printer.CurrentY = xFila: Printer.Print Rst("codsgi")
                Printer.CurrentX = 50:  Printer.CurrentY = xFila: Printer.Print Rst("descripcion")
                Printer.CurrentX = 130: Printer.CurrentY = xFila: Printer.Print Rst("descabrevia")
                Printer.CurrentX = 150: Printer.CurrentY = xFila: Printer.Print Rst("sumadecanpro")
                
                Rst.MoveNext
                xFila = xFila + 4
                If Rst.EOF = True Then
                    xFila = xFila + 4
                    Exit For
                End If
                If xFila >= 270 Then
                    Printer.NewPage
                    xFila = 6
                End If
            Next B
        End If
        
        If xFila >= 270 Then
            Printer.NewPage
            xFila = 6
        End If
    Next A
    
    Printer.EndDoc
    MsgBox "El documento se imprimio con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Sub

'*****************************************************************************************************
'* Nombre           : CmdProcesar_Click
'* Tipo             : SUB
'* Descripcion      : Genera todas las Guias
'* Parametros       :
'* Devuelve         :
'* Creado por       :
'* Modificado       : Jose Chacon 26/04/2011
'*****************************************************************************************************
Private Sub CmdProcesar_Click()
    Dim A As Integer
    Dim xNumGui As Double
    Dim Fila As Integer
    
    ProgressBar1.Max = Fg1.Rows - 1
    Frame2.Visible = True
    
    xNumGui = Val(TxtNumDoc.Text)
    Fila = 1
    For A = 1 To Fg1.Rows - 1
        ProgressBar1.Value = A
        Frame2.Refresh
        TxtNumDoc.Text = Format(xNumGui, "0000000000")
        GenerarUnaGuia Fg1, Fila
        If ocurrioError Then Fila = Fila + 1 Else xNumGui = xNumGui + 1
    Next A
    TxtNumDoc.Text = Format(xNumGui, "0000000000")
    Frame2.Visible = False
    
    MsgBox "Las guias se generaron con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstPunVen = Nothing
    ' Si no hay mas registros que procesar se cierra la ventana
    If Fg1.Rows = Fg1.FrozenRows Then Unload Me
End Sub

'*****************************************************************************************************
'* Nombre           : CmdGenGuiaUno_Click
'* Tipo             : SUB
'* Descripcion      : Genera solo una guia
'* Parametros       :
'* Devuelve         :
'* Creado por       :
'* Modificado       : Jose Chacon 27/04/2011
'*****************************************************************************************************
Private Sub CmdGenGuiaUno_Click()
    Dim xNumGui As Double
    xNumGui = Val(TxtNumDoc.Text)
    GenerarUnaGuia Fg1, Fg1.Row
    If Not ocurrioError Then
        xNumGui = xNumGui + 1
        TxtNumDoc.Text = Format(xNumGui, "0000000000")
        MsgBox "La guia se generó con èxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

'Private Sub CmdProcesar_Click2()
'    Dim A As Integer
'    Dim xNumGui As Double
'
'    Frame2.Left = 2250
'    Frame2.Top = 2790
'    ProgressBar1.Max = Fg1.Rows - 1
'    Frame2.Visible = True
'
'    xNumGui = Val(TxtNumDoc.Text)
'
'    For A = 1 To Fg1.Rows - 1
'        ProgressBar1.Value = A
'        Frame2.Refresh
'        If Fg1.TextMatrix(A, 3) = "" And Fg1.TextMatrix(A, 4) = "" Then
'
'        Else
'            TxtNumDoc.Text = Format(xNumGui, "0000000000")
'            'GenerarGuia A
'            GenerarUnaGuia Fg1, A
'            xNumGui = xNumGui + 1
'        End If
'    Next A
'    Frame2.Visible = False
'
'    MsgBox "Las guias se generaron con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'    Set RstPunVen = Nothing
'    Unload Me
'End Sub

Private Sub CmdSalir_Click()
    Set RstPunVen = Nothing
    Unload Me
End Sub

Private Sub CmeExp_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No hay pedidos para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    Dim xFun As New SGI2_funciones.formularios
    xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Pedidos", "", "", "pedidos.xls"
    Set xFun = Nothing
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If NulosC(Fg1.TextMatrix(Fg1.Row, 10)) <> "" Then
        MsgBox "No puede modificar los datos de un pedido procesado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    If Col = 3 Then
        Dim xCampos(4, 4) As String
        xCampos(0, 0) = "Chofer":      xCampos(0, 1) = "apenom":       xCampos(0, 2) = "4500":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Nº Brevete":  xCampos(1, 1) = "numbre":       xCampos(1, 2) = "1200":    xCampos(1, 3) = "C"
        xCampos(2, 0) = "Marca":       xCampos(2, 1) = "marca":        xCampos(2, 2) = "1450":    xCampos(2, 3) = "C"
        xCampos(3, 0) = "Nº Placa":    xCampos(3, 1) = "numpla":       xCampos(3, 2) = "1000":    xCampos(3, 3) = "C"
        
        xform.SQLCad = "SELECT mae_chofer.id, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+ +', '+[pla_empleados]![nom] AS apenom, mae_chofer.numbre, mae_vehiculo.marca, " _
            & " mae_vehiculo.numpla,  mae_chofer.idvehiculo FROM mae_vehiculo RIGHT JOIN (pla_empleados RIGHT JOIN mae_chofer ON pla_empleados.id = mae_chofer.idper) " _
            & " ON mae_vehiculo.id = mae_chofer.idvehiculo ORDER BY UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+ +', '+[pla_empleados]![nom] "

        xform.Titulo = "Buscando Chofer"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "apenom"
        xform.CampoBusca = "apenom"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Fg1.TextMatrix(Fg1.Row, 3) = NulosC(xRs("apenom"))
            Fg1.TextMatrix(Fg1.Row, 4) = NulosC(xRs("numpla"))
            Fg1.TextMatrix(Fg1.Row, 5) = NulosN(xRs("id"))
            Fg1.TextMatrix(Fg1.Row, 6) = NulosN(xRs("idvehiculo"))
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 4 Then
        Dim xCampos2(2, 4) As String
        xCampos2(0, 0) = "Nº Placa":    xCampos2(0, 1) = "numpla":    xCampos2(0, 2) = "2000":    xCampos2(0, 3) = "C"
        xCampos2(1, 0) = "Marca":       xCampos2(1, 1) = "marca":     xCampos2(1, 2) = "2000":    xCampos2(1, 3) = "C"
        
        xform.SQLCad = "SELECT mae_vehiculo.marca, mae_vehiculo.numpla, mae_vehiculo.id From mae_vehiculo ORDER BY mae_vehiculo.marca"
        
        xform.Titulo = "Buscando Unidades de Transporte"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "numpla"
        xform.CampoBusca = "numpla"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos2)
        
        If xRs.State = 1 Then
            Fg1.TextMatrix(Fg1.Row, 4) = NulosC(xRs("numpla"))
            Fg1.TextMatrix(Fg1.Row, 6) = NulosN(xRs("id"))
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_EnterCell()
    If Fg1.Col = 3 Or Fg1.Col = 4 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Fg1.Col = 3 Then
        If KeyCode = 46 Then
            Fg1.TextMatrix(Fg1.Row, 3) = ""
            Fg1.TextMatrix(Fg1.Row, 5) = ""
        End If
    End If
    
    If Fg1.Col = 4 Then
        If KeyCode = 46 Then
            Fg1.TextMatrix(Fg1.Row, 4) = ""
            Fg1.TextMatrix(Fg1.Row, 6) = ""
        End If
    End If
End Sub

Private Sub Fg1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = 3 Then
        
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Form_Activate
'* Tipo             : SUB
'* Descripcion      :
'* Parametros       :
'* Devuelve         :
'* Creado por       :
'* Modificado       : Jose Chacon 13/12/2010
'                       Modificacion en la consulta para que utilice la nueva tabla
'                       Modificacion de la consulta debido a errores de sintaxis
'*****************************************************************************************************
Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        pCargarGrid
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    iniciarCampos
End Sub


Private Sub pCargarGrid()
    Dim A As Integer
    
    cSQL = "SELECT ped_pedido.idpunvecli, ped_pedido.codcen AS codigo, vta_puntoVenta.descripcion, ped_pedido.fchemi, ped_pedido.fchent, mae_cliente.nombre AS nomcli, vta_puntoVenta.dir, vta_puntoVenta.idcli, ped_pedido.id AS idpedido, ped_pedido.numcen AS numorden " _
        + vbCr + "FROM (ped_pedido LEFT JOIN vta_puntoVenta ON ped_pedido.idpunvecli = vta_puntoVenta.id) LEFT JOIN mae_cliente ON vta_puntoVenta.idcli = mae_cliente.id " _
        + vbCr + "GROUP BY ped_pedido.idpunvecli, ped_pedido.codcen, vta_puntoVenta.descripcion, ped_pedido.fchemi, ped_pedido.fchent, mae_cliente.nombre, vta_puntoVenta.dir, vta_puntoVenta.idcli, ped_pedido.id, ped_pedido.numcen, ped_pedido.proceso, ped_pedido.anulado, ped_pedido.idprocped  " _
        + vbCr + "HAVING (((ped_pedido.proceso)=0) AND ((ped_pedido.anulado)=0) AND ((ped_pedido.idprocped)=2));"
        
    RST_Busq RstPunVen, cSQL, xCon

    If RstPunVen.RecordCount = 0 Then
        MsgBox "No se han encontrado pedidos levantados", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstPunVen = Nothing
        Unload Me
        Exit Sub
    End If
    RstPunVen.MoveFirst
    
    For A = 1 To RstPunVen.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(RstPunVen("codigo"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstPunVen("descripcion"))
        'Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstPunVen("apenom"))
        'Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstPunVen("placa"))
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosN(RstPunVen("idpunvecli"))
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(RstPunVen("fchemi"), "dd/mm/yy")
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(RstPunVen("fchent"), "dd/mm/yy")
        'Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(RstPunVen("estado"))
        Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosC(RstPunVen("nomcli"))
        Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosC(RstPunVen("idcli"))
        Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosC(RstPunVen("dir"))
        Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosC(RstPunVen("idpedido"))
        Fg1.TextMatrix(Fg1.Rows - 1, 15) = NulosC(RstPunVen("numorden"))
        RstPunVen.MoveNext
        If RstPunVen.EOF = True Then Exit For
    Next A
    
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT mae_emptra.id, mae_emptra.nombre, mae_emptra.default From mae_emptra " _
        & "WHERE (((mae_emptra.default)=-1))", xCon
        
    LblIdEmpTra.Caption = Rst("id")
    TxtEmpTra.Text = Rst("nombre")
    
    TxtMotivo.Text = "Venta"
    LblIdMotivo.Caption = "1"
    
    TxtNumSer.Text = "0001"
    TxtNumDoc.Text = HallaNumGuia(Trim(TxtNumSer.Text), xCon)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width <= 10860 Or Me.Height <= 7000 Then
        If Me.Width <= 10860 Then Me.Width = 10860
        If Me.Height <= 7000 Then Me.Height = 7000
        ElasticOne1.Width = Me.Width - 90
        ElasticOne1.Height = Me.Height - 380
        Frame1.Left = ((ElasticOne1.Width - Frame1.Width) / 2) - 60
        Exit Sub
    End If
    ElasticOne1.Width = Me.Width - 90
    ElasticOne1.Height = Me.Height - 380
    Frame1.Left = ((ElasticOne1.Width - Frame1.Width) / 2) - 60
End Sub

'*****************************************************************************************************
'* Nombre           : iniciarCampos
'* Tipo             : SUB
'* Descripcion      : inicializa los valores del grid
'* Parametros       :
'* Devuelve         :
'* Creado por       : Jose Chacon
'* Modificado       : Jose Chacon 27/04/2011
'*                      Se quita la propiedad de cambiar de lugar las columnas
'*****************************************************************************************************
Private Sub iniciarCampos()
    Fg1.Rows = 1
    
    Fg1.ColWidth(5) = 0
    Fg1.ColWidth(6) = 0
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(12) = 0
    Fg1.ColWidth(13) = 0
    Fg1.ColWidth(14) = 0
    
    Fg1.ColComboList(3) = "|..."
    Fg1.ColComboList(4) = "|..."
    
    Frame3.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    
    Me.Width = 14780
    
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.ExplorerBar = flexExSortShow
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ForeColorSel = &H0&
    Fg1.BackColorSel = &HC0E0FF
    
    Fg1.Select 0, 1, 0, Fg1.Cols - 1
    Fg1.FillStyle = flexFillRepeat
    Fg1.CellForeColor = &H800000
    Fg1.CellFontBold = True
    
    Fg1.GridLines = flexGridInset
    Fg1.RowHeight(0) = 300
    Fg1.ColWidth(0) = 0
    Fg1.ColWidth(1) = 1400
    Fg1.ColWidth(2) = 3750
    Fg1.ColWidth(8) = 1000
    Fg1.ColWidth(9) = 1000
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 2500
    Fg1.ColWidth(15) = 2150
End Sub

Private Sub TxtEmpTra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtMotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Sub GenerarUnaGuia(ByRef fg As VSFlexGrid, Fila As Integer)
    '===================================================================================================
    'Creado :   /  /  Por: Enrique Pollongo
    'Propósito: Generar una guia de remision por pedido seleccionado
    '
    'Entradas:  Ninguna
    '
    'Resultados: Registro de la guia en vta_guia
    '
    'Nota:       1.- Por defecto muestra la fecha actual para la fecha de emisión de la guia
    '            2.- Indicar el chofer (Por defecto muestra el vehículo)
    '            3.- Indicar el vehiculo (Si no es el correcto)
    '            4.- Clic en boton [Crear Guias - Guia Actual]

    'Modificado: 27/11/10 Por: Johan Castro
    '           Grabar el código de la guia en el pedido para hacer la vinculación entre las tablas
    '           vta_pedido.idtipdoc=1(1 Guias, 2=Ventas); vta_pedido.iddocven=vta_guia.id
    '           15/12/10 Por: Jose Chacon
    '           Modificar las consultas de la tabla vta_pedido,vta_pedidodet por ped_pedido,ped_pedidodet
    '           19/01/11 Por Johan Castro
    '           Cambiar el tipo de dato a variable xId a Double antes Integer
    '           Agregar lineas de codigo para registrar el historial de guia(nuevo registro)
    '           Linea de codigo para registrar historial de pedido esta deshabilitado
    '           26/04/11 Por Jose Chacon
    '           Se generaliza para todo tipo de generacion de guia individual o grupal
    '           Se elimina el parametro contador de tipo entero
    '===================================================================================================

    Dim RstCab As New ADODB.Recordset '--Cabecera de guias
    Dim RstDet As New ADODB.Recordset '--Detalle de guias
    Dim RstPed As New ADODB.Recordset '--Detalle de pedido
    Dim Rst As New ADODB.Recordset 'Cabecera de pedido
   
    Dim A, B, C As Integer
    Dim xId As Double
    Dim xCad, xFchGir, xFchEnt  As String
    Dim xNumGui As String
    With fg
        If .Rows = 1 Then Exit Sub
        
        If .TextMatrix(Fila, 3) = "" Then
            MsgBox "No ha especificado el chofer para la entrega del pedido Nº " + Trim(.TextMatrix(Fila, 1)), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            ocurrioError = True
            Exit Sub
        End If
        
        If .TextMatrix(Fila, 4) = "" Then
            MsgBox "No ha especificado la unidad de transporte para la entrega del pedido Nº " + Trim(.TextMatrix(Fila, 1)), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            ocurrioError = True
            Exit Sub
        End If
        
        If UCase(NulosC(.TextMatrix(Fila, 10))) = "PROCESADO" Then
            MsgBox "Ya se procesaron guias para el punto de venta especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            ocurrioError = True
            Exit Sub
        End If
        
        If Fila < 0 Then
            MsgBox "Seleccione un registro para grabar", vbInformation, xTitulo
            ocurrioError = True
            Exit Sub
        End If
        
        On Error GoTo LaCague
        
        xCon.BeginTrans
        xNumGui = TxtNumDoc.Text
            
        RST_Busq RstCab, "SELECT TOP 1 * FROM vta_guia", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM vta_guiadet", xCon
        
        xId = HallaCodigoTabla("vta_guia", xCon, "id")
        
        RstCab.AddNew
        RstCab("id") = xId
        RstCab("tipdoc") = 9
        RstCab("numser") = Format(NulosC(TxtNumSer.Text), "0000")
        RstCab("numdoc") = Format(xNumGui, "0000000000")
        RstCab("idcli") = NulosN(.TextMatrix(Fila, 12))
        RstCab("dircli") = .TextMatrix(Fila, 13)
        RstCab("fecgiro") = CDate(TxtFecha.Valor)
        RstCab("idpunven") = NulosN(.TextMatrix(Fila, 7))
        RstCab("idmottra") = NulosN(LblIdEmpTra.Caption)
        RstCab("idcho") = NulosN(.TextMatrix(Fila, 5))
        RstCab("idemptra") = NulosN(LblIdEmpTra.Caption)
        RstCab("idveh") = NulosN(.TextMatrix(Fila, 6))
        
        cSQL = "SELECT ped_pedido.* " _
            + vbCr + "From ped_pedido " _
            + vbCr + "WHERE (((ped_pedido.idpunvecli)=" & NulosN(.TextMatrix(Fila, 7)) & ") AND ((ped_pedido.anulado)=0) AND ((ped_pedido.proceso)=0) AND ((ped_pedido.id)=" & NulosN(.TextMatrix(Fila, 14)) & "))"
    
        RST_Busq Rst, cSQL, xCon
        
        cSQL = "SELECT alm_inventario.id, ped_pedidodet.idpeddet, mae_productoscen.codcen, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.canpro, alm_inventario.idunimed, ped_pedido.proceso " _
            + vbCr + "FROM (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) RIGHT JOIN (ped_pedido LEFT JOIN (mae_productoscen RIGHT JOIN ped_pedidodet ON mae_productoscen.codcen = ped_pedidodet.codpro) ON ped_pedido.id = ped_pedidodet.idped) ON alm_inventario.id = mae_productoscen.iditem " _
            + vbCr + "Where (((ped_pedido.idpunvecli) = " & NulosN(.TextMatrix(Fila, 7)) & ")) " _
            + vbCr + "GROUP BY alm_inventario.id, ped_pedidodet.idpeddet, mae_productoscen.codcen, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.canpro, alm_inventario.idunimed, ped_pedido.proceso, ped_pedido.id " _
            + vbCr + "HAVING (((ped_pedido.proceso)=0) AND ((ped_pedido.id)=" & NulosN(.TextMatrix(Fila, 14)) & "))"
    
        RST_Busq RstPed, cSQL, xCon
    
        If RstPed.RecordCount = 0 Then
            MsgBox "No se han especificado items en el pedido del punto de venta  " + .TextMatrix(Fila, 2) + Chr(13) _
                & " del cliente  " + .TextMatrix(Fila, 11), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set RstPed = Nothing
            Exit Sub
        End If
        
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            xCad = ""
            For C = 1 To Rst.RecordCount
                xCad = xCad + Left(Right(Rst("numcen"), 11), 10)
                '--indicar que se procesa el pedido
                Rst("proceso") = -1
                'hacer la vinculacion entre pedido y guias  1=Guias, 2=Factura
                Rst("idtipdoc") = 1
                Rst("iddocven") = xId
                
                Rst.Update
                Rst.MoveNext
                If Rst.EOF = True Then
                    Rst.MovePrevious
                    xFchGir = Rst("fchemi")
                    xFchEnt = Rst("fchent")
                    Exit For
                End If
                xCad = xCad + ", "
            Next C
        End If
        RstCab("numordcom") = Mid(Trim(xCad), 1, 50)
        RstCab("fchemiord") = CDate(xFchGir)
        RstCab("fchentord") = CDate(xFchEnt)
        RstCab("dirpunpar") = ""
        RstCab("dirpunlle") = Trim(.TextMatrix(Fila, 13))
        RstCab("tippro") = "3"
        RstCab("idtipdocref") = 5
        RstCab.Update
        
        'GRABAMOS EL DETALLE DE LA GUIA
        RstPed.MoveFirst
        For A = 1 To RstPed.RecordCount
            RstDet.AddNew
            RstDet("idgui") = xId
            RstDet("iditem") = RstPed("id")
            RstDet("idunimed") = RstPed("idunimed")
            RstDet("canpro") = RstPed("canpro")
            RstDet("iddocref") = RstPed("idpeddet")
            RstDet.Update
            ' Se actualiza la cantidad entregada del producto en el pedido
            xCon.Execute "UPDATE ped_pedidodet SET ped_pedidodet.canproent = " & RstPed("canpro") & " WHERE (((ped_pedidodet.idpeddet)=" & RstPed("idpeddet") & "))"
            
            RstPed.MoveNext
            If RstPed.EOF = True Then Exit For
        Next A
        
        Set Rst = Nothing
        Set RstPed = Nothing
        
        'grabamos el movimiento en la tabla var_edicion para guias
        GrabarOperacion xIdUsuario, 17, 1, Time, Time, Date, xCon, xId
        
        'Se elimina la fila Procesada
        .RemoveItem Fila
        xCon.CommitTrans
    End With
    ocurrioError = False
    Set RstCab = Nothing
    Set RstDet = Nothing
    Exit Sub
LaCague:
    ocurrioError = True
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Exit Sub
End Sub

'Sub GenerarUnaGuia2()
'    '===================================================================================================
'    'Creado :   /  /  Por: Enrique Pollongo
'    'Propósito: Generar una guia de remision por pedido seleccionado
'    '
'    'Entradas:  Ninguna
'    '
'    'Resultados: Registro de la guia en vta_guia
'    '
'    'Nota:       1.- Por defecto muestra la fecha actual para la fecha de emisión de la guia
'    '            2.- Indicar el chofer (Por defecto muestra el vehículo)
'    '            3.- Indicar el vehiculo (Si no es el correcto)
'    '            4.- Clic en boton [Crear Guias - Guia Actual]
'
'    'Modificado: 27/11/10 Por: Johan Castro
'    '           Grabar el código de la guia en el pedido para hacer la vinculación entre las tablas
'    '           vta_pedido.idtipdoc=1(1 Guias, 2=Ventas); vta_pedido.iddocven=vta_guia.id
'    '           15/12/10 Por: Jose Chacon
'    '           Modificar las consultas de la tabla vta_pedido,vta_pedidodet por ped_pedido,ped_pedidodet
'    '           19/01/11 Por Johan Castro
'    '           Cambiar el tipo de dato a variable xId a Double antes Integer
'    '           Agregar lineas de codigo para registrar el historial de guia(nuevo registro)
'    '           Linea de codigo para registrar historial de pedido esta deshabilitado
'    '===================================================================================================
'    Dim RstCab As New ADODB.Recordset '--Cabecera de guias
'    Dim RstDet As New ADODB.Recordset '--Detalle de guias
'    Dim RstPed As New ADODB.Recordset '--Detalle de pedido
'    Dim Rst As New ADODB.Recordset 'Cabecera de pedido
'
'    Dim A, B, C As Integer
'    Dim xId As Double
'    Dim xCad, xFchGir, xFchEnt  As String
'    Dim xNumGui As String
'
'    If Fg1.Row < 0 Then
'        MsgBox "Seleccione un registro para grabar", vbInformation, xTitulo
'        Exit Sub
'    End If
'
'    On Error GoTo LaCague
'
'    xCon.BeginTrans
'    xNumGui = TxtNumDoc.Text
'
'    RST_Busq RstCab, "SELECT TOP 1 * FROM vta_guia", xCon
'    RST_Busq RstDet, "SELECT TOP 1 * FROM vta_guiadet", xCon
'
'    xId = HallaCodigoTabla("vta_guia", xCon, "id")
'
'    RstCab.AddNew
'    RstCab("id") = xId
'    RstCab("tipdoc") = 9
'    RstCab("numser") = Format(NulosC(TxtNumSer.Text), "0000")
'    RstCab("numdoc") = xNumGui
'    RstCab("idcli") = NulosN(Fg1.TextMatrix(Fg1.Row, 12))
'    RstCab("dircli") = Fg1.TextMatrix(Fg1.Row, 13)
'    RstCab("fecgiro") = CDate(TxtFecha.Valor)
'    RstCab("idpunven") = NulosN(Fg1.TextMatrix(Fg1.Row, 7))
'    RstCab("idmottra") = NulosN(LblIdEmpTra.Caption)
'    RstCab("idcho") = NulosN(Fg1.TextMatrix(Fg1.Row, 5))
'    RstCab("idemptra") = NulosN(LblIdEmpTra.Caption)
'    RstCab("idveh") = NulosN(Fg1.TextMatrix(Fg1.Row, 6))
'
'    cSQL = "SELECT ped_pedido.* " _
'        + vbCr + "From ped_pedido " _
'        + vbCr + "WHERE (((ped_pedido.idpunvecli)=" & NulosN(Fg1.TextMatrix(Fg1.Row, 7)) & ") AND ((ped_pedido.anulado)=0) AND ((ped_pedido.proceso)=0) AND ((ped_pedido.id)=" & NulosN(Fg1.TextMatrix(Fg1.Row, 14)) & "))"
'
'    RST_Busq Rst, cSQL, xCon
'
'    cSQL = "SELECT alm_inventario.id, ped_pedidodet.idpeddet, mae_productoscen.codcen, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.canpro, alm_inventario.idunimed, ped_pedido.proceso " _
'        + vbCr + "FROM (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) RIGHT JOIN (ped_pedido LEFT JOIN (mae_productoscen RIGHT JOIN ped_pedidodet ON mae_productoscen.codcen = ped_pedidodet.codpro) ON ped_pedido.id = ped_pedidodet.idped) ON alm_inventario.id = mae_productoscen.iditem " _
'        + vbCr + "Where (((ped_pedido.idpunvecli) = " & NulosN(Fg1.TextMatrix(Fg1.Row, 7)) & ")) " _
'        + vbCr + "GROUP BY alm_inventario.id, ped_pedidodet.idpeddet, mae_productoscen.codcen, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.canpro, alm_inventario.idunimed, ped_pedido.proceso, ped_pedido.id " _
'        + vbCr + "HAVING (((ped_pedido.proceso)=0) AND ((ped_pedido.id)=" & NulosN(Fg1.TextMatrix(Fg1.Row, 14)) & "))"
'
'    RST_Busq RstPed, cSQL, xCon
'
'    If RstPed.RecordCount = 0 Then
'        MsgBox "No se han especificado items en el pedido del punto de venta  " + Fg1.TextMatrix(Fg1.Row, 2) + Chr(13) _
'            & " del cliente  " + Fg1.TextMatrix(Fg1.Row, 11), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Set RstPed = Nothing
'        Exit Sub
'    End If
'
'    If Rst.RecordCount <> 0 Then
'        Rst.MoveFirst
'        xCad = ""
'        For C = 1 To Rst.RecordCount
'            xCad = xCad + Left(Right(Rst("numcen"), 11), 10)
'            '--indicar que se procesa el pedido
'            Rst("proceso") = -1
'            'hacer la vinculacion entre pedido y guias  1=Guias, 2=Factura
'            Rst("idtipdoc") = 1
'            Rst("iddocven") = xId
'
'            Rst.Update
'            Rst.MoveNext
'            If Rst.EOF = True Then
'                Rst.MovePrevious
'                xFchGir = Rst("fchemi")
'                xFchEnt = Rst("fchent")
'                Exit For
'            End If
'            xCad = xCad + ", "
'        Next C
'    End If
'    RstCab("numordcom") = Mid(Trim(xCad), 1, 50)
'    RstCab("fchemiord") = CDate(xFchGir)
'    RstCab("fchentord") = CDate(xFchEnt)
'    RstCab("dirpunpar") = ""
'    RstCab("dirpunlle") = Trim(Fg1.TextMatrix(Fg1.Row, 13))
'    RstCab("tippro") = "3"
'    RstCab("idtipdocref") = 5
'    RstCab.Update
'
'    'GRABAMOS EL DETALLE DE LA GUIA
'    RstPed.MoveFirst
'    For A = 1 To RstPed.RecordCount
'        RstDet.AddNew
'        RstDet("idgui") = xId
'        RstDet("iditem") = RstPed("id")
'        RstDet("idunimed") = RstPed("idunimed")
'        RstDet("canpro") = RstPed("canpro")
'        RstDet("iddocref") = RstPed("idpeddet")
'        RstDet.Update
'        ' Se actualiza la cantidad entregada del producto en el pedido
'        xCon.Execute "UPDATE ped_pedidodet SET ped_pedidodet.canproent = " & RstPed("canpro") & " WHERE (((ped_pedidodet.idpeddet)=" & RstPed("idpeddet") & "))"
'
'        RstPed.MoveNext
'        If RstPed.EOF = True Then Exit For
'    Next A
'
'    Set Rst = Nothing
'    Set RstPed = Nothing
'
'    'grabamos el movimiento en la tabla var_edicion para guias
'    GrabarOperacion xIdUsuario, 17, 1, Time, Time, Date, xCon, xId
'
'    'Fg1.Rows = 1
'    Fg1.RemoveItem Fg1.Row
'    xCon.CommitTrans
'
'    MsgBox "La guia se generó con èxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'    Set RstCab = Nothing
'    Set RstDet = Nothing
'    Exit Sub
'LaCague:
'    xCon.RollbackTrans
'    Set RstCab = Nothing
'    Set RstDet = Nothing
'    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
'    Exit Sub
'End Sub
'
'Sub GenerarGuia(Fila As Integer)
'    '===================================================================================================
'    'Creado :   /  /  Por: Enrique Pollongo
'    'Propósito: Generar una guia re remision por pedido
'    '
'    'Entradas:  Fila=Posicion del registro
'    '
'    'Resultados: Registro de la guia en vta_guia
'    '
'    'Nota:       1.- Por defecto muestra la fecha actual para la fecha de emision de la guia
'    '            2.- Indicar el chofer (Por defecto muestra el vehiculo)
'    '            3.- Indicar el vehiculo (Si no es el correcto)
'    '            4.- Todas las guias
'
'    'Modificado: 27/11/10 Por: Johan Castro
'    '           Grabar el código de la guia en el pedido para hacer la vinculacion entre las tablas
'    '           vta_pedido.idtipdoc=1(1 Guias, 2=Ventas); vta_pedido.iddocven=vta_guia.id
'    '           15/12/10 Por: Jose Chacon
'    '           Modificar las consultas de la tabla vta_pedido,vta_pedidodet por ped_pedido,ped_pedidodet
'    '           19/01/11 Por: Johan Castro
'    '           Cambiar el tipo de dato a variable xId a Double antes Integer
'    '           Agregar lineas de codigo para registrar el historial guia(nuevo registro)
'    '           Linea de codigo para registrar historial de pedido esta deshabilitado
'    '===================================================================================================
'    Dim RstCab As New ADODB.Recordset '--Cabecera de guias
'    Dim RstDet As New ADODB.Recordset '--Detalle de guias
'    Dim RstPed As New ADODB.Recordset '--Detalle de pedido
'    Dim Rst As New ADODB.Recordset 'Cabecera de pedido
'
'    Dim A, B, C As Integer
'    Dim xId As Double
'    Dim xCad, xFchGir, xFchEnt  As String
'    Dim xNumGui As String
'
'    On Error GoTo LaCague
'
'    xCon.BeginTrans
'    xNumGui = TxtNumDoc.Text
'
'    RST_Busq RstCab, "SELECT TOP 1 * FROM vta_guia", xCon
'    RST_Busq RstDet, "SELECT TOP 1 * FROM vta_guiadet", xCon
'    xId = HallaCodigoTabla("vta_guia", xCon, "id")
'
'    RstCab.AddNew
'    RstCab("id") = xId
'    RstCab("tipdoc") = 9
'    RstCab("numser") = Format(NulosC(TxtNumSer.Text), "0000")
'    RstCab("numdoc") = xNumGui
'    RstCab("idcli") = Fg1.TextMatrix(Fila, 12)
'    RstCab("dircli") = Fg1.TextMatrix(Fila, 13)
'    RstCab("fecgiro") = CDate(TxtFecha.Valor)
'    RstCab("idpunven") = Fg1.TextMatrix(Fila, 7)
'    RstCab("idmottra") = NulosN(LblIdEmpTra.Caption)
'    RstCab("idcho") = NulosN(Fg1.TextMatrix(Fila, 5))
'    RstCab("idemptra") = NulosN(LblIdEmpTra.Caption)
'    RstCab("idveh") = NulosN(Fg1.TextMatrix(Fila, 6))
'
''    RST_Busq Rst, "SELECT vta_pedido.* From vta_pedido WHERE (((vta_pedido.idpunvecli)=" & NulosN(Fg1.TextMatrix(Fila, 7)) & ") AND ((vta_pedido.anulado)=0) AND ((vta_pedido.proceso)=0) " _
''        & " AND ((vta_pedido.id)=" & NulosN(Fg1.TextMatrix(Fila, 14)) & "))", xCon
''
''    RST_Busq RstPed, "SELECT alm_inventario.id, mae_productoscen.codcen, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, vta_pedidodet.canpro, " _
''        & " alm_inventario.idunimed, vta_pedido.proceso FROM (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) " _
''        & " RIGHT JOIN (vta_pedido LEFT JOIN (mae_productoscen RIGHT JOIN vta_pedidodet ON mae_productoscen.codcen = vta_pedidodet.codpro) ON vta_pedido.id = vta_pedidodet.idped) " _
''        & " ON alm_inventario.id = mae_productoscen.iditem Where (((vta_pedido.idpunvecli) = " & NulosN(Fg1.TextMatrix(Fila, 7)) & ")) GROUP BY alm_inventario.id, mae_productoscen.codcen, " _
''        & " alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, vta_pedidodet.canpro, alm_inventario.idunimed, vta_pedido.proceso, vta_pedido.id " _
''        & " HAVING (((vta_pedido.proceso)=0) AND ((vta_pedido.id)=" & NulosN(Fg1.TextMatrix(Fila, 14)) & "))", xCon
'
'    RST_Busq Rst, "SELECT ped_pedido.* From ped_pedido WHERE (((ped_pedido.idpunvecli)=" & NulosN(Fg1.TextMatrix(Fila, 7)) & ") AND ((ped_pedido.anulado)=0) AND ((ped_pedido.proceso)=0) " _
'        & " AND ((ped_pedido.id)=" & NulosN(Fg1.TextMatrix(Fila, 14)) & "))", xCon
'
'    RST_Busq RstPed, "SELECT alm_inventario.id, mae_productoscen.codcen, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.canpro, " _
'        & " alm_inventario.idunimed, ped_pedido.proceso FROM (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) " _
'        & " RIGHT JOIN (ped_pedido LEFT JOIN (mae_productoscen RIGHT JOIN ped_pedidodet ON mae_productoscen.codcen = ped_pedidodet.codpro) ON ped_pedido.id = ped_pedidodet.idped) " _
'        & " ON alm_inventario.id = mae_productoscen.iditem Where (((ped_pedido.idpunvecli) = " & NulosN(Fg1.TextMatrix(Fila, 7)) & ")) GROUP BY alm_inventario.id, mae_productoscen.codcen, " _
'        & " alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.canpro, alm_inventario.idunimed, ped_pedido.proceso, ped_pedido.id " _
'        & " HAVING (((ped_pedido.proceso)=0) AND ((ped_pedido.id)=" & NulosN(Fg1.TextMatrix(Fila, 14)) & "))", xCon
'
'    If RstPed.RecordCount = 0 Then
'        MsgBox "No se han especificado items en el pedido del punto de venta " + Fg1.TextMatrix(Fila, 2) + Chr(13) _
'            & " del cliente" + Fg1.TextMatrix(Fila, 11), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Set RstPed = Nothing
'        Exit Sub
'    End If
'
'    If Rst.RecordCount <> 0 Then
'        Rst.MoveFirst
'        xCad = ""
'        For C = 1 To Rst.RecordCount
'            xCad = xCad + Left(Right(Rst("numcen"), 11), 10)
'            '--indicar que se procesa el pedido
'            Rst("proceso") = -1
'            '--asignar el chofer
'            Rst("idcho") = NulosN(Fg1.TextMatrix(Fila, 5))
'            '--asignar el vehiculo
'            Rst("idcar") = NulosN(Fg1.TextMatrix(Fila, 6))
'            'hacer la vinculacion entre pedido y guias  1=Guias, 2=Factura
'            Rst("idtipdoc") = 1
'            Rst("iddocven") = xId
'
'''            'grabamos el movimiento en la tabla var_edicion para pedidos
'''            GrabarOperacion xIdUsuario, 224, 2, Time, Time, Date, xCon, Rst("id")
'
'            Rst.Update
'
'            Rst.MoveNext
'            If Rst.EOF = True Then
'                Rst.MovePrevious
'                xFchGir = Rst("fchemi")
'                xFchEnt = Rst("fchent")
'                Exit For
'            End If
'            xCad = xCad + ", "
'        Next C
'    End If
'    RstCab("numordcom") = Mid(Trim(xCad), 1, 50)
'    RstCab("fchemiord") = CDate(xFchGir)
'    RstCab("fchentord") = CDate(xFchEnt)
'    RstCab("dirpunpar") = ""
'    RstCab("dirpunlle") = Mid(Trim(Fg1.TextMatrix(Fila, 13)), 1, 100)
'    'RstCab("iddocven") = "" 'ESTE CAMPO NO SE LLENA AQUI
'    'RstCab("numlote") = ""  'ESTE CAMPO NO SE LLENA AQUI
'    'RstCab("fchpro") = ""   'ESTE CAMPO NO SE LLENA AQUI
'    RstCab("tippro") = "3"
'    RstCab.Update
'
'    'GRABAMOS EL DETALLE DE LA GUIA
'    RstPed.MoveFirst
'    For A = 1 To RstPed.RecordCount
'        RstDet.AddNew
'        RstDet("idgui") = xId
'        RstDet("iditem") = RstPed("id")
'        RstDet("idunimed") = RstPed("idunimed")
'        RstDet("canpro") = RstPed("canpro")
'        RstDet.Update
'
'        RstPed.MoveNext
'        If RstPed.EOF = True Then Exit For
'    Next A
'
'    Set Rst = Nothing
'    Set RstPed = Nothing
'
'    'grabamos el movimiento en la tabla var_edicion para guias
'    GrabarOperacion xIdUsuario, 17, 1, Time, Time, Date, xCon, xId
'
'    xCon.CommitTrans
'
'    Set RstCab = Nothing
'    Set RstDet = Nothing
'    Exit Sub
'
'LaCague:
'    xCon.RollbackTrans
'    Set RstCab = Nothing
'    Set RstDet = Nothing
'    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
'    Exit Sub
'End Sub
