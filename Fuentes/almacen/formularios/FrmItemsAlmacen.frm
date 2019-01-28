VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmItemsAlmacen 
   Caption         =   "Alamacen - Asignar Items a Almacenes"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "2"
      Height          =   780
      Left            =   3450
      TabIndex        =   15
      Top             =   3315
      Visible         =   0   'False
      Width           =   5280
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grabando Items"
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
         Left            =   150
         TabIndex        =   17
         Top             =   135
         Width           =   1350
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   750
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   5265
         X2              =   5265
         Y1              =   15
         Y2              =   765
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5325
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5325
         Y1              =   15
         Y2              =   15
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7170
      Top             =   15
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
            Picture         =   "FrmItemsAlmacen.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmItemsAlmacen.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6285
      Left            =   15
      TabIndex        =   1
      Top             =   1290
      Width           =   11835
      _cx             =   20876
      _cy             =   11086
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
      BackColorSel    =   64
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
      Rows            =   1
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmItemsAlmacen.frx":277E
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
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   6000
      TabIndex        =   3
      Top             =   300
      Width           =   5850
      Begin VB.CommandButton CmdSelItem 
         Caption         =   "Seleccionar Item"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1995
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   330
         Width           =   1860
      End
      Begin VB.CommandButton CmdAddTodo 
         Caption         =   "Agregar Todos los Items"
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   525
         Width           =   1860
      End
      Begin VB.CommandButton CmdDelTodo 
         Caption         =   "Eliminar Todo"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3900
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   525
         Width           =   1860
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Agregar Item"
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1860
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Eliminar Item"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3900
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   15
      TabIndex        =   4
      Top             =   300
      Width           =   4680
      Begin VB.CommandButton CmdBusTipiTem 
         Height          =   240
         Left            =   4230
         Picture         =   "FrmItemsAlmacen.frx":2921
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   390
         Width           =   240
      End
      Begin VB.TextBox TxtAlmacen 
         Height          =   300
         Left            =   1215
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "TxtAlmacen"
         Top             =   360
         Width           =   3285
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   195
         TabIndex        =   5
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.Frame Frame3 
      Height          =   960
      Left            =   4770
      TabIndex        =   7
      Top             =   300
      Width           =   1155
      Begin VB.CommandButton CmdMuestra 
         Height          =   525
         Left            =   165
         Picture         =   "FrmItemsAlmacen.frx":2A53
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         Width           =   810
      End
      Begin VB.Label LblIdAlmacen 
         AutoSize        =   -1  'True
         Caption         =   "LblIdAlmacen"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   75
         TabIndex        =   11
         Top             =   195
         Visible         =   0   'False
         Width           =   960
      End
   End
End
Attribute VB_Name = "FrmItemsAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim CaracteresNumericos2  As String
Dim Agregando As Boolean

Private Sub CmdAdd_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT alm_inventario.*, mae_moneda.simbolo, mae_unidades.abrev FROM mae_moneda INNER JOIN (mae_unidades INNER JOIN alm_inventario " _
        & " ON mae_unidades.id = alm_inventario.idunimed) ON mae_moneda.id = alm_inventario.idmon ORDER BY alm_inventario.descripcion"

    xform.Titulo = "Buscando Items"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            If BuscarItemCuadricula(xRs("id")) = False Then
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("codpro"))
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("descripcion"))
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRs("abrev"))
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRs("simbolo"))
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(xRs("preini")), "0.0000")
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(xRs("stckini")), "0.0000")
                Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(xRs("preuni")), "0.0000")
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(xRs("porgan")), "0.0000")
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(xRs("preven")), "0.0000")
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(xRs("stckact")), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = xRs("id")
                Fg1.SetFocus
            Else
                MsgBox "El item seleccionado ya existe en la lista de items del almacen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Function BuscarItemCuadricula(IdProducto As Integer) As Boolean
    Dim A As Integer
    BuscarItemCuadricula = False
    For A = 1 To Fg1.Rows - 1
        If Val(Fg1.TextMatrix(A, 11)) = IdProducto Then
            BuscarItemCuadricula = True
        End If
    Next A
End Function

Private Sub CmdAddTodo_Click()
    AgregarTodos
End Sub

Private Sub CmdBusTipiTem_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT alm_almacenes.* FROM alm_almacenes"
    
    xform.Titulo = "Buscando Almacenes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblIdAlmacen.Caption = xRs("id")
        TxtAlmacen.Text = xRs("descripcion")
        CmdMuestra.SetFocus
        If QueHace = 3 Then CmdMuestra_Click
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Sub Activar()
    CmdAdd.Enabled = Not CmdAdd.Enabled
    CmdDel.Enabled = Not CmdDel.Enabled
    CmdDelTodo.Enabled = Not CmdDelTodo.Enabled
    CmdSelItem.Enabled = Not CmdSelItem.Enabled
    CmdAddTodo.Enabled = Not CmdAddTodo.Enabled
End Sub

Private Sub CmdDel_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No hay items para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    Fg1.RemoveItem Fg1.Row
End Sub

Private Sub CmdDelTodo_Click()
    Dim Rpta As Integer
    
    Rpta = MsgBox("¿ Esta seguro de eliminar todos los items ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Fg1.Rows = 1
    End If
End Sub

Private Sub CmdMuestra_Click()
    If TxtAlmacen.Text = "" Then
        MsgBox "No ha especificado el almacen a procesar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtAlmacen.SetFocus
        Exit Sub
    End If
    
    MuestraAlmacen
End Sub

Sub MuestraAlmacen()
    If TxtAlmacen.Text = "" Then
        MsgBox "No ha especificado el nombre del almacen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtAlmacen.SetFocus
        Exit Sub
    End If
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    Fg1.Rows = 1
    RST_Busq Rst, "SELECT alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.desctec, alm_inventarioalmacen.*, mae_moneda.simbolo AS desmon, " _
        & " mae_unidades.abrev AS desuni FROM alm_inventarioalmacen LEFT JOIN (mae_unidades RIGHT JOIN (mae_moneda RIGHT JOIN alm_inventario " _
        & " ON mae_moneda.id = alm_inventario.idmon) ON mae_unidades.id = alm_inventario.idunimed) ON alm_inventarioalmacen.iditem = alm_inventario.id " _
        & " WHERE (((alm_inventarioalmacen.idalm)=" & Val(LblIdAlmacen.Caption) & ")) ORDER BY alm_inventario.descripcion", xCon
    
    If Rst.RecordCount <> 0 Then
        Frame4.Left = 3450
        Frame4.Top = 3315
        Label2.Caption = "Cargando Items"
        ProgressBar1.Max = Rst.RecordCount
        Frame4.Visible = True
        
        Rst.MoveFirst
        Agregando = True
        For A = 1 To Rst.RecordCount
            ProgressBar1.Value = A
            Frame4.Refresh
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(Rst("codpro"))
            Fg1.TextMatrix(A, 2) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(A, 3) = NulosC(Rst("desuni"))
            Fg1.TextMatrix(A, 4) = NulosC(Rst("desmon"))
            Fg1.TextMatrix(A, 5) = Format(Rst("preini"), "0.0000")
            Fg1.TextMatrix(A, 6) = Format(Rst("stckini"), "0.0000")
            Fg1.TextMatrix(A, 7) = Format(Rst("preuni"), "0.0000")
            Fg1.TextMatrix(A, 8) = Format(Rst("porgan"), "0.0000")
            Fg1.TextMatrix(A, 9) = Format(Rst("preven"), "0.0000")
            Fg1.TextMatrix(A, 10) = Format(Rst("stckact"), "0.0000")
            Fg1.TextMatrix(A, 11) = Rst("iditem")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        Agregando = False
        Frame4.Visible = False
    End If
End Sub

Sub Seleccionar()
    Dim xFrm As New eps_librerias.FormSeleccion
    Dim xCampos(4, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim xRs1 As New ADODB.Recordset
    
    xCampos(0, 0) = "Descripcion":     xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "6000":   xCampos(0, 3) = "C":     xCampos(0, 4) = "S"
    xCampos(1, 0) = "Codigo":          xCampos(1, 1) = "codpro":         xCampos(1, 2) = "1500":   xCampos(1, 3) = "C":     xCampos(1, 4) = "N"
    xCampos(2, 0) = "Uni. Med.":       xCampos(2, 1) = "abrev":          xCampos(2, 2) = "1200":   xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "Moneda":          xCampos(3, 1) = "simbolo":        xCampos(3, 2) = "1200":   xCampos(3, 3) = "C":     xCampos(3, 4) = "N"

    xFrm.SQLCad = "SELECT alm_inventario.*, mae_moneda.simbolo, mae_unidades.abrev FROM mae_moneda INNER JOIN (mae_unidades INNER JOIN alm_inventario " _
        & " ON mae_unidades.id = alm_inventario.idunimed) ON mae_moneda.id = alm_inventario.idmon ORDER BY alm_inventario.descripcion"

    xFrm.Titulo = "Buscando Items"
    
    Set xFrm.Coneccion = xCon
    Set xRs = xFrm.Seleccionar(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount = 0 Then
            Set xRs = Nothing
            Exit Sub
        End If
        Dim A As Integer
        xRs.MoveFirst
        
        'CARGAMOS LOS DOCUMENTOS ADJUNTOS Y LO MOSTRAMOS EN LA LISTA DE "DOCUMENTOS ADJUNTOS"
        For A = 1 To xRs.RecordCount
            If BuscarItemCuadricula(xRs("id")) = False Then
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("codpro"))
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("descripcion"))
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRs("abrev"))
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRs("simbolo"))
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(xRs("preini")), "0.0000")
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(xRs("stckini")), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(xRs("preuni")), "0.0000")
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(xRs("porgan")), "0.0000")
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(xRs("preven")), "0.0000")
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(xRs("stckact")), "0.00")
                Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(xRs("id"))
            End If
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
    End If
End Sub

Sub AgregarTodos()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    RST_Busq Rst, "SELECT alm_inventario.*, mae_unidades.abrev AS desunimed, mae_moneda.simbolo AS desmon " _
        & " FROM mae_moneda INNER JOIN (mae_unidades INNER JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) " _
        & " ON mae_moneda.id = alm_inventario.idmon ORDER BY alm_inventario.descripcion", xCon

    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Agregando = True
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(Rst("codpro"))
            Fg1.TextMatrix(A, 2) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(A, 3) = NulosC(Rst("desunimed"))
            Fg1.TextMatrix(A, 4) = NulosC(Rst("desmon"))
            Fg1.TextMatrix(A, 5) = Format(NulosN(Rst("preini")), "0.0000")
            Fg1.TextMatrix(A, 6) = Format(NulosN(Rst("stckini")), "0.0000")
            Fg1.TextMatrix(A, 7) = Format(NulosN(Rst("preuni")), "0.0000")
            Fg1.TextMatrix(A, 8) = Format(NulosN(Rst("porgan")), "0.0000")
            Fg1.TextMatrix(A, 9) = Format(NulosN(Rst("preven")), "0.0000")
            Fg1.TextMatrix(A, 10) = Format(NulosN(Rst("stckact")), "0.0000")
            Fg1.TextMatrix(A, 11) = Rst("id")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        Agregando = False
    End If
    Set Rst = Nothing
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdSelItem_Click()
    Seleccionar
End Sub

Private Sub fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(3, 4) As String
    
    If Col = 3 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "3500":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Abreviatura":   xCampos(1, 1) = "abrev":          xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":        xCampos(2, 1) = "id":             xCampos(2, 2) = "2000":         xCampos(2, 3) = "N"
        
        xform.SQLCad = "SELECT mae_unidades.id, mae_unidades.descripcion, mae_unidades.abrev From mae_unidades ORDER BY mae_unidades.descripcion"

        xform.Titulo = "Buscando Unidad de Medida"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 3) = xRs("abrev")
            End If
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If

    If Col = 4 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "3500":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Abreviatura":   xCampos(1, 1) = "abrev":          xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":        xCampos(2, 1) = "id":             xCampos(2, 2) = "2000":         xCampos(2, 3) = "N"
        
        xform.SQLCad = "SELECT mae_moneda.* From mae_moneda ORDER BY mae_moneda.descripcion"

        xform.Titulo = "Buscando Moneda"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 4) = xRs("simbolo")
            End If
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col = 7 Then
        If NulosN(Fg1.TextMatrix(Row, 8)) = 0 Then
            Fg1.TextMatrix(Row, 9) = Format(Fg1.TextMatrix(Row, 7), "0.0000")
        Else
            Fg1.TextMatrix(Row, 9) = (Val(Fg1.TextMatrix(Row, 7)) * ((Val(Fg1.TextMatrix(Row, 8)) / 100) + 1))
            Fg1.TextMatrix(Row, 9) = Format(Fg1.TextMatrix(Row, 9), "0.0000")
        End If
    End If
    If Col = 8 Then
        Fg1.TextMatrix(Row, 9) = Val(Fg1.TextMatrix(Row, 7)) * ((Val(Fg1.TextMatrix(Row, 8)) / 100) + 1)
        Fg1.TextMatrix(Row, 9) = Format(Fg1.TextMatrix(Row, 9), "0.0000")
    End If
    If Col = 9 Then
        If NulosN(Fg1.TextMatrix(Row, 7)) = 0 Then
            If NulosN(Fg1.TextMatrix(Row, 8)) = 0 Then
                Fg1.TextMatrix(Row, 7) = Format(Fg1.TextMatrix(Row, 9), "0.0000")
            Else
                Fg1.TextMatrix(Row, 7) = Val(Fg1.TextMatrix(Row, 9)) / ((Val(Fg1.TextMatrix(Row, 8)) / 100) + 1)
                Fg1.TextMatrix(Row, 7) = Format(Fg1.TextMatrix(Row, 7), "0.0000")
            End If
        Else
            Fg1.TextMatrix(Row, 8) = Val(Fg1.TextMatrix(Row, Col)) / Val(Fg1.TextMatrix(Row, Col - 2))
            Fg1.TextMatrix(Row, 8) = ((Val(Fg1.TextMatrix(Row, Col - 1)) * 100) - 100)
            Fg1.TextMatrix(Row, 8) = Format(Fg1.TextMatrix(Row, 8), "0.0000")
        End If
    End If
    Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "0.0000")
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col >= 5 Then If InStr(CaracteresNumericos2, Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        TxtAlmacen.Text = ""
        SeEjecuto = True
        TxtAlmacen.SetFocus
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    Fg1.Rows = 1
    Fg1.ColWidth(11) = 0
    Fg1.SelectionMode = flexSelectionByRow
    CaracteresNumericos2 = "0123456789." & Chr(8) & Chr(13)
    
    Fg1.FrozenCols = 2
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Modificar
    
    If Button.Index = 3 Then
        If Grabar = True Then
            Cancelar
        End If
    End If
    
    If Button.Index = 4 Then Cancelar
    
    If Button.Index = 6 Then
        Unload Me
    End If
End Sub

Sub Cancelar()
    QueHace = 3
    ActivaTool
    Activar
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.SetFocus
End Sub

Function Grabar() As Boolean
    Grabar = False
    If TxtAlmacen.Text = "" Then
        MsgBox "No ha especificado el almacen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtAlmacen.SetFocus
        Exit Function
    End If

    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para el almacen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    xCon.Execute "DELETE * FROM alm_inventarioalmacen WHERE idalm = " & Val(LblIdAlmacen.Caption) & " "
    
    RST_Busq Rst, "SELECT * FROM alm_inventarioalmacen", xCon
    If Fg1.Rows <> 1 Then
        Frame4.Left = 3450
        Frame4.Top = 3315
        ProgressBar1.Max = Fg1.Rows - 1
        Label2.Caption = "Grabando Items"
        Frame4.Visible = True
        
        For A = 1 To Fg1.Rows - 1
            ProgressBar1.Value = A
            Frame4.Refresh
            Rst.AddNew
            Rst("iditem") = Fg1.TextMatrix(A, 11)
            Rst("idalm") = Val(LblIdAlmacen.Caption)
            Rst("stckini") = NulosN(Fg1.TextMatrix(A, 6))
            Rst("stckact") = NulosN(Fg1.TextMatrix(A, 10))
            Rst("preini") = NulosN(Fg1.TextMatrix(A, 5))
            Rst("preuni") = NulosN(Fg1.TextMatrix(A, 7))
            Rst("porgan") = NulosN(Fg1.TextMatrix(A, 8))
            Rst("preven") = NulosN(Fg1.TextMatrix(A, 9))
            Rst("idunimed") = Busca_Codigo(Fg1.TextMatrix(A, 3), "abrev", "id", "mae_unidades", "C", xCon)
            Rst("idmon") = Busca_Codigo(Fg1.TextMatrix(A, 4), "simbolo", "id", "mae_moneda", "C", xCon)
            Rst.Update
        Next A
        Frame4.Visible = False
    End If
    
    xCon.CommitTrans
    Set Rst = Nothing
    MsgBox "Los items se agregaron con exito al " + NulosC(TxtAlmacen.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set Rst = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    Toolbar1.Buttons(4).Enabled = Not Toolbar1.Buttons(4).Enabled
    
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
End Sub

Sub Modificar()
    QueHace = 2
    ActivaTool
    Activar
    Fg1.Rows = 1
    CmdMuestra_Click
    Fg1.SelectionMode = flexSelectionFree
    Fg1.ColComboList(3) = "|..."
    Fg1.ColComboList(4) = "|..."
    Fg1.Editable = flexEDKbdMouse
    Fg1.SetFocus
End Sub

Private Sub TxtAlmacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtAlmacen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipiTem_Click
    End If
End Sub
