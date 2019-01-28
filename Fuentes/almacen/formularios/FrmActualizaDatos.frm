VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmActualizaDatos 
   Caption         =   "Almacen - Actualizacion de Precios"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   2925
      TabIndex        =   10
      Top             =   3315
      Visible         =   0   'False
      Width           =   6045
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   135
         TabIndex        =   11
         Top             =   465
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Guardando Items"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   210
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   6045
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   6030
         X2              =   6030
         Y1              =   15
         Y2              =   810
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   15
         Y2              =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   6030
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.CommandButton CmdBusFin 
      Height          =   240
      Left            =   7275
      Picture         =   "FrmActualizaDatos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   765
      Width           =   240
   End
   Begin VB.CommandButton CmdBusIni 
      Height          =   240
      Left            =   7275
      Picture         =   "FrmActualizaDatos.frx":0132
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   465
      Width           =   240
   End
   Begin VB.TextBox TxtFin 
      Height          =   300
      Left            =   915
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "TxtFin"
      Top             =   735
      Width           =   6630
   End
   Begin VB.TextBox TxtIni 
      Height          =   300
      Left            =   915
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "TxtIni"
      Top             =   435
      Width           =   6630
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6285
      Left            =   15
      TabIndex        =   3
      Top             =   1200
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmActualizaDatos.frx":0264
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
      Left            =   8145
      Top             =   -225
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
            Picture         =   "FrmActualizaDatos.frx":0429
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":096D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":0CFF
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":0E83
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":12D7
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":13EF
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":1933
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":1E77
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":1F8B
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":209F
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":24F3
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatos.frx":265F
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   4
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
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   7770
      TabIndex        =   9
      Top             =   300
      Width           =   4095
      Begin VB.CommandButton CmdMuestra 
         Height          =   525
         Left            =   1590
         Picture         =   "FrmActualizaDatos.frx":2BA7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1110
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   780
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   465
   End
End
Attribute VB_Name = "FrmActualizaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMACTUALIZADATOS.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARO PARA LA ACTUALIZACION RAPIDA Y DIRECTA DE LOS ITEMS REGISTRADOS EN EL
'*                  : SISTEMA, PERMITE MODIFICAR LOS SIGUIENTES DATOS: UNIDAD DE MEDIDA, PRECIO INICIAL
'*                    STOCK INICIAL, PRECIO DE COMPRA, PORCENTAJE DE GANANCIA, PRECIO DE VENTA, STOCK
'*                    ACTUAL, STOCK MINIMO, STOCK MAXIMO.
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 08/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit
Dim Rst As New ADODB.Recordset        ' RECORDSET PRINCIPAL PARA CARGAR LOS REGISTRO DE LA TABLA alm_inventario
Dim SeEjecuto As Boolean              ' ESPECIFICA SI EL EVENTO ACTIVATE YA SE EJECUTO, VARIABLE USADA COMO SWITCH
Dim QueHace As Integer                ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO 1 = NUEVO; 2 = MODIFICA; 3 = SOLO LECTURA
Dim Agregando As Boolean              ' VARIABLE UTILIZADA PARA CONTROLAR EL EVENTO RolColChamge DE LOS CONTROLES FlexGrid
Dim CaracteresNumericos2  As String   ' VARIABLE UTILIZADA PARA ALMACENAR CARACTERES NUMERICOS Y VALIDARLOS EN EL EVENTO KeyPress  DE
                                      ' DE LOS CUADROS DE TEXTO

Private Sub CmdBusFin_Click()
    ' BUSCAMOS EL REGISTRO FINAL PARA APLICAR EL RANGO DE BUSQUEDA
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev FROM mae_unidades " _
        & " RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed ORDER BY alm_inventario.codpro"
    
    xform.Titulo = "Buscando Items"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtFin.Text = xRs("descripcion")
            CmdMuestra.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusIni_Click()
    ' BUSCAMOS EL REGISTRO INICIAL PARA APLICAR EL RANGO DE BUSQUEDA
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev FROM mae_unidades " _
        & " RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed ORDER BY alm_inventario.codpro"
    
    xform.Titulo = "Buscando Items"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIni.Text = xRs("descripcion")
            TxtFin.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdMuestra_Click()
    ' VALIDAMOS QUE LOS RANGOS DE BUSQUEDA SEAN LOS CORRECTOS PARA APLICAR LA BUSQUEDA DE REGISTROS
    If NulosC(TxtIni.Text) <> "" And NulosC(TxtFin.Text) = "" Then
        ' mostramos los items desde el item inicial
        RST_Busq Rst, "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
            & " WHERE (((alm_inventario.descripcion)>='" & Trim(TxtIni.Text) & "')) ORDER BY alm_inventario.descripcion", xCon
    End If
    
    If NulosC(TxtIni.Text) = "" And NulosC(TxtFin.Text) <> "" Then
        ' mostramos los items hasta el item final
        RST_Busq Rst, "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
            & " WHERE (((alm_inventario.descripcion)<='" & Trim(TxtFin.Text) & "')) ORDER BY alm_inventario.descripcion", xCon
    End If
    
    If NulosC(TxtIni.Text) <> "" And NulosC(TxtFin.Text) <> "" Then
        'mostramos los itemas desde el item inicial hasta el item final
        RST_Busq Rst, "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
            & " WHERE (((alm_inventario.descripcion)>='" & TxtIni.Text & "' And (alm_inventario.descripcion)<='" & TxtFin.Text & "')) ORDER BY alm_inventario.descripcion", xCon
    End If
    
    If NulosC(TxtIni.Text) = "" And NulosC(TxtFin.Text) = "" Then
        'mostramos todos los items
        RST_Busq Rst, "SELECT alm_inventario.*, mae_unidades.abrev FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed", xCon
    End If
    
    If Rst.RecordCount = 0 Then
        MsgBox "No se han encontrado items con los valores requerido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.Rows = 1
        Set Rst = Nothing
        Exit Sub
    Else
        ' MOSTRAMOS LOS ITEMS ENCONTRADOS SEGUN CRITERIOS DE BUSQUEDA
        Dim A As Integer
        Rst.MoveFirst
        Fg1.Rows = 1
        Agregando = True
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(Rst("codpro"))
            Fg1.TextMatrix(A, 2) = Rst("descripcion")
            Fg1.TextMatrix(A, 3) = Rst("abrev")
            Fg1.TextMatrix(A, 4) = Format(Rst("preini"), "0.0000")
            Fg1.TextMatrix(A, 5) = Format(Rst("stckini"), "0.0000")
            Fg1.TextMatrix(A, 6) = Format(Rst("preuni"), "0.0000")
            Fg1.TextMatrix(A, 7) = Format(Rst("porgan"), "0.00")
            Fg1.TextMatrix(A, 8) = Format(Rst("preven"), "0.0000")
            Fg1.TextMatrix(A, 9) = Format(Rst("stckact"), "0.0000")
            Fg1.TextMatrix(A, 10) = Rst("id")
            Fg1.TextMatrix(A, 11) = Format(Rst("stckmin"), "0.00")
            Fg1.TextMatrix(A, 12) = Format(Rst("stckmax"), "0.00")
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
        Agregando = False
    End If
End Sub

Private Sub fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 3 Then
        ' MOSTRAMOS UNA LISTA DE LAS UNIDADES DE MEDIDA DISPONIBLES, SE REEMPLAZARA SEGUN CRITERIO DEL USUARIO
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
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
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    ' HACEMOS LO CALCULOS NECESARIO SEGUN SE VAYAN INGRESANDO LOS DATOS EN EL CONTROL FlexGrid
    If Agregando = True Then Exit Sub
    If Col = 6 Then   ' PRECIO DE COMPRA, SI SE CAMBIA EL PRECIO SE VUELVE A CALCULAR EL PRECIO DE VENTA EN FUNCION AL PORCENTAJE DE GANANCIA
        If NulosN(Fg1.TextMatrix(Row, 7)) = 0 Then
            Fg1.TextMatrix(Row, 8) = Format(Fg1.TextMatrix(Row, 6), "0.0000")
        Else
            Fg1.TextMatrix(Row, 8) = (Val(Fg1.TextMatrix(Row, 6)) * ((Val(Fg1.TextMatrix(Row, 7)) / 100) + 1))
            Fg1.TextMatrix(Row, 8) = Format(Fg1.TextMatrix(Row, 8), "0.0000")
        End If
    End If
    If Col = 7 Then ' PORCENTAJE DE GANANCIA, SI SE CAMBIA EL PORCENTAJE DE GANANCIA O EL PRECIO DE COMPRA SE VUELVE A CALCULAR EL PRECIO DE VENTA
        Fg1.TextMatrix(Row, 8) = Val(Fg1.TextMatrix(Row, 6)) * ((Val(Fg1.TextMatrix(Row, 7)) / 100) + 1)
        Fg1.TextMatrix(Row, 8) = Format(Fg1.TextMatrix(Row, 8), "0.0000")
    End If
    If Col = 8 Then ' PRECIO DE VENTA, SI SE CAMBIA EL PRECIO DE VENTA SE VUELVE A CALCULAR EL PORCENTAJE DE GANANCIA EN FUNCION
                    ' AL PRECIO DE VENTA / PRECIO DE COMPRA
        If NulosN(Fg1.TextMatrix(Row, 6)) = 0 Then
            If NulosN(Fg1.TextMatrix(Row, 7)) = 0 Then
                Fg1.TextMatrix(Row, 6) = Format(Fg1.TextMatrix(Row, 8), "0.0000")
            Else
                Fg1.TextMatrix(Row, 6) = Val(Fg1.TextMatrix(Row, 8)) / ((Val(Fg1.TextMatrix(Row, 7)) / 100) + 1)
                Fg1.TextMatrix(Row, 6) = Format(Fg1.TextMatrix(Row, 6), "0.0000")
            End If
        Else
            Fg1.TextMatrix(Row, 7) = Val(Fg1.TextMatrix(Row, Col)) / Val(Fg1.TextMatrix(Row, Col - 2))
            Fg1.TextMatrix(Row, 7) = ((Val(Fg1.TextMatrix(Row, Col - 1)) * 100) - 100)
            Fg1.TextMatrix(Row, 7) = Format(Fg1.TextMatrix(Row, 7), "0.0000")
        End If
    End If
    Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "0.0000")
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg1.Col = 1 Or Fg1.Col = 2 Then
        Fg1.Editable = flexEDNone
    Else
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col >= 4 Then If InStr(CaracteresNumericos2, Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO QUE SE EJECUTARA AL CARGAR EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        TxtIni.SetFocus
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO QUE SE EJECUTARA AL CARGAR EL FORMULARIO
    SeEjecuto = False
    QueHace = 3
    TxtIni.Text = ""
    TxtFin.Text = ""
    Fg1.ColWidth(10) = 0
    Fg1.Editable = flexEDNone
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.Rows = 1
    CaracteresNumericos2 = "0123456789." & Chr(8) & Chr(13)
    Fg1.FrozenCols = 2
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Modificar
    
    If Button.Index = 3 Then
        If Grabar = True Then
            Cancelar
            CmdMuestra_Click
        End If
    End If
    If Button.Index = 4 Then
        Cancelar
        CmdMuestra_Click
    End If
    
    If Button.Index = 6 Then
        Set Rst = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GRABA LOS REGISTRO EN LA TABLA Alm_inventario, DEVUELVE VERDADERO CUANDO TIENE
'*                    EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    If Fg1.Rows = 1 Then
        MsgBox "La lista se encuentra vacia no hay nada que grabar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Grabar = False
        Exit Function
    End If
    
    Dim A As Integer
    Dim xIdUni As Integer
    Frame2.Left = 3360
    Frame2.Top = 3075
    ProgressBar1.Max = Fg1.Rows - 1
    Frame2.Visible = True
    
    On Error GoTo LaCague
    xCon.BeginTrans
    For A = 1 To Fg1.Rows - 1
        ProgressBar1.Value = A
        Frame2.Refresh
        xIdUni = Busca_Codigo(Fg1.TextMatrix(A, 3), "abrev", "id", "mae_unidades", "C", xCon)
        ' ACTUALIZA LOS DATOS EN LA TABLA
        xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckini = " & Val(Fg1.TextMatrix(A, 5)) & ", alm_inventario.preini = " & Val(Fg1.TextMatrix(A, 4)) & ", " _
            & " alm_inventario.preuni = " & Val(Fg1.TextMatrix(A, 6)) & ", alm_inventario.idunimed = " & xIdUni & ", " _
            & " alm_inventario.stckact = " & Val(Fg1.TextMatrix(A, 9)) & ", " _
            & " alm_inventario.preven = " & Val(Fg1.TextMatrix(A, 8)) & ", alm_inventario.porgan = " & Val(Fg1.TextMatrix(A, 7)) & ", " _
            & " alm_inventario.stckmin = " & Val(Fg1.TextMatrix(A, 11)) & ", alm_inventario.stckmax = " & Val(Fg1.TextMatrix(A, 12)) & "" _
            & " WHERE (((alm_inventario.id)=" & Val(Fg1.TextMatrix(A, 10)) & "))"
    Next A
    
    Frame2.Visible = False
    xCon.CommitTrans
    Grabar = True
    Exit Function

LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELAR EL PROCESO DE INGRESO O MODIFICACION DE REGISTROS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Fg1.Editable = flexEDNone
    Frame1.Enabled = True
    CmdBusIni.Enabled = True
    CmdBusFin.Enabled = True
    ActivaTool
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColComboList(3) = ""
    TxtIni.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR REGISTROS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    Fg1.Editable = flexEDKbdMouse
    Frame1.Enabled = False
    CmdBusIni.Enabled = False
    CmdBusFin.Enabled = False
    ActivaTool
    Fg1.SelectionMode = flexSelectionFree
    Fg1.ColComboList(3) = "|..."
    Fg1.ColComboList(4) = "|..."
    Fg1.ColComboList(5) = "|..."
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    Toolbar1.Buttons(4).Enabled = Not Toolbar1.Buttons(4).Enabled
    
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
End Sub

Private Sub TxtFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtFin_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusFin_Click
    End If
    If KeyCode = 46 Then
        TxtFin.Text = ""
    End If
End Sub

Private Sub TxtIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIni_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusIni_Click
    End If
    If KeyCode = 46 Then
        TxtIni.Text = ""
    End If
End Sub
