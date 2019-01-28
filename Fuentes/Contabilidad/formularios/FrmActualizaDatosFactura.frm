VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmActualizaDatosFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas - Actualizar Datos de Ventas"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6510
      Left            =   15
      TabIndex        =   0
      Top             =   1095
      Width           =   11835
      _cx             =   20876
      _cy             =   11483
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmActualizaDatosFactura.frx":0000
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
      Top             =   0
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
            Picture         =   "FrmActualizaDatosFactura.frx":013E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":0682
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":0A14
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":0B98
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":0FEC
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":1104
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":1648
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":1B8C
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":1CA0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":1DB4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":2208
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActualizaDatosFactura.frx":2374
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   609
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
      Height          =   795
      Left            =   15
      TabIndex        =   2
      Top             =   285
      Width           =   11835
      Begin VB.CommandButton cmd_periodo 
         Height          =   255
         Index           =   0
         Left            =   3000
         Picture         =   "FrmActualizaDatosFactura.frx":28BC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   285
      End
      Begin VB.CommandButton CmdMuestra 
         Height          =   480
         Left            =   9885
         Picture         =   "FrmActualizaDatosFactura.frx":2C3E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes de Facturacion"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   315
         Width           =   1410
      End
      Begin VB.Label lbl_periodo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_periodo(0)"
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
         Height          =   300
         Index           =   0
         Left            =   1650
         TabIndex        =   5
         Top             =   285
         Width           =   1680
      End
   End
End
Attribute VB_Name = "FrmActualizaDatosFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim xFchIni, xFchFin As String

Private Sub cmd_periodo_Click(index As Integer)
    Dim mMesIni As Integer
    mMesIni = SeleccionaMes(xCon)
    lbl_periodo(0).Caption = Busca_Codigo(mMesIni, "id", "descripcion", "con_meses", "N", xCon)
    If mMesIni = 0 Then
        lbl_periodo(0).Caption = ""
    Else
        xFchIni = "01/" + Format(mMesIni, "00") + "/" + Format(Date, "YYYY")
        xFchFin = Format(HallaDiasMes(CDate(xFchIni)), "00") + "/" + Format(mMesIni, "00") + "/" + Format(Date, "YYYY")
    End If
End Sub

Private Sub CmdMuestra_Click()
    If NulosC(lbl_periodo(0).Caption) = "" Then
        MsgBox "No ha especificado el mes de facturacion a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        cmd_periodo(0).SetFocus
        Exit Sub
    End If
    
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Fg1.Rows = 1
    RST_Busq Rst, "SELECT mae_documento.abrev, vta_ventas.fchdoc, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, mae_moneda.simbolo, vta_ventas.imptotdoc, " _
        & " mae_cliente.numruc, mae_cliente.nombre, vta_ventas.idcli, vta_ventas.idmon, vta_ventas.id FROM ((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) " _
        & " LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
        & " WHERE (((vta_ventas.fchdoc)>=CDate('" & xFchIni & "') And (vta_ventas.fchdoc)<=CDate('" & xFchFin & "')) AND ((vta_ventas.anulado)<>-1)) " _
        & " ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]", xCon

    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("abrev")
            Fg1.TextMatrix(A, 2) = Rst("numdoc")
            Fg1.TextMatrix(A, 3) = Rst("fchdoc")
            Fg1.TextMatrix(A, 4) = Rst("simbolo")
            Fg1.TextMatrix(A, 5) = Rst("imptotdoc")
            Fg1.TextMatrix(A, 6) = Rst("numruc")
            Fg1.TextMatrix(A, 7) = Rst("nombre")
            Fg1.TextMatrix(A, 8) = Rst("id")
            Fg1.TextMatrix(A, 9) = Rst("idcli")
            Fg1.TextMatrix(A, 10) = Rst("idmon")
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 6 Then
        If QueHace = 3 Then Exit Sub
        Dim xform As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        
        Dim xCampos(2, 4) As String
        
        xCampos(0, 0) = "Cliente":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
        
        xform.SQLCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id, mae_cliente.idven From mae_cliente"
        
        xform.Titulo = "Buscando Cliente"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "nombre"
        xform.CampoBusca = "nombre"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                If xRs.RecordCount <> 0 Then
                    Fg1.TextMatrix(Fg1.Row, 6) = xRs("numruc")
                    Fg1.TextMatrix(Fg1.Row, 7) = xRs("nombre")
                    Fg1.TextMatrix(Fg1.Row, 9) = xRs("id")
                End If
            End If
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_EnterCell()
    If Fg1.Col = 6 Then
        If QueHace = 3 Then Exit Sub
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 3
    lbl_periodo(0).Caption = ""
    
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColComboList(6) = "|..."
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.index = 1 Then Modificar
    If Button.index = 3 Then
        If Grabar = True Then
            Cancelar
        End If
    End If
    
    If Button.index = 4 Then Cancelar
    
    If Button.index = 6 Then
        Unload Me
    End If

End Sub

Sub Cancelar()
    ActivaTool
    Fg1.Rows = 1
    QueHace = 3
    CmdMuestra_Click
End Sub

Function Grabar() As Boolean
    Dim A As Integer
    
On Error GoTo LaCague
    xCon.BeginTrans
    
    For A = 1 To Fg1.Rows - 1
        xCon.Execute "UPDATE vta_ventas SET vta_ventas.idcli = " & NulosN(Fg1.TextMatrix(A, 9)) & " WHERE (((vta_ventas.id)=" & NulosN(Fg1.TextMatrix(A, 8)) & "))"

    Next A
    
    xCon.CommitTrans
    Grabar = True
    Exit Function

LaCague:
    Grabar = False
    xCon.RollbackTrans
    MsgBox "No se pudo guardar los cambios por el siguiente motivo : " & Err.Description
    
End Function

Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    Toolbar1.Buttons(4).Enabled = Not Toolbar1.Buttons(4).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
End Sub

Sub Modificar()
    If Fg1.Rows = 1 Then
        MsgBox "No se ha especificado que documentos se modificaran", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    ActivaTool
    QueHace = 2
    Fg1.SelectionMode = flexSelectionFree
End Sub
