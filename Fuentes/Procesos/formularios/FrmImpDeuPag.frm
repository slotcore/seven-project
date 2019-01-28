VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmImpDeuPag 
   Caption         =   "Contabilidad - Inventario Inicial de Documentos x Pagar"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1200
      Left            =   2910
      TabIndex        =   12
      Top             =   3285
      Visible         =   0   'False
      Width           =   5805
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   300
         Left            =   150
         TabIndex        =   13
         Top             =   615
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Documentos : "
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   300
         Width           =   1935
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   5835
         Y1              =   1185
         Y2              =   1170
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   1
         X1              =   5790
         X2              =   5790
         Y1              =   15
         Y2              =   1155
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   30
         Y2              =   1170
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5820
         Y1              =   15
         Y2              =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   0
      TabIndex        =   2
      Top             =   345
      Width           =   11985
      Begin VB.CommandButton Command1 
         Caption         =   "Cargar"
         Enabled         =   0   'False
         Height          =   435
         Left            =   10245
         TabIndex        =   6
         Top             =   330
         Width           =   1620
      End
      Begin VB.CommandButton CmdBusCta 
         Enabled         =   0   'False
         Height          =   240
         Left            =   2805
         Picture         =   "FrmImpDeuPag.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   585
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Enabled         =   0   'False
         Height          =   240
         Left            =   7140
         Picture         =   "FrmImpDeuPag.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   270
         Width           =   240
      End
      Begin VB.TextBox TxtArchivo 
         Height          =   300
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "TxtArchivo"
         Top             =   240
         Width           =   6015
      End
      Begin VB.TextBox TxtCuenta 
         Height          =   300
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "TxtCuenta"
         Top             =   555
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivo"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   285
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label LblDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescripcion"
         Height          =   300
         Left            =   3120
         TabIndex        =   9
         Top             =   555
         Width           =   4290
      End
      Begin VB.Label LblIdCuenta 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCuenta"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   7515
         TabIndex        =   8
         Top             =   495
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   10125
         X2              =   10125
         Y1              =   150
         Y2              =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   10140
         X2              =   10140
         Y1              =   135
         Y2              =   915
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8850
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6030
      Left            =   0
      TabIndex        =   0
      Top             =   1395
      Width           =   11985
      _cx             =   21140
      _cy             =   10636
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
      BackColorSel    =   4210816
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
      Rows            =   50
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmImpDeuPag.frx":0264
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
      Left            =   8025
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
            Picture         =   "FrmImpDeuPag.frx":040A
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":0CE0
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":0E64
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":12B8
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":13D0
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":1914
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":1E58
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":1F6C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":2080
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":24D4
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpDeuPag.frx":2640
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar Inventario Inicial"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmImpDeuPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean

Private Sub CmdBusCta_Click()
    Dim xfrm As New SGI2_funciones.formularios
    Dim rst As New ADODB.Recordset
    Set rst = xfrm.SelePlanCuentas(xCon)
    If rst.State = 1 Then
        If rst.RecordCount <> 0 Then
            TxtCuenta.Text = Trim(rst("cuenta"))
            LblDescripcion.Caption = Trim(rst("descripcion"))
            LblIdCuenta.Caption = Trim(rst("id"))
        End If
    End If
    Set xfrm = Nothing
End Sub

Private Sub Command2_Click()
    'CommonDialog1.CancelError = True
    'Especificar las extensiones a usar
    CommonDialog1.DefaultExt = "*.xls"
    'CommonDialog1.Filter = "Cardfile (*.crd)|*.crd|Textos (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    CommonDialog1.Filter = "Documentos de Excel (*.xls)|*.xls"
    CommonDialog1.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
    Else
        TxtArchivo.Text = CommonDialog1.FileName
    End If
End Sub

Private Sub Command1_Click()
    If TxtArchivo.Text = "" Then
        MsgBox "No ha especificado el nombre del archivo que contiene los datos para el inventario inicial de documentos por cobrar", vbInformation + vbOKCancel + vbCritical, xTitulo
        TxtArchivo.SetFocus
        Exit Sub
    End If
    
    CargaDocumentos
End Sub

Sub CargaDocumentos()
    Dim xNumFilas As Integer
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    'Dim objExcel As New Excel.Application
    
    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo.Text)
    
    Frame2.Left = 3090
    Frame2.Top = 2910
    Label4.Caption = "Cargando registros para la importacion"
    Frame2.Visible = True
    
    xFilas = 3
    xNumFilas = 1
    
    Fg1.Rows = 1
    With objExcel.ActiveSheet
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        For A = 2 To 10000
            If NulosC(.Cells(A, 1)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit For
            End If
        Next A
        
        xNumFilas = xNumFilas + 1
        ProgressBar2.Max = xNumFilas
        
        For A = 2 To xNumFilas
            ProgressBar2.Value = A
            Frame2.Refresh
            
            If NulosC(.Cells(A, 1)) = "" Then Exit For
            Fg1.Rows = Fg1.Rows + 1
            For B = 1 To 13
                Fg1.TextMatrix(A - 1, B) = Trim(.Cells(A, B))
                If (B = 9) Or (B = 10) Or (B = 11) Or (B = 12) Or (B = 13) Then Fg1.TextMatrix(A - 1, B) = Format(Trim(.Cells(A, B)), "0.00")
                If (B = 2) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(Val(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_documento", "N", xCon)
                If (B = 3) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(Val(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_tipoproducto", "N", xCon)
                If (B = 5) Or (B = 6) Then Fg1.TextMatrix(A - 1, B) = Format(CDate(Trim(.Cells(A, B))), "dd/mm/yy")
                If (B = 7) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(Val(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_condpago", "N", xCon)
                If (B = 8) Then Fg1.TextMatrix(A - 1, B) = Busca_Codigo(Val(Fg1.TextMatrix(A - 1, B)), "id", "descripcion", "mae_moneda", "N", xCon)
                
            Next B
            
        Next A
    End With
    
    Frame2.Visible = False
    MsgBox "El proceso termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
    objExcel.WindowState = 2
    objExcel.Workbooks.Close
    
    Set objExcel = Nothing
    Exit Sub
End Sub

Function GrabarDocumentos() As Boolean
    If Val(LblIdCuenta.Caption) = 0 Then
        MsgBox "No ha especificado la cuenta contable para el inventario inicial de documentos por pagar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCuenta.SetFocus
        GrabarDocumentos = False
        Exit Function
    End If
    
    Dim A As Integer
    Dim rst As New ADODB.Recordset
    Dim Rstdoc As New ADODB.Recordset
    Dim RstDia As New ADODB.Recordset
    Dim RstTab As New ADODB.Recordset
    Dim RstTC As New ADODB.Recordset
    
    Dim xId As Integer
    Dim xNumAsiento As String
    
    Frame2.Left = 3090
    Frame2.Top = 2910
    Label4.Caption = "Importando registros"
    Frame2.Visible = True
    
    ProgressBar2.Max = Fg1.Rows - 1
    
    RST_Busq Rstdoc, "SELECT * FROM com_compras", xCon
    RST_Busq RstDia, "SELECT * FROM con_diario", xCon
    RST_Busq RstTab, "SELECT * FROM con_saldoinicial", xCon
    
    xNumAsiento = NuevoNumAsiento(1, 0, xCon)

    For A = 1 To Fg1.Rows - 1
        ProgressBar2.Value = A
        Frame2.Refresh
        Dim xCodMon As Integer
        RST_Busq rst, "SELECT * FROM mae_prov WHERE numruc = '" & Fg1.TextMatrix(A, 1) & "'", xCon
        If rst.RecordCount = 0 Then
            MsgBox "El Nº R.U.C. " + NulosC(Fg1.TextMatrix(A, 1)) + " no existe en el maestro de proveedores", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
        
        xId = HallaCodigoTabla("com_compras", xCon, "id")
        Rstdoc.AddNew
        Rstdoc("id") = xId
        Rstdoc("idtipo") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 3)), "descripcion", "id", "mae_tipoproducto", "C", xCon)
        Rstdoc("tipdoc") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 2)), "descripcion", "id", "mae_documento", "C", xCon)
        Rstdoc("idpro") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 1)), "numruc", "id", "mae_prov", "C", xCon)
        Rstdoc("numser") = Format(Mid(Fg1.TextMatrix(A, 4), 1, 4), "0000")
        Rstdoc("numdoc") = Format(Mid(Fg1.TextMatrix(A, 4), 6, 10), "0000000000")
        Rstdoc("fchreg") = CDate("01/01/" + Trim(Str(Val(AnoTra))))
        Rstdoc("fchdoc") = CDate(Fg1.TextMatrix(A, 5))
        Rstdoc("fchven") = CDate(Fg1.TextMatrix(A, 6))
        Rstdoc("idconpag") = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 7)), "descripcion", "id", "mae_condpago", "C", xCon)
        
        xCodMon = Busca_Codigo(NulosC(Fg1.TextMatrix(A, 8)), "descripcion", "id", "mae_moneda", "C", xCon)
        Rstdoc("idmon") = xCodMon
        
        Rstdoc("impbru") = Val(Fg1.TextMatrix(A, 9))
        Rstdoc("impigv") = NulosN(Fg1.TextMatrix(A, 10))
        Rstdoc("impisc") = NulosN(Fg1.TextMatrix(A, 11))
        Rstdoc("imptot") = NulosN(Fg1.TextMatrix(A, 12))
        Rstdoc("impsal") = NulosN(Fg1.TextMatrix(A, 13))
        
        If NulosN(Fg1.TextMatrix(A, 10)) = 0 Then
            Rstdoc("afecto") = 0
        Else
            Rstdoc("afecto") = -1
        End If
        Rstdoc("importado") = -1
        Rstdoc.Update
        
        RstDia.AddNew
        
        RstDia("año") = AnoTra
        RstDia("idmes") = 0
        RstDia("idlib") = 1
        RstDia("idmov") = xId
        RstDia("numasi") = xNumAsiento
        RstDia("idcue") = Val(LblIdCuenta.Caption)
        If xCodMon = 1 Then
            RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 13))
            RstDia("impdebdol") = 0
        Else
            Set RstTC = Nothing
            Set RstTC = BuscaConCriterio("SELECT * FROM con_tc WHERE fecha = cdate('" & Fg1.TextMatrix(A, 5) & "')", xCon)
            If RstTC.RecordCount <> 0 Then
                RstDia("tc") = RstTC("impven")
                RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 13)) * RstTC("impven")
                RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 13))
            Else
                RstDia("tc") = 0
                RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 13))
            End If
        End If
        RstDia.Update
        
        xCon.Execute "UPDATE com_compras SET com_compras.numreg = '" & "00" + Trim(xNumAsiento) & "' WHERE (((com_compras.id)=" & xId & "))"
        
        'grabamos el movimiento de la importacion de documentos por cobrar, se graba en esta tabla para cuando se quiera eliminar
        'el inventario inicial de documentos por cobrar
        RstTab.AddNew
        RstTab("iddoc") = xId
        RstTab("numasi") = xNumAsiento
        RstTab("idmes") = 0
        RstTab("tipo") = 2
        RstTab.Update
    Next A
    Frame2.Visible = False
    
    GrabarDocumentos = True
End Function

Sub MostrarImportados()
    Dim rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq rst, "SELECT mae_prov.numruc, mae_documento.descripcion AS descdoc, mae_tipoproducto.descripcion AS desctipcom, " _
        & " [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchdoc, com_compras.fchven, " _
        & " mae_condpago.descripcion AS descconpag, mae_moneda.descripcion AS descmon, com_compras.impbru, com_compras.impisc, " _
        & " com_compras.impigv, com_compras.imptot, com_compras.impsal " _
        & " FROM mae_moneda RIGHT JOIN (mae_condpago RIGHT JOIN (mae_tipoproducto RIGHT JOIN (mae_documento RIGHT JOIN " _
        & " (mae_prov RIGHT JOIN (con_saldoinicial LEFT JOIN com_compras ON con_saldoinicial.iddoc = com_compras.id) " _
        & " ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) ON mae_tipoproducto.id = com_compras.idtipo) " _
        & " ON mae_condpago.id = com_compras.idconpag) ON mae_moneda.id = com_compras.idmon Where (((con_saldoinicial.tipo) = 2)) " _
        & " ORDER BY com_compras.fchdoc", xCon

    Fg1.Rows = 1
    
    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = rst("numruc")
            Fg1.TextMatrix(A, 2) = rst("descdoc")
            Fg1.TextMatrix(A, 3) = rst("desctipcom")
            Fg1.TextMatrix(A, 4) = rst("numdoc")
            Fg1.TextMatrix(A, 5) = Format(rst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(A, 6) = Format(rst("fchven"), "dd/mm/yy")
            Fg1.TextMatrix(A, 7) = rst("descconpag") 'Format(Rst("impigv"), "0.00")
            Fg1.TextMatrix(A, 8) = rst("descmon") 'Format(Rst("impisc"), "0.00")
            Fg1.TextMatrix(A, 9) = Format(rst("impbru"), "0.00")
            Fg1.TextMatrix(A, 10) = Format(rst("impigv"), "0.00")
            Fg1.TextMatrix(A, 11) = Format(rst("impisc"), "0.00")
            Fg1.TextMatrix(A, 12) = Format(rst("imptot"), "0.00")
            Fg1.TextMatrix(A, 13) = Format(rst("impsal"), "0.00")
            'Fg1.TextMatrix(A, 14) = Rst("desctipitem")
            
            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
    End If
End Sub

Sub EliminarImportados()
    Dim rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq rst, "SELECT mae_prov.numruc, mae_documento.descripcion AS descdoc, mae_tipoproducto.descripcion AS desctipcom, " _
        & " [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchdoc, com_compras.fchven, " _
        & " mae_condpago.descripcion AS descconpag, mae_moneda.descripcion AS descmon, com_compras.impbru, com_compras.impisc, " _
        & " com_compras.impigv, com_compras.imptot, com_compras.impsal, con_saldoinicial.iddoc, con_saldoinicial.numasi, con_saldoinicial.idmes  " _
        & " FROM mae_moneda RIGHT JOIN (mae_condpago RIGHT JOIN (mae_tipoproducto RIGHT JOIN (mae_documento RIGHT JOIN " _
        & " (mae_prov RIGHT JOIN (con_saldoinicial LEFT JOIN com_compras ON con_saldoinicial.iddoc = com_compras.id) " _
        & " ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) ON mae_tipoproducto.id = com_compras.idtipo) " _
        & " ON mae_condpago.id = com_compras.idconpag) ON mae_moneda.id = com_compras.idmon Where (((con_saldoinicial.tipo) = 2)) " _
        & " ORDER BY com_compras.fchdoc", xCon
    
    Frame2.Left = 3090
    Frame2.Top = 2910
    Label4.Caption = ""
    Frame2.Visible = True
    
    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        'Eliminamos el diario
        Label4.Caption = "Eliminado asientos contables"
        Frame2.Refresh
        ProgressBar2.Max = rst.RecordCount
        For A = 1 To rst.RecordCount
            ProgressBar2.Value = A
            Frame2.Refresh
            xCon.Execute "DELETE con_diario.idmes, con_diario.idlib, con_diario.idmov, con_diario.numasi From con_diario " _
                & " WHERE (((con_diario.idmes)=" & rst("idmes") & ") AND ((con_diario.idlib)=1) AND ((con_diario.idmov)=" & rst("iddoc") & ") " _
                & " AND ((con_diario.numasi)='" & rst("numasi") & "'))"
            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
        
        'eliminamos los documentos importados
        Label4.Caption = "Eliminado registros de ventas"
        Frame2.Refresh
        ProgressBar2.Max = rst.RecordCount
        rst.MoveFirst
        For A = 1 To rst.RecordCount
            ProgressBar2.Value = A
            Frame2.Refresh
            xCon.Execute "DELETE com_compras.id From com_compras WHERE (((com_compras.id)=" & rst("iddoc") & "))"

            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
        
        'eliminamos los documentos de la tabla saldos iniciales
        xCon.Execute "DELETE * FROM con_saldoinicial WHERE tipo = 2"
    End If
    Fg1.Rows = 1
    Frame2.Visible = False
    'MsgBox "El inventario inicial de documentos por cobrar se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        MostrarImportados
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    TxtArchivo.Text = ""
    TxtCuenta.Text = ""
    LblDescripcion.Caption = ""
    LblIdCuenta.Caption = ""
    Fg1.Rows = 1
End Sub

Sub ActivarTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(4).Enabled = Not Toolbar1.Buttons(4).Enabled
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(7).Enabled = Not Toolbar1.Buttons(7).Enabled
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        Dim rst As New ADODB.Recordset
        
        RST_Busq rst, "SELECT * FROM con_saldoinicial WHERE tipo = 2", xCon
        If rst.RecordCount <> 0 Then
            MsgBox "Ya se importo datos para el inventario inicial de compras, elimine " + Chr(13) _
                & "el inventario inicial y vuelva a ejecutar esta opcion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set rst = Nothing
            Exit Sub
        End If
        
        Set rst = Nothing
        Command2.Enabled = True
        Command1.Enabled = True
        CmdBusCta.Enabled = True
        TxtArchivo.Text = ""
        TxtCuenta.Text = ""
        LblDescripcion.Caption = ""
        LblIdCuenta.Caption = ""
        
        MostrarImportados
        ActivarTool
    End If
    
    If Button.Index = 2 Then
        Dim Rpta As Integer
        
        Rpta = MsgBox("Esta seguro de eliminar el inventario inicial de documentos por pagar", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            EliminarImportados
            MsgBox "El inventario inicial de documentos por pagar se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    
    If Button.Index = 4 Then
        Command2.Enabled = False
        Command1.Enabled = False
        CmdBusCta.Enabled = False
        Command1.Enabled = False
        ActivarTool
    End If
    
    If Button.Index = 5 Then
        If GrabarDocumentos = True Then
            Command2.Enabled = False
            Command1.Enabled = False
            CmdBusCta.Enabled = False
            Command1.Enabled = False
            ActivarTool
            Fg1.Rows = 1
            TxtArchivo.Text = ""
            TxtCuenta.Text = ""
            LblDescripcion.Caption = ""
            LblIdCuenta.Caption = ""
            MostrarImportados
        End If
    End If
    
    If Button.Index = 7 Then
        Unload Me
    End If
End Sub


