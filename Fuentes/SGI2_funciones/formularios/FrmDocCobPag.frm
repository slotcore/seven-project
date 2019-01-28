VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDocCobPag 
   Caption         =   " Documentos por Cobrar y Pagar"
   ClientHeight    =   7425
   ClientLeft      =   -735
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleMode       =   0  'User
   ScaleWidth      =   1.52981e5
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   5595
      Left            =   15
      TabIndex        =   0
      Top             =   1305
      Width           =   11550
      _cx             =   20373
      _cy             =   9869
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   4210816
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      Rows            =   1
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmDocCobPag.frx":0000
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
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   1245
         Left            =   3405
         TabIndex        =   1
         Top             =   2235
         Visible         =   0   'False
         Width           =   5010
         Begin VB.Frame Frame6 
            Height          =   705
            Left            =   75
            TabIndex        =   2
            Top             =   390
            Width           =   4845
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   300
               Left            =   75
               TabIndex        =   3
               Top             =   240
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   529
               _Version        =   393216
               Appearance      =   0
               Scrolling       =   1
            End
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exportando a Excel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   150
            TabIndex        =   4
            Top             =   105
            Width           =   1665
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000002&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000002&
            Height          =   315
            Left            =   45
            Top             =   45
            Width           =   4905
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            BorderWidth     =   2
            Index           =   1
            X1              =   15
            X2              =   15
            Y1              =   0
            Y2              =   1200
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   0
            X1              =   4995
            X2              =   4995
            Y1              =   0
            Y2              =   1200
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            BorderWidth     =   2
            Index           =   1
            X1              =   0
            X2              =   4995
            Y1              =   15
            Y2              =   15
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderWidth     =   2
            Index           =   0
            X1              =   -15
            X2              =   4980
            Y1              =   1230
            Y2              =   1230
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a Excel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   990
      Left            =   0
      TabIndex        =   5
      Top             =   270
      Width           =   2880
      Begin VB.OptionButton optDocporpag 
         Caption         =   "Documentos por Pagar"
         Height          =   225
         Left            =   285
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   1965
      End
      Begin VB.OptionButton optDocporcob 
         Caption         =   "Documentos por Cobrar"
         Height          =   225
         Left            =   285
         TabIndex        =   6
         Top             =   600
         Width           =   1995
      End
   End
   Begin VB.Frame Frame2 
      Height          =   990
      Left            =   2895
      TabIndex        =   8
      Top             =   270
      Width           =   5790
      Begin VB.CommandButton cmdEnt 
         Height          =   240
         Left            =   5325
         Picture         =   "FrmDocCobPag.frx":016D
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   285
         Width           =   255
      End
      Begin VB.TextBox txtEnt 
         Height          =   285
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   255
         Width           =   4500
      End
      Begin VB.Label lblEnt 
         AutoSize        =   -1  'True
         Caption         =   "lblEnt"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3375
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lblNomEnt 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   390
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
            Picture         =   "FrmDocCobPag.frx":029F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":07E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":0B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":0CCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":1061
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":11E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":1639
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":1751
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":1C95
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":21D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":22ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":2401
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":2855
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocCobPag.frx":29C1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   990
      Left            =   8700
      TabIndex        =   20
      Top             =   270
      Width           =   2850
      Begin VB.CommandButton cmdPorVencer 
         Caption         =   "Mostrar por Vencer"
         Height          =   345
         Left            =   690
         TabIndex        =   22
         Top             =   555
         Width           =   1530
      End
      Begin VB.CommandButton CmdVen 
         Caption         =   "&Mostrar Vencidos"
         Height          =   345
         Left            =   675
         TabIndex        =   21
         Top             =   180
         Width           =   1530
      End
   End
   Begin VB.Frame Frame4 
      Height          =   570
      Left            =   15
      TabIndex        =   13
      Top             =   6855
      Width           =   11520
      Begin VB.TextBox TxtTotalSal 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00400000&
         Height          =   300
         Left            =   7845
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   990
      End
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00400000&
         Height          =   300
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   990
      End
      Begin VB.TextBox TxtNumReg 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº de Registros"
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   210
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "TOTALES S/. ==>"
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
         Left            =   5160
         TabIndex        =   17
         Top             =   225
         Width           =   1560
      End
   End
End
Attribute VB_Name = "FrmDocCobPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim T As Integer

Private Sub cmdEnt_Click()
    'Dim xform As New EPS_Buscar.Buscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String

    If Me.optDocporpag.Value = True Then
   
        xCampos(0, 0) = "Proveedor":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "6000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "RUC":        xCampos(1, 1) = "numruc":   xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
        
        'Solo mostramos los empleados que sean diferente a 6 = OBRERO
        xform.SQLCad = "SELECT * FROM mae_prov WHERE activo = -1"
        'and m_trabajadores.id not in (" & LblIdSuperv.Caption & ")
        xform.TITULO = "Buscando Proveedor"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "nombre"
        xform.CampoBusca = "nombre"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            txtEnt.Text = xRs("nombre")
            lblEnt.Caption = xRs("id")
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If

    If Me.optDocporcob.Value = True Then
        xCampos(0, 0) = "Cliente":  xCampos(0, 1) = "nombre":  xCampos(0, 2) = "6000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "RUC":      xCampos(1, 1) = "numruc":  xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
        
        'Solo mostramos los empleados que sean diferente a 6 = OBRERO
        xform.SQLCad = "SELECT * FROM mae_cliente WHERE activo = -1"
        'and m_trabajadores.id not in (" & LblIdSuperv.Caption & ")
        xform.TITULO = "Buscando Cliente"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "nombre"
        xform.CampoBusca = "nombre"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            txtEnt.Text = xRs("nombre")
            lblEnt.Caption = xRs("id")
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub cmdPorVencer_Click()
    Consultar 2
    T = 2
End Sub

Private Sub CmdVen_Click()
    Consultar 1
    T = 1
End Sub

Sub Consultar(Tipo As Integer)
    Dim rsTC As New ADODB.Recordset
    Dim CadWhere As String
    
    If Tipo = 1 Then
            CadWhere = "[fchven]-Date() <=0"
        Else
            CadWhere = "[fchven]-Date() >=1"
    End If
    
    If optDocporpag.Value = True Then
        RST_Busq Rs, "SELECT mae_prov.nombre, com_compras.idpro ,mae_documento.abrev, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc,  " _
                   & " mae_moneda.simbolo, com_compras.idmon as IdMoneda, com_compras.imptot, com_compras.impsal, com_compras.fchdoc, com_compras.fchven, " _
                   & " [fchven]-Date() AS Atrazo FROM mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_prov RIGHT JOIN com_compras " _
                   & " ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon " _
                   & " WHERE (((com_compras.impsal)<>0)) AND " & Trim(CadWhere) & " ", xCon
               
    End If
    If optDocporcob.Value = True Then
        RST_Busq Rs, "SELECT mae_cliente.nombre, vta_ventas.idcli, mae_documento.abrev, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, " _
                   & " mae_moneda.simbolo, vta_ventas.idmon as IdMoneda, vta_ventas.imptotdoc as imptot, vta_ventas.impsal, vta_ventas.fchdoc, vta_ventas.fchven, [fchven]-Date() AS Atrazo " _
                   & " FROM mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) " _
                   & " ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon " _
                   & " WHERE vta_ventas.anulado = 0 AND " & Trim(CadWhere) & "" _
                   & " AND vta_ventas.impsal <> 0 ", xCon
    End If
    
    If Me.optDocporpag.Value Then
        If Me.txtEnt.Text = "" Then
            Else
                Criterio = "idpro = '" & NulosC(lblEnt) & "'"
        End If
    End If
    
    If Me.optDocporcob.Value Then
        If Me.txtEnt.Text <> "" Then
            Criterio = "idcli = '" & NulosC(lblEnt) & "'"
        End If
    End If
    
    If Criterio <> "" Then Rs.Filter = Criterio
    
    Fg1.Rows = 1
    If Rs.RecordCount = 0 Then
        MsgBox "No se han encontrado registros", vbInformation, "Mensaje"
        Set Rs = Nothing
        TxtTotal.Text = "0.00"
        TxtTotalSal.Text = "0.00"
        TxtNumReg.Text = "0"
        Exit Sub
    End If
    
    ProgressBar1.Max = Rs.RecordCount
    Frame5.Visible = True
    Label7.Caption = "Procesando Registros"
    
    
    For a = 1 To Rs.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        ProgressBar1.Value = a
        Frame5.Refresh
        
        RST_Busq rsTC, "SELECT * FROM con_tc WHERE idmon = " & Rs("IdMoneda") & " " _
                     & " AND fecha = cdate('" & Rs("fchdoc") & "') ", xCon
        
        Fg1.TextMatrix(a, 1) = Rs("nombre")
        Fg1.TextMatrix(a, 2) = Rs("abrev")
        Fg1.TextMatrix(a, 3) = Rs("numdoc")
        Fg1.TextMatrix(a, 4) = Rs("simbolo")
        If rsTC.RecordCount <> 0 Then
            If Me.optDocporpag.Value = True Then
                Fg1.TextMatrix(a, 5) = Format(Val(Rs("imptot")) / Val(rsTC("impcom")), "0.00")
            End If
            If Me.optDocporcob.Value = True Then
                Fg1.TextMatrix(a, 5) = Format(Val(Rs("imptot")) * Val(rsTC("impven")), "0.00")
            End If
        End If
        Fg1.TextMatrix(a, 6) = Format(Rs("imptot"), "0.00")
        Fg1.TextMatrix(a, 7) = Format(Rs("impsal"), "0.00")
        Fg1.TextMatrix(a, 8) = Rs("fchdoc")
        Fg1.TextMatrix(a, 9) = Rs("fchven")
        Fg1.TextMatrix(a, 10) = Rs("Atrazo")
        Rs.MoveNext
    Next a
    
    Frame5.Visible = False
    
    
    Dim Tot, TotSal As Double
    
    For a = 1 To Fg1.Rows - 1
        Tot = Tot + Val(Fg1.TextMatrix(a, 6))
        TotSal = TotSal + Val(Fg1.TextMatrix(a, 7))
        Fg1.TextMatrix(a, 10) = Abs(Val(Fg1.TextMatrix(a, 10)))
    Next a
    
    TxtTotal.Text = Format(Tot, "0.00")
    TxtTotalSal.Text = Format(TotSal, "0.00")
    TxtNumReg.Text = Fg1.Rows - 1
    
End Sub

Private Sub Form_Load()
    Fg1.ColWidth(11) = 0
End Sub

Private Sub optDocporcob_Click()
    txtEnt.Text = ""
    lblEnt.Caption = ""
End Sub

Private Sub optDocporpag_Click()
    txtEnt.Text = ""
    lblEnt.Caption = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.index = 1 Then Imprimir
    If Button.index = 2 Then ExportarExcel
    If Button.index = 4 Then Unload Me

End Sub

Sub ExportarExcel()
    If Fg1.Rows = 1 Then
        MsgBox "No se ha registrado datos para exportar", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
        Exit Sub
    End If
    
    Dim a As Integer
    Dim B As Integer
    Dim xFilas As Integer
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    objExcel.Workbooks.Add
    objExcel.WindowState = xlMinimized
    Frame5.Visible = True
    Label7.Caption = "Exportando a Excel"
    
    xFilas = 7
    ProgressBar1.Max = Fg1.Rows - 1

    With objExcel.ActiveSheet
        
        .Cells(2, 2) = "EMPRESA"
        .Cells(2, 3) = xNomEmp
        .Cells(3, 2) = "RUC"
        .Cells(3, 3) = xNumRuc
        .Cells(4, 2) = "FECHA"
        .Cells(4, 3) = Date
        
        For a = 0 To Fg1.Rows - 1
            ProgressBar1.Value = a
            Frame5.Refresh
            For B = 0 To Fg1.Cols - 1
                If B <= 4 Then
                    .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(a, B)
                Else
                    If (B = 5) Or (B = 6) Or (B = 7) Or (B = 10) Then
                        .Cells(xFilas, B + 1) = Val(Fg1.TextMatrix(a, B))
                    Else
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(a, B)
                    End If
                End If
            Next B
            xFilas = xFilas + 1
        Next a
    End With
    
    Frame5.Visible = False
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
End Sub

Sub Imprimir()
    Dim a As Integer
    Dim RsTemp As New ADODB.Recordset
    RsTemp.CursorType = adOpenStatic
        
    RsTemp.Fields.Append "ent", adLongVarChar, 50, adFldIsNullable
    RsTemp.Fields.Append "td", adLongVarChar, 10, adFldIsNullable
    RsTemp.Fields.Append "numdoc", adLongVarChar, 15, adFldIsNullable
    RsTemp.Fields.Append "mon", adLongVarChar, 10, adFldIsNullable
    RsTemp.Fields.Append "totdol", adLongVarChar, 10, adFldIsNullable
    RsTemp.Fields.Append "totsol", adLongVarChar, 10, adFldIsNullable
    RsTemp.Fields.Append "salact", adLongVarChar, 10, adFldIsNullable
    RsTemp.Fields.Append "fecemi", adLongVarChar, 10, adFldIsNullable
    RsTemp.Fields.Append "fecven", adLongVarChar, 10, adFldIsNullable
    RsTemp.Fields.Append "atr", adLongVarChar, 10, adFldIsNullable
    
    RsTemp.Open
    
    For a = 1 To Fg1.Rows - 1
        RsTemp.AddNew
        RsTemp("ent") = Fg1.TextMatrix(a, 1)
        RsTemp("td") = Fg1.TextMatrix(a, 2)
        RsTemp("numdoc") = Fg1.TextMatrix(a, 3)
        RsTemp("mon") = Fg1.TextMatrix(a, 4)
        RsTemp("totdol") = Fg1.TextMatrix(a, 5)
        RsTemp("totsol") = Fg1.TextMatrix(a, 6)
        RsTemp("salact") = Fg1.TextMatrix(a, 7)
        RsTemp("fecemi") = Fg1.TextMatrix(a, 8)
        RsTemp("fecven") = Fg1.TextMatrix(a, 9)
        RsTemp("atr") = Fg1.TextMatrix(a, 10)
        RsTemp.Update
    Next a
    
    If Me.optDocporpag.Value = True Then
        rptDocPorPC.Sections("sección4").Controls("lblEntidad").Caption = "Proveedor"
        rptDocPorPC.Sections("sección4").Controls("lblEntidadNom").Caption = txtEnt.Text
        rptDocPorPC.Sections("sección4").Controls("lblReporte").Caption = "REPORTE DE DOCUMENTOS POR PAGAR"
    End If
    If T = 1 Then
            rptDocPorPC.Sections("sección4").Controls("lbldocumento").Caption = "DOCUMENTOS VENCIDOS"
        Else
            rptDocPorPC.Sections("sección4").Controls("lbldocumento").Caption = "DOCUMENTOS POR VENCER"
    End If
    If Me.optDocporcob.Value = True Then
        rptDocPorPC.Sections("sección4").Controls("lblEntidad").Caption = "Cliente"
        rptDocPorPC.Sections("sección4").Controls("lblEntidadNom").Caption = txtEnt.Text
        rptDocPorPC.Sections("sección4").Controls("lblReporte").Caption = "REPORTE DE DOCUMENTOS POR COBRAR"
    End If
        rptDocPorPC.Sections("sección4").Controls("lblemp").Caption = xNomEmp
        rptDocPorPC.Sections("sección4").Controls("lblruc").Caption = xNumRuc
        rptDocPorPC.Sections("sección5").Controls("totsol").Caption = TxtTotal.Text
        rptDocPorPC.Sections("sección5").Controls("totsal").Caption = TxtTotalSal.Text
        rptDocPorPC.Sections("sección5").Controls("lblTotReg").Caption = TxtNumReg.Text
    
    Set rptDocPorPC.DataSource = RsTemp
    Set RsTemp = Nothing
    rptDocPorPC.Orientation = rptOrientLandscape
    rptDocPorPC.Show

End Sub

Private Sub txtEnt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        txtEnt.Text = ""
        lblEnt.Caption = ""
    End If
End Sub
