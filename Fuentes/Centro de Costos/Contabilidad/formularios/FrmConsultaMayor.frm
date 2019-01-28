VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaMayor 
   Caption         =   "Contabilidad - Mayor"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1200
      Left            =   3420
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   5805
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   300
         Left            =   150
         TabIndex        =   8
         Top             =   615
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label LblProcesando 
         AutoSize        =   -1  'True
         Caption         =   "LblProcesando"
         Height          =   195
         Left            =   1845
         TabIndex        =   10
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Cuenta  : "
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   300
         Width           =   1590
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
      Height          =   975
      Left            =   15
      TabIndex        =   1
      Top             =   90
      Width           =   11970
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3285
         TabIndex        =   13
         Top             =   495
         Width           =   900
      End
      Begin VB.OptionButton OptDolares 
         Caption         =   "Dolares"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4245
         TabIndex        =   12
         Top             =   495
         Width           =   900
      End
      Begin VB.CommandButton CmdImprimir 
         Height          =   555
         Left            =   10815
         Picture         =   "FrmConsultaMayor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   270
         Width           =   630
      End
      Begin VB.CommandButton CmdMuestra 
         Height          =   555
         Left            =   10155
         Picture         =   "FrmConsultaMayor.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   630
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
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
         Left            =   1200
         TabIndex        =   3
         Top             =   540
         Width           =   1245
         _ExtentX        =   2196
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
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   2865
         X2              =   2865
         Y1              =   240
         Y2              =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Final"
         Height          =   195
         Left            =   105
         TabIndex        =   5
         Top             =   570
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Inicio"
         Height          =   195
         Left            =   105
         TabIndex        =   4
         Top             =   270
         Width           =   735
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6270
      Left            =   15
      TabIndex        =   0
      Top             =   1095
      Width           =   11955
      _cx             =   21087
      _cy             =   11060
      _ConvInfo       =   1
      Appearance      =   0
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsultaMayor.frx":074C
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
      ExplorerBar     =   2
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
Attribute VB_Name = "FrmConsultaMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstMayor As New ADODB.Recordset  'recorset para cargar las cuentas principales del
Dim RstTmp As New ADODB.Recordset  'recorset temporal
Dim SeEjecuto  As Boolean

Sub CargarMayor()
    Dim cadena As String
    Dim A As Integer
        
    PreparaRST_Tmp
    
    RST_Busq RstMayor, "SELECT con_planctas.id, con_planctas.cuenta, con_planctas.descripcion, " _
        & " (SELECT Sum([impdebsol]) AS total From con_diario WHERE (((con_diario.idcue)=con_planctas.id) AND " _
        & " ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND " _
        & " ((con_diario.idmes)<>0))) AS totdeb, " _
        & " (SELECT Sum([imphabsol]) AS total From con_diario WHERE (((con_diario.idcue)=con_planctas.id) AND " _
        & " ((con_diario.idmes)<>0) AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "')))) AS tothab " _
        & " From con_planctas " _
        & " WHERE ((((SELECT Sum([impdebsol]) AS total From con_diario WHERE (((con_diario.idcue)=con_planctas.id) " _
        & " AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "')) AND " _
        & " ((con_diario.idmes)<>0)))) Is Not Null) AND " _
        & " (((SELECT Sum([imphabsol]) AS total From con_diario WHERE (((con_diario.idcue)=con_planctas.id) " _
        & " AND ((con_diario.idmes)<>0) AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))))) Is Not Null)) " _
        & " ORDER BY con_planctas.cuenta", xCon
    
    'RstMayor.Filter = "cuenta like 10%"
    If RstMayor.RecordCount <> 0 Then
        RstMayor.MoveFirst
        
        For A = 1 To RstMayor.RecordCount
            cadena = cadena + Trim(RstMayor("cuenta"))
            RstMayor.MoveNext
            If RstMayor.EOF = True Then Exit For
            cadena = cadena + "|"
        Next A
        
        CargarDetalle
    
        Fg1.Rows = 1
        RstTmp.MoveFirst
        For A = 1 To RstTmp.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = RstTmp("numcue")
            Fg1.TextMatrix(A, 2) = Format(RstTmp("fchmov"), "dd//mm/yy")
            Fg1.TextMatrix(A, 3) = RstTmp("numasi")
            Fg1.TextMatrix(A, 4) = RstTmp("descri")
            Fg1.TextMatrix(A, 5) = Format(RstTmp("totdeb"), "0.00")
            Fg1.TextMatrix(A, 6) = Format(RstTmp("tothab"), "0.00")
          
            Fg1.TextMatrix(A, 7) = Format(RstTmp("impsal"), "0.00")
            Fg1.TextMatrix(A, 8) = Format(RstTmp("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(A, 9) = RstTmp("numdoc")
            Fg1.TextMatrix(A, 10) = RstTmp("nompro")
            
            RstTmp.MoveNext
            If RstTmp.EOF = True Then
                Exit For
            End If
        Next A
        
        Fg1.Redraw = False
        
        Fg1.ExplorerBar = flexExMove
        Fg1.OutlineBar = 1
        Fg1.ColComboList(1) = cadena
        Fg1.Subtotal flexSTSum, -1, 5, , , RGB(255, 0, 0), False, "TOTAL ==>"
        Fg1.Subtotal flexSTSum, -1, 6, , , RGB(255, 0, 0), False
        
        Fg1.Subtotal flexSTSum, 1, 5, , , &H800000, False, "Cta Nº %s"
        Fg1.Subtotal flexSTSum, 1, 6, , , &H800000, True
        Fg1.Redraw = True
        
        Dim xCad As String
        For A = 1 To Fg1.Rows - 1
            If Mid(Fg1.TextMatrix(A, 1), 1, 3) = "Cta" Then
                xCad = Trim(Mid(Fg1.TextMatrix(A, 1), 8, 20))
                Fg1.TextMatrix(A, 4) = Busca_Codigo(xCad, "cuenta", "descripcion", "con_planctas", "C", xCon)
            End If
        Next A
    End If
End Sub

Private Sub CmdMuestra_Click()
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    CargarMayor
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Fg1.Rows = 1
        SeEjecuto = True
    End If
End Sub

Sub CargarDetalle()
    Dim RstDetCta As New ADODB.Recordset 'recorset para hallar el detalle de la cuenta que se esta mostrando
    Dim RstDetCtaReg As New ADODB.Recordset 'recorset para hallar el registro del detalle de la cuenta que se esta mostrando
    Dim A, B As Integer
    Dim xSaldo As Double
    RstMayor.MoveFirst
    
    LblProcesando.Caption = ""
    Frame2.Visible = True
    ProgressBar2.Max = RstMayor.RecordCount
    
    For A = 1 To RstMayor.RecordCount
        LblProcesando.Caption = RstMayor("descripcion")
        ProgressBar2.Value = A
        Frame2.Refresh
        
        'cargamos el detalle de la cuenta
        'RST_Busq RstDetCta, "SELECT con_diario.idcue, con_diario.idlib, con_diario.numasi, con_diario.impdebsol, " _
            & " con_diario.imphabsol, con_cajabanco.tipmov, con_diario.idmov FROM con_diario LEFT JOIN con_cajabanco ON " _
            & " con_diario.idmov = con_cajabanco.id WHERE (((con_diario.idcue)=" & RstMayor("id") & "))", xCon
        RST_Busq RstDetCta, "SELECT con_diario.idcue, con_diario.idlib, con_diario.numasi, con_diario.impdebsol, " _
            & " con_diario.imphabsol, con_cajabanco.tipmov, con_diario.idmov, con_diario.fchasi " _
            & " FROM con_diario LEFT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id " _
            & " WHERE (((con_diario.idcue) = " & RstMayor("id") & ") AND ((con_diario.fchasi)>= CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<= CDate('" & TxtFchFin.Valor & "')))", xCon

        RstDetCta.MoveFirst
        xSaldo = 0
        
        If RstDetCta.RecordCount <> 0 Then
            For B = 1 To RstDetCta.RecordCount
                If RstDetCta("idlib") = 1 Then
                    'LIBRO COMPRAS
                    RST_Busq RstDetCtaReg, "SELECT con_diario.idcue, con_diario.idmov, con_diario.idlib, con_diario.numasi, " _
                        & " con_diario.impdebsol, con_diario.imphabsol, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, " _
                        & " com_compras.fchdoc AS fchmov, com_compras.fchdoc, mae_prov.nombre FROM mae_prov RIGHT JOIN " _
                        & " (con_diario LEFT JOIN com_compras ON con_diario.idmov = com_compras.id) ON mae_prov.id = com_compras.idpro " _
                        & " WHERE (((con_diario.idcue)=" & RstDetCta("idcue") & ") AND ((con_diario.idmov)=" & RstDetCta("idmov") & ") " _
                        & " AND ((con_diario.idlib)=1))", xCon
                End If
                
                If RstDetCta("idlib") = 2 Then
                    'LIBRO VENTAS
                    RST_Busq RstDetCtaReg, "SELECT con_diario.idcue, con_diario.idmov, con_diario.idlib, con_diario.numasi, " _
                        & " con_diario.impdebsol, con_diario.imphabsol, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, " _
                        & " vta_ventas.fchdoc AS fchmov, vta_ventas.fchdoc, mae_cliente.nombre " _
                        & " FROM con_diario LEFT JOIN (vta_ventas LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) " _
                        & " ON con_diario.idmov = vta_ventas.id WHERE (((con_diario.idcue)=" & RstDetCta("idcue") & ") " _
                        & " AND ((con_diario.idmov)=" & RstDetCta("idmov") & ") AND ((con_diario.idlib)=2))", xCon

                End If
                
                If RstDetCta("idlib") = 6 Then
                    'LIBRO CAJA Y BANCOS
                    If RstDetCta("tipmov") = 1 Then  'si es un ingreso buscamos en la tabla vta_ventas, notas de credito, etc
                        RST_Busq RstDetCtaReg, "SELECT con_diario.idcue, con_cajabanco.tipmov, con_diario.idmov, " _
                            & " con_diario.idlib, con_diario.numasi, con_diario.impdebsol, con_diario.imphabsol, " _
                            & " con_cajabancodet.iddoc, con_cajabancodet.idorigen FROM (con_diario LEFT JOIN con_cajabanco " _
                            & " ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_cajabancodet ON " _
                            & " con_cajabanco.id = con_cajabancodet.id WHERE (((con_diario.idcue)=" & RstDetCta("idcue") & ") " _
                            & " AND ((con_cajabanco.tipmov)=1) AND ((con_diario.idmov)=" & RstDetCta("idmov") & ") " _
                            & " AND ((con_diario.idlib)=6))", xCon
                        
                        If RstDetCtaReg("idorigen") = 4 Then   'buscamos el numero de documento en ventas tabla (vta_ventas)
                            RST_Busq RstDetCtaReg, "SELECT con_diario.idcue, con_cajabanco.tipmov, con_diario.idmov, con_cajabanco.fchope AS fchmov," _
                                & " con_diario.idlib, con_diario.numasi, con_diario.impdebsol, con_diario.imphabsol, " _
                                & " con_cajabancodet.iddoc, con_cajabancodet.idorigen, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, " _
                                & " vta_ventas.fchdoc, vta_ventas.fchven, mae_cliente.nombre, mae_cliente.numruc " _
                                & " FROM (((con_diario LEFT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id) " _
                                & " LEFT JOIN con_cajabancodet ON con_cajabanco.id = con_cajabancodet.id) LEFT JOIN vta_ventas " _
                                & " ON con_cajabancodet.iddoc = vta_ventas.id) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id " _
                                & " WHERE (((con_diario.idcue)=" & RstDetCta("idcue") & ") AND ((con_cajabanco.tipmov)=1) AND " _
                                & " ((con_diario.idmov)=" & RstDetCta("idmov") & ") AND ((con_diario.idlib)=6) " _
                                & " AND ((con_cajabancodet.idorigen)=4))", xCon
                        End If
                    End If
                    
                    If RstDetCta("tipmov") = 2 Then  'si es un egreso buscamos en la tablas de salidas compras, pèrcepciom, detraccion, notas de credito,
                        RST_Busq RstDetCtaReg, "SELECT con_diario.idcue, con_cajabanco.tipmov, con_diario.idmov, con_cajabanco.fchope AS fchmov," _
                            & " con_diario.idlib, con_diario.numasi, con_diario.impdebsol, con_diario.imphabsol, " _
                            & " con_cajabancodet.iddoc, con_cajabancodet.idorigen FROM (con_diario LEFT JOIN con_cajabanco " _
                            & " ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_cajabancodet ON " _
                            & " con_cajabanco.id = con_cajabancodet.id WHERE (((con_diario.idcue)=" & RstDetCta("idcue") & ") " _
                            & " AND ((con_cajabanco.tipmov)=2) AND ((con_diario.idmov)=" & RstDetCta("idmov") & ") " _
                            & " AND ((con_diario.idlib)=6))", xCon
                        
                        If RstDetCtaReg("idorigen") = 1 Then   'buscamos el numero de documento en compras tabla (com_compras)
                            RST_Busq RstDetCtaReg, "SELECT con_diario.idcue, con_cajabanco.tipmov, con_diario.idmov, con_cajabanco.fchope AS fchmov," _
                                & " con_diario.idlib, con_diario.numasi, con_diario.impdebsol, con_diario.imphabsol, " _
                                & " con_cajabancodet.iddoc, con_cajabancodet.idorigen, mae_prov.numruc, mae_prov.nombre, " _
                                & " [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchdoc, " _
                                & " com_compras.fchven FROM ((con_diario LEFT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id) " _
                                & " LEFT JOIN con_cajabancodet ON con_cajabanco.id = con_cajabancodet.id) " _
                                & " LEFT JOIN (mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro) " _
                                & " ON con_cajabancodet.iddoc = com_compras.id WHERE (((con_diario.idcue)=" & RstDetCta("idcue") & ") " _
                                & " AND ((con_cajabanco.tipmov)=2) AND ((con_diario.idmov)=" & RstDetCta("idmov") & ") AND ((con_diario.idlib)=6) " _
                                & " AND ((con_cajabancodet.idorigen)=1))", xCon
                        End If
                        
                        If RstDetCtaReg("idorigen") = 2 Then   'buscamos el numero de documento en percepciones tabla (con_percepciones)
                            RST_Busq RstDetCtaReg, "SELECT con_diario.idcue, con_cajabanco.tipmov, con_diario.idmov, con_cajabanco.fchope AS fchmov," _
                                & " con_diario.idlib, con_diario.numasi, con_diario.impdebsol, con_diario.imphabsol, " _
                                & " con_cajabancodet.iddoc, con_cajabancodet.idorigen, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc, " _
                                & " con_percepcion.fchdoc, mae_prov.numruc, mae_prov.nombre FROM (((con_diario LEFT JOIN " _
                                & " con_cajabanco ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_cajabancodet " _
                                & " ON con_cajabanco.id = con_cajabancodet.id) LEFT JOIN con_percepcion ON con_cajabancodet.iddoc = con_percepcion.id) " _
                                & " LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id WHERE " _
                                & " (((con_diario.idcue)=" & RstDetCta("idcue") & ") AND ((con_cajabanco.tipmov)=2) " _
                                & " AND ((con_diario.idmov)=" & RstDetCta("idmov") & ") AND ((con_diario.idlib) = 6) " _
                                & " AND ((con_cajabancodet.idorigen)=2))", xCon
                        End If
                    End If
                End If
                
                RstTmp.AddNew
                RstTmp("numcue") = RstMayor("cuenta")
                RstTmp("descri") = RstMayor("descripcion")
                RstTmp("idcuen") = RstMayor("id")
                RstTmp("totdeb") = RstDetCtaReg("impdebsol")
                RstTmp("tothab") = RstDetCtaReg("imphabsol")
                RstTmp("numdoc") = RstDetCtaReg("numdoc")
                RstTmp("fchmov") = RstDetCtaReg("fchmov")
                If RstDetCtaReg("impdebsol") <> 0 Then xSaldo = xSaldo + RstDetCtaReg("impdebsol")
                If RstDetCtaReg("imphabsol") <> 0 Then xSaldo = xSaldo - RstDetCtaReg("imphabsol")
                
                RstTmp("fchdoc") = RstDetCtaReg("fchdoc")
                RstTmp("nompro") = RstDetCtaReg("nombre")
                RstTmp("numasi") = RstDetCtaReg("numasi")
                
                RstTmp("impsal") = xSaldo
                RstTmp.Update
                
                RstDetCta.MoveNext
                If RstDetCta.EOF = True Then Exit For
            Next B
        End If
        
        RstMayor.MoveNext
        If RstMayor.EOF = True Then Exit For
    Next A
    Frame2.Visible = False
End Sub

Sub PreparaRST_Tmp()
    Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xCampos(10, 3) As String

    xCampos(0, 0) = "numcue":        xCampos(0, 1) = "C":      xCampos(0, 2) = "10"
    xCampos(1, 0) = "descri":        xCampos(1, 1) = "C":      xCampos(1, 2) = "100"
    xCampos(2, 0) = "idcuen":        xCampos(2, 1) = "N":      xCampos(2, 2) = "1"
    xCampos(3, 0) = "totdeb":        xCampos(3, 1) = "D":      xCampos(3, 2) = "8"
    xCampos(4, 0) = "tothab":        xCampos(4, 1) = "D":      xCampos(4, 2) = "8"
    xCampos(5, 0) = "numdoc":        xCampos(5, 1) = "C":      xCampos(5, 2) = "15"
    xCampos(6, 0) = "fchmov":        xCampos(6, 1) = "F":      xCampos(6, 2) = "8"
    xCampos(7, 0) = "impsal":        xCampos(7, 1) = "D":      xCampos(7, 2) = "8"
    xCampos(8, 0) = "fchdoc":        xCampos(8, 1) = "F":      xCampos(8, 2) = "8"
    xCampos(9, 0) = "nompro":        xCampos(9, 1) = "C":      xCampos(9, 2) = "100"
    xCampos(10, 0) = "numasi":       xCampos(10, 1) = "C":     xCampos(10, 2) = "20"
    Set RstTmp = xFun.CrearRstTMP(xCampos)
    RstTmp.Open
End Sub

Private Sub Form_Load()
    Fg1.ColWidth(0) = 0
    'Fg1.ColWidth(1) = 2000
    Fg1.MergeCol(0) = True
    Fg1.MergeCol(1) = True
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    SeEjecuto = False
    
    TxtFchIni.Valor = "01/03/07"
    TxtFchFin.Valor = "31/03/07"
End Sub

