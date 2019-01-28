VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmImpCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas - Importar Clientes"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6435
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1200
      Left            =   3090
      TabIndex        =   0
      Top             =   2865
      Visible         =   0   'False
      Width           =   5805
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   300
         Left            =   150
         TabIndex        =   1
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
         TabIndex        =   2
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
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6450
      Left            =   30
      TabIndex        =   3
      Top             =   1065
      Width           =   11790
      _cx             =   20796
      _cy             =   11377
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   21
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmImpCliente.frx":0000
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":0286
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":07CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":0B5C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":0CE0
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":1134
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":124C
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":1790
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":1CD4
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":1DE8
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":1EFC
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":2350
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":24BC
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmImpCliente.frx":2A04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
            ImageIndex      =   13
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Crear Formato a Importar"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   750
      Left            =   30
      TabIndex        =   4
      Top             =   285
      Width           =   11805
      Begin VB.CommandButton CmdBusFile 
         Enabled         =   0   'False
         Height          =   240
         Left            =   7140
         Picture         =   "FrmImpCliente.frx":2D96
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton CmdAgrupar 
         Caption         =   "Ordenar"
         Enabled         =   0   'False
         Height          =   435
         Left            =   10125
         TabIndex        =   5
         Top             =   195
         Width           =   1620
      End
      Begin VB.TextBox TxtArchivo 
         Height          =   300
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "TxtArchivo"
         Top             =   285
         Width           =   6015
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar"
         Enabled         =   0   'False
         Height          =   435
         Left            =   8520
         TabIndex        =   6
         Top             =   195
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivo"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   330
         Width           =   540
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   8310
         X2              =   8310
         Y1              =   150
         Y2              =   700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   8325
         X2              =   8325
         Y1              =   135
         Y2              =   700
      End
   End
End
Attribute VB_Name = "FrmImpCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean

Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(4).Enabled = Not Toolbar1.Buttons(4).Enabled
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(7).Enabled = Not Toolbar1.Buttons(7).Enabled
End Sub

Sub Blanquea()
    TxtArchivo.Text = ""
End Sub

Sub Bloquea()
    CmdBusFile.Enabled = Not CmdBusFile.Enabled
    CmdCargar.Enabled = Not CmdCargar.Enabled
    CmdAgrupar.Enabled = Not CmdAgrupar.Enabled
End Sub

Private Sub CmdAgrupar_Click()
    GRID_ORDENAR Fg1, 1, 1, , , flexSortNumericAscending
End Sub

Private Sub CmdBusFile_Click()
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

Sub Cancelar()
    ActivaTool
    Blanquea
    Bloquea
    QueHace = 3
    MuestraDatos
End Sub

Sub Modificar()
'    Dim Rst As New ADODB.Recordset
'
'    RST_Busq Rst, "SELECT var_importados.iddoc, var_importados.idtabla From var_importados WHERE (((var_importados.idtabla)=3))", xCon
'    If Rst.RecordCount <> 0 Then
'        MsgBox "Ya se importaron datos de clientes, para importar nuevos datos elimine los datos previos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Set Rst = Nothing
'        Exit Sub
'    End If
'    Set Rst = Nothing
    
    ActivaTool
    Fg1.TextMatrix(0, 1) = "Pendiente"
    Blanquea
    Bloquea
    QueHace = 2
    Fg1.Rows = 1
End Sub

Sub CargaDocumentos()
    Dim A&
    Dim B&
    Dim xNumFilas&
    On Error GoTo error
    Dim objExcel As Object
    
    Me.MousePointer = vbHourglass
    
    Set objExcel = CreateObject("Excel.Application")
    'Dim objExcel As New Excel.Application
    
    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo.Text)
    
    Frame2.Left = 3090
    Frame2.Top = 2910
    Label4.Caption = "Cargando registros para la importación"
    Frame2.Visible = True
    
    xNumFilas = 1
    
    Fg1.Rows = 1
    With objExcel.ActiveSheet
        
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        A = 4
        Do While NulosC(.Cells(A, 1)) <> ""
            If NulosC(.Cells(A, 1)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit Do
            End If
            A = A + 1
        Loop
        
        Fg1.Rows = 1
        xNumFilas = xNumFilas + 1
        ProgressBar2.Max = xNumFilas
        A = 4
        
        Do While NulosC(.Cells(A, 1)) <> ""
        
            ProgressBar2.Value = A - 3
            DoEvents
            
            Fg1.Rows = Fg1.Rows + 1
                '--tipo persona
                If NulosN(.Cells(A, 1)) <> 0 Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = Busca_Codigo(NulosN(.Cells(A, 1)), "id", "descripcion", "mae_tipoempresa", "N", xCon) 'Trim(.Cells(A, B))
                    Fg1.TextMatrix(Fg1.Rows - 1, 17) = NulosN(.Cells(A, 1))
                End If
                '--tipo doc
                If NulosN(.Cells(A, 2)) <> 0 Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = Busca_Codigo(NulosN(.Cells(A, 2)), "id", "descripcion", "mae_dociden", "N", xCon)  'Trim(.Cells(A, B))
                    Fg1.TextMatrix(Fg1.Rows - 1, 18) = NulosN(.Cells(A, 2))
                End If
                '--ruc
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(.Cells(A, 3))
                Fg1.TextMatrix(Fg1.Rows - 1, 16) = Busca_Codigo(NulosC(.Cells(A, 3)), "numruc", "id", "mae_cliente", "C", xCon)    'Trim(.Cells(A, B))
                If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 16)) = 0 Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = -1
                End If
                
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = .Cells(A, 4) '--proveedor
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(.Cells(A, 5)) 'Nombre 1
                Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(.Cells(A, 6)) 'Nombre 2
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(.Cells(A, 7)) 'Apellido 1
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(.Cells(A, 8)) 'Apellido 2
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(.Cells(A, 9)) 'direccion
                
                'departamento
                If NulosC(.Cells(A, 10)) <> "" Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 11) = Busca_Codigo(NulosN(.Cells(A, 10)), "id", "descripcion", "mae_departamentos", "N", xCon)
                    Fg1.TextMatrix(Fg1.Rows - 1, 19) = NulosC(.Cells(A, 10))
                End If
                
                'Distrito
                If NulosC(.Cells(A, 11)) <> "" Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 12) = Busca_Codigo(NulosN(.Cells(A, 11)), "id", "descripcion", "mae_distritos", "N", xCon)
                    Fg1.TextMatrix(Fg1.Rows - 1, 20) = NulosC(.Cells(A, 11))
                End If
                
                
                
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosC(.Cells(A, 12)) 'telefono
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosC(.Cells(A, 13)) 'fax
                Fg1.TextMatrix(Fg1.Rows - 1, 15) = NulosC(.Cells(A, 14)) 'email
                DoEvents
            A = A + 1
        Loop
    End With
    
    Label4.Caption = "Revizando datos Duplicados"
    DoEvents
    '--validar duplicados
    If Fg1.Rows > Fg1.FixedRows Then
        ProgressBar2.Max = Fg1.Rows - 1
        For A = 1 To Fg1.Rows - 1
            ProgressBar2.Value = A
            DoEvents
            If Fg1.FindRow(Fg1.TextMatrix(A, 4), A + 1, 4, False, False) <> -1 Then
                '--colocar como duplicado
                Fg1.TextMatrix(A, 1) = 0
                Fg1.TextMatrix(A, 20) = "Duplicado"
            End If
        Next A
    End If
    Frame2.Visible = False
    
    MsgBox "El proceso termino de cargar los datos con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 2
    objExcel.Workbooks.Close
    
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Frame2.Visible = False
    Me.MousePointer = vbDefault
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "CargaDocumentos"
End Sub

Private Sub CmdCargar_Click()
    If TxtArchivo.Text = "" Then
        MsgBox "No ha especificado el nombre del archivo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtArchivo.SetFocus
        Exit Sub
    End If
    CargaDocumentos
''''CargaDetraccionXLS
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nSQL As String
    If Col <> 2 And Col <> 3 Then Exit Sub
    Select Case Col
        Case 2 '--tipo persona
                
                ReDim xCampos(2, 4) As String
                xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "4500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
                xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":           xCampos(1, 2) = "700":      xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
                
                nSQL = "SELECT mae_tipoempresa.id, mae_tipoempresa.descripcion FROM mae_tipoempresa;"
                nTitulo = "Buscando Tipo de Persona"
            
        Case Else
            Exit Sub
    End Select

    Dim xRs As New ADODB.Recordset
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    Select Case Col
        Case 2 '--tipo persona
            Fg1.TextMatrix(Fg1.Row, 2) = NulosC(xRs.Fields("descripcion"))
            Fg1.TextMatrix(Fg1.Row, 17) = NulosC(xRs.Fields("id"))
            
    End Select
SALIR:

    Set xRs = Nothing

Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Fg_CellButtonClick(" & Row & "," & Col & ")", True, xTitulo


End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col = 1 Or Fg1.Col = 3 Or NulosN(Fg1.TextMatrix(Fg1.Row, 1)) = 0 Or Fg1.Col = Fg1.Cols - 1 Then
        Fg1.Editable = flexEDNone
    Else
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        QueHace = 3
        MuestraDatos
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    '--ocultar las columnas de los id's
    OCULTAR_COL Fg1, 16, 20
    GRID_COMBOLIST Fg1, 2 '--tipo persona
    GRID_COMBOLIST Fg1, 3 '--tipo doc
    Fg1.FrozenCols = 5
    '----

    QueHace = 3
    Blanquea
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim A&
    Dim mCuentaRegEliminar As Double
        
    
    Rpta = MsgBox("¿Esta seguro de eliminar la Importación de los Clientes Seleccionados?", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo)
    If Rpta = vbYes Then
        
        Frame2.Left = 3090
        Frame2.Top = 2910
        Label4.Caption = "Eliminando registros Importados"
        ProgressBar2.Max = Fg1.Rows - 1
        Frame2.Visible = True
        mCuentaRegEliminar = 0
        For A = 1 To Fg1.Rows - 1
        
            ProgressBar2.Value = A
            DoEvents
            If NulosN(Fg1.TextMatrix(A, 1)) = -1 Then
                xCon.Execute "DELETE * FROM mae_cliente WHERE id = " & NulosN(Fg1.TextMatrix(A, 16)) & ""
                           
                xCon.Execute "DELETE * FROM var_importados WHERE idtabla = 3 and iddoc = " & NulosN(Fg1.TextMatrix(A, 16))
                
                mCuentaRegEliminar = mCuentaRegEliminar + 1
            End If
            
        Next A
        
        Frame2.Visible = False
        If mCuentaRegEliminar <> 0 Then
            MsgBox "Los registros se eliminaron con éxito" & vbCr & "Total Registros Eliminados: " & mCuentaRegEliminar, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            MuestraDatos
        Else
            MsgBox "No hay registros importados a eliminar", vbInformation + vbOKOnly + vbCritical, xTitulo
        End If
    End If
End Sub

Function Grabar() As Boolean
    If Fg1.Rows = 1 Then
        MsgBox "No se ha cargando ningún registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Grabar = False
        Exit Function
    End If
    If MsgBox("Seguro desea grabar los Clientes", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Function
    
    Dim A As Integer
    Dim rst As New ADODB.Recordset
    Dim RstImp As New ADODB.Recordset
    Dim mTotalGrabados As Double
    Dim xId As Integer
    
'    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    Me.MousePointer = vbHourglass
    
    RST_Busq rst, "SELECT TOP 1 * FROM mae_cliente", xCon
    RST_Busq RstImp, "SELECT * FROM var_importados", xCon
    Frame2.Left = 3090
    Frame2.Top = 2910
    
    Label4.Caption = "Importando registros"
    
    ProgressBar2.Max = Fg1.Rows - 1
    Frame2.Visible = True
    mTotalGrabados = 0
    For A = 1 To Fg1.Rows - 1
        ProgressBar2.Value = A
        DoEvents
        If NulosN(Fg1.TextMatrix(A, 1)) = -1 Then
            mTotalGrabados = mTotalGrabados + 1
            xId = HallaCodigoTabla("mae_cliente", xCon, "id")
            
            rst.AddNew
           
            rst("id") = xId
            rst("tipper") = NulosN(Fg1.TextMatrix(A, 17))
            rst("idtipdoc") = NulosN(Fg1.TextMatrix(A, 18))
            rst("numruc") = NulosC(Trim(Fg1.TextMatrix(A, 4)))
            rst("nombre") = Left(NulosC(Fg1.TextMatrix(A, 5)), rst("nombre").DefinedSize)
            rst("nomcli1") = Left(NulosC(Fg1.TextMatrix(A, 6)), rst("nomcli1").DefinedSize)
            rst("nomcli2") = Left(NulosC(Fg1.TextMatrix(A, 7)), rst("nomcli2").DefinedSize)
            rst("apecli1") = Left(NulosC(Fg1.TextMatrix(A, 8)), rst("apecli1").DefinedSize)
            rst("apecli2") = Left(NulosC(Fg1.TextMatrix(A, 9)), rst("apecli2").DefinedSize)
            
            rst("dir") = NulosC(Fg1.TextMatrix(A, 10))
            rst("iddep") = NulosN(Fg1.TextMatrix(A, 19))
            rst("iddis") = NulosN(Fg1.TextMatrix(A, 20))
            rst("tel") = NulosC(Fg1.TextMatrix(A, 13))
            rst("fax") = NulosC(Fg1.TextMatrix(A, 14))
            rst("email") = NulosC(Fg1.TextMatrix(A, 15))
            rst("activo") = -1
            
            
            
            rst.Update
            
            RstImp.AddNew
            RstImp("iddoc") = xId
            RstImp("idtabla") = 3
            RstImp("idmes") = 0
            RstImp.Update
            
        End If
    Next A
    Frame2.Visible = False
    
    Me.MousePointer = vbDefault
    Set rst = Nothing
    Set RstImp = Nothing
    
    xCon.CommitTrans
    If mTotalGrabados = 0 Then
        MsgBox "Ningún Cliente a sido importado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Else
        MsgBox "Los datos se importaron con éxito" & vbCr & "Total Registros: " & mTotalGrabados, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    
    Fg1.Rows = 1
    Grabar = True
    Exit Function
LaCague:
    Frame2.Visible = False
    Me.MousePointer = vbDefault
    xCon.RollbackTrans
    Set rst = Nothing
    Set RstImp = Nothing
    MsgBox "No se pudo Importar por el siguiente motivo :" + Trim(Err.Description)
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este Listo para Importar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    Else
        
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Modificar
    
    If Button.Index = 2 Then Eliminar
    
    If Button.Index = 4 Then Cancelar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Blanquea
            Cancelar
        End If
    End If
    
    If Button.Index = 9 Then
        Unload Me
    End If
End Sub

Sub MuestraDatos()
    Dim rst As New ADODB.Recordset
    Dim A&
    Dim nSQL As String
    nSQL = "SELECT cli.*,reg.registros FROM " _
        + vbCr + " (SELECT mae_cliente.*, mae_tipoempresa.descripcion AS desctipper, mae_dociden.descripcion AS desctipdoc, mae_distritos.descripcion AS descdis, mae_departamentos.descripcion AS descdep FROM ((((var_importados LEFT JOIN mae_cliente ON var_importados.iddoc = mae_cliente.id) LEFT JOIN mae_tipoempresa ON mae_cliente.tipper = mae_tipoempresa.id) LEFT JOIN mae_distritos ON mae_cliente.iddis = mae_distritos.id) LEFT JOIN mae_departamentos ON mae_cliente.iddep = mae_departamentos.id) LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id WHERE (((var_importados.idtabla)=3)) ) AS cli " _
        + vbCr + " LEFT JOIN " _
        + vbCr + " (SELECT vta_ventas.idcli, Count(vta_ventas.id) AS registros FROM vta_ventas GROUP BY vta_ventas.idcli) AS reg " _
        + vbCr + " ON cli.id= reg.idcli"

    RST_Busq rst, nSQL, xCon
        
    Fg1.Rows = 1
    
    If rst.RecordCount <> 0 Then
        Fg1.TextMatrix(0, 1) = "Eliminar"
        rst.MoveFirst
    
        For A = 1 To rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            If NulosN(rst("registros")) <> 0 Then
                Fg1.TextMatrix(A, 1) = 0 '--no se puede eliminar
            Else
                Fg1.TextMatrix(A, 1) = -1 '--si se puede eliminar
            End If
            Fg1.TextMatrix(A, 2) = NulosC(rst("desctipper"))
            Fg1.TextMatrix(A, 3) = NulosC(rst("desctipdoc"))
            Fg1.TextMatrix(A, 4) = NulosC(rst("numruc"))
            Fg1.TextMatrix(A, 5) = NulosC(rst("nombre"))
            Fg1.TextMatrix(A, 6) = NulosC(rst("nomcli1"))
            Fg1.TextMatrix(A, 7) = NulosC(rst("nomcli2"))
            Fg1.TextMatrix(A, 8) = NulosC(rst("apecli1"))
            Fg1.TextMatrix(A, 9) = NulosC(rst("apecli2"))
            Fg1.TextMatrix(A, 10) = NulosC(rst("dir"))
            Fg1.TextMatrix(A, 11) = NulosC(rst("descdep"))
            Fg1.TextMatrix(A, 12) = NulosC(rst("descdis"))
            Fg1.TextMatrix(A, 13) = NulosC(rst("tel"))
            Fg1.TextMatrix(A, 14) = NulosC(rst("fax"))
            Fg1.TextMatrix(A, 15) = NulosC(rst("email"))
            
            Fg1.TextMatrix(A, 16) = NulosN(rst("id"))
            Fg1.TextMatrix(A, 17) = NulosN(rst("tipper"))
            Fg1.TextMatrix(A, 18) = NulosN(rst("idtipdoc"))
            Fg1.TextMatrix(A, 19) = NulosN(rst("iddis"))
            Fg1.TextMatrix(A, 20) = NulosN(rst("iddep"))
            
            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
    End If
    Set rst = Nothing
End Sub


Private Sub pCrearFormato()

    On Error GoTo error
    Dim objExcel As Object
    Dim k As Integer
    
    Set objExcel = CreateObject("Excel.Application")
    objExcel.SheetsInNewWorkbook = 1
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    
    With objExcel.ActiveSheet
        .Cells(1, 1) = "Importar Clientes"
        
        .Cells(3, 1) = "Tipo de Persona"
        .Cells(3, 2) = "Tipo Documento"
        .Cells(3, 3) = "Nº Documento"
        .Cells(3, 4) = "Cliente"
        .Cells(3, 5) = "Nombre 1"
        .Cells(3, 6) = "Nombre 2"
        .Cells(3, 7) = "Apellido 1"
        .Cells(3, 8) = "Apellido 2"
        .Cells(3, 9) = "Dirección"
        .Cells(3, 10) = "Departamento"
        .Cells(3, 11) = "Distrito"
        .Cells(3, 12) = "Teléfono"
        .Cells(3, 13) = "Fax"
        .Cells(3, 14) = "email"
        '---------
        .Columns(1).ColumnWidth = 8
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 15
        '--establecer el ancho
        
        .Columns(1).ColumnWidth = 15
        .Columns(2).ColumnWidth = 15
        .Columns(3).ColumnWidth = 12
        .Columns(4).ColumnWidth = 25.5
        
        .Columns(9).ColumnWidth = 31
        
        For k = 1 To 14
            .Cells(3, k).Font.Bold = True
        Next

                
    End With
    MsgBox "Proceda a ingresar la información según los Parámetros Solicitados" + vbCr + "Luego proceda a Importar...", vbInformation, xTitulo
    objExcel.Visible = True
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then pCrearFormato
End Sub

Private Sub TxtArchivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = vbKeyF5 Then CmdBusFile_Click
End Sub



Sub CargaDetraccionXLS()
    '--28/09/09
    '--archivo proporcionado por mario
    Dim A&
    Dim B&
    Dim xNumFilas&
    On Error GoTo error
    Dim objExcel As Object
    
    Me.MousePointer = vbHourglass
    
    Set objExcel = CreateObject("Excel.Application")
    'Dim objExcel As New Excel.Application
    
    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo.Text)
    
    Frame2.Left = 3090
    Frame2.Top = 2910
    Label4.Caption = "Cargando registros para la importación"
    Frame2.Visible = True
    
    'xNumFilas = 1
    
    Dim RstCab As New ADODB.Recordset
        
    RST_Busq RstCab, "select top 1 * FROM zzz_detraccion", xCon
    
    xNumFilas = 985

    With objExcel.ActiveSheet
        

        
        
        ProgressBar2.Max = xNumFilas
        ProgressBar2.Value = 1
        A = 1
        
        Dim xFchRecepcion As String
        Dim xFchPago As String
        Dim xEmpresa As String
 
        
        xEmpresa = "--"
        xFchPago = "--"
        xFchRecepcion = "--"
                
        
        Do While A < xNumFilas
            ProgressBar2 = A + 1
            If NulosN(.Cells(A, 7)) <> 0 Then
                RstCab.AddNew
                RstCab("empresa") = xEmpresa
                RstCab("fchrecepcion") = xFchRecepcion
                
                RstCab("fchpago") = xFchPago
                RstCab("numdocu") = NulosC(.Cells(A, 1))
                RstCab("fchdoc") = NulosC(.Cells(A, 2))
                RstCab("proveedor") = NulosC(.Cells(A, 3))
                RstCab("impdol") = NulosN(.Cells(A, 4))
                RstCab("imptc") = NulosN(.Cells(A, 5))
                RstCab("impsol") = NulosN(.Cells(A, 6))
                RstCab("impdetra") = NulosN(.Cells(A, 7))
                
                RstCab.Update
            End If
            A = A + 1
            
            If UCase(.Cells(A, 4)) = "RECIBIDO POR" Then
                If InStr(NulosC(.Cells(A + 5, 2)), "SAVAR") Or InStr(NulosC(.Cells(A + 6, 2)), "DEPOT") Then
                    xEmpresa = NulosC(.Cells(A + 5, 2))
                    xFchRecepcion = NulosC(.Cells(A + 8, 1))
                ElseIf InStr(NulosC(.Cells(A + 7, 2)), "SAVAR") Or InStr(NulosC(.Cells(A + 7, 2)), "DEPOT") Then
                    xEmpresa = NulosC(.Cells(A + 7, 2))
                    xFchRecepcion = NulosC(.Cells(A + 10, 1))
                
                ElseIf InStr(NulosC(.Cells(A + 8, 2)), "SAVAR") Or InStr(NulosC(.Cells(A + 8, 2)), "DEPOT") Then
                    xEmpresa = NulosC(.Cells(A + 8, 2))
                    xFchRecepcion = NulosC(.Cells(A + 11, 1))
                
                Else
                    xEmpresa = NulosC(.Cells(A + 7, 2))
                    xFchRecepcion = NulosC(.Cells(A + 10, 1))
                
                End If
                
                xFchPago = NulosC(.Cells(A + 1, 1))
                
            End If
            
            
            
        Loop
        
    End With
    
    
    DoEvents
    Frame2.Visible = False
    
    MsgBox "El proceso termino de cargar los datos con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 2
    objExcel.Workbooks.Close
    
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    TxtArchivo.Text = ""
    Exit Sub
error:
    Frame2.Visible = False
    Me.MousePointer = vbDefault
    Set objExcel = Nothing
    MsgBox Err.Description
    Err.Clear
End Sub

