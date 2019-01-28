VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConfPlant 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Platillas de Documento - Configuracion"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   3900
      Left            =   15
      TabIndex        =   2
      Top             =   3570
      Width           =   11685
      _cx             =   20611
      _cy             =   6879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   " Cabecera Documento | Detalle Documento "
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   0
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   3480
         Left            =   45
         TabIndex        =   3
         Top             =   375
         Width           =   11595
         Begin VB.Frame Frame3 
            Height          =   3480
            Left            =   10365
            TabIndex        =   8
            Top             =   -15
            Width           =   1230
            Begin VB.CommandButton cmdAgregarCC 
               Caption         =   "Agregar Campo"
               Height          =   570
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   1170
               Width           =   1020
            End
            Begin VB.CommandButton cmdQuitarCC 
               Caption         =   "Quitar Campo"
               Height          =   570
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   1800
               Width           =   1020
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   3390
            Left            =   30
            TabIndex        =   4
            Top             =   75
            Width           =   10290
            _cx             =   18150
            _cy             =   5980
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
            SelectionMode   =   1
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
            FormatString    =   $"FrmConfPlant.frx":0000
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   3480
         Left            =   12330
         TabIndex        =   5
         Top             =   375
         Width           =   11595
         Begin VB.Frame Frame4 
            Height          =   3495
            Left            =   10365
            TabIndex        =   11
            Top             =   -15
            Width           =   1230
            Begin VB.CommandButton cmdAgregarCD 
               Caption         =   "Agregar Campo"
               Height          =   570
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   1170
               Width           =   1020
            End
            Begin VB.CommandButton cmdQuitarCD 
               Caption         =   "Quitar Campo"
               Height          =   570
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   1800
               Width           =   1020
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg3 
            Height          =   3390
            Left            =   15
            TabIndex        =   7
            Top             =   75
            Width           =   10290
            _cx             =   18150
            _cy             =   5980
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
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConfPlant.frx":0134
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
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   2730
      Left            =   15
      TabIndex        =   1
      Top             =   750
      Width           =   11670
      _cx             =   20585
      _cy             =   4815
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
      ForeColorSel    =   16777215
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
      Rows            =   1
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConfPlant.frx":0250
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "Imprimir"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10950
      Top             =   855
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":03C2
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":0906
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":0A8A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":0EDE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":0FF6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":153A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":1A7E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":1B92
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":1CA6
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":20FA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConfPlant.frx":2266
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Lista de Documentos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   105
      TabIndex        =   0
      Top             =   405
      Width           =   11505
   End
End
Attribute VB_Name = "FrmConfPlant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rstdoc As New ADODB.Recordset
Dim RstCab As New ADODB.Recordset
Dim Rstdet As New ADODB.Recordset
Dim Mostrando As Boolean
Dim QueHace As Integer
Dim Letras, Tamaño As String

Private Sub cmdAgregarCC_Click()
    Fg2.Rows = Fg2.Rows + 1
End Sub

Private Sub cmdAgregarCD_Click()
    Fg3.Rows = Fg3.Rows + 1
End Sub

Private Sub cmdQuitarCC_Click()
    Fg2.RemoveItem (Fg2.Row)
End Sub

Private Sub cmdQuitarCD_Click()
    Fg3.RemoveItem (Fg3.Row)
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    If Col = 1 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
        
        xform.SQLCad = "SELECT mae_documento.* FROM mae_documento"
        
        xform.Titulo = "Buscando Documento"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 10) = xRs("id")
            End If
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 2 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Serie":         xCampos(1, 1) = "numser":         xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
        
        If Fg1.TextMatrix(Fg1.Row, 10) = "" Then
            MsgBox "Seleccione primero el documento", vbInformation + vbOKOnly, xTitulo
            Exit Sub
        End If
        
        xform.SQLCad = "SELECT mae_series.id, mae_series.iddoc, mae_documento.descripcion, format(mae_series.numser, '0000') as numser " _
                     & " FROM mae_documento LEFT JOIN mae_series ON mae_documento.id = mae_series.iddoc " _
                     & " WHERE (((mae_series.iddoc)=" & Fg1.TextMatrix(Fg1.Row, 10) & "))"
        
        xform.Titulo = "Buscando Serie del Documento"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "descripcion"
        xform.CampoBusca = "descripcion"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 2) = Format(xRs("numser"), "0000")
                Fg1.TextMatrix(Fg1.Row, 11) = xRs("id")
            End If
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub fg1_RowColChange()
    Dim A As Integer
    If Mostrando <> True Then
        Set RstCab = Nothing
        RST_Mant RstCab, "SELECT * FROM var_plantillacab WHERE idplan = " & Fg1.TextMatrix(Fg1.Row, 9) & " ORDER BY var_plantillacab.numord ASC", xCon
        'Set DataGrid2.DataSource = RstCab
        Fg2.Rows = 1
        For A = 1 To RstCab.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = RstCab("item")
            Fg2.TextMatrix(A, 2) = RstCab("descripcion")
            Fg2.TextMatrix(A, 3) = RstCab("numord")
            Fg2.TextMatrix(A, 4) = RstCab("ancho")
            Fg2.TextMatrix(A, 5) = RstCab("posx")
            Fg2.TextMatrix(A, 6) = RstCab("posy")
            Fg2.TextMatrix(A, 7) = RstCab("campo")
            Fg2.TextMatrix(A, 8) = NulosC(RstCab("formato"))
            Fg2.TextMatrix(A, 9) = RstCab("tipo")
            RstCab.MoveNext
        Next A
        
        Set Rstdet = Nothing
        RST_Mant Rstdet, "SELECT * FROM var_plantilladet WHERE idplan = " & Fg1.TextMatrix(Fg1.Row, 9) & " ORDER BY var_plantilladet.numord ASC", xCon
        'Set DataGrid3.DataSource = RstDet
        
        Fg3.Rows = 1
        For A = 1 To Rstdet.RecordCount
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(A, 1) = Rstdet("item")
            Fg3.TextMatrix(A, 2) = Rstdet("descripcion")
            Fg3.TextMatrix(A, 3) = Rstdet("numord")
            Fg3.TextMatrix(A, 4) = Rstdet("ancho")
            Fg3.TextMatrix(A, 5) = Rstdet("posx")
            Fg3.TextMatrix(A, 6) = Rstdet("posy")
            Fg3.TextMatrix(A, 7) = Rstdet("campo")
            Fg3.TextMatrix(A, 8) = NulosC(Rstdet("formato"))
            Rstdet.MoveNext
        Next
        
    End If
End Sub

Private Sub Form_Activate()
    Dim A As Integer
    
    Set Rstdoc = Nothing
    Set RstCab = Nothing
    Set Rstdet = Nothing
    
    Mostrando = True
    RST_Busq Rstdoc, "SELECT var_plantilladoc.*, mae_series.numser FROM mae_series RIGHT JOIN var_plantilladoc " _
                   & " ON mae_series.id = var_plantilladoc.idser", xCon
    
    Set RstCab = Nothing
    RST_Mant RstCab, "SELECT * FROM var_plantillacab WHERE idplan = " & Rstdoc("id") & " ORDER BY var_plantillacab.numord ASC", xCon

    Set Rstdet = Nothing
    RST_Mant Rstdet, "SELECT * FROM var_plantilladet WHERE idplan = " & Rstdoc("id") & " ORDER BY var_plantilladet.numord ASC", xCon
    
    Fg1.Rows = 1
    
    For A = 1 To Rstdoc.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(A, 1) = Rstdoc("descripcion")
        Fg1.TextMatrix(A, 2) = Format(Rstdoc("numser"), "0000")
        Fg1.TextMatrix(A, 3) = Rstdoc("numitem")
        Fg1.TextMatrix(A, 4) = Rstdoc("filaitem")
        Fg1.TextMatrix(A, 5) = Rstdoc("altocab")
        Fg1.TextMatrix(A, 6) = NulosC(Rstdoc("tipoletra"))
        Fg1.TextMatrix(A, 7) = NulosC(Rstdoc("tamañoletra"))
        Fg1.TextMatrix(A, 8) = NulosC(Rstdoc("colorletra"))
        Fg1.TextMatrix(A, 9) = NulosC(Rstdoc("id"))
        Fg1.TextMatrix(A, 10) = NulosC(Rstdoc("tipdoc"))
        Fg1.TextMatrix(A, 11) = NulosC(Rstdoc("idser"))
        Rstdoc.MoveNext
    Next A

    Fg2.Rows = 1
    For A = 1 To RstCab.RecordCount
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(A, 1) = RstCab("item")
        Fg2.TextMatrix(A, 2) = RstCab("descripcion")
        Fg2.TextMatrix(A, 3) = RstCab("numord")
        Fg2.TextMatrix(A, 4) = RstCab("ancho")
        Fg2.TextMatrix(A, 5) = RstCab("posx")
        Fg2.TextMatrix(A, 6) = RstCab("posy")
        Fg2.TextMatrix(A, 7) = RstCab("campo")
        Fg2.TextMatrix(A, 8) = NulosC(RstCab("formato"))
        Fg2.TextMatrix(A, 9) = RstCab("tipo")
        RstCab.MoveNext
    Next A
    
    Fg3.Rows = 1
    For A = 1 To Rstdet.RecordCount
        Fg3.Rows = Fg3.Rows + 1
        Fg3.TextMatrix(A, 1) = Rstdet("item")
        Fg3.TextMatrix(A, 2) = Rstdet("descripcion")
        Fg3.TextMatrix(A, 3) = Rstdet("numord")
        Fg3.TextMatrix(A, 4) = Rstdet("ancho")
        Fg3.TextMatrix(A, 5) = Rstdet("posx")
        Fg3.TextMatrix(A, 6) = Rstdet("posy")
        Fg3.TextMatrix(A, 7) = Rstdet("campo")
        Fg3.TextMatrix(A, 8) = NulosC(Rstdet("formato"))
        Rstdet.MoveNext
    Next
    
    Mostrando = False
    
End Sub

Private Sub Form_Load()
    TabOne1.CurrTab = 0
    Fg1.ColWidth(4) = 0
    Fg1.ColWidth(5) = 0
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    
    Fg2.ColWidth(9) = 0
    Fg2.ColWidth(10) = 0
    
    Fg3.ColWidth(9) = 0
    
    Letras = "|#Super Draft 15cpi;Super Draft 15cpi" & _
             "|#Times New Roman;Times New Roman" & _
             "|#Arial;Arial" & _
             "|#Arial Black;Arial Black" & _
             "|#Verdama;Verdama"

    Tamaño = "|#9;9" & _
             "|#10;10" & _
             "|#11;11" & _
             "|#12;12" & _
             "|#13;13"
             
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then
        Dim Rpta As Integer
        Rpta = MsgBox("Seguro de eliminar la plantilla", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            xCon.Execute "DELETE FROM plantilladoc WHERE id = " & Val(Fg1.TextMatrix(Fg1.Row, 8)) & ""
            xCon.Execute "DELETE FROM plantillacab WHERE idplan = " & Val(Fg1.TextMatrix(Fg1.Row, 8)) & ""
            xCon.Execute "DELETE FROM plantilladet WHERE idplan = " & Val(Fg1.TextMatrix(Fg1.Row, 8)) & ""
        End If
        Form_Activate
    End If
    
    If Button.Index = 5 Then
        Cancelar
    End If
    
    If Button.Index = 6 Then
        If Grabar = True Then
            Cancelar
        End If
        'Call Form_Activate
    End If
    
    If Button.Index = 14 Then Unload Me
End Sub

Sub Modificar()
    Dim A As Integer
    Mostrando = True
    Toolbar
    Rstdoc.Filter = "id = " & Fg1.TextMatrix(Fg1.Row, 9) & ""
    Fg1.Rows = 1
    For A = 1 To Rstdoc.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(A, 1) = Rstdoc("descripcion")
        Fg1.TextMatrix(A, 2) = Format(Rstdoc("numser"), "0000")
        Fg1.TextMatrix(A, 3) = Rstdoc("numitem")
        Fg1.TextMatrix(A, 4) = Rstdoc("filaitem")
        Fg1.TextMatrix(A, 5) = Rstdoc("altocab")
        Fg1.TextMatrix(A, 6) = NulosC(Rstdoc("tipoletra"))
        Fg1.TextMatrix(A, 7) = NulosC(Rstdoc("tamañoletra"))
        Fg1.TextMatrix(A, 8) = NulosC(Rstdoc("colorletra"))
        Fg1.TextMatrix(A, 9) = NulosC(Rstdoc("id"))
        Fg1.TextMatrix(A, 10) = NulosC(Rstdoc("tipdoc"))
        Fg1.TextMatrix(A, 11) = NulosC(Rstdoc("idser"))
        Rstdoc.MoveNext
    Next A
    
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    Fg3.Editable = flexEDKbdMouse
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(2) = "|..."
    Fg1.ColComboList(6) = Letras
    Fg1.ColComboList(7) = Tamaño
    
End Sub

Sub Nuevo()
    QueHace = 1
    Toolbar
    Mostrando = True
    Fg1.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    Fg2.Rows = 1
    Fg3.Rows = 1
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    Fg3.Editable = flexEDKbdMouse
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(2) = "|..."
    Fg1.ColComboList(6) = Letras
    Fg1.ColComboList(7) = Tamaño
    'Fg1.SelectionMode = flexSelectionFree
End Sub

Sub Cancelar()
    Toolbar
    QueHace = 3
    Form_Activate
    'Blanquea
    'Bloquea
    Fg1.Editable = flexEDNone
    Fg2.Editable = flexEDNone
    Fg3.Editable = flexEDNone
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    'lblModo1.Caption = "Listado de Equipos"
End Sub


Function Grabar() As Boolean
    Grabar = False
    Dim F As Integer
    
    For F = 1 To Fg1.Rows - 1
        If NulosC(Fg1.TextMatrix(F, 1)) = "" Then
            MsgBox "No ha especificado el tipo de documento para la plantilla", vbInformation, xTitulo
            Exit Function
        End If
        If NulosC(Fg1.TextMatrix(F, 1)) = "" Then
            MsgBox "No ha especificado el tipo de documento para la plantilla", vbInformation, xTitulo
            Exit Function
        End If
    Next F
    
    If Fg2.Rows = 1 Then
        MsgBox "No se ha especificado campos para la cabezera de la plantilla"
        Exit Function
    End If
    
    For F = 1 To Fg2.Rows - 1
        If NulosC(Fg2.TextMatrix(F, 2)) = "" Then
            MsgBox "No ha especificado la descripcion del campo", vbInformation, xTitulo
            Exit Function
        End If
        If NulosC(Fg2.TextMatrix(F, 4)) = "" Then
            MsgBox "No ha especificado la posición X del campo", vbInformation, xTitulo
            Exit Function
        End If
        If NulosC(Fg2.TextMatrix(F, 5)) = "" Then
            MsgBox "No ha especificado la posición Y del campo", vbInformation, xTitulo
            Exit Function
        End If
        If NulosC(Fg2.TextMatrix(F, 6)) = "" Then
            MsgBox "No ha especificado el nombre del campo", vbInformation, xTitulo
            Exit Function
        End If
    Next F
    
    If Fg3.Rows = 1 Then
        MsgBox "No se ha especificado campos para el detalle de la plantilla"
        Exit Function
    End If
    For F = 1 To Fg3.Rows - 1
        If NulosC(Fg3.TextMatrix(F, 2)) = "" Then
            MsgBox "No ha especificado la descripcion del campo", vbInformation, xTitulo
            Exit Function
        End If
        If NulosC(Fg3.TextMatrix(F, 4)) = "" Then
            MsgBox "No ha especificado la posición X del campo", vbInformation, xTitulo
            Exit Function
        End If
        If NulosC(Fg3.TextMatrix(F, 5)) = "" Then
            MsgBox "No ha especificado la posición Y del campo", vbInformation, xTitulo
            Exit Function
        End If
        If NulosC(Fg3.TextMatrix(F, 6)) = "" Then
            MsgBox "No ha especificado el nombre del campo", vbInformation, xTitulo
            Exit Function
        End If
    Next F
    
    Dim rsDoc As New ADODB.Recordset
    Dim rsPlaCab As New ADODB.Recordset
    Dim rsPlaDet As New ADODB.Recordset
    
    Dim xId As Integer
    Dim A As Integer
    
    On Error GoTo LaCague
    xCon.BeginTrans
       
    If QueHace = 1 Then
        xId = HallaCodigoTabla("var_plantilladoc", xCon, "id")
        RST_Busq rsDoc, "SELECT * FROM var_plantilladoc", xCon
        RST_Busq rsPlaCab, "SELECT * FROM var_plantillacab", xCon
        RST_Busq rsPlaDet, "SELECT * FROM var_plantilladet", xCon
        rsDoc.AddNew
        rsDoc("id") = xId
    Else
        Rstdoc.MoveFirst
        RST_Busq rsDoc, "SELECT * FROM var_plantilladoc WHERE id = " & Rstdoc("id") & "", xCon
        xCon.Execute "DELETE * FROM var_plantillacab WHERE idplan = " & Rstdoc("id") & ""
        xCon.Execute "DELETE * FROM var_plantilladet WHERE idplan = " & Rstdoc("id") & ""
        RST_Busq rsPlaCab, "SELECT * FROM var_plantillacab", xCon
        RST_Busq rsPlaDet, "SELECT * FROM var_plantilladet", xCon
        xId = rsDoc("id")
    End If
       
    'GUARDA DOCUMENTO PLANTILLA
    For A = 1 To Fg1.Rows - 1
        rsDoc("tipdoc") = Val(Fg1.TextMatrix(A, 10))
        rsDoc("descripcion") = NulosC(Fg1.TextMatrix(A, 1))
        rsDoc("numitem") = NulosN(Fg1.TextMatrix(A, 3))
        rsDoc("filaitem") = NulosN(Fg1.TextMatrix(A, 4))
        rsDoc("altocab") = NulosN(Fg1.TextMatrix(A, 5))
        rsDoc("tipoletra") = NulosC(Fg1.TextMatrix(A, 6))
        rsDoc("tamañoletra") = NulosN(Fg1.TextMatrix(A, 7))
        rsDoc("colorletra") = NulosC(Fg1.TextMatrix(A, 8))
        rsDoc("idser") = NulosN(Fg1.TextMatrix(A, 11))
    Next A
    rsDoc.Update
    
    'GUARDA CABECERA PLANTILLA
    For A = 1 To Fg2.Rows - 1
        rsPlaCab.AddNew
        rsPlaCab("idplan") = xId
        rsPlaCab("item") = NulosN(Fg2.TextMatrix(A, 1))
        rsPlaCab("descripcion") = NulosC(Fg2.TextMatrix(A, 2))
        rsPlaCab("numord") = NulosC(Fg2.TextMatrix(A, 3))
        rsPlaCab("ancho") = NulosN(Fg2.TextMatrix(A, 4))
        rsPlaCab("posx") = NulosN(Fg2.TextMatrix(A, 5))
        rsPlaCab("posy") = NulosN(Fg2.TextMatrix(A, 6))
        rsPlaCab("campo") = NulosC(Fg2.TextMatrix(A, 7))
        rsPlaCab("formato") = NulosC(Fg2.TextMatrix(A, 8))
        rsPlaCab("tipo") = NulosN(Fg2.TextMatrix(A, 9))
        rsPlaCab.Update
    Next A
    
    'GUARDA DETALLE PLANTILLA
    For A = 1 To Fg3.Rows - 1
        rsPlaDet.AddNew
        rsPlaDet("idplan") = xId
        rsPlaDet("item") = Fg3.TextMatrix(A, 1)
        rsPlaDet("descripcion") = Fg3.TextMatrix(A, 2)
        rsPlaDet("numord") = Fg3.TextMatrix(A, 3)
        rsPlaDet("ancho") = Fg3.TextMatrix(A, 4)
        rsPlaDet("posx") = Fg3.TextMatrix(A, 5)
        rsPlaDet("posy") = Fg3.TextMatrix(A, 6)
        rsPlaDet("campo") = Fg3.TextMatrix(A, 7)
        rsPlaDet("formato") = Fg3.TextMatrix(A, 8)
        'rsPlaDet("tipo") = NulosC(Fg3.TextMatrix(A, 8))
        rsPlaDet.Update
    Next A
    
    MsgBox "El registro se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Mensaje"
    xCon.CommitTrans
    Grabar = True
    
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set rsDoc = Nothing
    Set rsPlaCab = Nothing
    Set rsPlaDet = Nothing

    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Sub Toolbar()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
End Sub







