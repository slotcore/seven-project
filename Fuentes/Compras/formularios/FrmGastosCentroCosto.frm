VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmGastosCentroCosto 
   Caption         =   "Contabilidad - Gastos x Centro de Costos"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   3900
      Left            =   15
      TabIndex        =   10
      Top             =   1245
      Width           =   11880
      _cx             =   20955
      _cy             =   6879
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmGastosCentroCosto.frx":0000
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
      Height          =   1245
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   11865
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   5205
         TabIndex        =   12
         Top             =   450
         Width           =   2160
      End
      Begin VB.CommandButton CmdMuestra 
         Height          =   570
         Left            =   10230
         Picture         =   "FrmGastosCentroCosto.frx":00C0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   435
         Width           =   825
      End
      Begin VB.CommandButton CmdBusOrdCom 
         Height          =   240
         Left            =   1830
         Picture         =   "FrmGastosCentroCosto.frx":0502
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   885
         Width           =   240
      End
      Begin VB.TextBox TxtIdMon 
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Text            =   "TxtIdMon"
         Top             =   855
         Width           =   1020
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   525
         Width           =   1290
         _ExtentX        =   2275
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
         Left            =   3435
         TabIndex        =   2
         Top             =   525
         Width           =   1290
         _ExtentX        =   2275
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Niveles del Centro Costo"
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
         Left            =   5220
         TabIndex        =   13
         Top             =   225
         Width           =   2115
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   5100
         X2              =   5100
         Y1              =   180
         Y2              =   1155
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   5085
         X2              =   5085
         Y1              =   180
         Y2              =   1155
      End
      Begin VB.Label LblMoneda 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblMoneda"
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
         Left            =   2145
         TabIndex        =   9
         Top             =   855
         Width           =   2580
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Inicio"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   885
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo de Costeo"
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
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   195
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "FchFinal"
         Height          =   180
         Left            =   2640
         TabIndex        =   4
         Top             =   555
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Inicio"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   555
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmGastosCentroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FrmGastosCentroCosto
' DateTime  : 03/01/2007 9:00
' Author    : Enrique Pollongo Sierra
' Purpose   : gfgfgfgfgfgfg
'---------------------------------------------------------------------------------------
Option Explicit

Private Sub CmdMuestra_Click()
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT con_centrocosto.id, con_centrocosto.codigo, con_centrocosto.descripcion, (SELECT Sum([impcos]) AS total " _
        & " FROM com_compras RIGHT JOIN com_comprascosto ON com_compras.id = com_comprascosto.idcom " _
        & " WHERE (((com_comprascosto.idcencos)=con_centrocosto.id) AND ((com_compras.fchdoc)>=CDate('" & TxtFchIni.Valor & "') " _
        & " And (com_compras.fchdoc)<=CDate('" & TxtFchFin.Valor & "')))) AS total From con_centrocosto ORDER BY con_centrocosto.codigo", xCon
    
    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("codigo")
            Fg1.TextMatrix(A, 2) = Rst("descripcion")
            Fg1.TextMatrix(A, 3) = Format(Rst("total"), "0.00")
            Fg1.TextMatrix(A, 5) = Rst("id")
            
            If Len(Trim(Rst("codigo"))) = 2 Then
                RST_Busq Rst2, "SELECT Sum([impcos]) AS total FROM com_compras RIGHT JOIN (con_centrocosto RIGHT JOIN " _
                    & " com_comprascosto ON con_centrocosto.id = com_comprascosto.idcencos) ON com_compras.id = com_comprascosto.idcom " _
                    & " WHERE (((con_centrocosto.codigo) Like '" & Rst("codigo") & "%') AND ((com_compras.fchdoc)>=CDate('" & TxtFchIni.Valor & "') " _
                    & " And (com_compras.fchdoc)<=CDate('" & TxtFchFin.Valor & "')))", xCon
                
                Fg1.TextMatrix(A, 4) = Format(Rst2("total"), "0.00")
            End If
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
End Sub

Private Sub Form_Load()
    Fg1.ColWidth(5) = 0
    Fg1.Rows = 1
    List1.AddItem ("Nivel 1   01")
    List1.AddItem ("Nivel 2   0101")
    List1.AddItem ("Nivel 3   010101")
    List1.AddItem ("Nivel 4   01010101")
    List1.AddItem ("Nivel 5   0101010101")
    List1.AddItem ("Nivel 6   010101010101")
    List1.Selected(1) = True
End Sub



'SELECT con_centrocosto.id, con_centrocosto.codigo, con_centrocosto.descripcion, (SELECT Sum([impcos]) AS total
'From com_comprascosto
'WHERE (((com_comprascosto.idcencos)=con_centrocosto.id))) AS total
'From con_centrocosto
'ORDER BY con_centrocosto.codigo;

