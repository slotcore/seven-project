VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtipocambio 
   Caption         =   "Registro de tipo de cambio"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   Icon            =   "frmtipocambio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBGrid70.TDBGrid dg1 
      Height          =   3615
      Left            =   240
      TabIndex        =   15
      Top             =   1125
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6376
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
      MultiSelect     =   2
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.alignment=2,.valignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(48)  =   ":id=34,.fgcolor=&H80000012&,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   550
      Left            =   7695
      Picture         =   "frmtipocambio.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame fracons 
      Height          =   1005
      Left            =   5265
      TabIndex        =   9
      Top             =   60
      Width           =   3660
      Begin VB.CommandButton cmdbuscar 
         Height          =   675
         Left            =   2370
         Picture         =   "frmtipocambio.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpfecfin 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   38692
      End
      Begin MSComCtl2.DTPicker dtpfecini 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   38692
      End
      Begin VB.Label lblfecfin 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   645
         Width           =   705
      End
      Begin VB.Label lblfecini 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame fradatos 
      Height          =   645
      Left            =   255
      TabIndex        =   2
      Top             =   420
      Width           =   4935
      Begin VB.TextBox txttc_venta 
         Height          =   315
         Left            =   3990
         MaxLength       =   5
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txttc_compra 
         Height          =   315
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpfecha 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   675
         TabIndex        =   4
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16449539
         CurrentDate     =   38692
      End
      Begin VB.Label lblfecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lbltc_venta 
         AutoSize        =   -1  'True
         Caption         =   "Venta"
         Height          =   195
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lbltc_compra 
         AutoSize        =   -1  'True
         Caption         =   "Compra"
         Height          =   195
         Left            =   2010
         TabIndex        =   5
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.ComboBox cbomonedas 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   45
      Width           =   1335
   End
   Begin VB.Label lblmoneda 
      AutoSize        =   -1  'True
      Caption         =   "Moneda"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   90
      Width           =   585
   End
End
Attribute VB_Name = "frmtipocambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function GeneraCodigo() As Integer
Dim rstbus As New ADODB.Recordset

RST_Busq rstbus, "Select id from con_tc order by id", xCon

If rstbus.RecordCount = 0 Then
     GeneraCodigo = 1
Else
    rstbus.MoveLast
    GeneraCodigo = rstbus.Fields(0) + 1
    rstbus.Close
End If
Set rstbus = Nothing
End Function

Private Sub cbomonedas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then dtpfecha.SetFocus

End Sub

Private Sub cmdbuscar_Click()

Dim RS As New ADODB.Recordset
 RST_Busq RS, "SELECT MAE_Moneda.descripcion as [Moneda], con_tc.fecha , con_tc.impcom as [Compra], con_tc.impven as [Venta]" & _
 " FROM MAE_Moneda INNER JOIN con_tc ON MAE_Moneda.id = con_tc.idmoneda " & _
 " WHERE con_tc.Fecha >= cdate('" & dtpfecini.Value & "') and con_tc.Fecha <= cdate('" & dtpfecfin.Value & "') ORDER BY con_tc.fecha", xCon

Set Dg1.DataSource = RS

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub dtpfecfin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Dg1.SetFocus
End Sub

Private Sub dtpfecfin_LostFocus()
Call cmdbuscar_Click
End Sub

Private Sub dtpfecha_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RS As New ADODB.Recordset
If KeyCode = 13 Then
txttc_compra.Text = 0
txttc_venta.Text = 0

    RST_Busq RS, "Select fecha, impCom, impVen from con_tc Where con_tc.Fecha = #" & Format(dtpfecha, "mm/dd/yyyy") & "#", xCon
 
    If RS.RecordCount > 0 Then
        txttc_compra = RS!impcom
        txttc_venta = RS!impven
    End If

    RS.Close
    
    txttc_compra.SetFocus

End If
Set RS = Nothing
End Sub

Private Sub dtpfecini_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then dtpfecfin.SetFocus
End Sub

Private Sub Form_Load()
Dim RS As New ADODB.Recordset
RST_Busq RS, "Select * from MAE_Moneda", xCon

Do While Not RS.EOF
 cbomonedas.AddItem RS!descripcion & Space(100) & RS!id
 RS.MoveNext
Loop
If RS.RecordCount > 0 Then cbomonedas.ListIndex = 0


RST_Busq RS, "SELECT MAE_Moneda.Descripcion, con_tc.Fecha, con_tc.ImpCom as [Compra], con_tc.Impven as [Venta]" & _
 " FROM MAE_Moneda INNER JOIN con_tc ON MAE_Moneda.id = con_tc.id " & _
 " ORDER BY con_tc.fecha ", xCon
Set Dg1.DataSource = RS

        


 

End Sub

Private Sub txttc_compra_KeyPress(KeyAscii As Integer)



If KeyAscii = 13 Then txttc_venta.SetFocus



End Sub

Private Sub txttc_venta_KeyPress(KeyAscii As Integer)
Dim RS As New ADODB.Recordset

If KeyAscii = 13 Then

If Val(txttc_compra) <= 0 Then
MsgBox "Tipo de cambio de compra debe ser mayor a 0", vbExclamation, Me.Caption
txttc_compra.SetFocus
Exit Sub
End If

If Val(txttc_venta) <= 0 Then
MsgBox "Tipo de cambio de venta debe ser mayor a 0", vbExclamation, Me.Caption
txttc_venta.SetFocus
Exit Sub
End If


'Registramos el tipo de cambio
    
    RST_Busq RS, "Select * from con_tc where idmoneda = " & Val(Right(Me.cbomonedas, 3)) & "  and  fecha =#" & Format(dtpfecha, "mm/dd/yyyy") & "#", xCon
    If RS.RecordCount = 0 Then
        RS.AddNew
        RS!id = GeneraCodigo
    End If
    
    RS!idmoneda = Val(Right(Me.cbomonedas, 3))
    RS!fecha = dtpfecha
    RS!impcom = Val(txttc_compra)
    RS!impven = Val(txttc_venta)
    RS.Update
    
        
    RST_Busq RS, "SELECT MAE_Monedas.Descripcion as [Moneda], con_tc.Fecha, con_tc.ImpCom as [Compra] , con_tc.Impven as [Venta]" & _
     " FROM MAE_Monedas INNER JOIN con_tc ON MAE_Monedas.id = con_tc.idmoneda " & _
     " ORDER BY con_tc.fecha ", xCon

    Set Dg1.DataSource = RS
    dtpfecha.SetFocus

End If




End Sub
