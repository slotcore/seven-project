VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMantenimiento 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC=Salir"
      Height          =   435
      Left            =   8415
      TabIndex        =   11
      Top             =   345
      Width           =   945
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6705
      Left            =   30
      TabIndex        =   0
      Top             =   390
      Width           =   11490
      _cx             =   20267
      _cy             =   11827
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   8421376
      ForeColor       =   -2147483630
      FrontTabColor   =   12632064
      BackTabColor    =   8421376
      TabOutlineColor =   16777215
      FrontTabForeColor=   -2147483630
      Caption         =   "     Consulta     |    Detalles     "
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
      TabHeight       =   300
      TabCaptionPos   =   3
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6375
         Left            =   12105
         TabIndex        =   2
         Top             =   315
         Width           =   11460
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C000&
            Height          =   5895
            Left            =   105
            TabIndex        =   8
            Top             =   330
            Width           =   11235
            Begin VB.CommandButton CmdBusca 
               Height          =   240
               Index           =   0
               Left            =   2355
               Picture         =   "FrmMantenimiento.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   435
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               BackColor       =   &H00DEFDFE&
               Height          =   285
               Index           =   0
               Left            =   1740
               Locked          =   -1  'True
               TabIndex        =   9
               Text            =   "TxtTexto"
               Top             =   405
               Width           =   870
            End
            Begin VB.Label LblIdCampo 
               AutoSize        =   -1  'True
               Caption         =   "LblIdCampo"
               Height          =   195
               Index           =   0
               Left            =   8640
               TabIndex        =   14
               Top             =   330
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label LblDato 
               Appearance      =   0  'Flat
               BackColor       =   &H00808000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDato"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C0FFFF&
               Height          =   285
               Index           =   0
               Left            =   2640
               TabIndex        =   13
               Top             =   405
               Visible         =   0   'False
               Width           =   3795
            End
            Begin VB.Label LblCaption 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Label1"
               Height          =   195
               Index           =   0
               Left            =   255
               TabIndex        =   10
               Top             =   435
               Width           =   480
            End
         End
         Begin VB.Label LblTitulo2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Detalles"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   105
            TabIndex        =   7
            Top             =   75
            Width           =   10125
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6375
         Left            =   15
         TabIndex        =   1
         Top             =   315
         Width           =   11460
         Begin TrueOleDBGrid70.TDBGrid Db1 
            Height          =   5955
            Left            =   15
            TabIndex        =   4
            Top             =   405
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   10504
            _LayoutType     =   4
            _RowHeight      =   15
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Descripcion"
            Columns(0).DataField=   "baselegal"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   1
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=1"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1349"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1270"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
            PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632064
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=79,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDEFDFE&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=37,.fgcolor=&H8000000D&,.bold=-1"
            _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.bgcolor=&H80&,.fgcolor=&HFFFF&"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HC0FFC0&"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=36,.bgcolor=&H80&,.fgcolor=&HFFFF&"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=36,.bgcolor=&H80&"
            _StyleDefs(30)  =   ":id=18,.fgcolor=&HFFFF&,.locked=-1"
            _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(32)  =   ":id=17,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(33)  =   ":id=17,.fontname=MS Sans Serif"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(43)  =   "Named:id=33:Normal"
            _StyleDefs(44)  =   ":id=33,.parent=0"
            _StyleDefs(45)  =   "Named:id=34:Heading"
            _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(47)  =   ":id=34,.wraptext=-1"
            _StyleDefs(48)  =   "Named:id=35:Footing"
            _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   "Named:id=36:Selected"
            _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(52)  =   "Named:id=37:Caption"
            _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(54)  =   "Named:id=38:HighlightRow"
            _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(56)  =   "Named:id=39:EvenRow"
            _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(58)  =   "Named:id=40:OddRow"
            _StyleDefs(59)  =   ":id=40,.parent=33"
            _StyleDefs(60)  =   "Named:id=41:RecordSelector"
            _StyleDefs(61)  =   ":id=41,.parent=34"
            _StyleDefs(62)  =   "Named:id=42:FilterBar"
            _StyleDefs(63)  =   ":id=42,.parent=33"
         End
         Begin VB.Label LblTitulo1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   105
            TabIndex        =   5
            Top             =   75
            Width           =   10125
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   105
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
            Picture         =   "FrmMantenimiento.frx":0132
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantenimiento.frx":0676
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantenimiento.frx":07FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantenimiento.frx":0C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantenimiento.frx":0D66
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantenimiento.frx":12AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantenimiento.frx":17EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantenimiento.frx":1902
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantenimiento.frx":1A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantenimiento.frx":1E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMantenimiento.frx":1FD6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
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
         NumButtons      =   12
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
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10125
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000BBFDF&
      BackStyle       =   1  'Opaque
      Height          =   225
      Left            =   12180
      Top             =   2490
      Width           =   1245
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H002102B3&
      BackStyle       =   1  'Opaque
      Height          =   225
      Left            =   12120
      Top             =   1950
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0021A0FE&
      BackStyle       =   1  'Opaque
      Height          =   225
      Left            =   12150
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00235FFC&
      BackStyle       =   1  'Opaque
      Height          =   225
      Left            =   12090
      Top             =   1230
      Width           =   1245
   End
End
Attribute VB_Name = "FrmMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim RstCab As New ADODB.Recordset

Dim CaracteresNumericos As String
Dim QueHace As Integer
Dim xCampCla As String   'VARIABLE PARA SABE EL NOMBRE DEL CAMPO CLAVE DE LA TABLA

Private Sub CmdBusca_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    
    Dim X, A, B, C As Integer
    Dim Campos(2, 4) As String
    
    Dim xForm As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    For X = LBound(xVincula) To UBound(xVincula)
        If xCampos(Index, 1) = xVincula(X, 6) Then
            
            Dim xCad As String
            Dim xIndice As Integer
            
            '-------------------------------------
            'hallamos los rotulos para la busqueda
            xIndice = 0
            For C = 1 To Len(Trim(xVincula(X, 3)))
                If Mid(xVincula(X, 3), C, 1) <> "," Then
                    xCad = xCad + Mid(xVincula(X, 3), C, 1)
                Else
                    Campos(xIndice, 0) = xCad  'ROTULO PARA LA BUSQUEDA
                    xIndice = xIndice + 1:     xCad = ""
                End If
            Next C
            Campos(xIndice, 0) = xCad  'ROTULO PARA LA BUSQUEDA
            
            '---------------------------------------------
            'hallamos el nombre del campo para la busqueda
            xCad = ""
            xIndice = 0
            For C = 1 To Len(Trim(xVincula(X, 2)))
                If Mid(xVincula(X, 2), C, 1) <> "," Then
                    xCad = xCad + Mid(xVincula(X, 2), C, 1)
                Else
                    Campos(xIndice, 1) = xCad  'NOMBRE DEL CAMPO PARA LA BUSQUEDA
                    xIndice = xIndice + 1:     xCad = ""
                End If
            Next C
            Campos(xIndice, 1) = xCad  'NOMBRE DEL CAMPO PARA LA BUSQUEDA
            
            '---------------------------------------------
            'hallamos el tamaño del campo
            xCad = ""
            xIndice = 0
            For C = 1 To Len(Trim(xVincula(X, 4)))
                If Mid(xVincula(X, 4), C, 1) <> "," Then
                    xCad = xCad + Mid(xVincula(X, 4), C, 1)
                Else
                    Campos(xIndice, 2) = Val(xCad)  'TAMAÑO DEL CAMPO
                    xIndice = xIndice + 1:     xCad = ""
                End If
            Next C
            Campos(xIndice, 2) = Val(xCad)  'TAMAÑO DEL CAMPO
            
            '---------------------------------------------
            'hallamos el tipo del campo
            xCad = ""
            xIndice = 0
            For C = 1 To Len(Trim(xVincula(X, 5)))
                If Mid(xVincula(X, 5), C, 1) <> "," Then
                    xCad = xCad + Mid(xVincula(X, 5), C, 1)
                Else
                    Campos(xIndice, 3) = xCad  'TIPO DEL CAMPO
                    xIndice = xIndice + 1:     xCad = ""
                End If
            Next C
            Campos(xIndice, 3) = xCad  'TIPO DEL CAMPO
            
            xForm.SQLCad = "SELECT * FROM  " & xVincula(X, 0) & ""
            
            xForm.Titulo = "BUSCANDO"
            xForm.FormaBusca = Principio
            xForm.Criterio = ""
            xForm.Ordenado = Campos(0, 1)
            xForm.CampoBusca = Campos(0, 1)
            Set xForm.Coneccion = xConeccion
            Set xRs = xForm.BuscarReg(Campos)
            If xRs.State = 1 Then
                LblIdCampo(Index).Caption = xRs(xVincula(X, 1))
                TxtTexto(Index).Text = NulosC(xRs(xVincula(X, 9)))
                LblDato(Index).Caption = Trim(xRs(xVincula(X, 7)))
            End If
            Set xForm = Nothing
            Set xRs = Nothing
        End If
    Next X
End Sub

Private Sub CmdEsc_Click()
    Set RstCab = Nothing
    Unload Me
End Sub

Private Sub Db1_DblClick()
    TabOne1.CurrTab = 1
    MuestraDatos
    TxtTexto(0).SetFocus
End Sub

Private Sub Db1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TabOne1.CurrTab = 1
        MuestraDatos
        TxtTexto(0).SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        xCampCla = BuscaCampoLista(xCampoClave, 0, 1, xCampos)

        CaracteresNumericos = "0123456789.-" & Chr(8)
        SeEjecuto = True
        RST_Busq RstCab, xCadSQL, xConeccion
        PreparaGrid
        AgregandoControles
       
        Set Db1.DataSource = RstCab
        Db1.SetFocus
    End If
End Sub

Sub PreparaGrid()
    Dim A As Integer
    
    Dim C As TrueOleDBGrid70.Column
    
    For A = LBound(xCamposVista) To UBound(xCamposVista)
        If A >= 1 Then Set C = Db1.Columns.Add(A)
        Db1.Columns(A).Visible = True
        Db1.Columns(A).Caption = xCamposVista(A, 0)
        Db1.Columns(A).DataField = xCamposVista(A, 1)
        Db1.Columns(A).Width = Val(xCamposVista(A, 2))
        If Trim(xCamposVista(A, 4)) = "I" Then Db1.Columns(A).Alignment = dbgLeft
        If Trim(xCamposVista(A, 4)) = "D" Then Db1.Columns(A).Alignment = dbgRight
        If Trim(xCamposVista(A, 4)) = "C" Then Db1.Columns(A).Alignment = dbgCenter
        
        If A = UBound(xCamposVista) - 1 Then Exit For
    Next A
End Sub

Private Sub Form_Load()
    Me.Caption = xTituloForm
    SeEjecuto = False
    CmdEsc.Left = 15000
    TabOne1.CurrTab = 0
    QueHace = 3
    
    If xPermiteActualiza = 0 Then
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = False
        Toolbar1.Buttons(3).Visible = False
        Toolbar1.Buttons(4).Visible = False
        Toolbar1.Buttons(5).Visible = False
        Toolbar1.Buttons(6).Visible = False
        Toolbar1.Buttons(7).Visible = False
    End If
End Sub

Sub ActivaToolbar()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
End Sub

Sub Buscar()
    Dim xForm As New EPS_Buscar.Buscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim A As Integer
    Dim B As Integer
    Dim Campos(2, 4) As String
        
    TabOne1.CurrTab = 0
    For A = LBound(xCamposBusqueda) To UBound(xCamposBusqueda)
        Campos(A, 0) = BuscaCampoLista(Trim(xCamposBusqueda(A)), 1, 0, xCampos) 'ROTULO PARA LA BUSQUEDA
        Campos(A, 1) = BuscaCampoLista(Trim(xCamposBusqueda(A)), 1, 1, xCampos) 'NOMBRE DEL CAMPO PARA LA BUSQUEDA
        Campos(A, 2) = BuscaCampoLista(Trim(xCamposBusqueda(A)), 1, 3, xCampos) 'TAMAÑO DEL CAMPO
        Campos(A, 3) = BuscaCampoLista(Trim(xCamposBusqueda(A)), 1, 2, xCampos) 'TIPO DEL CAMPO
        
        If A >= (UBound(xCamposBusqueda) - 1) Then Exit For
    Next A
    
    'Campos(0, 0) = "Documento":     Campos(0, 1) = "descripcion":     Campos(0, 2) = "7000":    Campos(0, 3) = "C"
    'Campos(1, 0) = "Codigo":        Campos(1, 1) = "id":              Campos(1, 2) = "1200":    Campos(1, 3) = "C"
    xForm.SQLCad = "SELECT * FROM  " & xTabla & ""
    
    xForm.Titulo = "BUSCANDO"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = BuscaCampoLista(Trim(xCamposBusqueda(0)), 1, 1, xCampos)
    xForm.CampoBusca = BuscaCampoLista(Trim(xCamposBusqueda(0)), 1, 1, xCampos)
    Set xForm.Coneccion = xConeccion
    Set xRs = xForm.BuscarReg(Campos)
    If xRs.State = 1 Then
        RstCab.MoveFirst
        If BuscaCampoLista(Trim(xCamposBusqueda(A)), 1, 2, xCampos) = "C" Then
            RstCab.Find "" & BuscaCampoLista(Trim(xCamposBusqueda(A)), 1, 1, xCampos) & " = '" & xRs(BuscaCampoLista(Trim(xCamposBusqueda(A)), 1, 1, xCampos)) & "'"
        Else
            RstCab.Find "" & BuscaCampoLista(Trim(xCamposBusqueda(A)), 1, 1, xCampos) & " = " & xRs(BuscaCampoLista(Trim(xCamposBusqueda(A)), 1, 1, xCampos)) & ""
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o modificando un registro", vbInformation + vbOKOnly + vbDefaultButton1, "Mantenimiento de Tablas"
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        MuestraDatos
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then
        TabOne1.CurrTab = 0
        Dim Rpta As Integer
        Rpta = MsgBox("Esta seguro de eliminar el registro seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, "Mantenimiento de Tablas")
        If Rpta = vbYes Then
            If BuscaCampoLista(xCampoClave, 0, 2, xCampos) = "C" Then
                xConeccion.Execute "DELETE * FROM " & xTabla & " WHERE " & xCampCla & " ='" & RstCab(xCampCla) & "'"
            Else
                xConeccion.Execute "DELETE * FROM " & xTabla & " WHERE " & xCampCla & " =" & RstCab(xCampCla) & ""
            End If
            RstCab.Requery
            Db1.Refresh
        End If
    End If
    
    If Button.Index = 5 Then
        ActivaToolbar
        RecorreControles 2
        LblTitulo2.Caption = "Detalles"
        TabOne1.TabEnabled(0) = True
        QueHace = 3
        TabOne1.CurrTab = 0
    End If
    
    If Button.Index = 6 Then
        If Grabar = True Then
            ActivaToolbar
            RecorreControles 2
            LblTitulo2.Caption = "Detalles"
            TabOne1.TabEnabled(0) = True
            TabOne1.CurrTab = 0
            QueHace = 3
            RstCab.Requery
            Db1.Refresh
        End If
    End If
    
    If Button.Index = 9 Then
        RstCab.Filter = adFilterNone
        RstCab.Requery
        Db1.Refresh
    End If
    If Button.Index = 10 Then Buscar
    If Button.Index = 12 Then
        If QueHace <> 3 Then
            MsgBox "No puede salir del formulario miestras este agregando o modificando un registro", vbInformation + vbOKOnly + vbDefaultButton1, "Mantenimiento de Tablas"
            Exit Sub
        End If
        Set RstCab = Nothing
        Unload Me
    End If
End Sub

Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaToolbar
    LblTitulo2.Caption = "Agregando Datos"
    RecorreControles 1
    RecorreControles 2
    
    If xDefCampoClave = True Then
        'SE PERMITE EL INGRESO DEL CAMPO CLAVE
        
        TxtTexto.Item(0).Locked = Not TxtTexto.Item(0).Locked
        If (BuscaCampoLista(xCampoClave, 0, 2, xCampos)) <> "C" Then
            'MUESTRA EL COSIGO AUTOGENERADO SUGERIDO POR EL SISTEMA
            TxtTexto.Item(0).Text = HallaCodigoTabla(xTabla, xConeccion, xCampCla)
        Else
            TxtTexto.Item(0).Text = ""
        End If
    Else
        'NO SE PERMITE EL INGRESO DEL CAMPO CLAVE
        TxtTexto.Item(0).Text = HallaCodigoTabla(xTabla, xConeccion, xCampCla)
    End If
    TxtTexto.Item(0).SetFocus
End Sub

Sub Modificar()
    QueHace = 2
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaToolbar
    LblTitulo2.Caption = "Modificando Datos"
    RecorreControles 1
    MuestraDatos
    RecorreControles 2
    TxtTexto.Item(1).SetFocus
End Sub

Sub RecorreControles(Quehacer As Integer)
    '1 = LIMPIA LOS CONTROLES
    '2 = BLOQUEA LOS CONTROLES
    Dim A As Integer
    For A = LBound(xCampos) To UBound(xCampos)
        If A >= 1 Then
            'LIMPIAMOS LOS CONTROLES
            If Quehacer = 1 Then TxtTexto.Item(A).Text = "": LblDato.Item(A).Caption = "": LblIdCampo.Item(A).Caption = ""
            'BLOQUEAMOS LOS CONTROLES
            If Quehacer = 2 Then TxtTexto.Item(A).Locked = Not TxtTexto.Item(A).Locked
        End If
        
        If A = (UBound(xCampos) - 1) Then
            Exit For
        End If
    Next A
    'SE VUELVE A BLOQUEAR PARA QUE SOLO SE ACTIVE CUANDO SE PERMITA EL INGRESO DEL CAMPO CLAVE
    If Quehacer = 2 Then
        If TxtTexto.Item(0).Locked = False Then
            TxtTexto.Item(0).Locked = Not TxtTexto.Item(0).Locked
        End If
    End If
End Sub

Sub AgregandoControles()
    Dim I, A, X As Integer
    LblDato(0).Caption = ""
    
    For A = LBound(xCampos) To UBound(xCampos)
        If A >= 1 Then
            Load CmdBusca(A)
            Load TxtTexto(A)
            Load LblCaption(A)
            Load LblDato(A)
            Load LblIdCampo(A)
        End If
        
        TxtTexto(A).Width = xCampos(A, 3)
        LblCaption(A).Caption = xCampos(A, 0)
        
        If A >= 1 Then
            TxtTexto(A).Visible = True
            LblCaption(A).Visible = True
            TxtTexto(A).Top = TxtTexto(A - 1).Top + 300
            LblCaption(A).Top = LblCaption(A - 1).Top + 300
        End If
        
        For X = LBound(xVincula) To UBound(xVincula)
            If xVincula(X, 6) = xCampos(A, 1) Then
                CmdBusca(A).Visible = True
                CmdBusca(A).Left = (TxtTexto(A).Left + TxtTexto(A).Width) - 250
                CmdBusca(A).Top = TxtTexto(A).Top + 25
                
                LblDato(A).Visible = True
                LblDato(A).Top = TxtTexto(A).Top
                LblDato(A).Left = (TxtTexto(A).Left + TxtTexto(A).Width) + 20
                Exit For
            End If
        Next X
    
        If A = (UBound(xCampos) - 1) Then
            Exit For
        End If
    Next A
End Sub

Sub MuestraDatos()
    Dim A As Integer
    Dim X As Integer
    
On Error GoTo LaCague

    For A = LBound(xCampos) To UBound(xCampos)
        If xCampos(A, 2) = "C" Then
            TxtTexto(A).Text = NulosC(RstCab(xCampos(A, 1)))
        Else
            TxtTexto(A).Text = NulosN(RstCab(xCampos(A, 1)))
        End If
        
        For X = LBound(xVincula) To UBound(xVincula)
            If xVincula(X, 6) = xCampos(A, 1) Then
                LblIdCampo(A).Caption = TxtTexto(A).Text
                LblDato(A).Caption = Busca_Codigo(TxtTexto(A).Text, xVincula(X, 1), xVincula(X, 7), xVincula(X, 0), xVincula(X, 8), xConeccion)
                TxtTexto(A).Text = Busca_Codigo(TxtTexto(A).Text, xVincula(X, 1), xVincula(X, 9), xVincula(X, 0), xVincula(X, 8), xConeccion)
            End If
        Next X
            
        If A >= (UBound(xCampos) - 1) Then Exit For
    Next A

    Exit Sub
LaCague:
    If Err.Number = 3265 Then
        MsgBox "No se ha encontrado el siguiente campo: " + Trim(xCampos(A, 1))
    Else
        MsgBox Trim(Err.Description)
    End If
    Resume Next
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    
    If UCase(xCampos(Index, 2)) = "N" Then
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Function Grabar() As Boolean
    If xDefCampoClave = True Then
        If TxtTexto.Item(0).Text = "" Then
            MsgBox "No ha especificado el campo codigo del registro", vbInformation + vbOKOnly + vbDefaultButton1, "Sistema de Mantenimiento"
            TxtTexto.Item(0).SetFocus
            Exit Function
        End If
    End If

    Dim RstGra As New ADODB.Recordset
    
    On Error GoTo LaCague
    
    xConeccion.BeginTrans
    
    If QueHace = 1 Then
        RST_Mant RstGra, "SELECT * FROM  " & xTabla & "", xConeccion
        RstGra.AddNew
    Else
        If (BuscaCampoLista(xCampoClave, 0, 2, xCampos)) = "C" Then
            RST_Mant RstGra, "SELECT * FROM  " & xTabla & " WHERE " & xCampCla & " = '" & TxtTexto.Item(0).Text & "'", xConeccion
        Else
            RST_Mant RstGra, "SELECT * FROM  " & xTabla & " WHERE " & xCampCla & " = " & TxtTexto.Item(0).Text & "", xConeccion
        End If
    End If
    Dim A As Integer
    
    For A = LBound(xCampos) To UBound(xCampos)
        If xCampos(A, 2) = "C" Then
            If CmdBusca(A).Visible = False Then
                RstGra(xCampos(A, 1)) = NulosC(TxtTexto(A).Text)
            Else
                RstGra(xCampos(A, 1)) = NulosC(LblIdCampo(A).Caption)
            End If
        Else
            If CmdBusca(A).Visible = False Then
                RstGra(xCampos(A, 1)) = NulosN(TxtTexto(A).Text)
            Else
                RstGra(xCampos(A, 1)) = NulosN(LblIdCampo(A).Caption)
            End If
        End If
        
        If A >= (UBound(xCampos) - 1) Then Exit For
    Next A

    RstGra.Update
    xConeccion.CommitTrans
    
    MsgBox "El registro se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Sistema de Mantenimiento"
    Grabar = True
    Exit Function

LaCague:
'    Resume
    xConeccion.RollbackTrans
    Set RstGra = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function
