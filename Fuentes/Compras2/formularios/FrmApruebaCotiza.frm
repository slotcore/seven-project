VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.ocx"
Begin VB.Form FrmApruebaCotiza 
   Caption         =   "Compras - Aprobar Orden de Cotizacion"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PushButton5 
      Height          =   630
      Left            =   435
      TabIndex        =   0
      Top             =   5430
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   1111
      _StockProps     =   79
      Caption         =   "Ver Cotizacion"
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmApruebaCotiza.frx":0000
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   465
      Left            =   15
      TabIndex        =   1
      Top             =   0
      Width           =   9780
      _Version        =   786432
      _ExtentX        =   17251
      _ExtentY        =   820
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.Label Label1 
         Height          =   300
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   9615
         _Version        =   786432
         _ExtentX        =   16960
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Orden de Cotizacion Emitidas"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   630
      Left            =   2160
      TabIndex        =   3
      Top             =   5430
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   1111
      _StockProps     =   79
      Caption         =   "Aprobar Cotizacion"
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmApruebaCotiza.frx":031A
   End
   Begin TrueOleDBGrid70.TDBGrid Dg1 
      Height          =   4830
      Left            =   0
      TabIndex        =   4
      Top             =   495
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   8520
      _LayoutType     =   4
      _RowHeight      =   31
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nº Cotizacion"
      Columns(0).DataField=   "numdoc"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Fch. Emi."
      Columns(1).DataField=   "fchemi"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Area"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Solicitante"
      Columns(3).DataField=   "nombre"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nº Items"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   528
      Columns(5)._MaxComboItems=   5
      Columns(5).ValueItems(0)._DefaultItem=   0
      Columns(5).ValueItems(0).Value=   "Aprobada"
      Columns(5).ValueItems(0).Value.vt=   8
      Columns(5).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(0).DisplayValue(0)=   "bHQAANYLAABCTdYLAAAAAAAANgAAACgAAAAgAAAAHwAAAAEAGAAAAAAAoAsAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(0).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(0).DisplayValue(3)=   "///////////////////////////AwMCAgICAgICAgICAgICAgICAgICAgICAgIDAwMDAwMDAwMD/"
      Columns(5).ValueItems(0).DisplayValue(4)=   "///////////////////////////////////////////////////////////////////////AwMCA"
      Columns(5).ValueItems(0).DisplayValue(5)=   "gIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAgICAgICAgIDAwMDAwMDAwMD/////////////////"
      Columns(5).ValueItems(0).DisplayValue(6)=   "//////////////////////////////////////////+AgIAAgIAAgIAAgIAAgIAAgIAAgIAAgIAA"
      Columns(5).ValueItems(0).DisplayValue(7)=   "gIAAgIAAgIAAAAAAAACAgICAgICAgIDAwMDAwMD/////////////////////////////////////"
      Columns(5).ValueItems(0).DisplayValue(8)=   "//////////////8AgIAAgIAA//8A//8A//8A//8A//8A//8A//8A//8A//8AgIAAgIAAgIAAAAAA"
      Columns(5).ValueItems(0).DisplayValue(9)=   "AACAgICAgIDAwMDAwMD///////////////////////////////////////////8AgIAA//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(10)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AgIAAgIAAAACAgICAgIDAwMDAwMD/"
      Columns(5).ValueItems(0).DisplayValue(11)=   "//////////////////////////////////8AgIAA//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(12)=   "//8A//8A//8A//8A//8A//8A//8AgIAAgIAAAACAgICAgIDAwMDAwMD/////////////////////"
      Columns(5).ValueItems(0).DisplayValue(13)=   "//////8AgIAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(14)=   "//8A//8AgIAAgIAAAACAgICAgIDAwMD///////////////////////8AgIAA//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(15)=   "//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8AgIAAgIAAAACA"
      Columns(5).ValueItems(0).DisplayValue(16)=   "gIDAwMDAwMD///////////////////8AgIAA//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8A"
      Columns(5).ValueItems(0).DisplayValue(17)=   "AP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8AgIAAAACAgICAgIDAwMD/////////////"
      Columns(5).ValueItems(0).DisplayValue(18)=   "//8AgIAA//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(0).DisplayValue(19)=   "//8A//8A//8A//8A//8A//8AgIAAAACAgIDAwMDAwMD///////////8AgIAA//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(20)=   "//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(21)=   "gIAAAACAgICAgIDAwMD///////8AgIAA//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8A//8A"
      Columns(5).ValueItems(0).DisplayValue(22)=   "//8A//8A//8A//8A//8AAP8AAP8AAP8A//8A//8A//8A//8A//8AgIAAgIAAAACAgIDAwMD/////"
      Columns(5).ValueItems(0).DisplayValue(23)=   "//8AgIAA//8A//8A//8A//8AAP8A//8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(24)=   "//8AAP8AAP8A//8AAP8A//8A//8A//8AgIAAAACAgIDAwMD///////8AgIAA//8A//8A//8AAP8A"
      Columns(5).ValueItems(0).DisplayValue(25)=   "AP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8A"
      Columns(5).ValueItems(0).DisplayValue(26)=   "//8A//8AgIAAAACAgIDAwMD///////8AgIAA//8A//8A//8A//8AAP8AAP8AAP8AAP8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(27)=   "//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8A//8A//8A//8AgIAAAACAgIDAwMD/"
      Columns(5).ValueItems(0).DisplayValue(28)=   "//////8AgIAA//8A//8A//8A//8A//8A//8A//8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(29)=   "//8AAP8AAP8A//8A//8A//8A//8A//8A//8AgIAAAACAgIDAwMD///////8AgIAA//////8A//8A"
      Columns(5).ValueItems(0).DisplayValue(30)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(31)=   "//8A//8A//8AgIAAAACAgIDAwMD///////8AgIAA//////8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(32)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AgIAAAACAgID/"
      Columns(5).ValueItems(0).DisplayValue(33)=   "//////////8AgIAA//////////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(34)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8AgIAAgIAAAADAwMD///////////////8AgID/////"
      Columns(5).ValueItems(0).DisplayValue(35)=   "//8A//8A//8A//8A//8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(36)=   "//8A//8A//8AgIAAAACAgID///////////////////8AgIAA//////////8A//8A//8A//8AAP8A"
      Columns(5).ValueItems(0).DisplayValue(37)=   "AP8AAP8AAP8AAP8A//8A//8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8AgIAAAADAwMD/"
      Columns(5).ValueItems(0).DisplayValue(38)=   "//////////////////////8AgID///////8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(39)=   "AP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8AgIAAAACAgID///////////////////////////8A"
      Columns(5).ValueItems(0).DisplayValue(40)=   "gIAA//////////8A//8A//8AAP8AAP8AAP8AAP8AAP8A//8A//8AAP8AAP8AAP8AAP8AAP8A//8A"
      Columns(5).ValueItems(0).DisplayValue(41)=   "//8A//8A//8AgIAAAAD///////////////////////////////////8AgIAA//////////8A//8A"
      Columns(5).ValueItems(0).DisplayValue(42)=   "//8AAP8AAP8AAP8A//8A//8A//8A//8AAP8AAP8AAP8A//8A//8A//8A//8AgIAAAAD/////////"
      Columns(5).ValueItems(0).DisplayValue(43)=   "//////////////////////////////////8AgIAA//////////////8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(44)=   "//8A//8A//8A//8A//8A//8A//8A//8AgIAAgID/////////////////////////////////////"
      Columns(5).ValueItems(0).DisplayValue(45)=   "//////////////8AgIAA//////////////////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(46)=   "//8AgIAAgID///////////////////////////////////////////////////////////8AgIAA"
      Columns(5).ValueItems(0).DisplayValue(47)=   "gIAA//////////////////8A//8A//8A//8A//8A//8A//8A//8AgIAAgID/////////////////"
      Columns(5).ValueItems(0).DisplayValue(48)=   "//////////////////////////////////////////////////////8AgIAAgIAA//8A//8A//8A"
      Columns(5).ValueItems(0).DisplayValue(49)=   "//8A//8A//8A//8A//8AgIAAgID/////////////////////////////////////////////////"
      Columns(5).ValueItems(0).DisplayValue(50)=   "//////////////////////////////////////8AgIAAgIAAgIAAgIAAgIAAgIAAgIAAgID/////"
      Columns(5).ValueItems(0).DisplayValue(51)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(0).DisplayValue(52)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(0).DisplayValue(53)=   "//////////////////////8="
      Columns(5).ValueItems(0).DisplayValue.vt=   9
      Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(1)._DefaultItem=   0
      Columns(5).ValueItems(1).Value=   "Rechazada"
      Columns(5).ValueItems(1).Value.vt=   8
      Columns(5).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(1).DisplayValue(0)=   "bHQAADYMAABCTTYMAAAAAAAANgAAACgAAAAgAAAAIAAAAAEAGAAAAAAAAAwAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(3)=   "///////////////////////////////////AwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMD/////////"
      Columns(5).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(5)=   "///AwMCAgICAgICAgICAgICAgICAgICAgICAgIDAwMDAwMDAwMD/////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(6)=   "///////////////////////////////////////////////AwMCAgICAgICAgIAAAAAAAAAAAAAA"
      Columns(5).ValueItems(1).DisplayValue(7)=   "AAAAAAAAAACAgICAgICAgIDAwMDAwMDAwMD/////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(8)=   "//////////////////+AgIAAgIAAgIAAgIAAgIAAgIAAgIAAgIAAgIAAgICAgIAAAAAAAACAgICA"
      Columns(5).ValueItems(1).DisplayValue(9)=   "gICAgIDAwMDAwMD///////////////////////////////////////////////////8AgIAAgIAA"
      Columns(5).ValueItems(1).DisplayValue(10)=   "//8AgIAA//8AgIAA//8AgIAA//8AgIAA//8AgIAAgICAgIAAAAAAAACAgICAgIDAwMDAwMD/////"
      Columns(5).ValueItems(1).DisplayValue(11)=   "//////////////////////////////////////8AgIAA//8A//8A//8A//8A//8A//8AgIAA//8A"
      Columns(5).ValueItems(1).DisplayValue(12)=   "gIAA//8AgIAA//8AgIAAgIAAgICAgIAAAACAgICAgIDAwMDAwMD/////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(13)=   "//////////8AgIAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AgIAA//8AgIAA"
      Columns(5).ValueItems(1).DisplayValue(14)=   "//8AgICAgIAAAACAgICAgIDAwMDAwMD///////////////////////////8AgIAA//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(15)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AgIAA//8AgIAA//8AgICAgIAAAACAgICA"
      Columns(5).ValueItems(1).DisplayValue(16)=   "gIDAwMD///////////////////////8AgIAA//8A//8A//8A//8AAP8AAP8A//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(17)=   "//8A//8A//8A//8A//8A//8AAP8AAP8AgIAA//8AgICAgIAAAACAgIDAwMDAwMD/////////////"
      Columns(5).ValueItems(1).DisplayValue(18)=   "//////8AgIAA//8A//8A//8A//8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(19)=   "AP8AAP8A//8AgIAA//8AgIAAAACAgICAgIDAwMD///////////////8AgIAA//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(20)=   "//8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AgIAA//8AgIAAgICA"
      Columns(5).ValueItems(1).DisplayValue(21)=   "gIAAAACAgIDAwMDAwMD///////////8AgIAA//8A//8A//8A//8A//8A//8AAP8AAP8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(22)=   "//8A//8A//8A//8A//8A//8AAP8AAP8AgIAA//8AgIAA//8AgIAAgIAAAACAgICAgIDAwMD/////"
      Columns(5).ValueItems(1).DisplayValue(23)=   "//8AgIAA//8A//8A//8A//8A//8A//8A//8AAP8AAP8AAP8AAP8A//8A//8A//8A//8AAP8AAP8A"
      Columns(5).ValueItems(1).DisplayValue(24)=   "AP8AAP8A//8AgIAA//8AgIAA//8AgICAgIAAAACAgIDAwMD///////8AgIAA//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(25)=   "//8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8AgIAA//8A"
      Columns(5).ValueItems(1).DisplayValue(26)=   "gIAAgICAgIAAAACAgIDAwMD///////8AgIAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(27)=   "AP8AAP8AAP8AAP8AAP8AAP8A//8A//8A//8A//8A//8A//8AgIAA//8AgICAgIAAAACAgIDAwMD/"
      Columns(5).ValueItems(1).DisplayValue(28)=   "//////8AgIAA//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(29)=   "//8A//8A//8A//8A//8AgIAA//8AgIAAgICAgIAAAACAgIDAwMD///////8AgIAA//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(30)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(31)=   "gIAA//8AgICAgIAAAACAgIDAwMD///////8AgIAA//////8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(32)=   "//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8AgIAA//8AgIAAgICAgIAAAACAgIDA"
      Columns(5).ValueItems(1).DisplayValue(33)=   "wMD///////8AgIAA//////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(34)=   "//8A//8A//8A//8A//8A//8A//8AgIAA//8AgICAgIAAAACAgID///////////8AgIAA////////"
      Columns(5).ValueItems(1).DisplayValue(35)=   "//8A//8A//8A//8A//8AAP8AAP8A//8A//8A//8A//8A//8A//8A//8A//8AAP8AAP8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(36)=   "gIAA//8AgIAAgICAgIAAAADAwMD///////////////8AgID///////8A//8A//8A//8A//8AAP8A"
      Columns(5).ValueItems(1).DisplayValue(37)=   "AP8AAP8AAP8AAP8A//8A//8AAP8AAP8AAP8AAP8AAP8A//8AgIAA//8AgIAA//8AgIAAAACAgID/"
      Columns(5).ValueItems(1).DisplayValue(38)=   "//////////////////8AgIAA//////////8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(39)=   "AP8AAP8AAP8AAP8AAP8A//8A//8AgIAA//8AgICAgIAAAADAwMD///////////////////////8A"
      Columns(5).ValueItems(1).DisplayValue(40)=   "gID///////8A//8A//8A//8AAP8AAP8AAP8AAP8AAP8A//8A//8AAP8AAP8AAP8AAP8AAP8A//8A"
      Columns(5).ValueItems(1).DisplayValue(41)=   "gIAA//8AgIAAgIAAAACAgID///////////////////////////8AgIAA//////////8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(42)=   "//8AAP8AAP8AAP8A//8A//8A//8A//8AAP8AAP8AAP8A//8A//8A//8AgIAA//+AgIAAAAD/////"
      Columns(5).ValueItems(1).DisplayValue(43)=   "//////////////////////////////8AgIAA//////////8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(44)=   "//8A//8A//8A//8A//8A//8A//8AgIAA//8AgIAAAAD/////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(45)=   "//////////8AgIAA//////////////8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A//8A"
      Columns(5).ValueItems(1).DisplayValue(46)=   "gIAA//8AgICAgID///////////////////////////////////////////////////8AgIAA////"
      Columns(5).ValueItems(1).DisplayValue(47)=   "//////////////8A//8A//8A//8A//8A//8A//8AgIAA//8AgIAA//8AgICAgID/////////////"
      Columns(5).ValueItems(1).DisplayValue(48)=   "//////////////////////////////////////////////8AgIAAgIAA//////////////////8A"
      Columns(5).ValueItems(1).DisplayValue(49)=   "//8AgIAA//8AgIAA//8AgIAA//8AgIAAgID/////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(50)=   "//////////////////////////////8AgIAAgIAA//8AgIAA//8AgIAA//8AgIAA//8AgIAAgIAA"
      Columns(5).ValueItems(1).DisplayValue(51)=   "gID/////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(52)=   "//////////////8AgIAAgIAAgIAAgIAAgIAAgIAAgIAAgID/////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(53)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(1).DisplayValue(54)=   "//////////////////////////////////////////////////////////////////////////8="
      Columns(5).ValueItems(1).DisplayValue.vt=   9
      Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(2)._DefaultItem=   0
      Columns(5).ValueItems(2).Value=   "Pendiente"
      Columns(5).ValueItems(2).Value.vt=   8
      Columns(5).ValueItems(2).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(5).ValueItems(2).DisplayValue(0)=   "bHQAADYMAABCTTYMAAAAAAAANgAAACgAAAAgAAAAIAAAAAEAGAAAAAAAAAwAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(2).DisplayValue(1)=   "AAAAAAD///////////////////////////////////////////////8AAAAAAAAAAAAAAAAAAAAA"
      Columns(5).ValueItems(2).DisplayValue(2)=   "AAAAAAAAAAD/////////////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(3)=   "//////////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/"
      Columns(5).ValueItems(2).DisplayValue(4)=   "//////////////////////////////////////////////////////////////8AAAAAAAAAAAAA"
      Columns(5).ValueItems(2).DisplayValue(5)=   "AAAAAAAAAACAAACAgACAAAAAAACAAACAgACAAAAAAAAAAAAAAAAAAAAAAAD/////////////////"
      Columns(5).ValueItems(2).DisplayValue(6)=   "//////////////////////////////////8AAAAAAAAAAAAAAACAAACAgACAAAAAgACAAACAgAAA"
      Columns(5).ValueItems(2).DisplayValue(7)=   "gACAgACAAAAAgAAAgACAgACAAACAgACAAAAAAAD/////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(8)=   "//////8AAAAAAAAAAACAgACAAAAAgACAAAAAgACAAAAAgAAA/wCAgAAAgAAA/wCAAAAAgAAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(9)=   "gACAAAAAgACAAAAAAAD///////////////////////////////////8AAAAAAACAAACAgACAAAAA"
      Columns(5).ValueItems(2).DisplayValue(10)=   "gACAAAAAgACAAAAAgACAAAAAgAAAgACAgAAAgAAA/wAAgAAA/wAAgAAA/wCAAAAAgAAAgAAAAAD/"
      Columns(5).ValueItems(2).DisplayValue(11)=   "//////////////////////////8AAAAAAAAAAAAAAACAAAAAgACAAAAAgAAA/wAAgAAA/wAAgAAA"
      Columns(5).ValueItems(2).DisplayValue(12)=   "gAAA/wAAgAAAgAAAgAAAgAAAgAAAgAAAgAAAgAAAgAAA/wAAgAAAAAD///////////////////8A"
      Columns(5).ValueItems(2).DisplayValue(13)=   "AAAAAACAAACAgACAAAAA/wCAAACAgAAAgAAAgAAA/wAAgAAAgAAA/wAAgAAAgAAA/wAAgAAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(14)=   "/wAAgAAA/wAAgAAAgAAAgAAAgAAAgAAAAAD///////////////8AAAAAAAAAAACAAAAAgACAAAAA"
      Columns(5).ValueItems(2).DisplayValue(15)=   "gAAAgAAAgAAAgAAAgAAAgAAA/wAAgAAAgAAA/wAAgAAAgAAAgAAAgAAAgAAAgAAAgAAA/wAAgAAA"
      Columns(5).ValueItems(2).DisplayValue(16)=   "/wAAgAAAAAD///////////8AAAAAAAAAAACAAAAAgACAAAAAgAAA/wAAgAAA/wAAgAAAgAAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(17)=   "gAAA/wAAgAAAgAAAgAAAgAAA/wAAgAAAgAAAgAAA/wAAgAAA/wAAgAAAgAAAgAAAAAD///////8A"
      Columns(5).ValueItems(2).DisplayValue(18)=   "AAAAAACAAACAgAAAgACAgAAAgAAAgAAAgAAAgAAAgAAA/wAAgAAA/wAAgAAAgAAA/wAA/wAAgAAA"
      Columns(5).ValueItems(2).DisplayValue(19)=   "gAAA/wAA/wAAgAAAgAAAgAAAgAAAgAAAgAAA/wAAAAD///////8AAAAAAACAgAAAgAAA/wAAgAAA"
      Columns(5).ValueItems(2).DisplayValue(20)=   "/wAAgAAAgAAAgAAAgAAAgAAAgAAAgAAAgAAA/wAAgAAA/wAAgAAA/wAAgAAAgAAA/wAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(21)=   "gAAA/wAA/wAAgAAAAAD///8AAAAAAACAgACAAACAgAAAgACAgAAAgAAAgAAAgAAA/wAAgAAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(22)=   "gAAA/wAA/wAAgAAA/wAAgAAA/wAAgAAA/wAAgAAAgAAAgAAA/wAAgAAA/wAAgAAAgAAA/wAAAAAA"
      Columns(5).ValueItems(2).DisplayValue(23)=   "AAAAAACAAACAgACAAAAAgAAAgAAA/wAAgAAAgAAAgAAA/wAAgAAAgAAAgAAA/wAAgAAAgAAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(24)=   "gAAA/wAAgAAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAAgAAAAAAAAAAAAAAAAACAAACAgAAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(25)=   "gAAAgAAAgAAAgAAA/wAAgAAAgAAA/wAAgAAAgAAAgAAA/wAAgAAA/wAAgAAA/wAA/wAAgAAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(26)=   "gAAAgAAA/wAAgAAA/wAAgAAAAAAAAAAAAACAAACAgAAAgACAgAAA/wAAgAAAgAAA/wAAgAAAgAAA"
      Columns(5).ValueItems(2).DisplayValue(27)=   "/wAAgAAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAAgAAA/wAA/wAA/wAA/wAAgAAAgAAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(28)=   "AAAAAAAAAACAgACAAACAgAAAgAAAgAAAgAAAgAAAgAAAgAAA/wAAgAAA/wAAgAAA/wAAgAAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(29)=   "gAAA/wAA/wAAgAAA/wAA/wAAgAAAgAAA/wAAgAAA/wAAgAAAgAAAAAAAAACAgACAAACAgAAAgAAA"
      Columns(5).ValueItems(2).DisplayValue(30)=   "/wAAgAAAgAAA/wAAgAAA/wAAgAAA/wAA/wAA/wAAgAAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(31)=   "/wAA/wAA/wAA/wAA/wAAgAAA/wAAAAAAAACAAACAgAAAgAAA/wAAgAAAgAAAgAAA/wAAgAAAgAAA"
      Columns(5).ValueItems(2).DisplayValue(32)=   "/wAAgAAA/wAA/wAA/wAA/wAAgAAA/wAAgAAA/wAA/wAAgAAA/wAA/wAA/wAA/wAAgAAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(33)=   "/wAAAAAAAACAgACAAAAAgAAAgAAAgAAA/wAAgAAAgAAAgAAA/wAAgAAAgAAA/wAAgAAAgAAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(34)=   "/wAAgAAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAAgAAA/wAAAAD///8AAACAgACAAAAA"
      Columns(5).ValueItems(2).DisplayValue(35)=   "gAAA/wAAgAAAgAAA/wAAgAAAgAAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(36)=   "/wAA/wAA/wAA/wAA/wAA/wAA/wAAAAD///////8AAACAAAAAgAAAgAAAgAAA/wAAgAAAgAAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(37)=   "gAAAgAAAgAAAgAAA/wAAgAAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(38)=   "gAAAAAD///////8AAACAgACAAACAgAAA/wAAgAAAgAAAgAAA/wAA/wAA/wAA/wAAgAAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(39)=   "/wAA/wAA/wAA/wAA/wAA/wD///8A/wAA/wD///8A/wAA/wAA/wAA/wAAAAD///////////8AAACA"
      Columns(5).ValueItems(2).DisplayValue(40)=   "gAAAgAAAgAAA/wAAgAAA/wAAgAAAgAAAgAAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(41)=   "/wD///8A/wD///////8A/wAA/wAAAAD///////////////8AAACAAACAgAAAgAAAgAAAgAAAgAAA"
      Columns(5).ValueItems(2).DisplayValue(42)=   "/wAA/wAAgAAAgAAA/wAAgAAA/wAA/wAA/wAA/wAA/wAA/wD///8A/wAA/wD///8A/wAA/wD///8A"
      Columns(5).ValueItems(2).DisplayValue(43)=   "/wAAAAD///////////////////8AAAAAgACAgAAAgAAA/wAA/wAAgAAAgAAA/wAA/wAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(44)=   "/wAAgAAA/wAA/wAA/wAA/wAA/wAA/wD///8A/wD///8A/wAA/wAAAAD/////////////////////"
      Columns(5).ValueItems(2).DisplayValue(45)=   "//////8AAAAAgAAAgAAAgAAAgAAAgAAA/wAAgAAA/wAA/wAAgAAAgAAA/wAA/wAA/wAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(46)=   "/wAA/wAA/wAA/wAA/wAA/wAAAAD///////////////////////////////////8AAAAAgAAAgAAA"
      Columns(5).ValueItems(2).DisplayValue(47)=   "gAAAgAAAgAAA/wAA/wAAgAAA/wAA/wAAgAAA/wAA/wAA/wAA/wAA/wD///8A/wD///8A/wAAAAD/"
      Columns(5).ValueItems(2).DisplayValue(48)=   "//////////////////////////////////////////8AAAAAgAAAgAAAgAAA/wAAgAAAgAAAgAAA"
      Columns(5).ValueItems(2).DisplayValue(49)=   "gAAA/wAA/wAA/wAAgAAA/wAA/wAA/wAA/wAA/wAA/wAAAAD/////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(50)=   "//////////////////////8AAAAAAAAAgAAAgAAAgAAAgAAA/wAAgAAA/wAAgAAA/wAA/wAA/wAA"
      Columns(5).ValueItems(2).DisplayValue(51)=   "gAAA/wAA/wAAAAAAAAD/////////////////////////////////////////////////////////"
      Columns(5).ValueItems(2).DisplayValue(52)=   "//////8AAAAAAAAAAAAAgAAAgAAAgAAAgAAAgAAA/wAAgAAAgAAAAAAAAAAAAAD/////////////"
      Columns(5).ValueItems(2).DisplayValue(53)=   "//////////////////////////////////////////////////////////////////////8AAAAA"
      Columns(5).ValueItems(2).DisplayValue(54)=   "AAAAAAAAAAAAAAAAAAAAAAAAAAD///////////////////////////////////////////////8="
      Columns(5).ValueItems(2).DisplayValue.vt=   9
      Columns(5).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems.Count=   3
      Columns(5).Caption=   "Estado"
      Columns(5).DataField=   "descest"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2963"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2884"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=528"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1746"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1667"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=529"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=3069"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2990"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=532"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=4445"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=4366"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=532"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1455"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1376"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=532"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2487"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2408"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=532"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      DefColWidth     =   0
      HeadLines       =   1.5
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE2FEFD&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
      _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.valignment=2,.bgcolor=&H40&"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   630
      Left            =   3885
      TabIndex        =   5
      Top             =   5430
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   1111
      _StockProps     =   79
      Caption         =   "Rechazar Cotizacion"
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmApruebaCotiza.frx":0634
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   630
      Left            =   7590
      TabIndex        =   6
      Top             =   5430
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   1111
      _StockProps     =   79
      Caption         =   "Salir"
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmApruebaCotiza.frx":094E
   End
   Begin XtremeSuiteControls.PushButton PushButton4 
      Height          =   630
      Left            =   5610
      TabIndex        =   7
      Top             =   5430
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   1111
      _StockProps     =   79
      Caption         =   "Dejar Pendiente"
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmApruebaCotiza.frx":0C68
   End
End
Attribute VB_Name = "FrmApruebaCotiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset

Private Sub Form_Activate()
    RST_Busq Rst, "SELECT com_ordencot.id, [com_ordencot]![numser] & '-' & [com_ordencot]![numdoc] AS numdoc, com_ordencot.fchemi, " _
        & " pla_empleados.nombre, com_ordencot.idest, mae_estados.descripcion AS descest, com_ordencot.idsit " _
        & " FROM (pla_empleados RIGHT JOIN (com_ordencot LEFT JOIN com_usuario ON com_ordencot.idsol = com_usuario.id) ON pla_empleados.id = com_usuario.idper) " _
        & " LEFT JOIN mae_estados ON com_ordencot.idest = mae_estados.id WHERE (((com_ordencot.idsit)=0))", xCon
   
    Dg1.DataSource = Rst
End Sub

Private Sub PushButton1_Click()
    xCon.Execute "UPDATE com_ordencot SET com_ordencot.idest = 2 WHERE (((com_ordencot.id)=" & Rst("id") & "))"
    Rst.Requery
    Dg1.Refresh
End Sub

Private Sub PushButton2_Click()
    xCon.Execute "UPDATE com_ordencot SET com_ordencot.idest = 4 WHERE (((com_ordencot.id)=" & Rst("id") & "))"
    Rst.Requery
    Dg1.Refresh
End Sub

Private Sub PushButton3_Click()
    Set Rst = Nothing
    Unload Me
End Sub

Private Sub PushButton4_Click()
    xCon.Execute "UPDATE com_ordencot SET com_ordencot.idest = 1 WHERE (((com_ordencot.id)=" & Rst("id") & "))"
    Rst.Requery
    Dg1.Refresh
End Sub

Private Sub PushButton5_Click()
    FrmManOrdCotiza.DeDonde = 2
    FrmManOrdCotiza.xIdOR = Rst("id")
    FrmManOrdCotiza.Show vbModal
End Sub

