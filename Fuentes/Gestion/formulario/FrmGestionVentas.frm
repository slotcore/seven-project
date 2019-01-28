VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGestionVentas 
   Caption         =   "Gestión - Análisis de Ventas"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   615
   ClientWidth     =   12270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   12270
   Begin VB.Frame FraGraf1 
      Height          =   2385
      Left            =   4320
      TabIndex        =   47
      Top             =   3105
      Visible         =   0   'False
      Width           =   3525
      Begin VB.CommandButton CmdGrafCancel1 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   1800
         TabIndex        =   59
         Top             =   1950
         Width           =   1560
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mostrar"
         Height          =   765
         Left            =   180
         TabIndex        =   56
         Top             =   1500
         Width           =   1515
         Begin VB.CheckBox ChkLeyenda 
            Caption         =   "Leyenda"
            Height          =   195
            Left            =   210
            TabIndex        =   57
            Top             =   300
            Width           =   1005
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Con Datos"
         Height          =   1110
         Left            =   180
         TabIndex        =   53
         Top             =   360
         Width           =   1515
         Begin VB.OptionButton OptconDatosDetalle1 
            Caption         =   "Detallado"
            Height          =   210
            Left            =   165
            TabIndex        =   55
            Top             =   645
            Width           =   1035
         End
         Begin VB.OptionButton OptConDatoResum1 
            Caption         =   "Resumido"
            Height          =   195
            Left            =   165
            TabIndex        =   54
            Top             =   315
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin VB.CommandButton CmdGrafAcep1 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   1800
         TabIndex        =   52
         Top             =   1530
         Width           =   1560
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Gráfico"
         Height          =   1110
         Left            =   1800
         TabIndex        =   48
         Top             =   360
         Width           =   1560
         Begin VB.OptionButton OptTipGrafCircular 
            Caption         =   "Circular"
            Height          =   195
            Left            =   165
            TabIndex        =   51
            Top             =   795
            Width           =   1290
         End
         Begin VB.OptionButton OptTipGrafLinea 
            Caption         =   "Lineas"
            Height          =   195
            Left            =   165
            TabIndex        =   50
            Top             =   547
            Width           =   1290
         End
         Begin VB.OptionButton OptTipGrafBarra1 
            Caption         =   "Barras"
            Height          =   195
            Left            =   165
            TabIndex        =   49
            Top             =   300
            Value           =   -1  'True
            Width           =   1290
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "  Propiedades de gráfico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   58
         Top             =   0
         Width           =   3885
      End
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3375
      TabIndex        =   8
      Top             =   3615
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   9
         Top             =   465
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   34
         Top             =   795
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   5025
         TabIndex        =   36
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   5025
         TabIndex        =   35
         Top             =   495
         Width           =   825
      End
      Begin VB.Label lbl 
         Caption         =   "Interrumpir = ESC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   4275
         TabIndex        =   12
         Top             =   150
         Width           =   1530
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Procesando:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   11
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Ventas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   150
         Width           =   585
      End
      Begin VB.Shape Shape1 
         Height          =   1065
         Left            =   90
         Top             =   60
         Width           =   5805
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Gráfico"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4860
         Top             =   90
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":2A98
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmGestionVentas.frx":2E2A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fr 
      Height          =   2595
      Index           =   5
      Left            =   0
      TabIndex        =   2
      Top             =   330
      Width           =   11805
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
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
         Height          =   660
         Index           =   4
         Left            =   7140
         TabIndex        =   22
         Top             =   120
         Width           =   1515
         Begin VB.CommandButton CmdMas 
            Height          =   225
            Index           =   1
            Left            =   1200
            Picture         =   "FrmGestionVentas.frx":327C
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Reducir"
            Top             =   1110
            Width           =   285
         End
         Begin VB.CommandButton CmdMas 
            Height          =   225
            Index           =   0
            Left            =   1200
            Picture         =   "FrmGestionVentas.frx":337E
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Ampliar"
            Top             =   420
            Width           =   285
         End
         Begin VB.OptionButton opt_totalizar 
            Caption         =   "Precio Máx."
            Enabled         =   0   'False
            Height          =   195
            Index           =   4
            Left            =   75
            TabIndex        =   68
            Top             =   1125
            Width           =   1155
         End
         Begin VB.OptionButton opt_totalizar 
            Caption         =   "Precio Prom."
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   75
            TabIndex        =   67
            Top             =   891
            Width           =   1215
         End
         Begin VB.OptionButton opt_totalizar 
            Caption         =   "Precio Mìn."
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   66
            Top             =   659
            Width           =   1095
         End
         Begin VB.OptionButton opt_totalizar 
            Caption         =   "Cantidades"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   24
            Top             =   427
            Width           =   1155
         End
         Begin VB.OptionButton opt_totalizar 
            Caption         =   "Importe"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   23
            Top             =   195
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
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
         Height          =   690
         Index           =   7
         Left            =   2850
         TabIndex        =   63
         Top             =   840
         Width           =   1155
         Begin VB.OptionButton OptFecha 
            Caption         =   "Fch. Doc."
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   65
            Top             =   195
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton OptFecha 
            Caption         =   "Fch. Reg."
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   64
            Top             =   420
            Width           =   1005
         End
      End
      Begin VB.CheckBox chkAnioPasados 
         Caption         =   "Considerar Años Anteriores"
         Height          =   195
         Left            =   8700
         TabIndex        =   62
         Top             =   1185
         Width           =   2595
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar Importe"
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
         Height          =   495
         Index           =   6
         Left            =   8685
         TabIndex        =   43
         Top             =   120
         Width           =   3075
         Begin VB.OptionButton opt_importe 
            Caption         =   "Sólo Igv"
            Height          =   195
            Index           =   2
            Left            =   2040
            TabIndex        =   46
            Top             =   225
            Width           =   885
         End
         Begin VB.OptionButton opt_importe 
            Caption         =   "Sin Igv"
            Height          =   195
            Index           =   1
            Left            =   1125
            TabIndex        =   45
            Top             =   225
            Width           =   825
         End
         Begin VB.OptionButton opt_importe 
            Caption         =   "Con Igv"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   225
            Value           =   -1  'True
            Width           =   900
         End
      End
      Begin VB.ListBox ls 
         Height          =   1185
         Index           =   1
         Left            =   4050
         Style           =   1  'Checkbox
         TabIndex        =   40
         Top             =   315
         Width           =   1650
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
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
         Height          =   660
         Index           =   2
         Left            =   2850
         TabIndex        =   37
         Top             =   135
         Width           =   1155
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Trimestre"
            Height          =   210
            Index           =   1
            Left            =   60
            TabIndex        =   42
            Top             =   420
            Width           =   960
         End
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Mes"
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   39
            Top             =   195
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Semestre"
            Height          =   210
            Index           =   2
            Left            =   60
            TabIndex        =   38
            Top             =   735
            Visible         =   0   'False
            Width           =   960
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Tipo Consulta"
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
         Height          =   1380
         Index           =   1
         Left            =   30
         TabIndex        =   28
         Top             =   120
         Width           =   1500
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x T.Documento"
            Height          =   195
            Index           =   5
            Left            =   45
            TabIndex        =   61
            Top             =   1140
            Width           =   1410
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x T. Prod/Item"
            Height          =   195
            Index           =   4
            Left            =   45
            TabIndex        =   33
            Top             =   945
            Width           =   1380
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x Vendedor"
            Height          =   195
            Index           =   3
            Left            =   45
            TabIndex        =   32
            Top             =   750
            Width           =   1380
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x Pto de Venta"
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   31
            Top             =   555
            Width           =   1380
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x Año"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   30
            Top             =   165
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x Cliente"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   29
            Top             =   360
            Width           =   1380
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Presentación"
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
         Height          =   705
         Index           =   3
         Left            =   7140
         TabIndex        =   25
         Top             =   795
         Width           =   1485
         Begin VB.OptionButton opt_escala 
            Caption         =   "En Decimales"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   195
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton opt_escala 
            Caption         =   "En Miles"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   420
            Width           =   1275
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
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
         Height          =   1380
         Index           =   0
         Left            =   5760
         TabIndex        =   13
         Top             =   120
         Width           =   1335
         Begin VB.OptionButton opt_mon 
            Caption         =   "Todo en ME"
            Height          =   210
            Index           =   3
            Left            =   75
            TabIndex        =   17
            Top             =   945
            Width           =   1200
         End
         Begin VB.OptionButton opt_mon 
            Caption         =   "Todo en MN"
            Height          =   210
            Index           =   2
            Left            =   75
            TabIndex        =   16
            Top             =   710
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt_mon 
            Caption         =   "Solo ME"
            Height          =   210
            Index           =   1
            Left            =   75
            TabIndex        =   15
            Top             =   475
            Width           =   960
         End
         Begin VB.OptionButton opt_mon 
            Caption         =   "Solo MN"
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   14
            Top             =   240
            Width           =   1005
         End
      End
      Begin VB.ListBox ls 
         Height          =   1185
         Index           =   0
         Left            =   1590
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   315
         Width           =   1200
      End
      Begin VB.CheckBox ChkMostrarItem 
         Caption         =   "Mostrar item"
         Height          =   195
         Left            =   10110
         TabIndex        =   4
         Top             =   1575
         Width           =   1155
      End
      Begin VB.CommandButton CmdBusProducto 
         Height          =   225
         Left            =   9075
         Picture         =   "FrmGestionVentas.frx":3480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   855
         Width           =   210
      End
      Begin VB.TextBox TxtIdTipProd 
         Height          =   300
         Left            =   8700
         MaxLength       =   3
         TabIndex        =   1
         Top             =   810
         Width           =   615
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1020
         Index           =   0
         Left            =   60
         TabIndex        =   18
         Top             =   1515
         Width           =   2835
         _cx             =   5001
         _cy             =   1799
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
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmGestionVentas.frx":35B2
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1020
         Index           =   1
         Left            =   3010
         TabIndex        =   19
         Top             =   1515
         Width           =   2835
         _cx             =   5001
         _cy             =   1799
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
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmGestionVentas.frx":360D
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1020
         Index           =   2
         Left            =   5960
         TabIndex        =   20
         Top             =   1515
         Width           =   2835
         _cx             =   5001
         _cy             =   1799
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
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmGestionVentas.frx":366F
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1020
         Index           =   3
         Left            =   8910
         TabIndex        =   21
         Top             =   1515
         Width           =   2835
         _cx             =   5001
         _cy             =   1799
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
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmGestionVentas.frx":36CB
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Selec. Mes"
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
         Height          =   195
         Index           =   6
         Left            =   4095
         TabIndex        =   41
         Top             =   120
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Selec. Año"
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
         Height          =   195
         Index           =   3
         Left            =   1575
         TabIndex        =   7
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "T.Producto"
         Height          =   165
         Left            =   8700
         TabIndex        =   6
         Top             =   630
         Width           =   795
      End
      Begin VB.Label lblTipProducto 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   9330
         TabIndex        =   5
         Top             =   810
         Width           =   2340
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   4980
      Left            =   0
      TabIndex        =   71
      Top             =   2940
      Width           =   11820
      _cx             =   20849
      _cy             =   8784
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
      BackColor       =   14745342
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14745342
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmGestionVentas.frx":3723
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
   Begin VB.Menu mn_cliente 
      Caption         =   "Cliente"
      Visible         =   0   'False
      Begin VB.Menu mn_cliente1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_cliente2 
         Caption         =   "Seleccionar"
      End
   End
   Begin VB.Menu mn_ptoventa 
      Caption         =   "PtoVenta"
      Visible         =   0   'False
      Begin VB.Menu mn_ptoventa1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_ptoventa2 
         Caption         =   "Seleccionar"
      End
   End
   Begin VB.Menu mn_vendedor 
      Caption         =   "Vendedor"
      Visible         =   0   'False
      Begin VB.Menu mn_vendedor1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mn_vendedor2 
         Caption         =   "Seleccionar"
      End
   End
   Begin VB.Menu mn_item 
      Caption         =   "Item"
      Visible         =   0   'False
      Begin VB.Menu mn_item1 
         Caption         =   "Agegar"
      End
      Begin VB.Menu mn_item2 
         Caption         =   "Seleccionar"
      End
   End
End
Attribute VB_Name = "FrmGestionVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--PARA EXPORTAR A EXCEL
Dim Oleapp As Object
Dim vCantMeses As Integer
'--VARIABLES DE PROPIEDADES DE GRAFICO
Dim vLngTipoGrafico As Long, vTipoDato As Integer
Dim vTituloGraf As String, vViewLeyenda As Boolean
'--FIN PARA EXPORTAR A EXCEL


'-- ALMACENAR LOS TOTALES DE TODA LA CONSULTA
Dim ArrayTotalGral() As Double '--ALMACENAR TOTALES POR TODAS LAS FILAS

Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE
'------------
Dim ARR_ANYO() As String    '--ARRAY DE AÑOS SELECCIONADOS
Dim ARR_MES() As String      '--SE CARGARA CUANDO SE CARGA EL FORMULARIO Y CUANDO SE CAMBIE EL ESTILO(MES, TRIMESTRE,SEMESTRE)
Dim ARR_TMP() As String     '--DEPENDERA DEL ESTILO SOLO CARGARA LO QUE SELECCIONA


                            '--SE USA PARA DAR FORMATO DE LA GRILLA, SEGUN SELECCIONE EL USUARIO
Dim Q_TOTAL_ANYO As Integer '--INDICA LA CANTIDAD DE AÑOS DE BUSQUEDA,
                            '--EJ. 2004,2005 => Q_TOTAL_ANYO = 2
                            '--EJ. 2004,2005,2006 => Q_TOTAL_ANYO = 3
                            
Dim Q_COL_FILA As Integer   '--INDICA LA CANTIDAD DE COLUMNAS ANTES DE LOS MESES
                            '--EJ. IDCLI,CLIENTE => Q_COL_FILA=2
                            '--    IDCLI,ID_PTO_VTA,CLIENTE,PTO_VENTA => Q_COL_FILA=4
                            
Dim Q_POS_MES_INICIO As Integer '--INDICA LA POSICION INICIAL DE LA COLUMNA DEL PRIMER MES, NO CAMBIA
                            '--EJ. Q_POS_MES_INICIO = Q_COL_FILA +1

Dim Q_POS_MES As Integer    '--INDICA LA POSICION DEL MES, ESTO CAMBIA
                            '--UTIL PARA COLOCAR LOS DATOS EN EL GRID

Dim Q_COL_FILA_OCULTA As Integer '--INDICA LAS COLUMNAS QUE CONTENDRAN LOS ID'S, ESTOS SE OCULTARAN
                                '-- -1 NO SE OCULTA, <> -1 SE PROCEDE A ACULTAR
                                'EJ. CLIENTE  vta_ventas.idcli,
                                    'PUNTO DE VENTA vta_guia.idpunven
                                    'PRODUCTO   alm_inventario.tippro
                                    'ITEM       alm_inventario.id
                                    'EMPLEADO   vta_ventas.idven

Dim Q_POSICION_TOTAL  As Integer '--INDICA LA POCISION DE LA COLUMNA DONDE SE COLOCARA EL NOMBRE DEL TOTAL Y TOTAL_GRL
                                 '--OBTENDRA VALOR EN fGenerarConsulta()

Dim Q_COL_COMPARAR_GRUPO As Integer '--INDICA LA COLUMNA PARA COMPARAR EL GRUPO
                                    '--OBTENDRA VALOR EN fGenerarConsulta()

Dim Q_COL_ARR_TOTAL As Integer  '--NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS
                                '--OBTENDRA VALOR EN VALIDAR_CONSULTA()
                                '--SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                '--SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0

Dim M_FILA_INICIO_GRUPO As Long '--INDICA LA POSICION INICIAL DE GRUPO PARA LUEGO TOTALIZARLA
Dim Q_POS_ARRAY As Integer      '--INDICA LA POSICION DEL ARREGLO PARA ALMACENAR EL TOTAL GENERAL

Dim F_SELECCION As Boolean '--INDICA SI SE VA SELECCIONAR LOS REGISTROS DE: CLIENTES, PTO VENTA, VENDEDOR, PRODUCTO
                           '--FALSE = SELECCIONA UN REGISTRO; TRUE = SELECCIONAR VARIOS REGISTROS
'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long



Private Sub CmdBusProducto_Click()
On Error GoTo error
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha

    Dim xCampos(2, 4) As String

    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"

    xform.SQLCad = "SELECT id, descripcion FROM mae_tipoproducto "

    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If TxtIdTipProd.Text <> "" And TxtIdTipProd.Text <> CStr(xRs.Fields("id")) Then LimpiarGrid Fg(3), True
        TxtIdTipProd.Text = xRs("id")
        lblTipProducto.Caption = xRs("descripcion")
    End If
    
    ChkMostrarItem_Click
     
    Set xform = Nothing
    Set xRs = Nothing
    Exit Sub
error:
    Set xform = Nothing
    Set xRs = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub

Private Sub CONSULTAR()
    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    Dim rstTemp As New ADODB.Recordset '--rst temporal para consulta de años anteriores
    '--
    Dim CN_TMP As New ADODB.Connection '--Conexion temporal
    Dim Rst_RUTA As New ADODB.Recordset '--Lista de rutas segun años seleccionados
    
    Dim vStrSelect As String '--RECIBIR LA CONSULTA
    '--CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    Dim N_ANYO As String
    Dim nSQLAnyo As String
    Dim k As Integer '--recorrer todos los años
    Dim F_CARGAR_1RA_VEZ As Boolean '--TRUE::SE CARGA POR 1RA VEZ LA GRILLA
    
    If Validar_Consulta(N_ANYO) = False Then Exit Sub
    
    BAND_INTERRUMPIR = False
    
    '--CONFIGURAR LA PRESENTACION DE LA CONSULTA
    LimpiarGrid Fg1
    
    '--INVOCAR A ESTA FUNCION PARA OBTENER LOS VALORES DE
        '--Q_POS_MES , Q_POS_MES_INICIO
    fGenerarConsulta "-1"
    Configurar_Grilla
    
    '--CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    nSQLAnyo = " AND anotra IN (" + Left(N_ANYO, Len(N_ANYO) - 1) + ") "
    
    '--SI LA BASE DE BATOS PRINCIPAL EXISTE
    If ArchivoExiste(AP_RUTABD + "data.mdb") = False Then
        MsgBox "No existe la ruta a la Base de Datos Principal", vbCritical, "Mensaje..."
        Exit Sub
    End If
    
    '--ABRIENDO LA CONEXION PARA OBTENER EL LISTADO DE RUTAS A LAS BASES DE DATOS
    OPEN_CONEX_TMP CN_TMP, AP_RUTABD + "data.mdb"
    If CN_TMP.State = 0 Then Exit Sub
    
    RST_Busq rst_select, "SELECT ruta,anotra FROM mae_empresa WHERE numruc = '" + NumRUC + "' " + nSQLAnyo + " ORDER BY anotra ASC ", CN_TMP
    
    '--CARGAR RST TEMPORAL
    DEFINIR_RST_TMP Rst_RUTA, rst_select
    CARGAR_RST_TMP Rst_RUTA, rst_select
    If Rst_RUTA.RecordCount = 0 Then
        MsgBox "No hay Base de Datos", vbInformation
        Exit Sub
    End If
    Rst_RUTA.MoveFirst
    
    Set rst_select = Nothing
    
    CN_TMP.Close
    
    '----
    MousePointer = vbHourglass
    DoEvents
    '--Posicionar la barra de progreso
    PgBar(1).Min = 0
    PgBar(1).Value = 0
    PosicionarProgBar
    DoEvents
    PgBar(0).Min = 0
    PgBar(0).Max = Rst_RUTA.RecordCount
    
    '--definir el arreglo para acumular los subtotales y mostrar en total general
    ReDim ArrayTotalGral((Q_COL_ARR_TOTAL + 3) * Q_TOTAL_ANYO)
    
    For k = 0 To Rst_RUTA.RecordCount - 1
    
        lbl(4).Caption = "Año: " + CStr(Rst_RUTA.Fields(1))
        DoEvents
        PgBar(0).Value = k + 1
        
        '------------------------------------------------
        If k = 0 Then
            '--ENTRAR SOLO UNA VEZ
            vStrSelect = fGenerarConsulta(CStr(Rst_RUTA.Fields(1)))
        Else
            '--EN LOS DEMAS AÑO REEMPLAZAR EL AÑO ANTERIOR POR EL AÑO ACTUAL
            vStrSelect = Replace(vStrSelect, ARR_ANYO(k - 1), CStr(Rst_RUTA.Fields(1)))
        End If
        
        '------------------------------------------------
        If vStrSelect = "" Then GoTo salir
        '--SI EL ARCHIVO EXISTE
        If ArchivoExiste(AP_RUTABD + Rst_RUTA.Fields(0) & "") = False Then
            MsgBox "No existe la ruta a la Base de Datos Año: " + CStr(Rst_RUTA.Fields(1)), vbCritical, "Mensaje..."
            GoTo salir
        End If
        
        '--ABRIENDO LA CONEXION A LA BASE DE DATOS
        OPEN_CONEX_TMP CN_TMP, AP_RUTABD + Rst_RUTA.Fields(0) & ""
        If CN_TMP.State = 0 Then Exit Sub
        
        '--CARGADO EL RST
        Set rst_select = Nothing
        RST_Busq rst_select, vStrSelect, CN_TMP

        '--------------------------------------
        '--Generar primer grupo
        If opt_consulta(0).Value = True And (Me.TxtIdTipProd.Text <> "" Or ChkMostrarItem.Value = 1) Then
            Comparar_Grupo Rst_RUTA, False, CStr(Rst_RUTA.Fields(1)), 1
        End If
        '--------------------------------------
        
        If rst_select.RecordCount > 0 Then
            If F_CARGAR_1RA_VEZ = False Or opt_consulta(0).Value = True Then
                '--CARGA LOS DATOS DEL PRIMER AÑO
                CARGAR_DATOS rst_select, CStr(Rst_RUTA.Fields(1))
                F_CARGAR_1RA_VEZ = True
                
            Else
                '--CUANDO LOS DATOS ESTAN CARGADOS => AGREGAR DATOS A LOS DEMAS AÑOS
                
                '--definir la estructura del rst temporal
                '--agregar nuevo campo xsel para identificar los registros pendientes de mostrar en pantalla
                Set rstTemp = Nothing
                DEFINIR_RST_TMP rstTemp, rst_select, "xsel"
                '--cargar los datos al rst temporal
                '--asignar por defecto al nuevo campo valor = 0
                CARGAR_RST_TMP rstTemp, rst_select, "xsel", 0
                
                rst_select.Close
                '--Volver a asignar el recordset al incial
                Set rst_select = rstTemp.Clone
                '--limpiar el rst temporal
                Set rstTemp = Nothing
                
                CARGAR_DATOS_GRILLA_OTROS_ANYOS rst_select, CStr(Rst_RUTA.Fields(1))
            End If
        End If
        
        CN_TMP.Close
        '--------------------------------------
        Rst_RUTA.MoveNext
        
    Next k
    
    '---CUANDO LA CONSULTA ES X AÑOS COLOCAR LOS TOTALES
    If (NulosN(TxtIdTipProd.Text) = 0 And ChkMostrarItem.Value = 0) Or opt_consulta(0).Value = True Then
        '--Ordenar columnas cuando sea x Cliente
        If opt_consulta(1).Value = True Then
            GRID_ORDENAR Fg1, Fg1.FixedRows, 3, , , flexSortGenericAscending
        End If
        '---------------------------------------------------------------------
        
        CARGAR_DATOS_TOTALES True, "Tot Gen:", True, True, ARR_ANYO(k - 1)
    Else
        CARGAR_DATOS_TOTALES True, "Tot Gen:", True, False, ARR_ANYO(k - 1)
    End If
    
    PgBar(0).Value = PgBar(0).Max
salir:
    FraProgreso.Visible = False
    Set Rst_RUTA = Nothing
    Set rst_select = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    CN_TMP.Close
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub


Private Function CARGAR_DATOS(RST_ORIGEN As ADODB.Recordset, M_ANYO As String)
    '===================================================================================================
    'Creado : xxxxxx Por: Johan Castro
    'Propósito: Recorrer todos los registros del Recordset a fin de mostrarlos en pantalla
    '
    'Entradas:  RST_ORIGEN= Recordset que contiene todos los datos de la consulta
    '           M_ANYO= Año que se esta consultando
    '
    'Resultados: Consulta en pantalla
    '
    '===================================================================================================
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim BAND_ADD_REG As Boolean
    
    
    BAND_ADD_REG = True
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    
    PgBar(1).Min = 0
    PgBar(1).Value = 1
    PgBar(1).Max = RST_ORIGEN.RecordCount
    
    While Not RST_ORIGEN.EOF
    
        DoEvents
        '--Verificar si prescionan tecla ESC para interrumpir la cosulta
        If BAND_INTERRUMPIR = True Then Exit Function
        '---------------------------------------------------------
        '--armar grupo
        Comparar_Grupo RST_ORIGEN, BAND_ADD_REG, M_ANYO
        '---------------------------------------------------------
        '--agregar fila
        ADD_REG Fg1
        '--cargar datos al grid(datos en pantalla)
        CARGAR_DATOS_GRILLA RST_ORIGEN, M_ANYO, Fg1.Rows - 1
        '---------------------------------------------------------
        '--avanzar al siguiente registro
        RST_ORIGEN.MoveNext
        
        '--poner el total del ultimo grupo o Total General
        If RST_ORIGEN.EOF Then
            '--Agregar totales del grupo
            If Q_COL_COMPARAR_GRUPO <> -1 Then
                CARGAR_DATOS_TOTALES BAND_ADD_REG, "Total:", , , M_ANYO
            End If
        
'            If opt_consulta(0).Value = False Then
'                CARGAR_DATOS_TOTALES BAND_ADD_REG, "Total:", , , M_ANYO
'                '--agregar el total general
'                If NulosN(TxtIdTipProd.Text) <> 0 Or ChkMostrarItem.Value = 1 Then
'                    CARGAR_DATOS_TOTALES True, "Tot Gen:", True, True, M_ANYO
'                End If
'            End If
        Else
            '--Recorrer la barra de progreso
            PgBar(1).Value = PgBar(1).Value + 1
            
        End If
    Wend
    
    PgBar(1).Value = 0
    
    
    
End Function

Private Sub Comparar_Grupo(RST_ORIGEN As ADODB.Recordset, _
                            BAND_ADD_REG As Boolean, _
                            M_ANYO As String, _
                            Optional Q_COL_COMPARAR As Integer = -1)
    '===================================================================================================
    'Creado : xxxxxx Por: Johan Castro
    'Propósito: Agregar una fila para agregar nuevo grupo
    '
    'Entradas:  RST_ORIGEN=Recordset que contiene todos los datos de la consulta
    '           BAND_ADD_REG=Año que se esta consultando
    '           M_ANYO=Año que se esta consultando
    '           Q_COL_COMPARAR=Indica la posicion de la columna para generar el grupo
    '                          toma un valor en fGenerarConsulta()
    '
    'Resultados: Consulta en pantalla por grupos
    '
    '===================================================================================================

    '--COMPARA CUANDO CAMBIAR DE GRUPO
    Dim RST_TEPM_1 As New ADODB.Recordset
    
    '---------------------------------------------------------
    If Q_COL_COMPARAR_GRUPO = -1 Then GoTo salir
    '---------------------------------------------------------
    If Q_COL_COMPARAR = -1 Then Q_COL_COMPARAR = Q_COL_COMPARAR_GRUPO
    
    
    If RST_ORIGEN.Bookmark = 1 Then
        '--SE CARGA EN fGenerarConsulta() Q_COL_COMPARAR_GRUPO
        ADD_REG Fg1, Fila_grupo
        UNIR_CELDAS Fg1, Fg1.Rows - 1, Q_COL_COMPARAR + 1, Fg1.Rows - 1, Q_POS_MES_INICIO - 1, RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter:
        FORMATO_CELDA Fg1, Fg1.Rows - 1, Q_COL_COMPARAR_GRUPO + 1
        'Asignar posicion de inicio de grupo para totalizar los grupos
        M_FILA_INICIO_GRUPO = Fg1.Rows
    Else
    
        Set RST_TEPM_1 = RST_ORIGEN.Clone
        RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
        RST_TEPM_1.MovePrevious
        
        '--Comparar la actual fila de la grilla con el nuevo registro del recordset
        '--si son distintos entonces totalizar grupo e insertar nuevo grupo
        If RST_TEPM_1.Fields(Q_COL_COMPARAR) <> RST_ORIGEN.Fields(Q_COL_COMPARAR) Then
            
            '--Verificar que la fila anterior no este totalizada
            '--se da cuando muestra datos de otros años, por unica vez al inicio de la carga de datos
            If Fg1.TextMatrix(Fg1.Rows - 1, Q_POSICION_TOTAL) <> "Total:" Then
                '--Totalizar el grupo
                CARGAR_DATOS_TOTALES BAND_ADD_REG, "Total:", , , M_ANYO
            End If
            
            '--Agregar fila en grid para separar el grupo anterior con el nuevo grupo
            ADD_REG Fg1, Fila_en_Blanco
            UNIR_CELDAS Fg1, Fg1.Rows - 1, IIf(Q_COL_FILA_OCULTA = -1, 1, Q_COL_FILA_OCULTA + 1), Fg1.Rows - 1, Fg1.Cols - 1, " ", flexAlignLeftCenter
            
            '--Agregar registro para nuevo grupo
            ADD_REG Fg1, Fila_grupo
            UNIR_CELDAS Fg1, Fg1.Rows - 1, Q_COL_COMPARAR + 1, Fg1.Rows - 1, Q_POS_MES_INICIO - 1, RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter:
            FORMATO_CELDA Fg1, Fg1.Rows - 1, Q_COL_COMPARAR_GRUPO + 1
            
            'Asignar posicion de inicio de grupo para totalizar los grupos
            M_FILA_INICIO_GRUPO = Fg1.Rows
            
        End If
    End If
salir:
    Set RST_TEPM_1 = Nothing
End Sub



Private Function CARGAR_DATOS_GRILLA_OTROS_ANYOS(RST_ORIGEN As ADODB.Recordset, _
                                         M_ANYO As String)
                                         
    '===================================================================================================
    'Creado : xxxxxx Por: Johan Castro
    'Propósito: Cargar datos a la grilla de otro año(distinto al año inicial)
    '
    'Entradas:  RST_ORIGEN= Recordset que contiene todos los datos de la consulta
    '           M_ANYO= Año que se esta consultando
    '
    'Resultados: Consulta en pantalla
    '
    '===================================================================================================

                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim Q_ROW1 As Long '--Posicion de la fila de la grilla
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    
    '--Incializar la barra de progreso
    PgBar(1).Min = 0
    PgBar(1).Max = Fg1.Rows
    
'    Fg1.Row = 2
    Dim Q_ROW As Long '--INDICA LA POSICION DEL REGISTRO A AGREGAR DATOS
    Dim N_FILTRO As String '--INDICA EL FILTRO QUE SE TENDRA QUE HACER AL RECORDSET
                            '-- DEPENDE DE Q_COL_FILA_OCULTA
                            
    For Q_ROW = 2 To Fg1.Rows - 1
        Fg1.Row = Q_ROW
        PgBar(1).Value = Q_ROW
        N_FILTRO = ""
        '--CONCATENO MI FILTRO
        If Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_grupo Then
            M_FILA_INICIO_GRUPO = Fg1.Row + 1
        ElseIf Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_Total Then
        
            CARGAR_DATOS_TOTALES False, "Total:", , , M_ANYO, True
            
        ElseIf Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_Total_grl Then
''            CARGAR_DATOS_TOTALES True, "Tot Gen:", True, False, M_ANYO, True

        ElseIf Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_en_Blanco Then
        
        Else
            '--Generar el filtro segun datos de la grilla
            For Q_ROW1 = 0 To Q_COL_FILA_OCULTA - 1
                N_FILTRO = N_FILTRO + RST_ORIGEN.Fields(Q_ROW1).Name + "= " + Fg1.TextMatrix(Fg1.Row, Q_ROW1 + 1) + " AND "
            Next Q_ROW1
            
            N_FILTRO = Left(N_FILTRO, Len(N_FILTRO) - 5) '--QUITO EL ULTIMO AND
            RST_ORIGEN.Filter = N_FILTRO '--HACER EL FILTRO
            If RST_ORIGEN.RecordCount > 0 Then
                'actualizar campo xsel a 1
                RST_ORIGEN("xsel") = 1
                DoEvents
                '--SI SE NTERRUMPE EL PROCESO => SALIR
                If BAND_INTERRUMPIR = True Then Exit Function
                
                '--CARGAR_DATOS A LA GRILLA
                CARGAR_DATOS_GRILLA RST_ORIGEN, M_ANYO, Q_ROW, True
            End If
            
        End If
    Next Q_ROW
    
    '--limpiar
    RST_ORIGEN.Filter = ""
    RST_ORIGEN.Filter = "xsel=0"
    
    '--Verificar si existe registros pendientes de mostrar
    If RST_ORIGEN.RecordCount <> 0 Then
        '--definir nuevo recordset para trabajar los nuevos registros que no se han insertado en grilla
        Dim xRstTmp As New ADODB.Recordset
        '--definir estructura
        DEFINIR_RST_TMP xRstTmp, RST_ORIGEN
        '--cargar data a nuevo recordset
        CARGAR_RST_TMP xRstTmp, RST_ORIGEN
                
        If Fg1.TextMatrix(Fg1.Rows - 1, Q_POSICION_TOTAL) = "Total:" Then
            '--Agregar fila en grid para separar el grupo anterior con el nuevo grupo
            ADD_REG Fg1, Fila_en_Blanco
            UNIR_CELDAS Fg1, Fg1.Rows - 1, IIf(Q_COL_FILA_OCULTA = -1, 1, Q_COL_FILA_OCULTA + 1), Fg1.Rows - 1, Fg1.Cols - 1, " ", flexAlignLeftCenter
        End If
        
        '--Cargar datos a la grilla
        CARGAR_DATOS xRstTmp, M_ANYO
        
        '--Limpiar recordset temporal
        Set xRstTmp = Nothing
        
    End If
    
End Function


Private Function CARGAR_DATOS_GRILLA(RST_ORIGEN As ADODB.Recordset, _
                                        M_ANYO As String, _
                                         Q_ROW As Long, _
                                         Optional F_OTROS_ANYOS As Boolean = False)
    '===================================================================================================
    'Creado : xxxxxx Por: Johan Castro
    'Propósito: Agregar datos a nueva fila en la grilla
    '
    'Entradas:  RST_ORIGEN=Recordset que contiene todos los datos de la consulta
    '           M_ANYO=Año que se esta consultando
    '           Q_ROW=Fila actual la grilla para ingresar datos
    '           F_OTROS_ANYOS=Indica si el año es distindo al año incial de la consulta(menor año)
    '                         False:Año inicial; False:Distinto al Año inicial
    '
    'Resultados: Nueva fila de la grilla con datos
    '
    '===================================================================================================
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    
    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO
    Dim Q_POS As Integer
    Dim qCampo As Integer
    Dim vStrCampo As String
    
    '--Identificar el incremento para las columnas de meses o trimestre
    For Q_POS = 0 To UBound(ARR_ANYO) - 1
        If ARR_ANYO(Q_POS) = M_ANYO Then
            Q_INCREMENTO_X_COL = Q_POS
            Exit For
        End If
    Next
    
    '--IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
    If opt_consulta(0).Value = True Then Q_INCREMENTO_X_COL = 0
    
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    '-----------
    
    DoEvents

    
    For qCampo = 0 To RST_ORIGEN.Fields.Count - 1
        If BAND_INTERRUMPIR = True Then Exit Function
        vStrCampo = RST_ORIGEN.Fields(qCampo).Name
        
        If LCase(vStrCampo) = "ene" Then
            Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
        End If
                   
        '--COLOCANDO LOS VALORES EN LA GRILLA
        Select Case LCase(vStrCampo)
            Case "xsel" '--no insertar datos, es campo comodin que se genera cuando hay registros pendientes de otros años
            
            '--DE LOS MESES
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"
                '"ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"
                '"ene-mar","abr-jun","jul-sep","oct-dic"
                '"1re sem","2do sem"
                
                '--ARR_TMP(0, 2) INDICA LA PRIMERA COLUMNA A MOSTRAR
                If LCase(vStrCampo) = ARR_TMP(0, 2) Then Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    
                Fg1.TextMatrix(Q_ROW, Q_POS_MES) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                
                If opt_consulta(0).Value = True Then
                    Q_POS_MES = Q_POS_MES + 1
                Else
                    Q_POS_MES = Q_POS_MES + Q_TOTAL_ANYO
                End If
                
             '--DEL TOTAL DEL AÑO
            Case "total"
                '--determinar la posicion de la columna del total
                If opt_consulta(0).Value = True Then
                    Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * 1
                Else
                    Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * Q_TOTAL_ANYO + Q_INCREMENTO_X_COL
                End If
                
                '--colocar el valor del total
                Fg1.TextMatrix(Q_ROW, Q_POS_MES_TOTAL) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                
                '--TOTALIZAR POR FILA
                '--TOTAL GRL
                If opt_consulta(0).Value = False Then
                    If IsNumeric(Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1)) = False Then
                        Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                    Else
                        Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = NulosN(Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1)) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                    End If
                    Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = Format(NulosN(Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1)), FORMAT_MONTO)
                End If
                
            '--DE LOS DEMAS CAMPOS
            Case Else
                '--SOLO SE AGREGARAN EN EL PRIMER AÑO
                If F_OTROS_ANYOS = False Then Fg1.TextMatrix(Q_ROW, qCampo + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
                
        End Select
        '------------
    Next
End Function



Private Sub pImprimir()

    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.Formularios
    MousePointer = vbHourglass
    
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", " ", T_RPT_PERIODO + " ", False, True

    Set X_PRINT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub

Private Sub ChkMostrarItem_Click()
    If ChkMostrarItem.Value = 0 Then
        Fg(3).Enabled = False
        
    Else
        '--habilitar controles de seleccion
        habilitar opt_totalizar, True
        
        '--LIMPIAR GRILLA
        Fg(3).Enabled = True
        LimpiarGrid Fg(3), True
        GRID_COMBOLIST Fg(3)
    End If
    '--restablecer tipo de consulta
    If opt_consulta(0).Value = True Then opt_consulta_Click 0
    If opt_consulta(1).Value = True Then opt_consulta_Click 1
    If opt_consulta(2).Value = True Then opt_consulta_Click 2
    If opt_consulta(3).Value = True Then opt_consulta_Click 3
    If opt_consulta(4).Value = True Then opt_consulta_Click 4
    If opt_consulta(5).Value = True Then opt_consulta_Click 5

End Sub



Private Sub CmdMas_Click(Index As Integer)
'--muestra u oculta mas opciones de reportes
If Index = 0 Then
    fr(4).Height = 1380
    CmdMas(0).Visible = False
    CmdMas(1).Visible = True
Else
    fr(4).Height = 660
    CmdMas(0).Visible = True
    CmdMas(1).Visible = False
    '--Activar importes
    If opt_totalizar(0).Value = False Then opt_totalizar(0).Value = True
    
End If

End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nOrden As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nSQLNotIn As String
    
    If Col <> 2 Then Exit Sub
    Select Case Index
    Case 0 '--CLIENTE
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "R.U.C.":   xCampos(1, 1) = "numruc":    xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
'            xCampos(2, 0) = "Id":       xCampos(2, 1) = "id":        xCampos(2, 2) = "800":   xCampos(2, 3) = "N"
            '--si hay filtros
            nSQLNotIn = GRID_GENERAR_SQL_ID(Fg(0), 1, " and mae_cliente.id", "NOT IN", True)
            If NulosC(Fg(Index).TextMatrix(Row, Col)) <> "" Then
                nSQLNotIn = nSQLNotIn & " and (UCASE(mae_cliente.nombre) LIKE '%" & UCase(NulosC(Fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(mae_cliente.nombre) LIKE '%" & UCase(NulosC(Fg(Index).TextMatrix(Row, Col))) & "%' ) "
            End If
            '--------------
            nSQL = "SELECT 0 as xsel, id, nombre,numruc FROM mae_cliente where id <> 0 " & nSQLNotIn & "  order by nombre asc"
            
            nTitulo = "Buscando Clientes"
            nOrden = "nombre"
            nCampoBusca = "nombre"
            
    Case 1 '--PTO VENTA
            ReDim xCampos(1, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4000":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Cliente":  xCampos(1, 1) = "cliente":   xCampos(1, 2) = "3200":   xCampos(1, 3) = "C"
'            xCampos(2, 0) = "Id":   xCampos(2, 1) = "id":        xCampos(2, 2) = "800":   xCampos(2, 3) = "N"
           '--si hay filtros
            nSQLNotIn = GRID_GENERAR_SQL_ID(Fg(1), 1, " WHERE vta_puntoVenta.id", "NOT IN", True)
            If NulosC(Fg(Index).TextMatrix(Row, Col)) <> "" Then
                nSQLNotIn = IIf(nSQLNotIn = "", " WHERE ", nSQLNotIn & " AND ") & "  (UCASE(vta_puntoVenta.descripcion) LIKE '%" & UCase(NulosC(Fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(vta_puntoVenta.descripcion) LIKE '%" & UCase(NulosC(Fg(Index).TextMatrix(Row, Col))) & "%' ) "
            End If
            '------
            nSQL = "SELECT 0 as xsel, vta_puntoVenta.id, vta_puntoVenta.descripcion AS nombre, mae_cliente.nombre as cliente " _
                + vbCr + " FROM vta_puntoVenta INNER JOIN mae_cliente ON vta_puntoVenta.idcli = mae_cliente.id " & nSQLNotIn _
                + vbCr + " ORDER BY mae_cliente.nombre, vta_puntoVenta.descripcion;"
            
            nTitulo = "Buscando Punto de Venta"
            nOrden = "nombre"
            nCampoBusca = "nombre"
    Case 2 '--VENDEDOR
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":   xCampos(1, 1) = "id":        xCampos(1, 2) = "800":   xCampos(1, 3) = "N"
            '--si hay filtros
            nSQLNotIn = GRID_GENERAR_SQL_ID(Fg(2), 1, " WHERE vta_vendedores.id", "NOT IN", True)
            If NulosC(Fg(Index).TextMatrix(Row, Col)) <> "" Then
                nSQLNotIn = IIf(nSQLNotIn = "", " WHERE ", nSQLNotIn & " AND ") & "  (UCASE(pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom) LIKE '%" & UCase(NulosC(Fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom) LIKE '%" & UCase(NulosC(Fg(Index).TextMatrix(Row, Col))) & "%' ) "
            End If
            '-------------
            nSQL = "SELECT 0 as xsel, vta_vendedores.id, pla_empleados.apepat & ' ' &  pla_empleados.apemat & ' ' & pla_empleados.nom AS nombre " _
                + vbCr + " FROM pla_empleados INNER JOIN vta_vendedores ON pla_empleados.id = vta_vendedores.idper " & nSQLNotIn _
                + vbCr + " ORDER BY pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom ;"
    
            nTitulo = "Buscando Vendedores"
            nOrden = "nombre"
            nCampoBusca = "nombre"
    
    Case 3 '--ITEM
        If TxtIdTipProd.Text = "" Then
            MsgBox "Falta especificar el tipo de item...!", vbExclamation, xTitulo
            TxtIdTipProd.SetFocus
            Exit Sub
        End If
        '---
        ReDim xCampos(2, 3) As String
        xCampos(0, 0) = "Descripción":   xCampos(0, 1) = "nombre":         xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Código":        xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2500":    xCampos(1, 3) = "C"
        'xCampos(2, 0) = "Id":            xCampos(2, 1) = "id":             xCampos(2, 2) = "800":         xCampos(2, 3) = "N"
        
        nSQLNotIn = GRID_GENERAR_SQL_ID(Fg(3), 1, " and alm_inventario.id", "NOT IN", True)
        If NulosC(Fg(Index).TextMatrix(Row, Col)) <> "" Then
            nSQLNotIn = " AND (UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg(Index).TextMatrix(Row, Col))) & "%' ) "
        End If
        '-------------
        nSQL = "SELECT 0 as xsel, id, codpro, descripcion as nombre FROM alm_inventario WHERE tippro = " & NulosN(TxtIdTipProd.Text) & nSQLNotIn & ""
        nTitulo = "Buscando Tipo de Item"
        nOrden = "nombre"
        nCampoBusca = "nombre"
    
    End Select
    Fg(Index).TextMatrix(Row, Col) = ""
    Dim xRs As New ADODB.Recordset
    
    If F_SELECCION = False Then
        '--PERMITIRA MOSTRAR LA VENTANA PARA AGREGAR UN REGISTRO
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, nOrden, nCampoBusca, Principio

    Else
        '--PERMITIRA MOSTRAR LA VENTANA PARA SELECCIONAR UNO O VARIOS REGISTROS
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
    End If

    If xRs.State = 0 Then GoTo salir
    If xRs.RecordCount = 0 Then GoTo salir
    
    '--SI ES SELECCION POSICIONAR EL RECORDSET A LA PRIMERA FILA
    If F_SELECCION = True Then xRs.MoveFirst
    
    Do While Not xRs.EOF
        Fg(Index).TextMatrix(Row, 1) = NulosN(xRs("id"))
        Fg(Index).TextMatrix(Row, 2) = NulosC(xRs("nombre"))
        
        If Fg(Index).Row = Fg(Index).Rows - 1 Then Fg(Index).AddItem ""
        Fg(Index).Row = Fg(Index).Rows - 1:
        Fg(Index).Col = 2
        
        '--VERIFICAR SI SOLAMENTE SE AGREGA UN REGISTRO
        If F_SELECCION = False Then
            Exit Do
        Else
            Row = Fg(Index).Rows - 1
        End If
        xRs.MoveNext
        
    Loop
        
    '--REINICIANDO VALOR A VARIABLE
    F_SELECCION = False
    
salir:
    Set xRs = Nothing

Exit Sub
error:
    
    Set xRs = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub

Private Sub Fg_DblClick(Index As Integer)
    Fg_CellButtonClick Index, Fg(Index).Rows - 1, 2
End Sub

Private Sub Fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Fg(Index).Row = -2 Then Exit Sub
    Select Case KeyCode
        Case 45  'INSERTAR REGI
            Fg(Index).AddItem ""
            Fg(Index).Row = Fg(Index).Rows - 1: Fg(Index).Col = 1
        Case 46 'SUPRIMIR/DELETE
            If Fg(Index).Rows - 1 >= 2 Then
                Fg(Index).RemoveItem Fg(Index).Row
            Else
                LimpiarGrid Fg(Index), True
                GRID_COMBOLIST Fg(Index)
            End If
    End Select
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        Select Case Index
            Case 0 '--cliente
                PopupMenu mn_cliente
            Case 1 '--punto de venta
                PopupMenu mn_ptoventa
            Case 2 '--vendedor
                PopupMenu mn_vendedor
            Case 3 '--item
                PopupMenu mn_item
        End Select
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        '--interrumpir
        BAND_INTERRUMPIR = True
    End If
End Sub

Private Sub Form_Load()
On Error GoTo error
    Dim k As Integer
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    '--CARGAR DATOS
    CentrarFrm Me
    '--FORMATO DE LAS GRILLAS
    For k = 0 To Fg.Count - 1
        GRID_COMBOLIST Fg(k)
        Fg(k).Tag = Fg(k).FormatString
    Next k
    Fg1.Tag = Fg1.FormatString
    
    LimpiarGrid Fg1
    '--CARGAR LOS AÑOS
    If CARGAR_LISTA_ANYOS_ACTIVOS(ls(0), xCon) = False Then Exit Sub
    Llenar_Mes ls(1)
    '--CARGANDO LOS MESES
    CARGAR_ARR_XX ARR_MES(), X_MES
    '--SELECCIONAR EL AÑO ACTUAL
    ls_activar_chek ls(0), AnoTra
    ls_activar_chek ls(1)
    '--CONFIGURAR LA GRILLA
    Validar_Consulta "-1"
    fGenerarConsulta "-1"
    Configurar_Grilla
    Exit Sub
error:
    SHOW_ERROR
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    If Me.Height > 3500 Then
        Fg1.Top = 2970
        Fg1.Width = Me.Width - 150
        Fg1.Height = Me.Height - 3350
    End If
End Sub

Private Sub mn_cliente1_Click()
    F_SELECCION = False
    Fg_CellButtonClick 0, Fg(0).Row, Fg(0).Col
End Sub

Private Sub mn_cliente2_Click()
    F_SELECCION = True
    Fg_CellButtonClick 0, Fg(0).Row, Fg(0).Col
End Sub

Private Sub mn_item1_Click()
    F_SELECCION = False
    Fg_CellButtonClick 3, Fg(3).Row, Fg(3).Col
End Sub

Private Sub mn_item2_Click()
    F_SELECCION = True
    Fg_CellButtonClick 3, Fg(3).Row, Fg(3).Col
End Sub

Private Sub mn_ptoventa1_Click()
    F_SELECCION = False
    Fg_CellButtonClick 1, Fg(1).Row, Fg(1).Col
End Sub

Private Sub mn_ptoventa2_Click()
    F_SELECCION = True
    Fg_CellButtonClick 1, Fg(1).Row, Fg(1).Col
End Sub

Private Sub mn_vendedor1_Click()
    F_SELECCION = False
    Fg_CellButtonClick 2, Fg(2).Row, Fg(2).Col
End Sub

Private Sub mn_vendedor2_Click()
    F_SELECCION = True
    Fg_CellButtonClick 2, Fg(2).Row, Fg(2).Col
End Sub

Private Sub opt_consulta_Click(Index As Integer)

    If ChkMostrarItem.Value = 1 Then
        Select Case Index
            Case 0, 5 '--x Año
                opt_totalizar(1).Enabled = False
                opt_totalizar(2).Enabled = False
                opt_totalizar(3).Enabled = False
                opt_totalizar(4).Enabled = False
            Case Else
                habilitar opt_totalizar, True
        End Select
        Exit Sub
    End If
    
    Select Case Index
        Case 0, 1, 3, 5 '--x Año
            opt_totalizar(1).Enabled = False
            opt_totalizar(2).Enabled = False
            opt_totalizar(3).Enabled = False
            opt_totalizar(4).Enabled = False
            '--verificar que no este activo los precios
            If opt_totalizar(0).Value = False Then opt_totalizar(0).Value = True
        Case 2, 4     '--x T. Prod/Item
            habilitar opt_totalizar, True
            
        Case Else   '--Otras tipos de Consulta
        
    End Select

    If opt_totalizar(0).Value = True Then opt_totalizar_Click 0
    If opt_totalizar(1).Value = True Then opt_totalizar_Click 1
    If opt_totalizar(2).Value = True Then opt_totalizar_Click 2
    If opt_totalizar(3).Value = True Then opt_totalizar_Click 3
    If opt_totalizar(4).Value = True Then opt_totalizar_Click 4

End Sub

Private Sub opt_estilo_Click(Index As Integer)
    Select Case Index
        Case 0 '--MES
            Llenar_Mes ls(1)
            ls_activar_chek ls(1)
            CARGAR_ARR_XX ARR_MES(), X_MES
        Case 1 '--TRIMESTRE
            Llenar_Trimestre ls(1)
            ls_activar_chek ls(1)
            CARGAR_ARR_XX ARR_MES(), X_TRIMESTRE
        Case 2 '--SEMESTRE
            Llenar_Semestre ls(1)
            ls_activar_chek ls(1)
            CARGAR_ARR_XX ARR_MES(), X_SEMESTRE
    End Select
    lbl(6).Caption = "Selecc. " + opt_estilo(Index).Caption
    
End Sub

Private Sub opt_totalizar_Click(Index As Integer)
    If Index = 1 Then '--Cantidades
        habilitar opt_mon, False
        habilitar opt_importe, False
        habilitar opt_escala, False: opt_escala(0).Value = True
        opt_mon(0).Value = False: opt_mon(1).Value = False: opt_mon(2).Value = False: opt_mon(3).Value = False
        opt_importe(0).Value = False: opt_importe(1).Value = False: opt_importe(2).Value = False
        
    Else '--Importes o Precios
        If TxtIdTipProd.Text = "" And ChkMostrarItem.Value = 0 And (opt_consulta(2).Value = False And opt_consulta(4).Value = False) Then
            habilitar opt_importe, True
        Else
            habilitar opt_importe, False
            opt_mon(2).Enabled = False: opt_mon(3).Enabled = False
        End If
        
        habilitar opt_escala, True
        habilitar opt_mon, True
        opt_mon(2).Value = True
        opt_importe(0).Value = True
    
    End If
End Sub



Private Sub TxtIdTipProd_Change()
    If TxtIdTipProd.Text = "" Then
        lblTipProducto.Caption = ""
        If ChkMostrarItem.Value = 1 Then ChkMostrarItem.Value = 0
        LimpiarGrid Fg(3), True
        ChkMostrarItem_Click
    End If
End Sub

Private Sub TxtIdTipProd_KeyPress(KeyAscii As Integer)
    On Error GoTo error
    If KeyAscii = 13 Then
        Dim RsTipProd As New ADODB.Recordset
        RsTipProd.CursorLocation = adUseClient
        If TxtIdTipProd.Text <> "" Then
            Set RsTipProd = BuscaConCriterio("SELECT id, descripcion FROM mae_tipoproducto WHERE id =" & Val(TxtIdTipProd.Text) & "", xCon)
            If RsTipProd.RecordCount <> 0 Then
                lblTipProducto.Caption = RsTipProd("descripcion")
            Else
                lblTipProducto.Caption = ""
                TxtIdTipProd.Text = ""
            End If
        End If
        ChkMostrarItem_Click
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
    Set RsTipProd = Nothing
    Exit Sub
error:
    Set RsTipProd = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"


End Sub

Private Sub TxtIdTipProd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then  'TECHAL F5
        CmdBusProducto.Value = True
    End If
End Sub

'------
Private Function Validar_Consulta(N_ANYO As String) As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    Dim k As Integer
    N_ANYO = ""
    Q_TOTAL_ANYO = 0
    '--RECORRER AÑO A AÑO PARA CARGAR LA DATA
    For k = ls(0).ListCount - 1 To 0 Step -1
        ls(0).ListIndex = k
        If ls(0).Selected(k) = True Then
            N_ANYO = N_ANYO + ls(0).Text + ","
            Q_TOTAL_ANYO = Q_TOTAL_ANYO + 1
        End If
    Next
    
    If N_ANYO = "" Then
       MsgBox "Seleccione un Año como mínimo", vbCritical, "Mensaje..."
'       ls(0).SetFocus
       Exit Function
    End If
    Erase ARR_ANYO '--LIMPIAR ARRAY
    ARR_ANYO = Split(N_ANYO, ",") '--ASIGNANDO EL LISTADO DE AÑOS
    
    
    '----------------
    Q_COL_ARR_TOTAL = 0
    For k = ls(1).ListCount - 1 To 0 Step -1
        ls(1).ListIndex = k
        If ls(1).Selected(k) = True Then
            Q_COL_ARR_TOTAL = Q_COL_ARR_TOTAL + 1
        End If
    Next
    If Q_COL_ARR_TOTAL = 0 Then
       MsgBox Replace(lbl(6).Caption, "Selecc.", "Selecc. un ") + " como mínimo...", vbCritical, xTitulo
       ls(1).SetFocus
       Exit Function
    End If
    Q_COL_ARR_TOTAL = Q_COL_ARR_TOTAL - 1
    '-----------
    Erase ARR_TMP
    ReDim ARR_TMP(Q_COL_ARR_TOTAL, 2)
    Dim POS_ARR As Integer
    POS_ARR = 0
    For k = 0 To ls(1).ListCount - 1
        ls(1).ListIndex = k
        If ls(1).Selected(k) = True Then
            ARR_TMP(POS_ARR, 0) = ARR_MES(k, 0)
            ARR_TMP(POS_ARR, 1) = ARR_MES(k, 1)
            ARR_TMP(POS_ARR, 2) = ARR_MES(k, 2)
            POS_ARR = POS_ARR + 1
        End If
    Next
    '-----------
    Validar_Consulta = True

End Function

Private Function fGenerarConsulta(M_ANYO As String) As String
    '===================================================================================================
    'Creado : xxxxxx Por: Johan Castro
    'Propósito: Asignar constantes a variables para configurar la presentacion de la grilla
    '           Generar sentencia SQL para mostrar datos segun opciones de la consulta
    '
    'Entradas:  M_ANYO=Año incial
    '
    'Resultados: Sentencia SQL para mostrar en pantalla
    '            Parametros para configurar cabecera de grilla
    '
    'Nota:  Se generan todos los tipos de consultas a presentar en pantalla
    '===================================================================================================

    '--
    Dim nSQLIdItem As String        '--Sentencia SQL para Item de almacen
    Dim nSQLIdCli As String         '--Sentencia SQL para Clientes
    Dim nSQLIdPtoVta As String      '--Sentencia SQL para Punto de Ventas
    Dim nSQLIdVen As String         '--Sentencia SQL para Vendedores
    Dim k As Integer                '--Recorrer arreglo de meses o trimestre para armar el Pivot
    '------------------------------------------------------------------------------------
    '--GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim nSQLValor As String         '--Indica que campo de SubConsulta va a sumar
    Dim nSQLCampos As String        '--Indica que campos de SubConsulta va mostrar
    Dim nSQLWhere As String         '--Indica las condiciones de la SubConsulta elegidas por el usuario
'    Dim nSQLFrom As String
    Dim nSQLGroupBy As String       '--Depende de nSQLCampos
    Dim nSQLOrderBy As String       '--Indica que campos va Ordenar de SubConsulta
    Dim nSQLPivot As String         '--Indica la presentacion de las columas por mes o trimestre
    Dim nSQLPivotSalida As String   '--Ordena los valores por mes o trimestre(Ene,Feb,Mar,Etc.)
    
    
    '***********************************************************************
    '***********************************************************************
    'Dim nSQL As String
    Dim nSQLTabla As String     '--Sentencia SQL para indicar una Sub consulta de donde se extraeran los datos
    Dim TipoTabla As Integer    '--Indica tipos de SubConsultas que se utilizara
                                '--1:Año, Tipo documento, Cliente, Vendedor
                                '--2:Producto
                                '--3:Punto de venta

    
'''''    Q_COL_FILA_OCULTA       '--OCULTAR COLUMNAS
'''''    Q_COL_FILA              '--CANTIDAD DE COLUMNAS QUE SE MOSTRARAN DESCONTANDO LOS MESES Y LOS TOTALES
'''''    Q_POSICION_TOTAL        '--POSICION DE LA COLUMNA QUE SE PONDRA EL TOTAL Y TOTAL_GRL EJ. TOTAL.(COL=2)   S/. 15000
'''''    Q_COL_COMPARAR_GRUPO    '--NO HAY GRUPO
    
    '------------------------------------------------------------------------------------------------
    
    '--DEL AÑO
    nSQLWhere = " Year(vta_ventas.fchdoc)= " + M_ANYO + " "
    '--
    '--DEL CLIENTE
    nSQLIdCli = GRID_GENERAR_SQL_ID(Fg(0), 1, " AND vta_ventas.idcli", "IN")
    
    '--DEL LOS PUNTOS DE VENTAS
    nSQLIdPtoVta = GRID_GENERAR_SQL_ID(Fg(1), 1, " AND vta_guia.idpunven", "IN")
    
    '--DEL LOS VENDEDORES
    nSQLIdVen = GRID_GENERAR_SQL_ID(Fg(2), 1, " AND vta_ventas.idven", "IN")
    
    '--DEL TIPO DE PRODUCTO
    If TxtIdTipProd.Text <> "" Then nSQLWhere = nSQLWhere + " AND alm_inventario.tippro = " + NulosC(TxtIdTipProd.Text) + " "
    
    '--DEL ITEM
    nSQLIdItem = GRID_GENERAR_SQL_ID(Fg(3), 1, " AND alm_inventario.id", "IN")
    
    '--CONCATENAR FECHA + CLIENTE + PUNTO DE VENTA + VENDEDOR + ITEM
    nSQLWhere = nSQLWhere + nSQLIdCli + nSQLIdPtoVta + nSQLIdVen + nSQLIdItem
    '---------------
    '--DE LA MONEDA
    '--SOLO MN
    If opt_mon(0).Value = True Then nSQLWhere = nSQLWhere + " AND vta_ventas.idmon= 1 "
    '--SOLO ME
    If opt_mon(1).Value = True Then nSQLWhere = nSQLWhere + " AND vta_ventas.idmon= 2 "
    '---------------
    
    '--restringir apertura de documento
    If chkAnioPasados.Value = 0 Then nSQLWhere = nSQLWhere + " AND vta_ventas.numreg<>'000001'  "
        
    '------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------
    If opt_estilo(0).Value = True Then '--MES
        nSQLPivot = "FORMAT(vta_ventas.fchdoc,'m') "
        
    ElseIf opt_estilo(1).Value = True Then '--TRIMESTRE
        nSQLPivot = "FORMAT(vta_ventas.fchdoc,'q') "
        
    ElseIf opt_estilo(2).Value = True Then '--SEMESTRE
        nSQLPivot = "FORMAT(vta_ventas.fchdoc,'s') "
        
    End If
    
    '--DEL PIVOT
    For k = 0 To UBound(ARR_TMP)
        nSQLPivotSalida = nSQLPivotSalida + ARR_TMP(k, 2) + ","
    Next k
    
    nSQLPivotSalida = " IN (" + Left(nSQLPivotSalida, Len(nSQLPivotSalida) - 1) + ") "
    nSQLWhere = nSQLWhere + " AND " + nSQLPivot + nSQLPivotSalida
    '--Otro formato del tipo del Povot Salida
    'nSQLPivotSalida = " In ('Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic');"
    
    '--reemplazar para la consulta general
    nSQLPivot = Replace(nSQLPivot, "vta_ventas", "vista")
    
    '--muestra datos segun fecha de registro
    If OptFecha(1).Value = True Then
        nSQLPivot = Replace(nSQLPivot, ".fchdoc", ".fchreg")
        nSQLWhere = Replace(nSQLWhere, ".fchdoc", ".fchreg")
    End If
    '------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------
        
    
    
    If opt_consulta(0).Value = True Then '--X AÑO
'''        If (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Then '--AÑO/PRODUCTO
'''            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 4:        Q_POSICION_TOTAL = 2:        Q_COL_COMPARAR_GRUPO = 0
'''            T_RPT_TITULO = "RESUMEN DE VENTAS POR AÑO CON TIPO PRODUCTO"
'''            nSQLCampos = "YEAR(vta_ventas.fchdoc) AS idanyo,alm_inventario.tippro,  YEAR(vta_ventas.fchdoc) AS anyo, mae_tipoproducto.descripcion "
'''            nSQLGroupBy = "alm_inventario.tippro,YEAR(vta_ventas.fchdoc),mae_tipoproducto.descripcion "
'''            nSQLOrderBy = "mae_tipoproducto.descripcion "
'''        ElseIf ChkMostrarItem.Value = 1 Then '--AÑO/PRODUCTO/ITEM
'''            Q_COL_FILA_OCULTA = 0:       Q_COL_FILA = 6:        Q_POSICION_TOTAL = 3:          Q_COL_COMPARAR_GRUPO = 0
'''            T_RPT_TITULO = "RESUMEN DE VENTAS POR AÑO CON ITEM"
'''            nSQLCampos = "YEAR(vta_ventas.fchdoc) AS idanyo,alm_inventario.tippro,alm_inventario.id,  YEAR(vta_ventas.fchdoc) AS anyo,mae_tipoproducto.descripcion,alm_inventario.descripcion AS desctipcom "
'''            nSQLGroupBy = "alm_inventario.tippro,alm_inventario.id,  YEAR(vta_ventas.fchdoc),mae_tipoproducto.descripcion,alm_inventario.descripcion"
'''            nSQLOrderBy = "mae_tipoproducto.descripcion,alm_inventario.descripcion  "
'''        Else    '--SOLO AÑOS
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 2:       Q_POSICION_TOTAL = 2:           Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR AÑO"
            
            nSQLCampos = "vista.idanyo, vista.anyo "
            nSQLGroupBy = "vista.idanyo ,vista.anyo "
            nSQLOrderBy = "vista.anyo "
            
            TipoTabla = 1
            
            '--cambiar tipo tabla cuando muestra cantidades
            If opt_totalizar(1).Value = True Then TipoTabla = 2

            '--si consultan por tipo producto
            If NulosN(TxtIdTipProd.Text) <> 0 Then TipoTabla = 2

'''        End If
        
        '--cambiar tipo tabla cuando seleccione filtro por pto venta
            If nSQLIdPtoVta <> "" Then TipoTabla = 3
        
        
    ElseIf opt_consulta(1).Value = True Then '--X CLIENTE
        If (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Then '--CLIETNE/PRODUCTO
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 4:        Q_POSICION_TOTAL = 4:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR PROVEEDOR CON TIPO PRODUCTO"
            
            nSQLCampos = "vista.idcli,vista.tippro,  vista.nomcliente,vista.desctipcom "
            nSQLGroupBy = "vista.idcli,vista.tippro,  vista.nomcliente,vista.desctipcom "
            nSQLOrderBy = "vista.nomcliente,vista.desctipcom "
            
            nSQLWhere = nSQLWhere + " AND alm_inventario.tippro IS NOT NULL " '--SOLO LOS QUE TIENEN TIPO DE PRODUCTO
            
            TipoTabla = 2
            
        ElseIf ChkMostrarItem.Value = 1 Then '--CLIENTE/PRODUCTO/ITEM
            Q_COL_FILA_OCULTA = 3:        Q_COL_FILA = 7:        Q_POSICION_TOTAL = 7:        Q_COL_COMPARAR_GRUPO = 3
            T_RPT_TITULO = "RESUMEN DE VENTAS POR CLIENTE CON ITEM"
            nSQLCampos = "vista.idcli,vista.tippro,vista.iditem,  vista.nomcliente,vista.desctipcom,vista.codigo,vista.descitem "
            nSQLGroupBy = "vista.idcli,vista.tippro,vista.iditem,  vista.nomcliente,vista.desctipcom,vista.codigo,vista.descitem "
            nSQLOrderBy = "vista.nomcliente,vista.desctipcom,vista.descitem "
            
            nSQLWhere = nSQLWhere + " AND alm_inventario.tippro IS NOT NULL " '--SOLO LOS QUE TIENEN TIPO DE PRODUCTO
            
            TipoTabla = 2
            
        Else    '--SOLO CLIENTE
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 3:        Q_POSICION_TOTAL = 3:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR CLIENTE"
            nSQLCampos = " vista.idcli,vista.numruc, vista.nomcliente "
            nSQLGroupBy = "vta_ventas.idcli,vista.numruc,vista.nomcliente "
            nSQLOrderBy = "vista.nomcliente "
            
            TipoTabla = 1
            
            '--cambiar tipo tabla cuando muestra cantidades
            If opt_totalizar(1).Value = True Then TipoTabla = 2
            
        End If
    
        '--cambiar tipo tabla cuando seleccione filtro por pto venta
        If nSQLIdPtoVta <> "" Then TipoTabla = 3
    
    ElseIf opt_consulta(2).Value = True Then '--X PTO DE VENTA
        If (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Then '--X PTO DE VENTA/PRODUCTO
            Q_COL_FILA_OCULTA = 3:        Q_COL_FILA = 6:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 3
            T_RPT_TITULO = "RESUMEN DE VENTAS POR PUNTO DE VENTA CON TIPO PRODUCTO"
            nSQLCampos = "vista.idcli,vista.idpunven,vista.tippro,  vista.nomcliente,vista.descptovta,vista.desctipcom "
            nSQLGroupBy = "vista.idcli,vista.idpunven,vista.tippro,  vista.nomcliente,vista.descptovta,vista.desctipcom "
            nSQLOrderBy = "vista.nomcliente,vista.descptovta"
            
            nSQLWhere = nSQLWhere + " AND vta_guia.idpunven <>0 " '--SOLO LOS QUE TIENEN PUNTO DE VENTA
            
            TipoTabla = 3
            
        ElseIf ChkMostrarItem.Value = 1 Then '--X PTO DE VENTA/PRODUCTO/ITEM
            Q_COL_FILA_OCULTA = 4:        Q_COL_FILA = 9:        Q_POSICION_TOTAL = 9:        Q_COL_COMPARAR_GRUPO = 4
            T_RPT_TITULO = "RESUMEN DE VENTAS POR PUNTO DE VENTA CON ITEM"
            
            nSQLCampos = " vista.idcli,vista.idpunven,vista.tippro,vista.iditem,  vista.nomcliente,vista.descptovta,vista.desctipcom,vista.codigo,vista.descitem "
            nSQLGroupBy = " vista.idcli,vista.idpunven,vista.tippro,vista.iditem,  vista.nomcliente,vista.descptovta,vista.desctipcom,vista.codigo,vista.descitem "
            nSQLOrderBy = " vista.nomcliente,vista.descptovta"
            
            nSQLWhere = nSQLWhere + " AND vta_guia.idpunven <>0 " '--SOLO LOS QUE TIENEN PUNTO DE VENTA
            
            TipoTabla = 3
                        
        Else    '--X PTO DE VENTA
            Q_COL_FILA_OCULTA = 2:        Q_COL_FILA = 4:        Q_POSICION_TOTAL = 4:        Q_COL_COMPARAR_GRUPO = 2
            T_RPT_TITULO = "RESUMEN DE VENTAS POR PUNTO DE VENTA"
            
            nSQLCampos = "vista.idcli,vista.idpunven,  vista.nomcliente,vista.descptovta"
            nSQLGroupBy = "vista.idcli,vista.idpunven,  vista.nomcliente,vista.descptovta"
            nSQLOrderBy = "vista.nomcliente,vista.descptovta"
            
            nSQLWhere = nSQLWhere + " AND vta_guia.idpunven <>0 " '--SOLO LOS QUE TIENEN PUNTO DE VENTA
            
            TipoTabla = 3
            
        End If
        
        '--muestra datos segun fecha de registro(solo si selecciona por fch documento)
        If OptFecha(0).Value = True Then nSQLPivot = Replace(nSQLPivot, "vista.fchdoc", "vista.fecgiro")
        
        nSQLWhere = nSQLWhere + " AND (vta_guia.iddocven <>0 OR vta_guia.iddocven  IS NOT NULL) "
       
    ElseIf opt_consulta(3).Value = True Then '--X VENDEDOR
        If (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Then '--VENDEDOR/PRODUCTO
            Q_COL_FILA_OCULTA = 2:        Q_COL_FILA = 4:        Q_POSICION_TOTAL = 4:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR VENDEDOR CON TIPO PRODUCTO"
            nSQLCampos = "vista.idven,vista.tippro,  vista.vendedor,vista.desctipcom "
            nSQLGroupBy = "vista.idven,vista.tippro,  vista.vendedor,vista.desctipcom "
            nSQLOrderBy = "vista.vendedor"
            
            nSQLWhere = nSQLWhere + " AND vta_ventas.idven <> 0 " '--SOLO LOS QUE TIENEN VENDEDORES
            
            TipoTabla = 2
            
        ElseIf ChkMostrarItem.Value = 1 Then '--VENDEDOR/PRODUCTO/ITEM
            Q_COL_FILA_OCULTA = 3:        Q_COL_FILA = 7:        Q_POSICION_TOTAL = 7:        Q_COL_COMPARAR_GRUPO = 3
            T_RPT_TITULO = "RESUMEN DE VENTAS POR VENDEDOR CON ITEM"
            nSQLCampos = "vista.idven,vista.tippro,vista.iditem,  vista.vendedor,vista.desctipcom,vista.codigo,vista.descitem "
            nSQLGroupBy = "vista.idven,vista.tippro,vista.iditem,  vista.vendedor,vista.desctipcom,vista.codigo,vista.descitem "
            nSQLOrderBy = "vista.vendedor "
            nSQLWhere = nSQLWhere + " AND vta_ventas.idven <> 0 " '--SOLO LOS QUE TIENEN VENDEDORES
                        
            TipoTabla = 2
            
        Else    '--SOLO VENDEDOR
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 2:        Q_POSICION_TOTAL = 2:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR VENDEDOR"
            nSQLCampos = "vista.idven,  vista.vendedor "
            nSQLGroupBy = "vista.idven,  vista.vendedor "
            nSQLOrderBy = "vista.vendedor "
            
            nSQLWhere = nSQLWhere + " AND vta_ventas.idven <> 0 " '--SOLO LOS QUE TIENEN VENDEDORES
            
            TipoTabla = 1
            
            '--cambiar tipo tabla cuando muestra cantidades
            If opt_totalizar(1).Value = True Then TipoTabla = 2

        End If
    
        '--cambiar tipo tabla cuando seleccione filtro por pto venta
        If nSQLIdPtoVta <> "" Then TipoTabla = 3
            
    ElseIf opt_consulta(4).Value = True Then '--X PRODUCTO
        If TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0 Then
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 2:        Q_POSICION_TOTAL = 2:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR TIPO PRODUCTO"
            nSQLCampos = "vista.tippro,  vista.desctipcom "
            nSQLGroupBy = "vista.tippro,  vista.desctipcom "
            nSQLOrderBy = "vista.desctipcom "
            
            nSQLWhere = nSQLWhere + " AND alm_inventario.tippro IS NOT NULL " '--SOLO LOS QUE TIENEN TIPO DE PRODUCTO
            
            TipoTabla = 2
            
        ElseIf ChkMostrarItem.Value = 1 Then
            Q_COL_FILA_OCULTA = 2:        Q_COL_FILA = 5:        Q_POSICION_TOTAL = 5:        Q_COL_COMPARAR_GRUPO = 2
            T_RPT_TITULO = "RESUMEN DE VENTAS POR TIPO PRODUCTO CON ITEM"
            nSQLCampos = "vista.tippro,vista.iditem,  vista.desctipcom,vista.codigo,vista.descitem "
            nSQLGroupBy = "vista.tippro,vista.iditem,  vista.desctipcom,vista.codigo,vista.descitem "
            nSQLOrderBy = "vista.desctipcom,vista.descitem "
            
            nSQLWhere = nSQLWhere + " AND alm_inventario.tippro IS NOT NULL " '--SOLO LOS QUE TIENEN TIPO DE PRODUCTO
            
            TipoTabla = 2
        
        Else '--X FAMILIA
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 2:        Q_POSICION_TOTAL = 2:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR FAMILIA"
            nSQLCampos = "vista.idfam,  vista.descfam "
            nSQLGroupBy = "vista.idfam,  vista.descfam "
            nSQLOrderBy = "vista.descfam "
            
            nSQLWhere = nSQLWhere + " AND alm_inventario.idfam IS NOT NULL " '--SOLO LOS QUE TIENEN TIPO DE PRODUCTO
            
            TipoTabla = 2
            
        End If
        
        '--cambiar tipo tabla cuando seleccione filtro por pto venta
        If nSQLIdPtoVta <> "" Then TipoTabla = 3
        
    ElseIf opt_consulta(5).Value = True Then '--X TIPO DE DOCUMENTO
        If (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Then '--VENDEDOR/PRODUCTO
            Q_COL_FILA_OCULTA = 2:        Q_COL_FILA = 4:        Q_POSICION_TOTAL = 4:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR TIPO DE DOCUMENTO CON TIPO PRODUCTO"
            nSQLCampos = "vista.tipdoc,vista.tippro,  vista.tdocabrev,vista.desctipcom "
            nSQLGroupBy = "vista.tipdoc,vista.tippro,  vista.tdocabrev,vista.desctipcom "
            nSQLOrderBy = "vista.tdocabrev"
                       
            TipoTabla = 2
            
        ElseIf ChkMostrarItem.Value = 1 Then '--TIPO DE DOCUM
            Q_COL_FILA_OCULTA = 3:        Q_COL_FILA = 7:        Q_POSICION_TOTAL = 7:        Q_COL_COMPARAR_GRUPO = 3
            T_RPT_TITULO = "RESUMEN DE VENTAS POR TIPO DE DOCUMENTO CON ITEM"
            nSQLCampos = "vista.tipdoc,vista.tippro,vista.iditem,  vista.tdocabrev,vista.desctipcom,vista.codigo,vista.descitem "
            nSQLGroupBy = "vista.tipdoc,vista.tippro,vista.iditem,  vista.tdocabrev,vista.desctipcom,vista.codigo,vista.descitem "
            nSQLOrderBy = "vista.tdocabrev "
                        
            TipoTabla = 2
            
        Else    '--SOLO TIPO DE DOCUMENTO
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 2:        Q_POSICION_TOTAL = 2:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR TIPO DE DOCUMENTO"
            nSQLCampos = "vista.tipdoc,  vista.tdocabrev "
            nSQLGroupBy = "vista.tipdoc,  vista.tdocabrev "
            nSQLOrderBy = "vista.tdocabrev "
            
            TipoTabla = 1
            
        End If
   
        '--cambiar tipo tabla cuando seleccione filtro por pto venta
        If nSQLIdPtoVta <> "" Then TipoTabla = 3
        
    End If
    
    
    '**************************************************************************************************
    '**************************************************************************************************
    '--definir las consultas de donde se obtendran los datos
    If TipoTabla = 1 Then
        '*año
        '*tipo documento
        '*cliente
        '*vendedor
        '--consulta ventas resumen
        nSQLTabla = "( SELECT YEAR(vta_ventas.fchreg) AS anyo, Left([vta_ventas].[numreg],2) & Format(mae_libros.codsun,'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre as nomcliente, mae_documento.abrev AS tdocabrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc,vta_ventas.fchreg,vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, " _
                + vbCr + " IIf([vta_ventas].[impsal]<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, " _
                + vbCr + " IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] is null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
                + vbCr + " IIf(vta_ventas.tipdoc=7,(-1)*vta_ventas.imptotdoc,vta_ventas.imptotdoc) AS impreal, " _
                + vbCr + " IIf(vta_ventas.idmon=1,impreal,0) AS imptotmn, " _
                + vbCr + " IIf(vta_ventas.idmon=2,impreal,0) AS imptotme, " _
                + vbCr + " (imptotmn + imptotme * tipcam) as imptotexpmn, " _
                + vbCr + " (imptotme + iif(tipcam=0,0, imptotmn / tipcam) )  as imptotexpme, " _
                + vbCr + " IIf(vta_ventas.tipdoc=7,(-1)*vta_ventas.impigv,vta_ventas.impigv) AS impigvreal, " _
                + vbCr + " IIf(vta_ventas.idmon=1,impigvreal,0) AS impigvmn, " _
                + vbCr + " IIf(vta_ventas.idmon=2,impigvreal,0) AS impigvme, " _
                + vbCr + " (impigvmn + impigvme * tipcam) as impigvexpmn, " _
                + vbCr + " (impigvme + iif(tipcam=0,0, impigvmn / tipcam) )  as impigvexpme, " _
                + vbCr + " pla_empleados.numdoc AS numdni, pla_empleados.nombre AS vendedor, " _
                + vbCr + " YEAR(vta_ventas.fchreg) AS idanyo,vta_ventas.idcli , vta_ventas.tipdoc, vta_ventas.idmon,vta_ventas.idven " _
                + vbCr + " FROM ((mae_cliente RIGHT JOIN (((((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) ON mae_cliente.id = vta_ventas.idcli) " _
                + vbCr + " LEFT JOIN vta_vendedores ON vta_ventas.idven = vta_vendedores.id) LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id " _
                + vbCr + " WHERE  " & nSQLWhere _
                + vbCr + " ) AS vista "
    

    ElseIf TipoTabla = 2 Then
        '*producto
        '--ventas y detalle de ventas
        '--en esta consulta solo se considera la base imponible
                    
        nSQLTabla = " ( SELECT YEAR(vta_ventas.fchreg) AS anyo, Left([vta_ventas].[numreg],2) & Format(mae_libros.codsun,'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre AS nomcliente, mae_documento.abrev AS tdocabrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchreg,vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, " _
                + vbCr + " IIf([vta_ventas].[impsal]<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
                + vbCr + " mae_tipoproducto.descripcion AS desctipcom,mae_familia.descripcion AS descfam, alm_inventario.codpro as codigo ,alm_inventario.descripcion as descitem, mae_unidades.abrev AS prodabrev, IIf([vta_ventas].[tipdoc]=7,(-1) * vta_ventasdet.canpro,vta_ventasdet.canpro) as canpro, " _
                + vbCr + " IIf([vta_ventas].[tipdoc]=7,(-1)*[vta_ventasdet].[imptot],[vta_ventasdet].[imptot]) AS impreal, " _
                + vbCr + " IIf([vta_ventas].[idmon]=1,[impreal],0) AS imptotmn, " _
                + vbCr + " IIf([vta_ventas].[idmon]=2,[impreal],0) AS imptotme, " _
                + vbCr + " IIf([vta_ventas].[idmon]=1,[vta_ventasdet].[preuni],0) AS pumn, " _
                + vbCr + " IIf([vta_ventas].[idmon]=2,[vta_ventasdet].[preuni],0) AS pume, " _
                + vbCr + " IIf([vta_ventas].[idmon]=1,[imptotmn],[impreal]*[tipcam]) AS imptotexpmn, " _
                + vbCr + " IIf([vta_ventas].[idmon]=2,[imptotme],IIf([tipcam]=0,0,[impreal]/[tipcam])) As imptotexpme, " _
                + vbCr + " IIf([vta_ventas].[idmon]=1,[pumn],[pume]*[tipcam]) AS puexpmn, " _
                + vbCr + " IIf([vta_ventas].[idmon]=2,[pume],IIF( tipcam=0,0,[pumn]/[tipcam])) AS puexpme, " _
                + vbCr + " pla_empleados.numdoc AS numdni, pla_empleados.nombre AS vendedor, " _
                + vbCr + " YEAR(vta_ventas.fchreg) AS idanyo,vta_ventas.idcli , vta_ventas.tipdoc, vta_ventas.idmon ,vta_ventas.idven, alm_inventario.id AS iditem, alm_inventario.tippro,alm_inventario.idfam " _
                + vbCr + " FROM ((((mae_cliente RIGHT JOIN (((((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) ON mae_cliente.id = vta_ventas.idcli) INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta) " _
                + vbCr + " LEFT JOIN (mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventasdet.iditem = alm_inventario.id) LEFT JOIN (vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id) ON vta_ventas.idven = vta_vendedores.id) LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id " _
                + vbCr + " WHERE " & nSQLWhere _
                + vbCr + "  ) AS vista "
    
    ElseIf TipoTabla = 3 Then
    '*pto venta

        nSQLTabla = " ( SELECT YEAR(vta_ventas.fchreg) AS anyo, Left([vta_ventas].[numreg],2) & Format(mae_libros.codsun,'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre AS nomcliente, mae_documento.abrev AS tdocabrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchreg,vta_guia.fecgiro, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, " _
                + vbCr + " IIf([vta_ventas].[impsal]<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
                + vbCr + " mae_tipoproducto.descripcion AS desctipcom,mae_familia.descripcion AS descfam, alm_inventario.codpro as codigo ,alm_inventario.descripcion as descitem, mae_unidades.abrev AS prodabrev, vta_ventasdet.canpro, " _
                + vbCr + " IIf([vta_ventas].[tipdoc]=7,(-1)*vta_guiadet.canpro*vta_ventasdet.preuni,vta_guiadet.canpro * vta_ventasdet.preuni) AS impreal, " _
                + vbCr + " IIf([vta_ventas].[idmon]=1,[impreal],0) AS imptotmn, " _
                + vbCr + " IIf([vta_ventas].[idmon]=2,[impreal],0) AS imptotme, " _
                + vbCr + " IIf([vta_ventas].[idmon]=1,[vta_ventasdet].[preuni],0) AS pumn, " _
                + vbCr + " IIf([vta_ventas].[idmon]=2,[vta_ventasdet].[preuni],0) AS pume, " _
                + vbCr + " IIf([vta_ventas].[idmon]=1,[imptotmn],[impreal]*[tipcam]) AS imptotexpmn, " _
                + vbCr + " IIf([vta_ventas].[idmon]=2,[imptotme],IIf([tipcam]=0,0,[impreal]/[tipcam])) As imptotexpme, " _
                + vbCr + " IIf([vta_ventas].[idmon]=1,[pumn],[pume]*[tipcam]) AS puexpmn, " _
                + vbCr + " IIf([vta_ventas].[idmon]=2,[pume],IIF( tipcam=0,0,[pumn]/[tipcam])) AS puexpme, " _
                + vbCr + " pla_empleados.numdoc AS numdni, pla_empleados.nombre AS vendedor,vta_puntoVenta.descripcion AS descptovta, " _
                + vbCr + " YEAR(vta_ventas.fchreg) AS idanyo,vta_ventas.idcli , vta_ventas.tipdoc, vta_ventas.idmon ,vta_ventas.idven, alm_inventario.id AS iditem, alm_inventario.tippro,alm_inventario.idfam,vta_guia.idpunven " _
                + vbCr + " FROM ((vta_guiadet INNER JOIN vta_guia ON vta_guiadet.idgui = vta_guia.id) LEFT JOIN vta_puntoVenta ON vta_guia.idpunven = vta_puntoVenta.id) INNER JOIN (((((mae_cliente RIGHT JOIN (((((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) ON mae_cliente.id = vta_ventas.idcli)  " _
                + vbCr + " INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN (mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventasdet.iditem = alm_inventario.id) LEFT JOIN (vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id) ON vta_ventas.idven = vta_vendedores.id) LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) ON (vta_guia.iddocven = vta_ventas.id) AND (vta_guiadet.iditem = vta_ventasdet.iditem) " _
                + vbCr + " WHERE " & nSQLWhere _
                + vbCr + "  ) AS vista "

    End If

    '**************************************************************************************************
    '**************************************************************************************************
    
    
    '--DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA '--Q_COL_FILA + CAMPO_TOTAL
    
    
    
    '-----------------------------------------
    '-muestra importe
    If opt_totalizar(0).Value = True Then
        '--importe total
        If opt_importe(0).Value = True Then
            '--solo mn o solo me
            If opt_mon(0).Value = True Or opt_mon(1).Value = True Then nSQLValor = " Sum(vista.impreal) "
            '--expresado en mn
            If opt_mon(2).Value = True Then nSQLValor = " Sum(vista.imptotexpmn) "
            '--expresado en me
            If opt_mon(3).Value = True Then nSQLValor = " Sum(vista.imptotexpme) "
            
        '--base afecta e inafecta
        ElseIf opt_importe(1).Value = True Then
            '--solo mn o solo me
            If opt_mon(0).Value = True Or opt_mon(1).Value = True Then nSQLValor = " (Sum(vista.impreal) - Sum(vista.impigvreal)) "
            '--expresado en mn
            If opt_mon(2).Value = True Then nSQLValor = " (Sum(vista.imptotexpmn) - Sum(vista.impigvexpmn))"
            '--expresado en me
            If opt_mon(3).Value = True Then nSQLValor = " (Sum(vista.imptotexpme) - Sum(vista.impigvexpme))"
        
        '--igv
        Else
            '--solo mn o solo me
            If opt_mon(0).Value = True Or opt_mon(1).Value = True Then nSQLValor = " Sum(vista.impigvreal) "
            '--expresado en mn
            If opt_mon(2).Value = True Then nSQLValor = " Sum(vista.impigvexpmn)"
            '--expresado en me
            If opt_mon(3).Value = True Then nSQLValor = " Sum(vista.impigvexpme)"
                        
        End If
    
    '--muestra cantidades
    ElseIf opt_totalizar(1).Value = True Then
        nSQLValor = "SUM(vista.canpro) "
    '--muestra Precio minimo
    ElseIf opt_totalizar(2).Value = True Then
        If opt_mon(0).Value = True Then nSQLValor = " Min(vista.pumn) " '--solo en mn
        If opt_mon(1).Value = True Then nSQLValor = " Min(vista.pume) " '--solo en me
        If opt_mon(2).Value = True Then nSQLValor = " Min(vista.puexpmn) " '--expresado en mn
        If opt_mon(3).Value = True Then nSQLValor = " Min(vista.puexpme) " '--expresado en me
        
    '--muestra Precio promedio
    ElseIf opt_totalizar(3).Value = True Then
        If opt_mon(0).Value = True Then nSQLValor = " Avg(vista.pumn) " '--solo en mn
        If opt_mon(1).Value = True Then nSQLValor = " Avg(vista.pume) " '--solo en me
        If opt_mon(2).Value = True Then nSQLValor = " Avg(vista.puexpmn) " '--expresado en mn
        If opt_mon(3).Value = True Then nSQLValor = " Avg(vista.puexpme) " '--expresado en me

    '--muestra Precio maximo
    ElseIf opt_totalizar(4).Value = True Then
        If opt_mon(0).Value = True Then nSQLValor = " Max(vista.pumn) " '--solo en mn
        If opt_mon(1).Value = True Then nSQLValor = " Max(vista.pume) " '--solo en me
        If opt_mon(2).Value = True Then nSQLValor = " Max(vista.puexpmn) " '--expresado en mn
        If opt_mon(3).Value = True Then nSQLValor = " Max(vista.puexpme) " '--expresado en me
    End If
    
    '--expresar los valores a miles
    If opt_escala(1).Value Then nSQLValor = nSQLValor & "/1000 "
    
    '--generar la consulta
    fGenerarConsulta = " TRANSFORM " + nSQLValor + _
        vbCr + " SELECT " + nSQLCampos + "," + nSQLValor + " AS total " + _
        vbCr + " FROM " + nSQLTabla + _
        vbCr + " GROUP BY " + nSQLGroupBy + _
        vbCr + " ORDER BY " + nSQLOrderBy + _
        vbCr + " PIVOT " + nSQLPivot + nSQLPivotSalida
    

    '------------------------------------------------------------------------------------
    
        
End Function


Private Sub CARGAR_DATOS_TOTALES(BAND_ADD_TOTAL As Boolean, Nombre_total As String, _
                Optional Band_Total_gral As Boolean = False, _
                Optional band_forzar_suma As Boolean = False, _
                Optional M_ANYO As String, _
                Optional F_OTROS_ANYOS As Boolean = False)
                
    '===================================================================================================
    'Creado : xxxxxx Por: Johan Castro
    'Propósito: Agregar nueva fila o actualizar fila para totalizar grupos o todos los grupos
    '
    'Entradas:  BAND_ADD_TOTAL=Indica si agregara fila para totalizar
    '           Nombre_Total=Nombre del total
    '           band_forzar_suma=Indica si sumara todas las filas sin considerar los acumulados en arreglo
    '                            False:No forzar; True:Forzar suma
    '           M_ANYO=Indica año de consulta
    '           F_OTROS_ANYOS=Indica si el año es distinto al año incial(Menor Año)
    '                         False:Año actual; True:Otros años
    'Resultados: Nueva fila o fila actualizada con la suma de los grupos
    '
    '===================================================================================================
                
                
    Dim Q_MES As Integer
    Dim X_ROW As Long
    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO
    
    'On Error Resume Next
    If F_OTROS_ANYOS = False Then
        X_ROW = Fg1.Rows
        If BAND_ADD_TOTAL = True Then
            '--Agregando fila
            ADD_REG Fg1, IIf(Band_Total_gral = False, Fila_Total, Fila_Total_grl)
            
            'PONIENDO LOS NOMBRES DE LOS TOTALES  Q_POSICION_TOTAL SE OBTIENE DE fGenerarConsulta()
            Fg1.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
            FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL
        End If
    Else
        X_ROW = Fg1.Row
    End If
  
    '
'--------------------------
    
    For Q_MES = 0 To UBound(ARR_ANYO) - 1
        If ARR_ANYO(Q_MES) = M_ANYO Then
            Q_INCREMENTO_X_COL = Q_MES
            Exit For
        End If
    Next
    
    '--Identifica posicion incial de mes o trimestre
    If opt_consulta(0).Value = True Then Q_INCREMENTO_X_COL = 0
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    
    '--Identificar posicion incial de arreglo
    Q_POS_ARRAY = 0 + Q_INCREMENTO_X_COL
    
    If Band_Total_gral = False Then
        '--DE LOS MESES
        For Q_MES = 0 To Q_COL_ARR_TOTAL
            '--INTERRUMPIR EL PROCESO
            If BAND_INTERRUMPIR = True Then Exit Sub
                        
            Fg1.TextMatrix(X_ROW, Q_POS_MES) = Format(GRID_SUMAR_COL(Fg1, Q_POS_MES, M_FILA_INICIO_GRUPO, X_ROW - 1), FORMAT_MONTO)
                
            '--acumulando los subtotales
            ArrayTotalGral(Q_POS_ARRAY) = ArrayTotalGral(Q_POS_ARRAY) + NulosN(Fg1.TextMatrix(X_ROW, Q_POS_MES))
            
            FORMATO_CELDA Fg1, X_ROW, Q_POS_MES
            If opt_consulta(0).Value = True Then
                Q_POS_MES = Q_POS_MES + 1
                Q_POS_ARRAY = Q_POS_ARRAY + 1
            Else
                Q_POS_MES = Q_POS_MES + Q_TOTAL_ANYO
                Q_POS_ARRAY = Q_POS_ARRAY + Q_TOTAL_ANYO
            End If
            
        Next Q_MES
        
        '--Para los totales y el total general
        For Q_MES = Q_COL_ARR_TOTAL + 1 To Q_COL_ARR_TOTAL + 2
            If Q_MES = Q_COL_ARR_TOTAL + 1 Then '--TOTAL
                If opt_consulta(0).Value = True Then
                    Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * 1
                    Q_POS_ARRAY = (Q_COL_ARR_TOTAL + 1) * 1 + 1
                Else
                    Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * Q_TOTAL_ANYO + Q_INCREMENTO_X_COL

                    Q_POS_ARRAY = (Q_COL_ARR_TOTAL + 1) * Q_TOTAL_ANYO + Q_INCREMENTO_X_COL
                End If
                
            ElseIf Q_MES = Q_COL_ARR_TOTAL + 2 Then '--TOTAL GRAL
                Q_POS_MES_TOTAL = Fg1.Cols - 1
                Q_POS_ARRAY = (Q_COL_ARR_TOTAL + 2) * Q_TOTAL_ANYO
                '--restar del acumulado el subtotal, para volver a acumularlo con el nuevo subtotal
                ArrayTotalGral(Q_POS_ARRAY) = ArrayTotalGral(Q_POS_ARRAY) - NulosN(Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL))
            
            End If
            
            Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = Format(GRID_SUMAR_COL(Fg1, Q_POS_MES_TOTAL, M_FILA_INICIO_GRUPO, X_ROW - 1), FORMAT_MONTO)
            
            '--acumulando los subtotales
            ArrayTotalGral(Q_POS_ARRAY) = ArrayTotalGral(Q_POS_ARRAY) + NulosN(Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL))
            
            FORMATO_CELDA Fg1, X_ROW, Q_POS_MES_TOTAL

        Next Q_MES
    
    Else
        If band_forzar_suma = False Then
            '--colocando el total general
            For Q_MES = Q_POS_MES_INICIO To Fg1.Cols - 1
                Fg1.TextMatrix(X_ROW, Q_MES) = Format(ArrayTotalGral(Q_MES - Q_POS_MES_INICIO), FORMAT_MONTO)
                FORMATO_CELDA Fg1, X_ROW, Q_MES
            Next Q_MES
            
        Else
            For Q_MES = Q_POS_MES_INICIO To Fg1.Cols - 1
                Fg1.TextMatrix(X_ROW, Q_MES) = Format(GRID_SUMAR_COL(Fg1, Q_MES, Fg1.FixedRows, Fg1.Rows - 2), FORMAT_MONTO)
                FORMATO_CELDA Fg1, X_ROW, Q_MES
            Next Q_MES
                
        End If
    
    End If
    
'    ADD_REG fr1, Fila_en_Blanco
    Err.Clear
    
End Sub


Private Sub Configurar_Grilla()
    '===================================================================================================
    'Creado : xxxxxx Por: Johan Castro
    'Propósito: Configurar la Cabecera de la Grilla segun la consulta
    '
    'Entradas:  Ninguno
    '
    'Resultados: Grilla con formato, listo para insertar datos
    '
    'Nota:  Se configura todas las presentaciones posibles
    '===================================================================================================
    
    Dim M_ANCHO_COL_MES As Integer '--DEPENDERA DEL TIPO DE PRESENTACION
                                    '--EN DECIMALES, EN MILES
    Dim k&, j&
       
    Fg1.Clear
    Fg1.FrozenCols = 0
    
    
    If opt_estilo(0).Value = True Then '--MES
        M_ANCHO_COL_MES = 800
    ElseIf opt_estilo(1).Value = True Then '--TRIMESTRE
        M_ANCHO_COL_MES = 900
    ElseIf opt_estilo(2).Value = True Then '--SEMESTRE
        M_ANCHO_COL_MES = 1000
    End If
    
    If opt_escala(0).Value = True Then
        M_ANCHO_COL_MES = M_ANCHO_COL_MES + 250
    Else
        M_ANCHO_COL_MES = M_ANCHO_COL_MES
    End If
    
    With Fg1
        '-----
        '--DATOS DE FILA
        If opt_consulta(0).Value = True Then
        
            Fg1.Cols = Q_COL_FILA + (Q_COL_ARR_TOTAL + 1) + 1
            UNIR_CELDAS Fg1, 0, Q_COL_FILA, 0, Fg1.Cols - 1, " ", flexAlignCenterTop
            '--DATOS DE FILA
            .ColAlignment(2) = flexAlignLeftCenter
            .TextMatrix(1, 2) = "Año":         .ColWidth(2) = M_ANCHO_COL_MES
                        
            Q_POS_MES = Q_POS_MES_INICIO
            '--DATOS DE COLUMNAS
            For k = 0 To Q_COL_ARR_TOTAL '--MESES DEL AÑO
                '--COLOCANDO LOS MESES
                UNIR_CELDAS Fg1, 1, Q_POS_MES, 1, Q_POS_MES, ARR_TMP(k, 1), flexAlignCenterTop: .ColWidth(k) = M_ANCHO_COL_MES
                .ColAlignment(Q_POS_MES) = flexAlignRightBottom
                .Row = 0:   .Col = Q_POS_MES:   .CellAlignment = flexAlignCenterBottom
                Q_POS_MES = Q_POS_MES + 1
            Next k
            '--COLOCANDO EL TOTAL
            .TextMatrix(1, .Cols - 1) = "Total Gral":         .ColWidth(.Cols - 1) = M_ANCHO_COL_MES + 200
        Else
        
            '--CANTIDAD DE COLUMNAS
            Fg1.Cols = Q_COL_FILA + ((Q_COL_ARR_TOTAL + 2) * Q_TOTAL_ANYO) + 1
                                    '--total_mes+total_años
            '---
            If opt_consulta(1).Value = True Then '--X CLIENTE
                If (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Then '--CLIETNE/PRODUCTO
                    .TextMatrix(1, 3) = "Cliente":       .ColWidth(3) = 1500:   .ColAlignment(3) = flexAlignLeftCenter
                    .TextMatrix(1, 4) = "Producto":      .ColWidth(4) = 1000:   .ColAlignment(4) = flexAlignLeftCenter
                    .Row = 1:   .Col = 4:  .CellAlignment = flexAlignLeftCenter
                ElseIf ChkMostrarItem.Value = 1 Then '--CLIETNE/PRODUCTO/ITEM
                    .TextMatrix(1, 4) = "Cliente":       .ColWidth(4) = 1500:   .ColAlignment(4) = flexAlignLeftCenter
                    .TextMatrix(1, 5) = "Producto":      .ColWidth(5) = 1000:   .ColAlignment(5) = flexAlignLeftCenter
                    .TextMatrix(1, 6) = "Código":        .ColWidth(6) = 900:   .ColAlignment(6) = flexAlignLeftCenter
                    .TextMatrix(1, 7) = "Item":          .ColWidth(7) = 2000:   .ColAlignment(7) = flexAlignLeftCenter
                    .Row = 1:   .Col = 5:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 6:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 7:  .CellAlignment = flexAlignLeftCenter
                Else    '--SOLO CLIENTE
                    .TextMatrix(1, 2) = "Ruc":          .ColWidth(2) = 1100:   .ColAlignment(2) = flexAlignLeftCenter
                    .TextMatrix(1, 3) = "Cliente":       .ColWidth(3) = 2500:   .ColAlignment(3) = flexAlignLeftCenter
                    .Row = 1:   .Col = 2:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 3:  .CellAlignment = flexAlignLeftCenter
                End If
                
            ElseIf opt_consulta(2).Value = True Then '--X PTO DE VENTA
                If (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Then '--CLIENTE/PTO DE VENTA/PRODUCTO
                    .TextMatrix(1, 4) = "Cliente":        .ColWidth(4) = 1500:   .ColAlignment(4) = flexAlignLeftCenter
                    .TextMatrix(1, 5) = "Punto de Venta": .ColWidth(5) = 1800:   .ColAlignment(5) = flexAlignLeftCenter
                    .TextMatrix(1, 6) = "Producto":       .ColWidth(6) = 1000:   .ColAlignment(6) = flexAlignLeftCenter
                    .Row = 1:   .Col = 5:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 6:  .CellAlignment = flexAlignLeftCenter
                ElseIf ChkMostrarItem.Value = 1 Then '--CLIENTE/PTO DE VENTA/PRODUCTO/ITEM
                    .TextMatrix(1, 5) = "Cliente":        .ColWidth(5) = 1500:   .ColAlignment(5) = flexAlignLeftCenter
                    .TextMatrix(1, 6) = "Punto de Venta": .ColWidth(6) = 1800:   .ColAlignment(6) = flexAlignLeftCenter
                    .TextMatrix(1, 7) = "Producto":       .ColWidth(7) = 1000:   .ColAlignment(7) = flexAlignLeftCenter
                    .TextMatrix(1, 8) = "Código":         .ColWidth(8) = 900:    .ColAlignment(8) = flexAlignLeftCenter
                    .TextMatrix(1, 9) = "Item":           .ColWidth(9) = 2000:   .ColAlignment(9) = flexAlignLeftCenter
                    .Row = 1:   .Col = 6:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 7:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 8:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 9:  .CellAlignment = flexAlignLeftCenter
                Else    '--SOLO PTO DE VENTA
                .TextMatrix(1, 3) = "Cliente":            .ColWidth(3) = 2000:   .ColAlignment(3) = flexAlignLeftCenter
                .TextMatrix(1, 4) = "Punto de Venta":     .ColWidth(4) = 2500:   .ColAlignment(4) = flexAlignLeftCenter
                .Row = 1:   .Col = 3:  .CellAlignment = flexAlignLeftCenter
                .Row = 1:   .Col = 4:  .CellAlignment = flexAlignLeftCenter
                End If
                
            ElseIf opt_consulta(3).Value = True Then '--X VENDEDOR
                If (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0) Then '--VENDEDOR/PRODUCTO
                    .TextMatrix(1, 3) = "Vendedor":       .ColWidth(3) = 1500:   .ColAlignment(3) = flexAlignLeftCenter
                    .TextMatrix(1, 4) = "Producto":       .ColWidth(4) = 1000:   .ColAlignment(4) = flexAlignLeftCenter
                    .Row = 1:   .Col = 4:  .CellAlignment = flexAlignLeftCenter
                ElseIf ChkMostrarItem.Value = 1 Then '--VENDEDOR/PRODUCTO/ITEM
'                    .TextMatrix(1, 2) = "Vendedor":       .ColWidth(4) = 1500:   .ColAlignment(4) = flexAlignLeftCenter
                    .TextMatrix(1, 5) = "Producto":       .ColWidth(5) = 1000:   .ColAlignment(5) = flexAlignLeftCenter
                    .TextMatrix(1, 6) = "Código":         .ColWidth(6) = 900:    .ColAlignment(6) = flexAlignLeftCenter
                    .TextMatrix(1, 7) = "Item":           .ColWidth(7) = 2000:   .ColAlignment(7) = flexAlignLeftCenter
                    .Row = 1:   .Col = 5:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 6:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 7:  .CellAlignment = flexAlignLeftCenter
                Else    '--SOLO VENDEDOR
                    .TextMatrix(1, 2) = "Vendedor":       .ColWidth(2) = 2000:   .ColAlignment(2) = flexAlignLeftCenter
                    .Row = 1:   .Col = 2:  .CellAlignment = flexAlignLeftCenter
                End If
                
            ElseIf opt_consulta(4).Value = True Then '--X PRODUCTO / ITEM
                If TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0 Then
                    .TextMatrix(1, 2) = "Producto":      .ColWidth(2) = 1000:   .ColAlignment(2) = flexAlignLeftCenter
                    .Row = 1:   .Col = 2:  .CellAlignment = flexAlignLeftCenter
                ElseIf ChkMostrarItem.Value = 1 Then
                    .TextMatrix(1, 4) = "Código":      .ColWidth(4) = 1000:   .ColAlignment(4) = flexAlignLeftCenter
                    .TextMatrix(1, 5) = "Item":        .ColWidth(5) = 1800:    .ColAlignment(5) = flexAlignLeftCenter
'                    .TextMatrix(1, 7) = "Item":          .ColWidth(7) = 2000:   .ColAlignment(7) = flexAlignLeftCenter
                    .Row = 1:   .Col = 7:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 5:  .CellAlignment = flexAlignLeftCenter
                    '.Row = 1:   .Col = 7:  .CellAlignment = flexAlignLeftCenter
                Else
                    .TextMatrix(1, 2) = "Familia":       .ColWidth(2) = 2000:   .ColAlignment(2) = flexAlignLeftCenter
                    .Row = 1:   .Col = 2:  .CellAlignment = flexAlignLeftCenter
                End If
            ElseIf opt_consulta(5).Value = True Then '--X TIPO DOCUMENTO
                If TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0 Then
                    .TextMatrix(1, 3) = "T.D.":         .ColWidth(3) = 600:   .ColAlignment(3) = flexAlignLeftCenter
                    .TextMatrix(1, 4) = "Producto":     .ColWidth(4) = 1000:   .ColAlignment(4) = flexAlignLeftCenter
                    .Row = 1:   .Col = 3:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 4:  .CellAlignment = flexAlignLeftCenter
                ElseIf ChkMostrarItem.Value = 1 Then
                    .TextMatrix(1, 5) = "Producto":      .ColWidth(5) = 1000:   .ColAlignment(5) = flexAlignLeftCenter
                    .TextMatrix(1, 6) = "Código":        .ColWidth(6) = 900:    .ColAlignment(6) = flexAlignLeftCenter
                    .TextMatrix(1, 7) = "Item":          .ColWidth(7) = 2000:   .ColAlignment(7) = flexAlignLeftCenter
                    .Row = 1:   .Col = 5:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 6:  .CellAlignment = flexAlignLeftCenter
                    .Row = 1:   .Col = 7:  .CellAlignment = flexAlignLeftCenter
                Else
                    .TextMatrix(1, 2) = "T.D.":       .ColWidth(2) = 600:   .ColAlignment(2) = flexAlignLeftCenter
                    .Row = 1:   .Col = 2:  .CellAlignment = flexAlignLeftCenter
                End If
            End If
            Q_POS_MES = Q_POS_MES_INICIO
            
            '--DATOS DE COLUMNAS
            For k = 0 To Q_COL_ARR_TOTAL + 1 '--MESES DEL AÑO + TOTAL
                '--COLOCANDO LOS MESES Y AGRUPANDOLOS
                If k = Q_COL_ARR_TOTAL + 1 Then
                    UNIR_CELDAS Fg1, 0, Q_POS_MES, 0, Q_POS_MES + Q_TOTAL_ANYO - 1, "Totales", flexAlignRightCenter
                Else
                    UNIR_CELDAS Fg1, 0, Q_POS_MES, 0, Q_POS_MES + Q_TOTAL_ANYO - 1, IIf(Q_TOTAL_ANYO > 1, ARR_TMP(k, 0), ARR_TMP(k, 1)), flexAlignRightCenter
                End If
                
                .ColAlignment(Q_POS_MES) = flexAlignRightCenter
                
                .Row = 0:   .Col = Q_POS_MES:   .CellAlignment = flexAlignCenterCenter
    
                '--COLOCANDO LOS AÑOS
                For j = 0 To Q_TOTAL_ANYO - 1 '--CANTIDAD DE AÑOS SELECCIONADOS
                    If k = Q_COL_ARR_TOTAL + 1 Then
                        .ColWidth(Q_POS_MES + j) = M_ANCHO_COL_MES + 200 '--DE LOS AÑOS
                    Else
                        .ColWidth(Q_POS_MES + j) = M_ANCHO_COL_MES  '--DE LOS MESES
                    End If
                    UNIR_CELDAS Fg1, 1, Q_POS_MES + j, 1, Q_POS_MES + j, ARR_ANYO(j), flexAlignCenterCenter
                    .Row = 1:   .Col = Q_POS_MES + j:   .CellAlignment = flexAlignCenterCenter
                Next j
                
                Q_POS_MES = Q_POS_MES + Q_TOTAL_ANYO
                
            Next k
            
            '--COLOCANDO LOS TOTALES
            .ColWidth(.Cols - 1) = M_ANCHO_COL_MES + 400
            UNIR_CELDAS Fg1, 0, .Cols - 1, 0, .Cols - 1, "Total Gral.", flexAlignCenterTop
            
            'DEL TOTAL GRAL
            UNIR_CELDAS Fg1, 1, .Cols - 1, 1, .Cols - 1, "Total", flexAlignCenterTop
           .ColAlignment(.Cols - 1) = flexAlignRightCenter
            '--OCULTAR EL GRUPO
            .ColWidth(Q_COL_COMPARAR_GRUPO + 1) = 0
           
        End If
        .FrozenCols = Q_POS_MES_INICIO - 1
        .ColWidth(0) = 200
        '--Generar ID's a los campos ID'S
        For k = 1 To Q_COL_FILA_OCULTA
            .TextMatrix(1, k) = "ID" + CStr(k):         .ColWidth(k) = 500
        Next
        
        '--Ocultar columnas
        If Q_COL_FILA_OCULTA <> -1 Then OCULTAR_COL Fg1, 1, Q_COL_FILA_OCULTA

    End With
    DoEvents
    
End Sub



Sub PosicionarProgBar()
'--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
'    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    FraProgreso.Visible = True
End Sub


'---DEL GRAFICO
'--251007


Private Sub CmdGrafAcep1_Click()
    If OptTipGrafBarra1.Value = True Then
        vLngTipoGrafico = 51
    ElseIf OptTipGrafLinea.Value = True Then
        vLngTipoGrafico = 65
    ElseIf OptTipGrafCircular.Value = True Then
        vLngTipoGrafico = 5
    End If
    
    If OptConDatoResum1.Value = True Then
        vTipoDato = 0
    ElseIf OptconDatosDetalle1.Value = True Then
        vTipoDato = 1
    End If
    
    If ChkLeyenda.Value = 1 Then
        vViewLeyenda = True
    Else
        vViewLeyenda = False
    End If
    
    GrafEstilo_TotGral_0_1
    FraGraf1.Visible = False
End Sub

Private Sub CmdGrafCancel1_Click()
    FraGraf1.Visible = False
End Sub


Private Function fTituloGrafico() As String
    If OptConDatoResum1.Value = True Then
        fTituloGrafico = "RESUMIDO POR AÑO"
    ElseIf OptconDatosDetalle1.Value = True Then
        fTituloGrafico = "DETALLADO POR AÑO"
    End If
End Function

Private Sub GenerarGraf_TotGral_0_1(pRango As String, pTipoGraf As Long, pTitulo As String, pTipoDato As Integer)
    With Oleapp
        '--MACRO 1
    '    .Sheets("Hoja1").Select
    '    .Sheets("Hoja1").Name = "dato"
    '    .Range(pRango).Select
        .Charts.Add
        '.ActiveChart.ChartType = xlColumnClustered
        .ActiveChart.ChartType = pTipoGraf
        '.ActiveChart.SetSourceData Source:=Sheets("dato").Range("A3:B5"), PlotBy:=xlColumns
        If OptTipGrafLinea.Value = True Then
            .ActiveChart.SetSourceData Source:=.Sheets("datos").Range(pRango), PlotBy:=1
        Else
            If OptconDatosDetalle1.Value = True Then
                .ActiveChart.SetSourceData Source:=.Sheets("datos").Range(pRango), PlotBy:=1
            Else
                .ActiveChart.SetSourceData Source:=.Sheets("datos").Range(pRango), PlotBy:=2
            End If
        End If
        '.ActiveChart.Location Where:=xlLocationAsNewSheet
        .ActiveChart.Location Where:=1
'        If pTipoDato = 1 Then
'            ActiveChart.HasLegend = True
'        End If
        '----
        Select Case pTipoGraf
            Case 51 'BARRAS
                If pTipoDato = 0 Then
                    .ActiveChart.ChartArea.Select
                    .ActiveChart.ApplyDataLabels Type:=2, LegendKey:=False
                End If
            Case 5 'CIRCULAR
                .ActiveChart.HasLegend = True
                .ActiveChart.Legend.Select
                .Selection.Position = -4152
                .ActiveChart.ApplyDataLabels Type:=3, LegendKey:=True _
                    , HasLeaderLines:=True
        End Select
        '-----
        '--PONER TITULO
        .ActiveChart.ChartArea.Select
        With .ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = pTitulo
        End With
        On Error Resume Next
        .ActiveChart.ChartArea.Select
        .ActiveChart.HasLegend = vViewLeyenda
        
    End With
End Sub

Private Sub GrafEstilo_TotGral_0_1()
    'GRAFICO POR ANIO POR TOTAL GENERAL
    Dim i_row As Long, i_col As Long, fs As Variant, NFILA As Long
    Dim nArchivo As String, NCOLUMN As Long, vRangSelect As String
    Dim vColTotMesAnio As Long, vIniCol_Grilla As Integer, vColIndexVarible As Long
    'VARIABLES PARA TRABAJAR CON LA SELECCION DE CELDAS DE EXCEL
    Dim vRango1 As String, vRango2 As String, vRangoCelSelecTotal As String
    '----------------------------------------------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    If vTipoDato = 0 Then
        nArchivo = "C:\grafico_x_anio.XLS"
    Else
        nArchivo = "C:\grafico_x_anio_Detallado.XLS"
    End If
    Set Oleapp = CreateObject("excel.application")
    Oleapp.Visible = True
    With Oleapp
        .WindowState = 1
        .Workbooks.Add
        .Sheets(1).Select
        .Sheets(1).Name = "datos"
                
        NCOLUMN = 2 'COLUMNA INICIO PARA EXCEL
        vIniCol_Grilla = 3
        
        If vTipoDato = 0 Then
            vColTotMesAnio = vCantMeses + 1 'MESES + 1(TOTAL GENERAL)
        Else
            vColTotMesAnio = vCantMeses
        End If
        '--LE SUMA EL VALOR DE INICIO DE LA COLUMNA DE INICIO
        vColTotMesAnio = vColTotMesAnio + vIniCol_Grilla - 1
        '--PONEL EL ENCABEZADO DEL TOTAL GRAL. O DE LOS MESES
        If vTipoDato = 0 Then 'SOLO CON TOTAL GENERAL
            .Cells(3, NCOLUMN) = Fg1.TextMatrix(1, vColTotMesAnio)
        Else 'PARA DETALLADO
            For i_col = vIniCol_Grilla To vColTotMesAnio
                .Cells(3, NCOLUMN) = Fg1.TextMatrix(1, i_col)
                NCOLUMN = NCOLUMN + 1
            Next
        End If
        
        '--PONE LOS ANIO COMO REGISTROS EN LA COLUMNA 1 EN EXCEL
        NFILA = 4
        NCOLUMN = 2 'COLUMNA INICIO PARA EXCEL
        For i_row = 2 To Fg1.Rows - 2
            .Cells(NFILA, 1) = Trim(Fg1.TextMatrix(i_row, 2))
            NFILA = NFILA + 1
        Next
        
        'LLENAR LOS DATOS DEL DETALLE DE LA GRILLA
        If vTipoDato = 0 Then 'SOLO POR TOT GENERAL
            vColTotMesAnio = vCantMeses + 1 'MESES + 1(TOTAL GENERAL)
        Else 'SOLO PARA DETALLADO
            vColTotMesAnio = vCantMeses
        End If
        '--LE SUMA EL VALOR DE INICIO DE LA COLUMNA DE INICIO
        vColTotMesAnio = vColTotMesAnio + vIniCol_Grilla - 1
        NCOLUMN = 2 'COLUMNA INICIO PARA EXCEL
        NFILA = 4
        For i_row = 2 To Fg1.Rows - 2
            If vTipoDato = 0 Then
                .Cells(NFILA, NCOLUMN) = Fg1.TextMatrix(i_row, vColTotMesAnio)
                NFILA = NFILA + 1
            ElseIf vTipoDato = 1 Then
                NCOLUMN = 2
                For i_col = vIniCol_Grilla To vColTotMesAnio
                    .Cells(NFILA, NCOLUMN) = Fg1.TextMatrix(i_row, i_col)
                    NCOLUMN = NCOLUMN + 1
                Next
                NFILA = NFILA + 1
            End If
        Next
        '--GENERA EL GRAFICO
        vRango1 = .Cells(3, 1).Address
        If vTipoDato = 0 Then
            vRango2 = .Cells(NFILA - 1, 2).Address
        Else
            vColTotMesAnio = vCantMeses + 1
            vRango2 = .Cells(NFILA - 1, vColTotMesAnio).Address
        End If
        vRangSelect = vRango1 & ":" & vRango2
        
        vTituloGraf = fTituloGrafico
        'vLngTipoGrafico = 51 barras
'        vLngTipoGrafico = 5 'pie
        GenerarGraf_TotGral_0_1 vRangSelect, vLngTipoGrafico, vTituloGraf, vTipoDato
'        Oleapp.ActiveWorkbook.SaveAs (nArchivo)
        Oleapp.WindowState = 1
        '.ActiveWindow.Zoom = 75
    End With
'    vRangSelect = "A" & CStr(3) & ":M" & CStr(NFILA)
'    GeneraGrafico vRangSelect, "Grafico por Año"
'    Oleapp.Quit
    Set Oleapp = Nothing   ' la aplicación; después libera la referenci
    Set fs = Nothing
    MsgBox "Los datos han sido exportados correctamente", vbInformation, "Aviso"
End Sub

Private Sub UnirCeldaEnExcel(pRango As String)
    With Oleapp
        .Range(pRango).Select
        With .Selection
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
            .WrapText = False
            .Orientation = 0
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Selection.Merge
    End With
End Sub

Sub GraficoEstilo_0() 'SOLO POR ANIO
    Dim i_row As Long, i_col As Long, fs As Variant, NFILA As Long
    Dim nArchivo As String, NCOLUMN As Long, vRangSelect As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    nArchivo = "C:\XANIO_MESES_GRAFIC.XLS"
    
    Set Oleapp = CreateObject("excel.application")
    With Oleapp
        .Workbooks.Add
        .Sheets(1).Select
        .Sheets(1).Name = "datos"
               
'        .CELLS(1, 3).Value = "PADRÓN DE ARTICULOS"
        NCOLUMN = 2
        For i_col = 2 To Fg1.Cols - 2
            .Cells(3, NCOLUMN) = Fg1.TextMatrix(1, i_col)
            NCOLUMN = NCOLUMN + 1
        Next
               
        NFILA = 4: NCOLUMN = 1
        For i_row = 2 To Fg1.Rows - 2
            NCOLUMN = 1
            For i_col = 1 To Fg1.Cols - 2
                .Cells(NFILA, NCOLUMN) = Fg1.TextMatrix(i_row, i_col)
                NCOLUMN = NCOLUMN + 1
            Next
            NFILA = NFILA + 1
        Next
        NFILA = 3
        For i_row = 2 To Fg1.Rows - 2
            NFILA = NFILA + 1
        Next
        Oleapp.ActiveWorkbook.SaveAs (nArchivo)
    End With
    vRangSelect = "A" & CStr(3) & ":M" & CStr(NFILA)
    GeneraGrafico vRangSelect, "Grafico por Año"
    Oleapp.Quit
    Set Oleapp = Nothing   ' la aplicación; después libera la referenci
    Set fs = Nothing
    MsgBox "Los datos han sido exportados correctamente", vbInformation, "Aviso"
End Sub

Sub GraficoEstilo_1() 'SOLO POR PROVEEDOR
    Dim i_row As Long, i_col As Long, fs As Variant, NFILA As Long
    Dim nArchivo As String, NCOLUMN As Long, vRangSelect As String
    Dim vColTotMesAnio As Long, vIniCol_Grilla As Integer, vColIndexVarible As Long
    'VARIABLES PARA TRABAJAR CON LA SELECCION DE CELDAS DE EXCEL
    Dim vRango1 As String, vRango2 As String, vRangoCelSelecTotal As String
    '----------------------------------------------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    nArchivo = "C:\grafico_x_proveedor.XLS"
    Set Oleapp = CreateObject("excel.application")
    With Oleapp
        .Workbooks.Add
        .Sheets(1).Select
        .Sheets(1).Name = "datos"

        NCOLUMN = 2 'COLUMNA INICIO PARA EXCEL
        vIniCol_Grilla = 3
        vColTotMesAnio = (vCantMeses * Q_TOTAL_ANYO) - 1 'MESES X ANIOS
        '--PONE EL ENCABEZADO DE ANIOS
        For i_col = vIniCol_Grilla To vColTotMesAnio + vIniCol_Grilla
            .Cells(3, NCOLUMN) = Fg1.TextMatrix(0, i_col)
            .Cells(4, NCOLUMN) = Fg1.TextMatrix(1, i_col)
            NCOLUMN = NCOLUMN + 1
        Next
        '--UNE CELDAS DE LOS MESES
        'ESTA VARIABLE vCantMeses ME INDICA LA CANTIDAD DE MESES SELECCIONADOS
        vColTotMesAnio = (vCantMeses * Q_TOTAL_ANYO)
        vColIndexVarible = 2
        For i_col = 1 To vCantMeses
            vRango1 = .Cells(3, vColIndexVarible).Address
            vRango2 = .Cells(3, vColIndexVarible + (Q_TOTAL_ANYO - 1)).Address
            vRangoCelSelecTotal = vRango1 & ":" & vRango2 'ejemplo B3:C3
            On Error Resume Next
            UnirCeldaEnExcel vRangoCelSelecTotal
            vColIndexVarible = vColIndexVarible + Q_TOTAL_ANYO
        Next
        'LLENAR NOMBRES DE PROVEEDORES
        NFILA = 5
        For i_row = 2 To Fg1.Rows - 1
            .Cells(NFILA, 1) = Trim(Fg1.TextMatrix(i_row, 2))
            NFILA = NFILA + 1
        Next
        'LLENAR LOS DATOS DEL DETALLE DE LA GRILLA
        vColTotMesAnio = (vCantMeses * Q_TOTAL_ANYO) - 1
        NFILA = 5: NCOLUMN = 2
        For i_row = 2 To Fg1.Rows - 1
            NCOLUMN = 2
            For i_col = 3 To (3 + vColTotMesAnio)
                
                .Cells(NFILA, NCOLUMN) = Fg1.TextMatrix(i_row, i_col)
                NCOLUMN = NCOLUMN + 1
            Next
            NFILA = NFILA + 1
        Next
        Oleapp.ActiveWorkbook.SaveAs (nArchivo)
    End With
'    vRangSelect = "A" & CStr(3) & ":M" & CStr(NFILA)
'    GeneraGrafico vRangSelect, "Grafico por Año"
    Oleapp.Quit
    Set Oleapp = Nothing   ' la aplicación; después libera la referenci
    Set fs = Nothing
    MsgBox "Los datos han sido exportados correctamente", vbInformation, "Aviso"
End Sub

Sub GeneraGrafico(pRango As String, pTitGrafico As String)
    With Oleapp
        .Charts.Add
        .ActiveChart.ChartType = 65
        .ActiveChart.SetSourceData Source:=.Sheets("datos").Range(pRango), PlotBy:=1
        .ActiveChart.Location Where:=1
        .ActiveChart.ChartArea.Select
        .Selection.AutoScaleFont = True
        With .Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = -4142
            .ColorIndex = -4105
            .Background = -4105
        End With
        '-----------
'        .ActiveChart.ChartArea.Select
'        .ActiveChart.ApplyDataLabels AutoText:=True, LegendKey:=False, _
'        HasLeaderLines:=False, ShowSeriesName:=False, ShowCategoryName:=False, _
'        ShowValue:=True, ShowPercentage:=False, ShowBubbleSize:=False, Separator _
'        :=" "
        
        'PARA OFF 97
        .ActiveChart.ChartArea.Select
        .ActiveChart.ApplyDataLabels Type:=2, LegendKey:=False
        '------------
        With .ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = pTitGrafico
        End With
    End With
End Sub
'--FIN CODIGO DE GRAFICO------------------------------------------


Private Sub EXPORTAR()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.Formularios

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO, T_RPT_PERIODO, "", "Ventas"
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub



'************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then CONSULTAR
    If Button.Index = 5 Then EXPORTAR
    If Button.Index = 6 Then pVerGrafico
    If Button.Index = 7 Then pImprimir
    If Button.Index = 9 Then
        BAND_INTERRUMPIR = True
        Unload Me
    End If
End Sub

'************************************************

Private Sub pVerGrafico()
    If Fg1.Rows = 2 Then
        MsgBox "No hay datos para el gráfico.", vbInformation, xTitulo
        Exit Sub
    End If
    
    Dim vEstilo As Integer
'''    vEstilo = ESTILO_CONSULTA
    vCantMeses = Q_COL_ARR_TOTAL + 1
    FraGraf1.Left = (Me.Width - FraGraf1.Width) \ 2
    FraGraf1.Top = (Me.Height - FraGraf1.Height) \ 2
    FraGraf1.Visible = True

End Sub


Private Sub FraGraf1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    FraGraf1.ZOrder 0
End Sub

Private Sub FraGraf1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With FraGraf1
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub


Private Sub FraProgreso_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    FraProgreso.ZOrder 0
End Sub

Private Sub FraProgreso_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With FraProgreso
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub
