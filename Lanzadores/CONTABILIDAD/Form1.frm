VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command17 
      Caption         =   "Kardex Resumen"
      Height          =   735
      Left            =   750
      TabIndex        =   18
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Reporte de Stocks"
      Height          =   735
      Left            =   750
      TabIndex        =   17
      Top             =   3570
      Width           =   1275
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Libro de Costos"
      Height          =   735
      Left            =   810
      TabIndex        =   16
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Costo de Personal"
      Height          =   735
      Left            =   9660
      TabIndex        =   15
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Consulta de Costo Mov"
      Height          =   735
      Left            =   3240
      TabIndex        =   14
      Top             =   2040
      Width           =   1275
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Metodo de Valorizacion"
      Height          =   735
      Left            =   4800
      TabIndex        =   13
      Top             =   240
      Width           =   1275
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Informe de Kardex"
      Height          =   735
      Left            =   750
      TabIndex        =   12
      Top             =   4500
      Width           =   1275
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Consulta Diario"
      Height          =   735
      Left            =   3240
      TabIndex        =   11
      Top             =   5520
      Width           =   1275
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Analisis de Costo de Produccion"
      Height          =   735
      Left            =   3240
      TabIndex        =   10
      Top             =   3840
      Width           =   1275
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Balance de Comprobacion"
      Height          =   735
      Left            =   810
      TabIndex        =   9
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CommandButton RepCosto 
      Caption         =   "Reporte de costo"
      Height          =   735
      Left            =   8250
      TabIndex        =   8
      Top             =   1980
      Width           =   1275
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Consulta Mayor"
      Height          =   735
      Left            =   3240
      TabIndex        =   7
      Top             =   4680
      Width           =   1275
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Consulta Kardex Resum. Val."
      Height          =   735
      Left            =   4800
      TabIndex        =   6
      Top             =   1050
      Width           =   1275
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Retencion"
      Height          =   735
      Left            =   3270
      TabIndex        =   5
      Top             =   1050
      Width           =   1275
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Registro Produccion"
      Height          =   735
      Left            =   8310
      TabIndex        =   4
      Top             =   3600
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Consulta de Costo de Prod"
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   3000
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Kardex Resumen Valorizado"
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Registro Produccion V2"
      Height          =   735
      Left            =   9750
      TabIndex        =   1
      Top             =   3540
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   735
      Left            =   9630
      TabIndex        =   0
      Top             =   6120
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.MostrarKardexValorizado xCon
    Set xfrm = Nothing
End Sub

Private Sub Command10_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.HojaTrabajo xCon
    Set xfrm = Nothing
End Sub

Private Sub Command11_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.AnalisisCostoProduccion xCon
    Set xfrm = Nothing
End Sub

Private Sub Command12_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.VerDiario xCon
    Set xfrm = Nothing
End Sub

Private Sub Command13_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.InformeKardexVal xCon
    Set xfrm = Nothing
End Sub

Private Sub Command14_Click()
    Dim xfrm As New sgi2_contabilidad.Mantenimiento
    xfrm.IdMenu = 256
    xfrm.IdUsuario = 1
    xfrm.ManConfigVal xCon
    Set xfrm = Nothing
End Sub

Private Sub Command15_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.ConsultaCostoMovimiento xCon
    Set xfrm = Nothing
End Sub

Private Sub Command17_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.MostrarStockResumen xCon, True
    Set xfrm = Nothing
End Sub

Private Sub Command19_Click()
    Dim xfrm As New sgi2_contabilidad.Mantenimiento
    xfrm.IdMenu = 256
    xfrm.IdUsuario = 1
    xfrm.verLibroCosto xCon
    Set xfrm = Nothing
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.MostrarKardexValorizado xCon
    Set xfrm = Nothing
End Sub

Private Sub Command5_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.ConsultaCostoParte xCon
    Set xfrm = Nothing
End Sub

Private Sub Command7_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.Mantenimiento
    xfrm.IdMenu = 31
    xfrm.IdUsuario = 1
    xfrm.ManRetencion xCon, 7
    Set xfrm = Nothing
End Sub

Private Sub Command8_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.ConsultaKardexValResum xCon
    Set xfrm = Nothing
End Sub

Private Sub Command9_Click()
    Dim xfrm As New sgi2_contabilidad.Consultas
    xfrm.Mayor xCon
    Set xfrm = Nothing
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Main
End Sub
