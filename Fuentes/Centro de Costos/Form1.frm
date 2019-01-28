VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "PERCEPCIONE"
      Height          =   975
      Left            =   480
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      Caption         =   "08-06"
      Height          =   735
      Left            =   6600
      TabIndex        =   13
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Costo Ventas Resumen"
      Height          =   735
      Left            =   8760
      TabIndex        =   12
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Costo Ventas"
      Height          =   735
      Left            =   8760
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Hoja de Trabajo"
      Height          =   735
      Left            =   3720
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Libro de Costos"
      Height          =   735
      Left            =   8760
      TabIndex        =   9
      Top             =   180
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Detraccion"
      Height          =   735
      Left            =   3720
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Registro de Compras"
      Height          =   735
      Left            =   3720
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Diario"
      Height          =   735
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Centro de Costos x Area"
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Des - Ingresos"
      Height          =   735
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Des - Egresos"
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kardex"
      Height          =   735
      Left            =   450
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Retenciones"
      Height          =   735
      Left            =   450
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   735
      Left            =   8760
      TabIndex        =   0
      Top             =   4560
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.KeyPreview = True
    Main
End Sub
'PERCEPCION
Private Sub Command11_Click()
   
    Dim xfrm As New sgi2_contabilidad.Mantenimiento
xfrm.IdMenu = 92
xfrm.IdUsuario = 1
xfrm.ManPercepcion xCon, CInt(Mid(Date, 4, 2))
End Sub
'RETENCIONES
Private Sub Command1_Click()
Dim xfrm As New sgi2_contabilidad.Mantenimiento
xfrm.IdMenu = 92
xfrm.IdUsuario = 1
xfrm.ManRetencion xCon, CInt(Mid(Date, 4, 2))
End Sub
'KARDEX
Private Sub Command3_Click()
Dim xfrm As New sgi2_contabilidad.Consultas
xfrm.MostrarKardexValorizado xCon
Set xfrm = Nothing

End Sub
'DIARIO
Private Sub Command7_Click()
Dim xfrm As New sgi2_contabilidad.Consultas
xfrm.VerDiario xCon
Set xfrm = Nothing
End Sub
'REGISTRO DE COMPRAS
Private Sub Command8_Click()

Dim xfrm As New sgi2_contabilidad.Consultas
xfrm.VerRegCompras xCon

Set xfrm = Nothing

End Sub
'DETRACCION
Private Sub Command9_Click()

    ' EJECUTA MENU
    Dim xfrm As New sgi2_contabilidad.Mantenimiento
    xfrm.IdMenu = 1
    xfrm.IdUsuario = xIdUsuario
    xfrm.ManDetraccion xCon, AP_MESTRA, DET_Compra
    Set xfrm = Nothing
End Sub
'HOJA DE TRABAJO
Private Sub Command12_Click()
Dim xfrm As New sgi2_contabilidad.Consultas
xfrm.HojaTrabajo xCon

Set xfrm = Nothing

End Sub
'LIBRO DE COSTO
Private Sub Command10_Click()
Dim xfrm As New sgi2_contabilidad.Mantenimiento
xfrm.IdUsuario = 1
xfrm.IdMenu = 131
xfrm.verLibroCosto xCon
End Sub
'COSTO VENTAS
Private Sub Command13_Click()
Dim xfrm As New sgi2_contabilidad.Consultas

xfrm.MostrarCostosVenta xCon
End Sub
'COSTO VENTAS RESUMEN
Private Sub Command15_Click()
Dim xfrm As New sgi2_contabilidad.Consultas

xfrm.MostrarCostosVentaRes xCon
End Sub
'SALIR
Private Sub Command2_Click()
    Unload Me
End Sub

'no sirve
Private Sub Command4_Click()
Dim xfrm As New sgi2_contabilidad2.mantenimientos
xfrm.IdUsuario = 1
xfrm.IdMenu = 131
xfrm.ManDestinos 2, xCon

End Sub

Private Sub Command5_Click()
Dim xfrm As New sgi2_contabilidad2.mantenimientos
xfrm.IdUsuario = 1
xfrm.IdMenu = 130
xfrm.ManDestinos 1, xCon
End Sub

Private Sub Command6_Click()
Dim xfrm As New sgi2_contabilidad2.mantenimientos
xfrm.IdUsuario = 1
xfrm.ManCentroCostoArea xCon
End Sub







