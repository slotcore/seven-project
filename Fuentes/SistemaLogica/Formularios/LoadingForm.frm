VERSION 5.00
Object = "{62EC3EC3-A75A-11D1-AB74-004F4C006808}#1.0#0"; "MARCHOSO.OCX"
Begin VB.Form LoadingForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ejemplo de Marchoso.gif"
   ClientHeight    =   765
   ClientLeft      =   -60
   ClientTop       =   0
   ClientWidth     =   4170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   765
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MARCHOSOLib.Marchoso Marchoso1 
      Height          =   495
      Left            =   200
      TabIndex        =   0
      Top             =   120
      Width           =   495
      _Version        =   131072
      _ExtentX        =   882
      _ExtentY        =   882
      _StockProps     =   1
      BackColor       =   -2147483633
      FileName        =   ""
      AutoSize        =   0   'False
      Transparent     =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Procesando espere por favor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   940
      TabIndex        =   1
      Top             =   240
      Width           =   3060
   End
End
Attribute VB_Name = "LoadingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Iniciar()
    Marchoso1.FileName = App.Path & "\loading2.gif"
End Sub

Public Sub Detener()
    Marchoso1.FileName = ""
End Sub

Private Sub Form_Load()
    Iniciar
End Sub
