VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form MensajeBoxForm 
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList MensajeImageList 
      Left            =   4440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   700
      Left            =   120
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   4
      Top             =   200
      Width           =   700
   End
   Begin VB.CommandButton ActionCmd 
      Caption         =   "Copiar Mensaje"
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   1500
   End
   Begin VB.CommandButton ActionCmd 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1500
   End
   Begin VB.CommandButton ActionCmd 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MensajeLabel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1290
   End
End
Attribute VB_Name = "MensajeBoxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTipoMensaje As Long

Public Property Get TipoMensaje() As Long
    TipoMensaje = mTipoMensaje
End Property
Public Property Let TipoMensaje(ByVal NewValue As Long)
    mTipoMensaje = NewValue
End Property

