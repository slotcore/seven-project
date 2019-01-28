VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public NomRemitente As String
Public MailRemitente As String
Public MailDestino As String
Public Asunto As String
Public Cuerpo As String
Public ServidorSMTP As String

Function EnviarCorreo(ArchAdjunto() As String) As Boolean
    Dim A, xIndex As Integer
    xIndex = 0
    For A = 0 To UBound(ArchAdjunto)
        If F_NulosC(ArchAdjunto(A)) <> "" Then
            'UUfiles(A) = ArchAdjunto(A)
            UUfiles(xIndex) = UUEncodeFile(ArchAdjunto(A))
            xIndex = xIndex + 1
        End If
    Next A
    indexUUfiles = xIndex
    
    FrmMensaje.Fg1.Rows = 1
    For A = 0 To UBound(ArchAdjunto)
        If F_NulosC(ArchAdjunto(A)) <> "" Then
            FrmMensaje.Fg1.Rows = FrmMensaje.Fg1.Rows + 1
            FrmMensaje.Fg1.TextMatrix(FrmMensaje.Fg1.Rows - 1, 1) = ArchAdjunto(A)
        End If
    Next A
    FrmMensaje.txtUU.Text = "OC0001-00000001.pdf   OC0001-00000002.pdf   OC0001-00000003.pdf"
    FrmMensaje.txtSMTP.Text = ServidorSMTP
    FrmMensaje.txtFrom.Text = NomRemitente
    FrmMensaje.txtMailFrom.Text = MailRemitente
    FrmMensaje.txtMailTo = MailDestino
    FrmMensaje.txtSubject = Asunto
    FrmMensaje.txtBody = Cuerpo
    FrmMensaje.Show vbModal
    EnviarCorreo = FrmMensaje.SeEnvio
    Unload FrmMensaje
End Function
