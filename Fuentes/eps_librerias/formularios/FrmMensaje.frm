VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMensaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AnoniMMMail                                                    by The-Pirat 2004"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   6060
      TabIndex        =   22
      Top             =   30
      Width           =   2820
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   780
      Left            =   105
      TabIndex        =   21
      Top             =   3390
      Width           =   6165
      _cx             =   10874
      _cy             =   1376
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmMensaje.frx":0000
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
   Begin MSComctlLib.ProgressBar PB 
      Height          =   270
      Left            =   2025
      TabIndex        =   14
      Top             =   5820
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtUU 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1935
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5550
      Visible         =   0   'False
      Width           =   6135
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   8445
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtBody 
      Height          =   1695
      Left            =   90
      TabIndex        =   4
      Top             =   1635
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   2990
      _Version        =   393217
      TextRTF         =   $"FrmMensaje.frx":0050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   4695
      TabIndex        =   12
      Top             =   6135
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtSMTP 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6960
      TabIndex        =   6
      Text            =   "mx1.hotmail.com"
      Top             =   210
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtSubject 
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1020
      Width           =   7830
   End
   Begin VB.TextBox txtMailTo 
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   4935
   End
   Begin VB.TextBox txtMailFrom 
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   4935
   End
   Begin VB.TextBox txtFrom 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin MSWinsockLib.Winsock sck 
      Left            =   6480
      Top             =   465
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   585
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4425
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   6390
      TabIndex        =   17
      Top             =   3300
      Width           =   2505
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "Enviar"
         Height          =   375
         Left            =   375
         TabIndex        =   20
         Top             =   345
         Width           =   1815
      End
      Begin VB.CommandButton cmdCerrar 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   375
         TabIndex        =   19
         Top             =   1155
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdjuntar 
         Caption         =   "Adjuntar Archivo..."
         Height          =   375
         Left            =   375
         TabIndex        =   18
         Top             =   750
         Width           =   1815
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   135
      TabIndex        =   16
      Top             =   4185
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Mensaje"
      Height          =   195
      Left            =   135
      TabIndex        =   15
      Top             =   1410
      Width           =   600
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "SMTP Server:"
      Height          =   255
      Left            =   7125
      TabIndex        =   11
      Top             =   525
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Para"
      Height          =   195
      Left            =   135
      TabIndex        =   10
      Top             =   765
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Asunto"
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   1065
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Correo"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   450
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "De"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   150
      Width           =   210
   End
End
Attribute VB_Name = "FrmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SeEnvio As Boolean

Private Sub cmdAdjuntar_Click()
    If indexUUfiles > 9 Then
        MsgBox "No puede adjuntar más de 10 archivos", vbCritical, "Error"
        Exit Sub
    End If
    
    CD.DialogTitle = "Adjuntar Archivo..."
    CD.Filter = "Todos los archivos (*.*)|*.*"
    CD.Action = 1
    If CD.FileName = "" Then Exit Sub
    
    Me.Caption = "Codificando Archivo..."
    cmdEnviar.Enabled = False
    cmdAdjuntar.Enabled = False
    PB.Value = 0
    PB.Visible = True
    UUfiles(indexUUfiles) = UUEncodeFile(CD.FileName)
    
    txtUU.Visible = True
    indexUUfiles = indexUUfiles + 1
    txtUU.Text = txtUU.Text & CD.FileTitle & "   "
    cmdEnviar.Enabled = True
    cmdAdjuntar.Enabled = True
    PB.Visible = False
    Me.Caption = exCaption
End Sub

Private Sub cmdCancel_Click()
    Call DesConectar
End Sub

Private Sub cmdCerrar_Click()
    SeEnvio = False
    Me.Hide
End Sub

Private Sub cmdEnviar_Click()
    If txtSMTP = "" Or txtFrom = "" Or txtMailFrom = "" Or txtMailTo = "" Then
        MsgBox "Datos incompletos", vbCritical, "Error"
        Exit Sub
    End If

    If txtSubject = "" And txtBody = "" Then
        MsgBox "Debe escribir un Asunto o un Mensaje", vbCritical, "Error"
        Exit Sub
    End If
    
    If MsgBox("¿Confirma envío?", vbYesNo Or vbQuestion, "") = vbNo Then Exit Sub
    
    DServer = txtSMTP
    
    Enviar txtFrom, txtMailFrom, txtMailTo, txtSubject, txtBody.Text
    SeEnvio = True
    'MsgBox "El correo se envio con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Sub

Private Sub Form_Activate()
    txtFrom.SetFocus
    exCaption = Me.Caption
    'cmdEnviar_Click
End Sub

Private Sub Form_Load()
    txtStatus = ""
    SeEnvio = False
    'indexUUfiles = 0
End Sub

Private Sub sck_DataArrival(ByVal bytesTotal As Long)
    sck.GetData Respuesta
    Code = Left(Respuesta, 3)
    Call AddStatus("<- " & Respuesta)
    If Code >= 200 And Code <= 399 Then
        Select Case SendStatus
            Case CONECTED
                sck.SendData "HELO " & DHelo & vbCrLf
                SendStatus = MailFrom
            Case MailFrom
                sck.SendData "MAIL FROM:<" & DMailFrom & ">" & vbCrLf
                AddStatus ("-> MAIL FROM:<" & DMailFrom & ">")
                SendStatus = RCPTTO
            Case RCPTTO
                sck.SendData "RCPT TO:<" & DRcptTo & ">" & vbCrLf
                AddStatus ("-> RCPT TO:<" & DRcptTo & ">")
                SendStatus = DATAC
            Case DATAC
                sck.SendData "DATA" & vbCrLf
                SendStatus = MESSAGGE
            Case MESSAGGE
                sck.SendData "FROM: " & DFrom & vbCrLf
                AddStatus ("-> FROM: " & DFrom)
                sck.SendData "SUBJECT: " & DSubject & vbCrLf
                AddStatus ("-> SUBJECT: " & DSubject)
                sck.SendData "X-Priority: 1" & vbCrLf & "X-MSMail-Priority: High" & vbCrLf
                sck.SendData DMensaje & vbCrLf
                                
                Dim i As Byte, Buff As String
                If indexUUfiles > 0 Then
                    For i = 0 To indexUUfiles
                        Buff = Buff & UUfiles(i)
                    Next i
                    sck.SendData Buff
                End If
                
                sck.SendData vbCrLf & "." & vbCrLf
                
                SendStatus = QUIT
            Case QUIT
                Call AddStatus("*** MAIL ENVIADO OK ***")
                sck.SendData "QUIT" & vbCrLf
                Call DesConectar
                MsgBox "El correo se envio con exito a: " & txtMailTo.Text, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Me.Hide
                SeEnvio = True
        End Select
    Else
        Call DesConectar
    End If
End Sub

Private Sub sck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call AddStatus("Error nº:" & Number & " " & Description)
    Call DesConectar
End Sub


Private Sub txtBody_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long
If Shift <> 0 Then Exit Sub
If KeyCode = 9 Then
    i = txtBody.SelStart
    txtBody.Text = Left(txtBody.Text, i) & Chr(9) & Mid(txtBody.Text, i + 1)
    txtBody.SelStart = i + 1
    KeyCode = 0
End If
End Sub



