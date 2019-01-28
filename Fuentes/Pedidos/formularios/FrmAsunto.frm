VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmAsunto 
   Caption         =   "Ventas - Detalle Pedido"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Rt1 
      Height          =   2430
      Left            =   15
      TabIndex        =   13
      Top             =   1755
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4286
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"FrmAsunto.frx":0000
   End
   Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
      Height          =   300
      Left            =   1020
      TabIndex        =   2
      Top             =   810
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox TxtHorFin 
      Height          =   300
      Left            =   2430
      TabIndex        =   5
      Top             =   1125
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtSujeto 
      Height          =   300
      Left            =   1020
      TabIndex        =   0
      Text            =   "TxtSujeto"
      Top             =   105
      Width           =   6135
   End
   Begin VB.TextBox TxtDir 
      Height          =   300
      Left            =   1020
      TabIndex        =   1
      Text            =   "TxtDir"
      Top             =   405
      Width           =   6135
   End
   Begin MSMask.MaskEdBox TxtHorIni 
      Height          =   300
      Left            =   2430
      TabIndex        =   3
      Top             =   810
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
      Height          =   300
      Left            =   1020
      TabIndex        =   4
      Top             =   1125
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   30
      TabIndex        =   12
      Top             =   4125
      Width           =   8640
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   420
         Left            =   3540
         TabIndex        =   6
         Top             =   225
         Width           =   1560
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Final"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   1170
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Inicio"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   855
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   1530
      Width           =   855
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Sujeto"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   150
      Width           =   450
   End
   Begin VB.Label lblLocation 
      AutoSize        =   -1  'True
      Caption         =   "Direccion"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   450
      Width           =   675
   End
End
Attribute VB_Name = "FrmAsunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_pEditingEvent As CalendarEvent

Private Sub CmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    VerEstaEntrega
End Sub

Sub VerEstaEntrega()
    
    Dim HitTest As CalendarHitTestInfo
    Set HitTest = FrmCronoPedidos.CalendarControl.ActiveView.HitTest
    
    Set m_pEditingEvent = HitTest.ViewEvent.Event
    
    TxtSujeto.Text = m_pEditingEvent.Subject
    TxtDir.Text = m_pEditingEvent.Location
    TxtFchIni.Valor = Format(m_pEditingEvent.StartTime, "dd/mm/yyyy")
    TxtFchFin.Valor = Format(m_pEditingEvent.EndTime, "dd/mm/yyy")
    TxtHorIni.Text = Format(m_pEditingEvent.StartTime, "hh:mm:ss")
    TxtHorFin.Text = Format(m_pEditingEvent.EndTime, "hh:mm:ss")
    
    Rt1.Text = m_pEditingEvent.Body
    
    
    
'    Dim StartTime As Date, EndTime As Date
'    StartTime = DateFromString(cmbStartDate.Text, cmbStartTime.Text)
'    EndTime = DateFromString(cmbEndDate.Text, cmbEndTime.Text)
'
'    If chkAllDayEvent.Value = 1 Then
'        If DateDiff("s", TimeValue(EndTime), 0) = 0 Then
'            EndTime = EndTime + 1
'        End If
'    End If
'
'    If m_pEditingEvent.RecurrenceState <> xtpCalendarRecurrenceMaster Then
'        m_pEditingEvent.StartTime = StartTime
'        m_pEditingEvent.EndTime = EndTime
'    End If
'
'    m_pEditingEvent.Subject = txtSubject.Text
'    m_pEditingEvent.Location = txtLocation.Text
'    m_pEditingEvent.Body = txtBody
'    m_pEditingEvent.AllDayEvent = chkAllDayEvent.Value = 1
'    m_pEditingEvent.Label = cmbLabel.ItemData(cmbLabel.ListIndex)
'    m_pEditingEvent.BusyStatus = cmbShowTimeAs.ListIndex
'    If cmbSchedule.ListIndex >= 0 And cmbSchedule.ListIndex < cmbSchedule.ListCount Then
'        m_pEditingEvent.ScheduleID = cmbSchedule.ItemData(cmbSchedule.ListIndex)
'    End If
'
'    m_pEditingEvent.PrivateFlag = chkPrivate.Value = 1
'    m_pEditingEvent.MeetingFlag = chkMeeting.Value = 1
'
'    If Not chkReminder.Value = m_pEditingEvent.Reminder Then
'        m_pEditingEvent.Reminder = chkReminder.Value
'        m_pEditingEvent.ReminderSoundFile = "D:\Backup_10_12\Desktop\mustbuild.wav"
'    End If
'
'    If chkReminder.Value Then
'        If Not Val(cmbReminder.Text) = m_pEditingEvent.ReminderMinutesBeforeStart Then
'            m_pEditingEvent.ReminderMinutesBeforeStart = CalcStandardDurations_0m_2wLong(cmbReminder.Text)
'        End If
'    End If
'
'    txtMarkupText.Text = Trim(txtMarkupText.Text)
'    If Len(txtMarkupText.Text) > 0 Then
'        m_pEditingEvent.CustomProperties.Property("xtpMarkupText") = txtMarkupText.Text
'    Else
'        m_pEditingEvent.CustomProperties.Remove "xtpMarkupText"
'    End If

End Sub

