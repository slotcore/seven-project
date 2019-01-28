VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmManConcepto11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Concepto"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmManConcepto11.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   9300
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   90
      TabIndex        =   9
      Top             =   1095
      Width           =   9195
      Begin VB.TextBox txt 
         DataField       =   "t_conceptoboleta_nombrecorto"
         Height          =   315
         Index           =   3
         Left            =   7020
         MaxLength       =   30
         TabIndex        =   23
         Top             =   270
         Width           =   2130
      End
      Begin VB.CommandButton cmd_buscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5265
         MouseIcon       =   "FrmManConcepto11.frx":000C
         Picture         =   "FrmManConcepto11.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Buscar Concepto"
         Top             =   615
         Width           =   435
      End
      Begin VB.Frame Frame6 
         Caption         =   "Formula"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   105
         TabIndex        =   18
         Top             =   1305
         Width           =   9015
         Begin VB.CheckBox chk_formula 
            Caption         =   "Formula"
            DataField       =   "F_ConceptoBoleta_Formula"
            Height          =   225
            Left            =   75
            TabIndex        =   21
            Top             =   255
            Width           =   885
         End
         Begin VB.CommandButton cmd_formula 
            Caption         =   "Editar Formula"
            Height          =   645
            Left            =   30
            Picture         =   "FrmManConcepto11.frx":0418
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   540
            Width           =   1275
         End
         Begin VB.TextBox txt_formula 
            DataField       =   "N_ConceptoBoleta_Formula"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   1395
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   195
            Width           =   7545
         End
      End
      Begin VB.TextBox txt 
         DataField       =   "N_ConceptoBoleta_NomVariable"
         Height          =   330
         Index           =   1
         Left            =   1590
         MaxLength       =   45
         TabIndex        =   17
         Top             =   975
         Width           =   3630
      End
      Begin VB.TextBox txt 
         DataField       =   "n_conceptoboleta_nombre"
         Height          =   360
         Index           =   0
         Left            =   1590
         MaxLength       =   45
         TabIndex        =   16
         Top             =   615
         Width           =   3630
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1470
         TabIndex        =   13
         Top             =   2505
         Width           =   7125
         Begin VB.OptionButton opt_prestamo 
            Caption         =   "Suceptible a Prestamo"
            Height          =   240
            Index           =   1
            Left            =   3450
            TabIndex        =   15
            Tag             =   "F_ConceptoBoleta_APrestamo"
            Top             =   180
            Width           =   2190
         End
         Begin VB.OptionButton opt_prestamo 
            Caption         =   "No Suceptible a Prestamo"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Tag             =   "F_ConceptoBoleta_APrestamo"
            Top             =   180
            Width           =   2200
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1455
         TabIndex        =   10
         Top             =   2955
         Width           =   7125
         Begin VB.OptionButton opt_planilla 
            Caption         =   "No Considerar en Planilla"
            Height          =   255
            Index           =   1
            Left            =   3450
            TabIndex        =   12
            Tag             =   "F_ConceptoBoleta_planilla"
            Top             =   135
            Width           =   2130
         End
         Begin VB.OptionButton opt_planilla 
            Caption         =   "Considerar en Planilla"
            Height          =   285
            HelpContextID   =   1
            Index           =   0
            Left            =   135
            TabIndex        =   11
            Tag             =   "F_ConceptoBoleta_planilla"
            Top             =   120
            Value           =   -1  'True
            WhatsThisHelpID =   1
            Width           =   1950
         End
      End
      Begin MSDataListLib.DataCombo dcb_tipoconcepto 
         DataField       =   "c_tipoconcepto_codigo"
         Height          =   315
         Left            =   1590
         TabIndex        =   24
         Top             =   255
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Corto :"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   5880
         TabIndex        =   28
         ToolTipText     =   "Es nombre abreviado del concepto, con posibilidad de que aparesca en el reporte"
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Variable :"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   27
         Top             =   1035
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Concepto :"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   315
         Width           =   1350
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Concepto :"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   25
         Top             =   675
         Width           =   1380
      End
   End
   Begin VB.Frame aa 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   1958
      TabIndex        =   2
      Top             =   5805
      Width           =   5355
      Begin VB.CommandButton CMD 
         Caption         =   "Sali&r"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   5
         Left            =   4350
         MouseIcon       =   "FrmManConcepto11.frx":09A2
         MousePointer    =   99  'Custom
         Picture         =   "FrmManConcepto11.frx":0CAC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   250
         Width           =   855
      End
      Begin VB.CommandButton CMD 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   0
         Left            =   75
         MouseIcon       =   "FrmManConcepto11.frx":10EE
         MousePointer    =   99  'Custom
         Picture         =   "FrmManConcepto11.frx":13F8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   250
         Width           =   855
      End
      Begin VB.CommandButton CMD 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   1
         Left            =   930
         MouseIcon       =   "FrmManConcepto11.frx":192A
         MousePointer    =   99  'Custom
         Picture         =   "FrmManConcepto11.frx":1C34
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   250
         Width           =   855
      End
      Begin VB.CommandButton CMD 
         Caption         =   "Mo&dificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   2
         Left            =   1785
         MouseIcon       =   "FrmManConcepto11.frx":2166
         MousePointer    =   99  'Custom
         Picture         =   "FrmManConcepto11.frx":2470
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   250
         Width           =   855
      End
      Begin VB.CommandButton CMD 
         Caption         =   "Eli&minar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   3
         Left            =   2640
         MouseIcon       =   "FrmManConcepto11.frx":29A2
         MousePointer    =   99  'Custom
         Picture         =   "FrmManConcepto11.frx":2CAC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   250
         Width           =   855
      End
      Begin VB.CommandButton CMD 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   4
         Left            =   3495
         MouseIcon       =   "FrmManConcepto11.frx":2E36
         MousePointer    =   99  'Custom
         Picture         =   "FrmManConcepto11.frx":3140
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   250
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   90
      OleObjectBlob   =   "FrmManConcepto11.frx":3672
      Top             =   5955
   End
   Begin VB.TextBox txt 
      DataField       =   "t_conceptoboleta_comentario"
      Height          =   645
      Index           =   2
      Left            =   1065
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Tag             =   "null"
      Top             =   4605
      Width           =   8175
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario :"
      Height          =   210
      Index           =   2
      Left            =   90
      TabIndex        =   0
      Top             =   4635
      Width           =   900
   End
End
Attribute VB_Name = "FrmManConcepto11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst_Concepto As New ADODB.Recordset

Private Sub chk_cantidad_Click()

End Sub
Private Sub chk_formula_Click()
Dim Operacion As String
Operacion = OperacionActual(Me.Toolbar)
If Operacion = "NG" Then Exit Sub
  cmd_formula.Enabled = chk_formula.Value
If chk_formula.Value = 1 Then
   If rst_Concepto.State = 0 Then Exit Sub
   If rst_Concepto.EOF Or rst_Concepto.BOF Then Exit Sub
   txt_formula.Text = rst_Concepto.Fields("n_conceptoboleta_formula") & ""
Else
  txt_formula.Text = ""
End If

End Sub


Private Sub cmd_editar_pos_Click()
frm_pl_ordenarconceptos.Show 1
End Sub

Private Sub cmd_formula_Click()
If dcb_tipoconcepto.MatchedWithList = False Then
   MsgBox "Primero debe seleccionar Tipo de Concepto", vbExclamation
   Exit Sub
ElseIf txt(0).Text = "" Then
   MsgBox "Primero debe Ingresar Nombre de Concepto", vbExclamation
   Exit Sub
End If

frm_pl_mn_concepto_formula.Show 1
If txt_formula.Text <> "" Then
   chk_formula.Value = 1
End If
End Sub

Private Sub dcb_tipoconcepto_Change()
If dcb_tipoconcepto.BoundText <> "TCP00002" Then
   opt_prestamo(0).Value = True
   Frame3.Enabled = False
Else
   opt_prestamo(1).Value = True
   Frame3.Enabled = True
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
   If OperacionActual(Toolbar) = "NG" Then
      rst_Concepto.Requery
      poner_datos
   End If
Else
'  Navegador Me, Toolbar, KeyCode, Shift
End If
End Sub

Private Sub Form_Load()
HabilitarConcepto False
abrir_rst rst_Concepto, "select * from plan_concepto_boleta where f_conceptoboleta_invisible<>0 order by c_tipoconcepto_codigo asc,Q_conceptoboleta_posicion asc", adOpenDynamic, adLockReadOnly
PonerDatosDataCombo "plan_tipo_concepto", dcb_tipoconcepto, "", "c_tipoconcepto_codigo", "n_tipoconcepto_nombre"
poner_datos

Degrade pic1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
  If MsgBox("Seguro Desea Salir", vbYesNo + vbQuestion) = vbNo Then
     Cancel = 1
  Else
    Set rst_Concepto = Nothing
    Unload Me
  End If
End If
End Sub


Private Sub opt_prestamo_Click(Index As Integer)
If opt_prestamo(Index).Value = True And Index = 0 Then
''   chk_cuotas.Visible = False
'   chk_cuotas.Value = False
Else
'   chk_cuotas.Visible = True
End If
End Sub

''''Private Sub opt_tipo_Click()
''''If opt_tipo.Value = True Then
''''   chk_tipo(0).Value = 0
''''   chk_tipo(1).Value = 0
''''End If
''''End Sub

Sub accionesToolBar(Index As Integer)
Select Case Index
Case 1 'boton nuevo
        cmd_Click 0
Case 2 'boton grabar
        cmd_Click 1
Case 3 'boton modificar
        cmd_Click 2
Case 4 'boton eliminar
        cmd_Click 3
Case 5 'boton cancelar
        cmd_Click 4
Case 13 'boton salir
        cmd_Click 5
Case 7 'primer registro
        If rst_Concepto.RecordCount < 1 Then Exit Sub
        rst_Concepto.MoveFirst
        poner_datos
Case 8 'anterior registro
        If rst_Concepto.RecordCount < 1 Then Exit Sub
        rst_Concepto.MovePrevious
        If rst_Concepto.BOF Then rst_Concepto.MoveFirst
        poner_datos
    
        
Case 9  'siguiente registro
        If rst_Concepto.RecordCount < 1 Then Exit Sub
        rst_Concepto.MoveNext
        If rst_Concepto.EOF Then rst_Concepto.MoveLast
        poner_datos
         
Case 10  'ultimo registro
             If rst_Concepto.RecordCount < 1 Then Exit Sub
             rst_Concepto.MoveLast
             poner_datos
End Select
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
accionesToolBar Button.Index
 'If Button.Index >= 7 And Button.Index <= 10 Then poner_datos
End Sub
Sub cmd_Click(Index As Integer)
On Error GoTo error
Select Case Index
 Case 0  'boton nuevo
          nuevo
          dcb_tipoconcepto.SetFocus
          botones Me, Me.Toolbar, False
        
Case 1  'boton guardar
          If grabar_datos("N") = True Then
             HabilitarConcepto False
             botones Me, Me.Toolbar, True
             
          End If
          
 Case 2  'boton modificar despues nos ponemos de acuerdo como va a a hacer
          If rst_Concepto.RecordCount < 1 Or lbl_codigo.Caption = "" Then
             MsgBox "No hay datos que Modificar", vbExclamation
             Exit Sub
          End If
          
          If Left(lbl_codigo.Caption, 3) = "SYS" And Toolbar.Buttons(3).Tag = "Mo&dificar" Then
             MsgBox "Registro no se puede modificar porque es usado por el sistema" & vbCr & "Solo Puede Modificarse la formula", vbInformation
             Toolbar.Buttons(3).Tag = "Aceptar"
             'botones Me, Me.Toolbar, True
              ToolModifica Toolbar, txt, False
              txt(0).Enabled = False
              txt(1).Enabled = False
'              Frame3.Enabled = False
'              Frame4.Enabled = False
'              Frame5.Enabled = False
              'Frame7.Enabled = True
              cmd_formula.Enabled = True
              Frame6.Enabled = True
              
              Exit Sub
          End If
          
          
          If Toolbar.Buttons(3).Tag = "Aceptar" Then
            If grabar_datos("M") = True Then
               Toolbar.Buttons(3).Tag = "Mo&dificar"
               botones Me, Me.Toolbar, True
               HabilitarConcepto False
            End If
         ElseIf Toolbar.Buttons(3).Tag = "Mo&dificar" Then
             txt(1).Enabled = False
             ToolModifica Toolbar, txt, False
             Toolbar.Buttons(3).Tag = "Aceptar"
             HabilitarConcepto True
         End If
 
 Case 3  'boton eliminar
          If rst_Concepto.RecordCount < 1 Then
              MsgBox "No hay datos que Eliminar", vbExclamation
              Exit Sub
          End If
          
          If Left(lbl_codigo.Caption, 3) = "SYS" Then
             MsgBox "Registro no se puede eliminar porque es usado por el sistema", vbInformation
             Exit Sub
          End If
          
'''''''''''          abrir_rst rst, "SELECT Prestamo_Adelanto.C_Prestamo_Codigo From Prestamo_Adelanto WHERE Prestamo_Adelanto.c_conceptoboleta_codigo='" & lbl_Codigo.Caption & "' limit 1", , , True, , False
'''''''''''          If rst.RecordCount > 0 Then
'''''''''''                 MsgBox "No se puede eliminar registro porque esta siendo utilizado en otro Archivo", vbExclamation
'''''''''''                 Set rst = Nothing
'''''''''''                 Exit Sub
'''''''''''          Else
'''''''''''                abrir_rst rst, "select c_conceptoboleta_codigo from detalle_formula where c_conceptoboleta_codigoa='" & lbl_Codigo.Caption & "' limit 1", , , True, , False
'''''''''''                If rst.RecordCount > 0 Then
'''''''''''                   MsgBox "No se puede eliminar registro porque esta siendo utilizado en otro Archivo", vbExclamation
'''''''''''                   Set rst = Nothing
'''''''''''                   Exit Sub
'''''''''''                End If
'''''''''''                 Set rst = Nothing
'''''''''''          End If
'''''''''''
'''''''''''          abrir_rst rst, "select c_conceptoboleta_codigo from detalletipoboleta_concepto where c_conceptoboleta_codigo='" & lbl_Codigo.Caption & "' limit 1", , , True, , False
'''''''''''          If rst.RecordCount > 0 Then
'''''''''''                 MsgBox "No se puede eliminar registro porque esta siendo utilizado en otro Archivo", vbExclamation
'''''''''''                 Set rst = Nothing
'''''''''''                 Exit Sub
'''''''''''          End If
'''''''''''          Set rst = Nothing
'''''''''''
'''''''''''          abrir_rst rst, "select c_conceptoboleta_codigo from detalle_emp_conceptoboleta where c_conceptoboleta_codigo='" & lbl_Codigo.Caption & "' limit 1", , , True, , False
'''''''''''          If rst.RecordCount > 0 Then
'''''''''''                 MsgBox "No se puede eliminar registro porque esta siendo utilizado en otro Archivo", vbExclamation
'''''''''''                 Set rst = Nothing
'''''''''''                 Exit Sub
'''''''''''          End If
'''''''''''          Set rst = Nothing
'''''''''''
'''''''''''
'''''''''''           abrir_rst rst, "select c_conceptoboleta_codigo from detalleconceptoboleta where c_conceptoboleta_codigo='" & lbl_Codigo.Caption & "' limit 1", , , True, , False
'''''''''''          If rst.RecordCount > 0 Then
'''''''''''                 MsgBox "No se puede eliminar registro porque esta siendo utilizado en otro Archivo", vbExclamation
'''''''''''                 Set rst = Nothing
'''''''''''                 Exit Sub
'''''''''''          End If
'''''''''''          Set rst = Nothing
'''''''''''
'''''''''''          If MsgBox("Seguro desea eliminar el registro actual", vbYesNo + vbQuestion) = vbYes Then
'''''''''''                 bd.Execute "delete from detalle_formula where c_conceptoboleta_codigo='" & lbl_Codigo.Caption & "'"
'''''''''''                 rst_concepto.Delete
'''''''''''                 rst_concepto.UpdateBatch
'''''''''''                 rst_concepto.MoveFirst
'''''''''''                 poner_datos
'''''''''''          End If
'''''''''''          Set rst = Nothing
'''''''''''
         'botones Me, Me.Toolbar, true
         
         Dim cls As New cls_pl_mantenimiento
         Dim flag As String
         
         If MsgBox("Seguro desea eliminar el registro actual", vbYesNo + vbInformation) = vbNo Then Exit Sub
         Set cls = New cls_pl_mantenimiento
         With cls
                 .CadenaConexion = bd.ConnectionString
                 flag = .PL_Concepto_Eliminar(lbl_codigo.Caption)
         End With
         Set cls = Nothing
         If LCase(flag) = "ok" Then
            MsgBox "Se elimino registro", vbInformation
            LimpiarConcepto
         Else
           MsgBox flag, vbInformation
         End If
 Case 4  'boton cancelar
'          rst_concepto.CancelBatch
          'rst_Concepto.MoveFirst
          HabilitarConcepto False
          poner_datos
          Toolbar.Buttons(3).Tag = "Mo&dificar"
          botones Me, Me.Toolbar, True
         
 Case 5  'boton salir
         Form_QueryUnload 0, 0
End Select
Exit Sub
error:
MsgBox Err.Description, vbCritical
Err.Clear
End Sub
Sub poner_datos()
If rst_Concepto.RecordCount = 0 Or rst_Concepto.EOF = True Or rst_Concepto.BOF = True Then
   LimpiarConcepto
   Exit Sub
End If

'if lbl_codigo.Caption = rst_concepto.Fields("c_conceptoboleta_codigo") then exit sub

lbl_codigo.Caption = rst_Concepto.Fields("c_conceptoboleta_codigo")
For i = 0 To txt.Count - 1
    txt(i).Text = rst_Concepto.Fields(txt(i).DataField) & ""

Next

If txt(4).Text = "-1" Then txt(4).Text = ""

dcb_tipoconcepto.BoundText = rst_Concepto.Fields(dcb_tipoconcepto.DataField)
txt_formula.Text = rst_Concepto.Fields(txt_formula.DataField) & ""
txt_formula.Tag = rst_Concepto.Fields("formula2") & ""
chk_formula.Value = rst_Concepto.Fields(chk_formula.DataField)
opt_prestamo(rst_Concepto.Fields(opt_prestamo(0).Tag)).Value = True
opt_planilla(rst_Concepto.Fields(opt_planilla(0).Tag)).Value = True
End Sub
Sub nuevo()
HabilitarConcepto True
LimpiarConcepto
End Sub
Sub LimpiarConcepto()
LimpiaText txt
lbl_codigo.Caption = ""
dcb_tipoconcepto.Text = ""
chk_formula.Value = False

txt_formula.Tag = ""
txt_formula.Text = ""

opt_prestamo(0).Value = False
opt_prestamo(1).Value = False
opt_planilla(0).Value = False
opt_planilla(1).Value = False
End Sub
Sub HabilitarConcepto(band As Boolean)
dcb_tipoconcepto.Enabled = band
habilitar txt, band
If OperacionActual(Toolbar) = "M" Then
   txt(1).Enabled = False
End If
Frame4.Enabled = band
Frame3.Enabled = band
Frame6.Enabled = band
cmd_formula.Enabled = band
cmd_buscar.Enabled = Not band
End Sub

Function grabar_datos(Operacion As String) As Boolean
Dim flag As Integer, band As String
Dim cls As New cls_pl_mantenimiento

If dcb_tipoconcepto.MatchedWithList = False Then
   MsgBox "Seleccione el Tipo de Concepto ", vbCritical
   dcb_tipoconcepto.SetFocus
   grabar_datos = False
   Exit Function
End If

If UCase(Left(lbl_codigo.Caption, 3)) = "SYS" Then
   txt(3).Tag = "null" 'nombre corto
   txt(4).Tag = "null" 'posicion
End If
txt(4).Tag = "null"

flag = Validar(txt)
If Validar(txt) >= 0 Then
   MsgBox "Debe ingresar datos en el campo " & lbl(flag).Caption, vbCritical
   txt(flag).SetFocus
   grabar_datos = False
   Exit Function
End If


txt(3).Tag = ""
txt(4).Tag = ""

If chk_formula.Value = 1 And txt_formula.Text = "" Then
   MsgBox "Debe de Editar la formula", vbCritical
   grabar_datos = False
   Exit Function
End If

flag = OptSeleccionado(opt_planilla)
If flag = -1 Then
   MsgBox "Seleccione si se considera o no en planilla el Concepto", vbCritical
   grabar_datos = False
   Exit Function
End If

flag = OptSeleccionado(opt_prestamo)
If flag = -1 Then
   MsgBox "Seleccione si es Suceptible a Prestamo el Concepto", vbCritical
   grabar_datos = False
   Exit Function
End If
  
If MsgBox("Seguro desea guardar?", vbInformation + vbYesNo) = vbNo Then Exit Function


Set cls = New cls_pl_mantenimiento
With cls
        .CadenaConexion = bd.ConnectionString
      band = .PL_Concepto_Insertar_Modificar(lbl_codigo.Caption, txt(0).Text, _
                                                                  txt(2).Text, dcb_tipoconcepto.BoundText, _
                                                                  OptSeleccionado(opt_prestamo), 1, txt(1).Text, _
                                                                  OptSeleccionado(opt_planilla), txt_formula.Text, _
                                                                  chk_formula, 1, txt(3).Text, Cod_usuario, Operacion)
End With
Set cls = Nothing

If band = "-1" Then 'error al grabar
  grabar_datos = False
Else
    grabar_datos = True
    If Operacion = "M" Then
       MsgBox band, vbInformation
    Else
       lbl_codigo.Caption = band
    End If
    rst_Concepto.Requery
End If
End Function
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Then
   If (validar_letras(KeyAscii) = False And keryascii <> 13 And Chr(KeyAscii) <> "_") Or KeyAscii = 32 Then
      MsgBox "En este campo solo se permiten Letras", vbExclamation
      KeyAscii = 0
   End If
ElseIf Index = 4 Then
  If validar_numero(KeyAscii) = False Then
     MsgBox "En este campo solo se permiten numeros", vbCritical
     KeyAscii = 0
  End If
End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
txt(Index).Text = Trim(txt(Index).Text)
End Sub
