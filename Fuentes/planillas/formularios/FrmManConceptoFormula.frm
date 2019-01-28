VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmManConceptoFormula 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Concepto - Editor de Fórmulas"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "&Limpiar"
      Height          =   495
      Index           =   2
      Left            =   5805
      TabIndex        =   15
      ToolTipText     =   "Limpiar Formula"
      Top             =   4800
      Width           =   1320
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Probar"
      Height          =   495
      Index           =   3
      Left            =   7110
      TabIndex        =   14
      ToolTipText     =   "Probar Formula"
      Top             =   4800
      Width           =   1320
   End
   Begin VB.TextBox txt_formula 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "FrmManConceptoFormula.frx":0000
      Top             =   3585
      Width           =   10995
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agregar Constante"
      Height          =   675
      Left            =   15
      TabIndex        =   4
      Top             =   4620
      Width           =   3030
      Begin VB.TextBox txt_constante 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   6
         Text            =   "txt_constante"
         Top             =   270
         Width           =   1590
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Agregar"
         Height          =   345
         Index           =   4
         Left            =   1785
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.TextBox Text2 
      Height          =   1125
      Left            =   615
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   6150
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4605
      TabIndex        =   2
      Top             =   6225
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Height          =   495
      Index           =   0
      Left            =   8415
      TabIndex        =   1
      ToolTipText     =   "Aceptar Formula"
      Top             =   4800
      Width           =   1320
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      Height          =   495
      Index           =   1
      Left            =   9720
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   4800
      Width           =   1320
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1185
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoFormula.frx":000E
            Key             =   "TipoConcepto"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoFormula.frx":0462
            Key             =   "inicial"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoFormula.frx":08B6
            Key             =   "Concepto"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoFormula.frx":0D0A
            Key             =   "ConceptoA"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoFormula.frx":115E
            Key             =   "+"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoFormula.frx":15B2
            Key             =   "-"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoFormula.frx":1A06
            Key             =   "*"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoFormula.frx":1E5A
            Key             =   ")"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoFormula.frx":7908A
            Key             =   "/"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConceptoFormula.frx":F02BA
            Key             =   "("
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   3525
      Index           =   0
      Left            =   15
      TabIndex        =   8
      Top             =   30
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   6218
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Conceptos"
      TabPicture(0)   =   "FrmManConceptoFormula.frx":1674EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "arb_concepto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.TreeView arb_concepto 
         Height          =   3090
         Left            =   30
         TabIndex        =   9
         Top             =   375
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   5450
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
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
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   3525
      Index           =   1
      Left            =   4110
      TabIndex        =   10
      Top             =   30
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   6218
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Conceptos Usados en la Fórmula"
      TabPicture(0)   =   "FrmManConceptoFormula.frx":167506
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "arb_conceptoformula"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.TreeView arb_conceptoformula 
         Height          =   3090
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   5450
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
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
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   3525
      Index           =   2
      Left            =   8040
      TabIndex        =   12
      Top             =   30
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   6218
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Operadores Matemáticos"
      TabPicture(0)   =   "FrmManConceptoFormula.frx":167522
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "arb_operador"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.TreeView arb_operador 
         Height          =   3090
         Left            =   30
         TabIndex        =   13
         Top             =   375
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   5450
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
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
   End
End
Attribute VB_Name = "FrmManConceptoFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'node.key=Codigo de concepto
'node.tag=nombre variable de concepto
Dim RstConcepto As New ADODB.Recordset
Dim VarsConcepto() As String, CodsConcepto() As String, CantVars As Integer
Dim NDx As Node
Dim i As Integer


Private Sub pCargarDatos()
    Dim NomTipoConcepto As String, CodTipoconcepto As String, NomTipoConcepto_Old As String
    
    Dim NomCatConcepto As String, CodCatConcepto As String, NomCatConcepto_Old As String
    
    Dim NomConcepto As String, CodConcepto As String
    
    Dim nSQL As String
    '--Todos los conceptos
    arb_concepto.Nodes.Clear
    Set arb_concepto.ImageList = Me.ImageList1
    
    Set NDx = arb_concepto.Nodes.Add()
    NDx.Text = "Tipos de Conceptos"
    NDx.key = "tc"
    NDx.Root = "Tipos de Conceptos"
    NDx.Image = "inicial"
    Set NDx = Nothing
    '--Concepto para formulas
    arb_conceptoformula.Nodes.Clear
    Set arb_conceptoformula.ImageList = Me.ImageList1
    Set NDx = arb_conceptoformula.Nodes.Add()
    NDx.Text = "Conceptos Usados en Fórmula"
    NDx.key = "cf"
    NDx.Image = "inicial"
    Set NDx = Nothing
    '--------------------------------------------------------
    nSQL = "SELECT * FROM " _
        + vbCr + " (SELECT 'A' & pla_conceptocat.id  AS catid, pla_conceptocat.descripcion AS catnombre, 'B' & pla_conceptotipo.id   AS tipid, pla_conceptotipo.descripcion AS tipnombre, 'C' & pla_concepto.id  as codigo ,pla_concepto.id, pla_concepto.descripcion, pla_concepto.variable, pla_concepto.formula " _
        + vbCr + " FROM pla_conceptocat INNER JOIN (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) ON pla_conceptocat.id = pla_conceptotipo.idcat " _
        + vbCr + " WHERE (((pla_concepto.variable) Is Not Null)) " _
        + vbCr + " UNION " _
        + vbCr + " SELECT 'A99' AS catid, 'TIPOS DE HORAS' AS catnombre, 'BB' AS tipid, ' ' AS tipnombre, 'F' & [id] as codigo ,id  , mae_tipohora.descripcion, mae_tipohora.variable, '' AS formula " _
        + vbCr + " FROM mae_tipohora " _
        + vbCr + " WHERE (((mae_tipohora.concepto) = -1)) " _
        + vbCr + " UNION " _
        + vbCr + " SELECT 'B99' AS catid, 'OTROS' AS catnombre, 'CC' AS tipid, '' AS tipnombre, 'G' & [pla_conceptovarios].[id] AS codigo, pla_conceptovarios.id, pla_conceptovarios.descripcion, pla_conceptovarios.variable, pla_conceptovarios.formula " _
        + vbCr + " FROM pla_conceptovarios " _
        + vbCr + " ) AS vw " _
        + vbCr + " ORDER BY vw.catid, vw.tipnombre, vw.descripcion; " _

    RST_Busq RstConcepto, nSQL, xCon
    
    NomTipoConcepto_Old = ""
    NomCatConcepto_Old = ""
    While Not RstConcepto.EOF
        NomTipoConcepto_Old = NomTipoConcepto
        NomTipoConcepto = RstConcepto.Fields("tipnombre")
        CodTipoconcepto = RstConcepto.Fields("tipid")
        
        NomCatConcepto_Old = NomCatConcepto
        NomCatConcepto = RstConcepto.Fields("catnombre")
        CodCatConcepto = RstConcepto.Fields("catid")
        
        If NomCatConcepto_Old <> NomCatConcepto Then   'cambia de categoria se agrega nuevo nodo
            ' ---------para todo el listado de categorias de conceptos
            agregar_nodo arb_concepto, "tc", CodCatConcepto, NomCatConcepto, "TipoConcepto"
            '-----------Conceptos que van a ser usados en la formula
            agregar_nodo arb_conceptoformula, "cf", CodCatConcepto, NomCatConcepto, "TipoConcepto"

        End If
               
        If NomTipoConcepto_Old <> NomTipoConcepto And CodCatConcepto <> "A99" And CodCatConcepto <> "B99" Then 'cambia de tipo de concepto agrega nuevo nodo
            ' ---------para todo el listado de conceptos
            agregar_nodo arb_concepto, CodCatConcepto, CodTipoconcepto, NomTipoConcepto, "TipoConcepto"
            '-----------Conceptos que van a ser usados en la formula
            agregar_nodo arb_conceptoformula, CodCatConcepto, CodTipoconcepto, NomTipoConcepto, "TipoConcepto"
            '------------------------------------
        End If
    
        NomConcepto = RstConcepto.Fields("descripcion")
        CodConcepto = RstConcepto.Fields("codigo")
        
        If CodCatConcepto <> "A99" And CodCatConcepto <> "B99" Then
            agregar_nodo arb_concepto, CodTipoconcepto, CodConcepto, NomConcepto & " NomVar = " & RstConcepto.Fields("variable"), "Concepto", "ConceptoA"
        Else
            agregar_nodo arb_concepto, CodCatConcepto, CodConcepto, NomConcepto & " NomVar = " & RstConcepto.Fields("variable"), "Concepto", "ConceptoA"
        End If
        arb_concepto.Nodes(CodConcepto).Tag = RstConcepto.Fields("variable")
        
        RstConcepto.MoveNext
     Wend
     '--falta cargar los conceptos en el arbol de formulas PENDIENTE!!!!!!!!!!!
    
    '--Llenar los operadores matematicos
    arb_operador.Nodes.Clear
    Set arb_operador.ImageList = Me.ImageList1
    
    Set NDx = arb_operador.Nodes.Add()
    NDx.Text = "Operadores Matemáticos"
    NDx.key = "om"
    NDx.Image = "inicial"
    
    agregar_nodo arb_operador, "om", "+", "Suma (+)", "+"
    agregar_nodo arb_operador, "om", "-", "Resta (-)", "-"
    agregar_nodo arb_operador, "om", "*", "Producto (*)", "*"
    agregar_nodo arb_operador, "om", "/", "División (/)", "/"
    agregar_nodo arb_operador, "om", ")", "Paréntesis de Cierre Derecho", ")"
    agregar_nodo arb_operador, "om", "(", "Paréntesis de Cierre Izquierdo", "("
    
    
    arb_concepto.Nodes("tc").Expanded = True
    arb_conceptoformula.Nodes("cf").Expanded = True
    arb_operador.Nodes("om").Expanded = True
    
    '-----
    txt_formula.Text = ""
    txt_constante.Text = ""
    '-----
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    pCargarDatos
End Sub


Private Sub cmd_Click(Index As Integer)
    Dim xCantvars As Long
    Dim yCantvars As Long
    Select Case Index
        Case 0 '--aceptar
            PonerVar arb_conceptoformula, VarsConcepto, CodsConcepto
            xCantvars = CantVars
            yCantvars = CantVars
            SacarRepetidasArray VarsConcepto, xCantvars     'saca valores repetidos
            CantVars = xCantvars
            SacarRepetidasArray CodsConcepto, yCantvars
            
            'se le pasa la formula2
            FrmManConcepto.txt_formula.Tag = txt_formula.Text
            ' se pone la formula que ve el usuario
            FrmManConcepto.txt_formula.Text = txt_formula.Text
            Unload Me
            
        Case 1 '--salir
            Unload Me
                  
        Case 2 '--limpiar
            txt_formula.Text = ""
            txt_formula.Tag = ""
            Text2.Tag = ""
            Text2.Text = ""
            
        Case 3 '--probar
            PonerVar arb_conceptoformula, VarsConcepto, CodsConcepto
            xCantvars = CantVars
            yCantvars = CantVars
            SacarRepetidasArray VarsConcepto, xCantvars  'saca valores repetidos
            CantVars = xCantvars
            SacarRepetidasArray CodsConcepto, yCantvars
            
            List1.Clear
            If IsNumeric(txt_formula.Text) = False Then
                For i = 0 To CantVars - 1
                    List1.AddItem VarsConcepto(i)
                Next
            End If
            FrmManConceptoFormulaProbar.Show 1
        Case 4 '--agregar constante a formula
            txt_constante_KeyPress 13
    End Select
End Sub

Sub ArbDrag(arb As TreeView, NoKey As String, X As Single, Y As Single)
    'On Error Resume Next
    Set NDx = arb_concepto.HitTest(X, Y)
    
    
    If NDx Is Nothing Then
       arb.SelectedItem.Selected = False
       Exit Sub
    End If
    
    If NDx.Tag = "" Then
       Set NDx = Nothing
       Exit Sub
    End If
    
    If NDx.key = NoKey Then
       Set NDx = Nothing
       Exit Sub
    End If
    
    If NDx.Parent.key = NoKey Then
       Set NDx = Nothing
       Exit Sub
    End If
    'Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RstConcepto = Nothing
    Set FrmManConcepto_formula = Nothing
End Sub

Private Sub txt_constante_Change()
    If txt_constante.Text = "" Then Exit Sub
    If IsNumeric(txt_constante.Text) = False Then
       MsgBox "No es un valor numerico", vbInformation
       txt_constante.Text = ""
    End If
End Sub

Private Sub txt_constante_KeyPress(KeyAscii As Integer)
    If validar_numero(KeyAscii) = False And Chr(KeyAscii) <> "." And KeyAscii <> 13 Then
       MsgBox "Valor no es valido", infor
       KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Text2.Text = Text2.Text & " " & txt_constante.Text
        txt_formula.Text = txt_formula.Text & " " & txt_constante.Text
        txt_constante.Text = ""
    End If
End Sub
Private Sub PonerFormulaReal(Nomvar As String)
    txt_formula.Text = txt_formula.Text & " " & Nomvar
End Sub

Private Sub txt_constante_LostFocus()
    txt_constante.Text = Trim(txt_constante.Text)
End Sub

Sub PonerVar(arb As TreeView, Vars As Variant, Cods As Variant)
    Dim NDx As Node
    Dim i As Integer, J As Integer
    
    ReDim Vars(0)
    ReDim Cods(0)
    J = 1
    For i = 1 To arb.Nodes.Count
         Set NDx = arb.Nodes(i)
         If NDx.key <> "" And NDx.key <> "td" And NDx.key <> "tc" And NDx.key <> "cf" And Left(LCase(NDx.key), 3) <> "tcp" Then
             RstConcepto.MoveFirst
             RstConcepto.Find "codigo='" & NDx.key & "'"
             If RstConcepto.EOF = False Then
                ReDim Preserve Vars(J)
                ReDim Preserve Cods(J)
                Vars(J - 1) = RstConcepto.Fields("variable")
                Cods(J - 1) = RstConcepto.Fields("id")
                J = J + 1
             End If
        End If
    Next
    CantVars = J - 1
End Sub

Private Sub arb_concepto_DragDrop(Source As Control, X As Single, Y As Single)
    On Error Resume Next
    If Source.Name = "arb_conceptoformula" Then
        If arb_conceptoformula.SelectedItem Is Nothing Then Exit Sub
        
        agregar_nodo arb_concepto, arb_conceptoformula.SelectedItem.Parent.key, arb_conceptoformula.SelectedItem.key, arb_conceptoformula.SelectedItem.Text, arb_conceptoformula.SelectedItem.Image, arb_conceptoformula.SelectedItem.SelectedImage
        arb_concepto.Nodes(arb_conceptoformula.SelectedItem.key).Tag = arb_conceptoformula.SelectedItem.Tag
        eliminar_nodo arb_conceptoformula, , arb_conceptoformula.SelectedItem.key
        txt_formula.Text = ""
        Text2.Tag = ""
        Text2.Text = ""
        'arb_concepto.Nodes(arb_conceptoformula.SelectedItem.key).Parent.Expanded = True
    ElseIf Source.Name <> "arb_concepto" Then
        MsgBox "No se acepta este valor", vbInformation
    End If
    Err.Clear
End Sub

Private Sub arb_concepto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 And Shift <> 0 Then Exit Sub
    ArbDrag arb_concepto, "tc", X, Y
End Sub

Private Sub arb_concepto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 0 Then Exit Sub
    If Not NDx Is Nothing Then
       arb_concepto.Drag 1
       arb_concepto.DragIcon = arb_concepto.SelectedItem.CreateDragImage
       Set NDx = Nothing
    End If
End Sub

Private Sub arb_conceptoformula_DragDrop(Source As Control, X As Single, Y As Single)
    'On Error Resume Next 'MEJORAR
    If Source.Name = "arb_concepto" Then
       If UCase(arb_concepto.SelectedItem.key) = "TC" Then Exit Sub
       If arb_concepto.SelectedItem Is Nothing Then Exit Sub
       
        agregar_nodo arb_conceptoformula, arb_concepto.SelectedItem.Parent.key, arb_concepto.SelectedItem.key, arb_concepto.SelectedItem.Text, arb_concepto.SelectedItem.Image, arb_concepto.SelectedItem.SelectedImage
        arb_conceptoformula.Nodes(arb_concepto.SelectedItem.key).Tag = arb_concepto.SelectedItem.Tag
        eliminar_nodo arb_concepto, , arb_concepto.SelectedItem.key
        'txt_formula.Text = ""
        'Text2.Tag = ""
        'Text2.Text = ""
    ElseIf Source.Name <> "arb_conceptoformula" Then
         MsgBox "No se acepta este valor", vbInformation
    End If
    'Err.Clear
End Sub

Private Sub arb_conceptoformula_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then Exit Sub
    ArbDrag arb_conceptoformula, "cf", X, Y
End Sub

Private Sub arb_conceptoformula_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 0 Then Exit Sub
    If Not NDx Is Nothing Then
       arb_conceptoformula.Drag 1
       arb_conceptoformula.DragIcon = arb_conceptoformula.SelectedItem.CreateDragImage
       Set NDx = Nothing
    End If
    Err.Clear
End Sub

Private Sub arb_conceptoformula_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Tag <> "" Then
       
       '----------------Codigo de concepto y nomvariable
       PonerFormulaReal Node.Tag
       
    '   txt_formula.text = txt_formula.text & " " & Node.Tag
       'Text2.text = txt_formula.Tag
    End If
End Sub

Private Sub arb_operador_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Name = "arb_concepto" Or Source.Name = "arb_conceptoformula" Then
       MsgBox "No se acepta este valor", vbInformation
    End If
End Sub

Private Sub arb_operador_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.key <> "om" Then
       txt_formula.Text = txt_formula.Text & " " & Node.key
       Text2.Text = Text2.Text & " " & Node.key
    End If
End Sub




