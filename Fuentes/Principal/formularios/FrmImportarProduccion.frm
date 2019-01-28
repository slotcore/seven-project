VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImportarProduccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Produccion"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   420
      Left            =   1335
      TabIndex        =   4
      Top             =   30
      Width           =   1125
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   90
      TabIndex        =   1
      Top             =   525
      Width           =   5835
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   2
         Top             =   420
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   135
         Width           =   45
      End
      Begin VB.Shape Shape1 
         Height          =   750
         Left            =   45
         Top             =   30
         Width           =   5655
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar"
      Height          =   420
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3930
      TabIndex        =   5
      Top             =   75
      Width           =   45
   End
End
Attribute VB_Name = "FrmImportarProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--IMPORTAR LAS SIGUIENTES TABLAS:
'--PRD_Produccion
'--PRD_Detalle_Produccion
'--PRD_Parte_Produccion
'--ALM_Pedidos
'--ALM_Detalle_Pedido
'--MAE_Unid_Med
'--PRD_Receta
'--
'--AGREGAR UN CAMPO EN TABLA: MAE_Unid_Med :: nombre=>idunimed  ; tipo dato=>numerico
'--EN ESE CAMPO COLOCAR LOS CODIGOS QUE SE RELACIONAN CON
'--MAE_Unid_Med.Descripcion = mae_unidades.descripcion
'--COLOCAR mae_unidades.ID
'---

Private Sub Command1_Click()
    FraProgreso.Visible = True
    If MsgBox("Al importar la data se eliminará la informacion relacionada a la produccion" + vbCr + "Desea continuar???", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    
    Me.Command1.Enabled = False
    Command2.Enabled = False
    Label6.Caption = "Importando..."
    Label1.Caption = "Inicio:" + Format(Time, "hh:mm:ss AM/PM")
    PRO_IMPORTAR
    Label6.Caption = "Terminado"
    Me.Command1.Enabled = True
    Me.Command2.Enabled = True
End Sub

Sub PRO_IMPORTAR()

    Dim TMPRstPro As New ADODB.Recordset   '--PRODUCCION
    Dim TMPRstDet As New ADODB.Recordset   '--DETALLE DE PRODUCCION
    Dim TMPRstIns As New ADODB.Recordset   '--DETALLE DE PRODUCCION CON INSUMOS
    Dim N_SQL As String
    Dim M_ANYO As String '--AÑO DE TRABAJO
    Dim IdPro As Long
    Dim IdDet As Long
    
    AnoTra = "2007"
    If AnoTra = "" Then
        MsgBox "Seleccione el año de trabajo"
    End If
    M_ANYO = AnoTra
    
    On Error GoTo error
    '--AGREGAR LOS PRODUCTOS  TABLA:
    N_SQL = "SELECT PRD_Parte_Produccion.Fecha " _
            + vbCr + " FROM PRD_Parte_Produccion " _
            + vbCr + " GROUP BY Year([Fecha]), PRD_Parte_Produccion.Fecha " _
           + vbCr + " Having (((Year([Fecha])) = " + M_ANYO + ")) ORDER BY PRD_Parte_Produccion.Fecha;"
           
    RST_Busq TMPRstPro, N_SQL, xCon
    
    If TMPRstPro.State = 0 Then GoTo salir:
    If TMPRstPro.BOF = True Or TMPRstPro.EOF = True Or TMPRstPro.RecordCount = 0 Then GoTo salir:
    
    Dim RstPro As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim RstIns As New ADODB.Recordset
        
    RST_Busq RstPro, "SELECT TOP 1 * FROM pro_produccion", xCon
    RST_Busq RstDet, "SELECT TOP 1 * FROM pro_producciondet", xCon
    RST_Busq RstIns, "SELECT TOP 1 * FROM pro_producciondetins", xCon '--INSUMOS
    RST_Busq RstTar, "SELECT TOP 1 * FROM pro_producciondettar", xCon '--TAREAS
    
    xCon.BeginTrans
    

    xCon.Execute "DELETE * FROM pro_produccion "


    PgBar.Min = 0
    PgBar.Max = TMPRstPro.RecordCount

    TMPRstPro.MoveFirst
    Do While Not TMPRstPro.EOF
    
        If TMPRstPro.EOF = False Then PgBar.Value = CLng(TMPRstPro.Bookmark)
        
        IdPro = HallaCodigoTabla("pro_produccion", xCon, "id")
        
        RstPro.AddNew
        RstPro.Fields("id") = IdPro
        RstPro.Fields("num") = Format(IdPro, "000000")
        RstPro.Fields("dia") = TMPRstPro.Fields("fecha")
        RstPro.Fields("idprog") = 1
        
        RstPro.Update
        
        '--AGREGAR LOS DETALLES
        N_SQL = "SELECT Val(PRD_Parte_Produccion.Nro_Parte) AS idparte, pro_receta.id AS idrec, alm_inventario.id AS iditem, mae_unidades.id AS idunimed, alm_inventario.descripcion AS descrec, PRD_Parte_Produccion.Cantidad AS canreal, PRD_Produccion.HoraInicio AS horini, PRD_Produccion.HoraFin AS horfin, Left(alm_inventario.descripcion,5) AS nomrec, IIf([nomrec]='PULPA',1,2) AS idres,  mae_unidades.abrev, PRD_Parte_Produccion.Nro_Produccion, PRD_Parte_Produccion.Nro_lote as numlote, PRD_Parte_Produccion.Nro_Parte AS numparte, PRD_Produccion.Cod_Empresa, PRD_Receta.Cod_Receta, PRD_Parte_Produccion.Cod_Item " _
            + vbCr + " FROM (MAE_Unid_Med INNER JOIN mae_unidades ON MAE_Unid_Med.idunimed = mae_unidades.id) INNER JOIN (PRD_Produccion INNER JOIN (((PRD_Parte_Produccion INNER JOIN PRD_Receta ON PRD_Parte_Produccion.Cod_Receta = PRD_Receta.Cod_Receta) INNER JOIN pro_receta ON PRD_Receta.Cod_Receta = pro_receta.codrec) INNER JOIN alm_inventario ON PRD_Receta.Cod_Item = alm_inventario.codpro) ON (PRD_Produccion.Nro_Produccion = PRD_Parte_Produccion.Nro_Produccion) AND (PRD_Produccion.Cod_Empresa = PRD_Parte_Produccion.Cod_Empresa)) ON MAE_Unid_Med.Cod_Unidad = PRD_Parte_Produccion.Cod_Unidad " _
            + vbCr + " WHERE (((Year([Fecha])) = " + M_ANYO + ") And ((PRD_Parte_Produccion.Fecha) = #" + Format(TMPRstPro.Fields("fecha"), "mm/dd/yy") + "#)) " _
            + vbCr + " ORDER BY PRD_Parte_Produccion.Fecha, alm_inventario.descripcion, PRD_Produccion.HoraInicio;"
    
    
        RST_Busq TMPRstDet, N_SQL, xCon
        If TMPRstDet.EOF = False Or TMPRstDet.BOF = False Then TMPRstDet.MoveFirst
        Do While Not TMPRstDet.EOF
            DoEvents
            RstDet.AddNew
            'IdDet = HallaCodigoTabla("pro_producciondet", xCon, "id")
            RstDet.Fields("id") = TMPRstDet.Fields("idparte")
            RstDet.Fields("idpro") = IdPro
            RstDet.Fields("idrec") = NulosN(TMPRstDet.Fields("idrec"))
            RstDet.Fields("iditem") = NulosN(TMPRstDet.Fields("iditem"))
            RstDet.Fields("idunimed") = NulosN(TMPRstDet.Fields("idunimed"))
            RstDet.Fields("cantidad") = NulosN(TMPRstDet.Fields("canreal"))
            If IsNull(TMPRstDet.Fields("horini")) = False Then RstDet.Fields("horini") = TMPRstDet.Fields("horini")
            If IsNull(TMPRstDet.Fields("horfin")) = False Then RstDet.Fields("horfin") = TMPRstDet.Fields("horfin")
            RstDet.Fields("numparte") = Format(NulosN(TMPRstDet.Fields("numparte")), "000000")
            RstDet.Fields("numlote") = Format(NulosN(TMPRstDet.Fields("numlote")), "000000")
            RstDet.Fields("idres") = TMPRstDet.Fields("idres")
            
            RstDet.Update
            
           '---------------------------------------------------------------
            '---------INSUMOS DEL DETALLE DE PRODUCCION --- LOS INSUMOS
            
            N_SQL = "SELECT pro_receta.id AS idrec, alm_inventario.id AS iditem, mae_unidades.id AS idunimed, alm_inventario.descripcion AS insumo, ALM_Detalle_Pedido.Cantidad AS canutil, mae_unidades.abrev " _
                + vbCr + " FROM alm_inventario INNER JOIN (((MAE_Unid_Med INNER JOIN mae_unidades ON MAE_Unid_Med.idunimed = mae_unidades.id) INNER JOIN ((PRD_Parte_Produccion INNER JOIN ALM_Pedidos ON (PRD_Parte_Produccion.Cod_Empresa = ALM_Pedidos.Cod_Empresa) AND (PRD_Parte_Produccion.Nro_Produccion = ALM_Pedidos.Nro_Doc)) INNER JOIN ALM_Detalle_Pedido ON ALM_Pedidos.Nro_Pedido = ALM_Detalle_Pedido.Nro_Pedido) ON MAE_Unid_Med.Cod_Unidad = ALM_Detalle_Pedido.Cod_Unidad) INNER JOIN pro_receta ON ALM_Detalle_Pedido.Cod_Receta = pro_receta.codrec) ON alm_inventario.codpro = ALM_Detalle_Pedido.Cod_Item " _
                + vbCr + " GROUP BY pro_receta.id, alm_inventario.id, mae_unidades.id, PRD_Parte_Produccion.Nro_Produccion, alm_inventario.descripcion, ALM_Detalle_Pedido.Cantidad, mae_unidades.abrev, PRD_Parte_Produccion.Cod_Empresa, PRD_Parte_Produccion.Nro_Produccion, PRD_Parte_Produccion.Fecha, alm_inventario.descripcion, Year([Fecha]), PRD_Parte_Produccion.Nro_Parte, PRD_Parte_Produccion.Cod_Receta, ALM_Detalle_Pedido.Producto " _
                + vbCr + " Having (((PRD_Parte_Produccion.Cod_Empresa) = '" + TMPRstDet.Fields("Cod_Empresa") + "') And ((PRD_Parte_Produccion.Nro_Parte) = '" + TMPRstDet.Fields("numparte") + "') And ((PRD_Parte_Produccion.Cod_Receta) = '" + TMPRstDet.Fields("Cod_Receta") + "') And ((ALM_Detalle_Pedido.Producto) = '" + TMPRstDet.Fields("Cod_Item") + "')) " _
                + vbCr + " ORDER BY PRD_Parte_Produccion.Nro_Produccion, alm_inventario.descripcion, PRD_Parte_Produccion.Fecha, alm_inventario.descripcion;"
        
            RST_Busq TMPRstIns, N_SQL, xCon
            If TMPRstIns.EOF = False Or TMPRstIns.BOF = False Then TMPRstIns.MoveFirst
            Do While Not TMPRstIns.EOF
                DoEvents
                RstIns.AddNew
                
                RstIns.Fields("idparte") = TMPRstDet.Fields("idparte")
                RstIns.Fields("idrec") = NulosN(TMPRstIns.Fields("idrec"))
                RstIns.Fields("iditem") = NulosN(TMPRstIns.Fields("iditem"))
                RstIns.Fields("idunimed") = NulosN(TMPRstIns.Fields("idunimed"))
                RstIns.Fields("canutil") = NulosN(TMPRstIns.Fields("canutil"))
                RstIns.Update
                
                TMPRstIns.MoveNext
            Loop
            
            Set TMPRstIns = Nothing
            '---------------------------------------------------------------
            '---------------------------------------------------------------
            TMPRstDet.MoveNext
        Loop
        
        Set TMPRstDet = Nothing
        
        '--AGREGAR LAS TAREAS
        
        
        TMPRstPro.MoveNext
    
    Loop
    '----ACTUALIZANDO LAS UNIDADES DE RECETA
    xCon.Execute "UPDATE pro_producciondetins INNER JOIN pro_recetains ON (pro_producciondetins.idrec = pro_recetains.idrec) AND (pro_producciondetins.iditem = pro_recetains.iditem) SET pro_producciondetins.canpro = pro_recetains!canpro;"

    xCon.CommitTrans
    MsgBox "El proceso se realizó con éxito", vbInformation + vbOKOnly, xTitulo
    
salir:
    Set TMPRstPro = Nothing
    Set TMPRstIns = Nothing
    Me.Command1.Enabled = True
    Command2.Enabled = True
    Exit Sub
error:
Set TMPRstPro = Nothing
Set TMPRstDet = Nothing
Set TMPRstIns = Nothing
    xCon.RollbackTrans
    SHOW_ERROR "", "", True, "No se pudo guardar el registro por el siguiente motivo :"
    Me.Command1.Enabled = True
    Command2.Enabled = True
    Label6.Caption = ""
End Sub



Private Sub Command2_Click()
Unload Me
End Sub

