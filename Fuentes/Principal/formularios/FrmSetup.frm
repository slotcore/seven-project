VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones de Configuracion"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   Icon            =   "FrmSetup.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   30
      TabIndex        =   20
      Top             =   7530
      Width           =   7560
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   3675
         TabIndex        =   27
         Top             =   210
         Width           =   1290
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   2340
         TabIndex        =   21
         Top             =   210
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   30
      TabIndex        =   9
      Top             =   0
      Width           =   7560
      Begin VB.Frame Frame3 
         Caption         =   "Configuracion de Almacenes de Orden de Servicio"
         Height          =   1845
         Left            =   180
         TabIndex        =   42
         Top             =   5490
         Width           =   7245
         Begin VB.Frame Frame4 
            Height          =   1515
            Left            =   5700
            TabIndex        =   44
            Top             =   240
            Width           =   1450
            Begin VB.CommandButton AgregarAlmacenCmd 
               Caption         =   "Agregar"
               Height          =   330
               Left            =   50
               TabIndex        =   46
               Top             =   180
               Width           =   1305
            End
            Begin VB.CommandButton EliminarAlmacenCmd 
               Caption         =   "Eliminar"
               Height          =   330
               Left            =   50
               TabIndex        =   45
               Top             =   600
               Width           =   1305
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   1425
            Index           =   0
            Left            =   90
            TabIndex        =   43
            Top             =   330
            Width           =   5550
            _cx             =   9790
            _cy             =   2514
            _ConvInfo       =   1
            Appearance      =   2
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
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmSetup.frx":030A
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
      End
      Begin VB.CommandButton LimpiarCostosCmd 
         Caption         =   "&Limpiar"
         Height          =   390
         Left            =   2415
         TabIndex        =   41
         Top             =   4920
         Width           =   1530
      End
      Begin MSComCtl2.DTPicker InicioDTPicker 
         Height          =   300
         Left            =   2760
         TabIndex        =   36
         Top             =   4500
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85000193
         CurrentDate     =   41853
      End
      Begin VB.CommandButton VentasCmd 
         Caption         =   "&Generar"
         Height          =   390
         Left            =   5880
         TabIndex        =   35
         Top             =   4440
         Width           =   1530
      End
      Begin VB.CheckBox CkVerificar 
         Alignment       =   1  'Right Justify
         Caption         =   "Verificar Stock al Facturar"
         Height          =   255
         Left            =   390
         TabIndex        =   33
         Top             =   4080
         Width           =   2220
      End
      Begin VB.TextBox TxtAnno 
         Height          =   300
         Left            =   2415
         MaxLength       =   11
         TabIndex        =   31
         Text            =   "TxtAnno"
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton CmdBusTiDocEmp 
         Height          =   240
         Left            =   3045
         Picture         =   "FrmSetup.frx":039F
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   540
         Width           =   240
      End
      Begin VB.CommandButton CmdBusTipPer 
         Height          =   240
         Left            =   3045
         Picture         =   "FrmSetup.frx":04D1
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   225
         Width           =   240
      End
      Begin VB.TextBox TxtTipPer 
         Height          =   300
         Left            =   2415
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "TxtTipPer"
         Top             =   195
         Width           =   900
      End
      Begin VB.CommandButton CmdBusDocIden 
         Height          =   240
         Left            =   3045
         Picture         =   "FrmSetup.frx":0603
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2535
         Width           =   240
      End
      Begin VB.OptionButton OptProConNo 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3240
         TabIndex        =   19
         Top             =   3735
         Width           =   840
      End
      Begin VB.OptionButton OptProConSi 
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2400
         TabIndex        =   18
         Top             =   3735
         Width           =   840
      End
      Begin VB.TextBox TxtRutaData 
         Height          =   300
         Left            =   2415
         TabIndex        =   8
         Text            =   "TxtRutaData"
         Top             =   3330
         Width           =   5000
      End
      Begin VB.TextBox TxtNumDoc 
         Height          =   300
         Left            =   2415
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "TxtNumDoc"
         Top             =   2820
         Width           =   1695
      End
      Begin VB.TextBox TxtIdDoc 
         Height          =   300
         Left            =   2415
         MaxLength       =   1
         TabIndex        =   6
         Text            =   "TxtIdDoc"
         Top             =   2505
         Width           =   900
      End
      Begin VB.TextBox TxtRepLeg 
         Height          =   300
         Left            =   2415
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "TxtRepLeg"
         Top             =   2190
         Width           =   5000
      End
      Begin VB.TextBox TxtDir 
         Height          =   300
         Left            =   2415
         MaxLength       =   100
         TabIndex        =   4
         Text            =   "TxtDir"
         Top             =   1785
         Width           =   5000
      End
      Begin VB.TextBox TxtNomEmp 
         Height          =   300
         Left            =   2415
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "TxtNomEmp"
         Top             =   1470
         Width           =   5000
      End
      Begin VB.TextBox TxtRuc 
         Height          =   300
         Left            =   2415
         MaxLength       =   11
         TabIndex        =   2
         Text            =   "TxtRuc"
         Top             =   1155
         Width           =   1695
      End
      Begin VB.TextBox TxtTipDocEmp 
         Height          =   300
         Left            =   2415
         MaxLength       =   1
         TabIndex        =   1
         Text            =   "TxtTipDocEmp"
         Top             =   510
         Width           =   900
      End
      Begin MSComCtl2.DTPicker FinDTPicker 
         Height          =   300
         Left            =   4440
         TabIndex        =   37
         Top             =   4500
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   115081217
         CurrentDate     =   41853
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Limpiar Procesos de Costos"
         Height          =   195
         Index           =   10
         Left            =   390
         TabIndex        =   40
         Top             =   5040
         Width           =   1950
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "A"
         Height          =   195
         Left            =   4200
         TabIndex        =   39
         Top             =   4560
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "De"
         Height          =   195
         Left            =   2415
         TabIndex        =   38
         Top             =   4560
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Movimientos de Ventas"
         Height          =   195
         Index           =   9
         Left            =   390
         TabIndex        =   34
         Top             =   4560
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   435
         TabIndex        =   32
         Top             =   885
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento "
         Height          =   195
         Index           =   8
         Left            =   435
         TabIndex        =   30
         Top             =   555
         Width           =   1230
      End
      Begin VB.Label LblTipDocEmpresa 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTipDocEmpresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   3375
         TabIndex        =   29
         Top             =   510
         Width           =   4035
      End
      Begin VB.Label LblTipPer 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblTipPer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   3375
         TabIndex        =   26
         Top             =   195
         Width           =   4035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Empresa"
         Height          =   195
         Index           =   7
         Left            =   435
         TabIndex        =   24
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label LblDocIden 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDocIden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   3375
         TabIndex        =   23
         Top             =   2505
         Width           =   4035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Procesar Contablemente"
         Height          =   195
         Index           =   6
         Left            =   435
         TabIndex        =   17
         Top             =   3705
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Index           =   5
         Left            =   435
         TabIndex        =   16
         Top             =   1815
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ruta de Datos"
         Height          =   195
         Index           =   4
         Left            =   435
         TabIndex        =   15
         Top             =   3360
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rep. Legal"
         Height          =   195
         Index           =   3
         Left            =   435
         TabIndex        =   14
         Top             =   2235
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Identidad"
         Height          =   195
         Index           =   2
         Left            =   435
         TabIndex        =   13
         Top             =   2550
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Doc. Identidad"
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   12
         Top             =   2850
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   11
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº R.U.C."
         Height          =   195
         Left            =   435
         TabIndex        =   10
         Top             =   1200
         Width           =   705
      End
   End
End
Attribute VB_Name = "FrmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : FRMSETUP
'* Tipo              : FORMULARIO
'* Descripcion       : FORMULARIO QUE MUESTRA INFORMACION BASICA DE LA EMPRESA ACTUAL, PERMITE TAMBIEN
'*                     MODIFICAR LA INFORMACION.
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 08/07/09
'* VERSION           : 1.0
'*****************************************************************************************************

Option Explicit
Dim Rst As New ADODB.Recordset

'*****************************************************************************************************
'* Nombre Modulo  : Blanquea()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : BLANQUEA LOS CONTROLES DEL FORMULARIO, PARA EL INGRESO DE DATOS
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Blanquea()
    TxtRuc.Text = ""
    TxtNomEmp.Text = ""
    TxtAnno.Text = ""
    TxtDir.Text = ""
    TxtRepLeg.Text = ""
    TxtIdDoc.Text = ""
    TxtNumDoc.Text = ""
    TxtRutaData.Text = ""
    OptProConSi.Value = False
    OptProConNo = False
    CkVerificar.Value = False
End Sub

Private Sub AgregarAlmacenCmd_Click()
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos(2, 4) As String
    Dim nSQLId As String
    Dim cSQL As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    nTitulo = "Buscando Almacenes"
    ' generar la lista de personal para no considerar en la lista
    nSQLId = GENERAR_SQL_ID(fg(0), fg(0).ColIndex("IDALM"), " WHERE alm_almacenes.id", "NOT IN", True)
    
    cSQL = "SELECT alm_almacenes.* FROM alm_almacenes" & nSQLId
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                    "descripcion", "descripcion", Principio, ""
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    fg(0).Rows = fg(0).Rows + 1
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CODIGO")) = UCase(NulosC(xRs("codigo")))
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("DESCRIPCION")) = UCase(NulosC(xRs("descripcion")))
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("IDALM")) = UCase(NulosC(xRs("id")))
    Set xRs = Nothing
End Sub

Private Sub CmdAceptar_Click()
    ' VALIDAMOS QUE LA INFORMACION INGRESADA SE A LA CORRECTA
    If TxtAnno.Text = "" Then
        MsgBox "No ha especificado el Año de trabajo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtAnno.SetFocus
        Exit Sub
    End If
    If TxtRuc.Text = "" Then
        MsgBox "No ha especificado el Nº de R.U.C.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtRuc.SetFocus
        Exit Sub
    End If
    
    If TxtNomEmp.Text = "" Then
        MsgBox "No ha especificado el nombre de la empresa", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNomEmp.SetFocus
        Exit Sub
    End If
    
    If TxtDir.Text = "" Then
        MsgBox "No ha especificado la direccion de la empresa", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDir.SetFocus
        Exit Sub
    End If
    
    If TxtTipPer.Text = "" Then
        MsgBox "No ha especificado el tipo de empresa", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipPer.SetFocus
        Exit Sub
    End If
    
    If TxtRepLeg.Text = "" Then
        MsgBox "No ha especificado el nombre y apellido del representante legal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtRepLeg.SetFocus
        Exit Sub
    End If
    
    If TxtIdDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento de identidad del representante legal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdDoc.SetFocus
        Exit Sub
    End If
    
    If TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado el Nº del documento de identidad del representante legal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Sub
    End If
    
    ' ACTUALIZAMOS LOS DATOS DE LA EMPRESA
    Rst("anotra") = TxtAnno.Text
    Rst("numruc") = TxtRuc.Text
    Rst("nomemp") = TxtNomEmp.Text
    Rst("diremp") = TxtDir.Text
    Rst("idtipper") = NulosN(TxtTipPer.Text)
    Rst("repleg") = TxtRepLeg.Text
    Rst("numdocrepleg") = TxtNumDoc.Text
    Rst("iddocrepleg") = NulosN(TxtIdDoc.Text)
    
    If OptProConSi.Value = True Then
        Rst("procon") = -1
    Else
        Rst("procon") = 0
    End If
    
    Rst("stckvta") = CkVerificar.Value
    Rst.Update
    
    
    ' Se actualizan los almacenes
    ' Limpiar
    Dim F As New SistemaLogica.Funciones
    Dim dataBase As New SistemaData.EDataBase
    Dim A As Integer
    
    Set dataBase.Connection = xCon
    dataBase.BeginTrans
    dataBase.CommandText = "UPDATE alm_almacenes SET vismov = -1"
    dataBase.Execute
    dataBase.CommitTrans
    ' Actualizar
    For A = 1 To fg(0).Rows - 1
        dataBase.ClearParameter
        dataBase.BeginTrans
        dataBase.CommandText = "UPDATE alm_almacenes SET vismov = 0 WHERE id = @id"
        dataBase.AddParameter "@id", adInteger, adParamInput, F.NuloNumeric(fg(0).TextMatrix(A, fg(0).ColIndex("IDALM")))
        dataBase.Execute
        dataBase.CommitTrans
    Next A
    Set dataBase = Nothing
    
    MsgBox "Los datos fueron actualizados con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    CmdCancelar_Click
End Sub

Private Sub CmdBusDocIden_Click()
    ' DESPLEGAMOS LA LISTA DE DOCUMENTOS DE IDENTIDAD PARA EFECTUAR LA BUSQUEDA
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_dociden.* FROM mae_dociden"
    
    xform.Titulo = "Buscando Documento de Identidad"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdDoc.Text = xRs("id")
        LblDocIden = xRs("descripcion")
        TxtNumDoc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTiDocEmp_Click()
    ' DESPLEGAMOS LA LISTA DE TIPO DE DOCUMENTO PARA EFECTUAR LA BUSQUEDA
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_dociden.* FROM mae_dociden"
    
    xform.Titulo = "Buscando Documento de Identidad"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtTipDocEmp.Text = xRs("id")
        LblTipDocEmpresa.Caption = xRs("descripcion")
        TxtNumDoc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipPer_Click()
    ' DESPLEGAMOS LA LISTA DE TIPO DE EMPRESA PARA EFECTUAR LA BUSQUEDA
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_tipoempresa.* FROM mae_tipoempresa"
    
    xform.Titulo = "Buscando Tipo de Empresa"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtTipPer.Text = xRs("id")
        LblTipPer.Caption = xRs("descripcion")
        TxtRepLeg.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub EliminarAlmacenCmd_Click()
    If fg(0).Row < 1 Then
        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        Exit Sub
    End If
    
    If fg(0).Rows = fg(0).FixedRows Then
        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        Exit Sub
    End If
    
    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    fg(0).RemoveItem fg(0).Row
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO QUE SE EJECUTARA AL CARGAR EL FORMULARIO, CARGA LOS DATOS DE LA EMPRESA Y LOS MUESTRA EN PANTALLA
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    
    TxtAnno.Text = NulosC(Rst("anotra"))
    TxtRuc.Text = NulosC(Rst("numruc"))
    TxtNomEmp.Text = NulosC(Rst("nomemp"))
    TxtDir.Text = NulosC(Rst("diremp"))
    TxtRepLeg.Text = NulosC(Rst("repleg"))
    TxtNumDoc.Text = NulosC(Rst("numdocrepleg"))
    TxtTipDocEmp.Text = NulosN(Rst("idtipdoc"))
    TxtTipDocEmp_Validate True
    
    If NulosN(Rst("iddocrepleg")) <> 0 Then
        TxtIdDoc.Text = Rst("iddocrepleg")
        LblDocIden.Caption = Busca_Codigo(Rst("iddocrepleg"), "id", "descripcion", "mae_dociden", "N", xCon)
    Else
        TxtIdDoc.Text = ""
        LblDocIden.Caption = ""
    End If
    
    If NulosN(Rst("idtipper")) <> 0 Then
        TxtTipPer.Text = NulosN(Rst("idtipper"))
        LblTipPer.Caption = Busca_Codigo(Rst("idtipper"), "id", "descripcion", "mae_tipoempresa", "N", xCon)
    Else
        TxtTipPer.Text = ""
        LblTipPer.Caption = ""
    End If
    
    TxtRutaData.Text = NulosC(Rst("ruta"))
    
    If Rst("procon") = -1 Then
        OptProConSi.Value = True
    Else
        OptProConNo.Value = True
    End If
    
    If Rst("stckvta") = True Then
        CkVerificar.Value = 1
    Else
        CkVerificar.Value = 0
    End If
    Dim RstDet As New ADODB.Recordset
    Dim mSQL As String
    ' Se cargan almacenes de ordenes de servicio
    mSQL = "SELECT alm_almacenes.* " _
        + vbCr + "FROM alm_almacenes " _
        + vbCr + "WHERE alm_almacenes.vismov=0"
    
    Set RstDet = Nothing
    RST_Busq RstDet, mSQL, xCon
    
    If RstDet.State = 0 Then Exit Sub
    If RstDet.RecordCount = 0 Then Exit Sub
    fg(0).Rows = fg(0).FixedRows
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        While Not RstDet.EOF
            fg(0).Rows = fg(0).Rows + 1
            fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CODIGO")) = NulosC(RstDet("codigo"))
            fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("DESCRIPCION")) = NulosC(RstDet("descripcion"))
            fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("IDALM")) = NulosN(RstDet("id"))
            RstDet.MoveNext
        Wend
    End If
End Sub

Private Sub LimpiarCostosCmd_Click()
    Dim mDataBase As New SistemaData.EDataBase
    Dim F As New SistemaLogica.Funciones
On Error GoTo BloqueError
      
    If MsgBox("¿Seguro desea borrar los registros de costos del sistema?", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    
    xCon.BeginTrans
    Set mDataBase.Connection = xCon
    ' Se eliminan los registros de kardex
    mDataBase.CommandText = "DELETE FROM con_librocostotemp"
    mDataBase.Execute
    ' Se borran las referencias de movimiento detalle
    mDataBase.ClearParameter
    mDataBase.CommandText = "UPDATE alm_ingresodet SET alm_ingresodet.iddocref = 0"
    mDataBase.Execute
    ' Se confirma la operacion
    xCon.CommitTrans
    MsgBox "Los registros se borraron con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Exit Sub
    
BloqueError:
    xCon.RollbackTrans
    F.MostrarMensajeError Err.Description, "Error al limpiar costos"
End Sub

Private Sub TxtDir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then CmdBusDocIden_Click
End Sub

Private Sub TxtNomEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtRepLeg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtRutaData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipDocEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipDocEmp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTiDocEmp_Click
    End If
End Sub

Private Sub TxtTipDocEmp_Validate(Cancel As Boolean)
    If NulosC(TxtTipDocEmp.Text) <> "" Then
        LblTipDocEmpresa.Caption = Busca_Codigo(TxtTipDocEmp.Text, "id", "descripcion", "mae_dociden", "N", xCon)
        If NulosC(LblTipDocEmpresa.Caption) = "" Then
            TxtTipDocEmp.Text = ""
        End If
    End If
End Sub

Private Sub TxtTipPer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipPer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then CmdBusTipPer_Click
End Sub

Private Sub VentasCmd_Click()
    Dim FPROD As New ProduccionLogica.Funciones
    Dim FechaInicio As Date
    Dim FechaFin As Date
    Dim FechaInicioMovimientos As Date
        
    If MsgBox("Seguro desea generar los Movimientos de Almacén automáticos de Ventas y Guias de Remision", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
    Me.Refresh
    FechaInicio = InicioDTPicker.Value
    FechaFin = FinDTPicker.Value
    FechaInicioMovimientos = CDate("01/01/" & AnoTra)
    Me.MousePointer = vbHourglass
    If FPROD.GenerarMovimientosGuiasVentas(CLng(xIdUsuario), FechaInicio, FechaFin, FechaInicioMovimientos, xCon) Then
        MsgBox "Los movimientos se generaron con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    Me.MousePointer = vbDefault
End Sub
