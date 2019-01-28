VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11640
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   5835
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   8643
            MinWidth        =   8643
            Text            =   "Empresa : "
            TextSave        =   "Empresa : "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Usuario : "
            TextSave        =   "Usuario : "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5145
      Top             =   2115
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":07F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0E30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Seleccionar Empresa"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Agregar Usuario"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cambiar mes de Trabajo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda del sistema"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir del sistema"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu11 
      Caption         =   "&Almacen"
      Begin VB.Menu menu11_11 
         Caption         =   "Maestro de Almacenes"
      End
      Begin VB.Menu menu11_13 
         Caption         =   "Maestro Series por Almacen"
      End
      Begin VB.Menu menu11_4 
         Caption         =   "Maestro de Unidades"
      End
      Begin VB.Menu menu11_8 
         Caption         =   "Maestro de Tipo de Items"
      End
      Begin VB.Menu menu11_1 
         Caption         =   "Maestro de Familias"
      End
      Begin VB.Menu menu11_2 
         Caption         =   "Maestro de Clase"
      End
      Begin VB.Menu menu11_3 
         Caption         =   "Maestro de Sub Clase"
      End
      Begin VB.Menu menu11_5 
         Caption         =   "-"
      End
      Begin VB.Menu menu11_6 
         Caption         =   "Mantenimiento Items de Compra y Venta"
      End
      Begin VB.Menu menu11_10 
         Caption         =   "Actualizar Precios, Stock y Margenes"
      End
      Begin VB.Menu menu11_12 
         Caption         =   "Asignar Items a Almacenes"
      End
      Begin VB.Menu menu11_9 
         Caption         =   "Ingreso/Salidas de Almacen"
      End
      Begin VB.Menu menu11_14 
         Caption         =   "-"
      End
      Begin VB.Menu menu11_15 
         Caption         =   "Consulta Ingreso/Salidas de Almacen"
      End
      Begin VB.Menu menu11_7 
         Caption         =   "Kardex                                "
      End
   End
   Begin VB.Menu menu22 
      Caption         =   "&Compras"
      Begin VB.Menu menu22_1 
         Caption         =   "Maestro de Proveedores"
      End
      Begin VB.Menu menu22_5 
         Caption         =   "-"
      End
      Begin VB.Menu menu22_8 
         Caption         =   "Fijar Precios de Compra a Items "
      End
      Begin VB.Menu menu22_2 
         Caption         =   "Requerimientos & Orden de Compra"
         Begin VB.Menu menu22_2_1 
            Caption         =   "Orden de Requerimiento"
         End
         Begin VB.Menu menu22_2_2 
            Caption         =   "Orden de Cotizacion"
         End
         Begin VB.Menu menu22_2_3 
            Caption         =   "-"
         End
         Begin VB.Menu menu22_2_4 
            Caption         =   "Orden de Compra"
         End
      End
      Begin VB.Menu menu22_3 
         Caption         =   "Registrar Compras"
      End
      Begin VB.Menu menu22_9 
         Caption         =   "Renta de 4ta Categoria"
      End
      Begin VB.Menu menu22_4 
         Caption         =   "Gastos por Centro de Costo"
         Visible         =   0   'False
      End
      Begin VB.Menu menu22_11 
         Caption         =   "Gastos Reembolsables"
      End
      Begin VB.Menu menu22_6 
         Caption         =   "-"
      End
      Begin VB.Menu menu22_7 
         Caption         =   "Consulta de Compras"
      End
      Begin VB.Menu menu22_10 
         Caption         =   "Consulta de Honorarios"
      End
   End
   Begin VB.Menu menu33 
      Caption         =   "&Ventas"
      Begin VB.Menu menu33_1 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu menu33_2 
         Caption         =   "Maestro Puntos de Venta del Cliente"
      End
      Begin VB.Menu menu33_12 
         Caption         =   "Maestro Vendedores"
      End
      Begin VB.Menu menu33_13 
         Caption         =   "Maestro de Productos CEN"
      End
      Begin VB.Menu menu33_15 
         Caption         =   "Datos de Transporte"
         Begin VB.Menu menu33_15_1 
            Caption         =   "Maestro Empresas de Transporte"
         End
         Begin VB.Menu menu33_15_2 
            Caption         =   "Maestro de Choferes"
         End
         Begin VB.Menu menu33_15_3 
            Caption         =   "Maestro de Unidades de Transporte"
         End
         Begin VB.Menu menu33_15_4 
            Caption         =   "Maestro Motivos de Traslado"
         End
      End
      Begin VB.Menu menu33_17 
         Caption         =   "Maestro de Concepto NC & ND"
      End
      Begin VB.Menu menu33_3 
         Caption         =   "-"
      End
      Begin VB.Menu menu33_4 
         Caption         =   "Cotizaciones"
      End
      Begin VB.Menu menu33_14 
         Caption         =   "Teleprocesos CEN"
         Begin VB.Menu menu33_14_1 
            Caption         =   "Levantar Pedidos               "
         End
         Begin VB.Menu menu33_14_2 
            Caption         =   "Procesar Pedidos"
         End
      End
      Begin VB.Menu menu33_19 
         Caption         =   "-"
      End
      Begin VB.Menu menu33_20 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu menu33_21 
         Caption         =   "Cronograma de Entregas"
      End
      Begin VB.Menu menu33_22 
         Caption         =   "Reporte de Pedidos"
      End
      Begin VB.Menu menu33_23 
         Caption         =   "-"
      End
      Begin VB.Menu menu33_5 
         Caption         =   "Guias de Remision"
      End
      Begin VB.Menu menu33_6 
         Caption         =   "Registrar Ventas"
      End
      Begin VB.Menu menu33_7 
         Caption         =   "Liquidacion Gasto Debito"
      End
      Begin VB.Menu menu33_16 
         Caption         =   "-"
      End
      Begin VB.Menu menu33_11 
         Caption         =   "Consulta de Ventas"
      End
   End
   Begin VB.Menu menu44 
      Caption         =   "C&ontabilidad"
      Begin VB.Menu menu44_1 
         Caption         =   "Maestro de Detracciones"
      End
      Begin VB.Menu menu44_2 
         Caption         =   "Maestro de Percepciones"
      End
      Begin VB.Menu menu44_3 
         Caption         =   "Maestro de Retenciones"
      End
      Begin VB.Menu menu44_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu44_7 
         Caption         =   "Maestro Libros Contables"
      End
      Begin VB.Menu menu44_9 
         Caption         =   "Maestro de Documentos Contables"
      End
      Begin VB.Menu menu44_23 
         Caption         =   "Maestro de Impuestos"
      End
      Begin VB.Menu menu44_10 
         Caption         =   "Asignar Ctas. Contables a Documentos"
      End
      Begin VB.Menu menu44_8 
         Caption         =   "Plan de Cuentas"
      End
      Begin VB.Menu menu44_22 
         Caption         =   "Centros de Costos"
         Begin VB.Menu menu44_22_1 
            Caption         =   "Mantenimiento de Centro de Costos"
         End
         Begin VB.Menu menu44_22_2 
            Caption         =   "Asignar Centro de Costos a Areas"
         End
      End
      Begin VB.Menu menu44_24 
         Caption         =   "Tipo de Cambio"
      End
      Begin VB.Menu menu44_28 
         Caption         =   "Codigo Unico Sunat"
         Begin VB.Menu menu44_28_1 
            Caption         =   "Mantenimiento de Codificacion"
         End
         Begin VB.Menu menu44_28_2 
            Caption         =   "Mantenimieto de Formatos"
         End
      End
      Begin VB.Menu menu44_30 
         Caption         =   "Configuracion de Estado Financieros"
         Begin VB.Menu menu44_30_1 
            Caption         =   "Mantenimiento de Conceptos"
         End
         Begin VB.Menu menu44_30_2 
            Caption         =   "Mantenimiento de Balances, Estados Financieros y Otros"
         End
      End
      Begin VB.Menu menu44_11 
         Caption         =   "-"
      End
      Begin VB.Menu menu44_14 
         Caption         =   "Detracciones"
         Begin VB.Menu menu44_14_1 
            Caption         =   "Compras                                  "
         End
         Begin VB.Menu menu44_14_2 
            Caption         =   "Ventas"
         End
      End
      Begin VB.Menu menu44_15 
         Caption         =   "Percepcion"
      End
      Begin VB.Menu menu44_16 
         Caption         =   "Retencion"
      End
      Begin VB.Menu menu44_25 
         Caption         =   "Asientos Diversos"
      End
      Begin VB.Menu menu44_17 
         Caption         =   "-"
      End
      Begin VB.Menu menu44_12 
         Caption         =   "Registro de Compras"
      End
      Begin VB.Menu menu44_13 
         Caption         =   "Registro de Ventas"
      End
      Begin VB.Menu menu44_33 
         Caption         =   "Registro de Honorarios"
      End
      Begin VB.Menu menu44_34 
         Caption         =   "Analisis de Cuenta"
      End
      Begin VB.Menu menu44_18 
         Caption         =   "Libro Diario"
      End
      Begin VB.Menu menu44_19 
         Caption         =   "Libro Mayor"
      End
      Begin VB.Menu menu44_27 
         Caption         =   "Libro Bancos"
      End
      Begin VB.Menu menu44_20 
         Caption         =   "Balance de Comprobacion 14 columnas"
      End
      Begin VB.Menu menu44_29 
         Caption         =   "Estados Financieros"
      End
      Begin VB.Menu menu44_21 
         Caption         =   "Gastos x Centro de Costo"
         Visible         =   0   'False
      End
      Begin VB.Menu menu44_26 
         Caption         =   "Kardex valorizado"
      End
      Begin VB.Menu menu44_31 
         Caption         =   "DAOT"
      End
      Begin VB.Menu menu44_32 
         Caption         =   "Centro de Costos"
      End
   End
   Begin VB.Menu menu55 
      Caption         =   "&Tesoreria"
      Begin VB.Menu menu55_1 
         Caption         =   "Maestro Origen"
         Begin VB.Menu menu55_1_1 
            Caption         =   "Ingreso                  "
         End
         Begin VB.Menu menu55_1_2 
            Caption         =   "Egreso"
         End
      End
      Begin VB.Menu menu55_2 
         Caption         =   "Maestro Destino"
         Begin VB.Menu menu55_2_1 
            Caption         =   "Ingreso                  "
         End
         Begin VB.Menu menu55_2_2 
            Caption         =   "Egreso"
         End
      End
      Begin VB.Menu menu55_3 
         Caption         =   "Maestro de Bancos"
      End
      Begin VB.Menu menu55_4 
         Caption         =   "Maestro Cuentas de Banco"
      End
      Begin VB.Menu menu55_19 
         Caption         =   "Asignar Empleados a Tesoreria"
      End
      Begin VB.Menu menu55_13 
         Caption         =   "Conceptos de Abonos y Cargos de Banco"
      End
      Begin VB.Menu menu55_8 
         Caption         =   "Maestro Documentos de Caja y Bancos"
      End
      Begin VB.Menu menu55_5 
         Caption         =   "-"
      End
      Begin VB.Menu menu55_14 
         Caption         =   "Programar Pagos"
      End
      Begin VB.Menu menu55_15 
         Caption         =   "Emitir Cargos a Rendir"
      End
      Begin VB.Menu menu55_16 
         Caption         =   "Redir Cuentas"
      End
      Begin VB.Menu menu55_17 
         Caption         =   "-"
      End
      Begin VB.Menu menu55_6 
         Caption         =   "Ingresos                    "
      End
      Begin VB.Menu menu55_7 
         Caption         =   "Egresos"
      End
      Begin VB.Menu aaaaaa_1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu aaaaaa_2 
         Caption         =   "Imgresos antiguo"
         Visible         =   0   'False
      End
      Begin VB.Menu aaaaaa_3 
         Caption         =   "Egresos antiguo"
         Visible         =   0   'False
      End
      Begin VB.Menu aaaaaa_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu55_11 
         Caption         =   "Canje de Documentos"
      End
      Begin VB.Menu menu55_12 
         Caption         =   "Letras"
         Begin VB.Menu menu55_12_1 
            Caption         =   "Emision de Letras"
         End
         Begin VB.Menu menu55_12_2 
            Caption         =   "Planilla de Cobranza"
         End
      End
      Begin VB.Menu menu55_9 
         Caption         =   "-"
      End
      Begin VB.Menu menu55_10 
         Caption         =   "Analisis de Cliente Proveedor"
      End
      Begin VB.Menu menu55_18 
         Caption         =   "Consultar Programacion de Pagos"
      End
      Begin VB.Menu cccc 
         Caption         =   "Analisi Cta Cte Cliente"
      End
   End
   Begin VB.Menu menu66 
      Caption         =   "&Produccion"
      Begin VB.Menu menu66_1 
         Caption         =   "Mantenimiento de Tareas"
      End
      Begin VB.Menu menu66_2 
         Caption         =   "Mantenimiento de Recetas"
      End
      Begin VB.Menu menu66_7 
         Caption         =   "Mantenimiento de Estacionalidad"
      End
      Begin VB.Menu menu66_14 
         Caption         =   "Mantenimiento de Costo"
      End
      Begin VB.Menu menu66_6 
         Caption         =   "-"
      End
      Begin VB.Menu menu66_4 
         Caption         =   "Programar Producción"
      End
      Begin VB.Menu menu66_3 
         Caption         =   "Orden de Producción"
      End
      Begin VB.Menu menu66_5 
         Caption         =   "Parte de Producción"
         Visible         =   0   'False
      End
      Begin VB.Menu menu66_8 
         Caption         =   "-"
      End
      Begin VB.Menu menu66_10 
         Caption         =   "Grupos de Trabajo"
      End
      Begin VB.Menu menu66_11 
         Caption         =   "Personal de Produccion"
      End
      Begin VB.Menu menu66_12 
         Caption         =   "Tareas"
      End
      Begin VB.Menu menu66_18 
         Caption         =   "Consulta de Tareas"
      End
      Begin VB.Menu menu66_13 
         Caption         =   "-"
      End
      Begin VB.Menu menu66_15 
         Caption         =   "Registro de tareas"
      End
      Begin VB.Menu menu66_16 
         Caption         =   "-"
      End
      Begin VB.Menu menu66_9 
         Caption         =   "Consulta de Producción"
      End
   End
   Begin VB.Menu menu88 
      Caption         =   "P&lanilla"
      Begin VB.Menu menu88_0 
         Caption         =   "Maestros"
         Begin VB.Menu menu88_4 
            Caption         =   "Maestro de Areas"
         End
         Begin VB.Menu menu88_5 
            Caption         =   "Maestro de Cargos"
         End
         Begin VB.Menu menu88_6 
            Caption         =   "Maestro de Conceptos"
         End
         Begin VB.Menu menu88_30 
            Caption         =   "Maestro de Fondo de Pensiones"
         End
         Begin VB.Menu menu88_1 
            Caption         =   "Maestro de Empleados"
         End
      End
      Begin VB.Menu menu88_13 
         Caption         =   "-"
      End
      Begin VB.Menu menu88_19 
         Caption         =   "Maestro de Horarios"
      End
      Begin VB.Menu menu88_23 
         Caption         =   "Maestro de Licencias"
      End
      Begin VB.Menu menu88_25 
         Caption         =   "Maestro de Permisos"
      End
      Begin VB.Menu menu88_15 
         Caption         =   "Programación de dias Festivos"
      End
      Begin VB.Menu menu88_18 
         Caption         =   "-"
      End
      Begin VB.Menu menu88_14_a 
         Caption         =   "Controles"
         Begin VB.Menu menu88_14 
            Caption         =   "Control de Asistencia"
         End
         Begin VB.Menu menu88_24 
            Caption         =   "Control de Licencias"
         End
         Begin VB.Menu menu88_16 
            Caption         =   "Control de Permisos"
         End
         Begin VB.Menu menu88_17 
            Caption         =   "Control de Vacaciones"
         End
      End
      Begin VB.Menu menu88_20 
         Caption         =   "-"
      End
      Begin VB.Menu menu88_29 
         Caption         =   "Resumen de Horas"
      End
      Begin VB.Menu menu88_28 
         Caption         =   "Asignar Sueldo"
      End
      Begin VB.Menu menu88_2 
         Caption         =   "Registro de Boletas de Pago"
      End
      Begin VB.Menu menu88_11 
         Caption         =   "Comisión de Vendedores"
      End
      Begin VB.Menu menu88_36 
         Caption         =   "-"
      End
      Begin VB.Menu menu88_35 
         Caption         =   "Planillas de Producción"
      End
      Begin VB.Menu menu88_37 
         Caption         =   "Consulta de Planilla de Producción"
      End
      Begin VB.Menu menu88_10 
         Caption         =   "-"
      End
      Begin VB.Menu menu88_12 
         Caption         =   "Exportar Datos Sunat"
         Begin VB.Menu menu88_12_1 
            Caption         =   "Exportar Datos del Trabajador"
         End
      End
      Begin VB.Menu menu88_26 
         Caption         =   "-"
      End
      Begin VB.Menu menu88_40 
         Caption         =   "Impresion de Boletas"
      End
      Begin VB.Menu menu88_22 
         Caption         =   "Consulta de Asistencia"
      End
      Begin VB.Menu menu88_27 
         Caption         =   "Consulta de Planillas"
      End
   End
   Begin VB.Menu menuAA 
      Caption         =   "&Mantenimiento"
      Begin VB.Menu menuAA_1 
         Caption         =   "Maestro de Equipos"
      End
      Begin VB.Menu menuAA_5 
         Caption         =   "Maestro Clases de Equipo"
      End
      Begin VB.Menu menuAA_2 
         Caption         =   "Maestro de Areas"
      End
      Begin VB.Menu menuAA_3 
         Caption         =   "-"
      End
      Begin VB.Menu menuAA_4 
         Caption         =   "Orden de Trabajo"
      End
   End
   Begin VB.Menu menu_zz 
      Caption         =   "&Gestion"
      Begin VB.Menu menu_zz_1 
         Caption         =   "Análisis de Compras"
      End
      Begin VB.Menu menu_zz_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_zz_3 
         Caption         =   "Análisis de Ventas"
      End
      Begin VB.Menu menu_zz_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu_zz_5 
         Caption         =   "Análisis de Producción"
      End
      Begin VB.Menu menu_zz_6 
         Caption         =   "-"
      End
      Begin VB.Menu menu_zz_7 
         Caption         =   "Analisis de Tesoreria"
      End
      Begin VB.Menu menu_zz_8 
         Caption         =   "Planeamiento"
         Begin VB.Menu menu_zz_8_2 
            Caption         =   "Proyeccion de Ventas"
         End
         Begin VB.Menu menu_zz_8_1 
            Caption         =   "Plan de Ventas"
         End
         Begin VB.Menu menu_zz_8_3 
            Caption         =   "-"
         End
         Begin VB.Menu menu_zz_8_4 
            Caption         =   "Plan de Produccion"
         End
         Begin VB.Menu menu_zz_8_5 
            Caption         =   "-"
         End
         Begin VB.Menu menu_zz_8_6 
            Caption         =   "Plan de Abastecimiento"
         End
         Begin VB.Menu menu_zz_8_7 
            Caption         =   "-"
         End
         Begin VB.Menu menu_zz_8_8 
            Caption         =   "Plan de Produccion - Unificado"
         End
         Begin VB.Menu menu_zz_8_9 
            Caption         =   "Plan de Abastecimiento - Unificado"
         End
         Begin VB.Menu menu_zz_8_10 
            Caption         =   "-"
         End
         Begin VB.Menu menu_zz_8_11 
            Caption         =   "Produccion Total - Unificado"
         End
      End
   End
   Begin VB.Menu menu10 
      Caption         =   "?"
      Begin VB.Menu menu10_1 
         Caption         =   "Mantenimiento de Empresas"
      End
      Begin VB.Menu menu10_13 
         Caption         =   "Seleccionar Ruta de Acceso"
      End
      Begin VB.Menu menu10_166 
         Caption         =   "Vincular BD"
      End
      Begin VB.Menu menu10_14 
         Caption         =   "Tablas Unificadas"
         Begin VB.Menu menu10_14_1 
            Caption         =   "Mantenimiento Items de Compra y Venta"
         End
         Begin VB.Menu menu10_14_2 
            Caption         =   "Maestro de Unidades"
         End
         Begin VB.Menu menu10_14_3 
            Caption         =   "Maestro de Tipo de Items"
         End
         Begin VB.Menu menu10_14_4 
            Caption         =   "Maestro de Familias"
         End
         Begin VB.Menu menu10_14_5 
            Caption         =   "Maestro de Clase"
         End
         Begin VB.Menu menu10_14_6 
            Caption         =   "Maestro de Sub Clase"
         End
         Begin VB.Menu menu10_14_7 
            Caption         =   "-"
         End
         Begin VB.Menu menu10_14_8 
            Caption         =   "Maestro de Clientes"
         End
         Begin VB.Menu menu10_14_9 
            Caption         =   "Maestro de Proveedores"
         End
      End
      Begin VB.Menu menu10_8 
         Caption         =   "Apertura de Saldos "
         Enabled         =   0   'False
         Begin VB.Menu menu10_8_1 
            Caption         =   "Importar Documentos x Cobrar"
         End
         Begin VB.Menu menu10_8_2 
            Caption         =   "Importar Documentos x Pagar"
         End
      End
      Begin VB.Menu menu10_9 
         Caption         =   "Importacion de Datos"
         Begin VB.Menu menu10_9_1 
            Caption         =   "Clientes"
         End
         Begin VB.Menu menu10_9_2 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu menu10_9_3 
            Caption         =   "Plan de Cuentas"
         End
         Begin VB.Menu menu10_9_4 
            Caption         =   "Centro de Costos"
         End
         Begin VB.Menu menu10_9_5 
            Caption         =   "Items de Almacen"
         End
         Begin VB.Menu menu10_9_6 
            Caption         =   "Compras"
            Begin VB.Menu menu10_9_6_1 
               Caption         =   "Savar"
            End
            Begin VB.Menu menu10_9_6_2 
               Caption         =   "Estudio"
            End
         End
         Begin VB.Menu menu10_9_7 
            Caption         =   "Ventas"
            Begin VB.Menu menu10_9_7_1 
               Caption         =   "Savar"
            End
            Begin VB.Menu menu10_9_7_2 
               Caption         =   "Estudio"
            End
         End
      End
      Begin VB.Menu menu10_5 
         Caption         =   "Setup"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu10_6 
         Caption         =   "Usuarios"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu10_7 
         Caption         =   "Configurar Opciones de Usuario"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu10_15 
         Caption         =   "Transferir Operaciones"
      End
      Begin VB.Menu menu10_10 
         Caption         =   "Plantilla de Impresion"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu10_11 
         Caption         =   "Acceso  de usuarios para Compras"
      End
      Begin VB.Menu menu10_12 
         Caption         =   "Acceso  de usuarios para Produccion"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu10_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu10_45 
         Caption         =   "Corregir Asientos"
      End
      Begin VB.Menu menu10_3 
         Caption         =   "Cerrar Mes"
      End
      Begin VB.Menu menu10_4 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : MDIPRINCIPAL
'* Tipo              : FORMULARIO
'* Descripcion       :
'*
'*
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 01/09/09
'* VERSION           : 1.0
'*****************************************************************************************************

Option Explicit

Private Sub cccc_Click()
    Dim xFrm As New sgi2_cajabancos2.analisis
    xFrm.AnalisisClienteSavar xCon
    Set xFrm = Nothing
End Sub

Private Sub MDIForm_Activate()
    ' SEGUNDO EVENTO A AJECUTARSE DEL FORMULARIO, AQUI SE VALIDA QUE SE HAYA SELECCIONADO UNA EMPRESA
    If SeEjecutoEmp = False Then
        SeEjecutoEmp = True
        ' INFORMA AL SISTEMA SI HARA LOS PROCESOS CONTABLES
        If CONTABILIZAR = True Then
            AP_MESTRA = Val(Format(Date, "mm"))
            ' CARGA EL STATUS BAR DEL SISTEMA CON EM MES DE TRABAJO ACTUAL
            MDIPrincipal.StatusBar1.Panels(4).Text = "Mes : " + NulosC(Format(Date, "mmmm"))
        Else
            ' MUESTRA EN EN STATUS BAR DEL SISTEMA QUE NO SE REALIZARAN LOS PROCESOS CONTABLES
            MDIPrincipal.StatusBar1.Panels(4).Text = "NO CONTABLE"
        End If
    End If
End Sub

Private Sub MDIForm_Load()
    ' PRIMER EVENTO A EJECUTARSE DEL FORMULARIO, SE DEFINEN EL ATO Y ANCHO DEL FORMULARIO Y SU POSICION INCIAL
    ' EN LA PANTALLA
    Me.Width = 12000
    Me.Height = 8600
    Me.Left = 0
    Me.Top = 0
    
    On Error Resume Next                                        ' CONTROLADOR DE ERROR
    Me.Picture = LoadPicture(Trim(AP_RUTABM) + "picchu1.bmp")   ' CARGAMOS EL FONDO DEL FORMULARIO
    Err.Clear
    Me.Caption = AP_NOMSIS                                      ' CARGAMOS EN EL CAPTION DEL FORMULARIO EL NOMBRE DEL SISTEMA
    ActivarMenus                                                ' ACTIVAMOS LOS MENUS DEL SISTEMA
End Sub

Private Sub menu_zz_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_gestion.Compras
    xFrm.AnalizisCompras xCon
    Set xFrm = Nothing
End Sub

Private Sub menu_zz_3_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_gestion.Ventas
    xFrm.AnalizisVentas xCon
    Set xFrm = Nothing
End Sub

Private Sub menu_zz_5_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_gestion.produccion
    xFrm.AnalizisProduccion xCon
    Set xFrm = Nothing
End Sub

Private Sub menu_zz_8_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_gestion.Planeamiento
    xFrm.PlanVentas xCon
    Set xFrm = Nothing
End Sub

Private Sub menu_zz_8_11_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_gestion.Unificados
    xFrm.UnificadoProducido xCon
    Set xFrm = Nothing
End Sub

Private Sub menu_zz_8_2_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_gestion.Planeamiento
    xFrm.PlanVentasEstimado xCon
    Set xFrm = Nothing
End Sub

Private Sub menu_zz_8_4_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_gestion.Planeamiento
    xFrm.PlanProduccion xCon
    Set xFrm = Nothing
End Sub

Private Sub menu_zz_8_6_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_gestion.Planeamiento
    xFrm.PlanAbastecimiento xCon
    Set xFrm = Nothing
End Sub

Private Sub menu_zz_8_8_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_gestion.Unificados
    xFrm.UnificadoProduccion xCon
    Set xFrm = Nothing
End Sub

Private Sub menu_zz_8_9_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_gestion.Unificados
    xFrm.UnificadoAbastecimiento xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_1_Click()
    ' EJECUTA MENU
    FrmMantEmpresa.Show vbModal
End Sub

Private Sub menu10_10_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.ConfiguraImpresion xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_11_Click()
    Dim xFun As New Sgi2_Procesos.Procesos
    xFun.PersonalCompras xCon
    Set xFun = Nothing
End Sub

Private Sub menu10_12_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.PersonalProduccion xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_13_Click()
    ' EJECUTA MENU
    FrmManRutasRutas.Show vbModal
End Sub

Private Sub menu10_14_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New SGI2_almacen.Almacen
    xFrm.MantItem xCon, 2
    Set xFrm = Nothing
End Sub

Private Sub menu10_14_2_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.Almacen
    xFun.ManUnidades xCon, 2
    Set xFun = Nothing
End Sub

Private Sub menu10_14_3_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.Almacen
    xFun.ManTipoProducto xCon, 2
    Set xFun = Nothing
End Sub

Private Sub menu10_14_4_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.Almacen
    xFun.ManFamilia xCon, 2
    Set xFun = Nothing
End Sub

Private Sub menu10_14_5_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.Almacen
    xFun.ManClase xCon, 2
    Set xFun = Nothing
End Sub

Private Sub menu10_14_6_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.Almacen
    xFun.ManSubClase xCon, 2
    Set xFun = Nothing
End Sub

Private Sub menu10_14_8_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.Clientes xCon, 2
    Set xFrm = Nothing
End Sub

Private Sub menu10_14_9_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_compras.Compras
    xFrm.ManProveedor xCon, 2
    Set xFrm = Nothing
End Sub

Private Sub menu10_15_Click()
    ' EJECUTA MENU
    Dim xFun As New Sgi2_Procesos.Procesos
    xFun.TransferenciaOperaciones xCon
    Set xFun = Nothing
End Sub

Private Sub menu10_166_Click()
    ' EJECUTA MENU
    FrmVincularData.Show vbModal
End Sub

Private Sub menu10_3_Click()
    ' EJECUTA MENU
    FrmCierreMes.Show vbModal
End Sub

Private Sub menu10_4_Click()
    ' EJECUTA MENU
    Set xCon = Nothing
    Unload Me
    End
End Sub

Private Sub menu10_45_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.CorregirAsiento xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_5_Click()
    ' EJECUTA MENU
    FrmSetup.Show vbModal
End Sub

Private Sub menu10_6_Click()
    ' EJECUTA MENU
    FrmManUsuarios.Show vbModal
End Sub

Private Sub menu10_7_Click()
    ' EJECUTA MENU
    FrmManOpcionesUsuario.Show vbModal
End Sub

Private Sub menu10_8_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.InventarioCobranza xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_8_2_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.InventarioPagos xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_9_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.CargarClientes xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_9_2_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.CargarProveedores xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_9_3_Click()
    ' EJECUTA MENU
    Dim xFun As New Sgi2_Procesos.Procesos
    xFun.CargarPlandeCuentas xCon
    Set xFun = Nothing
End Sub

Private Sub menu10_9_6_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.CargarCompras xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_9_6_2_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.CargarCompras2 xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_9_7_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.CargarVentas xCon
    Set xFrm = Nothing
End Sub

Private Sub menu10_9_7_2_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.CargarVentasEstudio xCon
    Set xFrm = Nothing
End Sub

Private Sub menu11_1_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.Almacen
    xFun.ManFamilia xCon, 1
    Set xFun = Nothing
End Sub

Private Sub menu11_10_Click()
    ' EJECUTA MENU
    Dim xFrm As New SGI2_almacen.Almacen
    xFrm.ActualizaPrecio xCon
    Set xFrm = Nothing
End Sub

'Function CargarTabla(IdMantenimiento As Integer) As ADODB.Recordset
'    Dim RstCab As New ADODB.Recordset
'
'    RST_Busq RstCab, "SELECT * FROM mae_manformularios WHERE id = " & IdMantenimiento & "", xCon
'    If RstCab.RecordCount = 0 Then
'        Set CargarTabla = Nothing
'    Else
'        Set CargarTabla = RstCab
'    End If
'End Function
'
'Function CargarTablaCampos(IdMantenimiento As Integer) As ADODB.Recordset
'    Dim RstCab As New ADODB.Recordset
'
'    RST_Busq RstCab, "SELECT * FROM mae_manformulariosdet WHERE id = " & IdMantenimiento & " ORDER BY corr", xCon
'    If RstCab.RecordCount = 0 Then
'        Set CargarTablaCampos = Nothing
'    Else
'        Set CargarTablaCampos = RstCab
'    End If
'End Function

Private Sub menu11_11_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT alm_almacenes.* From alm_almacenes ORDER BY alm_almacenes.id"


    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":         xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":    xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "6000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "alm_almacenes"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Alamcenes"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu11_12_Click()
    ' EJECUTA MENU
    Dim xFrm As New SGI2_almacen.Almacen
    xFrm.CargarAlmacenes xCon
    Set xFrm = Nothing
End Sub

Private Sub menu11_13_Click()
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT alm_numseries.*, alm_almacenes.descripcion AS desalm, mae_documento.descripcion AS destipdoc FROM (alm_numseries " _
        & " LEFT JOIN alm_almacenes ON alm_numseries.idalm = alm_almacenes.id) LEFT JOIN mae_documento ON alm_numseries.idtipdoc = mae_documento.id " _
        & " ORDER BY alm_numseries.numser, mae_documento.descripcion"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":              xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "C"
    xCamposVista(1, 0) = "Almacen":             xCamposVista(1, 1) = "desalm":         xCamposVista(1, 2) = "2000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Tipo Documento":      xCamposVista(2, 1) = "destipdoc":      xCamposVista(2, 2) = "3000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Nº Serie":            xCamposVista(3, 1) = "numser":         xCamposVista(3, 2) = "1500":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "C"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":           xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Almacen":          xCampos(1, 1) = "idalm":        xCampos(1, 2) = "N":    xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Tipo Documento":   xCampos(2, 1) = "idtipdoc":     xCampos(2, 2) = "N":    xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Nº Serie":         xCampos(3, 1) = "numser":       xCampos(3, 2) = "C":    xCampos(3, 3) = "1500"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "alm_almacenes":         xVinculos(0, 1) = "id":            xVinculos(0, 2) = "id,descripcion":
    xVinculos(0, 3) = "Codigo,Descripcion":    xVinculos(0, 4) = "1000,3000":     xVinculos(0, 5) = "N,C":
    xVinculos(0, 6) = "idalm":                 xVinculos(0, 7) = "descripcion":   xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "id"
    
    xVinculos(1, 0) = "mae_documento":         xVinculos(1, 1) = "id":            xVinculos(1, 2) = "id,descripcion":
    xVinculos(1, 3) = "Codigo,Descripcion":    xVinculos(1, 4) = "1000,5000":     xVinculos(1, 5) = "N,C":
    xVinculos(1, 6) = "idtipdoc":              xVinculos(1, 7) = "descripcion":   xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "id"

    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "alm_numseries"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Series x Almacen"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu11_15_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_compras.Compras
    xFrm.ConsultaIngSalAlm xCon
    Set xFrm = Nothing
End Sub

Private Sub menu11_2_Click()
    ' EJECUTA MENU
    Dim xform As New SGI2_almacen.Almacen
    xform.ManClase xCon, 1
    Set xform = Nothing
End Sub

Private Sub menu11_3_Click()
    ' EJECUTA MENU
    Dim xform As New SGI2_almacen.Almacen
    xform.ManSubClase xCon, 1
    Set xform = Nothing
End Sub

Private Sub menu11_4_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.Almacen
    xFun.ManUnidades xCon, 1
    Set xFun = Nothing
End Sub

Private Sub menu11_6_Click()
    ' EJECUTA MENU
    Dim xFrm As New SGI2_almacen.Almacen
    xFrm.Idusuario = xIdUsuario
    xFrm.MantItem xCon, 1
    Set xFrm = Nothing
End Sub

Private Sub menu11_7_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.Consultas
    xFrm.MostrarStockResumen xCon, False
    Set xFrm = Nothing
End Sub

Private Sub menu11_8_Click()
    ' EJECUTA MENU
    Dim xFun As New SGI2_almacen.Almacen
    xFun.ManTipoProducto xCon, 1
    Set xFun = Nothing
End Sub

Private Sub menu11_9_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_compras.Compras
    xFrm.Idusuario = xIdUsuario
    xFrm.IngresoAlmacen xCon
    Set xFrm = Nothing
End Sub

Private Sub menu22_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_compras.Compras
    xFrm.Idusuario = xIdUsuario
    xFrm.ManProveedor xCon, 1
    Set xFrm = Nothing
End Sub

Private Sub menu22_10_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_compras.Compras
    xFrm.RepHonorario xCon
    Set xFrm = Nothing
End Sub

Private Sub menu22_11_Click()
'    Dim xFrm As New sgi2_compras.Compras
'    xFrm.Reembolsables xCon
'    Set xFrm = Nothing
End Sub

Private Sub menu22_2_1_Click()
    Dim xFun As New seven_compras2.Compras
    xFun.ManOrdenrequerimiento xCon
    Set xFun = Nothing
End Sub

Private Sub menu22_2_2_Click()
    Dim xFun As New seven_compras2.Compras
    xFun.ManOrdenCotizacion xCon
    Set xFun = Nothing
End Sub

Private Sub menu22_2_4_Click()
    Dim xFun As New seven_compras2.Compras
    xFun.ManOrdenCompra xIdUsuario, xCon
    Set xFun = Nothing
End Sub

Private Sub menu22_2_Click()
    ' EJECUTA MENU
    'Dim xFrm As New sgi2_compras.Compras
    'xFrm.OrdenCompra xCon
    'Set xFrm = Nothing
End Sub

Private Sub menu22_3_Click()
    ' EJECUTA MENU
    Dim xform As New sgi2_compras.Compras
    xform.Idusuario = xIdUsuario
    xform.RegCompras2 xCon, AP_MESTRA, 0
    Set xform = Nothing
End Sub

Private Sub menu22_7_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_compras.Compras
    xFrm.RepCompras xCon
    Set xFrm = Nothing
End Sub

Private Sub menu22_8_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_compras.Compras
    xFrm.AsignaPrecioItem xCon
    Set xFrm = Nothing
End Sub

Private Sub menu22_9_Click()
    ' EJECUTA MENU
    Dim xform As New sgi2_compras.Compras
    xform.Idusuario = xIdUsuario
    xform.RegHonorarios xCon, AP_MESTRA, 0
    Set xform = Nothing
End Sub

Private Sub menu33_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.Idusuario = xIdUsuario
    xFrm.Clientes xCon, 1
    Set xFrm = Nothing
End Sub

Private Sub menu33_10_Click()
    ' EJECUTA MENU
'    Dim xFun As New SGI2_puntoventa.Punto
'    xFun.PuntoVenta xCon, xIdUsuario
'    Set xFun = Nothing
'    Dim xFun As New SGI2_PuntoVenta2.PuntoVenta
'    xFun.PuntoVenta xCon
'    Set xFun = Nothing
End Sub

Private Sub menu33_11_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.ReporteVentas xCon
    Set xFrm = Nothing
End Sub

Private Sub menu33_12_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT vta_vendedores.id, vta_vendedores.basico, vta_vendedores.comision, UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat] +', '+[pla_empleados]![nom]) AS apenom, " _
        & " vta_vendedores.idper FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":              xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Apellidos y Nombres": xCamposVista(1, 1) = "apenom":         xCamposVista(1, 2) = "4000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Basico":              xCamposVista(2, 1) = "basico":         xCamposVista(2, 2) = "1500":   xCamposVista(2, 3) = "N":    xCamposVista(2, 4) = "D"
    xCamposVista(3, 0) = "Comision":            xCamposVista(3, 1) = "comision":       xCamposVista(3, 2) = "1100":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "D"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Vendedor":       xCampos(1, 1) = "idper":        xCampos(1, 2) = "N":    xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Basico":         xCampos(2, 1) = "basico":       xCampos(2, 2) = "N":    xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Comision":       xCampos(3, 1) = "comision":     xCampos(3, 2) = "N":    xCampos(3, 3) = "1000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "pla_empleados":             xVinculos(0, 1) = "id":                xVinculos(0, 2) = "apepat,nom,id":
    xVinculos(0, 3) = "Apellidos,Nombres,Codigo":  xVinculos(0, 4) = "3000,3000,1500":    xVinculos(0, 5) = "C,C,C":
    xVinculos(0, 6) = "idper":                     xVinculos(0, 7) = "apepat":               xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "id"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "apenom"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "vta_vendedores"
    xform.CampoOrdenado = "apenom"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Vendedores"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu33_13_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.MantProdCen xCon
    Set xFrm = Nothing
End Sub

Private Sub menu33_14_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.LevantarPedidos xCon
    Set xFrm = Nothing
End Sub

Private Sub menu33_14_2_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.MuestraPedidos xCon
    Set xFrm = Nothing
End Sub

Private Sub menu33_15_1_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_emptra.* From mae_emptra ORDER BY mae_emptra.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Nº R.U.C.":            xCamposVista(1, 1) = "numruc":         xCamposVista(1, 2) = "1200":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Descripcion":          xCamposVista(2, 1) = "nombre":         xCamposVista(2, 2) = "4000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Direccion":            xCamposVista(3, 1) = "direccion":      xCamposVista(3, 2) = "4000":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Nº R.U.C.":      xCampos(1, 1) = "numruc":       xCampos(1, 2) = "C":    xCampos(1, 3) = "1200"
    xCampos(2, 0) = "Descripcion":    xCampos(2, 1) = "nombre":       xCampos(2, 2) = "C":    xCampos(2, 3) = "5000"
    xCampos(3, 0) = "Direccion":      xCampos(3, 1) = "direccion":    xCampos(3, 2) = "C":    xCampos(3, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "nombre"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_emptra"
    xform.CampoOrdenado = "nombre"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Empresas de Transporte"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu33_15_2_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_chofer.id, mae_chofer.idvehiculo, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS apenom, " _
        & " mae_chofer.numbre, mae_vehiculo.marca, mae_vehiculo.numpla, mae_chofer.categoria, mae_chofer.idper FROM mae_vehiculo RIGHT JOIN (pla_empleados RIGHT JOIN " _
        & " mae_chofer ON pla_empleados.id = mae_chofer.idper) ON mae_vehiculo.id = mae_chofer.idvehiculo " _
        & " ORDER BY UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom]"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Apellido y Nombres":   xCamposVista(0, 1) = "apenom":         xCamposVista(0, 2) = "5000":   xCamposVista(0, 3) = "C":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Nº Brevete":           xCamposVista(1, 1) = "numbre":         xCamposVista(1, 2) = "1200":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Marca":                xCamposVista(2, 1) = "marca":          xCamposVista(2, 2) = "1500":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Nº Placa":             xCamposVista(3, 1) = "numpla":         xCamposVista(3, 2) = "1200":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":          xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Nom. Apellidos":  xCampos(1, 1) = "idper":        xCampos(1, 2) = "N":    xCampos(1, 3) = "3000"
    xCampos(2, 0) = "Nº Brevete":      xCampos(2, 1) = "numbre":       xCampos(2, 2) = "C":    xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Categoria":       xCampos(3, 1) = "categoria":    xCampos(3, 2) = "C":    xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Vehiculo":        xCampos(4, 1) = "idvehiculo":   xCampos(4, 2) = "N":    xCampos(4, 3) = "1500"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "pla_empleados":                            xVinculos(0, 1) = "id":                   xVinculos(0, 2) = "apepat,nom,id":
    xVinculos(0, 3) = "Apell. Pat.,Nombres,Codigo":               xVinculos(0, 4) = "2000,2000,1000":       xVinculos(0, 5) = "C,C,N":
    xVinculos(0, 6) = "idper":                                    xVinculos(0, 7) = "nom":                  xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "apepat"

    xVinculos(1, 0) = "mae_vehiculo":               xVinculos(1, 1) = "id":              xVinculos(1, 2) = "marca,numpla,id":
    xVinculos(1, 3) = "Marca,Nº Placa,Codigo":      xVinculos(1, 4) = "2000,2000,1000":  xVinculos(1, 5) = "C,C,N":
    xVinculos(1, 6) = "idvehiculo":                 xVinculos(1, 7) = "numpla":          xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "marca"

    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "idper"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_chofer"
    xform.CampoOrdenado = "id"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Choferes"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu33_15_3_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(3) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_vehiculo.*  From mae_vehiculo ORDER BY mae_vehiculo.marca"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":          xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1500":   xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Marca":           xCamposVista(1, 1) = "marca":          xCamposVista(1, 2) = "2000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Nº Placa":        xCamposVista(2, 1) = "numpla":         xCamposVista(2, 2) = "2000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":        xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Marca":         xCampos(1, 1) = "marca":        xCampos(1, 2) = "C":    xCampos(1, 3) = "2000"
    xCampos(2, 0) = "Nº Placa":      xCampos(2, 1) = "numpla":       xCampos(2, 2) = "C":    xCampos(2, 3) = "2000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":      xVinculos(0, 1) = "":      xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":      xVinculos(0, 4) = "":      xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":      xVinculos(0, 7) = "":      xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""

    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "marca"
    xCampoBusca(1) = "numpla"
    xCampoBusca(2) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_vehiculo"
    xform.CampoOrdenado = "id"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Vehiculos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu33_15_4_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_mottra.* From mae_mottra ORDER BY mae_mottra.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_mottra"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Motivos de Traslado"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu33_18_Click()
    ' EJECUTA MENU
'    Dim xFrm As New sgi2_ventas.ventas
'    xFrm.Empaques xCon
'    Set xFrm = Nothing
End Sub

Private Sub menu33_17_Click()
    Dim xFun As New sgi2_ventas.Ventas
    xFun.ManConceptoNC_ND xCon
    Set xFun = Nothing
End Sub

Private Sub menu33_2_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.PuntosVenta xCon
    Set xFrm = Nothing
End Sub

Private Sub menu33_20_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_pedidos.Pedidos
    xFrm.Pedidos xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu33_21_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_pedidos.Pedidos
    xFun.MostrarCronogramaEntregas xCon
    Set xFun = Nothing
End Sub

Private Sub menu33_4_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.Cotizaciones xCon
    Set xFrm = Nothing
End Sub

Private Sub menu33_5_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.Idusuario = xIdUsuario
    xFrm.GuiasRemision xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu33_6_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.Idusuario = xIdUsuario
    xFrm.Ventas xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu33_7_Click()
    Dim xFrm As New sgi2_ventas.Ventas
    xFrm.LiqGasDebito xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu44_1_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_detraccion.* From mae_detraccion ORDER BY mae_detraccion.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":         xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":    xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Tasas":          xCamposVista(2, 1) = "tasa":           xCamposVista(2, 2) = "1100":   xCamposVista(2, 3) = "N":    xCamposVista(2, 4) = "D"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "6000"
    xCampos(2, 0) = "Tasa":           xCampos(2, 1) = "tasa":         xCampos(2, 2) = "N":    xCampos(2, 3) = "1100"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_detraccion"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Detracciones"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu44_10_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.mantenimiento
    xFrm.ManCtaDocumento xCon
    xFrm.Idusuario = xIdUsuario
    Set xFrm = Nothing
End Sub

Private Sub menu44_12_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.Consultas
    xFrm.VerRegCompras xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_13_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.Consultas
    xFrm.VerRegVentas xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_14_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.mantenimiento
    xFrm.Idusuario = xIdUsuario
    xFrm.ManDetraccion xCon, AP_MESTRA, DET_Compra
    Set xFrm = Nothing
End Sub

Private Sub menu44_14_2_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.mantenimiento
    xFrm.Idusuario = xIdUsuario
    xFrm.ManDetraccion xCon, AP_MESTRA, DET_Venta
    Set xFrm = Nothing
End Sub

Private Sub menu44_15_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.mantenimiento
    xFrm.Idusuario = xIdUsuario
    xFrm.ManPercepcion xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu44_16_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.mantenimiento
    xFrm.Idusuario = xIdUsuario
    xFrm.ManRetencion xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu44_18_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.Consultas
    xFrm.VerDiario xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_19_Click()
    ' EJECUTA MENU
    Dim xFor As New sgi2_contabilidad.Consultas
    xFor.Mayor xCon
    Set xFor = Nothing
End Sub

Private Sub menu44_2_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(5, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_percepcion.id, mae_percepcion.idcuencom, mae_percepcion.idcuenven, mae_percepcion.descripcion, " _
        & " mae_percepcion.tasa, con_planctas.cuenta AS cuentacom, con_planctas_1.cuenta AS cuentaven, " _
        & " con_planctas.descripcion AS descuecom, con_planctas_1.descripcion AS descueven FROM (mae_percepcion LEFT JOIN con_planctas ON " _
        & " mae_percepcion.idcuencom = con_planctas.id) LEFT JOIN con_planctas AS con_planctas_1 ON " _
        & " mae_percepcion.idcuenven = con_planctas_1.id ORDER BY mae_percepcion.descripcion"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "5000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Tasa":               xCamposVista(2, 1) = "tasa":           xCamposVista(2, 2) = "1100":   xCamposVista(2, 3) = "N":    xCamposVista(2, 4) = "D"
    xCamposVista(3, 0) = "Cta. Compra":        xCamposVista(3, 1) = "cuentacom":      xCamposVista(3, 2) = "1300":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "I"
    xCamposVista(4, 0) = "Cta. Venta":         xCamposVista(4, 1) = "cuentaven":      xCamposVista(4, 2) = "1300":   xCamposVista(4, 3) = "N":    xCamposVista(4, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Tasa":           xCampos(2, 1) = "tasa":         xCampos(2, 2) = "N":    xCampos(2, 3) = "1100"
    xCampos(3, 0) = "Cuenta Compra":  xCampos(3, 1) = "idcuencom":    xCampos(3, 2) = "N":    xCampos(3, 3) = "1100"
    xCampos(4, 0) = "Cuenta Venta":   xCampos(4, 1) = "idcuenven":    xCampos(4, 2) = "N":    xCampos(4, 3) = "1100"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "con_planctas":       xVinculos(0, 1) = "id":          xVinculos(0, 2) = "cuenta,descripcion":
    xVinculos(0, 3) = "Cuenta,Descripcion": xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "C,C":
    xVinculos(0, 6) = "idcuencom":          xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "cuenta"
    
    xVinculos(1, 0) = "con_planctas":       xVinculos(1, 1) = "id":          xVinculos(1, 2) = "cuenta,descripcion":
    xVinculos(1, 3) = "Cuenta,Descripcion": xVinculos(1, 4) = "1100,4000":   xVinculos(1, 5) = "C,C":
    xVinculos(1, 6) = "idcuenven":          xVinculos(1, 7) = "descripcion": xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "cuenta"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_percepcion"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Percepciones"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu44_20_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.Consultas
    xFrm.HojaTrabajo xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_22_1_Click()
    Dim xFrm As New sgi2_contabilidad.mantenimiento
    xFrm.Idusuario = xIdUsuario
    xFrm.ManCentroCostos xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_22_2_Click()
    Dim xFrm As New sgi2_contabilidad2.mantenimientos
    xFrm.ManCentroCostoArea xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_23_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(6, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(6, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_impuestos.*, con_planctas.cuenta, con_planctas.descripcion AS desccuen " _
        & " FROM mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id"

   
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "3000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Abreviatura":        xCamposVista(2, 1) = "abrev":          xCamposVista(2, 2) = "1100":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "C"
    xCamposVista(3, 0) = "Tasa":               xCamposVista(3, 1) = "tasa":           xCamposVista(3, 2) = "1100":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "R"
    xCamposVista(4, 0) = "Cuenta":             xCamposVista(4, 1) = "cuenta":         xCamposVista(4, 2) = "1100":   xCamposVista(4, 3) = "C":    xCamposVista(4, 4) = "I"
    xCamposVista(5, 0) = "Descripcion Cuenta": xCamposVista(5, 1) = "desccuen":       xCamposVista(5, 2) = "3000":   xCamposVista(5, 3) = "C":    xCamposVista(5, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "abrev":        xCampos(2, 2) = "C":    xCampos(2, 3) = "1100"
    xCampos(3, 0) = "Tasa":           xCampos(3, 1) = "tasa":         xCampos(3, 2) = "N":    xCampos(3, 3) = "1100"
    xCampos(4, 0) = "Cuenta Compra":  xCampos(4, 1) = "idcuen":       xCampos(4, 2) = "N":    xCampos(4, 3) = "1100"
    xCampos(5, 0) = "Cuenta Venta":   xCampos(5, 1) = "idcuenvta":    xCampos(5, 2) = "N":    xCampos(5, 3) = "1100"
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "con_planctas":       xVinculos(0, 1) = "id":          xVinculos(0, 2) = "cuenta,descripcion":
    xVinculos(0, 3) = "Cuenta,Descripcion": xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "C,C":
    xVinculos(0, 6) = "idcuen":             xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "cuenta"
    
    xVinculos(1, 0) = "con_planctas":       xVinculos(1, 1) = "id":          xVinculos(1, 2) = "cuenta,descripcion":
    xVinculos(1, 3) = "Cuenta,Descripcion": xVinculos(1, 4) = "1100,4000":   xVinculos(1, 5) = "C,C":
    xVinculos(1, 6) = "idcuenvta":          xVinculos(1, 7) = "descripcion": xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "cuenta"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_impuestos"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Impuestos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu44_24_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad2.mantenimientos
    xFrm.ManTC xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_25_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.mantenimiento
    xFrm.Idusuario = xIdUsuario
    xFrm.ManProviciones xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu44_26_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.Consultas
    xFrm.MostrarStockResumen xCon, True
    Set xFrm = Nothing
End Sub

Private Sub menu44_27_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_cajabancos.cajabancos
    xFrm.Librobancos xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_28_1_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(5, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT var_codificacion.*, var_formatos.formato FROM var_codificacion LEFT JOIN var_formatos ON var_codificacion.iddato = var_formatos.id" _
        & " ORDER BY var_codificacion.orden"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":   xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "C"
    xCamposVista(1, 0) = "Orden":              xCamposVista(1, 1) = "orden":          xCamposVista(1, 2) = "1000":   xCamposVista(1, 3) = "N":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Descripcion":        xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "4000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Formato":            xCamposVista(3, 1) = "formato":        xCamposVista(3, 2) = "1200":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "I"
    xCamposVista(4, 0) = "Activo":             xCamposVista(4, 1) = "activo":         xCamposVista(4, 2) = "1200":   xCamposVista(4, 3) = "N":    xCamposVista(4, 4) = "C"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":          xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "900"
    xCampos(1, 0) = "Descripcion":     xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Formato":         xCampos(2, 1) = "iddato":       xCampos(2, 2) = "N":    xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Orden":           xCampos(3, 1) = "orden":        xCampos(3, 2) = "N":    xCampos(3, 3) = "1000"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
'    xVinculos(0, 0) = "mae_doccajabantipo":  xVinculos(0, 1) = "id":          xVinculos(0, 2) = "id,descripcion":
'    xVinculos(0, 3) = "Codigo,Descripcion":  xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "N,C":
'    xVinculos(0, 6) = "tipo":                xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
'    xVinculos(0, 9) = "id"
    
    xVinculos(0, 0) = "var_formatos":                xVinculos(0, 1) = "id":               xVinculos(0, 2) = "id,descripcion,formato":
    xVinculos(0, 3) = "Codigo,Descripcion,Formato":  xVinculos(0, 4) = "1100,4000,1100":   xVinculos(0, 5) = "N,C,C":
    xVinculos(0, 6) = "iddato":                      xVinculos(0, 7) = "descripcion":      xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "formato"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    
    xform.CadSQLVista = xConsulta
    xform.Tabla = "var_codificacion"
    xform.CampoOrdenado = "orden"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Codificacion"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu44_28_2_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT var_formatos.* FROM var_formatos ORDER BY id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":   xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "C"
    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "4000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Formato":            xCamposVista(2, 1) = "formato":        xCamposVista(2, 2) = "1200":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":          xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "900"
    xCampos(1, 0) = "Descripcion":     xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Formato":         xCampos(2, 1) = "formato":      xCampos(2, 2) = "C":    xCampos(2, 3) = "1200"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
'    xVinculos(0, 0) = "var_formatos":                xVinculos(0, 1) = "id":               xVinculos(0, 2) = "id,descripcion,formato":
'    xVinculos(0, 3) = "Codigo,Descripcion,Formato":  xVinculos(0, 4) = "1100,4000,1100":   xVinculos(0, 5) = "N,C,C":
'    xVinculos(0, 6) = "iddato":                      xVinculos(0, 7) = "descripcion":      xVinculos(0, 8) = "N":
'    xVinculos(0, 9) = "formato"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    
    xform.CadSQLVista = xConsulta
    xform.Tabla = "var_formatos"
    xform.CampoOrdenado = "orden"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Formato para Codificacion"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu44_29_Click()
    ' EJECUTA MENU
'    Dim xFrm As New sgi2_contabilidad2.estadosfinancieros
'    xFrm.BalanceGeneral xCon
'    Set xFrm = Nothing
    Dim xFrm As New sgi2_contabilidad3.mantenimientos
    xFrm.VerInforme xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_3_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(5, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_retencion.id, mae_retencion.idcuencom, mae_retencion.idcuenven, mae_retencion.descripcion, " _
        & " mae_retencion.tasa, mae_retencion.defaul, con_planctas.cuenta AS cuentacom, con_planctas_1.cuenta AS cuentaven, " _
        & " con_planctas.descripcion AS descuecom, con_planctas_1.descripcion AS descueven FROM (mae_retencion LEFT JOIN con_planctas ON " _
        & " mae_retencion.idcuencom = con_planctas.id) LEFT JOIN con_planctas AS con_planctas_1 ON " _
        & " mae_retencion.idcuenven = con_planctas_1.id ORDER BY mae_retencion.descripcion"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "5000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Tasa":               xCamposVista(2, 1) = "tasa":           xCamposVista(2, 2) = "1100":   xCamposVista(2, 3) = "N":    xCamposVista(2, 4) = "D"
    xCamposVista(3, 0) = "Cta. Compra":        xCamposVista(3, 1) = "cuentacom":      xCamposVista(3, 2) = "1300":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "I"
    xCamposVista(4, 0) = "Cta. Venta":         xCamposVista(4, 1) = "cuentaven":      xCamposVista(4, 2) = "1300":   xCamposVista(4, 3) = "N":    xCamposVista(4, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Tasa":           xCampos(2, 1) = "tasa":         xCampos(2, 2) = "N":    xCampos(2, 3) = "1100"
    xCampos(3, 0) = "Cuenta Compra":  xCampos(3, 1) = "idcuencom":    xCampos(3, 2) = "N":    xCampos(3, 3) = "1100"
    xCampos(4, 0) = "Cuenta Venta":   xCampos(4, 1) = "idcuenven":    xCampos(4, 2) = "N":    xCampos(4, 3) = "1100"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "con_planctas":       xVinculos(0, 1) = "id":          xVinculos(0, 2) = "cuenta,descripcion":
    xVinculos(0, 3) = "Cuenta,Descripcion": xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "C,C":
    xVinculos(0, 6) = "idcuencom":          xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "cuenta"
    
    xVinculos(1, 0) = "con_planctas":       xVinculos(1, 1) = "id":          xVinculos(1, 2) = "cuenta,descripcion":
    xVinculos(1, 3) = "Cuenta,Descripcion": xVinculos(1, 4) = "1100,4000":   xVinculos(1, 5) = "C,C":
    xVinculos(1, 6) = "idcuenven":          xVinculos(1, 7) = "descripcion": xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "cuenta"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_retencion"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Retenciones"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu44_30_1_Click()
    ' EJECUTA MENU
'    Dim xFrm As New sgi2_contabilidad2.mantenimientos
'    xFrm.ManBalance xCon
'    Set xFrm = Nothing

    Dim xFrm As New sgi2_contabilidad3.mantenimientos
    xFrm.ManConcepto xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_30_2_Click()
    ' EJECUTA MENU
'    Dim xFrm As New sgi2_contabilidad2.mantenimientos
'    xFrm.ManEstados xCon
'    Set xFrm = Nothing
    Dim xFrm As New sgi2_contabilidad3.mantenimientos
    xFrm.ManInforme xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_31_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.Reportes
    xFun.RepDAOT xCon
    Set xFun = Nothing
End Sub

Private Sub menu44_32_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.Reportes
    xFun.CentroCostos xCon
    Set xFun = Nothing
End Sub

Private Sub menu44_33_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.Consultas
    xFrm.VerRegRent4ta xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_34_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad2.estadosfinancieros
    xFrm.AnalisisCuenta xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_7_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(0, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_libros.* FROM mae_libros"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "D"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "5000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    
'    xVinculos(0, 0) = "mae_impuestos":       xVinculos(0, 1) = "id":            xVinculos(0, 2) = "id,descripcion":
'    xVinculos(0, 3) = "Codigo,Descripcion":  xVinculos(0, 4) = "1000,5000":     xVinculos(0, 5) = "N,C":
'    xVinculos(0, 6) = "idimp":               xVinculos(0, 7) = "descripcion":   xVinculos(0, 8) = "N"
'    xVinculos(0, 9) = "id"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_libros"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Libros"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu44_8_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.mantenimiento
    xFrm.Idusuario = xIdUsuario
    xFrm.ManPlanCuentas xCon
    Set xFrm = Nothing
End Sub

Private Sub menu44_9_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(6, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(6, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_documento.*, mae_impuestos.descripcion AS descimp, mae_impuestos.tasa" _
        & " FROM mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Cod. Sunat":           xCamposVista(1, 1) = "codsun":         xCamposVista(1, 2) = "1100":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "C"
    xCamposVista(2, 0) = "Descripcion":          xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "3000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Abreviatura":          xCamposVista(3, 1) = "abrev":          xCamposVista(3, 2) = "1100":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "C"
    xCamposVista(4, 0) = "Impuesto":             xCamposVista(4, 1) = "descimp":        xCamposVista(4, 2) = "3000":   xCamposVista(4, 3) = "C":    xCamposVista(4, 4) = "I"
    xCamposVista(5, 0) = "Tasa":                 xCamposVista(5, 1) = "tasa":           xCamposVista(5, 2) = "1000":   xCamposVista(5, 3) = "C":    xCamposVista(5, 4) = "D"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "abrev":        xCampos(2, 2) = "C":    xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Cod. Sunat":     xCampos(3, 1) = "codsun":       xCampos(3, 2) = "C":    xCampos(3, 3) = "1200"
    xCampos(4, 0) = "Impuesto":       xCampos(4, 1) = "idimp":        xCampos(4, 2) = "N":    xCampos(4, 3) = "1200"
    xCampos(5, 0) = "Observaciones":  xCampos(5, 1) = "observacion":  xCampos(5, 2) = "M":    xCampos(5, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    
    xVinculos(0, 0) = "mae_impuestos":       xVinculos(0, 1) = "id":            xVinculos(0, 2) = "id,descripcion":
    xVinculos(0, 3) = "Codigo,Descripcion":  xVinculos(0, 4) = "1000,5000":     xVinculos(0, 5) = "N,C":
    xVinculos(0, 6) = "idimp":               xVinculos(0, 7) = "descripcion":   xVinculos(0, 8) = "N"
    xVinculos(0, 9) = "id"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_documento"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Documentos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu55_1_1_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.ManOrigenes 1, xCon
    Set xFun = Nothing

End Sub

Private Sub menu55_1_2_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.ManOrigenes 2, xCon
    Set xFun = Nothing
End Sub

Private Sub menu55_1_Click()
    ' EJECUTA MENU
End Sub

Private Sub menu55_10_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_cajabancos.cajabancos
    xFrm.ConsultaCtaCte xCon
    Set xFrm = Nothing
End Sub

Private Sub menu55_11_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_cajabancos.cajabancos
    xFrm.CanjeDocumentos xCon
    Set xFrm = Nothing
End Sub

Private Sub menu55_12_1_Click()
    ' EJECUTA MENU
    Dim xLet As New sgi2_letras.letras
    xLet.ManLetras AP_MESTRA, xCon
    Set xLet = Nothing
End Sub

Private Sub menu55_12_2_Click()
    ' EJECUTA MENU
'    Dim xLet As New sgi2_letras.letras
'    xLet.ManPlanilla AP_MESTRA, xCon
'    Set xLet = Nothing
End Sub

Private Sub menu55_13_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_bancoscarabo.* From mae_bancoscarabo ORDER BY mae_bancoscarabo.id"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":    xCamposVista(2, 3) = "C":    xCamposVista(1, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_bancoscarabo"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Conceptos Carho y Abono de Bancos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu55_14_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_cajabancos.cajabancos
    xFrm.ProgramarPagos xCon, xIdUsuario, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu55_15_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_cajabancos.cajabancos
    xFrm.EmiRendir xCon, xIdUsuario, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu55_16_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_cajabancos.cajabancos
    xFrm.DevRendir xCon, xIdUsuario, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu55_18_Click()
    ' EJECUTA MENU
'    Dim xFrm As New sgi2_cajabancos.cajabancos
'    xFrm.ConsultaProgramacionPagos xCon, AP_MESTRA
'    Set xFrm = Nothing
    Dim xFrm As New sgi2_cajabancos2.analisis
    xFrm.AnalisisCtaCte xCon
    Set xFrm = Nothing
End Sub

Private Sub menu55_19_Click()
    ' EJECUTA MENU
    Dim xFrm As New Sgi2_Procesos.Procesos
    xFrm.PersonalTesoreria xCon
    Set xFrm = Nothing
End Sub

Private Sub menu55_2_1_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.ManDestinos 1, xCon
    Set xFun = Nothing
End Sub

Private Sub menu55_2_2_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_contabilidad2.mantenimientos
    xFun.ManDestinos 2, xCon
    Set xFun = Nothing
End Sub

Private Sub menu55_2_Click()
    ' EJECUTA MENU
End Sub

Private Sub menu55_3_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_bancos.* From mae_bancos ORDER BY mae_bancos.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":    xCamposVista(2, 3) = "C":    xCamposVista(1, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_bancos"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Bancos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu55_4_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(3, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(6, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_banconumcta.*, mae_bancos.descripcion AS descban, con_planctas.cuenta, " _
        & " con_planctas.descripcion AS desccuen, mae_moneda.simbolo AS moneda " _
        & " FROM (con_planctas RIGHT JOIN (mae_bancos RIGHT JOIN mae_banconumcta ON mae_bancos.id = mae_banconumcta.idban) " _
        & " ON con_planctas.id = mae_banconumcta.idcuen) LEFT JOIN mae_moneda ON mae_banconumcta.idmon = mae_moneda.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":              xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Banco":               xCamposVista(1, 1) = "descban":        xCamposVista(1, 2) = "2000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Nº Cuenta":           xCamposVista(2, 1) = "numcue":         xCamposVista(2, 2) = "1500":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Moneda":              xCamposVista(3, 1) = "moneda":         xCamposVista(3, 2) = "1100":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "C"
    xCamposVista(4, 0) = "Cta. Contable":       xCamposVista(4, 1) = "cuenta":         xCamposVista(4, 2) = "1200":   xCamposVista(4, 3) = "C":    xCamposVista(4, 4) = "I"
    xCamposVista(5, 0) = "Descripcion Cuenta":  xCamposVista(5, 1) = "desccuen":       xCamposVista(5, 2) = "6000":   xCamposVista(5, 3) = "C":    xCamposVista(5, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Banco":          xCampos(1, 1) = "idban":        xCampos(1, 2) = "N":    xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Cuenta":         xCampos(2, 1) = "numcue":       xCampos(2, 2) = "C":    xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Moneda":         xCampos(3, 1) = "idmon":        xCampos(3, 2) = "N":    xCampos(3, 3) = "1000"
    xCampos(4, 0) = "Cuenta":         xCampos(4, 1) = "idcuen":       xCampos(4, 2) = "N":    xCampos(4, 3) = "1500"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "mae_bancos":            xVinculos(0, 1) = "id":            xVinculos(0, 2) = "id,descripcion":
    xVinculos(0, 3) = "Codigo,Descripcion":    xVinculos(0, 4) = "1000,2000":     xVinculos(0, 5) = "N,C":
    xVinculos(0, 6) = "idban":                 xVinculos(0, 7) = "descripcion":   xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "id"
    
    xVinculos(1, 0) = "mae_moneda":            xVinculos(1, 1) = "id":            xVinculos(1, 2) = "id,descripcion":
    xVinculos(1, 3) = "Codigo,Descripcion":    xVinculos(1, 4) = "1000,2000":     xVinculos(1, 5) = "N,C":
    xVinculos(1, 6) = "idmon":                 xVinculos(1, 7) = "descripcion":   xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "id"

    xVinculos(2, 0) = "con_planctas":          xVinculos(2, 1) = "id":            xVinculos(2, 2) = "cuenta,descripcion":
    xVinculos(2, 3) = "Nº Cuenta,Descripcion": xVinculos(2, 4) = "1500,3000":     xVinculos(2, 5) = "C,C":
    xVinculos(2, 6) = "idcuen":                xVinculos(2, 7) = "descripcion":   xVinculos(2, 8) = "N":
    xVinculos(2, 9) = "cuenta"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_banconumcta"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Cuentas de Banco"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu55_6_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_cajabancos.cajabancos
    xFrm.Idusuario = xIdUsuario
    xFrm.IngresoCajaBanco2 xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu55_7_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_cajabancos.cajabancos
    xFrm.Idusuario = xIdUsuario
    xFrm.EgresoCajaBanco2 xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu55_8_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_doccajaban.*, mae_doccajabantipo.descripcion AS desctipo " _
        & " FROM mae_doccajabantipo INNER JOIN mae_doccajaban ON mae_doccajabantipo.id = mae_doccajaban.tipo"
   
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":             xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":   xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":        xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "4000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Abreviatura":        xCamposVista(2, 1) = "abrev":          xCamposVista(2, 2) = "1200":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Tipo":               xCamposVista(3, 1) = "desctipo":       xCamposVista(3, 2) = "1200":   xCamposVista(3, 3) = "N":    xCamposVista(3, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":          xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "900"
    xCampos(1, 0) = "Descripcion":     xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":     xCampos(2, 1) = "abrev":        xCampos(2, 2) = "C":    xCampos(2, 3) = "1300"
    xCampos(3, 0) = "Tipo":            xCampos(3, 1) = "tipo":         xCampos(3, 2) = "N":    xCampos(3, 3) = "1100"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "mae_doccajabantipo":  xVinculos(0, 1) = "id":          xVinculos(0, 2) = "id,descripcion":
    xVinculos(0, 3) = "Codigo,Descripcion":  xVinculos(0, 4) = "1100,4000":   xVinculos(0, 5) = "N,C":
    xVinculos(0, 6) = "tipo":                xVinculos(0, 7) = "descripcion": xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "id"
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_doccajaban"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Documentos Caja y Bancos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu66_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.Idusuario = xIdUsuario
    xFrm.ManTareas xCon
    Set xFrm = Nothing
End Sub

Private Sub menu66_10_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.ManGrupos xCon
    Set xFrm = Nothing
End Sub

Private Sub menu66_11_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.PersonalProduccion xCon
    Set xFrm = Nothing
End Sub

Private Sub menu66_12_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.DistribucionTareas xCon
    Set xFrm = Nothing
End Sub

Private Sub menu66_14_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.Idusuario = xIdUsuario
    xFrm.ConfigurarCosto xCon
    Set xFrm = Nothing
End Sub

Private Sub menu66_15_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.IngresoTareas xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu66_18_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.RepTarea xCon
    Set xFrm = Nothing
End Sub

Private Sub menu66_2_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.Idusuario = xIdUsuario
    xFrm.MamRecetas xCon
    Set xFrm = Nothing
End Sub

Private Sub menu66_3_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.OrdenProduccion xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu66_4_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.ProgramaProduccion xCon, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu66_7_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.Idusuario = xIdUsuario
    xFrm.Estacionalidad xCon
    Set xFrm = Nothing
End Sub

Private Sub menu66_9_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.RepProduccion xCon
    Set xFrm = Nothing
End Sub

Private Sub menu77_1_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_moneda.* From mae_moneda ORDER BY mae_moneda.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "3000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Abreviatura":          xCamposVista(2, 1) = "simbolo":        xCamposVista(2, 2) = "1100":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "C"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "simbolo":      xCampos(2, 2) = "C":    xCampos(2, 3) = "1100"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_moneda"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Monedas"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu77_10_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad.mantenimiento
    xFrm.ManCtaDocumento xCon
    Set xFrm = Nothing
End Sub

Private Sub menu77_2_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_distritos.* From mae_distritos ORDER BY mae_distritos.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":         xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":    xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "6000"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_distritos"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Distritos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu77_3_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_departamentos.* From mae_departamentos ORDER BY mae_departamentos.id"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":         xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":    xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "6000"
        
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":       xVinculos(0, 1) = "":          xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":       xVinculos(0, 4) = "":          xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":       xVinculos(0, 7) = "":          xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_departamentos"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Departamentos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu77_4_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_dociden.* From mae_dociden ORDER BY mae_dociden.id"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Cod. Sunat":           xCamposVista(1, 1) = "iddoc":          xCamposVista(1, 2) = "1100":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "C"
    xCamposVista(2, 0) = "Descripcion":          xCamposVista(2, 1) = "descripcion":    xCamposVista(2, 2) = "3000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Cod. Sunat":     xCampos(2, 1) = "iddoc":        xCampos(2, 2) = "C":    xCampos(2, 3) = "1200"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_dociden"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Documentos de Identidad"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu77_5_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(4, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_condpago.* From mae_condpago ORDER BY mae_condpago.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "3000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Abreviatura":          xCamposVista(2, 1) = "abrev":          xCamposVista(2, 2) = "3000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Nº Dias":              xCamposVista(3, 1) = "numdia":         xCamposVista(3, 2) = "1000":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "D"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "abrev":        xCampos(2, 2) = "C":    xCampos(2, 3) = "3000"
    xCampos(3, 0) = "Nº Dias":        xCampos(3, 1) = "numdia":       xCampos(3, 2) = "C":    xCampos(3, 3) = "1000"
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_condpago"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Condicion de Pago"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu77_7_1_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_emptra.* From mae_emptra ORDER BY mae_emptra.id"
    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Nº R.U.C.":            xCamposVista(1, 1) = "numruc":         xCamposVista(1, 2) = "1200":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Descripcion":          xCamposVista(2, 1) = "nombre":         xCamposVista(2, 2) = "6000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Nº R.U.C.":      xCampos(1, 1) = "numruc":       xCampos(1, 2) = "C":    xCampos(1, 3) = "1200"
    xCampos(2, 0) = "Descripcion":    xCampos(2, 1) = "nombre":       xCampos(2, 2) = "C":    xCampos(2, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_emptra"
    xform.CampoOrdenado = "nombre"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Empresas de Transporte"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu77_7_2_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_mottra.* From mae_mottra ORDER BY mae_mottra.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "900":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_mottra"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Motivos de Traslado"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu77_7_3_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(5, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(4, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_chofer.id, mae_chofer.idvehiculo, UCase(pla_empleados!ape)+', '+pla_empleados!nom AS apenom, mae_chofer.numbre, " _
        & " mae_vehiculo.marca, mae_vehiculo.numpla, mae_chofer.categoria, mae_chofer.idper FROM pla_empleados RIGHT JOIN (mae_vehiculo " _
        & " RIGHT JOIN mae_chofer ON mae_vehiculo.id = mae_chofer.idvehiculo) ON pla_empleados.id = mae_chofer.idper " _
        & " ORDER BY UCase(pla_empleados!ape)+', '+pla_empleados!nom"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Apellido y Nombres":   xCamposVista(0, 1) = "apenom":         xCamposVista(0, 2) = "5000":   xCamposVista(0, 3) = "C":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Nº Brevete":           xCamposVista(1, 1) = "numbre":         xCamposVista(1, 2) = "1200":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Marca":                xCamposVista(2, 1) = "marca":          xCamposVista(2, 2) = "1500":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"
    xCamposVista(3, 0) = "Nº Placa":             xCamposVista(3, 1) = "numpla":         xCamposVista(3, 2) = "1200":   xCamposVista(3, 3) = "C":    xCamposVista(3, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":          xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Nom. Apellidos":  xCampos(1, 1) = "idper":        xCampos(1, 2) = "N":    xCampos(1, 3) = "3000"
    xCampos(2, 0) = "Nº Brevete":      xCampos(2, 1) = "numbre":       xCampos(2, 2) = "C":    xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Categoria":       xCampos(3, 1) = "categoria":    xCampos(3, 2) = "C":    xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Vehiculo":        xCampos(4, 1) = "idvehiculo":   xCampos(4, 2) = "N":    xCampos(4, 3) = "1500"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "pla_empleados":              xVinculos(0, 1) = "id":              xVinculos(0, 2) = "ape,nom,id":
    xVinculos(0, 3) = "Apellidos,Nombres,Codigo":   xVinculos(0, 4) = "2000,2000,1000":  xVinculos(0, 5) = "C,C,N":
    xVinculos(0, 6) = "idper":                      xVinculos(0, 7) = "nom":             xVinculos(0, 8) = "N":
    xVinculos(0, 9) = "ape"

    xVinculos(1, 0) = "mae_vehiculo":               xVinculos(1, 1) = "id":              xVinculos(1, 2) = "marca,numpla,id":
    xVinculos(1, 3) = "Marca,Nº Placa,Codigo":      xVinculos(1, 4) = "2000,2000,1000":  xVinculos(1, 5) = "C,C,N":
    xVinculos(1, 6) = "idvehiculo":                 xVinculos(1, 7) = "numpla":          xVinculos(1, 8) = "N":
    xVinculos(1, 9) = "marca"

    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "idper"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_chofer"
    xform.CampoOrdenado = "id"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Choferes"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon

End Sub

Private Sub menu77_7_4_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(3, 4) As String
    Dim xVinculos(1, 10) As String
    Dim xCampoBusca(3) As String
    Dim xCamposVista(3, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT mae_vehiculo.*  From mae_vehiculo ORDER BY mae_vehiculo.marca"

    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":          xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1500":   xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Marca":           xCamposVista(1, 1) = "marca":          xCamposVista(1, 2) = "2000":   xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    xCamposVista(2, 0) = "Nº Placa":        xCamposVista(2, 1) = "numpla":         xCamposVista(2, 2) = "2000":   xCamposVista(2, 3) = "C":    xCamposVista(2, 4) = "I"

    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":        xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "800"
    xCampos(1, 0) = "Marca":         xCampos(1, 1) = "marca":        xCampos(1, 2) = "C":    xCampos(1, 3) = "2000"
    xCampos(2, 0) = "Nº Placa":      xCampos(2, 1) = "numpla":       xCampos(2, 2) = "C":    xCampos(2, 3) = "2000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "":      xVinculos(0, 1) = "":      xVinculos(0, 2) = "":
    xVinculos(0, 3) = "":      xVinculos(0, 4) = "":      xVinculos(0, 5) = "":
    xVinculos(0, 6) = "":      xVinculos(0, 7) = "":      xVinculos(0, 8) = "":
    xVinculos(0, 9) = ""

    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "marca"
    xCampoBusca(1) = "numpla"
    xCampoBusca(2) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "mae_vehiculo"
    xform.CampoOrdenado = "id"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Vehiculos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu77_9_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_contabilidad2.mantenimientos
    xFrm.ManTC xCon
    Set xFrm = Nothing
End Sub

Private Sub menu88_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas.planillas
    xFrm.ManNomina xCon, AP_RUTDATTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_11_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_planillas3.planillas
    xFun.Comision xCon
    Set xFun = Nothing
End Sub

Private Sub menu88_12_1_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_planillas.planillas
    xFun.ExportarNomina xCon, AP_RUTDATTRA
    Set xFun = Nothing
End Sub

Private Sub menu88_14_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_planillas2.planillas
    xFun.Asistencia xCon, AP_RUTDATTRA, AP_MESTRA
    Set xFun = Nothing
End Sub

Private Sub menu88_15_Click()
    ' EJECUTA MENU
    Dim xFun As New sgi2_planillas2.planillas
    xFun.ManDiasFestivos xCon, AP_RUTDATTRA
    Set xFun = Nothing
End Sub

Private Sub menu88_16_Click()
    ' EJECUTA MENU
'    Dim xFrm As New sgi2_planillas2.planillas
'    xFrm.Permiso xCon, AP_RUTDATTRA, AP_MESTRA
'    Set xFrm = Nothing
End Sub

Private Sub menu88_17_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas2.planillas
    xFrm.Vacaciones xCon, AP_RUTDATTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_19_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas2.planillas
    xFrm.ManHorario xCon, AP_RUTDATTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_2_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas3.planillas
    xFrm.BoletaPago xCon, AP_RUTDATTRA, AP_MESTRA
    Set xFrm = Nothing
End Sub


Private Sub menu88_22_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas2.planillas
    xFrm.ConsAsistencia xCon, AP_RUTDATTRA, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_23_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas2.planillas
    xFrm.ManTipoLicencia xCon, AP_RUTDATTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_24_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas2.planillas
    xFrm.Licencia xCon, AP_RUTDATTRA, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_25_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas2.planillas
    xFrm.ManTipoPermiso xCon, AP_RUTDATTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_27_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas3.planillas
    xFrm.ConsPlanilla xCon, AP_RUTDATTRA, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_28_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas3.planillas
    xFrm.AsignarSueldo xCon, AP_RUTDATTRA, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_29_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas2.planillas
    xFrm.ResumenHoras xCon, AP_RUTDATTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_30_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas3.planillas
    xFrm.ManRegimenPension xCon, AP_RUTDATTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_35_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.CostoProduccion xCon
    Set xFrm = Nothing
End Sub

Private Sub menu88_37_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.RepCosto xCon
    Set xFrm = Nothing
End Sub

Private Sub menu88_4_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT pla_area.* From pla_area ORDER BY pla_area.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":    xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "pla_area"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Areas"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu88_40_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas3.planillas
    xFrm.ImprimirBoletas xCon, AP_RUTDATTRA, AP_MESTRA
    Set xFrm = Nothing
End Sub

Private Sub menu88_5_Click()
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT pla_cargos.* From pla_cargos ORDER BY pla_cargos.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":    xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "pla_cargos"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Cargos"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menu88_6_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_planillas3.planillas
    xFrm.ManConcepto xCon, AP_RUTDATTRA
    Set xFrm = Nothing
End Sub

Private Sub menuAA_1_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_mantenimiento.mantenimiento
    xFrm.MantEquipos xCon
    Set xFrm = Nothing
End Sub

Private Sub menuAA_2_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT pla_area.* From pla_area ORDER BY pla_area.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":    xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "pla_area"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Mantenimiento - Areas"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon
End Sub

Private Sub menuAA_4_Click()
    ' EJECUTA MENU
    Dim xFrm As New sgi2_mantenimiento.mantenimiento
    xFrm.mantenimiento xCon, xIdUsuario
    Set xFrm = Nothing
End Sub

Private Sub menuAA_5_Click()
    ' EJECUTA MENU
    Dim xform As New Eps_MantTablas.mantenimiento
    Dim xNivelUsuario As Integer
    Dim xCampos(2, 4) As String
    Dim xVinculos(2, 10) As String
    Dim xCampoBusca(2) As String
    Dim xCamposVista(2, 4) As String
    Dim xConsulta As String
    
    'SENTENCIA SQL PARA MOSTRAR LA CONSULTA DE LA PESTAÑA CONSULTA
    xConsulta = "SELECT man_equipoclase.* From man_equipoclase ORDER BY man_equipoclase.id"

    
    'CAMPOS PARA LA VISTA DE LA PESTAÑA CONSULTA
    xCamposVista(0, 0) = "Codigo":               xCamposVista(0, 1) = "id":             xCamposVista(0, 2) = "1000":    xCamposVista(0, 3) = "N":    xCamposVista(0, 4) = "I"
    xCamposVista(1, 0) = "Descripcion":          xCamposVista(1, 1) = "descripcion":    xCamposVista(1, 2) = "6000":    xCamposVista(1, 3) = "C":    xCamposVista(1, 4) = "I"
    
    'LISTA DE CAMPOS PARA LA PESTAÑA DETALLE
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "id":           xCampos(0, 2) = "N":    xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "C":    xCampos(1, 3) = "5000"
    
    '0 = nombre de la tabla
    '1 = nombre del campo con el que se iniciara la busqueda
    '2 = lista de campos
    '3 = lista de rotulos para los campos
    '4 = tamaño de los campos
    '5 = tipo de los campos
    '6 = nombre del campo con el que se vincula el array anterior
    '7 = campo que devolvera la busqueda
    '8 = tipo del campo que iniciara la busqueda
    '9 = campo opcional que se mostrara en vez del campo relacionado con la tabla padre
    xVinculos(0, 0) = "": xVinculos(0, 1) = "": xVinculos(0, 2) = "": xVinculos(0, 3) = "": xVinculos(0, 4) = ""
    xVinculos(0, 5) = "": xVinculos(0, 6) = "": xVinculos(0, 7) = "": xVinculos(0, 8) = "": xVinculos(0, 9) = ""
    
    'CAMPOS PARA EFECTUAR LA BUSQUEDA
    xCampoBusca(0) = "descripcion"
    xCampoBusca(1) = "id"
    
    If xNivelUsuario = 0 Then
        xform.PermiteActualiza = False
    Else
        xform.PermiteActualiza = True
    End If
    xform.CadSQLVista = xConsulta
    xform.Tabla = "man_equipoclase"
    xform.CampoOrdenado = "descripcion"
    xform.CampoClave = "Codigo"
    xform.PermiteActualiza = True
    xform.TituloFormulario = "Maestro - Clases de Equipo"
    xform.MantTablas xCampos, xCampoBusca, xVinculos, xCamposVista, xCon

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'EJECUTO LAS OPCIONES LDE TOOLBAR EL FORMULARIO
    If Button.Index = 1 Then
        ' CONECTA  A LA BASE DE DATOS DE ENLACE Y NOS PERMITE SELECCIONAR UNA NUEVA EMPRESA DE TRABAJO
        AbrirDataEnlace
        FrmSelEmp2.Show vbModal
    End If
    
    If Button.Index = 2 Then
        ' LLAMA AL FORMULARIO DE USUARIOS
        'FrmManUsuarios.Show vbModal
    End If
    
    If Button.Index = 3 Then
        ' PERMITE CAMBIAS EL MES DE TRABAJO
        Dim xMesPro As Integer
        xMesPro = AP_MESTRA
        AP_MESTRA = SeleccionaMes(xCon)
        
        If AP_MESTRA = 0 Then
            AP_MESTRA = xMesPro
        End If
        
        ' BUSCAR EL NOMBRE DEL MES DE TRABAJO ACTUAL
        Dim NomMes As String
        NomMes = Busca_Codigo(AP_MESTRA, "id", "descripcion", "con_meses", "N", xCon)
        ' MUESTRA EN EL STATUS BAR DEL FORMULARIO EL NOMBRE DEL MES TRABAJO ACTUAL
        MDIPrincipal.StatusBar1.Panels(4).Text = "Mes : " + Trim(NomMes)
    End If
    
    If Button.Index = 5 Then
        '1 = muestras el contenido
        '2 = muestra el indice
        'PARA MOSTRAR LA AYUDA... DEL SISTEMA
        HtmlHelp ByVal 0&, Trim(App.Path) + "\ayuda\manual.chm", 0, ByVal 0&
        
    End If
    
    If Button.Index = 7 Then
        ' CIERRA LA CONECCION A LA BASE DE DATOS Y PERMITE SALIR DEL SISTEMA
        xCon.Close
        Set xCon = Nothing
        Unload Me
        End
    End If
End Sub

