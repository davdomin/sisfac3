VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Menu Principal"
   ClientHeight    =   9810
   ClientLeft      =   165
   ClientTop       =   -720
   ClientWidth     =   10440
   Icon            =   "menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1140
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   2011
      ButtonWidth     =   2011
      ButtonHeight    =   1905
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Empleados"
            Key             =   ""
            Description     =   "Empleados"
            Object.ToolTipText     =   "Empleados"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Productos"
            Key             =   ""
            Object.ToolTipText     =   "Productos"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pedidos"
            Key             =   ""
            Description     =   "Pedidos"
            Object.ToolTipText     =   "Pedidos"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Gastos"
            Key             =   ""
            Description     =   "Gastos"
            Object.ToolTipText     =   "Gastos"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reposición"
            Key             =   ""
            Description     =   "Reposición"
            Object.ToolTipText     =   "Reponer productos"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reportes"
            Key             =   ""
            Description     =   "Reportes"
            Object.ToolTipText     =   "Reportes"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cerrar Sesion"
            Key             =   ""
            Object.ToolTipText     =   "Cerrar Sesion"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   9555
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image logo 
      Height          =   4275
      Left            =   3960
      Picture         =   "menu.frx":058A
      Top             =   2520
      Width           =   6735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   50
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "menu.frx":5E280
            Key             =   ""
            Object.Tag             =   "Empleados"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "menu.frx":60082
            Key             =   ""
            Object.Tag             =   "barras"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "menu.frx":61E84
            Key             =   ""
            Object.Tag             =   "pedidos"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "menu.frx":639D6
            Key             =   ""
            Object.Tag             =   "gastos"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "menu.frx":64E58
            Key             =   ""
            Object.Tag             =   "Reponer"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "menu.frx":663AA
            Key             =   ""
            Object.Tag             =   "reportes"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "menu.frx":67EFC
            Key             =   ""
            Object.Tag             =   "SESSION"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   -3720
      Top             =   -3360
      Width           =   24000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   15
      Left            =   2640
      TabIndex        =   1
      Top             =   9840
      Width           =   9375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   1800
      Width           =   7215
   End
   Begin VB.Menu MnuArch 
      Caption         =   "Archivos"
      Begin VB.Menu ConfProduct 
         Caption         =   "Configuracion Productos"
         Begin VB.Menu mnuProductos 
            Caption         =   "Productos"
         End
         Begin VB.Menu MnuTiposP 
            Caption         =   "Tipos"
         End
      End
      Begin VB.Menu ConfEmp 
         Caption         =   "Configuracion Empleados"
         Begin VB.Menu Emplea 
            Caption         =   "Empleados"
         End
         Begin VB.Menu mnuCargadito 
            Caption         =   "Cargos"
         End
      End
      Begin VB.Menu Raya 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuC 
         Caption         =   "Clientes"
      End
      Begin VB.Menu Provee 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu Raya2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu CS 
         Caption         =   "Cerrar Session"
      End
      Begin VB.Menu Salgase 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Pedoid 
      Caption         =   "Pedidos"
      Begin VB.Menu FrmPedidos 
         Caption         =   "Crear Pedidos"
      End
      Begin VB.Menu cp 
         Caption         =   "Crear Proforma"
      End
      Begin VB.Menu EntregarPedidos 
         Caption         =   "Entregar Pedidos"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu cuentaBancaria 
      Caption         =   "Cuenta Bancaria"
      Begin VB.Menu cuentaCorriente 
         Caption         =   "Cuenta Corriente"
      End
   End
   Begin VB.Menu Repo 
      Caption         =   "Reposiciones"
      Begin VB.Menu mnuRepo 
         Caption         =   "Reposicion"
      End
      Begin VB.Menu produ 
         Caption         =   "Ordenes de Compra"
      End
   End
   Begin VB.Menu MnuReporte 
      Caption         =   "Reportes"
      Begin VB.Menu Rpetepe 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu mnuKardex 
         Caption         =   "Kardex"
      End
      Begin VB.Menu mnurptcis 
         Caption         =   "Cierres de caja"
      End
      Begin VB.Menu MnuClient 
         Caption         =   "Clientes"
      End
      Begin VB.Menu Prop 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu ventasMerca 
         Caption         =   "Ventas por mercaderia"
      End
      Begin VB.Menu MnuAlmacen 
         Caption         =   "Inventario"
      End
      Begin VB.Menu MnuInventario 
         Caption         =   "Inventario"
      End
      Begin VB.Menu mnustokcminimo 
         Caption         =   "Stock Minimo"
      End
      Begin VB.Menu mnuProductosRep 
         Caption         =   "Productos"
      End
      Begin VB.Menu RepMer 
         Caption         =   "Reposicion de Mercancia"
      End
      Begin VB.Menu cpp 
         Caption         =   "Cuentas por pagar"
      End
      Begin VB.Menu TPE 
         Caption         =   "Trabajos por Empleados"
         Visible         =   0   'False
      End
      Begin VB.Menu HDV 
         Caption         =   "Historial de Vehiculos"
         Visible         =   0   'False
      End
      Begin VB.Menu REPOGAS 
         Caption         =   "Gastos"
      End
      Begin VB.Menu mnuRepEstadisticos 
         Caption         =   "Estadisticas"
      End
   End
   Begin VB.Menu MnuMantiene 
      Caption         =   "Mantenimiento"
      Begin VB.Menu PlanificarInventario 
         Caption         =   "Planificar Inventario"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu Ini 
         Caption         =   "Inicializar Sistema"
      End
      Begin VB.Menu BG 
         Caption         =   "Borrado General"
      End
      Begin VB.Menu EN 
         Caption         =   "Establecer numero de factura"
      End
      Begin VB.Menu bo 
         Caption         =   "Establecer numero de boleta"
      End
      Begin VB.Menu RESP 
         Caption         =   "Respaldo de la Informacion"
      End
      Begin VB.Menu trans 
         Caption         =   "Transpasar"
      End
      Begin VB.Menu mnuCambioSunat 
         Caption         =   "Cambio SUNAT"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AA_Click()
    Form20.Show vbModal
End Sub

Private Sub AlimentarAPID_Click()
    Form37.Show vbModal
End Sub

Private Sub Banquito_Click()
    Form24.Show vbModal
End Sub

Private Sub BG_Click()
    Datos.Borrado_General
End Sub

Private Sub bo_Click()
Dim rs As New ADODB.Recordset
Dim numero As Integer
    rs.Open "SELECT * FROM Boleta ORDER BY nbol DESC", Conexion
    If Not rs.EOF Then
        numero = rs(0)
    Else
        numero = 1
    End If
    numero = Val(InputBox("Numero de la Boleta realizada", , numero))
    Conexion.Execute "insert into Boleta (nbol) values(" & numero & ")"
    MsgBox "Boleta Establecida"

End Sub

Private Sub CatAlmacen_Click()
    Form3.Show vbModal
End Sub

Private Sub CDC_Click()
    Form44.Show vbModal
End Sub

Private Sub cGastos_Click()
    Form22.Show vbModal
End Sub


Private Sub CP_Click()
    Form26.Show vbModal
End Sub

Private Sub cpp_Click()
    Form33.Show vbModal
End Sub

Private Sub CS_Click()
    Me.Enabled = False
    Form17.Show vbModal
End Sub

Private Sub cuentaCorriente_Click()
form41.Show vbModal
End Sub

Private Sub Ejecutar_Click()
    Form14.Show vbModal
End Sub

Private Sub Emplea_Click()
    Form8.Show vbModal
End Sub

Private Sub EN_Click()
Dim rs As New ADODB.Recordset
Dim numero As Integer
    rs.Open "SELECT * FROM Factura ORDER BY nfac DESC", Conexion
    If Not rs.EOF Then
        numero = rs(0)
    Else
        numero = 1
    End If
    numero = Val(InputBox("Numero de la ultima factura realizada", , numero))
    Conexion.Execute "insert into Factura (nfac) values(" & numero & ")"
    MsgBox "Factura Establecida"

End Sub

Private Sub EntregarPedidos_Click()
    Form15.Show vbModal
End Sub
Sub MostrarMenu()
    If Not Produccion Then
        Me.MnuAlmacen.Visible = False
        Me.EntregarPedidos.Visible = False
'        Me.Toolbar1.Buttons(5).Visible = False
 '       Me.Toolbar1.Buttons(1).Visible = False
        
    End If
    
End Sub

Private Sub Form_Initialize()
   'Label1.Caption = "Software registrado para " & NombreEmpresa
    Me.mnuCambioSunat.Enabled = Proyecto.cambioSunat
    Me.Enabled = False
    Label2.Caption = UsuarioSession
    StatusBar1.Panels(1).Text = "Software registrado para " & NombreEmpresa
    MostrarMenu
End Sub

Private Sub Form_Load()
    'Image1.Left = (Me.ScaleWidth + 1 / 2) - Image1.Width
End Sub

Private Sub FrmPedidos_Click()
    Form10.Show vbModal
End Sub

Private Sub HDV_Click()
    Form36.Show vbModal
End Sub

Private Sub Ini_Click()
    If inicializar Then
        EN_Click
        bo_Click
    End If
End Sub

Private Sub Marcas_Click()
   Form27.Show vbModal
End Sub

Private Sub menuCotizacion_Click()
    Form50.Show vbModal
End Sub

Private Sub mnuAlma_Click()
    Form2.Show vbModal
End Sub

Private Sub MnuAlmacen_Click()
Dim f(1) As String
    f(1) = ""
    Datos.CargarReporte "", App.Path & "\reportes\rptinventario.rpt", f
End Sub

Private Sub MnuC_Click()
    Form7.Show vbModal
End Sub

Private Sub mnuCambioSunat_Click()
Form52.Show
End Sub

Private Sub mnuCargadito_Click()
    Form9.Show vbModal
End Sub

Private Sub MnuClient_Click()
Dim f(1) As String
    f(1) = ""
    Datos.CargarReporte "", App.Path & "\reportes\clientes.rpt", f
End Sub

Private Sub MnuOrden_Click()
    Form12.Show vbModal
End Sub

Private Sub MnuInventario_Click()
Dim f(1) As String
    f(1) = ""
    Datos.CargarReporte "", App.Path & "\reportes\rptinventario.rpt", f

End Sub

Private Sub mnuKardex_Click()
    frmKardex.Show vbModal
End Sub

Private Sub MnuProductos_Click()
    Form5.Show vbModal
End Sub
Sub inventario()

Dim f(1) As String
    f(1) = ""
    Datos.CargarReporte "{Productos.Tipo}<>'Servicio'", App.Path & "\reportes\productos.rpt", f
End Sub

Private Sub mnuProductosRep_Click()
    inventario
End Sub

Private Sub mnuRepEstadisticos_Click()
    Form53.Show vbModal
End Sub

Private Sub mnuRepo_Click()
    Form49.Show vbModal
End Sub

Private Sub mnurptcis_Click()
    Form45.Show vbModal
End Sub

Private Sub mnustokcminimo_Click()
Dim f(1) As String
    f(1) = ""
    Datos.CargarReporte "", App.Path & "\reportes\rptminimo.rpt", f

End Sub

Private Sub MnuTiposP_Click()
    Form6.Show vbModal
End Sub

Private Sub mnuUsuarios_Click()
    Form18.Show vbModal
End Sub

Private Sub Mod_Click()
    Form28.Show vbModal
End Sub

Private Sub NE_Click()
    Form30.Show vbModal
End Sub

Private Sub Notas_Click()
    Form40.Show vbModal
End Sub

Private Sub PlanificarInventario_Click()
    Form48.Show vbModal
End Sub

Private Sub PN_Click()
    Form32.Show vbModal
End Sub

Private Sub produ_Click()
    frmOrdenCompra.Show vbModal
End Sub

Private Sub Prop_Click()
    Form46.Show vbModal
End Sub

Private Sub Provee_Click()
    Form23.Show vbModal
End Sub

Private Sub RepMer_Click()
    Form21.Show vbModal
End Sub

Private Sub REPOGAS_Click()
    Form35.Show vbModal
End Sub

Private Sub RESP_Click()
    Datos.Respaldo
    MsgBox "Respaldada"
End Sub

Private Sub Rpetepe_Click()
    Form16.Show vbModal
End Sub

Private Sub Salgase_Click()
    End
End Sub

Private Sub SD_Click()
    Form31.Show vbModal
End Sub

Private Sub TipoAl_Click()
    Form4.Show vbModal
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Dim Opcion As String
    Opcion = UCase(Button.Caption)
    Select Case Opcion
        Case "ALMACEN"
            Form2.Show vbModal
        Case "PRODUCTOS"
            Form5.Show vbModal
        Case "PEDIDOS"
            Form10.Show vbModal
        Case "ORDEN/TRABAJO"
            Form12.Show vbModal
        Case "REPORTES"
            Form16.Show vbModal
        Case "REPOSICIÓN"
            mnuRepo_Click
        Case "EMPLEADOS"
            Form18.Show vbModal
        Case "ENTREGAS"
            Form30.Show vbModal
        Case "PRESUPUESTO"
            Form26.Show vbModal
        Case "PAG/NOMINA"
            Form32.Show vbModal
        Case "CERRAR SESION"
            CS_Click
    End Select
    
 
End Sub

Private Sub TPE_Click()
    Form34.Show vbModal
End Sub

Private Sub trans_Click()
    BorrarEspacios
    Dim conex2 As New ADODB.Connection
    conex2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\datos.mdb;Persist Security Info=False"
    Dim rs1 As New ADODB.Recordset
    rs1.Open "select *from productos", conex2
    While Not rs1.EOF
        barra = rs1("barraS")
        stock = rs1("stock")
        ALMACEN = rs1("ALMACEN")
        codProducto = rs1(0)
        Conexion.Execute "update productos set barras='" & barra & "',almacen='" & ALMACEN & "',STOCK='" & stock & "' where codProducto=" & codProducto
        
        rs1.MoveNext
    Wend
    MsgBox "listo"
    
End Sub

Private Sub Vehi_Click()
    Form29.Show vbModal
End Sub

Private Sub ventasMerca_Click()
    Form43.Show vbModal
End Sub
