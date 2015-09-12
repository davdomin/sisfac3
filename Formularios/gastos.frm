VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form22 
   Caption         =   "Gastos"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   10215
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form22"
   ScaleHeight     =   8700
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo4 
      Height          =   360
      ItemData        =   "gastos.frx":0000
      Left            =   2520
      List            =   "gastos.frx":0019
      TabIndex        =   47
      Text            =   "Combo3"
      Top             =   3840
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Height          =   360
      ItemData        =   "gastos.frx":0066
      Left            =   120
      List            =   "gastos.frx":0073
      TabIndex        =   45
      Text            =   "Combo3"
      Top             =   3840
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2400
      TabIndex        =   43
      Top             =   6720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112984065
      CurrentDate     =   40822
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9000
      Picture         =   "gastos.frx":009C
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   1920
      Width           =   495
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   120
      TabIndex        =   36
      Top             =   9720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   5640
      TabIndex        =   35
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   8235
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   4920
      TabIndex        =   34
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   8235
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   4200
      TabIndex        =   33
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   8235
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   3480
      TabIndex        =   32
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   8235
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   31
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6720
      TabIndex        =   30
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5520
      TabIndex        =   29
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4200
      TabIndex        =   28
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1800
      TabIndex        =   26
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   600
      TabIndex        =   25
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   120
      TabIndex        =   24
      Top             =   9000
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox Text11 
      Height          =   360
      Left            =   6720
      TabIndex        =   21
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Height          =   360
      Left            =   6720
      TabIndex        =   19
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Height          =   360
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   8895
   End
   Begin VB.TextBox Text8 
      Height          =   360
      Left            =   7680
      TabIndex        =   15
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   360
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   6615
   End
   Begin VB.TextBox Text6 
      Height          =   360
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   8895
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   2400
      TabIndex        =   9
      Top             =   1920
      Width           =   6615
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo de Pago"
      Height          =   255
      Left            =   2520
      TabIndex        =   46
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pago"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Prox. Fecha de Pago"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "*Opcional"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   135
      Left            =   6720
      TabIndex        =   41
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "*Opcional"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   135
      Left            =   9120
      TabIndex        =   40
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "*Opcional"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   135
      Left            =   9120
      TabIndex        =   39
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "*Opcional"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   135
      Left            =   9120
      TabIndex        =   38
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   9480
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Cuenta"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   8760
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Cancelado"
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Total "
      Height          =   255
      Left            =   4800
      TabIndex        =   18
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Documento"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección "
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RUC Empresa/Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora:"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero Gasto"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String
Sub GuardarCuenta()
Dim nroCuenta As String
Dim codBanco As Integer
Dim codProveedor As Integer

    nroCuenta = Datos.MostrarCampo("Cuentas", "nroCuenta", "nroCuenta='" & Combo1.Text & "'")
    If nroCuenta = "" Then
        codProveedor = Val(Datos.MostrarCampo("Proveedores", "codProveedor", "rif='" & Text4.Text & "'"))
        codBanco = Val(Datos.MostrarCampo("Bancos", "codBanco", "nombre='" & Combo2.Text & "'"))
        If codProveedor > 0 And codBanco > 0 Then
            Conexion.Execute "insert into cuentas(nroCuenta,codBanco,codProveedor) values (" _
            & "'" & Combo1.Text & "'," _
            & "" & codBanco & "," _
            & "" & codProveedor & ")"
        End If
    End If
        
End Sub
Sub guardarProveedor()
    Dim codProveedor As Integer
    codProveedor = Val(Datos.MostrarCampo("Proveedores", "codProveedor", "rif='" & Text4.Text & "'"))
    If codProveedor = 0 Then
        Conexion.Execute "insert into Proveedores(codProveedor,rif,razonsocial,direccion,forma,tiempopago,telefono) values(" _
        & "" & Datos.generarCodigo("Proveedores", "codProveedor") & "," _
        & "'" & Text4.Text & "'," _
        & "'" & Text5.Text & "'," _
        & "'" & Text6.Text & "'," _
        & "'" & Combo3.Text & "'," _
        & "'" & Combo4.Text & "'," _
        & "'" & Text7.Text & "')"
    Else
        Conexion.Execute "update proveedores set " _
        & "rif='" & Text4.Text & "'," _
        & "razonsocial='" & Text5.Text & "'," _
        & "direccion='" & Text6.Text & "'," _
        & "forma='" & Combo3.Text & "'," _
        & "tiempopago='" & Combo4.Text & "'," _
        & "telefono='" & Text7.Text & "' " _
        & "where codProveedor=" & codProveedor & ""
    End If
    
    
End Sub

Sub CargarTablas()
    Tabla = "Gastos"
    campoClave = "codGasto"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    Dim codProveedor As Integer
    codProveedor = Val(Datos.MostrarCampo("Proveedores", "codProveedor", "rif='" & Text4.Text & "'"))
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Val(Text1.Text), Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "codGasto," _
        & "fecha," _
        & "hora," _
        & "codProveedor," _
        & "documento," _
        & "concepto," _
        & "total," _
        & "pagado,proxFecha," _
        & "nrocuenta)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "'" & Text3.Text & "'," _
        & "'" & codProveedor & "'," _
        & "'" & Text8.Text & "'," _
        & "'" & Text9.Text & "'," _
        & "" & Val(Text10.Text) & "," _
        & "" & Val(Text11.Text) & "," _
        & "'" & Format(DTPicker1.value, "mm/dd/yyyy") & "'," _
        & "'" & Combo1.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "fecha='" & Text2.Text & "'," _
            & "hora='" & Text3.Text & "'," _
            & "codProveedor=" & codProveedor & "," _
            & "documento='" & Text8.Text & "'," _
            & "concepto='" & Text9.Text & "'," _
            & "total='" & Text10.Text & "'," _
            & "pagado='" & Text11.Text & "'," _
            & "nroCuenta='" & Combo1.Text & "' " _
            & "Where codGasto=" & Text1.Text & ""
    End If
    SqlActualizacion = iSql
    
End Function

Sub bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Text5.Enabled = es
    Text6.Enabled = es
    Text7.Enabled = es
    Text8.Enabled = es
    Text9.Enabled = es
    Text10.Enabled = es
    Text11.Enabled = es
    Combo1.Enabled = es
    Combo2.Enabled = es
    Combo3.Enabled = es
    Combo4.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
    Label7.Enabled = es
    Label8.Enabled = es
    Label9.Enabled = es
    Label10.Enabled = es
    Label11.Enabled = es
    Label12.Enabled = es
    Label17.Enabled = es
    Label18.Enabled = es
    Label19.Enabled = es
    Label13.Enabled = es
    DTPicker1.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Combo1.Text = ""
    Combo3.Text = ""
    Combo4.Text = ""
    DTPicker1.value = Date
End Sub
Sub BuscarProveedor()
On Error Resume Next
    Dim codProveedor As Integer
    codProveedor = Val(Datos.MostrarCampo("Proveedores", "codProveedor", "rif='" & Text4.Text & "'"))
    Text5.Text = Datos.MostrarCampo("Proveedores", "razonsocial", "rif='" & Text4.Text & "'")
    Text6.Text = Datos.MostrarCampo("Proveedores", "direccion", "rif='" & Text4.Text & "'")
    Text7.Text = Datos.MostrarCampo("Proveedores", "telefono", "rif='" & Text4.Text & "'")
    Combo3.Text = Datos.MostrarCampo("Proveedores", "forma", "rif='" & Text4.Text & "'")
    Combo4.Text = Datos.MostrarCampo("Proveedores", "tiempopago", "rif='" & Text4.Text & "'")
    If codProveedor > 0 Then
        Datos.llenarCombo "select nroCuenta from cuentas where codProveedor=" & codProveedor, Combo1
    Else
        Combo1.Clear
    End If
End Sub
Sub buscarCuenta()
Dim codBanco As Integer
    codBanco = Val(Datos.MostrarCampo("Cuentas", "codBanco", "nroCuenta='" & Combo1.Text & "'"))
    Combo2.Text = Datos.MostrarCampo("Bancos", "nombre", "codBanco=" & codBanco & "")
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text4.Text = Datos.MostrarCampo("Proveedores", "rif", "codProveedor=" & rs(3))
    Text8.Text = rs(4)
    Text9.Text = rs(5)
    Text10.Text = rs(6)
    DTPicker1.value = rs("proxFecha")
    Text11.Text = rs(7)
    BuscarProveedor
    Combo1.Text = rs(8)
    
    buscarCuenta
End Sub

Private Sub Combo1_Change()
    buscarCuenta
End Sub

Private Sub Combo1_Click()
    buscarCuenta
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
    Text2.Text = Date
    Text3.Text = Time
    Text4.SetFocus
End Sub


Private Sub Command12_Click()
    MostrarCatalogo "select rif, razonsocial,direccion,telefono  from proveedores"
    Text4.Text = Catalogo.Resultado
    
End Sub

Private Sub Command2_Click()
    guardarProveedor
    GuardarCuenta
    Formularios.Guardar Me
End Sub
Private Sub Command3_Click()
    Formularios.cancelar Me
End Sub
Private Sub Command4_Click()
    Formularios.modificar Me
End Sub
Private Sub Command5_Click()
    Formularios.Eliminar Me
End Sub
Private Sub Command6_Click()
    MostrarCatalogo "select * from [" & Tabla & "]"
    Text1.Text = Catalogo.Resultado
    Formularios.Buscar Me
End Sub
Private Sub Command7_Click()
    Unload Me
End Sub
Private Sub Command8_Click()
    Formularios.Primero Me
End Sub
Private Sub Command9_Click()
    Formularios.Anterior Me
End Sub
Private Sub Command10_Click()
    Formularios.Siguiente Me
End Sub
Private Sub Command11_Click()
    Formularios.Ultimo Me
End Sub
Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Formularios.Botones Me, 1
    Formularios.cancelar Me
    bloquear False
    CargarTablas
    Datos.llenarCombo "select nombre from bancos", Combo2
End Sub

Private Sub Text4_Change()
    BuscarProveedor
End Sub
