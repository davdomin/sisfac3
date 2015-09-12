VERSION 5.00
Begin VB.Form Form23 
   Caption         =   "Proveedores"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form23"
   ScaleHeight     =   5655
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      Caption         =   "&Cuentas Bancarias"
      Height          =   375
      Left            =   720
      TabIndex        =   22
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3420
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3420
      TabIndex        =   2
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3420
      TabIndex        =   3
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3420
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   4260
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   4815
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   4815
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   4815
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   4815
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Default         =   -1  'True
      Height          =   225
      Left            =   -840
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   5100
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cod/Ruc:"
      Height          =   255
      Left            =   1260
      TabIndex        =   20
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social"
      Height          =   255
      Left            =   1260
      TabIndex        =   19
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      Height          =   255
      Left            =   1260
      TabIndex        =   18
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
      Height          =   255
      Left            =   1260
      TabIndex        =   17
      Top             =   2640
      Width           =   2055
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String

Sub CargarTablas()
    Tabla = "Proveedores"
    campoClave = "codProveedor"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Val(Text1.Text), Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "codProveedor," _
        & "Rif," _
        & "razonsocial," _
        & "direccion," _
        & "telefono)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "'" & Text3.Text & "'," _
        & "'" & Text4.Text & "'," _
        & "'" & Text5.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "rif='" & Text2.Text & "'," _
            & "RazonSocial='" & Text3.Text & "'," _
            & "direccion='" & Text4.Text & "'," _
            & "telefono='" & Text5.Text & "' " _
            & "Where codproveedor=" & Text1.Text & ""
    End If
   ' a = InputBox(iSql, iSql, iSql)
    SqlActualizacion = iSql
    
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Text5.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text4.Text = rs(3)
    Text5.Text = rs(4)
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
End Sub

Private Sub Command12_Click()
    If Command4.Enabled Then
        Form25.Text1.Text = Text2.Text
        Form25.Text2.Text = Text3.Text
        Form25.mostrarCuentas
        Form25.Show vbModal
        
    End If
End Sub

Private Sub Command13_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Command2_Click()
    Formularios.Guardar Me
End Sub

Private Sub Command3_Click()
    Formularios.Cancelar Me
End Sub

Private Sub Command4_Click()
    Formularios.Modificar Me
End Sub

Private Sub Command5_Click()
    Formularios.Eliminar Me
End Sub

Private Sub Command6_Click()
    MostrarCatalogo "select codProveedor, razonsocial,direccion,telefono ,rif from [" & Tabla & "]"
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
    Bloquear False
    CargarTablas
End Sub







