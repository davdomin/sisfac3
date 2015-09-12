VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Empleados"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Default         =   -1  'True
      Height          =   225
      Left            =   -240
      TabIndex        =   28
      Top             =   -360
      Width           =   375
   End
   Begin VB.TextBox Text8 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3900
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Catalogo1 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "Permite Buscar en el catalogo"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6540
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3900
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3900
      TabIndex        =   2
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3900
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3900
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3900
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3900
      TabIndex        =   6
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   4980
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
      Left            =   3240
      TabIndex        =   13
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   5535
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
      Left            =   3960
      TabIndex        =   12
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   5535
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
      Left            =   4680
      TabIndex        =   11
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   5535
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
      Left            =   5400
      TabIndex        =   10
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   5535
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   1680
      TabIndex        =   27
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   5580
      TabIndex        =   26
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI:"
      Height          =   255
      Left            =   1740
      TabIndex        =   25
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombres y Apellidos:"
      Height          =   255
      Left            =   1740
      TabIndex        =   24
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      Height          =   255
      Left            =   1740
      TabIndex        =   23
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
      Height          =   255
      Left            =   1740
      TabIndex        =   22
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Correo Electronico"
      Height          =   255
      Left            =   1740
      TabIndex        =   21
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo que Ocupa"
      Height          =   255
      Left            =   1740
      TabIndex        =   20
      Top             =   3120
      Width           =   2055
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String

Sub CargarTablas()
    Tabla = "Empleados"
    campoClave = "codEmpleado"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Val(Text1.Text), Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "codEmpleado," _
        & "cedula," _
        & "nombre," _
        & "direccion," _
        & "telefono," _
        & "correoe," _
        & "clave," _
        & "codCargo)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "'" & Text3.Text & "'," _
        & "'" & Text4.Text & "'," _
        & "'" & Text5.Text & "'," _
        & "'" & Text6.Text & "'," _
        & "'" & Text8.Text & "'," _
        & "'" & Text7.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "cedula='" & Text2.Text & "'," _
            & "nombre=" & Text3.Text & "," _
            & "direccion=" & Text4.Text & "," _
            & "telefono=" & Text5.Text & "," _
            & "correoe=" & Text6.Text & "," _
            & "codCargo=" & Text8.Text & "," _
            & "Where codEmpleado=" & Text1.Text & ""
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
    Text6.Enabled = es
    Text7.Enabled = es
    Text8.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
    Label7.Enabled = es
    Label8.Enabled = es
    Catalogo1.Enabled = es
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
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text4.Text = rs(3)
    Text5.Text = rs(4)
    Text6.Text = rs(5)
    Text7.Text = rs(6)
    Text8.Text = rs(7)

End Sub

Private Sub Catalogo1_Click()
    MostrarCatalogo "select *from cargos"
    Text7.Text = Catalogo.Resultado
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
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
    MostrarCatalogo "select codEmpleado, nombre,direccion,telefono ,cedula from [" & Tabla & "]"
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







