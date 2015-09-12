VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Registro de Clientes"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   4410
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Default         =   -1  'True
      Height          =   225
      Left            =   -240
      TabIndex        =   25
      Top             =   -360
      Width           =   375
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
      Left            =   5880
      TabIndex        =   17
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   3855
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
      Left            =   5160
      TabIndex        =   16
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   3855
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
      Left            =   4440
      TabIndex        =   15
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   3855
      Width           =   615
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
      Left            =   3720
      TabIndex        =   14
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   3855
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6960
      TabIndex        =   12
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   3300
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   3300
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   3300
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4380
      TabIndex        =   6
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   4380
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4380
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4380
      TabIndex        =   3
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4380
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4380
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7020
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      Height          =   255
      Left            =   2220
      TabIndex        =   24
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Correo Electronico"
      Height          =   255
      Left            =   2220
      TabIndex        =   23
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
      Height          =   255
      Left            =   2220
      TabIndex        =   22
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      Height          =   255
      Left            =   2220
      TabIndex        =   21
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social"
      Height          =   255
      Left            =   2220
      TabIndex        =   20
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DNI/RUC:"
      Height          =   255
      Left            =   2220
      TabIndex        =   19
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   6060
      TabIndex        =   18
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String

Sub CargarTablas()
    Tabla = "Clientes"
    campoClave = "codCliente"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Val(Text1.Text), Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "codCliente," _
        & "cedRif," _
        & "razonsocial," _
        & "direccion," _
        & "telefono," _
        & "correoe," _
        & "fax)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "'" & Text3.Text & "'," _
        & "'" & Text4.Text & "'," _
        & "'" & Text5.Text & "'," _
        & "'" & Text6.Text & "'," _
        & "'" & Text7.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "cedrif='" & Text2.Text & "'," _
            & "RazonSocial='" & Text3.Text & "'," _
            & "direccion='" & Text4.Text & "'," _
            & "telefono='" & Text5.Text & "'," _
            & "correoe='" & Text6.Text & "'," _
            & "fax='" & Text7.Text & "'" _
            & "Where codCliente=" & Text1.Text & ""
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
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
    Label7.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
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
    MostrarCatalogo "select codCliente, razonsocial,direccion,telefono ,cedrif from [" & Tabla & "]"
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





