VERSION 5.00
Begin VB.Form Form18 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   4515
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   4515
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   4515
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   4515
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "usuarios.frx":0000
      Left            =   3960
      List            =   "usuarios.frx":000D
      TabIndex        =   7
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   3960
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   5100
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Alias"
      Height          =   255
      Left            =   1740
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   4140
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String

Sub CargarTablas()
    Tabla = "Usuarios"
    campoClave = "CodUsuario"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Text1.Text, Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "CodUsuario," _
        & "alias," _
        & "nivel," _
        & "clave)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "'" & Combo1.Text & "'," _
        & "'" & Text3.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "alias='" & Text2.Text & "'," _
            & "Nivel='" & Combo1.Text & "'," _
            & "clave='" & Text3.Text & "'" _
            & "Where CodUsuario=" & Text1.Text & ""
    End If
'    InputBox iSql, iSql, iSql
    SqlActualizacion = iSql
    
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Combo1.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Combo1.Text = ""
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Combo1.Text = rs(3)
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
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
    MostrarCatalogo "select *from [" & Tabla & "]"
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



