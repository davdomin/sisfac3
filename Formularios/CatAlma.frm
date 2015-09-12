VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Categorias del Almacen"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1935
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3030
      TabIndex        =   1
      Top             =   2295
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   323
      TabIndex        =   12
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1523
      TabIndex        =   11
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2723
      TabIndex        =   10
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   3923
      TabIndex        =   9
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   4260
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5243
      TabIndex        =   8
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6443
      TabIndex        =   7
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7643
      TabIndex        =   6
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   3203
      TabIndex        =   5
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   4815
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   3923
      TabIndex        =   4
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   4815
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   4643
      TabIndex        =   3
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   4815
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   5363
      TabIndex        =   2
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   4815
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   4710
      TabIndex        =   14
      Top             =   1935
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   870
      TabIndex        =   13
      Top             =   2295
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String

Sub CargarTablas()
    Tabla = "Categorias"
    campoClave = "CodCategoria"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Text1.Text, Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "CodCategoria," _
        & "Descripcion)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "Descripcion='" & Text2.Text & "'" _
            & "Where CodCategoria=" & Text1.Text & ""
    End If
    SqlActualizacion = iSql
    
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
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
    Formularios.modificar Me
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

