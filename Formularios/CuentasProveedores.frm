VERSION 5.00
Begin VB.Form Form25 
   Caption         =   "Form25"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form25"
   ScaleHeight     =   10.319
   ScaleMode       =   0  'User
   ScaleWidth      =   12.383
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   2460
      Left            =   3600
      TabIndex        =   9
      Top             =   2040
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   2460
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   3600
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Cuenta:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Razon Social"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Rif del Proveedor:"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Suprimir(nroCuenta As String)
    Conexion.Execute "delete from Cuentas where nroCuenta='" & nroCuenta & "'"
    MsgBox "Cuenta Suprimida"
    mostrarCuentas
End Sub

Sub Seleccionar(Indice As Integer)
    If Indice <> -1 Then
        List1.Selected(Indice) = True
        List2.Selected(Indice) = True
    End If
End Sub

Sub agregarCuenta()
    Dim nroCuenta As String
    Dim codProveedor As Integer
    Dim codBanco As Integer
    Dim sql As String
    codProveedor = Val(Datos.MostrarCampo("Proveedores", "codProveedor", "rif='" & Text1.Text & "'"))
    codBanco = Val(Datos.MostrarCampo("Bancos", "codBanco", "Nombre='" & Combo1.Text & "'"))
    If codProveedor > 0 And codBanco > 0 Then
        sql = "insert into cuentas(nroCuenta,codProveedor,codBanco) values(" _
        & "'" & Text3.Text & "'," _
        & "" & codProveedor & "," _
        & "" & codBanco & ")"
        Conexion.Execute sql
        MsgBox "Cuenta Agregada"
        mostrarCuentas
    End If
End Sub
Sub mostrarCuentas()
    Dim codProveedor As Integer
    Dim rs As New ADODB.Recordset
    List1.Clear
    List2.Clear
    codProveedor = Datos.MostrarCampo("Proveedores", "codProveedor", "rif='" & Text1.Text & "'")
    rs.Open "select nroCuenta,Nombre from Cuentas,Bancos where " _
    & " Cuentas.codBanco=Bancos.codBanco and codProveedor=" & codProveedor, Conexion
    While Not rs.EOF
        List1.AddItem rs(0)
        List2.AddItem rs(1)
        rs.MoveNext
    Wend
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Combo1.Text = ""
    List1.Clear
    List2.Clear
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        agregarCuenta
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Limpiar
    Datos.llenarCombo "select Nombre from Bancos", Combo1
End Sub

Private Sub List1_Click()
    Seleccionar List1.ListIndex
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nroCuenta  As String
    If KeyCode = 46 Then
        nroCuenta = List1.List(List1.ListIndex)
        Suprimir (nroCuenta)
    End If

End Sub

Private Sub List2_Click()
    Seleccionar List2.ListIndex
End Sub
