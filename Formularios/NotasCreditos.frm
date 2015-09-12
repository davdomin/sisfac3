VERSION 5.00
Begin VB.Form Form40 
   Caption         =   "Notas de debito"
   ClientHeight    =   7230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form40"
   ScaleHeight     =   7230
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   5400
      TabIndex        =   24
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   2280
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   25
      Top             =   5160
      Width           =   1200
   End
   Begin VB.TextBox Text5 
      Height          =   390
      Left            =   5760
      TabIndex        =   5
      Top             =   2880
      Width           =   2040
   End
   Begin VB.TextBox Text9 
      Height          =   390
      Left            =   8400
      TabIndex        =   9
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   495
      Left            =   8040
      TabIndex        =   21
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   6720
      TabIndex        =   20
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   3960
      TabIndex        =   19
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   2520
      TabIndex        =   18
      Top             =   6600
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   390
      ItemData        =   "NotasCreditos.frx":0000
      Left            =   2880
      List            =   "NotasCreditos.frx":000A
      TabIndex        =   4
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      Height          =   390
      Left            =   8400
      TabIndex        =   8
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   390
      Left            =   5520
      TabIndex        =   7
      Top             =   3960
      Width           =   2040
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "NotasCreditos.frx":001F
      Top             =   3840
      Width           =   4935
   End
   Begin VB.TextBox Text4 
      Height          =   390
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   2040
   End
   Begin VB.TextBox Text3 
      Height          =   390
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   9735
   End
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   2040
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   8880
      TabIndex        =   0
      Text            =   "Text7.Text = """""
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Doc."
      Height          =   255
      Left            =   5760
      TabIndex        =   23
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Imponible:"
      Height          =   495
      Left            =   6360
      TabIndex        =   22
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Documento:"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "IGV"
      Height          =   495
      Left            =   7560
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto:"
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Razon de la devolucion:"
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DNI RUC:"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre y Apellido del solicitante"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Nota"
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form40"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Buscar()
    Dim t As New ADODB.Recordset
    Dim a As String
    t.Open "select *from credito where codCredito=" & Val(Text1.Text), Conexion
    If t.EOF Then
        a = Text1.Text
        Limpiar
        Text1.Text = a
        Command2.Enabled = True
    Else
        Mostrar t
        Command1.Enabled = False
    End If
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs("nCred")
    Text2.Text = rs("fecha")
    Text3.Text = rs("persona")
    Text4.Text = rs("dni")
    Combo1.Text = rs("tipoDoc")
    Text5.Text = rs("nFac")
    Text6.Text = rs("razon")
    Text7.Text = rs("Monto")
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = Date
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Combo1.Text = ""
End Sub
Sub bloquear(st As Boolean)
    Text1.Enabled = st
    Text2.Enabled = st
    Text3.Enabled = st
    Text4.Enabled = st
    Text5.Enabled = st
    Text6.Enabled = st
    Text7.Enabled = st
    Text8.Enabled = st
    Combo1.Enabled = st
End Sub

Sub cancelar()
    bloquear False
    Limpiar
End Sub
Private Sub Command1_Click()
    Limpiar
    bloquear True
    Text2.Text = Date
    Text3.SetFocus
    Text1.Text = Datos.generarCodigo("credito", "codCredito")
End Sub

Private Sub Command2_Click()

Dim sql As String
Dim filtro As String
Dim Archivo As String
    sql = "insert into credito(codCredito,fecha,persona,dni,tipoDoc,nFac,razon,monto,igv) values(" _
    & "" & Val(Text1.Text) & "," _
    & "'" & Text2.Text & "'," _
    & "'" & Text3.Text & "'," _
    & "'" & Text4.Text & "'," _
    & "'" & Combo1.Text & "'," _
    & "'" & Text5.Text & "'," _
    & "'" & Text6.Text & "'," _
    & "" & Val(Text7.Text) & "," _
    & "" & Val(Text8.Text) & ")"

    Conexion.Execute sql
    MsgBox "Datos guardados"
    Dim Formulas() As String
    filtro = "{credito.codCredito}=" & Val(Text1.Text)
    Archivo = App.Path & "\reportes\Ncred.rpt"

    
    Datos.CargarReporte filtro, Archivo, Formulas
    
    cancelar
    
End Sub

Private Sub Command3_Click()
    cancelar
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    Dim Formulas() As String
    Dim filtro As String
    Dim Archivo As String
    filtro = "{credito.codCredito}=" & Val(Text1.Text)
    Archivo = App.Path & "\reportes\Ncred.rpt"
    Datos.CargarReporte filtro, Archivo, Formulas

End Sub

Private Sub Form_Load()
    Limpiar
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    Buscar
End Sub

Private Sub Text7_Change()
    Text9.Text = Round(Formularios.sinIva(Val(Text7.Text)), 3)
    Text8.Text = Round(Val(Text9.Text) * (Proyecto.PIVa), 3)
End Sub
