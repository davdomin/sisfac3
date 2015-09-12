VERSION 5.00
Begin VB.Form Form39 
   Caption         =   "Pagos"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Segoe UI Symbol"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form39"
   ScaleHeight     =   6780
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto: 000000,000."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   9975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Pedido:"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form39"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public numero As Long
Sub BuscarPedido()
    Dim codPedido As Integer
    codPedido = Val(Text1.Text)
    Label2.Caption = "Cliente: "
    Label3.Caption = "Monto: "
    If codPedido <> 0 Then
        Dim codCliente As Integer
        Dim monto As Double
        codCliente = Val(Datos.MostrarCampo("PedidoEnc", "codCliente", "codPedido=" & codPedido))
        monto = Val(Datos.MostrarCampo("PedidoEnc", "total", "codPedido=" & codPedido))
        
        If codCliente <> 0 Then
            Dim nombrePaciente As String
            nombrePaciente = Datos.MostrarCampo("Clientes", "razonsocial", "codCliente=" & codCliente)
            Label2.Caption = "Cliente: " & nombrePaciente
            Label3.Caption = "Monto: " & monto
        End If
    End If
End Sub

Private Sub Command1_Click()
    numero = Val(Text1.Text)
    Unload Me
End Sub

Private Sub Command2_Click()
    numero = -1
    Unload Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Label2.Caption = "Cliente: "
    Text1.Text = ""
    Label3.Caption = "Monto: "
End Sub

Private Sub Text1_Change()
BuscarPedido
End Sub
