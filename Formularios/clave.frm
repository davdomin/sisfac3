VERSION 5.00
Begin VB.Form Form17 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form17"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      Picture         =   "clave.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Default         =   -1  'True
      Height          =   495
      Left            =   1080
      Picture         =   "clave.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   2520
      Picture         =   "clave.frx":0A44
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Otro Producto de Shirley Sistemas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   1200
      Picture         =   "clave.frx":1E1E
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SisFac 2.0 Sistema de Facturación      2011"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   0
      Picture         =   "clave.frx":31F8
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim rs As New ADODB.Recordset
    Dim aux As VbMsgBoxResult
    rs.Open "select *from usuarios where clave='" & Text1.Text & "'", Conexion
    If rs.EOF Then
        aux = MsgBox("Acceso denegado")
    Else
        Form1.MnuMantiene.Enabled = True
        Form1.MnuReporte.Enabled = True
        Form1.cuentaBancaria.Enabled = True
        Form1.Repo.Enabled = True
        Form1.produ.Enabled = True
        Form1.Toolbar1.Buttons(1).Enabled = True
        Form1.Toolbar1.Buttons(3).Enabled = True
        Form1.Toolbar1.Buttons(5).Enabled = True
        Form1.Toolbar1.Buttons(6).Enabled = True
        Form1.Toolbar1.Buttons(9).Enabled = True
        Form1.Toolbar1.Buttons(11).Enabled = True
        Form1.Toolbar1.Buttons(12).Enabled = True
        
        
        
        codUsuario = rs(0)
        NivelEntro = rs("nivel")
        If LCase(NivelEntro) <> "administrador" Then
            Form1.MnuMantiene.Enabled = False
            Form1.cuentaBancaria.Enabled = False
            Form1.Repo.Enabled = False
            Form1.produ.Enabled = False
        
            Form1.Toolbar1.Buttons(1).Enabled = False
            Form1.Toolbar1.Buttons(3).Enabled = False
            Form1.Toolbar1.Buttons(5).Enabled = True
            Form1.Toolbar1.Buttons(6).Enabled = False
            Form1.Toolbar1.Buttons(9).Enabled = False
            Form1.Toolbar1.Buttons(11).Enabled = False
            Form1.Toolbar1.Buttons(12).Enabled = True
        
        
        End If
        If LCase(NivelEntro) = "vendedor" Then
            Form1.MnuReporte.Enabled = False
        End If
        UsuarioSession = rs(1)
        Unload Me
        Form1.Show
        Form1.Enabled = True
    End If
    
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
'    Label3.Caption = Proyecto.NombreEmpresa
Text1.Text = ""
End Sub

