VERSION 5.00
Begin VB.Form Form46 
   Caption         =   "Reporte "
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form46"
   ScaleHeight     =   4485
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frmRptProProve.frx":0000
      Left            =   1080
      List            =   "frmRptProProve.frx":000D
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2760
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   3960
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mercaderia"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
End
Attribute VB_Name = "Form46"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function filtro()
    Dim f As String
    If Combo1.Text <> "" Then
        f = f & "{productos.descripcion}='" & Combo1.Text & "'"
    End If
    filtro = f
    
End Function
Private Sub Command1_Click()
    Dim f As String
    Dim Archivo As String
    Dim Formulas() As String
    Dim cont As Integer
    cont = 0
    
    ReDim Preserve Formulas(cont)

    cont = cont + 1

    
    
    f = filtro
    Archivo = App.Path & "\reportes\proprove.rpt"
    
    
    Datos.CargarReporte filtro, Archivo, Formulas

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Datos.llenarCombo "select descripcion from productos order by descripcion", Combo1
    
    
    DTPicker1.value = Date
    DTPicker2.value = Date
End Sub


